#!/usr/bin/env python3
"""
PDF structured translator: pdfplumber → Claude API → reportlab
Extracts tables and text, translates via Claude, rebuilds as clean English PDF.

Usage: python translate_pdf_v2.py <input.pdf> [output.pdf]
Requires: ANTHROPIC_API_KEY environment variable
"""
import pdfplumber
import anthropic
import json
import os
import re
import sys
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, HRFlowable,
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

CLAUDE_MODEL = "claude-sonnet-4-6"
PAGE_W, PAGE_H = A4
MARGIN = 1.8 * cm
TABLE_WIDTH = PAGE_W - 2 * MARGIN


# ── Styles ────────────────────────────────────────────────────────────────────

def setup_styles():
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle('H1', parent=styles['Heading1'],
        fontSize=13, spaceBefore=14, spaceAfter=5,
        textColor=colors.HexColor('#1a3a5c')))
    styles.add(ParagraphStyle('H2', parent=styles['Heading2'],
        fontSize=11, spaceBefore=10, spaceAfter=4,
        textColor=colors.HexColor('#1a3a5c')))
    styles.add(ParagraphStyle('H3', parent=styles['Heading3'],
        fontSize=9.5, spaceBefore=7, spaceAfter=3,
        textColor=colors.HexColor('#2d5a8e')))
    styles.add(ParagraphStyle('Body', parent=styles['Normal'],
        fontSize=9, leading=14, spaceAfter=4))
    styles.add(ParagraphStyle('Cell', parent=styles['Normal'],
        fontSize=7.5, leading=10))
    styles.add(ParagraphStyle('CellBold', parent=styles['Normal'],
        fontSize=7.5, leading=10, fontName='Helvetica-Bold'))
    return styles


def make_table_style(has_header=True):
    base = [
        ('FONTNAME',    (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE',    (0, 0), (-1, -1), 7.5),
        ('LEADING',     (0, 0), (-1, -1), 10),
        ('GRID',        (0, 0), (-1, -1), 0.4, colors.HexColor('#aaaaaa')),
        ('VALIGN',      (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 5),
        ('RIGHTPADDING',(0, 0), (-1, -1), 5),
        ('TOPPADDING',  (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING',(0,0), (-1, -1), 4),
    ]
    if has_header:
        base += [
            ('BACKGROUND',  (0, 0), (-1, 0), colors.HexColor('#e8edf2')),
            ('TEXTCOLOR',   (0, 0), (-1, 0), colors.HexColor('#1a1a1a')),
            ('FONTNAME',    (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE',    (0, 0), (-1, 0), 8),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1),
             [colors.white, colors.HexColor('#f4f7fa')]),
        ]
    else:
        base += [
            ('ROWBACKGROUNDS', (0, 0), (-1, -1),
             [colors.white, colors.HexColor('#f4f7fa')]),
        ]
    return TableStyle(base)


# ── Extraction ────────────────────────────────────────────────────────────────

def extract_page(page):
    """Return (text: str, tables: list[list[list[str]]])."""
    found = page.find_tables()
    tables = []
    for t in found:
        rows = t.extract()
        cleaned = []
        for row in rows:
            r = [str(c).strip() if c is not None else '' for c in row]
            if any(r):
                cleaned.append(r)
        if cleaned:
            tables.append(cleaned)
    text = page.extract_text() or ''
    return text, tables


# ── Translation ───────────────────────────────────────────────────────────────

def _call_claude(client, prompt, max_tokens=8096):
    msg = client.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=max_tokens,
        messages=[{"role": "user", "content": prompt}],
    )
    raw = msg.content[0].text.strip()
    raw = re.sub(r'^```(?:json)?\s*', '', raw)
    raw = re.sub(r'\s*```\s*$', '', raw)
    m = re.search(r'\{.*\}', raw, re.DOTALL)
    return json.loads(m.group() if m else raw)


def translate_text_block(client, page_num, text):
    """Translate just the text portion of a page."""
    prompt = f"""Translate this Chinese financial report text (page {page_num}) to English.
Return a JSON object: {{"elements": [{{"type":"heading","level":1,"text":"..."}},{{"type":"paragraph","text":"..."}}]}}
Rules: use standard accounting English, express Chinese yuan as CNY (not RMB), keep numbers/dates unchanged, return ONLY raw JSON.

TEXT:
{text}"""
    return _call_claude(client, prompt)


def translate_table(client, page_num, table_idx, rows):
    """Translate a single table."""
    prompt = f"""Translate this Chinese financial table (page {page_num}, table {table_idx}) to English.
Return a JSON object: {{"elements": [{{"type":"table","rows":[["Header1","Header2"],["val","val"]]}}]}}
Rules: translate all Chinese cell text, keep all numbers/percentages/dates unchanged, empty cells stay "", return ONLY raw JSON.

TABLE:
{json.dumps(rows, ensure_ascii=False)}"""
    return _call_claude(client, prompt)


def translate_page(client, page_num, text, tables):
    """Translate page content. Falls back to chunked mode if combined call fails."""
    tables_str = json.dumps(tables, ensure_ascii=False, indent=2) if tables else '[]'
    combined_prompt = f"""Translate page {page_num} of a Chinese quarterly report (Loctek Ergonomics Q1 2026) to English.
Return a single JSON object (no markdown):
{{"elements":[{{"type":"heading","level":1,"text":"..."}},{{"type":"paragraph","text":"..."}},{{"type":"table","rows":[["H1","H2"],["v1","v2"]]}}]}}
Rules: heading levels 1/2/3 for major/sub/minor; translate ALL Chinese including table cells; keep numbers/dates/percentages unchanged; standard English accounting terms; express Chinese yuan as CNY (not RMB); return ONLY raw JSON.

=== PAGE TEXT ===
{text}

=== TABLES ===
{tables_str}"""

    # Try combined first
    try:
        return _call_claude(client, combined_prompt, max_tokens=8096)
    except Exception:
        pass

    # Fall back: translate text and each table separately, then merge
    print(' [chunked]', end='', flush=True)
    elements = []
    if text.strip():
        try:
            result = translate_text_block(client, page_num, text)
            elements.extend(result.get('elements', []))
        except Exception:
            elements.append({"type": "paragraph", "text": f"[Text translation error — page {page_num}]"})

    for idx, rows in enumerate(tables):
        try:
            result = translate_table(client, page_num, idx + 1, rows)
            elements.extend(result.get('elements', []))
        except Exception:
            elements.append({"type": "paragraph", "text": f"[Table {idx+1} translation error — page {page_num}]"})

    return {"elements": elements}


# ── Rendering ─────────────────────────────────────────────────────────────────

def safe_text(text):
    return (str(text)
            .replace('&', '&amp;')
            .replace('<', '&lt;')
            .replace('>', '&gt;')
            .replace('\n', '<br/>'))


def build_table_flowable(rows, styles):
    if not rows:
        return None
    n_cols = max(len(r) for r in rows)
    # Pad rows to same length
    padded = [r + [''] * (n_cols - len(r)) for r in rows]

    # Heuristic column widths: first col wider (labels), rest equal
    if n_cols == 1:
        col_widths = [TABLE_WIDTH]
    elif n_cols == 2:
        col_widths = [TABLE_WIDTH * 0.55, TABLE_WIDTH * 0.45]
    elif n_cols == 3:
        col_widths = [TABLE_WIDTH * 0.45, TABLE_WIDTH * 0.275, TABLE_WIDTH * 0.275]
    else:
        first = TABLE_WIDTH * 0.35
        rest = (TABLE_WIDTH - first) / (n_cols - 1)
        col_widths = [first] + [rest] * (n_cols - 1)

    data = []
    for i, row in enumerate(padded):
        style = styles['CellBold'] if i == 0 else styles['Cell']
        data.append([Paragraph(safe_text(cell), style) for cell in row])

    tbl = Table(data, colWidths=col_widths, repeatRows=1)
    tbl.setStyle(make_table_style(has_header=True))
    return tbl


def page_flowables(page_data, styles, is_first):
    out = []
    if not is_first:
        out.append(PageBreak())
    out.append(HRFlowable(width='100%', thickness=0.5,
                           color=colors.HexColor('#cccccc'), spaceAfter=6))

    for elem in page_data.get('elements', []):
        kind = elem.get('type', 'paragraph')
        text = elem.get('text', '').strip()

        if kind == 'heading':
            lvl = elem.get('level', 1)
            s = styles.get(f'H{lvl}', styles['H2'])
            if text:
                out.append(Paragraph(safe_text(text), s))

        elif kind == 'paragraph':
            if text:
                out.append(Paragraph(safe_text(text), styles['Body']))

        elif kind == 'table':
            rows = elem.get('rows', [])
            tbl = build_table_flowable(rows, styles)
            if tbl:
                out.append(Spacer(1, 4))
                out.append(tbl)
                out.append(Spacer(1, 8))

    return out


def footer(canvas, doc):
    canvas.saveState()
    canvas.setFont('Helvetica', 7.5)
    canvas.setFillColor(colors.HexColor('#888888'))
    canvas.drawString(MARGIN, MARGIN * 0.6,
                      'Loctek Ergonomic Technology Co., Ltd. — Q1 2026 Quarterly Report (Translated)')
    canvas.drawRightString(PAGE_W - MARGIN, MARGIN * 0.6, f'Page {doc.page}')
    canvas.restoreState()


# ── Main ──────────────────────────────────────────────────────────────────────

def translate_pdf(input_path, output_path):
    api_key = os.environ.get('ANTHROPIC_API_KEY')
    if not api_key:
        sys.exit('Error: ANTHROPIC_API_KEY is not set.')

    client = anthropic.Anthropic(api_key=api_key)
    styles = setup_styles()
    all_flowables = []

    with pdfplumber.open(input_path) as pdf:
        total = len(pdf.pages)
        print(f'Translating {total} pages from: {input_path}')

        for i, page in enumerate(pdf.pages):
            print(f'  Page {i+1}/{total} ...', end='', flush=True)
            text, tables = extract_page(page)

            if not text.strip() and not tables:
                print(' (empty, skipped)')
                continue

            page_data = translate_page(client, i + 1, text, tables)
            n_elem = len(page_data.get('elements', []))
            flowables = page_flowables(page_data, styles, is_first=(i == 0))
            all_flowables.extend(flowables)
            print(f' {n_elem} elements')

    doc = SimpleDocTemplate(
        output_path, pagesize=A4,
        leftMargin=MARGIN, rightMargin=MARGIN,
        topMargin=MARGIN, bottomMargin=MARGIN * 1.5,
    )
    doc.build(all_flowables, onFirstPage=footer, onLaterPages=footer)
    print(f'\nSaved → {output_path}')


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('Usage: python translate_pdf_v2.py <input.pdf> [output.pdf]')
        sys.exit(1)
    inp = sys.argv[1]
    out = (sys.argv[2] if len(sys.argv) > 2
           else str(Path(inp).parent / (Path(inp).stem + '-en-v2.pdf')))
    translate_pdf(inp, out)
