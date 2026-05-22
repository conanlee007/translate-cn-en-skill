#!/usr/bin/env python3
"""
DOCX Chinese → English translator with full format preservation.
Reads every paragraph and table cell, translates Chinese text via Claude API,
writes back in-place — all original fonts, colors, table borders, styles preserved.

Usage: python translate_docx.py <input.docx> [output.docx]
Requires: ANTHROPIC_API_KEY environment variable
"""
import os
import re
import sys
import copy
from pathlib import Path

import anthropic
from docx import Document
from docx.oxml.ns import qn

CLAUDE_MODEL = "claude-sonnet-4-6"
BATCH_SIZE = 40  # texts per Claude call

CJK_RE = re.compile(r'[一-鿿㐀-䶿＀-￯　-〿⺀-⻿]')


def has_cjk(text: str) -> bool:
    return bool(CJK_RE.search(text))


# ── Claude translation ────────────────────────────────────────────────────────

def translate_batch(client: anthropic.Anthropic, texts: list[str]) -> list[str]:
    """Translate a batch of Chinese strings to English, preserving order and count."""
    indexed = {i: t for i, t in enumerate(texts)}
    prompt = (
        "Translate the following Chinese texts to English. "
        "Return ONLY a JSON array with exactly the same number of elements in the same order.\n"
        "Rules:\n"
        "- Use standard English financial/accounting terminology\n"
        "- Express Chinese yuan as CNY (not RMB)\n"
        "- Keep numbers, percentages, dates, stock codes, and '--' exactly unchanged\n"
        "- Keep empty strings as empty strings\n"
        "- No explanations, no markdown fences — only the JSON array\n\n"
        + __import__('json').dumps(list(indexed.values()), ensure_ascii=False)
    )
    msg = client.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=8096,
        messages=[{"role": "user", "content": prompt}],
    )
    raw = msg.content[0].text.strip()
    raw = re.sub(r'^```(?:json)?\s*', '', raw)
    raw = re.sub(r'\s*```\s*$', '', raw)
    m = re.search(r'\[.*\]', raw, re.DOTALL)
    import json
    parsed = json.loads(m.group() if m else raw)
    # Safety: pad if Claude returned fewer items
    if len(parsed) < len(texts):
        parsed += texts[len(parsed):]
    return [str(x) for x in parsed[:len(texts)]]


def translate_all(client: anthropic.Anthropic, texts: list[str]) -> list[str]:
    """Translate a list of texts in batches."""
    results = []
    for i in range(0, len(texts), BATCH_SIZE):
        batch = texts[i:i + BATCH_SIZE]
        translated = translate_batch(client, batch)
        results.extend(translated)
        print(f"    translated {min(i + BATCH_SIZE, len(texts))}/{len(texts)} texts", flush=True)
    return results


# ── DOCX text extraction ──────────────────────────────────────────────────────

def para_full_text(para) -> str:
    return ''.join(run.text for run in para.runs)


def iter_header_footer_paras(doc: Document):
    """Yield all paragraphs from headers and footers across all sections."""
    for section in doc.sections:
        for hf in [
            section.header, section.footer,
            section.first_page_header, section.first_page_footer,
            section.even_page_header, section.even_page_footer,
        ]:
            if hf is not None:
                yield from hf.paragraphs
                for table in hf.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            yield from cell.paragraphs


def collect_paragraph_texts(doc: Document) -> tuple[list, list[str]]:
    """Return (refs, texts) for all paragraphs (body + headers/footers) with CJK content."""
    refs, texts = [], []
    # Body paragraphs
    for para in doc.paragraphs:
        full = para_full_text(para)
        if has_cjk(full):
            refs.append(('para', para))
            texts.append(full)
    # Header / footer paragraphs
    for para in iter_header_footer_paras(doc):
        full = para_full_text(para)
        if has_cjk(full):
            refs.append(('para', para))
            texts.append(full)
    return refs, texts


def collect_table_texts(doc: Document) -> tuple[list, list[str]]:
    """Return (refs, texts) for all table cells with CJK content."""
    refs, texts = [], []
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    full = para_full_text(para)
                    if has_cjk(full):
                        refs.append(('cell_para', para))
                        texts.append(full)
    return refs, texts


# ── DOCX text writing ─────────────────────────────────────────────────────────

FONT_NAME = 'Garamond'


def set_run_font(run, font_name: str):
    """Set font at run level, including the XML-level eastAsia override."""
    run.font.name = font_name
    # Also set the eastAsia font in the XML so CJK font fallback doesn't override
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = rPr.makeelement(qn('w:rFonts'))
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)
    rFonts.set(qn('w:eastAsia'), font_name)
    rFonts.set(qn('w:cs'), font_name)


def set_para_text(para, new_text: str):
    """Replace paragraph text and update font; preserve other run formatting."""
    if not para.runs:
        return
    first_run = para.runs[0]
    first_run.text = new_text
    set_run_font(first_run, FONT_NAME)
    for run in para.runs[1:]:
        run.text = ''


# ── Main ──────────────────────────────────────────────────────────────────────

def translate_docx(input_path: str, output_path: str):
    api_key = os.environ.get('ANTHROPIC_API_KEY')
    if not api_key:
        sys.exit('Error: ANTHROPIC_API_KEY is not set.')

    client = anthropic.Anthropic(api_key=api_key)
    doc = Document(input_path)

    # Collect all translatable text
    print('Collecting text from paragraphs...')
    para_refs, para_texts = collect_paragraph_texts(doc)
    print(f'  {len(para_texts)} paragraphs with Chinese text')

    print('Collecting text from tables...')
    cell_refs, cell_texts = collect_table_texts(doc)
    print(f'  {len(cell_texts)} table cell paragraphs with Chinese text')

    all_refs = para_refs + cell_refs
    all_texts = para_texts + cell_texts
    print(f'\nTranslating {len(all_texts)} text blocks total...')

    all_translated = translate_all(client, all_texts)

    # Write translations back
    print('\nWriting translations back...')
    for (kind, para), translated in zip(all_refs, all_translated):
        set_para_text(para, translated)

    doc.save(output_path)
    print(f'\nSaved → {output_path}')


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('Usage: python translate_docx.py <input.docx> [output.docx]')
        sys.exit(1)
    inp = sys.argv[1]
    out = (sys.argv[2] if len(sys.argv) > 2
           else str(Path(inp).parent / (Path(inp).stem + '-EN.docx')))
    translate_docx(inp, out)
