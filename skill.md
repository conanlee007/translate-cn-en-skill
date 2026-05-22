---
name: translate-cn-en
description: >
  将中文文档（DOCX 或 PDF）翻译成英文，完整保留原始排版与格式。
  触发词：翻译、中文翻译成英文、translate、Chinese to English、翻译这个文件、
  帮我翻译这个 Word、帮我翻译这个 PDF、translate this document、
  translate this report、翻译报告、translate DOCX、translate PDF。
---

# 中文 → 英文文档翻译 Skill

将中文 Word（DOCX）或 PDF 文档翻译成英文。

- **DOCX**：直接读取文档语义结构，翻译所有段落、表格单元格、页眉页脚，字体统一替换为 Garamond，100% 保留原始格式（字号、颜色、表格边框、合并单元格、样式）
- **PDF**：提取文字与表格结构，Claude 翻译后用 reportlab 重新排版生成干净的英文 PDF

## 前置配置（首次使用）

### 1. 安装依赖

```bash
export PATH="$HOME/.local/bin:$PATH"

# 安装 uv（如已安装跳过）
curl -LsSf https://astral.sh/uv/install.sh | sh

# 安装 pdf2zh（包含 pymupdf、pdfplumber 等）
uv tool install --python 3.12 pdf2zh

# 安装额外依赖
uv pip install anthropic python-docx pdfplumber reportlab \
  --python ~/.local/share/uv/tools/pdf2zh/bin/python
```

### 2. 配置 API Key

```bash
# 写入 ~/.zshrc（永久生效）
echo 'export ANTHROPIC_API_KEY="sk-ant-..."' >> ~/.zshrc
source ~/.zshrc
```

### 3. 下载翻译脚本

将本 repo 中的脚本放到你习惯的目录，例如 `~/scripts/`：

```bash
mkdir -p ~/scripts
# 下载两个脚本
curl -o ~/scripts/translate_docx.py \
  https://raw.githubusercontent.com/conanlee007/translate-cn-en-skill/main/translate_docx.py
curl -o ~/scripts/translate_pdf.py \
  https://raw.githubusercontent.com/conanlee007/translate-cn-en-skill/main/translate_pdf.py
```

---

## 执行步骤

### Step 1 — 判断文件类型

检查用户提供的文件扩展名：
- `.docx` → 走 DOCX 流程
- `.pdf` → 走 PDF 流程

如果用户没有给出路径，询问文件位置。

### Step 2A — DOCX 翻译

```bash
export PATH="$HOME/.local/bin:$PATH"
~/.local/share/uv/tools/pdf2zh/bin/python ~/scripts/translate_docx.py \
  "/path/to/input.docx"
```

输出文件自动保存在同目录，文件名加 `-EN` 后缀：
`input-EN.docx`

**覆盖范围：**
- ✅ 正文段落
- ✅ 所有表格单元格
- ✅ 页眉 / 页脚
- ✅ 字体统一替换为 Garamond
- ✅ 原始格式（字号、加粗、颜色、表格样式）完整保留

### Step 2B — PDF 翻译

```bash
export PATH="$HOME/.local/bin:$PATH"
~/.local/share/uv/tools/pdf2zh/bin/python ~/scripts/translate_pdf.py \
  "/path/to/input.pdf"
```

输出文件自动保存在同目录，文件名加 `-en-v2` 后缀：
`input-en-v2.pdf`

**说明：**
- PDF 方案重新排版，布局风格接近原文但非像素级复刻
- 表格、标题、正文全部翻译
- 货币单位统一为 CNY

### Step 3 — 输出结果

翻译完成后告知用户：
- 输出文件的完整路径
- 翻译的文本块数量（脚本会打印）
- 如有报错页面，说明哪页失败及原因

---

## 翻译规范

Claude API 调用时遵守以下规则（已内置在脚本中）：

| 规则 | 说明 |
|---|---|
| 货币表达 | 人民币统一写 CNY，不用 RMB |
| 数字保留 | 所有数字、百分比、日期原样不动 |
| 财务术语 | 使用标准英文会计术语 |
| 股票代码 | 原样保留 |
| 空单元格 | 保持为空，不填充内容 |
| 字体 | DOCX 输出统一 Garamond |

---

## 常见问题

| 问题 | 处理方式 |
|---|---|
| `ANTHROPIC_API_KEY not set` | 检查 ~/.zshrc 是否已写入并 `source ~/.zshrc` |
| PDF 某页翻译失败 | 脚本自动 chunked 重试，若仍失败会在该页显示错误提示 |
| 扫描版 PDF（无文字层）| 本工具不支持，需先 OCR 处理 |
| 表格合并单元格丢失 | DOCX 模式完整保留；PDF 模式因重新排版可能丢失 |
