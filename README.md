# translate-cn-en-skill

Claude Code Skill：将中文 Word / PDF 文档翻译成英文，完整保留原始排版。

## 安装到 Claude Code

```bash
claude skill install https://github.com/conanlee007/translate-cn-en-skill
```

安装完成后，在 Claude Code 中直接说"帮我翻译这个文件"或"translate this document"即可触发。

## 功能

- **DOCX**：正文 + 表格 + 页眉页脚全覆盖，字体换 Garamond，原始格式 100% 保留
- **PDF**：提取结构后重新排版，生成干净的英文 PDF

## 前置要求

- macOS（Apple Silicon 或 Intel）
- `ANTHROPIC_API_KEY` 环境变量
- 首次运行脚本会自动提示安装依赖

详细配置见 [skill.md](./skill.md)。
