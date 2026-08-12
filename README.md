# ms-office-com-skill

Agent skill for **MS Office COM automation on Windows** — Word / Excel / PowerPoint via `win32com.client` (pywin32).

Designed for AI agents that edit real Office documents: Hermes, Claude Code, Codex, OpenClaw, Cursor, etc.

## Why COM?

Pure-library approaches (python-docx / openpyxl / python-pptx) flatten or lose complex structure:
merged tables, styles, outline levels, headers/footers, fields, images anchored in paragraphs.
The COM interface drives the real Office application and preserves everything — this skill is
built for that path.

## Contents

- `SKILL.md` — the skill itself (golden rules, Word/Excel/PPT core patterns, find/replace danger zone, styles & outline, table surgery, cells & formats, slides & shapes, health checks, budget cascades)
- `references/win32com-cheatsheet.md` — verified working snippets for Word/Excel/PowerPoint
- `references/pitfalls.md` — real incident transcripts (infinite TOC duplication, stale-process corruption, long-string Find error, image-container paragraph loss, style off-by-one)

## Install

### Hermes Agent

```bash
# copy into a profile's skills dir
cp -r ms-office-com-skill ~/AppData/Local/hermes/profiles/<profile>/skills/productivity/ms-office-com-skill
# or use the skills hub / curator
```

### Claude Code

```bash
mkdir -p ~/.claude/skills && cp -r ms-office-com-skill ~/.claude/skills/
```

### Codex (OpenAI)

```bash
mkdir -p ~/.codex/skills && cp -r ms-office-com-skill ~/.codex/skills/
```

### OpenClaw / Cursor / others

Any agent that supports Anthropic-style SKILL.md with YAML frontmatter (`name`, `description`).
Copy the folder into the agent's skills directory.

## Requirements

- **Windows 11 + Microsoft Office 2024 (LTSC)** — primary target; works on Office 2016+ (COM object model stable across these versions)
- Microsoft Office installed (Word / Excel / PowerPoint)
- Python with pywin32: `pip install pywin32`

## Feature Map (by tool)

| Tool | Capabilities |
|---|---|
| **Word** | full read (incl. tables), safe find/replace (paragraph-level, self-match-safe), styles & outline levels, **font rules (中文字体 仿宋_GB2312 + 西文字体 Times New Roman via Name/NameFarEast, GB/T 9704-2012 公文格式速查)**, **font substitution (生僻字缺字形自动回退, Application.SubstituteFont)**, **paragraph/image/table alignment (正文两端对齐+首行缩进2字符 via CharacterUnitFirstLineIndent, 图片/表格居中无缩进)**, table surgery (add/delete columns, merged-cell-safe reads, append tables), **table auto-fit (AI 表格超页边距修复 via AutoFitBehavior 窗口/内容)**, **default page numbers (页脚 PAGE 域, 小五 9pt Times New Roman)**, image insert/guard, **export PDF**, **TOC refresh + field update**, **per-section headers/footers**, **comments (批注) read/write**, **track changes accept/reject** |
| **Excel** | cell/range read/write (batch), formulas, merge, formatting, row/col insert/delete, **PivotTable**, **charts**, **conditional formatting**, **data validation**, **freeze panes**, **autofilter**, **named ranges**, **sheet→PDF export**, SaveAs format constants |
| **PowerPoint** | slide read/add, textboxes, tables, pictures, **export slides to PNG (visual verification)**, **speaker notes**, **template/theme apply**, **slide master**, **transitions**, SaveAs format constants |
| **通用** | crash recovery (taskkill), post-edit health checks, budget cascade verification, **COM object discovery (runtime property/method enumeration)**, **Dispatch vs DispatchEx vs EnsureDispatch selection**, **gen_py cache corruption recovery**, **multi-thread CoInitialize pattern**, **Excel merged-cell read/write rules**, Chinese 公文 formatting conventions |

## License

MIT
