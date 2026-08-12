# ms-office-com-skill

Agent skill for **MS Office COM automation on Windows** — Word / Excel / PowerPoint via `win32com.client` (pywin32).

> ⚠️ **SCOPE: 本 skill 仅适用于 Windows + Microsoft Office + COM (win32com.client)。**
> ❌ 不适用于 Linux/macOS、WPS Office、python-docx/openpyxl/python-pptx 纯库路线、Microsoft Graph 云端 API。
> 其他组合请勿使用本 skill。

Designed for AI agents that edit real Office documents: Hermes, Claude Code, Codex, OpenClaw, Cursor, etc.

## Why COM?

Pure-library approaches (python-docx / openpyxl / python-pptx) flatten or lose complex structure:
merged tables, styles, outline levels, headers/footers, fields, images anchored in paragraphs.
The COM interface drives the real Office application and preserves everything — this skill is
built for that path.

## Contents

- `SKILL.md` — the skill itself (golden rules, Word/Excel/PPT core patterns, find/replace danger zone, styles & outline, fonts & substitution, alignment & table cell defaults, GB/T 7714 references, superscript/subscript semantics, table surgery, cells & formats, slides & shapes, dialog-hang prevention, dispatch selection, health checks, budget cascades)
- `references/win32com-cheatsheet.md` — verified working snippets for Word/Excel/PowerPoint
- `references/pitfalls.md` — real incident transcripts (infinite TOC duplication, stale-process corruption, long-string Find error, image-container paragraph loss, style off-by-one, LLM-generation pitfalls A1-A10, file-corruption pitfalls B1-B3)

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
| **Word** | full read (incl. tables), safe find/replace (paragraph-level, self-match-safe), styles & outline levels, **font rules (中文字体 仿宋_GB2312 + 西文字体 Times New Roman via Name/NameFarEast, GB/T 9704-2012 公文格式速查)**, **font substitution (生僻字缺字形自动回退, Application.SubstituteFont)**, **paragraph/image/table alignment (正文两端对齐+首行缩进2字符 via CharacterUnitFirstLineIndent, 标题/表题/图题居中无缩进, 图片/表格居中无缩进, 图片段单倍行距)**, **table cell defaults (文字上下/左右居中+无缩进+单倍行距, 表头加粗+HeadingFormat 跨页重复标题行)**, **GB/T 7714-2015 references (顺序编码制, 各文献类型格式, 正文上标[1]引用)**, **superscript/subscript semantics (先判语义: m3→m³上标, CO2→CO₂下标, Ca2+→Ca²⁺上标, 数值不动)**, table surgery (add/delete columns, merged-cell-safe reads, append tables), **table auto-fit (AI 表格超页边距修复 via AutoFitBehavior 窗口/内容)**, **default page numbers (页脚 PAGE 域, 小五 9pt Times New Roman)**, image insert/guard, **export PDF**, **TOC refresh + field update**, **per-section headers/footers**, **comments (批注) read/write**, **track changes accept/reject** |
| **Excel** | cell/range read/write (batch), formulas, merge, formatting, row/col insert/delete, **PivotTable**, **charts**, **conditional formatting**, **data validation**, **freeze panes**, **autofilter**, **named ranges**, **sheet→PDF export**, SaveAs format constants |
| **PowerPoint** | slide read/add, textboxes, tables, pictures, **export slides to PNG (visual verification)**, **speaker notes**, **template/theme apply**, **slide master**, **transitions**, SaveAs format constants |
| **通用** | crash recovery (taskkill), post-edit health checks, budget cascade verification, **COM object discovery (runtime property/method enumeration)**, **Dispatch vs DispatchEx vs EnsureDispatch selection**, **dialog-box hang prevention (DispatchEx + DisplayAlerts + AutomationSecurity + is_locked + Open/Close full params + thread timeout kill)**, **gen_py cache corruption recovery**, **multi-thread CoInitialize pattern**, **Excel merged-cell read/write rules**, Chinese 公文 formatting conventions |

## License

MIT
