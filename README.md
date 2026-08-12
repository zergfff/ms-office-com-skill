# word-com-chinese-skill

Agent skill for **MS Word COM automation on Windows** — 专用于中文公文/报告文档（.docx/.doc）的生成与修改，通过 `win32com.client` (pywin32) 驱动真实 Word。

> ⚠️ **SCOPE: 本 skill 仅适用于 Windows + Microsoft Word + COM (win32com.client)，专用于中文文档（公文、报告、方案）。**
> ❌ 不适用于 Linux/macOS、WPS Office、Excel/PowerPoint、python-docx 纯库路线、Microsoft Graph 云端 API。
> 其他组合请勿使用本 skill。

Designed for AI agents that edit real Chinese government documents (公文): Hermes, Claude Code, Codex, OpenClaw, Cursor, etc.

## Why COM?

Pure-library approaches (python-docx) flatten or lose complex structure:
merged tables, styles, outline levels, headers/footers, fields, images anchored in paragraphs.
The COM interface drives the real Word application and preserves everything — this skill is
built for that path.

## Contents

- `SKILL.md` — the skill itself (golden rules, Word core patterns, find/replace danger zone, styles & outline, fonts & substitution, GB/T 9704-2012 公文格式, alignment & table cell defaults, GB/T 7714 references, superscript/subscript semantics, table surgery & auto-fit, page numbers, PDF export with heading bookmarks, dialog-hang prevention, dispatch selection, health checks, budget cascades)
- `references/win32com-cheatsheet.md` — verified working snippets for Word (safe bulk replace, style constants, table column add/delete, merged-cell reads, image insert/resize, fonts, PDF export, page numbers, references, superscript/subscript)
- `references/pitfalls.md` — real incident transcripts (infinite TOC duplication, stale-process corruption, long-string Find error, image-container paragraph loss, style off-by-one, LLM-generation pitfalls A1-A10, file-corruption pitfalls B1-B3)
- `scripts/ensure_fonts.py` — 检测本机是否缺少公文字体（仿宋_GB2312/楷体_GB2312/方正小标宋简体等），缺失则自动下载安装（用户级，无需管理员）

## Install

### Hermes Agent

```bash
# copy into a profile's skills dir
cp -r word-com-chinese-skill ~/AppData/Local/hermes/profiles/<profile>/skills/productivity/word-com-chinese-skill
# or use the skills hub / curator
```

### Claude Code

```bash
mkdir -p ~/.claude/skills && cp -r word-com-chinese-skill ~/.claude/skills/
```

### Codex (OpenAI)

```bash
mkdir -p ~/.codex/skills && cp -r word-com-chinese-skill ~/.codex/skills/
```

### OpenClaw / Cursor / others

Any agent that supports Anthropic-style SKILL.md with YAML frontmatter (`name`, `description`).
Copy the folder into the agent's skills directory.

## Requirements

- **Windows 11 + Microsoft Word 2024 (LTSC)** — primary target; works on Word 2016+ (COM object model stable across these versions)
- Microsoft Word installed
- Python with pywin32: `pip install pywin32`

## Feature Map

| 模块 | 能力 |
|---|---|
| **格式** | GB/T 9704-2012 公文格式速查：标题 方正小标宋简体2号居中 / 一级标题 黑体3号 / 二级标题 楷体_GB2312 3号 / 正文 仿宋_GB2312 3号两端对齐首行缩进2字符行距28磅；字体双属性（中文字体 NameFarEast + 西文字体 Name）；缺字自动回退（SubstituteFont）；标题/表题/图题居中无缩进；图片段单倍行距；表格内文字上下/左右居中+无缩进+单倍行距；表头加粗+跨页重复标题行 |
| **内容** | GB/T 7714-2015 参考文献（顺序编码制、各文献类型、正文上标[1]引用）；上下角标语义判定（m3→m³上标、CO2→CO₂下标、Ca2+→Ca²⁺上标、数值不动） |
| **表格** | 表格手术（加列/删空列/合并单元格安全读写/尾部追加表）；AI 表格超页边距修复（AutoFitBehavior 窗口/内容）；表头重复 |
| **交付** | 页脚页码（PAGE 域，小五 9pt Times New Roman）；PDF 导出（CreateBookmarks=1 按标题建书签，先刷新TOC/域）；目录自动刷新 |
| **安全** | 防弹窗卡死（DispatchEx + DisplayAlerts + AutomationSecurity + is_locked + Open/Close 全参数 + 线程超时强杀）；崩溃恢复（taskkill）；每次编辑后健康检查；预算级联核对 |
| **字体** | `scripts/ensure_fonts.py` 自动检测+下载安装缺失的公文字体（仿宋_GB2312/楷体_GB2312/方正小标宋简体/方正楷体_GBK/黑体等，用户级安装无需管理员）；**GitHub 源快速失败策略（连接10s/总限时30s 超时即跳过，国内被墙不傻等；失败可设 HTTPS_PROXY 代理重试）** |

## License

MIT
