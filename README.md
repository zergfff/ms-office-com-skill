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

- Windows OS
- Microsoft Office installed (Word / Excel / PowerPoint)
- Python with pywin32: `pip install pywin32`

## License

MIT
