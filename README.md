# /read-excel — Claude Code Skill

A Claude Code skill that reads and understands Excel (.xlsx) files, especially API specification workbooks with complex formatting.

## What it handles

- **Merged cells** — only the top-left cell of each merge is shown; continuation cells suppressed to prevent header duplication
- **Multiple table blocks per sheet** — blank-row-separated regions rendered as individual markdown tables
- **Section labels** — single-cell merged rows (e.g. "Request Body") rendered as prose before the table they introduce
- **Sparse rows** — 1-cell rows as plain text, 2-cell rows as `**key:** value`
- **Embedded images** — extracted to temp files, then viewed by Claude (mermaid diagrams transcribed, screenshots described)

## Install

**One-line install:**

```bash
mkdir -p ~/.claude/skills/read-excel && curl -sL https://raw.githubusercontent.com/leonatez/claude-skill-read-excel/main/SKILL.md -o ~/.claude/skills/read-excel/SKILL.md
```

**Or clone and run the installer:**

```bash
git clone https://github.com/leonatez/claude-skill-read-excel.git
cd claude-skill-read-excel
bash install.sh
```

## Prerequisites

The skill runs Python with `openpyxl` to parse Excel files:

```bash
pip install openpyxl
```

## Usage

In Claude Code, type:

```
/read-excel /path/to/file.xlsx
```

## What it does (5 steps)

1. **Resolve path** — verifies the file exists
2. **Extract content** — runs a Python script that converts every sheet to structured markdown (merged-cell aware, multi-table detection) and saves embedded images to `/tmp/`
3. **Read images** — Claude views each extracted image and transcribes diagrams, describes screenshots
4. **Classify sheets** — assigns each sheet a kind: `api_spec`, `error_code`, `edge_case`, `mapping`, `flow`, or `metadata`
5. **Structured report** — produces: Document Overview, Sheets, Flows, API Catalogue, Embedded Visuals, Key Patterns

## Uninstall

```bash
rm -rf ~/.claude/skills/read-excel
```

## License

MIT
