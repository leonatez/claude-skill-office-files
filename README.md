# Office File Skills — Claude Code

Seven Claude Code skills for reading, editing, and processing Office files (`.xlsx`, `.docx`, `.pptx`) and PDFs.
All skills share the same core principle: **inspect first, replicate style exactly**.

## Skills

| Skill | Command | Purpose |
|-------|---------|---------|
| read-excel | `/read-excel` | Read & analyze `.xlsx` — merged cells, multi-table sheets, embedded images, API spec parsing |
| edit-excel | `/edit-excel` | Edit `.xlsx` with formula-first enforcement and LibreOffice formula recalculation |
| read-docx  | `/read-docx`  | Read & analyze `.docx` — pandoc markdown conversion + structured python-docx extraction |
| edit-docx  | `/edit-docx`  | Edit `.docx` with exact style replication + tracked changes / redlining workflow |
| read-pptx  | `/read-pptx`  | Read & analyze `.pptx` — slides, shapes, positions, fonts |
| edit-pptx  | `/edit-pptx`  | Edit `.pptx` + template bulk-replace pipeline (inventory → rearrange → replace → thumbnail) |
| pdf        | `/pdf`        | Extract, merge, split, fill forms, and create PDF files |

---

## Quick Install (new machine)

```bash
git clone https://github.com/leonatez/claude-skill-read-excel.git
cd claude-skill-read-excel
bash install.sh
```

`install.sh` does two things:
1. Installs all Python dependencies (pinned versions from `requirements.txt`)
2. Copies each `SKILL.md` into `~/.claude/skills/<skill-name>/`

### Dependencies (installed automatically by `install.sh`)

```
openpyxl==3.1.5      # xlsx read/write
python-pptx==1.0.2   # pptx read/write
python-docx==1.2.0   # docx read/write
lxml==6.0.2          # XML manipulation
```

To install manually on a new machine:

```bash
pip3 install -r requirements.txt
```

---

## Usage

In Claude Code, type the slash command with the file path:

```
/read-excel /path/to/file.xlsx
/edit-excel /path/to/file.xlsx   add a new sheet called "Summary" with totals
/read-docx  /path/to/file.docx
/edit-docx  /path/to/file.docx   add a new section "Process Flow" with mermaid diagram after paragraph 29
/read-pptx  /path/to/file.pptx
/edit-pptx  /path/to/file.pptx   add a new slide after slide 31 about terminal commands
```

---

## How they work

Every edit skill follows the same three-phase pattern:

```
1. INSPECT  — run a Python script to extract exact shape geometry,
              font sizes (EMU/twips), colours (hex), and spacing values
              from the existing file. Never guess.

2. BUILD    — construct new content using raw XML (lxml), replicating
              the captured values exactly.

3. VERIFY   — reload the file and print a summary of the new content
              to confirm correctness.
```

This approach ensures new content is indistinguishable from the original in terms
of visual style, even for files with complex layouts.

---

## Per-skill details

### `/read-excel`
- Merged-cell-aware markdown renderer (continuation cells suppressed)
- Multi-table-block detection per sheet (blank-row-separated regions)
- Embedded image extraction → Claude views and transcribes diagrams
- Sheet classification: `api_spec`, `error_code`, `edge_case`, `mapping`, `flow`, `metadata`
- Output capped at 150 rows/sheet to prevent context overflow
- Multi-file batch mode with CSV field extraction

### `/edit-excel`
- Inspects fill colours, fonts, column widths from reference sheet before writing
- Provides `label()` and `value()` helpers pre-configured with Calibri font
- Adds `freeze_panes` on new sheets automatically

### `/read-docx`
- Extracts every paragraph with style name, alignment, indent, and run-level formatting
- Shows table contents (first 8 rows per table)
- Reports font sizes in both EMU and points for easy reuse in edit scripts

### `/edit-docx`
- Inspects raw `<w:pPr>` XML of existing paragraphs to capture `w:ind`, `w:jc`, `w:sz`
- Provides `make_paragraph()`, `make_code_paragraph()`, and `empty_paragraph()` helpers
- Inserts at exact position using `body.insert(idx, elem)`
- Documents XML-special-character escaping rules

### `/read-pptx`
- Lists all slides with layout name, shape positions (in inches), and text content
- Captures font sizes, bold, and colour for each run
- Designed to feed directly into `/edit-pptx`

### `/edit-pptx`
- Inspects target slide XML to extract exact EMU values before inserting
- Appends slide then moves `sldIdLst` entry to the correct position
- Provides `make_textbox()` and `make_badge()` helpers covering the common shape types
- Documents the `<a:rPr b="1">` vs `<a:b/>` gotcha

---

## Uninstall

```bash
rm -rf ~/.claude/skills/read-excel \
       ~/.claude/skills/edit-excel \
       ~/.claude/skills/read-docx \
       ~/.claude/skills/edit-docx \
       ~/.claude/skills/read-pptx \
       ~/.claude/skills/edit-pptx
```

## License

MIT
