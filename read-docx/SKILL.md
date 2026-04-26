---
name: read-docx
version: 2.0.0
description: |
  Read and understand Word (.docx) files. Extracts all paragraphs, headings, tables,
  run-level formatting (font, size, bold, colour, italic), and document structure.
  Supports both structured extraction (python-docx) and full markdown conversion (pandoc).
  Produces a structured markdown summary. Use when asked to "read", "analyze",
  "understand", or "summarize" a Word document.
dependencies:
  - python-docx==1.2.0
  - lxml==6.0.2
allowed-tools:
  - Bash
  - Read
---

You have been invoked to read and understand a Word (.docx) file. Follow every step below.

---

## Step 0 — Ensure dependencies

```bash
pip3 install python-docx==1.2.0 2>/dev/null | grep -E "^(Successfully|Already|Requirement)" || true
```

---

## Step 1 — Resolve and verify the file

Extract the file path from the skill argument or the user's message.

```bash
ls -lh "<FILE_PATH>"
```

---

## Step 2 — Choose extraction method

### Option A: Full text with pandoc (best for content understanding)

If `pandoc` is available, use it — it produces clean markdown preserving structure,
lists, tables, and even tracked changes:

```bash
# Check if pandoc is available
pandoc --version 2>/dev/null && echo "pandoc available" || echo "pandoc not found"

# Convert to markdown (shows accepted state of tracked changes)
pandoc "<FILE_PATH>" -o /tmp/doc_content.md
cat /tmp/doc_content.md

# To see ALL tracked changes (insertions and deletions):
pandoc --track-changes=all "<FILE_PATH>" -o /tmp/doc_changes.md
cat /tmp/doc_changes.md
```

### Option B: Structured extraction with python-docx (best for formatting details)

Use this when you need exact style names, font sizes, colours, and indentation values
(e.g., before editing with `/edit-docx`):

```bash
FILE_PATH="<FILE_PATH>"
python3 - << 'PYEOF'
import os, subprocess
file_path = subprocess.check_output("echo \"$FILE_PATH\"", shell=True).decode().strip()

from docx import Document

doc = Document(file_path)
print(f"=== DOCUMENT OVERVIEW ===")
print(f"Paragraphs : {len(doc.paragraphs)}")
print(f"Tables     : {len(doc.tables)}")
print(f"Sections   : {len(doc.sections)}")
print()

# Paragraph inventory
print("=== PARAGRAPHS ===")
for i, para in enumerate(doc.paragraphs):
    text = para.text.strip()
    if not text:
        continue
    style  = para.style.name
    align  = para.paragraph_format.alignment
    indent = para.paragraph_format.left_indent

    run_info = ""
    for run in para.runs:
        if run.text.strip():
            b  = run.bold
            sz = run.font.size
            it = run.italic
            try:
                col = run.font.color.rgb
            except:
                col = None
            run_info = f"bold={b} size={sz} italic={it} color={col}"
            break

    print(f"[{i:3d}] style='{style}' | {run_info}")
    print(f"       text='{text[:100]}'")

# Table inventory
print("\n=== TABLES ===")
for ti, table in enumerate(doc.tables):
    print(f"\n--- Table {ti} ({len(table.rows)} rows x {len(table.columns)} cols) ---")
    for ri, row in enumerate(table.rows[:8]):
        cells = [cell.text.strip()[:35] for cell in row.cells]
        print(f"  Row {ri}: {cells}")
    if len(table.rows) > 8:
        print(f"  ... ({len(table.rows) - 8} more rows)")

print("\n=== EXTRACTION COMPLETE ===")
PYEOF
```

---

## Step 3 — Synthesize a structured summary

From the extraction output, produce this report:

### Document Overview
- **File**: filename
- **Language**: detected language(s)
- **Purpose**: what this document is about (one sentence)

### Structure
List the logical sections found (headings, labelled section breaks) in order.

### Key Content
Summarize the most important information in each section.

### Tables
For each table: describe its purpose and list the column headers.

### Formatting Conventions Used
Note: font size(s), heading styles, colour usage, indent patterns — this is essential
context if the document will later be edited with `/edit-docx`.

---

## Common pitfalls

| Symptom | Cause | Fix |
|---------|-------|-----|
| Vietnamese / non-ASCII text garbled | Terminal encoding | The extraction script handles UTF-8 natively |
| Paragraphs missing | Paragraphs inside table cells are not in `doc.paragraphs` | Iterate `table.rows[i].cells[j].paragraphs` separately |
| Run font returns None | Style inherited from paragraph/document default | `None` means "inherit" — check paragraph style for the effective value |
| Tracked changes not visible | pandoc shows accepted state by default | Use `--track-changes=all` to see insertions/deletions |
