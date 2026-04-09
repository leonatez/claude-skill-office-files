---
name: edit-excel
version: 1.0.0
description: |
  Add sheets, write data, and edit existing content in .xlsx files while matching
  the original file's styling (fonts, fills, column widths, borders, alignment).
  Use when asked to "add a sheet", "write data to Excel", "update", or "edit" an xlsx file.
dependencies:
  - openpyxl==3.1.5
allowed-tools:
  - Bash
  - Read
---

You have been invoked to edit an Excel (.xlsx) file. The file path and instruction are
in the skill argument or the user's message. Follow every step below without skipping.

---

## Step 0 — Ensure dependencies

```bash
pip3 install openpyxl==3.1.5 2>/dev/null | grep -E "^(Successfully|Already|Requirement)" || true
```

---

## Step 1 — Resolve and verify the file

Extract the file path from the skill argument or the user's message.

```bash
ls -lh "<FILE_PATH>"
```

If the file does not exist, tell the user and stop.

---

## Step 2 — Inspect existing style (ALWAYS do this before writing)

Run this script to capture the styling patterns used in existing sheets.
This ensures new content is visually consistent with the original.

```bash
FILE_PATH="<FILE_PATH>"
python3 - << 'PYEOF'
import os, subprocess
file_path = subprocess.check_output("echo \"$FILE_PATH\"", shell=True).decode().strip()

import openpyxl
wb = openpyxl.load_workbook(file_path, data_only=True)

print("=== SHEET INVENTORY ===")
for name in wb.sheetnames:
    ws = wb[name]
    rows = sum(1 for r in ws.iter_rows(values_only=True) if any(c is not None for c in r))
    print(f"  [{name}]  {rows} data rows x {ws.max_column} cols")

print("\n=== STYLE SAMPLE (first API/data sheet) ===")
# Sample the first sheet that has data in multiple columns
for name in wb.sheetnames:
    ws = wb[name]
    if (ws.max_column or 0) >= 3 and (ws.max_row or 0) >= 5:
        print(f"\nSheet: {name}")
        for ri, row in enumerate(ws.iter_rows(min_row=1, max_row=12), start=1):
            non_empty = []
            for c in row:
                if c.value:
                    fill = c.fill.fgColor.rgb if c.fill and c.fill.fgColor else 'none'
                    bold = c.font.bold if c.font else False
                    sz   = c.font.size if c.font else None
                    non_empty.append(f"C{c.column}='{str(c.value)[:20]}' b={bold} sz={sz} fill={fill}")
            if non_empty:
                print(f"  R{ri}: {non_empty}")
        print(f"  Col widths: { {k: v.width for k, v in list(ws.column_dimensions.items())[:8]} }")
        break

wb.close()
PYEOF
```

Study the output. Note the fill colours used for section labels, title rows, and comment
columns — you will replicate these exactly in the new content.

---

## Step 3 — Write the edit script

Based on the inspection and the user's instruction, write and run a Python script that:

1. Loads the workbook with `openpyxl.load_workbook(file_path)` (no `data_only=True` — needed for writing)
2. Creates or modifies sheets using the same colours, fonts, column widths, and alignment
3. Freezes panes on new sheets with `ws.freeze_panes = "A2"`
4. Saves the file back to the same path with `wb.save(file_path)`

### Style reference (replicate from existing sheets, never invent new styles):

```python
from openpyxl.styles import Font, PatternFill, Alignment

# Capture from inspection:
LABEL_FILL  = PatternFill("solid", fgColor="9FC5E8")  # section label rows (light blue)
TITLE_FILL  = PatternFill("solid", fgColor="FFFF00")  # sheet title row (yellow)
COMMENT_FILL= PatternFill("solid", fgColor="00FF00")  # comment columns (green)

def label(ws, row, col, text, fill=None, bold=True, sz=11):
    c = ws.cell(row=row, column=col, value=text)
    c.font = Font(bold=bold, size=sz, name="Calibri")
    c.alignment = Alignment(wrap_text=False, vertical="top")
    if fill: c.fill = fill
    return c

def value(ws, row, col, text, bold=False, sz=11, wrap=True, fill=None):
    c = ws.cell(row=row, column=col, value=text)
    c.font = Font(bold=bold, size=sz, name="Calibri")
    c.alignment = Alignment(wrap_text=wrap, vertical="top")
    if fill: c.fill = fill
    return c
```

### Column widths — set these to match the reference sheet:

```python
from openpyxl.utils import get_column_letter
for col, w in enumerate([35, 12, 45, 12, 35, 12, 18, 22, 22], start=1):
    ws.column_dimensions[get_column_letter(col)].width = w
```

---

## Step 4 — Verify the result

After saving, reload and verify:

```bash
FILE_PATH="<FILE_PATH>"
python3 - << 'PYEOF'
import os, subprocess, openpyxl
file_path = subprocess.check_output("echo \"$FILE_PATH\"", shell=True).decode().strip()
wb = openpyxl.load_workbook(file_path, data_only=True)
print(f"Total sheets: {len(wb.sheetnames)}")
for name in wb.sheetnames[-3:]:
    ws = wb[name]
    rows = sum(1 for r in ws.iter_rows(values_only=True) if any(c is not None for c in r))
    print(f"  [{name}]  {rows} data rows x {ws.max_column} cols")
# Spot-check new sheet
last = wb[wb.sheetnames[-1]]
print(f"\nNew sheet '{last.title}' first 5 non-empty rows:")
for ri, row in enumerate(last.iter_rows(min_row=1, max_row=48), start=1):
    vals = [str(c.value)[:25] for c in row if c.value]
    if vals: print(f"  R{ri}: {vals}")
wb.close()
PYEOF
```

Report what was added/changed to the user.

---

## Common pitfalls

| Symptom | Cause | Fix |
|---------|-------|-----|
| File loads but no styles written | Opened with `data_only=True` | Open without `data_only` for writes |
| Merged cells lost on new sheet | Not using `ws.merge_cells()` | Call merge after writing cells |
| Column widths ignored | Using `ws.column_dimensions` before creating sheet | Set widths after `wb.create_sheet()` |
| Old values persist in reloaded sheet | `data_only=True` caches formula results | Use `data_only=True` only for reads |
