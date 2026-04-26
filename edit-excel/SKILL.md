---
name: edit-excel
version: 2.0.0
description: |
  Add sheets, write data, and edit existing content in .xlsx files while matching
  the original file's styling (fonts, fills, column widths, borders, alignment).
  Enforces formula-first best practices and recalculates formulas via LibreOffice.
  Use when asked to "add a sheet", "write data to Excel", "update", or "edit" an xlsx file.
dependencies:
  - openpyxl==3.1.5
  - pandas
allowed-tools:
  - Bash
  - Read
---

You have been invoked to edit an Excel (.xlsx) file. Follow every step below without skipping.

---

## Step 0 — Ensure dependencies

```bash
pip3 install openpyxl==3.1.5 pandas 2>/dev/null | grep -E "^(Successfully|Already|Requirement)" || true
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

print("\n=== STYLE SAMPLE (first data sheet) ===")
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

Study the output. Note fill colours, font sizes, and column widths for replication.

---

## Step 3 — Write the edit script

### CRITICAL: Use Excel formulas, not hardcoded Python values

**Always let Excel calculate; never hardcode computed results.**

```python
# ❌ WRONG — hardcodes the result, file won't recalculate when data changes
sheet['B10'] = df['Sales'].sum()

# ✅ CORRECT — stays dynamic
sheet['B10'] = '=SUM(B2:B9)'
```

This applies to all totals, percentages, ratios, averages, and growth rates.

### Common workflow

```python
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# Load existing file (no data_only — needed for writing)
wb = load_workbook(file_path)
sheet = wb.active  # or wb['SheetName']

# Write formula
sheet['B10'] = '=SUM(B2:B9)'

# Style helpers — replicate fill/font values from inspection output
LABEL_FILL  = PatternFill("solid", fgColor="9FC5E8")  # section labels
TITLE_FILL  = PatternFill("solid", fgColor="FFFF00")  # title rows
COMMENT_FILL= PatternFill("solid", fgColor="00FF00")  # comment columns

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

# Column widths
for col, w in enumerate([35, 12, 45, 12, 35, 12, 18, 22, 22], start=1):
    sheet.column_dimensions[get_column_letter(col)].width = w

# Freeze header row
sheet.freeze_panes = "A2"

wb.save(file_path)
print("Saved.")
```

### Financial model colour coding (use when building financial models)

Unless the file already has established conventions:

| Text colour | Meaning |
|---|---|
| Blue (0,0,255) | Hardcoded inputs users will change |
| Black (0,0,0) | All formulas and calculations |
| Green (0,128,0) | Links from other sheets in this workbook |
| Red (255,0,0) | Links to other files |

| Background | Meaning |
|---|---|
| Yellow | Key assumptions needing attention |

Number format rules:
- Years as text strings ("2024", not 2024)
- Currency: `$#,##0` with units in header ("Revenue ($mm)")
- Zeros display as "-" using `$#,##0;($#,##0);-`
- Negatives in parentheses: (123) not -123

---

## Step 4 — Recalculate formulas (MANDATORY when formulas are used)

`openpyxl` writes formulas as strings — values will be blank or stale until recalculated.
Run the bundled `recalc.py` script to evaluate all formulas via LibreOffice:

```bash
# Locate the script (installed alongside this skill)
SKILL_DIR="$(python3 -c "import pathlib; print(next(p for p in [pathlib.Path.home()/'.claude/skills/edit-excel', pathlib.Path('edit-excel')] if (p/'recalc.py').exists()))")"
python3 "$SKILL_DIR/recalc.py" "<FILE_PATH>"
```

Interpret the JSON output:

```json
{
  "status": "success",
  "total_formulas": 42,
  "total_errors": 0
}
```

If `status` is `errors_found`, check `error_summary` for locations and fix them:

| Error | Cause | Fix |
|---|---|---|
| `#REF!` | Deleted row/col referenced | Update cell references |
| `#DIV/0!` | Division by zero | Add `IF(B2=0,"",A2/B2)` guard |
| `#VALUE!` | Wrong data type | Check formula inputs are numeric |
| `#NAME?` | Misspelled function | Check function name spelling |

---

## Step 5 — Verify the result

```bash
FILE_PATH="<FILE_PATH>"
python3 - << 'PYEOF'
import subprocess, openpyxl
file_path = subprocess.check_output("echo \"$FILE_PATH\"", shell=True).decode().strip()
wb = openpyxl.load_workbook(file_path, data_only=True)
print(f"Total sheets: {len(wb.sheetnames)}")
for name in wb.sheetnames[-3:]:
    ws = wb[name]
    rows = sum(1 for r in ws.iter_rows(values_only=True) if any(c is not None for c in r))
    print(f"  [{name}]  {rows} data rows x {ws.max_column} cols")
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

## Formula verification checklist

- [ ] Test 2–3 sample references before building the full model
- [ ] Column mapping: confirm Excel columns match (column 64 = BL, not BK)
- [ ] Row offset: Excel rows are 1-indexed (DataFrame row 5 = Excel row 6)
- [ ] NaN handling: check nulls with `pd.notna()`
- [ ] Division by zero: check denominators before using `/`
- [ ] Cross-sheet references: use correct format `Sheet1!A1`

---

## Common pitfalls

| Symptom | Cause | Fix |
|---------|-------|-----|
| File loads but no styles written | Opened with `data_only=True` | Open without `data_only` for writes |
| Merged cells lost on new sheet | Not using `ws.merge_cells()` | Call merge after writing cells |
| Column widths ignored | Setting before `create_sheet()` | Set widths after creating the sheet |
| Old values persist on reload | `data_only=True` caches stale results | Only use `data_only=True` for reads |
| Formulas show as strings in Excel | Not recalculated after save | Run `recalc.py` after saving |
