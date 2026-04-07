---
name: read-excel
version: 1.0.0
description: |
  Read and understand Excel (.xlsx) files — especially API specification workbooks
  with merged cells, multiple table regions per sheet, and embedded images (mermaid
  flow diagrams, screenshots). Extracts every sheet to structured markdown and saves
  embedded images so Claude can describe them. Use when asked to "read", "analyze",
  "understand", or "extract from" an Excel file.
allowed-tools:
  - Bash
  - Read
---

You have been invoked to read and understand an Excel file. The user's file path (or
description of which file) is in the skill argument or their message. Follow every
step below without skipping.

---

## Step 1 — Resolve the file path

Extract the file path from the skill argument or the user's message. If no path was
given, ask: "Which Excel file should I read? Please provide the full path."

Verify the file exists:

```bash
ls -lh "<FILE_PATH>"
```

If it does not exist, tell the user and stop.

---

## Step 2 — Run the extraction script

Execute this Python script. It produces:
- A sheet inventory (name, dimensions, embedded image count)
- Full markdown rendering of each sheet (merged-cell aware, multi-table detection)
- Embedded images saved to a temp directory

```bash
python3 << 'PYEOF' "<FILE_PATH>"
import sys, os, io
import openpyxl

file_path = sys.argv[1]
out_dir = f"/tmp/excel_read_{os.getpid()}"
os.makedirs(out_dir, exist_ok=True)

wb = openpyxl.load_workbook(file_path, data_only=True)

# ── Sheet inventory ────────────────────────────────────────────────
print("=== SHEET INVENTORY ===")
for name in wb.sheetnames:
    ws_ro = openpyxl.load_workbook(file_path, read_only=True, data_only=True)[name]
    rows = sum(1 for r in ws_ro.iter_rows(values_only=True) if any(c is not None for c in r))
    ws = wb[name]
    imgs = len(getattr(ws, "_images", []))
    print(f"  [{name}]  {rows} rows  ×  {ws.max_column or 0} cols  |  {imgs} embedded image(s)")
print()


# ── Merged-cell-aware markdown renderer ───────────────────────────
def sheet_to_markdown(ws) -> str:
    # Build merged-cell maps
    merge_tl: dict = {}          # (row, col) → value at top-left of merge
    continuations: set = set()   # cells that are overflow of a merge (show as empty)
    for merge in ws.merged_cells.ranges:
        tl_val = ws.cell(merge.min_row, merge.min_col).value
        merge_tl[(merge.min_row, merge.min_col)] = tl_val
        for r in range(merge.min_row, merge.max_row + 1):
            for c in range(merge.min_col, merge.max_col + 1):
                if not (r == merge.min_row and c == merge.min_col):
                    continuations.add((r, c))

    max_row = ws.max_row or 0
    max_col = ws.max_column or 0

    def get_val(r, c):
        if (r, c) in continuations:
            return None
        return merge_tl.get((r, c), ws.cell(r, c).value)

    def cs(v) -> str:
        if v is None:
            return ""
        return str(v).strip().replace("\n", " ")

    # Build 2-D grid
    grid = [[cs(get_val(r, c)) for c in range(1, max_col + 1)]
            for r in range(1, max_row + 1)]
    if not grid:
        return "(empty)"

    def row_empty(row):
        return all(v == "" for v in row)

    output = []
    i = 0
    while i < len(grid):
        # Skip empty rows
        if row_empty(grid[i]):
            i += 1
            continue

        # Collect contiguous non-empty block
        block_start = i
        while i < len(grid) and not row_empty(grid[i]):
            i += 1
        block = grid[block_start:i]

        # Columns used in this block
        used_cols = [c for c in range(max_col) if any(row[c] != "" for row in block)]
        if not used_cols:
            continue

        # Find first row with ≥ 2 non-empty cells → use as table header.
        # This skips single-cell section labels (e.g. "Request Body") that
        # precede the actual field table.
        header_idx = next(
            (idx for idx, row in enumerate(block)
             if sum(1 for c in used_cols if row[c] != "") >= 2),
            None,
        )

        if header_idx is not None and len(block) > header_idx + 1:
            # Render any prefix rows (section labels) as prose
            for row in block[:header_idx]:
                ne = [row[c] for c in used_cols if row[c] != ""]
                if len(ne) == 1:
                    output.append(f"**{ne[0]}**")
                elif ne:
                    output.append("  ".join(ne))

            # Render table from detected header row onward
            hdrs = [block[header_idx][c] for c in used_cols]
            output.append("| " + " | ".join(hdrs) + " |")
            output.append("| " + " | ".join(["---"] * len(used_cols)) + " |")
            for row in block[header_idx + 1:]:
                cells = [row[c] if c < len(row) else "" for c in used_cols]
                output.append("| " + " | ".join(cells) + " |")
            output.append("")

        else:
            # Sparse block — render as prose / key-value pairs
            for row in block:
                ne = [v for v in row if v != ""]
                if len(ne) == 1:
                    output.append(ne[0])
                elif len(ne) == 2:
                    output.append(f"**{ne[0]}:** {ne[1]}")
                elif ne:
                    output.append("  ".join(ne))
            output.append("")

    return "\n".join(output)


# ── Render each sheet ─────────────────────────────────────────────
print("=== SHEET CONTENTS ===")
for name in wb.sheetnames:
    ws = wb[name]
    print(f"\n## Sheet: {name}\n")
    print(sheet_to_markdown(ws))
    print("\n---")


# ── Extract embedded images ───────────────────────────────────────
print("\n=== EMBEDDED IMAGES ===")
image_paths = []
for name in wb.sheetnames:
    ws = wb[name]
    for idx, img in enumerate(getattr(ws, "_images", [])):
        try:
            img_bytes = img._data()
            # Detect format from magic bytes
            if img_bytes[:4] == b'\x89PNG':
                ext = "png"
            elif img_bytes[:2] == b'\xff\xd8':
                ext = "jpg"
            elif img_bytes[:4] == b'GIF8':
                ext = "gif"
            else:
                ext = "png"
            img_path = os.path.join(out_dir, f"{name}_img{idx + 1}.{ext}")
            with open(img_path, "wb") as f:
                f.write(img_bytes)
            image_paths.append(img_path)
            print(f"  SAVED  sheet='{name}'  index={idx + 1}  size={len(img_bytes):,}B  path={img_path}")
        except Exception as e:
            print(f"  FAILED sheet='{name}'  index={idx + 1}  error={e}")

if not image_paths:
    print("  (none found)")

wb.close()
print(f"\nImages directory: {out_dir}")
print("=== EXTRACTION COMPLETE ===")
PYEOF
```

> **Note on merged cells**: openpyxl's `data_only=True` resolves formula results.
> Merged cell continuations are suppressed (only the top-left cell shows the value)
> so headers don't repeat across every spanned column/row.

---

## Step 3 — Read every embedded image

For each `SAVED` line in the output, use the **Read tool** to open the image path printed.

When you view each image:
- If it is a **sequence / flow diagram**: transcribe the flow as numbered steps
  (actor → actor: action). Note API names, conditions, arrows.
- If it is a **mermaid diagram rendered as PNG**: reconstruct the mermaid source if
  possible, or describe each node and edge in plain text.
- If it is a **table screenshot**: note key column headers and a few representative rows.
- If it is a **UML or swimlane diagram**: identify actors, lanes, and the sequence of
  interactions.

---

## Step 4 — Classify each sheet

Based on the content extracted, classify each sheet using these categories (same as the
unapi parsing pipeline):

| Kind         | Signs                                                                 |
|--------------|-----------------------------------------------------------------------|
| `api_spec`   | Tables with URL/path, request fields, response fields, HTTP method   |
| `error_code` | Tables of result/error codes with descriptions or HTTP status        |
| `edge_case`  | Conditions, retry logic, timeout handling, fallback flows            |
| `mapping`    | Lookup tables, code-to-description mappings, enum value lists        |
| `flow`       | Sequence overview, step-by-step call order, mermaid/swimlane diagrams|
| `metadata`   | Version, changelog, environment URLs, overview/intro text            |

List each sheet with its classified kind and a one-sentence description.

---

## Step 5 — Synthesize a structured summary

Produce this report:

### Document Overview
- **File**: `<filename>`
- **Purpose**: what system/integration this describes
- **Partner / system names** found in headers or URL paths

### Sheets
For each sheet: `[kind]` — one sentence description

### Flows Identified
List each distinct flow name (often corresponds to a business process like
"Disbursement", "Repayment", "Inquiry").

### APIs Catalogue
For each API found across all sheets:
```
Method  Path                     Sheet              Purpose (one line)
POST    /v1/order/create         API orderCreation  Initiates a new loan order
...
```

### Embedded Visuals
For each image: sheet name, what the diagram shows, key information gleaned.

### Key Patterns
- **Authentication**: token type, header name, signing scheme
- **Error conventions**: code format, where error tables live
- **Shared fields**: fields repeated across multiple APIs (e.g. `partnerRefId`)
- **Gotchas / edge cases**: anything unusual that affects integration

---

## Interpretation guide for API spec Excel files

Keep these heuristics in mind while reading:

**Merged cells as section labels** — A row with a single merged cell spanning all
columns is a section divider (e.g. "Request Body", "Response Header"). The table
immediately below it belongs to that section.

**Multiple table blocks per sheet** — One sheet often contains the full spec for one
API: a URL/method block, then a request header table, then a request body table, then
a response table, all separated by blank rows. Treat each contiguous block separately.

**Implicit required/optional** — If there is no "Required" column, look for bold or
colored cells, asterisk markers, or infer from context (primary key fields are almost
always required).

**Flow sheet vs. API sheet** — A sheet with mostly a diagram image and few text cells
is a flow/sequence overview. A sheet with dense tables of field names is an API spec
sheet.

**Embedded images are often the most important content** — In API spec Excel files,
mermaid/sequence diagrams embedded as images show the call ordering that is impossible
to infer from individual API sheets alone. Always describe them in detail.
