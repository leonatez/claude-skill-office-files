---
name: edit-docx
version: 2.1.0
description: |
  Add sections, paragraphs, tables, code blocks, and Mermaid diagrams to Word (.docx)
  files while matching the original document's heading styles, fonts, sizes, indentation,
  colours, and alignment. Also supports tracked changes (redlining) for legal, business,
  and academic documents. Use when asked to "add", "insert", "update", "edit", "redline",
  or "add a diagram / flowchart / sequence diagram".
dependencies:
  - python-docx==1.2.0
  - lxml==6.0.2
allowed-tools:
  - Bash
  - Read
---

You have been invoked to edit a Word (.docx) file. Follow every step below without skipping.
The cardinal rule: **always inspect before you write** — never guess styles.

---

## Workflow Decision Tree

- **Adding new content (sections, paragraphs, tables, images)** → Sections A–C, then **Section D (JSON spec, recommended)**
- **Custom styling that doesn't fit the JSON spec** → Sections A–C, then **Section D-Advanced (raw XML)**
- **Tracking changes for review (redlining)** → Section E (Tracked Changes Workflow)
- **Visual check of output** → Section F
- **Embed a Mermaid diagram as an image** → Section G

---

## Section A — Dependencies

```bash
pip3 install python-docx==1.2.0 lxml==6.0.2 2>/dev/null | grep -E "^(Successfully|Already|Requirement)" || true
```

---

## Section B — Resolve and verify the file

```bash
ls -lh "<FILE_PATH>"
```

---

## Section C — Inspect paragraph styles (ALWAYS before editing)

```bash
FILE_PATH="<FILE_PATH>"
python3 - << 'PYEOF'
import os, subprocess
file_path = subprocess.check_output("echo \"$FILE_PATH\"", shell=True).decode().strip()

from docx import Document
from lxml import etree

doc = Document(file_path)
print(f"Total paragraphs: {len(doc.paragraphs)}\n")

for i, para in enumerate(doc.paragraphs):
    text = para.text.strip()
    style = para.style.name
    pf = para.paragraph_format
    align  = pf.alignment
    indent = pf.left_indent
    run_info = ""
    for run in para.runs:
        if run.text.strip():
            b  = run.bold
            sz = run.font.size
            it = run.italic
            try:    col = str(run.font.color.rgb)
            except: col = "inherit"
            run_info = f"bold={b} size_emu={sz} italic={it} color={col}"
            break
    if text:
        print(f"[{i:3d}] style='{style}' align={align} indent={indent}")
        print(f"       {run_info}")
        print(f"       '{text[:90]}'")

print("\n=== RAW XML of bold/section-header paragraphs ===")
for i, para in enumerate(doc.paragraphs):
    text = para.text.strip()
    if text and any(para.runs) and para.runs[0].bold:
        xml = etree.tostring(para._element, pretty_print=True).decode()
        relevant = [l for l in xml.split('\n') if any(tag in l for tag in
            ['<w:pPr', '<w:ind ', '<w:jc ', '<w:rPr', '<w:b/>', '<w:b ',
             '<w:sz ', '<w:color', '<w:t>', '</w:t', '</w:r>', '</w:pPr'])]
        print(f"\n  Para {i}: '{text[:40]}'")
        for line in relevant[:20]:
            print(f"    {line.strip()}")
        if i > 5:
            break
PYEOF
```

Write down these values before proceeding:
- `w:left` from `<w:ind>` (twips)
- `w:val` from `<w:jc>` ("both"=justify, "l"=left)
- `w:val` from `<w:sz>` (half-points: 20=10pt, 24=12pt)
- Whether `<w:b/>` is present (bold)
- `w:val` from `<w:color>` (hex)

---

## Section D — Write a JSON spec and run docx-apply-json-spec.py (recommended)

This is the token-efficient path. You write a compact JSON spec; the script handles
all OOXML internally — no XML in your output.

### Locate the script

```bash
DOCX_EDIT="$(python3 -c "
import pathlib, sys
candidates = [
    pathlib.Path.home() / '.claude/skills/edit-docx/docx-apply-json-spec.py',
    pathlib.Path('edit-docx/docx-apply-json-spec.py'),
]
found = next((str(p) for p in candidates if p.exists()), None)
print(found or sys.exit('docx-apply-json-spec.py not found — run install.sh first'))
")"
```

### Supported element types

| `type` | Required fields | Optional fields |
|---|---|---|
| `paragraph` | `text` | `bold`, `italic`, `size` (pt), `color` (hex), `font`, `align`, `style` |
| `heading` | `text` | `level` (1–6, default 1) |
| `image` | `path` | `width_inches`, `align` (center/left/right) |
| `table` | `rows` | `headers`, `style` (default "Table Grid") |
| `code` | `lines` | `size` (pt, default 9) — uses Courier New |
| `empty` | — | — |
| `pagebreak` | — | — |

### Write the spec and run

```bash
python3 "$DOCX_EDIT" - << 'JSON'
{
  "file": "<FILE_PATH>",
  "insert_after": 12,
  "elements": [
    {"type": "empty"},
    {"type": "heading", "text": "Process Flow", "level": 1},
    {"type": "paragraph", "text": "See diagram below.", "italic": true},
    {"type": "image", "path": "/tmp/diagram.png", "width_inches": 5.5},
    {"type": "empty"},
    {"type": "table",
     "headers": ["Step", "Actor", "Action"],
     "rows": [
       ["1", "User",  "Submit form"],
       ["2", "API",   "Validate & persist"],
       ["3", "Queue", "Async fulfil"]
     ]
    }
  ]
}
JSON
```

**Insertion actions** (pick one):
- `"insert_after": N` — insert after paragraph N (0-based, from Section C output)
- `"insert_before": N` — insert before paragraph N
- `"append": true` — append to end of document

### Verify

```bash
FILE_PATH="<FILE_PATH>"
python3 - << 'PYEOF'
import subprocess
file_path = subprocess.check_output("echo \"$FILE_PATH\"", shell=True).decode().strip()
from docx import Document
doc = Document(file_path)
print(f"Total paragraphs: {len(doc.paragraphs)}")
for i, p in enumerate(doc.paragraphs[-20:]):
    idx = len(doc.paragraphs) - 20 + i
    print(f"  [{idx}] [{p.style.name}] '{p.text[:70]}'")
PYEOF
```

---

## Section D-Advanced — Write the edit script (raw XML, complex cases only)

Use this only when Section D's element types don't cover your needs (e.g. custom
list styles, complex run formatting, nested tables). Build a Python script that
inserts content using raw XML — the most reliable way to match the original
document's exact style.

```python
from docx import Document
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def make_paragraph(text, bold=False, size_half_pt=20, color_hex=None,
                   left_twips=283, italic=False, align="both",
                   font="Times New Roman"):
    color_xml  = f'<w:color w:val="{color_hex}"/>' if color_hex else ""
    bold_xml   = "<w:b/><w:bCs/>" if bold else ""
    italic_xml = "<w:i/><w:iCs/>" if italic else ""
    return etree.fromstring(f"""<w:p xmlns:w="{W}">
  <w:pPr>
    <w:ind w:left="{left_twips}"/>
    <w:jc w:val="{align}"/>
    <w:rPr>
      {bold_xml}
      <w:sz w:val="{size_half_pt}"/>
      <w:szCs w:val="{size_half_pt}"/>
      {color_xml}
    </w:rPr>
  </w:pPr>
  <w:r>
    <w:rPr>
      {bold_xml}{italic_xml}
      <w:sz w:val="{size_half_pt}"/>
      <w:szCs w:val="{size_half_pt}"/>
      {color_xml}
    </w:rPr>
    <w:t xml:space="preserve">{text}</w:t>
  </w:r>
</w:p>""")


def make_code_paragraph(text, size_half_pt=18, left_twips=283):
    return etree.fromstring(f"""<w:p xmlns:w="{W}">
  <w:pPr>
    <w:ind w:left="{left_twips}"/>
    <w:jc w:val="left"/>
    <w:rPr>
      <w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/>
      <w:sz w:val="{size_half_pt}"/>
      <w:szCs w:val="{size_half_pt}"/>
    </w:rPr>
  </w:pPr>
  <w:r>
    <w:rPr>
      <w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/>
      <w:sz w:val="{size_half_pt}"/>
      <w:szCs w:val="{size_half_pt}"/>
    </w:rPr>
    <w:t xml:space="preserve">{text}</w:t>
  </w:r>
</w:p>""")


def empty_paragraph(left_twips=283, align="both"):
    return etree.fromstring(f"""<w:p xmlns:w="{W}">
  <w:pPr>
    <w:ind w:left="{left_twips}"/>
    <w:jc w:val="{align}"/>
  </w:pPr>
</w:p>""")


def insert_after(body, anchor_para_elem, new_elements):
    idx = list(body).index(anchor_para_elem) + 1
    for elem in new_elements:
        body.insert(idx, elem)
        idx += 1


# Usage
doc = Document(file_path)
body = doc.element.body
anchor = doc.paragraphs[N]._element  # insert after paragraph N

new_elems = [
    empty_paragraph(),
    make_paragraph("NEW SECTION HEADING:", bold=True, size_half_pt=20, left_twips=283),
    make_paragraph("Body text here.", size_half_pt=20, color_hex="212121", left_twips=283),
    empty_paragraph(),
    make_code_paragraph("```mermaid"),
    make_code_paragraph("sequenceDiagram"),
    make_code_paragraph("    A->>B: message"),
    make_code_paragraph("```"),
]

insert_after(body, anchor, new_elems)
doc.save(file_path)
print("Saved.")
```

### XML special character escaping

| Character | Escape |
|-----------|--------|
| `&` | `&amp;` |
| `<` | `&lt;` |
| `>` | `&gt;` |
| `"` | `&quot;` |

### Verify insertion

```bash
FILE_PATH="<FILE_PATH>"
python3 - << 'PYEOF'
import subprocess
file_path = subprocess.check_output("echo \"$FILE_PATH\"", shell=True).decode().strip()
from docx import Document
doc = Document(file_path)
print(f"Total paragraphs: {len(doc.paragraphs)}")
for i, p in enumerate(doc.paragraphs[-30:]):
    idx = len(doc.paragraphs) - 30 + i
    print(f"  [{idx}] '{p.text[:80]}'")
PYEOF
```

---

## Section E — Tracked Changes (Redlining) Workflow

Use this for legal, business, academic, or government documents where reviewers must
accept or reject each change individually.

**Principle: Minimal, precise edits.** Only mark text that actually changes.
Never replace an entire sentence to change one word.

### Example — changing "30 days" to "60 days":

```python
# ❌ BAD — replaces entire sentence, hard to review
'<w:del><w:r><w:delText>The term is 30 days.</w:delText></w:r></w:del>'
'<w:ins><w:r><w:t>The term is 60 days.</w:t></w:r></w:ins>'

# ✅ GOOD — only marks what changed, preserves surrounding runs with original RSID
'<w:r w:rsidR="00AB12CD"><w:t xml:space="preserve">The term is </w:t></w:r>'
'<w:del w:id="1" w:author="Claude" w:date="2026-01-01T00:00:00Z">'
'  <w:r><w:delText>30</w:delText></w:r>'
'</w:del>'
'<w:ins w:id="2" w:author="Claude" w:date="2026-01-01T00:00:00Z">'
'  <w:r><w:t>60</w:t></w:r>'
'</w:ins>'
'<w:r w:rsidR="00AB12CD"><w:t xml:space="preserve"> days.</w:t></w:r>'
```

### Step E1 — Convert to markdown to read the document

```bash
pandoc --track-changes=all "<FILE_PATH>" -o current.md
```

Read `current.md` to understand the full document content and identify all changes needed.

### Step E2 — Unpack the document for XML editing

```bash
SCRIPTS="<path_to_skill>/ooxml/scripts"
python3 "$SCRIPTS/unpack.py" "<FILE_PATH>" unpacked/
# Note the suggested RSID printed — use it for all your w:rsidR values
```

### Step E3 — Plan changes in batches (3–10 changes per batch)

Group related changes. Do NOT use markdown line numbers — they don't map to XML.
Use grep patterns with unique surrounding text to locate changes in XML:

```bash
grep -n "30 days" unpacked/word/document.xml
```

### Step E4 — Implement each batch

For each batch:
1. `grep` for the exact text in `unpacked/word/document.xml` to see how it's split across `<w:r>` elements
2. Write and run a Python script using the OOXML patterns below
3. Verify with `pandoc --track-changes=all` before moving to the next batch

```python
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
RSID = "AABBCCDD"   # Use the RSID suggested by unpack.py
AUTHOR = "Claude"
DATE = "2026-01-01T00:00:00Z"

def make_del(old_text, change_id):
    return etree.fromstring(f"""<w:del xmlns:w="{W}"
        w:id="{change_id}" w:author="{AUTHOR}" w:date="{DATE}">
  <w:r><w:delText xml:space="preserve">{old_text}</w:delText></w:r>
</w:del>""")

def make_ins(new_text, change_id):
    return etree.fromstring(f"""<w:ins xmlns:w="{W}"
        w:id="{change_id}" w:author="{AUTHOR}" w:date="{DATE}">
  <w:r w:rsidR="{RSID}"><w:t xml:space="preserve">{new_text}</w:t></w:r>
</w:ins>""")
```

### Step E5 — Repack and verify

```bash
SCRIPTS="<path_to_skill>/ooxml/scripts"
python3 "$SCRIPTS/pack.py" unpacked/ "<OUTPUT_FILE>"
pandoc --track-changes=all "<OUTPUT_FILE>" -o verification.md
grep "old phrase" verification.md   # should NOT appear
grep "new phrase" verification.md   # should appear
```

---

## Section F — Visual verification (convert to images)

```bash
soffice --headless --convert-to pdf "<FILE_PATH>"
pdftoppm -jpeg -r 150 output.pdf page
# Creates page-1.jpg, page-2.jpg, …
```

Read the images to visually confirm the document looks correct.

---

## Section G — Embed Mermaid Diagram as Image

Use this when the user asks for a flowchart, sequence diagram, ERD, or any diagram
inside a `.docx` document (e.g. a PRD or FRD). The diagram is rendered to PNG first,
then embedded as an inline image at the target paragraph.

### Step G1 — Locate the shared render script

```bash
MERMAID_SCRIPT="$(python3 -c "
import pathlib, sys
candidates = [
    pathlib.Path.home() / '.claude/skills/mermaid/mermaid-render.py',
    pathlib.Path('mermaid/mermaid-render.py'),
]
found = next((str(p) for p in candidates if p.exists()), None)
print(found or sys.exit('mermaid-render.py not found — run install.sh first'))
")"
echo "Script: $MERMAID_SCRIPT"
```

### Step G2 — Write the Mermaid diagram and render to PNG

Write the diagram the user described (or generate an appropriate one), save it to a
temp `.mmd` file, and render:

```bash
cat > /tmp/diagram.mmd << 'MERMAID'
flowchart LR
    A[Start] --> B{Decision}
    B -->|Yes| C[Action]
    B -->|No| D[End]
MERMAID

python3 "$MERMAID_SCRIPT" --input /tmp/diagram.mmd --output /tmp/diagram.png --theme default
ls -lh /tmp/diagram.png
```

Themes: `default` (white bg), `dark`, `forest`, `neutral`. Use `default` for
light-background documents.

### Step G3 — Embed the PNG at the target paragraph

```python
from docx import Document
from docx.shared import Inches
from lxml import etree

doc = Document(file_path)
body = doc.element.body

# Inspect paragraph count if you need to pick the insertion index
print(f"Total paragraphs: {len(doc.paragraphs)}")

# Add the picture paragraph (python-docx appends to end first)
img_para = doc.add_paragraph()
run = img_para.add_run()
run.add_picture("/tmp/diagram.png", width=Inches(5.5))  # adjust width to taste

# Optionally centre the image paragraph
from docx.enum.text import WD_ALIGN_PARAGRAPH
img_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

# Move element to correct position (insert AFTER paragraph N)
anchor = doc.paragraphs[N]._element
body.remove(img_para._element)
idx = list(body).index(anchor) + 1
body.insert(idx, img_para._element)

doc.save(file_path)
print("Diagram embedded.")
```

**Width guidance** — pick based on document margins and diagram complexity:

| Diagram type | Recommended width |
|---|---|
| Simple flowchart / sequence | `Inches(5.5)` |
| Wide ERD or multi-lane swimlane | `Inches(6.5)` |
| Small icon / status diagram | `Inches(3.0)` |

### Step G4 — Verify

Run Section D's verify snippet to confirm paragraph count increased and the image
paragraph appears at the right index, then run Section F to visually confirm.

---

## Common pitfalls

| Symptom | Cause | Fix |
|---------|-------|-----|
| `&` in text breaks XML | Unescaped ampersand | Use `&amp;` |
| New paragraph at wrong position | Wrong `index()` in `insert_after` | Print `list(body).index(anchor)` to verify |
| Style doesn't match | Guessed values | Always run Section C first |
| Bullet style not replicated | List styles need `<w:numPr>` | Copy `<w:pPr>` from an existing list paragraph |
| Vietnamese/special chars show as `?` | Encoding issue | python-docx handles UTF-8; check source file encoding |
| Tracked change shows wrong range | Replaced too much text | Only mark the minimal changed text, reuse surrounding `<w:r>` elements |
