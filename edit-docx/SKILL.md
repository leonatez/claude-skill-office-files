---
name: edit-docx
version: 2.0.0
description: |
  Add sections, paragraphs, tables, and code blocks to Word (.docx) files while
  matching the original document's heading styles, fonts, sizes, indentation, colours,
  and alignment. Also supports tracked changes (redlining) for legal, business, and
  academic documents. Use when asked to "add", "insert", "update", "edit", or "redline".
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

- **Adding new content (sections, paragraphs, tables)** → Sections A–D below
- **Tracking changes for review (redlining)** → Section E (Tracked Changes Workflow)
- **Visual check of output** → Section F

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

## Section D — Write the edit script

Build a Python script that inserts content using raw XML — the most reliable way to
match the original document's exact style.

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

## Common pitfalls

| Symptom | Cause | Fix |
|---------|-------|-----|
| `&` in text breaks XML | Unescaped ampersand | Use `&amp;` |
| New paragraph at wrong position | Wrong `index()` in `insert_after` | Print `list(body).index(anchor)` to verify |
| Style doesn't match | Guessed values | Always run Section C first |
| Bullet style not replicated | List styles need `<w:numPr>` | Copy `<w:pPr>` from an existing list paragraph |
| Vietnamese/special chars show as `?` | Encoding issue | python-docx handles UTF-8; check source file encoding |
| Tracked change shows wrong range | Replaced too much text | Only mark the minimal changed text, reuse surrounding `<w:r>` elements |
