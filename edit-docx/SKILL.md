---
name: edit-docx
version: 1.0.0
description: |
  Add sections, paragraphs, tables, and code blocks to Word (.docx) files while
  matching the original document's heading styles, fonts, sizes, indentation, colours,
  and alignment. Use when asked to "add", "insert", "update", or "edit" a Word document.
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

## Step 0 — Ensure dependencies

```bash
pip3 install python-docx==1.2.0 lxml==6.0.2 2>/dev/null | grep -E "^(Successfully|Already|Requirement)" || true
```

---

## Step 1 — Resolve and verify the file

```bash
ls -lh "<FILE_PATH>"
```

---

## Step 2 — Inspect paragraph styles (ALWAYS before editing)

Run this to capture the exact XML properties of paragraphs near the insertion point.
You will replicate these in Step 3.

```bash
FILE_PATH="<FILE_PATH>"
python3 - << 'PYEOF'
import os, subprocess
file_path = subprocess.check_output("echo \"$FILE_PATH\"", shell=True).decode().strip()

from docx import Document
from lxml import etree

doc = Document(file_path)
print(f"Total paragraphs: {len(doc.paragraphs)}\n")

# Print all paragraphs with style info
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
            sz = run.font.size   # in EMU; divide by 12700 for pt
            it = run.italic
            try:    col = str(run.font.color.rgb)
            except: col = "inherit"
            run_info = f"bold={b} size_emu={sz} italic={it} color={col}"
            break

    if text:
        print(f"[{i:3d}] style='{style}' align={align} indent={indent}")
        print(f"       {run_info}")
        print(f"       '{text[:90]}'")

# Show raw XML for key paragraphs (section headers and body text samples)
print("\n=== RAW XML of section-header-style paragraphs ===")
for i, para in enumerate(doc.paragraphs):
    text = para.text.strip()
    if text and any(para.runs) and para.runs[0].bold:
        # Condensed XML: only pPr and first run
        xml = etree.tostring(para._element, pretty_print=True).decode()
        # Extract only the relevant tags
        relevant = [l for l in xml.split('\n') if any(tag in l for tag in
            ['<w:pPr', '<w:ind ', '<w:jc ', '<w:rPr', '<w:b/>', '<w:b ',
             '<w:sz ', '<w:color', '<w:t>', '</w:t', '</w:r>', '</w:pPr'])]
        print(f"\n  Para {i}: '{text[:40]}'")
        for line in relevant[:20]:
            print(f"    {line.strip()}")
        if i > 5:  # enough samples
            break

PYEOF
```

Write down these values before proceeding:
- `w:left` value (twips) from `<w:ind>`
- `w:val` from `<w:jc>` (alignment: "both" = justify, "l" = left)
- `w:val` from `<w:sz>` (half-points, e.g. 20 = 10pt)
- Whether `<w:b/>` is present (bold)
- `w:val` from `<w:color>` (hex, e.g. "212121")

---

## Step 3 — Write the edit script

Build a Python script that inserts new content using raw XML elements — this is the
most reliable way to match the original document's exact style.

### Core pattern

```python
from docx import Document
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def make_paragraph(text, bold=False, size_half_pt=20, color_hex=None,
                   left_twips=283, italic=False, align="both",
                   font="Times New Roman"):
    """Build a <w:p> element matching the document's existing style.
    
    Args:
        size_half_pt: font size in half-points (20 = 10pt, 24 = 12pt)
        left_twips:   left indent in twips (283 ≈ 0.2 in, 456 ≈ 0.32 in)
        align:        "both" (justify), "l" (left), "ctr" (center), "r" (right)
    """
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
    """Monospace paragraph for code blocks (e.g. mermaid diagrams)."""
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
    """Insert a list of <w:p> elements immediately after anchor_para_elem."""
    idx = list(body).index(anchor_para_elem) + 1
    for elem in new_elements:
        body.insert(idx, elem)
        idx += 1
```

### Insertion pattern (inserting after paragraph N)

```python
doc = Document(file_path)
body = doc.element.body
anchor = doc.paragraphs[N]._element   # paragraph to insert after

new_elems = [
    empty_paragraph(),
    make_paragraph("NEW SECTION HEADING:", bold=True, size_half_pt=20, left_twips=283),
    make_paragraph("Body text here.", bold=False, size_half_pt=20,
                   color_hex="212121", left_twips=283),
    # Mermaid / code block:
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

### Important: XML-special characters

Escape these in any text string passed to `make_paragraph` or `make_code_paragraph`:

| Character | Escape    |
|-----------|-----------|
| `&`       | `&amp;`   |
| `<`       | `&lt;`    |
| `>`       | `&gt;`    |
| `"`       | `&quot;`  |

---

## Step 4 — Verify the result

```bash
FILE_PATH="<FILE_PATH>"
python3 - << 'PYEOF'
import subprocess
file_path = subprocess.check_output("echo \"$FILE_PATH\"", shell=True).decode().strip()
from docx import Document
doc = Document(file_path)
print(f"Total paragraphs: {len(doc.paragraphs)}")
# Print last 30 paragraphs to confirm insertion
for i, p in enumerate(doc.paragraphs[-30:]):
    idx = len(doc.paragraphs) - 30 + i
    print(f"  [{idx}] '{p.text[:80]}'")
PYEOF
```

Confirm the new paragraphs appear at the correct position with the correct text.

---

## Common pitfalls

| Symptom | Cause | Fix |
|---------|-------|-----|
| `&` in text breaks XML | Unescaped ampersand | Use `&amp;` in the text string |
| New paragraph appears at wrong position | Wrong `index()` in insert_after | Print `list(body).index(anchor)` to verify |
| Paragraph style doesn't match | Guessed values instead of inspecting | Always run Step 2 first |
| Bullet/list paragraph style not replicated | List styles need `<w:numPr>` XML | Copy the `<w:pPr>` block verbatim from an existing list paragraph's XML |
| Vietnamese/special chars show as `?` | Bytes written without UTF-8 | python-docx handles encoding automatically; ensure source file opened correctly |
