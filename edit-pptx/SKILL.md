---
name: edit-pptx
version: 2.0.0
description: |
  Add, insert, or edit slides in PowerPoint (.pptx) files matching the original
  presentation's layout, shape positions, font sizes, colours, and formatting.
  Also supports template-based bulk generation: inventory → replace → rearrange pipeline.
  Use when asked to "add a slide", "insert", "update", "edit", or "generate from template".
dependencies:
  - python-pptx==1.0.2
  - lxml==6.0.2
  - pillow
allowed-tools:
  - Bash
  - Read
---

You have been invoked to edit a PowerPoint (.pptx) file. Follow every step without skipping.
The cardinal rule: **always inspect before you write** — never guess positions or colours.

---

## Workflow Decision Tree

- **Add/edit a few slides in an existing file** → use Sections A–D below
- **Populate an existing template with new content** → use Section E (Template Workflow)
- **Visual thumbnail overview of any presentation** → use Section F

---

## Section A — Dependencies

```bash
pip3 install python-pptx==1.0.2 lxml==6.0.2 pillow 2>/dev/null | grep -E "^(Successfully|Already|Requirement)" || true
```

---

## Section B — Resolve and verify the file

```bash
ls -lh "<FILE_PATH>"
```

---

## Section C — Inspect slides near the insertion point (ALWAYS do this first)

```bash
FILE_PATH="<FILE_PATH>"
TARGET_SLIDE=1   # slide to insert AFTER (1-based)
python3 - << 'PYEOF'
import os, subprocess
file_path = subprocess.check_output("echo \"$FILE_PATH\"", shell=True).decode().strip()
target = int(subprocess.check_output("echo \"$TARGET_SLIDE\"", shell=True).decode().strip())

from pptx import Presentation
from lxml import etree

prs = Presentation(file_path)
slides = list(prs.slides)
print(f"Total slides: {len(slides)}")
print(f"Slide size  : {prs.slide_width.emu} x {prs.slide_height.emu} EMU")
print()

for si in [target - 1, target, target + 1]:
    if 0 <= si - 1 < len(slides):
        slide = slides[si - 1]
        print(f"=== Slide {si} | Layout: {slide.slide_layout.name} ===")
        for shape in slide.shapes:
            l = shape.left  or 0; t = shape.top    or 0
            w = shape.width or 0; h = shape.height or 0
            print(f"  '{shape.name}': left={l} top={t} width={w} height={h}")
            try:
                fill = shape.fill
                if str(fill.type) == "SOLID (1)":
                    print(f"    fill_color={fill.fore_color.rgb}")
            except: pass
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    txt = para.text.strip()
                    if txt:
                        for run in para.runs:
                            if run.text.strip():
                                try:    col = str(run.font.color.rgb)
                                except: col = "inherit"
                                print(f"    TEXT: '{txt[:60]}' sz={run.font.size} bold={run.font.bold} color={col}")
                                break
        print()

slide = slides[target - 1]
print(f"=== XML of shapes in slide {target} ===")
for shape in slide.shapes:
    if shape.has_text_frame and shape.text_frame.text.strip():
        xml = etree.tostring(shape._element, pretty_print=True).decode()
        relevant = [l for l in xml.split('\n') if any(tag in l for tag in
            ['<a:off', '<a:ext', '<a:solidFill', '<a:srgbClr', '<a:schemeClr',
             '<a:bodyPr', '<a:pPr', '<a:rPr', '<a:t>', 'val=',
             '<a:lnSpc', '<a:spcBef', '<a:spcAft', '<a:buNone', '<a:ind'])]
        print(f"\n  Shape: '{shape.name}'")
        for line in relevant[:35]:
            print(f"    {line.strip()}")
PYEOF
```

Record EMU values for all key shapes. Use them in Section D.

---

## Section D — Write the edit script

### Insert a slide at a specific position

```python
from pptx import Presentation
from lxml import etree

prs = Presentation(file_path)
slides = list(prs.slides)
INSERT_AFTER = 5  # 1-based

layout = slides[INSERT_AFTER - 1].slide_layout
new_slide = prs.slides.add_slide(layout)

# Clear auto-generated placeholder shapes
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
sp_tree = new_slide.shapes._spTree
for sp in sp_tree.findall(f".//{{{P}}}sp"):
    sp_tree.remove(sp)

# ... add shapes below ...

# Move slide to correct position
sldIdLst = prs.slides._sldIdLst
entry = list(sldIdLst)[-1]
sldIdLst.remove(entry)
sldIdLst.insert(INSERT_AFTER, entry)

prs.save(file_path)
```

### Shape helpers (use EMU values from Section C)

```python
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"

def make_textbox(shape_id, name, left, top, width, height,
                 paragraphs, is_title=False, body_anchor="t"):
    """
    paragraphs: list of dicts — keys: text, bold, size_pt (half-pts e.g. 3500=35pt),
                color (hex or None), italic, indent_emu, line_spc (pct*1000), spc_before (pts*100)
    """
    ph_xml = '<p:nvPr><p:ph type="title"/></p:nvPr>' if is_title else '<p:nvPr/>'
    para_xml = ""
    for p in paragraphs:
        b      = "<a:b/>" if p.get("bold") else ""
        i_     = "<a:i/>" if p.get("italic") else ""
        col    = p.get("color")
        color_xml = f'<a:solidFill><a:srgbClr val="{col}"/></a:solidFill>' if col else \
                    '<a:solidFill><a:schemeClr val="dk1"/></a:solidFill>'
        sz    = p.get("size_pt", 3500)
        indent = p.get("indent_emu", 0)
        lspc  = p.get("line_spc", 115000)
        spbef = p.get("spc_before", 1200)
        para_xml += f"""
    <a:p>
      <a:pPr indent="{indent}" lvl="0" marL="0" rtl="0" algn="l">
        <a:lnSpc><a:spcPct val="{lspc}"/></a:lnSpc>
        <a:spcBef><a:spcPts val="{spbef}"/></a:spcBef>
        <a:spcAft><a:spcPts val="0"/></a:spcAft>
        <a:buNone/>
      </a:pPr>
      <a:r>
        <a:rPr {b} lang="en-US" sz="{sz}">{color_xml}</a:rPr>
        <a:t>{p["text"]}</a:t>
      </a:r>
    </a:p>"""
    return etree.fromstring(f"""<p:sp xmlns:p="{P}" xmlns:a="{A}"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:nvSpPr>
    <p:cNvPr id="{shape_id}" name="{name}"/>
    <p:cNvSpPr txBox="1"/>
    {ph_xml}
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="{left}" y="{top}"/><a:ext cx="{width}" cy="{height}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:noFill/><a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody>
    <a:bodyPr anchorCtr="0" anchor="{body_anchor}" bIns="0" lIns="0"
              spcFirstLastPara="1" rIns="0" wrap="square" tIns="0">
      <a:noAutofit/>
    </a:bodyPr>
    <a:lstStyle/>{para_xml}
  </p:txBody>
</p:sp>""")


def make_badge(shape_id, name, left, top, width, height,
               label, fill_color, font_color="FFFFFF", sz=2200):
    """Solid-fill rectangular badge."""
    return etree.fromstring(f"""<p:sp xmlns:p="{P}" xmlns:a="{A}"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:nvSpPr>
    <p:cNvPr id="{shape_id}" name="{name}"/>
    <p:cNvSpPr txBox="1"/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="{left}" y="{top}"/><a:ext cx="{width}" cy="{height}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="{fill_color}"/></a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody>
    <a:bodyPr anchorCtr="0" anchor="ctr" bIns="91425" lIns="91425"
              spcFirstLastPara="1" rIns="91425" wrap="square" tIns="91425">
      <a:noAutofit/>
    </a:bodyPr>
    <a:lstStyle/>
    <a:p>
      <a:pPr algn="ctr"/>
      <a:r>
        <a:rPr b="1" lang="en-US" sz="{sz}">
          <a:solidFill><a:srgbClr val="{font_color}"/></a:solidFill>
        </a:rPr>
        <a:t>{label}</a:t>
      </a:r>
    </a:p>
  </p:txBody>
</p:sp>""")
```

### Verify after editing

```bash
FILE_PATH="<FILE_PATH>"
TARGET_SLIDE=5
python3 - << 'PYEOF'
import subprocess
file_path = subprocess.check_output("echo \"$FILE_PATH\"", shell=True).decode().strip()
target = int(subprocess.check_output("echo \"$TARGET_SLIDE\"", shell=True).decode().strip())
from pptx import Presentation
prs = Presentation(file_path)
slides = list(prs.slides)
print(f"Total slides: {len(slides)}")
slide = slides[target]
print(f"Slide {target + 1} layout: {slide.slide_layout.name}")
for shape in slide.shapes:
    if shape.has_text_frame:
        print(f"  Shape '{shape.name}': '{shape.text_frame.text[:80]}'")
PYEOF
```

---

## Section E — Template Workflow (bulk content replacement)

Use this when populating a slide template with new content across many slides.

### Step E1 — Visual overview

```bash
# Locate scripts directory
SCRIPTS="<path_to_skill>/edit-pptx/scripts"

# Generate thumbnail grid
python3 "$SCRIPTS/thumbnail.py" template.pptx template-thumbs --cols 4
# → Produces template-thumbs.jpg (or -1.jpg, -2.jpg for large decks)
```

Read the thumbnail image to understand slide layouts visually.

### Step E2 — Extract text inventory

```bash
python3 "$SCRIPTS/inventory.py" template.pptx inventory.json
```

Read `inventory.json` to see all text shapes: their positions, placeholder types,
font sizes, and existing text. Structure:

```json
{
  "slide-0": {
    "shape-0": {
      "placeholder_type": "TITLE",
      "left": 1.5, "top": 0.5, "width": 7.0, "height": 1.2,
      "default_font_size": 28.0,
      "paragraphs": [{"text": "Original title", "bold": true}]
    }
  }
}
```

### Step E3 — Rearrange slides (if needed)

If you need specific slides in a specific order (with duplicates):

```bash
python3 "$SCRIPTS/rearrange.py" template.pptx working.pptx 0,3,3,7,12
# Slide indices are 0-based. Same index = duplicate that slide.
```

### Step E4 — Create replacements JSON

Based on the inventory, build `replacements.json`. Only shapes with `"paragraphs"` get new content;
all other text shapes are cleared automatically.

```json
{
  "slide-0": {
    "shape-0": {
      "paragraphs": [
        {"text": "New Title", "bold": true, "alignment": "CENTER"}
      ]
    },
    "shape-1": {
      "paragraphs": [
        {"text": "Bullet item", "bullet": true, "level": 0},
        {"text": "Another item", "bullet": true, "level": 0}
      ]
    }
  }
}
```

Paragraph properties (all optional):

| Property | Type | Notes |
|---|---|---|
| `text` | string | Required |
| `bold` / `italic` / `underline` | bool | |
| `font_size` | float | Points |
| `font_name` | string | |
| `color` | string | RGB hex e.g. `"FF0000"` |
| `theme_color` | string | e.g. `"DARK_1"` |
| `alignment` | string | `"LEFT"` `"CENTER"` `"RIGHT"` `"JUSTIFY"` |
| `bullet` | bool | Set `true` for bullet points |
| `level` | int | Required when `bullet: true` |
| `space_before` / `space_after` | float | Points |
| `line_spacing` | float | Points |

**Rules:**
- Do NOT include bullet symbols (•, -, *) in text — they're added automatically when `bullet: true`
- Bullets default to LEFT alignment; don't override unless needed
- Match `default_font_size` from inventory to avoid overflow

### Step E5 — Apply replacements

```bash
python3 "$SCRIPTS/replace.py" working.pptx replacements.json output.pptx
```

The script validates all shape keys exist and errors if text overflow worsens.

### Step E6 — Verify visually

```bash
python3 "$SCRIPTS/thumbnail.py" output.pptx output-thumbs --cols 4
```

Read the thumbnail image and check for text cutoff, overlap, or positioning issues.

---

## Section F — Thumbnail grid

Generate a visual overview of any presentation at any time:

```bash
SCRIPTS="<path_to_skill>/edit-pptx/scripts"
python3 "$SCRIPTS/thumbnail.py" presentation.pptx [output_prefix] [--cols 4]
```

Grid limits by column count: 3=12 slides, 4=20, 5=30 (default), 6=42.
Multiple grid files are created automatically for large decks.

---

## Common pitfalls

| Symptom | Cause | Fix |
|---------|-------|-----|
| `'list' object has no attribute 'rId'` | Slides accessed as list slice | Use `list(prs.slides)[i]`, never `prs.slides[i:j]` |
| New slide appears at end, not at position N | Forgot to move `sldIdLst` entry | Always do the `sldIdLst.insert()` step after adding |
| Shapes invisible or wrong position | Guessed EMU values | Run Section C and copy the exact values |
| `<a:b/>` in XML not working | Attribute vs element mismatch | Use `<a:rPr b="1" ...>` (attribute) not `<a:rPr><a:b/>` |
| Text encoding breaks XML | Special chars (`<`, `>`, `&`) | Escape: `&lt;` `&gt;` `&amp;` |
| replace.py reports shape not found | Wrong shape key | Re-run inventory.py and check the exact key |
| Thumbnails show text cutoff | Text overflow after replacement | Shorten content or reduce font size |
