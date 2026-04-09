---
name: edit-pptx
version: 1.0.0
description: |
  Add, insert, or edit slides in PowerPoint (.pptx) files while matching the original
  presentation's layout, shape positions, font sizes, colours, and formatting.
  Use when asked to "add a slide", "insert", "update", or "edit" a PowerPoint file.
dependencies:
  - python-pptx==1.0.2
  - lxml==6.0.2
allowed-tools:
  - Bash
  - Read
---

You have been invoked to edit a PowerPoint (.pptx) file. Follow every step without skipping.
The cardinal rule: **always inspect before you write** — never guess positions or colours.

---

## Step 0 — Ensure dependencies

```bash
pip3 install python-pptx==1.0.2 lxml==6.0.2 2>/dev/null | grep -E "^(Successfully|Already|Requirement)" || true
```

---

## Step 1 — Resolve and verify the file

```bash
ls -lh "<FILE_PATH>"
```

---

## Step 2 — Inspect slides near the insertion point (ALWAYS do this first)

Identify the slide number to insert after, then run this script to capture exact shape
geometry, colours, font sizes, and XML structure.

```bash
FILE_PATH="<FILE_PATH>"
TARGET_SLIDE=31   # slide to insert AFTER
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

# Inspect the target slide and its neighbours
for si in [target - 1, target, target + 1]:
    if 0 <= si - 1 < len(slides):
        slide = slides[si - 1]
        print(f"=== Slide {si} | Layout: {slide.slide_layout.name} ===")
        for shape in slide.shapes:
            l = shape.left  or 0; t = shape.top    or 0
            w = shape.width or 0; h = shape.height or 0
            print(f"  '{shape.name}': left={l} top={t} width={w} height={h}")
            # Fill colour
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

# Print condensed XML of key shapes from target slide
slide = slides[target - 1]
print(f"=== XML of shapes in slide {target} ===")
for shape in slide.shapes:
    if shape.has_text_frame and len(shape.text_frame.text.strip()) > 0:
        xml = etree.tostring(shape._element, pretty_print=True).decode()
        relevant = [l for l in xml.split('\n') if any(tag in l for tag in
            ['<a:off', '<a:ext', '<a:solidFill', '<a:srgbClr', '<a:schemeClr',
             '<a:bodyPr', '<a:pPr', '<a:rPr', '<a:t>', '<a:t ', 'val=',
             '<a:lnSpc', '<a:spcBef', '<a:spcAft', '<a:buNone', '<a:ind'])]
        print(f"\n  Shape: '{shape.name}'")
        for line in relevant[:35]:
            print(f"    {line.strip()}")

PYEOF
```

Record the values (in EMU) for the target slide's shapes. You will use these in Step 3.

---

## Step 3 — Write the edit script

### How to add a slide at a specific position

python-pptx can only append slides to the end. To insert at position N, append then
move the slide's `<p:sldId>` entry in the presentation's slide ID list:

```python
from pptx import Presentation
from lxml import etree

prs = Presentation(file_path)

# Add slide at end using the same layout as slide N
slides = list(prs.slides)
layout = slides[INSERT_AFTER - 1].slide_layout
new_slide = prs.slides.add_slide(layout)

# Clear all auto-generated placeholder shapes
sp_tree = new_slide.shapes._spTree
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
for sp in sp_tree.findall(f".//{{{P}}}sp"):
    sp_tree.remove(sp)

# ... add shapes to sp_tree (see below) ...

# Move the new slide from the end to position INSERT_AFTER + 1
sldIdLst = prs.slides._sldIdLst
children = list(sldIdLst)
new_entry = children[-1]
sldIdLst.remove(new_entry)
sldIdLst.insert(INSERT_AFTER, new_entry)  # 0-based index

prs.save(file_path)
```

### Adding shapes via raw XML

Use the EMU values from Step 2. All positions are in EMU (914400 EMU = 1 inch).

```python
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"

def make_textbox(shape_id, name, left, top, width, height,
                 paragraphs, is_title=False, body_anchor="t"):
    """
    paragraphs: list of dicts with keys:
      text      (str)
      bold      (bool, default False)
      size_pt   (int half-points, e.g. 3500 = 35pt)
      color     (str hex or None for inherit)
      italic    (bool, default False)
      indent_emu (int, 0 for no indent, 457200 for sub-indent)
      line_spc  (int percent*1000, e.g. 115000 for 115%)
      spc_before (int pts*100, e.g. 1200 for 12pt)
    """
    ph_xml = '<p:nvPr><p:ph type="title"/></p:nvPr>' if is_title else '<p:nvPr/>'
    
    para_xml = ""
    for p in paragraphs:
        b      = "<a:b/>" if p.get("bold") else ""
        i_     = "<a:i/>" if p.get("italic") else ""
        col    = p.get("color")
        color_xml = f'<a:solidFill><a:srgbClr val="{col}"/></a:solidFill>' if col else \
                    '<a:solidFill><a:schemeClr val="dk1"/></a:solidFill>'
        sz     = p.get("size_pt", 3500)
        indent = p.get("indent_emu", 0)
        lspc   = p.get("line_spc", 115000)
        spbef  = p.get("spc_before", 1200)
        text   = p["text"]
        
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
        <a:t>{text}</a:t>
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
    """Solid-fill rectangular badge (e.g. company logo badge)."""
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

### Usage example

```python
sp_tree.append(make_textbox(
    shape_id=401, name="title_shape",
    left=2794930, top=164802, width=18086100, height=1462800,
    is_title=True,
    paragraphs=[{"text": "My New Slide Title", "bold": False, "size_pt": 3500}]
))

sp_tree.append(make_textbox(
    shape_id=402, name="body_shape",
    left=572100, top=1627600, width=23239800, height=10875300,
    paragraphs=[
        {"text": "Step 1:", "bold": True,  "size_pt": 3500, "spc_before": 1200},
        {"text": "Step 2:", "bold": True,  "size_pt": 3500, "spc_before": 1200},
        {"text": "For eg: example command", "bold": False, "size_pt": 3500,
         "italic": True, "indent_emu": 457200, "spc_before": 1200},
    ]
))

sp_tree.append(make_badge(
    shape_id=403, name="badge",
    left=20707200, top=512700, width=1846200, height=521300,
    label="VPBank", fill_color="6AA84F"
))
```

---

## Step 4 — Verify the result

```bash
FILE_PATH="<FILE_PATH>"
TARGET_SLIDE=31
python3 - << 'PYEOF'
import subprocess
file_path = subprocess.check_output("echo \"$FILE_PATH\"", shell=True).decode().strip()
target = int(subprocess.check_output("echo \"$TARGET_SLIDE\"", shell=True).decode().strip())
from pptx import Presentation
prs = Presentation(file_path)
slides = list(prs.slides)
print(f"Total slides: {len(slides)}")
new_slide = slides[target]   # 0-based: slide after INSERT_AFTER
print(f"New slide layout: {new_slide.slide_layout.name}")
for shape in new_slide.shapes:
    if shape.has_text_frame:
        print(f"  Shape '{shape.name}': '{shape.text_frame.text[:80]}'")
PYEOF
```

---

## Common pitfalls

| Symptom | Cause | Fix |
|---------|-------|-----|
| `'list' object has no attribute 'rId'` | Slides accessed as list slice | Use `list(prs.slides)[i]`, never `prs.slides[i:j]` |
| New slide appears at end, not at position N | Forgot to move `sldIdLst` entry | Always do the `sldIdLst.insert()` step after adding |
| Shapes invisible or wrong position | Guessed EMU values | Run Step 2 and copy the exact values |
| `<a:b/>` in XML not working | Attribute vs element mismatch | Use `<a:rPr b="1" ...>` (attribute) not `<a:rPr><a:b/>` (sub-element) |
| Text encoding breaks XML | Special chars (`<`, `>`, `&`) in text | Escape: `&lt;` `&gt;` `&amp;` |
| Slide count correct but content empty | Shapes appended to wrong tree | Always use `new_slide.shapes._spTree` not `slide.shapes._spTree` |
