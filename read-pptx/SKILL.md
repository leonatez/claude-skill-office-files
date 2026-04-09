---
name: read-pptx
version: 1.0.0
description: |
  Read and understand PowerPoint (.pptx) files. Extracts all slides with their shapes,
  text content, run-level formatting (font size, bold, colour), positions, and layout
  names. Produces a structured summary. Use when asked to "read", "analyze", or
  "understand" a PowerPoint file.
dependencies:
  - python-pptx==1.0.2
  - lxml==6.0.2
allowed-tools:
  - Bash
  - Read
---

You have been invoked to read and understand a PowerPoint (.pptx) file. Follow every step.

---

## Step 0 — Ensure dependencies

```bash
pip3 install python-pptx==1.0.2 2>/dev/null | grep -E "^(Successfully|Already|Requirement)" || true
```

---

## Step 1 — Resolve and verify the file

```bash
ls -lh "<FILE_PATH>"
```

---

## Step 2 — Extract all slides

```bash
FILE_PATH="<FILE_PATH>"
python3 - << 'PYEOF'
import os, subprocess
file_path = subprocess.check_output("echo \"$FILE_PATH\"", shell=True).decode().strip()

from pptx import Presentation

prs = Presentation(file_path)
slides = list(prs.slides)

print(f"=== PRESENTATION OVERVIEW ===")
print(f"Total slides : {len(slides)}")
print(f"Slide size   : {prs.slide_width.inches:.2f} x {prs.slide_height.inches:.2f} inches")
print()

for si, slide in enumerate(slides, start=1):
    layout = slide.slide_layout.name
    print(f"--- Slide {si} | Layout: {layout} ---")
    for shape in slide.shapes:
        left  = shape.left  / 914400 if shape.left  else 0
        top   = shape.top   / 914400 if shape.top   else 0
        w     = shape.width / 914400 if shape.width else 0
        h     = shape.height/ 914400 if shape.height else 0
        print(f"  [{shape.shape_type}] '{shape.name}' pos=({left:.2f},{top:.2f}) size=({w:.2f}x{h:.2f})")
        if shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                text = para.text.strip()
                if text:
                    run_info = ""
                    for run in para.runs:
                        if run.text.strip():
                            try:    col = str(run.font.color.rgb)
                            except: col = "inherit"
                            run_info = f"sz={run.font.size} bold={run.font.bold} color={col}"
                            break
                    print(f"    TEXT: '{text[:90]}' | {run_info}")
    print()

print("=== EXTRACTION COMPLETE ===")
PYEOF
```

---

## Step 3 — Synthesize a structured summary

### Presentation Overview
- **File**: filename
- **Slide count**: N
- **Slide dimensions**: W x H inches
- **Purpose**: what this presentation is about

### Slide-by-slide summary
For each slide: slide number, layout name, main title text, key content bullets.

### Styling Conventions
Document for use with `/edit-pptx`:
- Title shape: position, font size, bold, colour
- Body text shape: position, font size, line spacing
- Accent colours found (hex values)
- Badge/logo shapes and their fill colours

---

## Common pitfalls

| Symptom | Cause | Fix |
|---------|-------|-----|
| `'list' object has no attribute 'rId'` | Slides indexed as list slice | Always use `list(prs.slides)[i]` not `prs.slides[i:j]` |
| Shape has no text | Image or drawing shape | Check `shape.has_text_frame` before accessing `.text_frame` |
| Font colour raises exception | Colour type is inherited (not explicit) | Wrap `run.font.color.rgb` in try/except |
