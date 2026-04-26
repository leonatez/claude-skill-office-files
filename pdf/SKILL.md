---
name: pdf
version: 1.0.0
description: |
  Extract text/tables, merge, split, rotate, fill forms, add watermarks, and create PDFs.
  Supports both searchable PDFs and scanned/image PDFs (via OCR). Use when asked to
  "read a PDF", "extract from PDF", "fill PDF form", "merge PDFs", or "create a PDF".
dependencies:
  - pypdf>=4.0.0
  - pdfplumber>=0.10.0
  - reportlab>=4.0.0
allowed-tools:
  - Bash
  - Read
---

You have been invoked to process a PDF file. Follow every step below.

---

## Step 0 — Ensure dependencies

```bash
pip3 install "pypdf>=4.0.0" "pdfplumber>=0.10.0" "reportlab>=4.0.0" \
  2>/dev/null | grep -E "^(Successfully|Already|Requirement)" || true
```

---

## Step 1 — Resolve and verify the file

```bash
ls -lh "<FILE_PATH>"
```

---

## Step 2 — Choose the right tool

| Task | Best tool |
|------|-----------|
| Extract text | pdfplumber |
| Extract tables | pdfplumber |
| Merge / split / rotate | pypdf |
| Fill form fields | pypdf |
| Create new PDF | reportlab |
| OCR scanned PDF | pytesseract + pdf2image |
| Command-line merge | qpdf |

---

## Common workflows

### Extract text

```python
import pdfplumber

with pdfplumber.open("<FILE_PATH>") as pdf:
    print(f"Pages: {len(pdf.pages)}")
    for i, page in enumerate(pdf.pages, 1):
        text = page.extract_text()
        if text:
            print(f"\n--- Page {i} ---")
            print(text)
```

### Extract tables

```python
import pdfplumber, pandas as pd

with pdfplumber.open("<FILE_PATH>") as pdf:
    all_tables = []
    for i, page in enumerate(pdf.pages, 1):
        tables = page.extract_tables()
        for j, table in enumerate(tables):
            if table:
                df = pd.DataFrame(table[1:], columns=table[0])
                print(f"\nPage {i}, Table {j+1}:")
                print(df.to_string())
                all_tables.append(df)
```

### Merge PDFs

```python
from pypdf import PdfWriter, PdfReader

writer = PdfWriter()
for pdf_file in ["doc1.pdf", "doc2.pdf", "doc3.pdf"]:
    reader = PdfReader(pdf_file)
    for page in reader.pages:
        writer.add_page(page)

with open("merged.pdf", "wb") as f:
    writer.write(f)
```

### Split PDF

```python
from pypdf import PdfReader, PdfWriter

reader = PdfReader("<FILE_PATH>")
for i, page in enumerate(reader.pages):
    writer = PdfWriter()
    writer.add_page(page)
    with open(f"page_{i+1}.pdf", "wb") as f:
        writer.write(f)
```

### Rotate pages

```python
from pypdf import PdfReader, PdfWriter

reader = PdfReader("<FILE_PATH>")
writer = PdfWriter()
for page in reader.pages:
    page.rotate(90)  # 90, 180, or 270
    writer.add_page(page)

with open("rotated.pdf", "wb") as f:
    writer.write(f)
```

### Extract metadata

```python
from pypdf import PdfReader

reader = PdfReader("<FILE_PATH>")
meta = reader.metadata
print(f"Title  : {meta.title}")
print(f"Author : {meta.author}")
print(f"Pages  : {len(reader.pages)}")
```

### Fill PDF form fields

```python
from pypdf import PdfReader, PdfWriter

reader = PdfReader("<FILE_PATH>")
writer = PdfWriter()
writer.append(reader)

# List available fields first
fields = reader.get_fields()
print("Available fields:", list(fields.keys()) if fields else "None (not a form)")

# Fill fields
writer.update_page_form_field_values(
    writer.pages[0],
    {
        "field_name_1": "Value 1",
        "field_name_2": "Value 2",
    }
)

with open("filled.pdf", "wb") as f:
    writer.write(f)
```

### Create a new PDF

```python
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

doc = SimpleDocTemplate("output.pdf", pagesize=A4)
styles = getSampleStyleSheet()
story = []

story.append(Paragraph("Report Title", styles['Title']))
story.append(Spacer(1, 12))
story.append(Paragraph("Body text here.", styles['Normal']))

# Table
data = [["Column A", "Column B", "Column C"],
        ["Row 1a", "Row 1b", "Row 1c"],
        ["Row 2a", "Row 2b", "Row 2c"]]
t = Table(data)
t.setStyle(TableStyle([
    ('BACKGROUND', (0,0), (-1,0), colors.grey),
    ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
    ('GRID', (0,0), (-1,-1), 1, colors.black),
    ('ALIGN', (0,0), (-1,-1), 'CENTER'),
]))
story.append(t)
doc.build(story)
```

### OCR a scanned PDF

```bash
pip3 install pytesseract pdf2image 2>/dev/null | grep -E "^(Successfully|Already)" || true
```

```python
import pytesseract
from pdf2image import convert_from_path

images = convert_from_path("<FILE_PATH>")
for i, image in enumerate(images, 1):
    text = pytesseract.image_to_string(image)
    print(f"\n--- Page {i} ---")
    print(text)
```

### Add watermark

```python
from pypdf import PdfReader, PdfWriter

watermark = PdfReader("watermark.pdf").pages[0]
reader = PdfReader("<FILE_PATH>")
writer = PdfWriter()

for page in reader.pages:
    page.merge_page(watermark)
    writer.add_page(page)

with open("watermarked.pdf", "wb") as f:
    writer.write(f)
```

### Password protect

```python
from pypdf import PdfReader, PdfWriter

reader = PdfReader("<FILE_PATH>")
writer = PdfWriter()
for page in reader.pages:
    writer.add_page(page)

writer.encrypt("userpassword", "ownerpassword")

with open("encrypted.pdf", "wb") as f:
    writer.write(f)
```

---

## Command-line alternatives

```bash
# Merge
qpdf --empty --pages file1.pdf file2.pdf -- merged.pdf

# Split pages 1-5
qpdf input.pdf --pages . 1-5 -- pages1-5.pdf

# Extract text
pdftotext input.pdf output.txt
pdftotext -layout input.pdf output.txt  # preserve layout

# Extract images
pdfimages -j input.pdf output_prefix
```

---

## Quick reference

| Task | Tool | Notes |
|------|------|-------|
| Text extraction | pdfplumber | Layout-aware |
| Table extraction | pdfplumber | Returns list of lists |
| Merge/split/rotate | pypdf | Pure Python, no external deps |
| Form filling | pypdf | AcroForm fields only |
| Create from scratch | reportlab | Full control over layout |
| OCR | pytesseract | Needs tesseract installed |
| CLI merge | qpdf | Fastest for large files |

---

## Common pitfalls

| Symptom | Cause | Fix |
|---------|-------|-----|
| Empty text from extraction | Scanned PDF (images only) | Use OCR path with pytesseract |
| Form fields not found | Not an AcroForm PDF | Check `reader.get_fields()` — may return None |
| Tables extracted as None | Complex/borderless tables | Try `page.extract_table(table_settings={...})` with custom settings |
| Unicode characters garbled | Encoding issue | Use `pdfplumber` instead of `pdftotext` |
| Large PDF slow to process | Loading all pages | Use `pdf.pages[start:end]` slice |
