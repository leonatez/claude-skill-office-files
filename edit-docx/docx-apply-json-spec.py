#!/usr/bin/env python3
"""
Apply structured JSON edits to a .docx file.
Claude writes JSON — this script handles all OOXML internally.

Usage:
    python3 docx-apply-json-spec.py spec.json
    echo '{...}' | python3 docx-apply-json-spec.py

Spec format:
    {
      "file": "/path/to/doc.docx",
      "insert_after": 12,          # paragraph index (0-based); or "append": true
      "elements": [
        {"type": "heading",   "text": "Title",    "level": 1},
        {"type": "paragraph", "text": "Body.",    "bold": false, "size": 11},
        {"type": "image",     "path": "/tmp/x.png", "width_inches": 5.5},
        {"type": "table",     "headers": ["A","B"], "rows": [["1","2"]]},
        {"type": "code",      "lines": ["def foo():", "    pass"]},
        {"type": "empty"},
        {"type": "pagebreak"}
      ]
    }
"""
import argparse
import json
import sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

_ALIGN = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    "both": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


def _apply_run_fmt(run, spec):
    if "bold" in spec:
        run.bold = spec["bold"]
    if "italic" in spec:
        run.italic = spec["italic"]
    if "size" in spec:
        run.font.size = Pt(spec["size"])
    if "font" in spec:
        run.font.name = spec["font"]
    if "color" in spec:
        h = spec["color"].lstrip("#")
        run.font.color.rgb = RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _add_paragraph(doc, spec):
    style = spec.get("style")
    para = doc.add_paragraph(spec.get("text", ""), style=style)
    if para.runs:
        _apply_run_fmt(para.runs[0], spec)
    if "align" in spec:
        para.alignment = _ALIGN.get(spec["align"].lower(), WD_ALIGN_PARAGRAPH.LEFT)
    return para


def _add_heading(doc, spec):
    return doc.add_heading(spec.get("text", ""), level=spec.get("level", 1))


def _add_image(doc, spec):
    para = doc.add_paragraph()
    run = para.add_run()
    kwargs = {}
    if "width_inches" in spec:
        kwargs["width"] = Inches(spec["width_inches"])
    run.add_picture(spec["path"], **kwargs)
    para.alignment = _ALIGN.get(spec.get("align", "center"), WD_ALIGN_PARAGRAPH.CENTER)
    return para


def _add_table(doc, spec):
    headers = spec.get("headers", [])
    rows = spec.get("rows", [])
    cols = len(headers) or (len(rows[0]) if rows else 1)
    tbl = doc.add_table(rows=len(rows) + (1 if headers else 0), cols=cols)
    tbl.style = spec.get("style", "Table Grid")
    offset = 0
    if headers:
        for ci, h in enumerate(headers):
            cell = tbl.rows[0].cells[ci]
            cell.text = str(h)
            if cell.paragraphs[0].runs:
                cell.paragraphs[0].runs[0].bold = True
        offset = 1
    for ri, row_data in enumerate(rows):
        for ci, val in enumerate(row_data):
            tbl.rows[ri + offset].cells[ci].text = str(val)
    return tbl


def _add_code(doc, spec):
    lines = spec.get("lines") or [spec.get("text", "")]
    paras = []
    for line in lines:
        para = doc.add_paragraph()
        run = para.add_run(line)
        run.font.name = "Courier New"
        run.font.size = Pt(spec.get("size", 9))
        paras.append(para)
    return paras


def _add_empty(doc, _=None):
    return doc.add_paragraph()


def _add_pagebreak(doc, _=None):
    para = doc.add_paragraph()
    run = para.add_run()
    br = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    run._r.append(br)
    return para


_BUILDERS = {
    "paragraph": _add_paragraph,
    "heading": _add_heading,
    "image": _add_image,
    "table": _add_table,
    "code": _add_code,
    "empty": _add_empty,
    "pagebreak": _add_pagebreak,
}


def apply_spec(spec: dict) -> int:
    file_path = Path(spec["file"])
    if not file_path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    doc = Document(str(file_path))
    body = doc.element.body

    # Snapshot BEFORE adding: original body-level paragraph elements only.
    # Must happen before add_* calls because add_table() also creates paragraphs inside cells,
    # which would pollute doc.paragraphs and shift indices.
    original_para_elems = [p._element for p in doc.paragraphs]
    before_ids = {id(e) for e in list(body)}

    # Build all new elements directly in the target doc (preserves image relationships)
    for elem_spec in spec.get("elements", []):
        etype = elem_spec.get("type", "paragraph")
        builder = _BUILDERS.get(etype)
        if not builder:
            raise ValueError(f"Unknown element type: {etype!r}. Valid: {list(_BUILDERS)}")
        builder(doc, elem_spec)

    # Collect new direct body children (paragraphs + tables) added during the loop
    new_elems = [e for e in list(body) if id(e) not in before_ids]

    if not new_elems:
        print("Warning: no elements were added.", file=sys.stderr)
        return 0

    # Move to insertion point (default: insert_after; append stays at end)
    if not spec.get("append"):
        anchor_idx = spec.get("insert_after", spec.get("insert_before"))
        if anchor_idx is not None:
            if anchor_idx >= len(original_para_elems):
                raise IndexError(
                    f"insert_after={anchor_idx} out of range "
                    f"(doc has {len(original_para_elems)} paragraphs)"
                )
            anchor_elem = original_para_elems[anchor_idx]
            insert_at = list(body).index(anchor_elem)
            insert_at += 0 if spec.get("insert_before") else 1

            for e in new_elems:
                body.remove(e)
            for i, e in enumerate(new_elems):
                body.insert(insert_at + i, e)

    doc.save(str(file_path))
    print(f"Saved {file_path} — {len(new_elems)} element(s) inserted.")
    return len(new_elems)


def main() -> None:
    parser = argparse.ArgumentParser(description="Apply JSON-spec edits to a .docx file")
    parser.add_argument("spec", nargs="?", help="JSON spec file path (omit to read from stdin)")
    args = parser.parse_args()

    if not args.spec or args.spec == "-":
        raw = sys.stdin.read()
    else:
        raw = Path(args.spec).read_text(encoding="utf-8")
    spec = json.loads(raw)
    apply_spec(spec)


if __name__ == "__main__":
    main()
