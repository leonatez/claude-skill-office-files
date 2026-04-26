#!/usr/bin/env python3
"""
Pack an unpacked Office document directory back into a .docx / .pptx / .xlsx file.

Usage:
    python pack.py <input_dir> <output_file>

XML files are condensed (pretty-print whitespace removed) before packing
so the result is byte-compatible with files produced by Office applications.
"""

import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

try:
    import defusedxml.minidom as minidom
except ImportError:
    import xml.dom.minidom as minidom


def _condense_xml(xml_file):
    """Remove pretty-print whitespace from an XML file, preserving w:t content."""
    try:
        dom = minidom.parse(str(xml_file))
    except Exception:
        return  # leave unparseable files alone

    def strip_whitespace(node):
        for child in list(node.childNodes):
            tag = getattr(child, "tagName", "") or ""
            if tag.endswith(":t") or tag == "t":
                continue  # never strip inside text nodes
            if child.nodeType == child.TEXT_NODE and child.nodeValue.strip() == "":
                node.removeChild(child)
            elif child.nodeType == child.COMMENT_NODE:
                node.removeChild(child)
            else:
                strip_whitespace(child)

    strip_whitespace(dom.documentElement)
    xml_file.write_bytes(dom.toxml(encoding="UTF-8"))


def pack(input_dir, output_file):
    src = Path(input_dir)
    dst = Path(output_file)

    if not src.is_dir():
        raise ValueError(f"Not a directory: {input_dir}")
    if dst.suffix.lower() not in {".docx", ".pptx", ".xlsx"}:
        raise ValueError(f"Output must be .docx, .pptx, or .xlsx: {output_file}")

    with tempfile.TemporaryDirectory() as tmp:
        staging = Path(tmp) / "content"
        shutil.copytree(src, staging)

        for pattern in ("*.xml", "*.rels"):
            for f in staging.rglob(pattern):
                _condense_xml(f)

        dst.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zf:
            for f in staging.rglob("*"):
                if f.is_file():
                    zf.write(f, f.relative_to(staging))

    print(f"Packed to: {output_file}")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print(__doc__)
        sys.exit(1)
    pack(sys.argv[1], sys.argv[2])
