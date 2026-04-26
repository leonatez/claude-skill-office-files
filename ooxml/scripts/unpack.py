#!/usr/bin/env python3
"""
Unpack a .docx / .pptx / .xlsx file to a directory for XML editing.

Usage:
    python unpack.py <office_file> <output_dir>

All XML and .rels files are pretty-printed for easy reading and diffing.
For .docx files, a suggested RSID is printed for use in tracked-change edits.
"""

import random
import sys
import zipfile
from pathlib import Path

try:
    import defusedxml.minidom as minidom
except ImportError:
    import xml.dom.minidom as minidom  # fallback (less secure for untrusted files)


def unpack(office_file, output_dir):
    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(office_file) as zf:
        zf.extractall(out)

    for pattern in ("*.xml", "*.rels"):
        for f in out.rglob(pattern):
            try:
                content = f.read_bytes()
                dom = minidom.parseString(content)
                f.write_bytes(dom.toprettyxml(indent="  ", encoding="utf-8"))
            except Exception:
                pass  # leave non-parseable files as-is

    if str(office_file).endswith(".docx"):
        rsid = "".join(random.choices("0123456789ABCDEF", k=8))
        print(f"Suggested RSID for this edit session: {rsid}")

    print(f"Unpacked to: {output_dir}")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print(__doc__)
        sys.exit(1)
    unpack(sys.argv[1], sys.argv[2])
