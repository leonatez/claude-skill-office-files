#!/usr/bin/env python3
"""
Generate a thumbnail grid image from a PowerPoint file.

Usage:
    python thumbnail.py input.pptx [output_prefix] [--cols N]

Examples:
    python thumbnail.py deck.pptx                  # → thumbnails.jpg
    python thumbnail.py deck.pptx grid --cols 4    # → grid.jpg or grid-1.jpg, grid-2.jpg …

Grid sizes by column count:
    3 cols → max 12 slides/grid (3×4)
    4 cols → max 20 slides/grid (4×5)
    5 cols → max 30 slides/grid (5×6)  [default]
    6 cols → max 42 slides/grid (6×7)

Requires: pillow, python-pptx, LibreOffice (soffice), poppler (pdftoppm)
"""

import argparse
import subprocess
import sys
import tempfile
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

THUMB_W = 300
DPI = 100
JPEG_Q = 95
PADDING = 20
BORDER = 2
FONT_RATIO = 0.12
LABEL_PAD_RATIO = 0.4
MAX_COLS = 6
DEFAULT_COLS = 5


def _convert_to_images(pptx_path, tmp_dir):
    pdf = tmp_dir / f"{pptx_path.stem}.pdf"
    r = subprocess.run(
        ["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(tmp_dir), str(pptx_path)],
        capture_output=True, text=True,
    )
    if r.returncode != 0 or not pdf.exists():
        raise RuntimeError(f"LibreOffice PDF conversion failed:\n{r.stderr}")

    r2 = subprocess.run(
        ["pdftoppm", "-jpeg", "-r", str(DPI), str(pdf), str(tmp_dir / "slide")],
        capture_output=True, text=True,
    )
    if r2.returncode != 0:
        raise RuntimeError(f"pdftoppm failed:\n{r2.stderr}")

    return sorted(tmp_dir.glob("slide-*.jpg"))


def _make_grid(images, cols, start_idx):
    font_size = int(THUMB_W * FONT_RATIO)
    lpad = int(font_size * LABEL_PAD_RATIO)

    with Image.open(images[0]) as img:
        aspect = img.height / img.width
    h = int(THUMB_W * aspect)

    rows = (len(images) + cols - 1) // cols
    gw = cols * THUMB_W + (cols + 1) * PADDING
    gh = rows * (h + font_size + lpad * 2) + (rows + 1) * PADDING
    grid = Image.new("RGB", (gw, gh), "white")
    draw = ImageDraw.Draw(grid)

    try:
        font = ImageFont.load_default(size=font_size)
    except Exception:
        font = ImageFont.load_default()

    for i, img_path in enumerate(images):
        row, col = divmod(i, cols)
        x = col * THUMB_W + (col + 1) * PADDING
        y_base = row * (h + font_size + lpad * 2) + (row + 1) * PADDING

        label = str(start_idx + i)
        bbox = draw.textbbox((0, 0), label, font=font)
        tw = bbox[2] - bbox[0]
        draw.text((x + (THUMB_W - tw) // 2, y_base + lpad), label, fill="black", font=font)

        y_thumb = y_base + lpad + font_size + lpad
        with Image.open(img_path) as img:
            img.thumbnail((THUMB_W, h), Image.Resampling.LANCZOS)
            iw, ih = img.size
            tx = x + (THUMB_W - iw) // 2
            ty = y_thumb + (h - ih) // 2
            grid.paste(img, (tx, ty))
            if BORDER:
                draw.rectangle(
                    [(tx - BORDER, ty - BORDER), (tx + iw + BORDER - 1, ty + ih + BORDER - 1)],
                    outline="gray", width=BORDER,
                )

    return grid


def generate(pptx_path, output_prefix="thumbnails", cols=DEFAULT_COLS):
    cols = min(cols, MAX_COLS)
    max_per_grid = cols * (cols + 1)

    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        print(f"Converting {pptx_path} to images…")
        images = _convert_to_images(Path(pptx_path), tmp_dir)
        if not images:
            raise RuntimeError("No slides found")
        print(f"Found {len(images)} slides")

        out_base = Path(f"{output_prefix}.jpg")
        files = []
        for chunk_i, start in enumerate(range(0, len(images), max_per_grid)):
            chunk = images[start:start + max_per_grid]
            grid = _make_grid(chunk, cols, start)

            if len(images) <= max_per_grid:
                out = out_base
            else:
                out = out_base.parent / f"{out_base.stem}-{chunk_i + 1}{out_base.suffix}"

            out.parent.mkdir(parents=True, exist_ok=True)
            grid.save(str(out), quality=JPEG_Q)
            files.append(str(out))

        print(f"Created {len(files)} grid(s):")
        for f in files:
            print(f"  - {f}")
        return files


def main():
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("input", help="Input .pptx file")
    parser.add_argument("output_prefix", nargs="?", default="thumbnails",
                        help="Output filename prefix (default: thumbnails)")
    parser.add_argument("--cols", type=int, default=DEFAULT_COLS,
                        help=f"Columns per grid (default: {DEFAULT_COLS}, max: {MAX_COLS})")
    args = parser.parse_args()

    if not Path(args.input).exists():
        print(f"Error: {args.input} not found")
        sys.exit(1)

    try:
        generate(args.input, args.output_prefix, args.cols)
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
