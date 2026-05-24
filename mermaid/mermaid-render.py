#!/usr/bin/env python3
"""
Render a Mermaid diagram to PNG.

Methods tried in order:
  1. mmdc via npx @mermaid-js/mermaid-cli  (offline, requires Node.js)
  2. mermaid.ink public API                 (online, no Node.js needed)

Usage:
    python3 mermaid-render.py --input diagram.mmd --output /tmp/out.png
    python3 mermaid-render.py --output /tmp/out.png < diagram.mmd
    python3 mermaid-render.py --text "flowchart LR\nA-->B" --output /tmp/out.png
    python3 mermaid-render.py --input diagram.mmd --output /tmp/out.png --theme dark
"""
import argparse
import base64
import os
import subprocess
import sys
import tempfile
import urllib.request
from pathlib import Path

THEMES = ("default", "dark", "forest", "neutral")

# Puppeteer config to disable Chrome sandbox (required on Ubuntu with AppArmor restrictions)
_PUPPETEER_CONFIG = '{"args":["--no-sandbox","--disable-setuid-sandbox"]}'


def _render_via_mmdc(mmd_path: str, out_path: str, theme: str) -> bool:
    """Render using mmdc (mermaid-cli) via npx — works offline with Node.js."""
    import tempfile as _tf

    # Write puppeteer config to a temp file (mmdc requires a file, not inline JSON)
    with _tf.NamedTemporaryFile(mode="w", suffix=".json", delete=False) as cfg:
        cfg.write(_PUPPETEER_CONFIG)
        cfg_path = cfg.name

    try:
        result = subprocess.run(
            [
                "npx", "-y", "@mermaid-js/mermaid-cli",
                "-i", mmd_path,
                "-o", out_path,
                "--theme", theme,
                "--backgroundColor", "white",
                "--puppeteerConfigFile", cfg_path,
            ],
            capture_output=True,
            text=True,
            timeout=90,
        )
        if result.returncode != 0:
            print(f"[mmdc stderr] {result.stderr[:400]}", file=sys.stderr)
        return result.returncode == 0 and Path(out_path).exists() and Path(out_path).stat().st_size > 500
    except (FileNotFoundError, subprocess.TimeoutExpired, OSError) as e:
        print(f"[mmdc] not available: {e}", file=sys.stderr)
        return False
    finally:
        os.unlink(cfg_path)


def _render_via_mermaid_ink(content: str, out_path: str) -> bool:
    """Fallback: mermaid.ink free public API — needs internet, no install."""
    try:
        encoded = base64.urlsafe_b64encode(content.encode("utf-8")).decode("ascii")
        url = f"https://mermaid.ink/img/{encoded}?type=png"
        req = urllib.request.Request(url, headers={"User-Agent": "mermaid-render/1.0"})
        with urllib.request.urlopen(req, timeout=30) as resp:
            data = resp.read()
        if data and len(data) > 500:
            Path(out_path).write_bytes(data)
            return True
        return False
    except Exception as e:
        print(f"[mermaid.ink] failed: {e}", file=sys.stderr)
        return False


def render(content: str, output_path: str, theme: str = "default") -> str:
    """Render mermaid content to PNG. Returns the output path on success."""
    with tempfile.NamedTemporaryFile(
        mode="w", suffix=".mmd", delete=False, encoding="utf-8"
    ) as f:
        f.write(content)
        mmd_path = f.name

    try:
        if _render_via_mmdc(mmd_path, output_path, theme):
            return output_path

        print("[mermaid] mmdc failed — trying mermaid.ink API...", file=sys.stderr)
        if _render_via_mermaid_ink(content, output_path):
            return output_path

        raise RuntimeError(
            "Mermaid rendering failed via all methods.\n"
            "  • mmdc: npx unavailable or returned error (see stderr above)\n"
            "  • mermaid.ink: network unreachable or request blocked\n"
            "Fix: install Node.js (https://nodejs.org) or check network connectivity."
        )
    finally:
        os.unlink(mmd_path)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Render a Mermaid diagram to PNG",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    src = parser.add_mutually_exclusive_group()
    src.add_argument("--input", "-i", metavar="FILE", help=".mmd input file")
    src.add_argument("--text", "-t", metavar="TEXT", help="Diagram text inline")
    parser.add_argument("--output", "-o", required=True, metavar="PNG", help="Output PNG path")
    parser.add_argument("--theme", default="default", choices=THEMES, help="Mermaid theme")
    args = parser.parse_args()

    if args.text:
        content = args.text.replace("\\n", "\n")
    elif args.input:
        content = Path(args.input).read_text(encoding="utf-8")
    else:
        content = sys.stdin.read()

    if not content.strip():
        print("Error: empty diagram content", file=sys.stderr)
        sys.exit(1)

    out = render(content, args.output, args.theme)
    print(out)


if __name__ == "__main__":
    main()
