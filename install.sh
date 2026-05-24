#!/usr/bin/env bash
# install.sh — installs all office-file skills + Python dependencies
# Usage: bash install.sh
set -e

REPO_DIR="$(cd "$(dirname "$0")" && pwd)"
SKILLS_DIR="$HOME/.claude/skills"

echo "=== Installing Python dependencies ==="
pip3 install -r "$REPO_DIR/requirements.txt"

echo ""
echo "=== Installing Claude Code skills ==="

SKILLS="read-excel edit-excel read-docx edit-docx read-pptx edit-pptx pdf"

for skill in $SKILLS; do
  src="$REPO_DIR/$skill/SKILL.md"
  dst="$SKILLS_DIR/$skill"
  if [ -f "$src" ]; then
    mkdir -p "$dst"
    cp "$src" "$dst/SKILL.md"

    # Copy bundled scripts if present
    if [ -d "$REPO_DIR/$skill/scripts" ]; then
      mkdir -p "$dst/scripts"
      cp -r "$REPO_DIR/$skill/scripts/." "$dst/scripts/"
      echo "  Installed /$skill (+ scripts)  ->  $dst"
    elif [ -f "$REPO_DIR/$skill/recalc.py" ] || [ -f "$REPO_DIR/$skill/docx-apply-json-spec.py" ]; then
      [ -f "$REPO_DIR/$skill/recalc.py" ] && cp "$REPO_DIR/$skill/recalc.py" "$dst/recalc.py"
      [ -f "$REPO_DIR/$skill/docx-apply-json-spec.py" ] && cp "$REPO_DIR/$skill/docx-apply-json-spec.py" "$dst/docx-apply-json-spec.py"
      echo "  Installed /$skill (+ scripts)  ->  $dst"
    else
      echo "  Installed /$skill  ->  $dst"
    fi
  else
    echo "  SKIP $skill (SKILL.md not found)"
  fi
done

# Install shared ooxml scripts
if [ -d "$REPO_DIR/ooxml/scripts" ]; then
  mkdir -p "$SKILLS_DIR/ooxml/scripts"
  cp -r "$REPO_DIR/ooxml/scripts/." "$SKILLS_DIR/ooxml/scripts/"
  echo "  Installed shared ooxml scripts  ->  $SKILLS_DIR/ooxml/scripts"
fi

# Install shared mermaid rendering script
if [ -f "$REPO_DIR/mermaid/mermaid-render.py" ]; then
  mkdir -p "$SKILLS_DIR/mermaid"
  cp "$REPO_DIR/mermaid/mermaid-render.py" "$SKILLS_DIR/mermaid/mermaid-render.py"
  echo "  Installed shared mermaid/mermaid-render.py  ->  $SKILLS_DIR/mermaid/"
fi

echo ""
echo "Done. Available slash commands:"
echo "  /read-excel   -- read & analyze .xlsx files (merged cells, images, API spec parsing)"
echo "  /edit-excel   -- edit .xlsx files with formula recalculation via LibreOffice"
echo "  /read-docx    -- read & analyze .docx files (pandoc + python-docx)"
echo "  /edit-docx    -- edit .docx files with tracked changes / redlining support"
echo "  /read-pptx    -- read & analyze .pptx files"
echo "  /edit-pptx    -- edit .pptx files + template bulk-replace workflow"
echo "  /pdf          -- read, extract, merge, split, fill, and create PDF files"
echo ""
echo "Optional: install system tools for full feature support:"
echo "  sudo apt-get install pandoc libreoffice poppler-utils"
echo "  # macOS: brew install pandoc libreoffice poppler"
