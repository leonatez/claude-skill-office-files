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

for skill in read-excel edit-excel read-docx edit-docx read-pptx edit-pptx; do
  src="$REPO_DIR/$skill/SKILL.md"
  dst="$SKILLS_DIR/$skill"
  if [ -f "$src" ]; then
    mkdir -p "$dst"
    cp "$src" "$dst/SKILL.md"
    echo "  Installed /$skill  ->  $dst"
  else
    echo "  SKIP $skill (SKILL.md not found)"
  fi
done

echo ""
echo "Done. Available slash commands:"
echo "  /read-excel   -- read & analyze .xlsx files"
echo "  /edit-excel   -- add sheets / write data to .xlsx files"
echo "  /read-docx    -- read & analyze .docx files"
echo "  /edit-docx    -- add sections / edit content in .docx files"
echo "  /read-pptx    -- read & analyze .pptx files"
echo "  /edit-pptx    -- add slides / edit content in .pptx files"
