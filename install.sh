#!/usr/bin/env bash
set -e
SKILL_DIR="$HOME/.claude/skills/read-excel"
mkdir -p "$SKILL_DIR"
cp "$(dirname "$0")/SKILL.md" "$SKILL_DIR/SKILL.md"
echo "Installed /read-excel skill to $SKILL_DIR"
