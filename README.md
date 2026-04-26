# Office File Skills — Claude Code

Seven Claude Code skills for reading, editing, and processing Office files (`.xlsx`, `.docx`, `.pptx`) and PDFs.

---

## Quick Install (new machine)

```bash
git clone https://github.com/leonatez/claude-skill-read-excel.git
cd claude-skill-read-excel
bash install.sh
```

`install.sh` does two things:
1. Installs all Python dependencies (pinned versions from `requirements.txt`)
2. Copies each `SKILL.md` into `~/.claude/skills/<skill-name>/`
---

## Usage

In Claude Code, type the slash command with the file path:

```
/read-excel /path/to/file.xlsx
/edit-excel /path/to/file.xlsx   add a new sheet called "Summary" with totals
/read-docx  /path/to/file.docx
/edit-docx  /path/to/file.docx   add a new section "Process Flow" with mermaid diagram after paragraph 29
/read-pptx  /path/to/file.pptx
/edit-pptx  /path/to/file.pptx   add a new slide after slide 31 about terminal commands
```

---

## Uninstall

```bash
rm -rf ~/.claude/skills/read-excel \
       ~/.claude/skills/edit-excel \
       ~/.claude/skills/read-docx \
       ~/.claude/skills/edit-docx \
       ~/.claude/skills/read-pptx \
       ~/.claude/skills/edit-pptx
```

## License

MIT
