# Office File Skills — Claude Code

Eight Claude Code skills for reading, editing, and processing Office files (`.xlsx`, `.docx`, `.pptx`) and PDFs — with token-efficient workflows that replace exploratory trial-and-error with structured, script-backed pipelines.

---

## Why skills save tokens

Working with Office files without skills forces Claude to:
1. **Explore the binary format** through multiple trial-and-error turns
2. **Invent XML or API calls** that may be wrong on the first try
3. **Debug errors** that accumulate context and multiply cost

These skills eliminate that by:

| Mechanism | What it does | Token impact |
|---|---|---|
| **Inspection scripts** | Decode binary → structured text Claude can use directly | Replace 3–5 exploration turns with 1 bash call |
| **JSON spec pipeline** (`edit-docx`) | Claude writes JSON; script handles all OOXML | ~700 tokens → ~120 tokens for same edit |
| **Template pipeline** (`edit-pptx`) | `inventory.py → replacements.json → replace.py` | Claude writes JSON, never touches XML |
| **Typed Python API** (`edit-excel`) | openpyxl handles XML; Claude writes `ws["A1"] = value` | No XML in output at all |
| **Pitfalls tables** | Pre-loaded edge cases (XML escaping, index off-by-one) | Each avoided bug saves ~3k–5k debug tokens |
| **Mermaid rendering** | Diagrams rendered to PNG offline, embedded as images | No round-trips to explain diagram format |

---

## Quick install

```bash
git clone https://github.com/leonatez/claude-skill-read-excel.git
cd claude-skill-read-excel
bash install.sh
```

`install.sh` installs Python dependencies and copies each skill into `~/.claude/skills/<skill-name>/`.

**Optional system tools** for full feature support:

```bash
sudo apt-get install pandoc libreoffice poppler-utils
# macOS: brew install pandoc libreoffice poppler
# Node.js (for offline mermaid rendering): https://nodejs.org
```

---

## Skills

### Read skills

| Command | Use for |
|---|---|
| `/read-excel` | Inventory sheets, read data, merged cells, images, parse API specs from xlsx |
| `/read-docx` | Read and analyze Word documents (pandoc + python-docx) |
| `/read-pptx` | Read and analyze PowerPoint presentations |

```
/read-excel /path/to/file.xlsx
/read-docx  /path/to/file.docx
/read-pptx  /path/to/file.pptx
```

---

### Edit skills

#### `/edit-docx` — Edit Word documents

```
/edit-docx /path/to/file.docx   add a "Process Flow" section after paragraph 12 with a sequence diagram
```

**Workflows:**

- **JSON spec (recommended)** — Claude writes a compact JSON spec; `docx-apply-json-spec.py` handles all OOXML:

  ```json
  {
    "file": "doc.docx",
    "insert_after": 12,
    "elements": [
      {"type": "heading",   "text": "Process Flow", "level": 1},
      {"type": "paragraph", "text": "See diagram below.", "italic": true},
      {"type": "image",     "path": "/tmp/diagram.png", "width_inches": 5.5},
      {"type": "table",     "headers": ["Step","Actor","Action"], "rows": [["1","User","Submit"]]},
      {"type": "code",      "lines": ["POST /api/orders"]},
      {"type": "empty"}
    ]
  }
  ```

  Supported types: `paragraph`, `heading`, `image`, `table`, `code`, `empty`, `pagebreak`.

- **Tracked changes / redlining** — for legal/business review documents
- **Mermaid diagrams** — render flowchart/sequence/ERD to PNG and embed as inline image

---

#### `/edit-excel` — Edit Excel workbooks

```
/edit-excel /path/to/file.xlsx   add a Summary sheet with totals from all other sheets
```

- Inspects existing styles (fills, fonts, column widths) before writing
- Formula-first: always writes `=SUM(B2:B9)` not hardcoded values
- Recalculates formulas via LibreOffice after saving
- Supports embedding Mermaid diagrams into sheets

---

#### `/edit-pptx` — Edit PowerPoint presentations

```
/edit-pptx /path/to/file.pptx   add a slide after slide 5 about the deployment architecture
```

**Workflows:**

- **Direct slide editing** — inspect shapes/positions → write new slide with matching layout
- **Template bulk replace** — JSON-driven pipeline:
  ```
  inventory.py → inventory.json → replacements.json → replace.py → output.pptx
  ```
  Claude writes JSON, never XML.
- **Thumbnail grid** — visual overview of all slides
- **Mermaid diagrams** — render and embed as picture shapes on any slide

---

#### `/pdf` — Process PDF files

```
/pdf /path/to/file.pdf   extract all tables as CSV
```

Read, extract, merge, split, fill forms, and create PDFs.

---

## Mermaid diagram rendering

All three edit skills support rendering Mermaid diagrams to PNG and embedding them
directly in the document. The shared `mermaid/mermaid-render.py` script handles rendering:

```bash
python3 ~/.claude/skills/mermaid/mermaid-render.py \
  --input diagram.mmd \
  --output /tmp/diagram.png \
  --theme default   # default | dark | forest | neutral
```

**Rendering methods** (tried in order):
1. `npx @mermaid-js/mermaid-cli` — offline, requires Node.js
2. `mermaid.ink` public API — online fallback, no install needed

**Usage example in a PRD:**
```
/edit-docx product-spec.docx  add a "System Architecture" section with this sequence diagram:
sequenceDiagram
    User->>API: POST /order
    API->>DB: Insert
    DB-->>API: OK
    API-->>User: 201 Created
```

Claude generates the Mermaid script → renders to PNG → embeds as inline image.

---

## Project structure

```
.
├── edit-docx/
│   ├── SKILL.md                    # skill instructions
│   └── docx-apply-json-spec.py     # JSON spec → OOXML pipeline
├── edit-excel/
│   ├── SKILL.md
│   └── recalc.py                   # formula recalculation via LibreOffice
├── edit-pptx/
│   ├── SKILL.md
│   └── scripts/
│       ├── inventory.py            # extract slide content to JSON
│       ├── replace.py              # apply replacements.json to pptx
│       ├── rearrange.py            # reorder/duplicate slides
│       └── thumbnail.py            # generate visual slide grid
├── mermaid/
│   └── mermaid-render.py           # shared Mermaid → PNG renderer
├── ooxml/
│   └── scripts/                    # unpack/pack .docx for raw XML editing
├── read-docx/SKILL.md
├── read-excel/SKILL.md
├── read-pptx/SKILL.md
├── pdf/SKILL.md
├── requirements.txt
└── install.sh
```

---

## Uninstall

```bash
rm -rf ~/.claude/skills/read-excel \
       ~/.claude/skills/edit-excel \
       ~/.claude/skills/read-docx  \
       ~/.claude/skills/edit-docx  \
       ~/.claude/skills/read-pptx  \
       ~/.claude/skills/edit-pptx  \
       ~/.claude/skills/pdf        \
       ~/.claude/skills/mermaid
```

---

## License

MIT
