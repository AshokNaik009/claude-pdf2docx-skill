# claude-pdf2docx-skill

A Claude skill that converts PDFs into **fully editable** Word `.docx` files that preserve the original layout, fonts, tables, and design — native paragraphs and tables, never text frames or "Drawing" objects.

## How it works

A hybrid pipeline: deterministic scripts for extraction and verification, with Claude (the agent executing the skill) as the semantic interpretation layer in the middle.

1. **Extract** — `scripts/extract_layout.py` dumps every text span (`text, font, size, bold, x, y`), table cell grid, and vector drawing (rules, fills) to JSON via PyMuPDF + pdfplumber.
2. **Interpret** — the agent reads the JSON alongside rasterized page images (`scripts/rasterize.py`) and infers the semantic structure: title hierarchy, column groups, table regions, alignment intent, decorative elements.
3. **Render** — the agent writes a bespoke docx-js (npm `docx`) script producing native Word tables and paragraphs. The output document defaults to the **Plus Jakarta Sans** font (original PDF fonts preserved on request).
4. **Verify** — `scripts/verify.sh` round-trips DOCX → PDF → JPG so the agent can visually diff against the original and iterate. Passing the original PDF also **validates the page count** of the rendered document (non-zero exit on mismatch).

Simple text-heavy PDFs are routed through `pdf2docx` quick mode instead; scanned (image-only) pages are detected and routed to OCR.

## Why not the obvious alternatives

- **LibreOffice `writer_pdf_import`** is pixel-faithful but wraps every text run in an absolutely-positioned frame — uneditable "Drawing" objects in Word/Google Docs. Kept only as an explicit pixel-perfect fallback.
- **`pdf2docx` alone** produces real paragraphs and tables but degrades formatting on complex layouts (column drift, mangled multi-column sections). Kept as quick mode for simple documents.

## Layout

```
pdf-to-docx/
├── SKILL.md              # workflow instructions for Claude
└── scripts/
    ├── extract_layout.py # PyMuPDF spans + pdfplumber tables → JSON
    ├── rasterize.py      # pdftoppm page images for visual reference
    └── verify.sh         # docx → pdf → jpg round-trip check
```

## Requirements

- Python: `pymupdf`, `pdfplumber`, `pdf2docx` (in the Claude container: `pip install ... --break-system-packages`)
- Poppler (`pdftoppm`) and LibreOffice (`soffice`) — preinstalled in the Claude container
- npm `docx` package — preinstalled in the Claude container
