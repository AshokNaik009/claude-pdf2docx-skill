---
name: pdf-to-docx
description: Convert PDF files into fully editable Word .docx documents that preserve the original layout, fonts, tables, and design. Use this skill WHENEVER the user wants a PDF converted to Word/docx/editable form — including phrases like "PDF to Word", "make this PDF editable", "convert invoice/report/contract to docx", or complaints that a previous conversion produced text boxes, frames, "Drawings", or broken formatting. Also use when the user wants to extract a PDF's layout/structure or rebuild a PDF-based template in Word. Do NOT use for docx→pdf (use LibreOffice directly) or for reading PDF content only.
---

# PDF → editable DOCX

Convert a PDF into a Word document made of **native paragraphs and tables** —
never text frames, never absolute positioning — while preserving the original
design. You are the semantic interpretation layer: deterministic scripts
extract and verify; you interpret the layout and write a bespoke renderer.

## Workflow

1. **Setup.** `pip install pymupdf pdfplumber pdf2docx --break-system-packages -q`
   (`--break-system-packages` is required in this container). Copy the input
   PDF from `/mnt/user-data/uploads/` to your working directory.

2. **Extract.** Run `python scripts/extract_layout.py input.pdf --out layout.json`.
   Read the per-page stdout summary (`page 1: 68 spans, 1 table, 2 drawings`)
   before opening the JSON.

3. **Scanned-page check.** If any page reports `has_text: false`, STOP and tell
   the user those pages are image-only and need OCR first:
   `ocrmypdf input.pdf ocr.pdf`, then restart from step 2 with `ocr.pdf`.
   Never silently produce empty pages.

4. **Rasterize.** Run `python scripts/rasterize.py input.pdf --dpi 100` and
   `view` every page image. These images are the ground-truth design reference.

5. **Triage** (your judgment, per document):
   - **Simple flowing text** (letters, essays, plain reports) → quick mode:
     ```python
     from pdf2docx import Converter
     cv = Converter("input.pdf"); cv.convert("output.docx"); cv.close()
     # page ranges: cv.convert(out, start=0, end=None) — 0-indexed start, exclusive end
     ```
     Jump straight to step 8. If verification passes, you are done — massively cheaper.
   - **Complex layout** (invoices, forms, multi-column, heavy tables,
     decorative elements) → full hybrid path, steps 6–8.

6. **Interpret.** Read `layout.json` alongside the page images and derive the
   semantic structure. Write it down as a brief plan before coding:
   - Title hierarchy from `size`/`bold` (e.g., a 32pt bold span is the title).
   - Column groups from x-clustering of spans (e.g., a 3-column address block
     clusters at x≈44/218/392).
   - Tables from the `tables` array (bbox + cell grid).
   - Alignment intent: right-aligned numbers have ragged left `x` but aligned
     right edges — **cluster on `x1`, not `x`**.
   - Decorative rules and fills from `drawings` (color, position, width).

7. **Render.** Write a bespoke Node script using the preinstalled `docx` npm
   package (require it directly; only `npm install docx` if the require fails).
   Rebuild the document with NATIVE Word elements. Apply every gotcha in the
   reference section below. Set the document's default font to
   **"Plus Jakarta Sans"** (the skill's standard output font) unless the user
   asks to preserve the PDF's original fonts. For multi-page PDFs, interpret
   and render page by page, inserting a `PageBreak` between pages.

8. **Verify loop — mandatory, never skip.** Run
   `bash scripts/verify.sh output.docx input.pdf` — always pass the original
   PDF so the script validates the page count; a mismatch exits non-zero and
   means content is overflowing or underfilling pages. Then `view` the
   rendered JPGs side by side against the step-4 originals. Check every page,
   not just page 1: title size/weight, column count and widths, table
   alignment, decorative rules, totals-block alignment, signature layout,
   page breaks. Fix the renderer and re-verify until faithful.
   Typical: 1–3 iterations.

9. **Deliver.** Copy the final `.docx` to `/mnt/user-data/outputs/` and present
   it. Summarize what is editable (real paragraphs, N native tables) and any
   conscious deviations (e.g., font substituted because the original wasn't
   embedded).

## docx-js gotchas (apply all of these)

- **Page size** defaults to A4. For US Letter set
  `page: { size: { width: 12240, height: 15840 } }` (DXA units; 1440 = 1 inch).
- **Tables need dual widths**: `columnWidths` on the `Table` AND `width` on
  every cell, both `WidthType.DXA`. `PERCENTAGE` breaks in Google Docs.
  Column widths must sum to the table width.
- **Table shading**: `ShadingType.CLEAR`, never `SOLID` (renders black).
- **Bullets**: never insert `•` literally — use a numbering config with
  `LevelFormat.BULLET`.
- **Line breaks**: never use `\n` — emit separate `Paragraph` elements.
- **`PageBreak`** must live inside a `Paragraph`.
- **Horizontal rule** = paragraph bottom border, NOT a table.
- **Borderless layout tables** (header/address/totals grids) need every border
  AND `insideHorizontal`/`insideVertical` set to `BorderStyle.NONE`.
- **Fonts**: default the document to "Plus Jakarta Sans" (set it as the
  default run font in `styles`). If the user asks to preserve original fonts,
  use the PDF's font family names (e.g., "Roboto") instead; Word falls back
  gracefully when a font isn't installed. Note substitutions to the user.

## Fallback: pixel-perfect over editable

If the user explicitly prioritizes pixel-perfect visuals over editability,
offer the LibreOffice PDF import instead:

```bash
soffice --headless --infilter="writer_pdf_import" \
  --convert-to docx:"MS Word 2007 XML" --outdir out/ input.pdf
```

**Warn them first**: every text run becomes an absolutely-positioned frame —
uneditable "Drawing" objects in Google Docs and Word. This is inherent to the
filter; there is no flag that produces flowing text.

## Environment notes

- LibreOffice (`soffice`) and Poppler (`pdftoppm`) are preinstalled.
- The `docx` npm package is preinstalled — require it directly.
- If a parallel/crashed soffice instance makes conversion hang, `verify.sh`
  already isolates the profile via `-env:UserInstallation`; do the same for
  any manual soffice call, and wrap it in `timeout 120`.
