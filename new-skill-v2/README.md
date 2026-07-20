# pdf-to-docx

Convert a PDF into a **faithful, editable Word (`.docx`)** document — preserving
multi-column layouts, tables, headers/footers, page numbers, and figures.

The project is a pipeline with **four interchangeable output engines**. They
evolved over time (Markdown → HTML/pandoc → vision-LLM code-gen → Docling); all
four still work, but they are not equal. This README gives an honest assessment
of each so you (or an agent) know which to reach for.

---

## TL;DR — which engine should I use?

```bash
node index.js all <input.pdf> --docling      # ⭐ deterministic, no LLM, ~seconds — best for structured docs
node index.js all <input.pdf>                # default: per-page vision-LLM code-gen (needs `claude` CLI)
node index.js all <input.pdf> --pandoc       # semantic HTML → pandoc (good for styled CVs)
node index.js all <input.pdf> --markdown     # GFM → markdown-docx (simplest, lowest-fidelity — legacy)
```

| Engine | Flag | How it works | LLM? | Status | Best for |
|---|---|---|---|---|---|
| **Docling** | `--docling` | IBM Docling (Python) parses the whole PDF → `doc.json`; `src/doclingcode.js` generates editable `docx.js` modules from element geometry. | ❌ none | ⭐ **Primary / most developed** | Structured docs: papers, credit memos, tables, multi-column, quadrant layouts. Deterministic & fast. |
| **Vision code-gen** | *(none)* | Each complex page image → `claude -p` writes a `docx.js` `buildPage()` module (guided by the `docxjs` skill). Clean text pages are reconstructed deterministically and never hit the LLM. | ✅ per page | **Active** | Pages Docling mis-parses; visually complex / scanned pages where you want the model's judgment. |
| **Pandoc HTML** | `--pandoc` | Each page → `claude -p` emits semantic HTML; `pandoc` builds the `.docx`. A reference doc + Lua filter map CSS-ish classes to real Word styles (sidebar/intro shading). | ✅ per page | **Active** (niche) | Styled CVs / resumes with colored sidebars. |
| **Markdown** | `--markdown` | Each page → `claude -p` emits pure GFM; `markdown-docx` builds the `.docx`. | ✅ per page | ⚠️ **Legacy / lowest fidelity** | Simple, mostly-linear text docs. Superseded by the above; kept because it still works. |

> **Honest note:** `--markdown` was the earliest approach and drops the most
> structure (no real multi-column, no rich table styling). `--pandoc` improved on
> it (HTML structure survives, but pandoc ignores CSS/colors). `--docling` is the
> newest and, for structured documents, the most reliable — and it costs **$0 /
> zero tokens**. Start with `--docling`; fall back to the vision engine for pages
> it can't parse.

---

## Setup

### 1. Node dependencies (required for all engines)
```bash
npm install          # docx, markdown-docx, mupdf
```

### 2. The `claude` CLI (required only for the LLM engines: default, --pandoc, --markdown)
The LLM engines shell out to `claude -p` (Claude Code headless). Make sure the
`claude` CLI is installed and authenticated. **`--docling` does not need it.**

### 3. Docling (required only for `--docling`)
```bash
python3.12 -m venv .venv
.venv/bin/pip install docling      # pulls torch etc (~GB, one-time)
```
The venv path is `config.pythonBin` (`.venv/bin/python`). First run downloads
Docling's layout/table model weights (one-time). **Full install & verification
steps: [`DOCLING_SETUP.md`](DOCLING_SETUP.md).**

### 4. pandoc (required only for `--pandoc`)
```bash
brew install pandoc
```

---

## Usage

The pipeline is **two-phase**, so you can inspect/edit the intermediate work dir
before producing the final `.docx`:

```bash
node index.js run   <input.pdf>          # phase 1: parse → work/<name>/ (per-page modules + session.json)
node index.js merge <work/<name>>        # phase 2: assemble → outputs/<name>.docx
node index.js all   <input.pdf>          # both phases in one shot
node index.js <input.pdf>                # shorthand for `all`
```

### Flags (override `config.js`)
| Flag | Meaning |
|---|---|
| `--docling` / `--pandoc` / `--markdown` | Pick an output engine (mutually exclusive; omit = default vision code-gen). |
| `--fresh` | Ignore any prior session and reprocess every page from scratch. |
| `--llm auto\|all\|none` | Which pages go to the LLM (default `auto`: only low-confidence/complex pages). |
| `--model <name>` | Model for the LLM engines (default `sonnet`; `opus` used for escalation). |
| `--concurrency <n>` | Parallel `claude -p` workers (default 4). |
| `--dpi <n>` | Render DPI for page images/assets (default 200). |
| `--retries <n>` | Extra attempts per failed page. |
| `--out <name>` | Output filename stem. |
| `--work <dir>` | Working-dirs root (default `work/`). |
| `--ref <path>` | Override the pandoc reference `.docx` (`--pandoc` only). |

Outputs land in `outputs/<stem>.docx` (plus the intermediate `.md`/`.html`).

---

## Architecture

```
                       ┌─────────────── phase 1: run ───────────────┐   ┌──── phase 2: merge ────┐
  input.pdf ──▶ render pages (mupdf) ──▶ per-page reconstruction ──▶ work/<name>/  ──▶ assemble ──▶ outputs/<name>.docx
                                              │                        (page modules
                                              │                         + session.json)
        default / --pandoc / --markdown ──────┤ each complex page → `claude -p`
        --docling ────────────────────────────┘ whole PDF → Docling → doc.json → generated modules
```

- **Resumable:** `session.json` tracks each page's status/mode/tokens/timing. Re-running continues where it left off unless `--fresh`.
- **Deterministic-first (LLM engines):** pages with a healthy text layer are reconstructed from the PDF's own text/geometry and **never** sent to the model (see `config.minCharsForText`, `coverageMin`). Only complex/scanned pages hit `claude -p`.
- **Common merge:** every engine ultimately produces per-page content that `src/docx.js` (`renderDocx`) or the pandoc/markdown renderer packs into the final `.docx`.

### Source map (`src/`)
| File | Role |
|---|---|
| `pipeline.js` | Phase-1 orchestration: render, classify, per-page reconstruction, session state, parallel workers. |
| `pdf.js` / `png.js` | PDF rendering & structured text extraction (mupdf); PNG encoding. |
| `quality.js` | Per-page classification — clean-text (deterministic) vs complex (LLM). |
| `extract.js` / `layout.js` / `tables.js` / `geometry.js` | Deterministic reconstruction: figure/asset extraction, column/paragraph segmentation, table detection, bbox clustering. |
| `worker.js` | The LLM workers: builds prompts and runs `claude -p` for the code-gen / HTML / Markdown engines. |
| `docx.js` | **The shared final renderer** — executes generated `buildPage()` modules and packs the `.docx`; applies doc-level running header/footer + page-number start. |
| `doclingengine.js` | `--docling` entry: spawns venv Python extractor, then calls `doclingcode.js`. |
| `doclingcode.js` | **Docling → `docx.js` code generation** (layout X-Y-cut, tables, headers/footers). See below. |
| `htmldocx.js` | `--pandoc` engine: assemble HTML, run pandoc. |
| `mddocx.js` / `markdown.js` | `--markdown` engine: assemble GFM, run markdown-docx; GFM sanitizer. |
| `merge.js` / `session.js` / `verify.js` | Document model assembly; session load/save; zero-LLM structural verifier + coverage gate. |

### Scripts & assets
- `scripts/docling_extract.py` — PDF → `doc.json` + images (OCR disabled for text-layer PDFs; ~4s for 7 pages). Wired via `config.doclingScript`.
- `scripts/build-cv-ref.mjs` — regenerates `assets/cv-ref.docx`, the pandoc reference doc.
- `assets/cv-ref.docx`, `assets/cv-classmap.lua` — named Word styles + HTML-class→style map for `--pandoc`.
- `.claude/skills/docxjs/` — the authoring skill the vision code-gen engine reads (return contract, table/layout/image rules).

---

## The Docling engine in depth (`--docling`)

This is the most developed path. Given `doc.json` (Docling's whole-document
parse), `src/doclingcode.js` deterministically emits editable `docx.js` modules.
Everything below keys off **geometry and Docling's own semantic flags** — never
example-specific text — so it generalizes across document types.

**Running headers/footers + live page numbers** (`detectHeadersFooters`)
- Two-layer detection: (1) anything Docling tags `furniture` / `page_header` / `page_footer` is routed out of the body; (2) as a safety net for classifier misses, body text in the top/bottom 12% margin band whose *normalized* text repeats on ≥2 pages is also treated as running furniture.
- Page numbers become **live Word `PAGE` fields**; `pageNumberStart` is extrapolated from the actual footer values (e.g. footer "536" on sheet 6 ⇒ start 531), not from body text.

**Table fidelity** (from Docling's per-cell flags — it exposes structure but **no colors**)
- `column_header` → shaded header row + bold. `row_header` (a left label column) → shaded + bold. `row_section` → shaded divider row.
- Fills are light/readable defaults because the real colors are **not in Docling's data** (it drops fill/border/text color). Real colors only survive via the vision/`--pandoc` engines, which see the image.

**Layout reconstruction — geometry-driven X-Y cut** (`emitPageModule`)
- Elements are ordered **top→bottom by bbox**, *not* by Docling's reading order (which scrambles on 2-D pages).
- An element that spans the central axis (a centered title, a full-width table/figure, or a caption over one) is emitted **full-width** and breaks the surrounding band.
- Everything else is assigned to the left/right column by center-x and stacked in place — so a **column-width table stays nested inside its own column**. Verified on a 2×2 credit-memo "Four Quadrant Analysis" (title on top, two columns each flowing top→bottom with their financial tables nested) and on a 2-column journal paper.

**Known Docling-path limitations**
- **Dark section-bar fill** (e.g. a "THE COMPANY" bar) is not reproduced deterministically — Docling drops the color, and shading every heading grey would wrongly bar plain headers. Dark bars *are* handled on the vision/`--pandoc` engines.
- Figure captions Docling places **outside** `body.children` are not rendered (a Docling quirk, not a bug here).

---

## Gotchas & notes

- **`docx` fixed table layout:** a `Table` with `columnWidths` but no `layout: TableLayoutType.FIXED` + per-cell `width` makes Word autofit and collapse columns to ~1 char. Always pair them. (Codified in `doclingcode.js` and the `docxjs` skill's `tables.md`.)
- **macOS `~/Downloads` is TCC-blocked** for the CLI: `existsSync` passes but reads fail with `EPERM`. Run against PDFs inside the project directory.
- **`work/` and `outputs/` are regenerable** scratch/output — safe to delete; a run rebuilds them. `work/<name>/_docling/doc.json` is the Docling parse cache.
- **No git yet.** Dead/experimental files were moved to `archive/` (see `archive/README.md`) rather than deleted, so nothing is lost. Consider `git init` + a `.gitignore` for `node_modules/`, `work/`, `outputs/`, `.venv/`, `.DS_Store`.

---

## Extending this (for a developer or agent)

- **Add an output engine:** produce per-page content and hand it to a renderer. The cleanest contract is `src/docx.js`'s `buildPage(docx, ctx)` returning `{ sections, styles, header, footer, pageNumberStart }` — `renderDocx` already consumes all of these.
- **Improve the Docling path:** work in `src/doclingcode.js`. The layout logic is `emitPageModule` (X-Y cut) and `detectHeadersFooters`. Verify by deleting outputs and running `node index.js all <pdf> --docling --fresh`, then opening the `.docx` in Word (not just LibreOffice — LibreOffice's PDF export can hide table-layout bugs).
- **Tune the LLM engines:** prompts live in `src/worker.js` (`buildPrompt` / `buildHtmlPrompt` / `buildMarkdownPrompt`); authoring rules live in `.claude/skills/docxjs/`.
- **Sample PDFs in-repo:** `newtest-document.pdf` (2-column journal paper), `Four_Quadrant_Analysis.docx.pdf` (2×2 credit memo), `AshokNaik2026.pdf` (CV).
