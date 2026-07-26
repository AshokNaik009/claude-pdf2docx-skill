# detfast — deterministic-fast PDF → editable DOCX

A self-contained engine that turns a PDF into an **editable** Word document. It is
built around one principle from the project handoff:

> Facts we can extract deterministically (text, table grids, **fills, borders,
> the outer frame**) must be **injected**, never decided by the model. Given the
> fills, the model applies them 4/4; asked to infer the frame, it drops it. So we
> extract the facts up front and only ask the model for the cheap part.

This folder is **portable** — copy or move it anywhere, run `npm install`, and it
works on its own. It has no dependency on the parent multi-engine repo.

---

## Pipeline

```
PDF page
  → pdfextract.py (PyMuPDF)  ── text + coords + VECTOR DRAWINGS (fills/borders/frame)
  → route:
       clean text/tables      → deterministic block builder  (ZERO LLM)
       fills / novel layout    → claude -p emits a COMPACT LAYOUT JSON (not code)
  → deterministic renderer     → docx, STAMPING the injected fills/borders/frame
  → merge → .docx
  → extractandVerify.py gate   → confirms fills landed, no table overflows its
                                 container, images embedded, word coverage (no LibreOffice)
```

**Why compact JSON, not docx.js code?** The measured bottleneck was *output
length* — the model writing ~6 KB of docx.js code took ≈136 s/page. Emitting a
small layout JSON that a deterministic renderer expands is ~2× faster (~63 s),
and the fills/borders/frame are stamped by the host, so the model never spends
tokens on them (or gets a chance to be wrong).

If the JSON path can't pass verification, a page falls back to docx.js code-gen,
then to the page image — so it never ships a worse result than the legacy path.

---

## Requirements

- **Node** ≥ 18 (`npm install` pulls `docx` + `mupdf` — both pure JS/WASM, no native build).
- **Python 3.9+** with **PyMuPDF** on the interpreter named by `pyExtractBin`
  (default `python3`):  `python3 -m pip install pymupdf`
- **`claude` CLI** on `PATH` — only needed for complex pages (the compact-JSON /
  code-gen paths). A `--llm none` run needs no `claude`.
- `extractandVerify.py` (the gate) is **pure stdlib** — nothing to install.

---

## Usage

```bash
npm install                       # first time only

# Full run (deterministic pages + model for complex pages):
node cli.js all input.pdf --out mydoc

# Deterministic only — no claude calls at all:
node cli.js all input.pdf --out mydoc --llm none

# Two-phase (inspect the per-page work dir between steps):
node cli.js run   input.pdf
node cli.js merge work/input
```

Output → `outputs/mydoc.docx`. The final step prints the soffice-free
verification gate. Re-verify any docx manually:

```bash
python3 extractandVerify.py outputs/mydoc.docx            # human summary
python3 extractandVerify.py outputs/mydoc.docx --json report.json
```

### Flags

| Flag | Default | Meaning |
|------|---------|---------|
| `--out NAME` | pdf stem | output file stem |
| `--llm auto\|all\|none` | `auto` | which pages go to the model |
| `--model M` | `sonnet` | `claude --model` for reconstruction |
| `--concurrency N` | 4 | parallel page workers |
| `--dpi N` | 200 | render DPI for image assets / fallbacks |
| `--no-compact-json` | — | disable the fast path; use docx.js code-gen |
| `--no-py-extract` | — | use the Node mupdf binding instead of pdfextract.py |
| `--no-gate` | — | skip the extractandVerify.py output check |
| `--fresh` | — | ignore any prior session; reprocess every page |

All defaults live in `config.js`.

---

## Layout-JSON schema (model output)

The model emits **only** this (transcription + block assembly — no colors, no frame):

```jsonc
{
  "title": "…",                 // optional
  "frame": false,               // optional; host injects the real frame anyway
  "styles": { "font": "…", "size": 22 },   // size = half-points (22 = 11pt)
  "columns": [                  // one object per side-by-side column
    { "blocks": [ /* Block, … */ ] }
  ]
}
```

`Block` is one of:

```jsonc
{ "t":"bar",  "text":"…" }                                  // full-width colored band (host fills the color)
{ "t":"h",    "text":"…", "level":1 }
{ "t":"p",    "text":"…" }
{ "t":"list", "ordered":false, "items":["…"] }
{ "t":"table","header":true, "cols":[2,1], "rows":[ ["A","B"], ["…","…"] ] }
{ "t":"img",  "file":"img_1.png" }
```
Table cells are a string or `{ "text":"…", "bold":true, "align":"right" }`.
Numeric columns are auto right-aligned; `header` (row 0) defaults true.

---

## How injection works

`pdfextract.py page_drawings()` returns, per page:

```jsonc
{
  "fills":   [ { "fill":"#3C78D8", "bbox":{…}, "text":"THE COMPANY" }, … ],
  "borders": [ { "color":"#000000", "width":1.0, "count":71 } ],
  "frame":   { "color":"#000000", "width":1.0, "bbox":{…} }   // or null
}
```

`src/pipeline.js buildInjected()` turns that into `{ fills, border, frame }`, stored
in each page manifest. `src/docx.js renderLayoutJson()` then:

- binds a `bar`/table-header fill by matching the block's words to the words a
  fill sits behind (`matchFill`), and picks contrasting text automatically;
- draws table gridlines at the **injected** stroke weight (docx `size` is in
  eighths of a point — 1 pt → `size: 8`);
- draws the **outer frame** whenever the page is ruled — an explicit frame rect,
  or (on a multi-column page) a dominant border stroke — as an outer layout table
  with the border + `insideVertical` rule. That ruling *is* the frame the model
  kept dropping.

**Widths are sized to the container, not the page.** Every nested table/bar/image
is built at its column's width (`columnWidth − padding`), so nothing overflows and
gets clipped off the right edge. Images are also capped to their column width.

**No figure is dropped.** Any extracted image asset the model didn't reference is
appended as an `img` block (`ensureImages`), so a figure is never lost even if the
model omits it.

---

## File map

| File | Role |
|------|------|
| `cli.js` | entry point (`run` / `merge` / `all`) |
| `config.js` | all defaults |
| `pdfextract.py` | **PyMuPDF** facts: structured text + `text_prompt` + `drawings` (fills/borders/**frame**) |
| `extractandVerify.py` | soffice-free `.docx` inspector — the verify gate (recurses nested tables) |
| `src/pipeline.js` | per-page orchestration + routing + the compact-JSON→code-gen→image ladder |
| `src/worker.js` | `claude -p` workers: `generatePageJson` (fast), `generatePageCode*` (fallback) |
| `src/docx.js` | `renderLayoutJson` (compact-JSON + injection), `executeModule`, `renderDocx` |
| `src/gate.js` | runs `extractandVerify.py`, checks injected fills landed + word coverage |
| `src/extract.js`, `tables.js`, `layout.js`, `geometry.js` | deterministic (zero-LLM) reconstruction |
| `src/verify.js` | zero-LLM structural verifier + word-coverage fidelity check |
| `src/pdf.js`, `png.js`, `quality.js`, `session.js`, `merge.js` | rendering, routing, state, ordering |
| `.claude/skills/docxjs/` | the docx.js authoring skill the workers read |

---

## Status vs. the handoff's open steps

- ✅ Compact-JSON worker + deterministic renderer wired as a real pipeline path (`--compact-json`, on by default).
- ✅ Frame + fills injected deterministically from `get_drawings()` (frame detection added to `pdfextract.py`).
- ✅ `extractandVerify.py` gate wired end-to-end (extended to recurse nested tables so injected fills are actually seen).
- ⏳ Multi-page wall-clock on a real doc still worth measuring with live `claude` (the machinery is verified; the numbers depend on your model/quota).
- ⏳ Targeted fixes from the gate's diff (today the gate **reports** mismatches; it does not yet re-drive individual pages).
