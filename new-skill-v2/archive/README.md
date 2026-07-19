# archive/ — superseded & scratch files (safe to ignore, kept for reference)

Nothing in this folder is imported by the running code. It was moved here (not
deleted) on 2026-07-19 because the project has **no git history** yet — if
anything turns out to still be needed, restore it with `mv`.

Verified before moving: none of these are referenced anywhere under `index.js`,
`config.js`, `src/`, or `scripts/`. After moving, `node index.js all <pdf>
--docling` still builds correctly on both sample PDFs.

| Archived path | What it was | Why it's dead | Restore |
|---|---|---|---|
| `src/docling.js` | First-cut standalone Docling→docx renderer (`renderDoclingDocx` / `buildDoclingDocx`) — produced a single-column docx directly, no code-gen. | Superseded by `src/doclingcode.js`, which generates editable per-page `docx.js` modules and feeds the normal merge. Nothing imports `docling.js`. | `mv archive/src/docling.js src/` |
| `scripts/docling_noocr.py` | Early Docling extractor variant with OCR disabled. | `scripts/docling_extract.py` (the one wired into `config.doclingScript`) already sets `do_ocr=False`, so this is redundant. | `mv archive/scripts/docling_noocr.py scripts/` |
| `root-scratch/_htmltest.mjs` | One-off harness that hand-built a page and ran the `--pandoc` HTML→docx path. | Ad-hoc test scratch; never imported. | `mv archive/root-scratch/_htmltest.mjs .` |
| `root-scratch/_p2.mjs` | One-off harness that rendered PDF page 2 and ran the vision-LLM `generatePageCode` path. | Ad-hoc test scratch; never imported. | `mv archive/root-scratch/_p2.mjs .` |
| `root-scratch/_p2/p2.png` | Output image from `_p2.mjs`. | Scratch output. | `mv archive/root-scratch/_p2 .` |

If you later add git, you can delete this folder entirely — history will hold it.
