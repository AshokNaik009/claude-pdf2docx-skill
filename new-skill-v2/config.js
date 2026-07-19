// Central defaults for the PDF → structured-model → DOCX pipeline.
// Every value is overridable from the CLI (see index.js parseFlags).
export const CONFIG = {
  workRoot: "work",        // per-document working dirs live under here
  outputsDir: "outputs",   // final .docx / .md land here

  dpi: 200,                // render DPI for page images & cropped assets

  // ---- Quality check (which pages go to the LLM) --------------------------
  // Pages with a healthy text layer are reconstructed deterministically and
  // NEVER hit Claude. Only low-confidence / complex pages are sent.
  minCharsForText: 90,     // fewer chars than this ⇒ treat as scanned ⇒ LLM
  columnGapPts: 40,        // horizontal gap (pt) that signals a real column split
  llmMode: "auto",         // auto | all | none  (force every / no page to the LLM)

  // ---- Headless Claude workers -------------------------------------------
  model: "sonnet",         // `claude --model` for reconstruction (diagram: Sonnet)
  escalationModel: "opus", // used ONLY when Sonnet still fails verification after
                            // verifyRetries attempts — reserved for the rare hardest pages
  concurrency: 4,          // parallel `claude -p` workers
  claudeTimeoutMs: 240000, // per-page worker timeout
  retries: 1,              // extra attempts for a failed page
  verifyRetries: 1,        // extra in-place regenerations when the zero-LLM verifier
                            // (src/verify.js) rejects a generated page's structure
  coverageMin: 0.6,        // deterministic-path fidelity gate (src/verify.js coverage()):
                            // below this fraction of the page's words reconstructed,
                            // escalate the page to vision instead of shipping it

  // ---- Heading detection --------------------------------------------------
  headingRatio: 1.15,      // line size ≥ bodySize * ratio ⇒ candidate heading

  pageBreakBetween: true,  // insert a page break between pages in the DOCX

  // ---- Markdown output strategy (--markdown) ------------------------------
  // When true, complex pages are reconstructed by having Claude emit pure GFM
  // Markdown from the page image (NOT docx.js code), and the final .docx is
  // produced by running the whole document through the markdown-docx library.
  // markdown-docx is a CommonMark+GFM converter — it does NOT render raw HTML
  // tags (they'd appear as literal text), so the worker is constrained to pure
  // GFM and a defensive sanitizer strips any HTML that slips through.
  markdown: false,

  // ---- Pandoc HTML output strategy (--pandoc) ----------------------------
  // When true, complex pages are reconstructed as semantic HTML (multi-column
  // layouts expressed as <table> so columns survive), and the final .docx is
  // produced by the pandoc CLI (`pandoc -f html -t docx`). Unlike markdown-docx,
  // pandoc parses HTML structure and embeds <img> assets as real Word media —
  // it only drops CSS styling (colors, flex layout).
  pandoc: false,

  // Optional pandoc reference .docx (named styles: heading colors, fonts, the
  // Sidebar/IntroBox shaded styles). Used only in --pandoc mode when it exists.
  // Override with --ref <path>. Build/refresh it with scripts/build-cv-ref.mjs.
  pandocRefDoc: "assets/cv-ref.docx",
  // Optional pandoc Lua filter mapping HTML classes → the named styles above.
  pandocLuaFilter: "assets/cv-classmap.lua",

  // ---- Docling output strategy (--docling) -------------------------------
  // When true, the whole PDF is parsed by Docling (IBM, Python — its own
  // layout/table models) into doc.json, then per-page docx.js buildPage modules
  // are generated from it (src/doclingcode.js) and the normal merge builds the
  // .docx. This bypasses the per-page vision-LLM loop entirely: no Claude calls,
  // deterministic, ~seconds. Two-column layout is reconstructed from element
  // bboxes. Requires the project venv with `docling` installed.
  docling: false,
  pythonBin: ".venv/bin/python",             // venv python that has docling
  doclingScript: "scripts/docling_extract.py", // PDF → doc.json + images extractor
};
