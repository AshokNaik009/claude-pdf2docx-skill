// Central defaults for the deterministic-fast PDF → DOCX engine.
// Every value is overridable from the CLI (see cli.js parseFlags).
//
// This is the standalone "deterministic-fast" engine extracted from the larger
// multi-engine repo: extract facts (PyMuPDF) → clean pages built with ZERO LLM;
// complex pages have the model emit a COMPACT LAYOUT JSON (not docx.js code,
// ~2× faster) that a deterministic renderer expands while STAMPING the exact
// fills/borders/frame recovered from the PDF's vector layer. Verified without
// LibreOffice by extractandVerify.py.
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
  model: "sonnet",         // `claude --model` for reconstruction
  escalationModel: "opus", // used ONLY when the default model still fails verification
  concurrency: 4,          // parallel `claude -p` workers
  claudeTimeoutMs: 240000, // per-page worker timeout
  retries: 1,              // extra attempts for a failed page
  verifyRetries: 1,        // extra in-place regenerations when the zero-LLM verifier rejects a page
  coverageMin: 0.6,        // deterministic-path fidelity gate (src/verify.js coverage())

  // ---- Heading detection --------------------------------------------------
  headingRatio: 1.15,      // line size ≥ bodySize * ratio ⇒ candidate heading

  pageBreakBetween: true,  // insert a page break between pages in the DOCX

  // ---- Compact-JSON model path (the ~2×-faster architecture) --------------
  // When true (default), complex pages are reconstructed by having the model
  // emit a small LAYOUT JSON instead of ~6KB of docx.js code — the model output
  // length was measured to be the real bottleneck (~136s → ~63s). A
  // deterministic renderer (src/docx.js renderLayoutJson) expands the JSON and
  // stamps the injected fills/borders/frame. If the JSON path can't pass
  // verification, the page falls back to the docx.js code-gen path, then to the
  // page image — so this never ships a worse result than the legacy path.
  compactJson: true,

  // ---- Deterministic Python extraction (facts up front) -------------------
  // Source each page's structured text + geometry + VECTOR DRAWINGS (fills,
  // borders, frame) from pdfextract.py (PyMuPDF). The text becomes a compact
  // prompt for the model (no vision tokens); the drawings are INJECTED into the
  // renderer, never decided by the model. Runs on the system Python that has
  // PyMuPDF, NOT a docling venv.
  pyExtract: true,
  pyExtractBin: "python3",                   // interpreter that has PyMuPDF (fitz)
  pyExtractScript: "pdfextract.py",

  // ---- Soffice-free verification gate (extractandVerify.py) ---------------
  // After the merge, inspect the OUTPUT .docx deterministically (no LibreOffice)
  // and check the injected fills actually landed + word coverage held. Reports
  // mismatches instead of shipping silently. Disable with --no-gate.
  gate: true,
  gateBin: "python3",
  gateScript: "extractandVerify.py",
};
