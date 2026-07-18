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
  concurrency: 4,          // parallel `claude -p` workers
  claudeTimeoutMs: 240000, // per-page worker timeout
  retries: 1,              // extra attempts for a failed page

  // ---- Heading detection --------------------------------------------------
  headingRatio: 1.15,      // line size ≥ bodySize * ratio ⇒ candidate heading

  pageBreakBetween: true,  // insert a page break between pages in the DOCX
};
