// QUALITY CHECK (diagram box 2): decide whether a page can be reconstructed
// deterministically from its text layer, or whether it is low-confidence /
// complex and must be sent to headless Claude.
//
// Returns { mode: "text" | "llm", reason, stats }.

function lineSizes(blocks) {
  const sizes = [];
  for (const b of blocks) for (const ln of b.lines || []) if (ln.font) sizes.push(ln.font.size);
  return sizes;
}

/** Detect a genuine multi-column layout: two x-bands separated by a wide gap. */
function looksMultiColumn(blocks, pageWidth, gapPts) {
  const centers = [];
  for (const b of blocks) {
    if (b.type !== "text" || !b.bbox) continue;
    centers.push(b.bbox.x + b.bbox.w / 2);
  }
  if (centers.length < 6) return false;
  const mid = pageWidth / 2;
  const left = centers.filter((c) => c < mid - gapPts / 2);
  const right = centers.filter((c) => c > mid + gapPts / 2);
  // Both halves substantially populated ⇒ real columns (not just indentation).
  return left.length >= 3 && right.length >= 3 &&
    left.length / centers.length > 0.25 && right.length / centers.length > 0.25;
}

/**
 * Heuristic table detector: several rows whose short cells align on repeated
 * x-positions. Deliberately conservative — a false positive only means a clean
 * page is (harmlessly) reconstructed by the LLM instead.
 */
function looksTabular(blocks) {
  const allLines = [];
  for (const b of blocks) for (const ln of b.lines || []) allLines.push(ln.text || "");

  // Signal A — geometry: recurring column x-positions of short cells.
  const shortX = [];
  for (const b of blocks)
    for (const ln of b.lines || []) {
      const words = (ln.text || "").trim().split(/\s+/).filter(Boolean).length;
      if (words > 0 && words <= 4 && ln.bbox) shortX.push(Math.round(ln.bbox.x / 5) * 5);
    }
  const colCounts = new Map();
  for (const x of shortX) colCounts.set(x, (colCounts.get(x) || 0) + 1);
  const alignedCols = [...colCounts.values()].filter((n) => n >= 3).length;
  if (shortX.length >= 8 && alignedCols >= 2) return true;

  // Signal B — content: rows with 2+ numeric groups (financial tables collapse
  // into single lines like "Revenue 1,316,892 1,305,077 1,493,000").
  const numRow = /(\$?-?[\d,]*\.?\d+%?)(\s+\$?-?[\d,]*\.?\d+%?){1,}/;
  const numericRows = allLines.filter((t) => {
    const groups = (t.match(/-?[\d,]*\.?\d+%?/g) || []).filter((g) => /\d/.test(g));
    return groups.length >= 2 && numRow.test(t);
  }).length;
  return numericRows >= 3;
}

export function classifyPage(structured, pageWidth, cfg) {
  const blocks = structured.blocks || [];
  const chars = blocks.reduce(
    (n, b) => n + (b.lines || []).reduce((m, ln) => m + (ln.text || "").length, 0),
    0
  );
  const imageBlocks = blocks.filter((b) => b.type === "image").length;
  const stats = { chars, blocks: blocks.length, images: imageBlocks, sizes: lineSizes(blocks).length };

  if (cfg.llmMode === "all") return { mode: "llm", reason: "forced (--llm all)", stats };
  if (cfg.llmMode === "none") return { mode: "text", reason: "forced (--llm none)", stats };

  // Scanned pages have no usable text layer to deterministically reconstruct
  // from at all — straight to vision. Multi-column and tabular pages now go
  // through the deterministic-first path too (src/tables.js, src/layout.js):
  // pipeline.js verifies the result for free (src/verify.js) and escalates
  // to vision itself if the reconstruction doesn't hold up, so routing them
  // here no longer needs to assume they'll fail.
  if (chars < cfg.minCharsForText) return { mode: "llm", reason: "low text (scanned)", stats };
  if (looksMultiColumn(blocks, pageWidth, cfg.columnGapPts))
    return { mode: "text", reason: "multi-column (deterministic)", stats };
  if (looksTabular(blocks)) return { mode: "text", reason: "tabular region (deterministic)", stats };

  return { mode: "text", reason: "clean text layer", stats };
}
