// Deterministic multi-column layout reconstruction (no LLM). Splits a page's
// text blocks into "columns" segments (N side-by-side bands, read entire band
// then entire next band — never row-interleaved) and "full" segments (a
// full-width interruption like a table or figure caption spanning the whole
// content width), matching the two/N-column-then-full-width-then-columns
// pattern the docxjs skill's layout.md documents on the emit side.
//
// mupdf sometimes emits ONE BLOCK PER VISUAL LINE rather than per paragraph
// (verified on real academic-paper PDFs) — reconstructing straight from raw
// blocks would insert a spurious paragraph break after every line. So each
// column band's blocks are re-merged into paragraphs first, using a gap
// threshold relative to that band's OWN measured line height (not a fixed
// constant — line height varies by font size across documents).
import { clusterByGap, nearestBandIndex } from "./geometry.js";
import { isRowShaped } from "./tables.js";

const FULL_WIDTH_RATIO = 1.6;  // a block wider than this * median column width is "full width"
const PARA_GAP_FACTOR = 1.8;   // a gap bigger than this * the band's typical line gap breaks a paragraph
const MIN_BAND_BLOCKS = 6;     // fewer narrow blocks than this ⇒ not enough signal to call it multi-column
const MIN_MEDIAN_TEXT_LEN = 40; // chars; below this the page reads as short entries, not column prose
const MIN_BAND_POPULATION = 3;  // fewer blocks than this at a clustered x ⇒ noise, not a real column

const hasText = (b) => (b.lines || []).some((ln) => (ln.text || "").trim());
const modalSize = (b) => {
  const counts = new Map();
  for (const ln of b.lines || []) {
    if (!ln.font) continue;
    counts.set(ln.font.size, (counts.get(ln.font.size) || 0) + (ln.text || "").length);
  }
  let best = null, max = -1;
  for (const [size, w] of counts) if (w > max) { max = w; best = size; }
  return best;
};

/**
 * @returns {{bands:{x0:number,x1:number}[], fullWidthMin:number} | null}
 *   null means the page doesn't show enough multi-column signal.
 */
function blockTextLen(b) {
  return (b.lines || []).map((ln) => ln.text || "").join(" ").trim().length;
}

function detectColumnBands(blocks, cfg) {
  const textBlocks = blocks.filter((b) => b.type === "text" && b.bbox && hasText(b));
  if (textBlocks.length < MIN_BAND_BLOCKS) return null;

  // A borderless table (no same-y anchor row for tables.js to key off — a
  // "company | lender | amount" summary where every cell wraps to several
  // short lines) can still cluster into x-aligned bands and get mistaken for
  // genuine column-body prose. Real prose blocks run long (median ~54 chars
  // measured on a real academic PDF); this table's cells measured 19-34.
  // Below this, it isn't page-column text — leave it single-column rather
  // than confidently reordering table cells as if they were columns of prose.
  const lens = textBlocks.map(blockTextLen).sort((a, b) => a - b);
  if (lens[Math.floor(lens.length / 2)] < MIN_MEDIAN_TEXT_LEN) return null;

  const widths = textBlocks.map((b) => b.bbox.w).sort((a, b) => a - b);
  const median = widths[Math.floor(widths.length / 2)];
  const fullWidthMin = median * FULL_WIDTH_RATIO;

  // Row-shaped blocks (multiple same-y, x-separated lines — tabular/tab-stop
  // content, see tables.js) must not VOTE on where the column bands are: a
  // table's cell x-positions or a résumé's "role ... dates" alignment look
  // exactly like column starts otherwise. The primary defense is that the
  // caller runs table detection first and removes real table blocks entirely
  // before this ever sees them; this filter is a second layer for tab-stopped
  // lines that don't reach a table region's own row/coherence thresholds.
  const narrow = textBlocks.filter((b) => b.bbox.w < fullWidthMin && !isRowShaped(b));
  if (narrow.length < MIN_BAND_BLOCKS) return null;

  // A single stray block (a centered caption between two real columns) is
  // enough to fracture gap clustering into a spurious extra "column" — a real
  // column is populated by many blocks, not one. Drop bands too sparse to
  // trust; those blocks get reassigned to their nearest surviving band later
  // (nearestBandIndex, in the segmentation step), not dropped from the page.
  let bands = clusterByGap(narrow.map((b) => b.bbox.x), cfg.columnGapPts);
  const population = bands.map((band) => narrow.filter((b) => b.bbox.x >= band.x0 - 1 && b.bbox.x <= band.x1 + 1).length);
  bands = bands.filter((_, i) => population[i] >= MIN_BAND_POPULATION);
  if (bands.length < 2) return null;
  return { bands, fullWidthMin };
}

/**
 * Re-merge adjacent same-band blocks into paragraphs. Blocks merge only if
 * their vertical gap looks like an ordinary line break (relative to that
 * band's own measured line spacing) AND their dominant font size matches
 * (so a heading immediately followed by body text is never swallowed into
 * the paragraph that follows it).
 */
export function mergeParagraphs(blocksInYOrder) {
  if (!blocksInYOrder.length) return [];
  const gaps = [];
  for (let i = 1; i < blocksInYOrder.length; i++) {
    gaps.push(blocksInYOrder[i].bbox.y - (blocksInYOrder[i - 1].bbox.y + blocksInYOrder[i - 1].bbox.h));
  }
  gaps.sort((a, b) => a - b);
  const typicalGap = gaps.length ? gaps[Math.floor(gaps.length / 2)] : 4;
  const breakThreshold = Math.max(typicalGap * PARA_GAP_FACTOR, 6);

  const merged = [];
  let cur = null;
  for (const b of blocksInYOrder) {
    if (cur) {
      const gap = b.bbox.y - (cur.bbox.y + cur.bbox.h);
      const sameSize = modalSize(cur) === null || modalSize(b) === null ||
        Math.abs(modalSize(cur) - modalSize(b)) < 0.5;
      if (gap <= breakThreshold && sameSize && !b._breakBefore) {
        cur.lines.push(...(b.lines || []));
        cur.bbox = {
          x: Math.min(cur.bbox.x, b.bbox.x),
          y: Math.min(cur.bbox.y, b.bbox.y),
          w: Math.max(cur.bbox.x + cur.bbox.w, b.bbox.x + b.bbox.w) - Math.min(cur.bbox.x, b.bbox.x),
          h: Math.max(cur.bbox.y + cur.bbox.h, b.bbox.y + b.bbox.h) - Math.min(cur.bbox.y, b.bbox.y),
        };
        continue;
      }
    }
    cur = { type: "text", bbox: { ...b.bbox }, lines: [...(b.lines || [])] };
    merged.push(cur);
  }
  return merged;
}

/**
 * Segment a page into "columns" (N-band) and "full" (full-width) runs, in
 * reading order. Each "columns" segment's `bands` is an array (one per band,
 * left to right) of already paragraph-merged pseudo-blocks — read the whole
 * first band, then the whole next band, never interleaved line by line.
 *
 * `blocks` should have table-region blocks already removed by the caller
 * (src/tables.js findTableRegions) — a table's own cell x-positions look
 * exactly like column starts otherwise and will be misread as page layout.
 * @param {object[]} blocks  reading-order-sorted, non-table text blocks
 * @returns {Array<{type:"columns",count:number,bands:object[][]} | {type:"full",blocks:object[]}> | null}
 *   null means the page isn't multi-column — caller should use the plain
 *   single-column path instead.
 */
export function buildLayoutSegments(blocks, cfg) {
  const raw = blocks.filter((b) => b.type === "text" && b.bbox && hasText(b));
  raw.sort((a, b) => (Math.abs(a.bbox.y - b.bbox.y) > 4 ? a.bbox.y - b.bbox.y : a.bbox.x - b.bbox.x));

  const colInfo = detectColumnBands(raw, cfg);
  if (!colInfo) return null;
  const { bands, fullWidthMin } = colInfo;

  const runs = [];
  for (const b of raw) {
    const isFull = b.bbox.w >= fullWidthMin;
    const kind = isFull ? "full" : "columns";
    const last = runs[runs.length - 1];
    // `_breakBefore` (set by extract.js when a table region was removed right
    // before this block) forces a new run even if the kind matches the
    // previous one — otherwise text from well below a carved-out table would
    // silently merge with text from above it into one column run.
    if (last && last.kind === kind && !b._breakBefore) last.blocks.push(b);
    else runs.push({ kind, blocks: [b] });
  }

  return runs.map((run) => {
    if (run.kind === "full") {
      return { type: "full", blocks: mergeParagraphs(run.blocks) };
    }
    const perBand = bands.map(() => []);
    for (const b of run.blocks) perBand[nearestBandIndex(b.bbox.x, bands)].push(b);
    return {
      type: "columns",
      count: bands.length,
      bands: perBand.map((blocksInBand) =>
        mergeParagraphs([...blocksInBand].sort((a, b) => a.bbox.y - b.bbox.y))
      ),
    };
  });
}
