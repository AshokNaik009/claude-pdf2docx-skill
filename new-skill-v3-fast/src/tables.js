// Deterministic table detection + grid extraction from mupdf's structured-text
// blocks (no LLM, no vector/border signal — verified empirically absent from
// every real-world sample PDF this was built against; mupdf's `walk({onVector})`
// reported zero vector blocks across 6 real documents / ~12k text blocks, so
// table geometry here is inferred purely from text-line alignment).
//
// Two things make this tractable:
//  - mupdf frequently groups a table ROW's cells as sibling `lines` sharing one
//    y-coordinate within a single block (a "row-shaped" block).
//  - a row LABEL that wraps to two lines lands in its own block, separate from
//    its numeric row — so rows are assembled by vertical (y-range) overlap
//    across blocks, not by "one block = one row".
import { clusterByGap, nearestBandIndex } from "./geometry.js";

const MIN_ROWS = 3;         // fewer than this ⇒ not worth treating as a table
const ROW_GAP_MAX = 30;     // pt; gap beyond this breaks a candidate table run
const X_BUCKET = 8;         // pt; lines within this x tolerance are one cell (wrapped text)
const BAND_GAP = 25;        // pt; gap beyond this starts a new column band
// mupdf occasionally splits one justified prose line into word-level "lines" at
// the same y (measured ~8-9pt apart — ordinary inter-word spacing). Real table
// columns measured 53-167pt apart in sample documents. This threshold rejects
// the former as a false "row-shaped" match while accepting the latter.
const MIN_CELL_GAP = 25;
const NUMERIC = /^-?\$?[\d,]*\.?\d+%?$/;

const isBold = (font) => font?.weight === "bold";
const isItalic = (font) => font?.style === "italic";

function yRange(b) {
  return { top: b.bbox.y, bottom: b.bbox.y + b.bbox.h };
}

function yOverlaps(a, b) {
  const ra = yRange(a), rb = yRange(b);
  return ra.top < rb.bottom && rb.top < ra.bottom;
}

function verticalGap(a, b) {
  return Math.max(0, b.bbox.y - (a.bbox.y + a.bbox.h));
}

/** A block whose lines include >=2 entries at (nearly) the same y but distinct x. */
export function isRowShaped(block) {
  const lines = (block.lines || []).filter((ln) => ln.bbox && (ln.text || "").trim());
  if (lines.length < 2) return false;
  const byY = [];
  for (const ln of lines) {
    let grp = byY.find((g) => Math.abs(g[0].bbox.y - ln.bbox.y) <= 2);
    if (!grp) { grp = []; byY.push(grp); }
    grp.push(ln);
  }
  return byY.some((grp) => {
    if (grp.length < 2) return false;
    const sorted = [...grp].sort((a, b) => a.bbox.x - b.bbox.x);
    for (let i = 1; i < sorted.length; i++) {
      const prevEnd = sorted[i - 1].bbox.x + sorted[i - 1].bbox.w;
      if (sorted[i].bbox.x - prevEnd < MIN_CELL_GAP) return false; // word-spacing, not column spacing
    }
    return true;
  });
}

/**
 * Find maximal contiguous runs of table-like blocks in a reading-order block array.
 * @returns {{start:number, end:number}[]} index ranges (inclusive) into `blocks`
 */
export function findTableRegions(blocks) {
  const rowShaped = blocks.map(isRowShaped);
  const regions = [];
  let i = 0;
  while (i < blocks.length) {
    if (!rowShaped[i]) { i++; continue; }
    // Extend forward. Accumulate the row-in-progress's y-range across ALL
    // absorbed blocks (not just the most recent one) — a row whose cells wrap
    // to different heights in different columns (a 2-line label next to a
    // 4-line description) still needs to be recognized as one row.
    let end = i, rowCount = 1, j = i + 1;
    let accTop = blocks[i].bbox.y, accBottom = blocks[i].bbox.y + blocks[i].bbox.h;
    while (j < blocks.length) {
      const b = blocks[j];
      const bTop = b.bbox.y, bBottom = b.bbox.y + b.bbox.h;
      if (rowShaped[j]) {
        end = j; rowCount++; j++;
        accTop = bTop; accBottom = bBottom; // a fresh row-shaped block starts a new accumulation
        continue;
      }
      const overlapsAcc = bTop < accBottom && accTop < bBottom;
      const continuesAcc = bTop >= accBottom && bTop - accBottom < ROW_GAP_MAX;
      if (overlapsAcc || continuesAcc) {
        end = j; accTop = Math.min(accTop, bTop); accBottom = Math.max(accBottom, bBottom); j++;
        continue;
      }
      break;
    }
    // Extend backward: absorb an immediately-preceding label block that isn't
    // itself row-shaped (e.g. a header's leading cell wrapped to two lines)
    // but overlaps the first accepted row's range.
    let start = i;
    while (start > 0) {
      const p = blocks[start - 1];
      const pTop = p.bbox.y, pBottom = p.bbox.y + p.bbox.h;
      const firstTop = blocks[start].bbox.y, firstBottom = blocks[start].bbox.y + blocks[start].bbox.h;
      if (pTop < firstBottom && firstTop < pBottom) start--;
      else break;
    }
    if (rowCount >= MIN_ROWS) regions.push({ start, end });
    i = end + 1;
  }
  return regions;
}

/**
 * Merge a region's blocks into rows, line by line. A row boundary is detected
 * by RE-ENCOUNTERING the leftmost column after at least one other column has
 * been seen — this is more robust than a y-gap threshold, because the gap
 * between rows and the gap between wrapped lines within one row can be the
 * same size (both are just "one line height") depending on the document.
 */
function assembleRows(blocks, start, end) {
  const leftColKey = Math.round(blocks[start].bbox.x / X_BUCKET) * X_BUCKET;
  const rows = [];
  let current = new Map(); // xBucket -> lines[]
  let sawOtherCol = false;
  const flush = () => { if (current.size) rows.push(current); current = new Map(); sawOtherCol = false; };

  for (let idx = start; idx <= end; idx++) {
    for (const ln of blocks[idx].lines || []) {
      if (!ln.bbox || !(ln.text || "").trim()) continue;
      const key = Math.round(ln.bbox.x / X_BUCKET) * X_BUCKET;
      if (key === leftColKey) {
        if (current.has(leftColKey) && sawOtherCol) flush();
      } else if (current.has(leftColKey)) {
        sawOtherCol = true;
      }
      if (!current.has(key)) current.set(key, []);
      current.get(key).push(ln);
    }
  }
  flush();

  return rows.map((cellMap) => {
    const cells = [];
    for (const [x, grp] of cellMap) {
      grp.sort((a, b) => a.bbox.y - b.bbox.y);
      const text = grp.map((ln) => (ln.text || "").trim()).filter(Boolean).join(" ");
      if (!text) continue;
      cells.push({
        x,
        xEnd: Math.max(...grp.map((ln) => ln.bbox.x + ln.bbox.w)),
        text,
        bold: grp.every((ln) => isBold(ln.font)),
        italic: grp.some((ln) => isItalic(ln.font)),
      });
    }
    return cells.sort((a, b) => a.x - b.x);
  });
}

/** Cluster cell x-starts across all rows into column bands via gap detection. */
function columnBands(rows) {
  const xs = [...new Set(rows.flatMap((r) => r.map((c) => c.x)))];
  return clusterByGap(xs, BAND_GAP);
}

const bandIndexFor = nearestBandIndex;

/**
 * Reject regions that merely LOOK aligned (e.g. a résumé's "role ... dates"
 * line, right-tab-stopped rather than tabular). Real tables — verified against
 * sample financial and academic tables — have each row's populated-column set
 * either a superset or subset of every other row's (columns drop off from the
 * right as data gets sparser). A résumé's per-row second column jumps between
 * unrelated positions row to row, which breaks this "chain" property, because
 * it isn't a table at all.
 */
function hasNestedColumnUsage(grid) {
  const sigs = grid.map((row) => new Set(row.flatMap((c, i) => (c ? [i] : []))));
  for (let i = 0; i < sigs.length; i++) {
    for (let j = i + 1; j < sigs.length; j++) {
      const a = sigs[i], b = sigs[j];
      const aSubB = [...a].every((x) => b.has(x));
      const bSubA = [...b].every((x) => a.has(x));
      if (!aSubB && !bSubA) return false;
    }
  }
  return true;
}

/**
 * Build a { rows, numericCols } grid from a detected table region, or `null`
 * if the region doesn't hold up as a coherent table on closer inspection.
 * `rows` is an array of arrays of `{text, bold, italic, columnSpan} | null`
 * (null = a cell covered by a preceding columnSpan — omit it when rendering).
 */
export function buildTableGrid(blocks, region) {
  const rawRows = assembleRows(blocks, region.start, region.end);
  const bands = columnBands(rawRows);
  if (bands.length < 2) return null;
  const grid = rawRows.map((cells) => {
    const row = new Array(bands.length).fill(null);
    for (const cell of cells) {
      const col = bandIndexFor(cell.x, bands);
      // Column-span: the cell's text extends past the next band(s) with nothing
      // else claiming them in this row (geometric inference — no border signal
      // is available; row-span is deliberately NOT inferred for the same reason,
      // an unrecoverable ambiguity here is safer left as a visibly empty cell
      // than guessed wrong, and the free verifier flags empty cells for review).
      let span = 1;
      while (col + span < bands.length && cell.xEnd > bands[col + span].x0 + 4 &&
             !cells.some((c) => c !== cell && bandIndexFor(c.x, bands) === col + span)) {
        span++;
      }
      row[col] = { text: cell.text, bold: cell.bold, italic: cell.italic, columnSpan: span };
      for (let k = 1; k < span; k++) row[col + k] = undefined; // covered — omit entirely
    }
    return row;
  });

  if (!hasNestedColumnUsage(grid)) return null;

  const numericCols = new Set();
  for (let c = 0; c < bands.length; c++) {
    const dataCells = grid.slice(1).map((r) => r[c]).filter((cell) => cell && cell.text);
    if (dataCells.length && dataCells.every((cell) => NUMERIC.test(cell.text.trim()))) {
      numericCols.add(c);
    }
  }
  // Relative column widths from measured band positions, so a narrow label
  // column and a wide description column (e.g. the heat-transfer table) don't
  // get force-split into equal thirds.
  const relativeWidths = bands.map((b, i) =>
    i < bands.length - 1 ? bands[i + 1].x0 - b.x0 : Math.max(b.x1 - b.x0 + 60, 60)
  );
  return { rows: grid, numericCols, relativeWidths };
}
