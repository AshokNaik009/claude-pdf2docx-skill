// Deterministic per-page reconstruction from the PDF text layer (the "text
// path" — no LLM). Builds the block model directly from mupdf's structured
// output, and extracts embedded image regions as PNG assets (preserved as-is,
// referenced by both the text path and the LLM path).
import { writeFileSync } from "node:fs";
import { join } from "node:path";
import { findTableRegions, buildTableGrid } from "./tables.js";
import { buildLayoutSegments, mergeParagraphs } from "./layout.js";

/** Most common line font size on the page = body text size. */
function bodySize(blocks) {
  const counts = new Map();
  for (const b of blocks)
    for (const ln of b.lines || [])
      if (ln.font) counts.set(ln.font.size, (counts.get(ln.font.size) || 0) + (ln.text || "").length);
  let best = 11, max = -1;
  for (const [size, weight] of counts) if (weight > max) { max = weight; best = size; }
  return best;
}

/** Distinct heading sizes above body, largest → level 1. */
function headingLevels(blocks, body, ratio) {
  const sizes = new Set();
  for (const b of blocks)
    for (const ln of b.lines || [])
      if (ln.font && ln.font.size >= body * ratio) sizes.add(Math.round(ln.font.size * 10) / 10);
  const sorted = [...sizes].sort((a, b) => b - a);
  const map = new Map();
  sorted.forEach((s, idx) => map.set(s, Math.min(4, idx + 1)));
  return map; // size -> heading level
}

const BULLET = /^\s*([•·▪◦‣]|[-*])\s+(.*)$/;
const NUMBERED = /^\s*(\d+[.)])\s+(.*)$/;

/**
 * Extract every image block on the page as a PNG asset.
 * @returns Map from block index → { type:"image", path, width, height }
 */
export function extractImageAssets(render, structured, pageDir, relDir) {
  const map = new Map();
  const blocks = structured.blocks || [];
  let n = 0;
  blocks.forEach((b, idx) => {
    if (b.type !== "image" || !b.bbox) return;
    n++;
    const { png, width, height } = render.crop(b.bbox);
    const file = `img_${n}.png`;
    writeFileSync(join(pageDir, file), png);
    map.set(idx, { type: "image", path: join(relDir, file), width, height });
  });
  return map;
}

const FIGURE_CAPTION = /^(fig\.?|figure)\s*\d+/i;
const MIN_FIGURE_HEIGHT = 60;   // pt; smaller than this is just normal paragraph spacing, not a figure
const MAX_FIGURE_HEIGHT = 500;  // pt; cap runaway crops if the gap-detection heuristic misfires
const PAGE_MARGIN = 72;         // pt; matches docx.js's DEFAULT_PROPS 1in margins

/**
 * A vector-drawn diagram (arrows, boxes, rules — not an embedded raster/photo)
 * produces NO "image" block in mupdf's structured text at all (verified
 * empirically: zero vector blocks across every real sample PDF this pipeline
 * was built against), so extractImageAssets never sees it and the page would
 * otherwise lose the figure entirely, keeping only its caption. This can't be
 * detected structurally — but a figure's CAPTION can (a block starting
 * "Fig. N" / "Figure N", the same pattern the LLM-generation skill's own
 * `caption()` helper produces), and the diagram is almost always exactly the
 * vertical gap between that caption and whatever content ends right above it.
 * Crop that gap from the already-rendered page raster — the same raw
 * mechanism extractImageAssets uses for real embedded images — and treat it
 * as one more image asset. A real embedded image found nearby suppresses
 * this (no double-cropping); a too-small gap (ordinary paragraph spacing, no
 * room for a diagram) skips it rather than guessing.
 * @returns {{bbox: object, asset: {type:"image", path, width, height}}[]}
 */
export function extractFigureRegions(structured, render, pageDir, relDir, existingImageAssets, pageWidth) {
  const blocks = structured.blocks || [];
  const textBlocks = blocks
    .filter((b) => b.type === "text" && b.bbox && (b.lines || []).some((ln) => (ln.text || "").trim()))
    .sort((a, b) => (Math.abs(a.bbox.y - b.bbox.y) > 4 ? a.bbox.y - b.bbox.y : a.bbox.x - b.bbox.x));

  const imageBboxes = [...existingImageAssets.keys()].map((idx) => blocks[idx]?.bbox).filter(Boolean);
  const overlapsExistingImage = (yTop, yBottom) =>
    imageBboxes.some((b) => yTop < b.y + b.h && b.y < yBottom);

  const out = [];
  let figN = 0;
  for (let i = 0; i < textBlocks.length; i++) {
    const b = textBlocks[i];
    const text = (b.lines || []).map((ln) => ln.text || "").join(" ").trim();
    if (!FIGURE_CAPTION.test(text)) continue;

    const prev = textBlocks[i - 1];
    const yTop = Math.max(prev ? prev.bbox.y + prev.bbox.h : 0, b.bbox.y - MAX_FIGURE_HEIGHT);
    const yBottom = b.bbox.y;
    if (yBottom - yTop < MIN_FIGURE_HEIGHT || overlapsExistingImage(yTop, yBottom)) continue;

    // Span the full content width (page width minus margins), not just the
    // caption's own width — a diagram is usually wider than its caption text.
    const bbox = { x: PAGE_MARGIN, y: yTop, w: pageWidth - 2 * PAGE_MARGIN, h: yBottom - yTop };

    figN++;
    const { png, width, height } = render.crop(bbox);
    const file = `fig_${figN}.png`;
    writeFileSync(join(pageDir, file), png);
    out.push({ bbox, asset: { type: "image", path: join(relDir, file), width, height } });
  }
  return out;
}

/**
 * Classify an ALREADY reading-order-sorted array of `{bbox, lines}` text
 * blocks (no images — those are handled as their own segments by the caller)
 * into heading/paragraph/list doc-blocks. `body`/`levels` are page-wide font
 * stats (see bodySize/headingLevels) so heading detection stays consistent
 * across column bands rather than being skewed by one band's smaller sample.
 */
export function classifyMergedBlocks(blocks, body, levels) {
  const out = [];
  let pendingList = null;
  const flushList = () => { if (pendingList) { out.push(pendingList); pendingList = null; } };

  for (const b of blocks) {
    const lines = (b.lines || []).map((ln) => (ln.text || "").replace(/\s+$/g, ""));
    const text = lines.join(" ").replace(/\s+/g, " ").trim();
    if (!text) continue;

    const maxSize = Math.max(...(b.lines || []).map((ln) => ln.font?.size || body));
    const rounded = Math.round(maxSize * 10) / 10;

    // Heading: whole block is large and short (a line or two).
    if (levels.has(rounded) && lines.length <= 2 && text.length <= 120) {
      flushList();
      out.push({ type: "heading", level: levels.get(rounded), text });
      continue;
    }

    // List detection per line.
    const bulletLines = lines.filter((l) => BULLET.test(l) || NUMBERED.test(l));
    if (bulletLines.length && bulletLines.length >= lines.length - 1) {
      const ordered = NUMBERED.test(lines[0]);
      const items = lines.map((l) => {
        const m = l.match(BULLET) || l.match(NUMBERED);
        return m ? m[2].trim() : l.trim();
      }).filter(Boolean);
      if (!pendingList || pendingList.ordered !== ordered) { flushList(); pendingList = { type: "list", ordered, items: [] }; }
      pendingList.items.push(...items);
      continue;
    }

    flushList();
    out.push({ type: "paragraph", text });
  }
  flushList();
  return out;
}

const readingOrderSort = (a, b) => {
  const ay = a.bbox?.y ?? 0, by = b.bbox?.y ?? 0;
  if (Math.abs(ay - by) > 4) return ay - by;
  return (a.bbox?.x ?? 0) - (b.bbox?.x ?? 0);
};

const blockTop = (b) => b.bbox?.y ?? 0;

/**
 * Top-level deterministic reconstruction for one page: detects table regions
 * (src/tables.js) and removes them from consideration BEFORE column-layout
 * detection ever sees them (a table's own cell x-positions look exactly like
 * page columns otherwise — see src/layout.js's isRowShaped filter for the
 * matching defense-in-depth check). Whatever text remains is segmented into
 * column bands / full-width runs (src/layout.js); images become their own
 * full-width segments, matching the skill's "figures get their own
 * CONTINUOUS section" guidance. All segments are then merged into one
 * Y-ordered list and returned as a JSON-safe `{ sections }` descriptor — the
 * SAME shape docx.js's deterministic renderer and the LLM path both produce,
 * so pipeline.js/docx.js don't need to know which path built a given page.
 *
 * @returns {{ sections: Array<{properties: object, blocks: object[]}> }}
 */
export function buildPageSections(structured, imageAssets, cfg) {
  const allBlocks = structured.blocks || [];
  const textBlocks = allBlocks
    .map((b, idx) => ({ b, idx }))
    .filter(({ b }) => b.type === "text")
    .sort((p, q) => readingOrderSort(p.b, q.b));

  const textOnly = textBlocks.map(({ b }) => b);
  const tableRegions = findTableRegions(textOnly);
  const consumed = new Set();
  const tableSegments = [];
  for (const region of tableRegions) {
    const grid = buildTableGrid(textOnly, region);
    if (!grid) continue; // coherence check rejected it — leave those blocks for normal text handling
    for (let i = region.start; i <= region.end; i++) consumed.add(i);
    tableSegments.push({
      y: blockTop(textOnly[region.start]),
      properties: { type: "CONTINUOUS" },
      blocks: [{ type: "table", rows: grid.rows, numericCols: [...grid.numericCols], relativeWidths: grid.relativeWidths }],
    });
  }

  // A block immediately following a removed table region is NOT vertically
  // continuous with whatever residual text preceded the table — without this
  // marker, layout segmentation (which only sees the residual blocks, with no
  // knowledge a table was carved out between them) would treat "before the
  // table" and "after the table" as one uninterrupted column run and merge
  // them, corrupting reading order (a block from well below the table could
  // land ahead of correctly-ordered content via the "whole band0 then whole
  // band1" convention). See src/layout.js's run-building loop for the other half.
  const hasBlockText = (b) => (b.lines || []).some((ln) => (ln.text || "").trim());
  const residualText = [];
  let afterConsumedRun = false;
  textOnly.forEach((b, i) => {
    if (consumed.has(i)) { afterConsumedRun = true; return; }
    // Skip whitespace-only blocks without clearing the pending flag — layout.js
    // filters these out too (hasText), so marking one of them would lose the
    // break entirely instead of carrying it to the next block that survives.
    if (!hasBlockText(b)) { residualText.push(b); return; }
    residualText.push(afterConsumedRun ? { ...b, _breakBefore: true } : b);
    afterConsumedRun = false;
  });

  const body = bodySize(allBlocks);
  const levels = headingLevels(allBlocks, body, cfg.headingRatio);

  const layoutSegments = buildLayoutSegments(residualText, cfg);
  const textSegments = layoutSegments
    ? layoutSegments.map((seg) =>
        seg.type === "columns"
          ? {
              y: Math.min(...seg.bands.flatMap((band) => band.map(blockTop)), Infinity),
              properties: { type: "CONTINUOUS", column: { count: seg.count, space: 560 } },
              blocks: seg.bands.flatMap((band) => classifyMergedBlocks(band, body, levels)),
            }
          : {
              y: Math.min(...seg.blocks.map(blockTop), Infinity),
              properties: { type: "CONTINUOUS" },
              blocks: classifyMergedBlocks(seg.blocks, body, levels),
            }
      )
    : residualText.length
    ? [{
        y: blockTop(residualText[0]),
        properties: { type: "CONTINUOUS" },
        blocks: classifyMergedBlocks(mergeParagraphs(residualText), body, levels),
      }]
    : [];

  const imageSegments = allBlocks
    .map((b, idx) => ({ b, idx }))
    .filter(({ b, idx }) => b.type === "image" && b.bbox && imageAssets.has(idx))
    .map(({ b, idx }) => ({
      y: blockTop(b),
      properties: { type: "CONTINUOUS" },
      blocks: [imageAssets.get(idx)],
    }));

  const sections = [...tableSegments, ...textSegments, ...imageSegments]
    .sort((a, b) => a.y - b.y)
    .map(({ y, ...s }) => s);

  if (!sections.length) return { sections: [{ properties: {}, blocks: [] }] };
  return { sections };
}
