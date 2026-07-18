// Deterministic per-page reconstruction from the PDF text layer (the "text
// path" — no LLM). Builds the block model directly from mupdf's structured
// output, and extracts embedded image regions as PNG assets (preserved as-is,
// referenced by both the text path and the LLM path).
import { writeFileSync } from "node:fs";
import { join } from "node:path";

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

/**
 * Build the block model for a clean-text page. Image blocks are interleaved in
 * reading order using the pre-extracted asset map.
 */
export function buildTextBlocks(structured, imageAssets, cfg) {
  const raw = (structured.blocks || []).map((b, idx) => ({ b, idx }));
  // Reading order: top-to-bottom, then left-to-right.
  raw.sort((p, q) => {
    const ay = p.b.bbox?.y ?? 0, by = q.b.bbox?.y ?? 0;
    if (Math.abs(ay - by) > 4) return ay - by;
    return (p.b.bbox?.x ?? 0) - (q.b.bbox?.x ?? 0);
  });

  const body = bodySize(structured.blocks || []);
  const levels = headingLevels(structured.blocks || [], body, cfg.headingRatio);
  const out = [];
  let pendingList = null;
  const flushList = () => { if (pendingList) { out.push(pendingList); pendingList = null; } };

  for (const { b, idx } of raw) {
    if (b.type === "image") {
      flushList();
      if (imageAssets.has(idx)) out.push(imageAssets.get(idx));
      continue;
    }
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
