// Docling doc.json  ->  per-page docx.js `buildPage` modules (+ manifests +
// session.json) that the NORMAL merge pipeline consumes (src/docx.js
// executeModule / renderDocx). No LLM, no pandoc.
//
// Layout reconstruction: each element carries a prov bbox. On a page we split
// elements by their horizontal center into a left/right column and emit the
// two-column band as a BORDERLESS 2-cell table (deterministic, robust to
// unequal column lengths — the same trick the --pandoc output used). Full-width
// elements (tables, figures, spanning headers) break the band and render across
// the page. The output is real docx.js code files you can open, edit, and merge.
import { readFileSync, writeFileSync, mkdirSync, copyFileSync, existsSync } from "node:fs";
import { join, basename } from "node:path";

const FULL_FRAC = 0.62;  // bbox wider than this fraction of the page ⇒ full-width
const BAND_FRAC = 0.12;  // top/bottom 12% of the page ⇒ header/footer margin band
const DROP_LABELS = new Set(["page_header", "page_footer"]);
const pad = (n) => String(n).padStart(4, "0");

// Normalize text for cross-page comparison: lowercase, digits→#, ws collapsed.
// "536" and "537" both become "###" so a running page-number footer groups.
const norm = (s) => (s || "").toLowerCase().replace(/\d+/g, "#").replace(/\s+/g, " ").trim();

// Footer strings that carry a live page number: "5", "5 of 20", "Page 5 of 20".
const isPageNumberText = (s) =>
  /^\d+$/.test(s) || /^\d+\s+of\s+\d+$/i.test(s) || /^page\s+\d+(\s+of\s+\d+)?$/i.test(s);

const modeOf = (arr) => {
  const m = new Map(); let best = arr[0], bc = 0;
  for (const x of arr) { const c = (m.get(x) || 0) + 1; m.set(x, c); if (c > bc) { bc = c; best = x; } }
  return best;
};

/** Flatten body.children into leaf elements in reading order (recurse groups). */
function flatten(doc, children, out = []) {
  for (const ref of children || []) {
    const [, kind, idxStr] = ref.$ref.split("/");
    const node = doc[kind]?.[Number(idxStr)];
    if (!node) continue;
    if (kind === "groups") { flatten(doc, node.children, out); continue; }
    out.push({ kind, idx: Number(idxStr), node });
  }
  return out;
}

const bboxOf = (node) => node.prov?.[0]?.bbox || null;
const pageOf = (node) => node.prov?.[0]?.page_no ?? null;
const pageWidth = (doc, pageNo) =>
  doc.pages?.[String(pageNo)]?.size?.width || doc.pages?.[pageNo]?.size?.width || 684;
const pageHeight = (doc, pageNo) =>
  doc.pages?.[String(pageNo)]?.size?.height || doc.pages?.[pageNo]?.size?.height || 864;

// Docling bbox origin is BOTTOMLEFT, so a high `t`/`b` means near the page top.
const isFurniture = (node) => node.content_layer === "furniture" || DROP_LABELS.has(node.label);
function bandOf(doc, node) {
  const b = bboxOf(node), p = pageOf(node);
  if (!b || p == null) return null;
  const frac = ((b.t + b.b) / 2) / pageHeight(doc, p);
  if (frac > 1 - BAND_FRAC) return "header";
  if (frac < BAND_FRAC) return "footer";
  return null;
}

/**
 * Detect running headers/footers + a live page number, layout-agnostically.
 * Layer 1 (authoritative): anything Docling tags furniture / page_header /
 * page_footer is routed OUT of the body. Layer 2 (research-standard safety net
 * for classifier misses, incl. docling bug #3015): body text sitting in the
 * top/bottom margin band whose NORMALIZED text repeats on ≥2 pages is a running
 * element too. A group is only shown as a running H/F line when it repeats on
 * ≥2 pages; single-occurrence furniture is dropped from the body but not shown.
 * Returns { headerLines, footer, pageNumberStart, removeSet } — removeSet holds
 * "kind:idx" keys to exclude from the body flow.
 */
function detectHeadersFooters(doc, bodyLeaves, furnitureLeaves, pnumOf) {
  const cands = [];
  const add = (l, layer) => {
    const raw = (l.node.text ?? l.node.orig ?? "").trim();
    if (!raw) return;
    let band = bandOf(doc, l.node);
    if (!band) { // furniture without a usable bbox: fall back to its label
      if (l.node.label === "page_header") band = "header";
      else if (l.node.label === "page_footer") band = "footer";
      else return;
    }
    const b = bboxOf(l.node);
    cands.push({ key: `${l.kind}:${l.idx}`, layer, band, page: pageOf(l.node),
      cy: b ? (b.t + b.b) / 2 : 0, raw, norm: norm(raw) });
  };
  // Layer 1 — furniture from the body tree AND the furniture root.
  for (const l of [...bodyLeaves, ...furnitureLeaves])
    if (l.kind === "texts" && isFurniture(l.node)) add(l, 1);
  // Layer 2 — non-furniture body text sitting in a margin band.
  for (const l of bodyLeaves)
    if (l.kind === "texts" && !isFurniture(l.node) && bandOf(doc, l.node)) add(l, 2);

  const groups = new Map();
  for (const c of cands) {
    const gk = `${c.band}|${c.norm}`;
    if (!groups.has(gk)) groups.set(gk, []);
    groups.get(gk).push(c);
  }
  const removeSet = new Set();
  const headerGroups = [], footerGroups = [];
  for (const arr of groups.values()) {
    for (const c of arr) if (c.layer === 1) removeSet.add(c.key); // always de-inline furniture
    if (new Set(arr.map((c) => c.page)).size < 2) continue;       // running ⇒ ≥2 pages
    for (const c of arr) removeSet.add(c.key);
    const rep = arr.slice().sort((a, b) => a.page - b.page)[0];
    (rep.band === "header" ? headerGroups : footerGroups).push({ rep, arr });
  }
  headerGroups.sort((a, b) => b.rep.cy - a.rep.cy); // top of page first
  footerGroups.sort((a, b) => b.rep.cy - a.rep.cy);

  const headerLines = headerGroups.map((g) => g.rep.raw);

  let footer = null, pageNumberStart;
  const numGroup = footerGroups.find((g) => isPageNumberText(g.rep.raw));
  if (numGroup) {
    footer = { kind: "pagenum", template: numGroup.rep.raw };
    const starts = numGroup.arr
      .map((c) => { const m = c.raw.match(/\d+/); return m ? Number(m[0]) - (pnumOf(c.page) - 1) : null; })
      .filter((x) => x != null);
    if (starts.length) pageNumberStart = modeOf(starts);
  } else if (footerGroups.length) {
    footer = { kind: "literal", template: footerGroups[0].rep.raw };
  }
  return { headerLines, footer, pageNumberStart, removeSet };
}

/** JS expr for a Footer paragraph's `children`, page numbers → live fields. */
function footerChildrenCode(template) {
  let intSeen = 0;
  return "[ " + template.split(/(\d+)/).filter((s) => s !== "").map((p) => {
    if (/^\d+$/.test(p) && ++intSeen <= 2)
      return `new TextRun({ children: [docx.PageNumber.${intSeen === 1 ? "CURRENT" : "TOTAL_PAGES"}] })`;
    return `new TextRun(${S(p)})`;
  }).join(", ") + " ]";
}

/** Statements + return-object fields that attach a running header/footer. */
function hfEmit(hf) {
  const stmts = [], fields = [];
  if (hf.headerLines?.length) {
    const paras = hf.headerLines.map((t) =>
      `new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [ new TextRun({ text: ${S(t)}, italics: true, size: 18 }) ] })`);
    stmts.push(`  const __header = new docx.Header({ children: [ ${paras.join(", ")} ] });`);
    fields.push("header: __header");
  }
  if (hf.footer) {
    stmts.push(`  const __footer = new docx.Footer({ children: [ new Paragraph({ alignment: AlignmentType.CENTER, children: ${footerChildrenCode(hf.footer.template)} }) ] });`);
    fields.push("footer: __footer");
  }
  if (hf.pageNumberStart != null) fields.push(`pageNumberStart: ${Number(hf.pageNumberStart)}`);
  return { stmts, fields };
}

/** Build the { grid } data (JSON-serializable) for a Docling table. */
function tableGrid(node) {
  const grid = node.data?.grid;
  if (!Array.isArray(grid)) return null;
  return grid.map((row, r) =>
    row.map((cell, c) => {
      if (!cell) return null;
      const originR = cell.start_row_offset_idx ?? r;
      const originC = cell.start_col_offset_idx ?? c;
      if (originR !== r || originC !== c) return null; // covered by a span
      return {
        t: (cell.text || "").trim(),
        h: !!cell.column_header,   // top header row
        rh: !!cell.row_header,     // left column acting as a header (row label)
        sec: !!cell.row_section,   // a section-divider row inside the table
        cs: cell.col_span > 1 ? cell.col_span : 1,
        rs: cell.row_span > 1 ? cell.row_span : 1,
      };
    })
  );
}

// ---- code emission ---------------------------------------------------------

const S = (v) => JSON.stringify(v); // safe JS string / JSON literal

/** A column-cell child expression, e.g. body("..") / head("..",2) / cap(".."). */
function kidExpr(node) {
  const text = (node.text ?? node.orig ?? "").trim();
  switch (node.label) {
    case "title": return `head(${S(text)}, 1)`;
    case "section_header": return `head(${S(text)}, 2)`;
    case "caption": return `cap(${S(text)})`;
    default: return `body(${S(text)})`;
  }
}

/** Statements that push a full-width element onto `children`. */
function fullStatements(el, figFile) {
  const { kind, node } = el;
  if (kind === "tables") {
    const grid = tableGrid(node);
    return grid ? [`  children.push(gridTable(${S(grid)}));`] : [];
  }
  if (kind === "pictures") {
    return figFile
      ? [`  children.push(fig(${S(figFile)}));`]
      : [`  children.push(body("[figure not extracted]"));`];
  }
  // full-width text/header/caption
  return [`  children.push(${kidExpr(node)});`];
}

const PRELUDE = `export function buildPage(docx, ctx) {
  const { Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle,
          AlignmentType, HeadingLevel, VerticalAlign, TableLayoutType } = docx;
  const DXA = (n) => ({ size: n, type: WidthType.DXA });
  const NONE = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  const NB = { top: NONE, bottom: NONE, left: NONE, right: NONE };
  const TB = { top: NONE, bottom: NONE, left: NONE, right: NONE, insideHorizontal: NONE, insideVertical: NONE };
  const GRID = { style: BorderStyle.SINGLE, size: 4, color: "A6A6A6" };
  const runs = (t) => t.split(/(et al\\.)/g).filter((s) => s !== "")
    .map((s) => s === "et al." ? new TextRun({ text: s, italics: true }) : new TextRun(s));
  const body = (t) => new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 140, line: 264 }, children: runs(t) });
  const head = (t, l = 2) => new Paragraph({
    heading: l === 1 ? HeadingLevel.HEADING_1 : l === 2 ? HeadingLevel.HEADING_2 : HeadingLevel.HEADING_3,
    spacing: { before: 200, after: 100 }, children: [new TextRun({ text: t, bold: true })] });
  const cap = (t) => new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 100, after: 140 }, children: [new TextRun({ text: t, bold: true })] });
  const twoCol = (L, R) => new Table({
    width: DXA(9360), columnWidths: [4680, 4680], layout: TableLayoutType.FIXED, borders: TB,
    rows: [ new TableRow({ children: [
      new TableCell({ width: DXA(4680), borders: NB, margins: { right: 200 }, children: L.length ? L : [new Paragraph({})] }),
      new TableCell({ width: DXA(4680), borders: NB, margins: { left: 200 }, children: R.length ? R : [new Paragraph({})] }),
    ] }) ] });
  // Cell emphasis is driven by Docling's semantic flags (column_header /
  // row_header / row_section) — the ONLY structure the parser exposes; it drops
  // real fill/text/border colors, so we apply light, readable defaults (dark
  // text stays legible) rather than guessing brand colors.
  const HFILL = "D9D9D9", SFILL = "E4E4E4", RFILL = "F2F2F2";
  const gridTable = (grid, tw = 9360) => { const ncols = Math.max(1, ...grid.map((r) => r.reduce((s, c) => s + (c ? (c.cs || 1) : 0), 0)));
    const cw = Math.floor(tw / ncols); return new Table({
    width: DXA(tw), columnWidths: Array(ncols).fill(cw), layout: TableLayoutType.FIXED,
    borders: { top: GRID, bottom: GRID, left: GRID, right: GRID, insideHorizontal: GRID, insideVertical: GRID },
    rows: grid.map((row) => new TableRow({
      tableHeader: row.some((c) => c && c.h),
      children: row.filter(Boolean).map((c) => {
        const fill = c.h ? HFILL : c.sec ? SFILL : c.rh ? RFILL : undefined;
        return new TableCell({
        width: DXA(cw * (c.cs || 1)),
        columnSpan: c.cs > 1 ? c.cs : undefined, rowSpan: c.rs > 1 ? c.rs : undefined,
        verticalAlign: VerticalAlign.TOP, margins: { top: 40, bottom: 40, left: 100, right: 100 },
        shading: fill ? { fill } : undefined,
        children: [ new Paragraph({ children: [ new TextRun({ text: c.t || " ", bold: !!(c.h || c.rh || c.sec) }) ] }) ],
      }); }),
    })) }); };
  const fig = (file) => { try { return new Paragraph({ alignment: AlignmentType.CENTER, keepNext: true, spacing: { before: 100 }, children: [ ctx.image(file) ] }); }
    catch (e) { return new Paragraph({ children: [ new TextRun({ text: "[image: " + file + "]", italics: true }) ] }); } };
  const children = [];
`;

const COL_DXA = 4200; // usable width of a two-column cell for a nested table

/** A child expression for an element sitting INSIDE a column cell — a text
 * paragraph, a column-width nested table, or a figure. */
function colChildExpr(el, figFileFor) {
  const { kind, node } = el;
  if (kind === "tables") { const g = tableGrid(node); return g ? `gridTable(${S(g)}, ${COL_DXA})` : `body("")`; }
  if (kind === "pictures") { const f = figFileFor(el.idx); return f ? `fig(${S(f)})` : `body("[figure not extracted]")`; }
  return kidExpr(node);
}

/** Generate the page.mjs source for one page's leaves. `hf` (doc-level running
 * header/footer, only passed for the first page) is attached to the returned
 * section object — renderDocx picks it up from the first page that provides it,
 * exactly as it already does for `styles`.
 *
 * Layout reconstruction is a geometry-driven X-Y cut, NOT Docling's reading
 * order (which scrambles on 2-D pages like the credit-memo quadrant layout):
 * elements are ordered top→bottom by bbox; an element that spans the central
 * axis (a title, a full-width table/figure or its caption) is emitted
 * full-width and breaks the surrounding two-column band; everything else is
 * assigned to the left/right column by its center-x and stacked in place — so a
 * column-width table stays nested inside its own column. */
function emitPageModule(leaves, doc, figFileFor, hf = null) {
  const pw = pageWidth(doc, pageOf(leaves.find((l) => pageOf(l.node))?.node) ?? 1);
  const mid = pw / 2, EPS = 0.05 * pw;
  const cx = (n) => { const b = bboxOf(n); return b ? (b.l + b.r) / 2 : mid; };
  const cy = (n) => { const b = bboxOf(n); return b ? (b.t + b.b) / 2 : 0; };
  // Spanning = wider than most of the page OR straddling the central axis with
  // real extent on both sides (a centered title, a caption over a wide object).
  const isSpanning = (n) => {
    const b = bboxOf(n);
    if (!b) return true;
    return (b.r - b.l) > FULL_FRAC * pw || (b.l < mid - EPS && b.r > mid + EPS);
  };
  const side = (n) => (cx(n) < mid ? "L" : "R");

  const items = leaves.filter((el) => {
    const { kind, node } = el;
    if (kind === "texts" && DROP_LABELS.has(node.label)) return false;
    const text = (node.text ?? node.orig ?? "").trim();
    if (kind === "texts" && (!text || (text.length <= 2 && (bboxOf(node)?.r - bboxOf(node)?.l || 0) < 10))) return false; // stray "."
    return true;
  });
  // Top→bottom (BOTTOMLEFT origin ⇒ larger cy is higher); ties left→right.
  items.sort((a, b) => cy(b.node) - cy(a.node) || cx(a.node) - cx(b.node));

  const lines = [PRELUDE];
  let band = []; // pending column elements between two full-width breaks

  const flushBand = () => {
    if (!band.length) return;
    const L = band.filter((el) => side(el.node) === "L").map((el) => colChildExpr(el, figFileFor));
    const R = band.filter((el) => side(el.node) === "R").map((el) => colChildExpr(el, figFileFor));
    if (!L.length || !R.length) {
      for (const e of [...L, ...R]) lines.push(`  children.push(${e});`); // single-column region
    } else {
      lines.push(`  children.push(twoCol([${L.join(", ")}], [${R.join(", ")}]));`);
    }
    band = [];
  };

  for (const el of items) {
    if (isSpanning(el.node)) {
      flushBand();
      lines.push(...fullStatements(el, el.kind === "pictures" ? figFileFor(el.idx) : null));
    } else {
      band.push(el);
    }
  }
  flushBand();
  const { stmts, fields } = hf ? hfEmit(hf) : { stmts: [], fields: [] };
  lines.push(...stmts);
  lines.push(`  return { sections: [ { children } ]${fields.length ? ", " + fields.join(", ") : ""} };\n}\n`);
  return lines.join("\n");
}

/** Main: read a docling out dir, write a merge-ready work dir of docx.js code. */
export function generateDoclingCode(doclingOutDir, workDir, pdfName = "document") {
  const doc = JSON.parse(readFileSync(join(doclingOutDir, "doc.json"), "utf8"));
  const bodyLeaves = flatten(doc, doc.body?.children);
  const furnitureLeaves = flatten(doc, doc.furniture?.children);

  // Sequential page index (1-based) for a docling page_no — used for the live
  // page number's start value.
  const allPages = [...new Set(bodyLeaves.map((l) => pageOf(l.node)).filter((p) => p != null))].sort((a, b) => a - b);
  const pnumOf = (p) => allPages.indexOf(p) + 1;

  const hf = detectHeadersFooters(doc, bodyLeaves, furnitureLeaves, pnumOf);
  const leaves = bodyLeaves.filter((l) => !hf.removeSet.has(`${l.kind}:${l.idx}`));

  // Group leaves by page in reading order; keep page appearance order.
  const pagesOrder = [];
  const byPage = new Map();
  for (const l of leaves) {
    let p = pageOf(l.node);
    if (p == null) p = pagesOrder[pagesOrder.length - 1] ?? 1; // attach no-geometry leaf to current page
    if (!byPage.has(p)) { byPage.set(p, []); pagesOrder.push(p); }
    byPage.get(p).push(l);
  }

  mkdirSync(join(workDir, "pages"), { recursive: true });
  mkdirSync(join(workDir, "assets"), { recursive: true });

  // Map each picture index -> the page it belongs to, so we copy the right PNG.
  const picPage = new Map();
  doc.pictures?.forEach((p, i) => picPage.set(i, pageOf(p)));

  const sessionPages = [];
  pagesOrder.sort((a, b) => a - b).forEach((pageNo, order) => {
    const pnum = order + 1; // sequential page index for filenames
    const relAssets = `assets/page_${pad(pnum)}`;
    const assetsAbs = join(workDir, relAssets);
    mkdirSync(assetsAbs, { recursive: true });

    // Copy figures for this page and build idx->filename map.
    let figN = 0;
    const figMap = new Map();
    for (const [i, pg] of picPage) {
      if (pg !== pageNo) continue;
      figN++;
      const src = join(doclingOutDir, "images", `pic_${i + 1}.png`);
      const file = `fig_${figN}.png`;
      if (existsSync(src)) copyFileSync(src, join(assetsAbs, file));
      figMap.set(i, file);
    }

    const leavesForPage = byPage.get(pageNo);
    const code = emitPageModule(leavesForPage, doc, (idx) => figMap.get(idx) || null, order === 0 ? hf : null);
    const codeRel = `pages/page_${pad(pnum)}/page.mjs`;
    mkdirSync(join(workDir, "pages", `page_${pad(pnum)}`), { recursive: true });
    writeFileSync(join(workDir, codeRel), code);

    const fallbackText = leavesForPage
      .filter((l) => l.kind === "texts")
      .map((l) => (l.node.text ?? l.node.orig ?? "").trim())
      .filter(Boolean);

    const manifest = {
      page: pnum,
      mode: "llm",
      reason: "docling doc.json",
      codeFile: codeRel,
      fallbackText,
      assets: [...figMap.values()].map((f) => `${relAssets}/${f}`),
      assetsDir: relAssets,
    };
    writeFileSync(join(workDir, "pages", `page_${pad(pnum)}.json`), JSON.stringify(manifest, null, 2));

    sessionPages.push({
      page: pnum, status: "completed", mode: "llm", reason: "docling",
      manifest: `pages/page_${pad(pnum)}.json`, assets: relAssets, error: null,
    });
  });

  const session = {
    sessionId: `${basename(workDir)}-${Date.now().toString(36)}`,
    document: pdfName,
    totalPages: sessionPages.length,
    engine: "docling",
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    pages: sessionPages,
  };
  writeFileSync(join(workDir, "session.json"), JSON.stringify(session, null, 2));
  return { workDir, pages: sessionPages.length };
}
