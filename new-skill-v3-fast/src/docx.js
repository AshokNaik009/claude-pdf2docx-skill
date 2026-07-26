// DOCX assembly (MERGE → DOCX GENERATION). Executes each page's Claude-generated
// `buildPage(docx, ctx)` module, falls back to deterministic block rendering for
// text pages (and to plain paragraphs if a generated module throws), inserts
// page breaks, and packs the final .docx. Native Word elements only — no frames.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { pathToFileURL } from "node:url";
import * as docx from "docx";

const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, SectionType,
  Table, TableRow, TableCell, WidthType, BorderStyle, ImageRun, AlignmentType,
} = docx;

const CONTENT_DXA = 9360; // 6.5in content width
const CONTENT_PX = 624;

// US Letter, 1in margins, single column — used for text pages and fallbacks.
const DEFAULT_PROPS = {
  page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } },
};

const HEADING = {
  1: HeadingLevel.HEADING_1, 2: HeadingLevel.HEADING_2,
  3: HeadingLevel.HEADING_3, 4: HeadingLevel.HEADING_4,
};

/** Read width/height from a PNG's IHDR (bytes 16..24). */
function pngSize(buf) {
  return { width: buf.readUInt32BE(16), height: buf.readUInt32BE(20) };
}

/**
 * Import and run a generated page module; returns { sections, styles } where
 * `sections` is one or more Word section objects. Accepts a single section,
 * an array of sections, a bare children array, or { sections, styles }.
 */
export async function executeModule(codeAbsPath, assetsAbsDir) {
  const url = pathToFileURL(codeAbsPath).href + "?v=" + Date.now(); // bust import cache on retries
  const mod = await import(url);
  if (typeof mod.buildPage !== "function") throw new Error("module has no buildPage export");
  const result = mod.buildPage(docx, makeCtx(assetsAbsDir));

  let sections, styles;
  // Doc-level running header/footer + live page-number start a page module may
  // carry (like `styles`); renderDocx applies them from the first page that does.
  let header, footer, pageNumberStart;
  if (Array.isArray(result)) {
    sections = result.every((s) => Array.isArray(s?.children)) ? result : [{ children: result }];
  } else if (result && Array.isArray(result.sections)) {
    ({ sections, styles, header, footer, pageNumberStart } = result);
  } else if (result && Array.isArray(result.children)) {
    sections = [{ properties: result.properties, children: result.children }];
    ({ styles, header, footer, pageNumberStart } = result);
  } else {
    throw new Error("buildPage must return a section, an array of sections, or { sections, styles }");
  }
  if (!sections.length || !sections.every((s) => Array.isArray(s.children))) {
    throw new Error("every section must have a children array");
  }
  // Only the first section needs a default page setup; later ones inherit it.
  sections = sections.map((s, i) => ({
    properties: s.properties ?? (i === 0 ? DEFAULT_PROPS : {}),
    children: s.children,
  }));
  return { sections, styles, header, footer, pageNumberStart };
}

/** ctx passed to generated modules — currently exposes image embedding. */
export function makeCtx(assetsAbsDir) {
  return {
    image(file, opts = {}) {
      const data = readFileSync(join(assetsAbsDir, file));
      let { width, height } = pngSize(data);
      const maxW = opts.maxWidth || CONTENT_PX;
      if (width > maxW) { height = Math.round((height * maxW) / width); width = maxW; }
      return new ImageRun({ type: "png", data, transformation: { width, height } });
    },
  };
}

// ---- Deterministic rendering for text pages / fallbacks ---------------------

/**
 * Build a real docx.Table from a src/tables.js grid: `{ rows, numericCols,
 * relativeWidths }`. `rows[r][c]` is `{text,bold,italic,columnSpan}` | null
 * (genuinely empty cell) | undefined (covered by a preceding columnSpan —
 * omitted entirely, per docx's covered-cell convention).
 */
function tableElement(block) {
  const { rows, relativeWidths } = block;
  const numericCols = new Set(block.numericCols || []); // JSON manifest stores an array
  if (!rows.length) return null;
  const cols = relativeWidths.length;
  const totalRel = relativeWidths.reduce((s, w) => s + w, 0);
  const columnWidths = relativeWidths.map((w) => Math.round((w / totalRel) * CONTENT_DXA));
  const line = { style: BorderStyle.SINGLE, size: 4, color: "A6A6A6" };
  const margins = { top: 80, bottom: 80, left: 120, right: 120 };
  return new Table({
    width: { size: CONTENT_DXA, type: WidthType.DXA },
    columnWidths,
    rows: rows.map((row, ri) =>
      new TableRow({
        tableHeader: ri === 0,
        children: row
          .map((cell, ci) => {
            if (cell === undefined) return null; // covered by an earlier columnSpan
            const isHeader = ri === 0;
            const numeric = numericCols.has(ci);
            const text = cell?.text ?? "";
            return new TableCell({
              width: { size: columnWidths[ci], type: WidthType.DXA },
              margins, borders: { top: line, bottom: line, left: line, right: line },
              verticalAlign: "center",
              ...(cell?.columnSpan > 1 ? { columnSpan: cell.columnSpan } : {}),
              children: [
                new Paragraph({
                  alignment: numeric ? AlignmentType.RIGHT : undefined,
                  children: [new TextRun({ text, bold: isHeader || !!cell?.bold, italics: !!cell?.italic })],
                }),
              ],
            });
          })
          .filter(Boolean),
      })
    ),
    borders: { top: line, bottom: line, left: line, right: line, insideHorizontal: line, insideVertical: line },
  });
}

export function blockToChildren(block, ctx) {
  switch (block.type) {
    case "heading":
      return [new Paragraph({ heading: HEADING[Math.min(4, block.level || 1)], children: [new TextRun(block.text)], spacing: { before: 240, after: 120 } })];
    case "paragraph":
      return [new Paragraph({ children: [new TextRun(block.text)], spacing: { after: 160, line: 276 } })];
    case "list":
      return block.items.map((it, idx) => new Paragraph({ children: [new TextRun((block.ordered ? `${idx + 1}. ` : "• ") + it)], indent: { left: 360 }, spacing: { after: 80 } }));
    case "table": {
      const t = tableElement(block);
      return t ? [t, new Paragraph({ children: [], spacing: { after: 160 } })] : [];
    }
    case "image":
      try { return [new Paragraph({ children: [ctx.image(block.path.split("/").pop())] })]; }
      catch { return [new Paragraph({ children: [new TextRun(`[image: ${block.path}]`)] })]; }
    default:
      return [];
  }
}

export const blocksToChildren = (blocks, ctx) => blocks.flatMap((b) => blockToChildren(b, ctx));
const textToChildren = (paras) =>
  (paras || []).map((t) => new Paragraph({ children: [new TextRun(t)], spacing: { after: 160, line: 276 } }));

// ---- Per-page children (generated module OR deterministic) ------------------

const SECTION_TYPE = { CONTINUOUS: SectionType.CONTINUOUS };

/**
 * Render the JSON-safe `{ sections: [{properties, blocks}] }` descriptor that
 * src/extract.js's buildPageSections() writes into the manifest (tables,
 * multi-column bands, and images already merged into one Y-ordered list) into
 * real docx sections. `properties.type` arrives as the string "CONTINUOUS"
 * (JSON can't hold the docx enum) and is resolved here; the first section
 * gets the default page setup, matching executeModule's LLM-path convention.
 */
export function renderJsonSections(sections, ctx) {
  return sections.map((s, i) => {
    const { type, ...rest } = s.properties || {};
    const properties = { ...(i === 0 ? DEFAULT_PROPS : {}), ...rest };
    if (type) properties.type = SECTION_TYPE[type] ?? type;
    return { properties, children: blocksToChildren(s.blocks || [], ctx) };
  });
}

// ---- Compact layout-JSON path (the ~2×-faster model output) -----------------
// The model emits a SMALL layout JSON (transcription + block assembly ONLY, no
// docx.js code — that is what makes it ~2× faster; see HANDOFF §"Measured
// findings"). This renderer expands it to real docx sections AND stamps the
// injected fills/borders/frame that pdfextract get_drawings() recovered — facts
// the model is never asked to decide. Given the fills it applies them 4/4; asked
// to infer the outer frame it drops it, three times — so we inject, not ask.
//
// Layout schema:
//   { title?, frame?:bool, styles?:{font,size}, columns:[ { blocks:[Block] } ] }
//   Block = {t:"bar",text,fill?,color?} | {t:"h",text,level?} | {t:"p",text}
//         | {t:"list",ordered?,items:[...]} | {t:"img",file}
//         | {t:"table", fill?, header?:bool, cols?:[relWidths], rows:[[cell,...]] }
//   cell = string | {text,bold?,italic?,align?}
// injected = { fills:[{fill,text,bbox}], border:{color,width}?, frame:{color,width}? }

const normHex = (h) => (h ? String(h).replace(/^#/, "").toUpperCase().slice(0, 6).padStart(6, "0") : null);

/** Relative luminance of a #RRGGBB → pick white or near-black text for contrast. */
function contrastText(hex) {
  const h = normHex(hex);
  if (!h) return undefined;
  const r = parseInt(h.slice(0, 2), 16), g = parseInt(h.slice(2, 4), 16), b = parseInt(h.slice(4, 6), 16);
  const lum = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
  return lum < 0.55 ? "FFFFFF" : "1A1A1A";
}

/** docx border `size` is in EIGHTHS of a point — a 1pt PDF stroke is size 8. */
const borderSize = (widthPt) => Math.max(2, Math.round((widthPt || 1) * 8));

const SIG_WORD = /[\p{L}\p{N}]{3,}/gu;
const sigWords = (s) => new Set((String(s || "").toLowerCase().match(SIG_WORD) || []));

/**
 * Match a block's text against the injected fills (each carries the words it
 * sits behind). A fill applies when the block and the fill share ≥2 significant
 * words, or one text contains the other — enough to bind "TRANSACTION RATIONALE"
 * to the green bar behind it without over-matching unrelated prose.
 * @returns {{fill:string, color:string}|null}
 */
function matchFill(text, injected) {
  const fills = injected?.fills || [];
  if (!fills.length) return null;
  const bw = sigWords(text);
  if (!bw.size) return null;
  let best = null, bestScore = 0;
  for (const f of fills) {
    if (!f.fill) continue;
    const fw = sigWords(f.text);
    if (!fw.size) continue;
    let shared = 0;
    for (const w of bw) if (fw.has(w)) shared++;
    const contains = fw.size <= bw.size ? [...fw].every((w) => bw.has(w)) : [...bw].every((w) => fw.has(w));
    const score = shared + (contains ? 1 : 0);
    if (score >= 2 && score > bestScore) { bestScore = score; best = f; }
  }
  if (!best) return null;
  const fill = normHex(best.fill);
  return { fill, color: contrastText(fill) };
}

/** A single-cell shaded bar (a colored section heading band), sized to `width`. */
function barElement(block, injected, width = CONTENT_DXA) {
  const explicit = block.fill ? { fill: normHex(block.fill), color: block.color && normHex(block.color) } : null;
  const stamped = explicit || matchFill(block.text, injected);
  const fill = stamped?.fill;
  const color = block.color ? normHex(block.color) : stamped?.color;
  const noBorder = { style: BorderStyle.NONE, size: 0, color: "auto" };
  return new Table({
    width: { size: width, type: WidthType.DXA },
    columnWidths: [width],
    borders: { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder, insideHorizontal: noBorder, insideVertical: noBorder },
    rows: [new TableRow({
      children: [new TableCell({
        width: { size: width, type: WidthType.DXA },
        ...(fill ? { shading: { fill } } : {}),
        margins: { top: 80, bottom: 80, left: 160, right: 160 },
        children: [new Paragraph({ children: [new TextRun({ text: block.text || "", bold: true, ...(color ? { color } : {}) })] })],
      })],
    })],
  });
}

const NUMERIC_CELL = /^-?\$?[\d,]*\.?\d+%?$/;
const cellText = (c) => (c == null ? "" : typeof c === "object" ? (c.text ?? "") : String(c));

/** A real docx table from a compact {rows, cols?, header?, fill?} block, sized to `width`. */
function jsonTableElement(block, injected, width = CONTENT_DXA) {
  const rows = block.rows || [];
  if (!rows.length) return null;
  const nCols = Math.max(...rows.map((r) => r.length));
  const rel = (block.cols && block.cols.length === nCols) ? block.cols : new Array(nCols).fill(1);
  const totalRel = rel.reduce((s, w) => s + w, 0);
  // Column widths sum to EXACTLY `width` (the containing column/page width) so a
  // nested table can never overflow its cell and get clipped off the right edge.
  const columnWidths = rel.map((w) => Math.max(1, Math.floor((w / totalRel) * width)));
  const drift = width - columnWidths.reduce((s, w) => s + w, 0);
  columnWidths[columnWidths.length - 1] += drift; // absorb rounding so the total is exact
  const hasHeader = block.header !== false;

  // Right-align columns whose data cells are all numeric.
  const numericCols = new Set();
  for (let c = 0; c < nCols; c++) {
    const data = rows.slice(hasHeader ? 1 : 0).map((r) => cellText(r[c]).trim()).filter(Boolean);
    if (data.length && data.every((t) => NUMERIC_CELL.test(t))) numericCols.add(c);
  }

  // Border weight comes from the injected stroke (1pt → size 8); a frame/border
  // means "draw gridlines", its absence means a clean editorial table.
  const bw = injected?.border || injected?.frame;
  const line = bw
    ? { style: BorderStyle.SINGLE, size: borderSize(bw.width), color: normHex(bw.color) || "808080" }
    : { style: BorderStyle.SINGLE, size: 4, color: "A6A6A6" };
  const headerFill = block.fill ? normHex(block.fill) : matchFill(cellText(rows[0]?.[0]) + " " + cellText(rows[0]?.[1]), injected)?.fill;
  const margins = { top: 60, bottom: 60, left: 120, right: 120 };

  return new Table({
    width: { size: width, type: WidthType.DXA },
    columnWidths,
    borders: { top: line, bottom: line, left: line, right: line, insideHorizontal: line, insideVertical: line },
    rows: rows.map((row, ri) => {
      const isHeader = hasHeader && ri === 0;
      return new TableRow({
        tableHeader: isHeader,
        children: Array.from({ length: nCols }, (_, ci) => {
          const cell = row[ci];
          const txt = cellText(cell);
          const bold = isHeader || (typeof cell === "object" && !!cell?.bold);
          const align = (typeof cell === "object" && cell?.align) ? cell.align
            : numericCols.has(ci) ? "right" : undefined;
          const color = isHeader && headerFill ? contrastText(headerFill) : undefined;
          return new TableCell({
            width: { size: columnWidths[ci], type: WidthType.DXA },
            margins, verticalAlign: "center",
            ...(isHeader && headerFill ? { shading: { fill: headerFill } } : {}),
            children: [new Paragraph({
              alignment: align === "right" ? AlignmentType.RIGHT : align === "center" ? AlignmentType.CENTER : undefined,
              children: [new TextRun({ text: txt, bold, italics: typeof cell === "object" && !!cell?.italic, ...(color ? { color } : {}) })],
            })],
          });
        }),
      });
    }),
  });
}

const spacer = () => new Paragraph({ children: [], spacing: { after: 120 } });

function jsonBlockToChildren(block, ctx, injected, width = CONTENT_DXA) {
  switch (block.t) {
    case "bar": return [barElement(block, injected, width), spacer()];
    case "h": return [new Paragraph({ heading: HEADING[Math.min(4, block.level || 1)], children: [new TextRun(block.text || "")], spacing: { before: 200, after: 100 } })];
    case "p": return [new Paragraph({ children: [new TextRun(block.text || "")], spacing: { after: 140, line: 276 } })];
    case "list": return (block.items || []).map((it, i) => new Paragraph({ children: [new TextRun((block.ordered ? `${i + 1}. ` : "• ") + it)], indent: { left: 360 }, spacing: { after: 60 } }));
    case "table": { const t = jsonTableElement(block, injected, width); return t ? [t, spacer()] : []; }
    case "img":
      // Cap the image at the container width (CONTENT_DXA→CONTENT_PX is /15) so a
      // figure in a narrow column can't overflow past the right edge either.
      try { return [new Paragraph({ children: [ctx.image(String(block.file).split("/").pop(), { maxWidth: Math.max(80, Math.floor(width / 15)) })] })]; }
      catch { return [new Paragraph({ children: [new TextRun(`[image: ${block.file}]`)] })]; }
    default: return [];
  }
}

// Cell padding (dxa) applied to each column cell in the outer layout table; the
// content inside a column is sized to (columnWidth − both paddings) so nested
// tables/bars fit exactly and never spill past the right border.
const CELL_PAD = 120;

/**
 * Expand a compact layout JSON (model output) into { sections, styles }, the
 * same shape executeModule/renderJsonSections return, stamping injected
 * fills/borders/frame. Multiple columns are rendered as an outer 1-row layout
 * table (one cell per column) so columns survive into Word; when a frame is
 * present that outer table carries the border + insideVertical rule, which is
 * exactly the frame the model kept dropping.
 */
export function renderLayoutJson(layout, ctx, injected = {}) {
  const columns = (layout.columns && layout.columns.length) ? layout.columns : [{ blocks: layout.blocks || [] }];
  const n = columns.length;

  // Draw the outer frame when the model says so, OR when the PDF's vector layer
  // shows the page is ruled: an explicit frame rect, or (for a multi-column page)
  // a dominant border stroke — that ruling IS the frame around the quadrant grid.
  const ruled = !!injected.frame || (n > 1 && !!injected.border);
  const wantFrame = !!layout.frame || ruled;

  const styles = (layout.styles && (layout.styles.font || layout.styles.size))
    ? { default: { document: { run: { ...(layout.styles.font ? { font: layout.styles.font } : {}), ...(layout.styles.size ? { size: layout.styles.size } : {}) } } } }
    : undefined;

  const title = layout.title
    ? [new Paragraph({ heading: HeadingLevel.HEADING_1, alignment: AlignmentType.CENTER, children: [new TextRun({ text: layout.title, bold: true })], spacing: { after: 160 } })]
    : [];

  // Single column, no frame → a plain single-section page at full content width.
  if (n === 1 && !wantFrame) {
    const kids = (columns[0].blocks || []).flatMap((b) => jsonBlockToChildren(b, ctx, injected, CONTENT_DXA));
    return { sections: [{ properties: DEFAULT_PROPS, children: [...title, ...(kids.length ? kids : [new Paragraph({ children: [] })])] }], styles };
  }

  // Multi-column (or framed): an outer 1-row layout table, one cell per column.
  // Each column's content is sized to (columnWidth − both cell paddings) so
  // nested tables/bars fit exactly inside the cell and never clip on the right.
  const colWidths = Array.from({ length: n }, () => Math.floor(CONTENT_DXA / n));
  colWidths[n - 1] += CONTENT_DXA - colWidths.reduce((s, w) => s + w, 0); // exact total
  const innerWidths = colWidths.map((w) => w - 2 * CELL_PAD);

  const colChildren = columns.map((col, i) => {
    const kids = (col.blocks || []).flatMap((b) => jsonBlockToChildren(b, ctx, injected, innerWidths[i]));
    return kids.length ? kids : [new Paragraph({ children: [] })];
  });

  const fsrc = injected.frame || injected.border;
  const on = { style: BorderStyle.SINGLE, size: fsrc ? borderSize(fsrc.width) : 8, color: normHex(fsrc?.color) || "000000" };
  const off = { style: BorderStyle.NONE, size: 0, color: "auto" };
  const outerBorders = wantFrame
    ? { top: on, bottom: on, left: on, right: on, insideHorizontal: off, insideVertical: on }
    : { top: off, bottom: off, left: off, right: off, insideHorizontal: off, insideVertical: off };

  const outer = new Table({
    width: { size: CONTENT_DXA, type: WidthType.DXA },
    columnWidths: colWidths,
    borders: outerBorders,
    rows: [new TableRow({
      children: colChildren.map((kids, i) => new TableCell({
        width: { size: colWidths[i], type: WidthType.DXA },
        margins: { top: 120, bottom: 120, left: CELL_PAD, right: CELL_PAD },
        verticalAlign: "top",
        children: kids,
      })),
    })],
  });

  return { sections: [{ properties: DEFAULT_PROPS, children: [...title, outer] }], styles };
}

/** Return { sections, styles } for a page (one or more Word sections). */
async function sectionsForPage(workDir, page, warn) {
  const assetsAbsDir = page.assetsDir ? join(workDir, page.assetsDir) : workDir;

  if (page.mode === "json" && page.layout) {
    try {
      return renderLayoutJson(page.layout, makeCtx(assetsAbsDir), page.injected || {});
    } catch (e) {
      warn(`page ${page.page}: layout JSON render failed (${e.message}); using plain-text fallback`);
      return { sections: [{ properties: DEFAULT_PROPS, children: textToChildren(page.fallbackText) }] };
    }
  }
  if (page.mode === "llm" && page.codeFile) {
    try {
      return await executeModule(join(workDir, page.codeFile), assetsAbsDir);
    } catch (e) {
      warn(`page ${page.page}: generated code failed (${e.message}); using plain-text fallback`);
      return { sections: [{ properties: DEFAULT_PROPS, children: textToChildren(page.fallbackText) }] };
    }
  }
  if (page.sections) {
    return { sections: renderJsonSections(page.sections, makeCtx(assetsAbsDir)) };
  }
  return { sections: [{ properties: DEFAULT_PROPS, children: blocksToChildren(page.blocks || [], makeCtx(assetsAbsDir)) }] };
}

/**
 * Build the final .docx Buffer. A page may emit several Word sections (e.g.
 * two-column text → continuous full-width table → two columns). The FIRST
 * section of each page (after the first page) is forced to start a new page;
 * a page's own continuation sections keep their CONTINUOUS type.
 */
export async function renderDocx(workDir, model, cfg, warn = console.warn) {
  const all = [];
  let docStyles, docHeader, docFooter, docPageStart;
  for (const page of model.pages) {
    const { sections, styles, header, footer, pageNumberStart } = await sectionsForPage(workDir, page, warn);
    if (styles && !docStyles) docStyles = styles;
    if (header && !docHeader) docHeader = header;
    if (footer && !docFooter) docFooter = footer;
    if (pageNumberStart != null && docPageStart == null) docPageStart = pageNumberStart;
    sections.forEach((s, si) => {
      if (all.length > 0 && si === 0) {
        s.properties = { ...s.properties, type: SectionType.NEXT_PAGE }; // page break between source pages
      }
      all.push(s);
    });
  }
  // Word headers/footers are per-section but inherit forward (linkedToPrevious):
  // set them on the FIRST section only and every later section picks them up.
  if (all.length) {
    if (docHeader) all[0].headers = { default: docHeader };
    if (docFooter) all[0].footers = { default: docFooter };
    if (docPageStart != null) {
      const props = all[0].properties = { ...all[0].properties };
      props.page = { ...props.page, pageNumbers: { ...props.page?.pageNumbers, start: docPageStart } };
    }
  }
  const doc = new Document(docStyles ? { styles: docStyles, sections: all } : { sections: all });
  return Packer.toBuffer(doc);
}
