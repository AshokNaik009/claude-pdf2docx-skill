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
  Table, TableRow, TableCell, WidthType, BorderStyle, ImageRun,
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
  if (Array.isArray(result)) {
    sections = result.every((s) => Array.isArray(s?.children)) ? result : [{ children: result }];
  } else if (result && Array.isArray(result.sections)) {
    sections = result.sections; styles = result.styles;
  } else if (result && Array.isArray(result.children)) {
    sections = [{ properties: result.properties, children: result.children }]; styles = result.styles;
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
  return { sections, styles };
}

/** ctx passed to generated modules — currently exposes image embedding. */
function makeCtx(assetsAbsDir) {
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

function tableElement(block) {
  const rows = block.rows.filter((r) => r.length);
  if (!rows.length) return null;
  const cols = Math.max(...rows.map((r) => r.length));
  const colW = Math.floor(CONTENT_DXA / cols);
  const line = { style: BorderStyle.SINGLE, size: 4, color: "A6A6A6" };
  const margins = { top: 80, bottom: 80, left: 120, right: 120 };
  return new Table({
    width: { size: CONTENT_DXA, type: WidthType.DXA },
    columnWidths: Array.from({ length: cols }, () => colW),
    rows: rows.map((r, ri) =>
      new TableRow({
        tableHeader: ri === 0,
        children: Array.from({ length: cols }, (_, ci) =>
          new TableCell({
            width: { size: colW, type: WidthType.DXA },
            margins, borders: { top: line, bottom: line, left: line, right: line },
            verticalAlign: "center",
            children: [new Paragraph({ children: [new TextRun({ text: r[ci] ?? "", bold: ri === 0 })] })],
          })
        ),
      })
    ),
    borders: { top: line, bottom: line, left: line, right: line, insideHorizontal: line, insideVertical: line },
  });
}

function blockToChildren(block, ctx) {
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

const blocksToChildren = (blocks, ctx) => blocks.flatMap((b) => blockToChildren(b, ctx));
const textToChildren = (paras) =>
  (paras || []).map((t) => new Paragraph({ children: [new TextRun(t)], spacing: { after: 160, line: 276 } }));

// ---- Per-page children (generated module OR deterministic) ------------------

/** Return { sections, styles } for a page (one or more Word sections). */
async function sectionsForPage(workDir, page, warn) {
  const assetsAbsDir = page.assetsDir ? join(workDir, page.assetsDir) : workDir;

  if (page.mode === "llm" && page.codeFile) {
    try {
      return await executeModule(join(workDir, page.codeFile), assetsAbsDir);
    } catch (e) {
      warn(`page ${page.page}: generated code failed (${e.message}); using plain-text fallback`);
      return { sections: [{ properties: DEFAULT_PROPS, children: textToChildren(page.fallbackText) }] };
    }
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
  let docStyles;
  for (const page of model.pages) {
    const { sections, styles } = await sectionsForPage(workDir, page, warn);
    if (styles && !docStyles) docStyles = styles;
    sections.forEach((s, si) => {
      if (all.length > 0 && si === 0) {
        s.properties = { ...s.properties, type: SectionType.NEXT_PAGE }; // page break between source pages
      }
      all.push(s);
    });
  }
  const doc = new Document(docStyles ? { styles: docStyles, sections: all } : { sections: all });
  return Packer.toBuffer(doc);
}
