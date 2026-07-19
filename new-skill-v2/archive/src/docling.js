// Deterministic Docling → DOCX converter.
//
// Consumes the DoclingDocument JSON produced by scripts/docling_extract.py and
// builds a real Word file with the `docx` library — no LLM, no pandoc. It walks
// body.children in reading order and maps each element:
//   page_header / page_footer  -> dropped (running headers/footers = noise)
//   title                       -> Title paragraph
//   section_header             -> bold heading
//   caption                    -> centered, bold label + rest
//   text                        -> justified body paragraph
//   table                       -> docx.Table rebuilt from the cell grid (spans, header row)
//   picture                     -> centered image (from the extracted images/ dir)
// Groups (e.g. list containers) are flattened by recursing into their children.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import * as docx from "docx";

const { Document, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
        AlignmentType, WidthType, VerticalAlign, HeadingLevel } = docx;

const LETTER_CONTENT_PX = 624; // Letter page minus 1" margins each side

/** Resolve a Docling "#/texts/3" style $ref against the loaded document. */
function deref(doc, ref) {
  const [, kind, idx] = ref.split("/"); // "#","texts","3"
  return { kind, idx: Number(idx), node: doc[kind]?.[Number(idx)] };
}

/** Read intrinsic PNG dimensions from the IHDR chunk (no image library). */
function pngSize(bytes) {
  // 8-byte signature, 4-byte length, 4-byte "IHDR", then width/height uint32 BE.
  const width = bytes.readUInt32BE(16);
  const height = bytes.readUInt32BE(20);
  return { width, height };
}

/** Italicize every "et al." so citations render correctly. */
function runsFromText(text) {
  return text
    .split(/(et al\.)/g)
    .filter((s) => s !== "")
    .map((s) =>
      s === "et al." ? new TextRun({ text: s, italics: true }) : new TextRun(s)
    );
}

function bodyPara(text) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { after: 160, line: 276 },
    children: runsFromText(text),
  });
}

function headingPara(text, level) {
  return new Paragraph({
    heading: level,
    spacing: { before: 220, after: 120 },
    children: [new TextRun({ text, bold: true })],
  });
}

function captionPara(text) {
  // Split a leading "Fig. 2:" / "Table 1:" label so it renders bold.
  const m = text.match(/^((?:Fig\.?|Figure|Table|Scheme)\s*\S*[:.])\s*(.*)$/s);
  const children = m
    ? [new TextRun({ text: m[1] + " ", bold: true }), new TextRun(m[2])]
    : [new TextRun({ text, italics: true })];
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 120, after: 160 },
    children,
  });
}

/** Build a docx.Table from a Docling table's cell grid, honoring spans + header. */
function buildTable(tableNode) {
  const data = tableNode.data || {};
  const grid = data.grid; // grid[r][c] cell objects (spans duplicated across covered cells)
  if (!Array.isArray(grid) || grid.length === 0) return null;

  const rows = grid.map((gridRow, r) => {
    const cells = [];
    for (let c = 0; c < gridRow.length; c++) {
      const cell = gridRow[c];
      if (!cell) continue;
      // Emit a cell only at its top-left origin; skip the covered positions so
      // rowSpan/columnSpan line up (docx.js expects covered cells omitted).
      const originRow = cell.start_row_offset_idx ?? r;
      const originCol = cell.start_col_offset_idx ?? c;
      if (originRow !== r || originCol !== c) continue;

      const isHeader = !!cell.column_header;
      const text = (cell.text || "").trim();
      cells.push(
        new TableCell({
          columnSpan: cell.col_span > 1 ? cell.col_span : undefined,
          rowSpan: cell.row_span > 1 ? cell.row_span : undefined,
          verticalAlign: VerticalAlign.TOP,
          shading: isHeader ? { fill: "EFEFEF" } : undefined,
          children: [
            new Paragraph({
              children: [new TextRun({ text: text || " ", bold: isHeader })],
            }),
          ],
        })
      );
    }
    // A TableRow with zero cells throws; guard fully-covered rows (shouldn't happen).
    if (cells.length === 0) return null;
    return new TableRow({
      tableHeader: gridRow.some((cc) => cc && cc.column_header),
      children: cells,
    });
  }).filter(Boolean);

  if (rows.length === 0) return null;
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows,
  });
}

/** Center-placed figure sized to the content width, aspect ratio preserved. */
function buildPicture(pictureNode, pictureIndex, imagesDir) {
  // scripts/docling_extract.py saves pictures as images/pic_<index+1>.png
  const file = join(imagesDir, `pic_${pictureIndex + 1}.png`);
  let bytes;
  try {
    bytes = readFileSync(file);
  } catch {
    return new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: "[figure not extracted]", italics: true })],
    });
  }
  const { width: iw, height: ih } = pngSize(bytes);
  const w = Math.min(LETTER_CONTENT_PX, iw);
  const h = Math.round((ih / iw) * w);
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    keepNext: true, // tie figure to the caption that follows
    spacing: { before: 120 },
    children: [
      new ImageRun({ type: "png", data: bytes, transformation: { width: w, height: h } }),
    ],
  });
}

const HEADING_LEVELS = [
  HeadingLevel.HEADING_1,
  HeadingLevel.HEADING_2,
  HeadingLevel.HEADING_3,
];

/** Walk body.children (recursing groups) → flat array of docx elements. */
function walk(doc, children, imagesDir, out) {
  let pictureCount = 0;
  const visit = (refList) => {
    for (const ref of refList) {
      const { kind, idx, node } = deref(doc, ref.$ref);
      if (!node) continue;

      if (kind === "groups") {
        visit(node.children || []); // flatten list/inline groups
        continue;
      }
      if (kind === "tables") {
        const t = buildTable(node);
        if (t) out.push(t);
        continue;
      }
      if (kind === "pictures") {
        out.push(buildPicture(node, idx, imagesDir));
        pictureCount++;
        continue;
      }
      if (kind === "texts") {
        const label = node.label;
        const text = (node.text ?? node.orig ?? "").trim();
        if (!text) continue;
        if (label === "page_header" || label === "page_footer") continue; // drop noise
        if (label === "title") { out.push(headingPara(text, HEADING_LEVELS[0])); continue; }
        if (label === "section_header") {
          const lvl = HEADING_LEVELS[Math.min((node.level ?? 1) - 1, 2)] ?? HEADING_LEVELS[1];
          out.push(headingPara(text, lvl));
          continue;
        }
        if (label === "caption") { out.push(captionPara(text)); continue; }
        out.push(bodyPara(text)); // default: body text
      }
    }
  };
  visit(children);
  return out;
}

/** Build a docx.Document from a loaded DoclingDocument + its images dir. */
export function buildDoclingDocx(doc, imagesDir) {
  const elements = walk(doc, doc.body?.children || [], imagesDir, []);
  return new Document({
    styles: { default: { document: { run: { font: "Calibri", size: 21 } } } },
    sections: [
      {
        properties: { page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
        children: elements.length ? elements : [new Paragraph({})],
      },
    ],
  });
}

/** Convenience: load doc.json from a docling output dir and pack to a Buffer. */
export async function renderDoclingDocx(doclingOutDir) {
  const doc = JSON.parse(readFileSync(join(doclingOutDir, "doc.json"), "utf8"));
  const document = buildDoclingDocx(doc, join(doclingOutDir, "images"));
  return docx.Packer.toBuffer(document);
}
