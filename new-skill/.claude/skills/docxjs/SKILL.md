---
name: docxjs
description: >-
  Generate Word (.docx) documents in JavaScript with the `docx` npm library (docx.js) — the
  browser/Node library using `new docx.Document(...)`, `docx.Table`, and `docx.Packer.toBlob`/
  `saveAs`. Use whenever the user wants JS/TS/browser code to build a Word doc, mentions docx.js,
  `Packer.toBlob`, `saveAs`, or a `generate()`/`generateDocument()` returning a .docx. Applies to
  ANY domain — academic papers and especially banking, asset management, and finance: factsheets,
  portfolio/holdings tables, returns tables, fee schedules, credit memos, board packs, and
  financial statements. Reach for it for multi-column layouts, justified text, section breaks
  mixing column counts, tables with header rows, merged (row/col-spanned) cells, right- or
  decimal-aligned numeric columns, totals rows, currency/percent formatting, captions, embedded
  images, footnotes, page numbers. DIFFERENT from the Python `python-docx` skill — use this for
  the JavaScript library, even when the user pastes a docx.js snippet. Never hand-write OOXML.
---

# Generating .docx with docx.js

The `docx` npm package (a.k.a. docx.js) builds Word documents from a tree of JS objects. The same
code runs in the browser (download via FileSaver's `saveAs`) and in Node (write bytes to disk). This
skill covers the building blocks plus the layout/table edge cases that trip people up. The patterns
are **domain-neutral** — the same two-column bodies, section breaks, and table constructs serve an
academic paper, a fund factsheet, a portfolio-holdings appendix, or a credit memo. Finance
deliverables lean hardest on the table features (numeric alignment, totals rows, currency/percent),
so those get their own section in `references/tables.md` and a dedicated worked example.

## When you reach for the reference files

Keep this file for the common 80%. Go deeper when the task needs it:

- **`references/tables.md`** — borders/styles, column widths, cell margins & shading, header-row
  repeat, `columnSpan`, `rowSpan`/`vMerge`, multi-paragraph & list cells, vertical alignment,
  nested tables, keeping rows from splitting — **plus a Finance section**: right-aligned and
  decimal-aligned numeric columns, currency/percent formatting, thousands separators, totals and
  subtotal rows, RAG/status shading, and full recipes for a holdings table and a periodic-returns
  table. Read it for anything beyond a plain grid.
- **`references/layout.md`** — sections & continuous breaks, mixing column counts on one page,
  column/page breaks, margins, page size & orientation, headers/footers (default/first/even),
  page numbering, tab stops, and the full units cheat-sheet. Read it for any multi-section or
  multi-column document (factsheets, board packs, reports).
- **`scripts/generate-example.js`** — a complete, runnable academic excerpt (two-column justified
  text → full-width captioned table → figure → back to two columns).
- **`scripts/generate-finance-example.js`** — a complete, runnable **fund factsheet / portfolio
  page**: a two-column commentary header, a right-aligned periodic-returns table, a holdings table
  with a totals row and thousands separators, and a disclosures footnote. Copy from whichever
  worked example is closer to the task.

## Units (memorize these — wrong units is the #1 bug)

| Property | Unit | Example |
|---|---|---|
| `size` (font) | half-points | `22` = 11pt |
| `spacing.before` / `after` | twips (1/20 pt) | `200` ≈ 10pt |
| `spacing.line` | 240ths | `240` single, `288` ≈1.15, `360` 1.5, `480` double |
| `indent.left` / `firstLine` / `hanging` | twips | `720` = 0.5" |
| `page.margin.*`, `column.space` | twips | `1440` = 1" |
| `ImageRun.transformation` w/h | pixels | `{ width: 400, height: 260 }` |
| table/cell width (`PERCENTAGE`) | percent number | `size: 100` = 100% |
| table/cell width (`DXA`) | twips | `size: 2400` |

## The two harnesses

Write the document body once as a `generateDocument()` that **returns a `docx.Document`**. Only the
final packing step differs between browser and Node.

**Browser** (needs `docx` UMD + FileSaver globals loaded via `<script>`):

```js
docx.Packer.toBlob(generateDocument()).then((blob) => saveAs(blob, "output.docx"));
```

CDN tags that work in a real browser:

```html
<script src="https://cdn.jsdelivr.net/npm/docx@9/build/index.umd.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js"></script>
```

**Node** (for testing/verifying, or server-side generation):

```js
const docx = require("docx");
const fs = require("fs");
docx.Packer.toBuffer(generateDocument()).then((buf) => fs.writeFileSync("output.docx", buf));
```

Everything below is library-identical across both; only swap the last line.

## Core building blocks

```js
// A run = a span of text with one set of formatting.
new docx.TextRun("plain");
new docx.TextRun({ text: "bold", bold: true, italics: true, size: 24, color: "444444", font: "Calibri" });

// A paragraph = a line/block made of runs.
new docx.Paragraph({
  alignment: docx.AlignmentType.JUSTIFIED,      // LEFT | CENTER | RIGHT | JUSTIFIED
  spacing: { after: 180, line: 288 },           // space after + 1.15 line spacing
  children: [ new docx.TextRun("...") ],
});
```

Set a **document-wide default font/size** once instead of on every run:

```js
new docx.Document({
  styles: { default: { document: { run: { font: "Palatino Linotype", size: 22 } } } },
  sections: [ /* ... */ ],
});
```

## Reusable helpers (put these at the top of the file)

These make academic bodies clean. The citation helper auto-italicizes `et al.` — a recurring need.

```js
// Italicize every "et al." so citations render like (Boulet et al. 2010) with et al. in italics.
function runsFromText(text) {
  return text.split(/(et al\.)/g)
    .filter((s) => s !== "")
    .map((s) => s === "et al."
      ? new docx.TextRun({ text: s, italics: true })
      : new docx.TextRun(s));
}

// Justified body paragraph with sensible spacing.
function bodyPara(text) {
  return new docx.Paragraph({
    alignment: docx.AlignmentType.JUSTIFIED,
    spacing: { after: 180, line: 288 },
    children: runsFromText(text),
  });
}

// Bold sub-heading, e.g. "Evaporation of Moisture".
function heading(text) {
  return new docx.Paragraph({
    spacing: { before: 200, after: 120 },
    children: [ new docx.TextRun({ text, bold: true }) ],
  });
}

// Centered caption: bold label + normal rest, e.g. "Fig. 2: ..." or "Table 1: ...".
function caption(label, rest) {
  return new docx.Paragraph({
    alignment: docx.AlignmentType.CENTER,
    spacing: { before: 120, after: 120 },
    children: [ new docx.TextRun({ text: label, bold: true }), new docx.TextRun(rest) ],
  });
}
```

If a paragraph mixes formatting beyond `et al.` (e.g. an inline bold term), build its `children`
array by hand with multiple `TextRun`s rather than forcing it through `runsFromText`.

## Multi-column layouts and mixing column counts (the headline edge case)

Two columns come from **one property on the section**, and text flows/balances automatically:

```js
sections: [{
  properties: {
    column: { count: 2, space: 560, separate: false }, // space = twips gutter; separate draws a line
    page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } },
  },
  children: [ bodyPara("..."), heading("..."), bodyPara("...") ],
}]
```

A very common academic layout is **two-column text, then a full-width table or figure, then back
to two columns** (exactly the Table 1 / Fig. 2 pattern). Word does this with *continuous section
breaks*. In docx.js, **each element of the `sections` array is a Word section**, so emit three
sections and mark the later ones `type: docx.SectionType.CONTINUOUS` so they don't start new pages:

```js
sections: [
  { properties: { column: { count: 2, space: 560 } },
    children: [ bodyPara("...two-column intro...") ] },

  { properties: { type: docx.SectionType.CONTINUOUS },        // single column, no page break
    children: [ caption("Table 1: ", "Mechanism of heat transfer"), table /* full width */ ] },

  { properties: { type: docx.SectionType.CONTINUOUS, column: { count: 2, space: 560 } },
    children: [ bodyPara("...back to two columns...") ] },
]
```

Put document-wide `styles` and page `margin`/`size` on the **first** section (or the document
`styles` block); continuation sections inherit page setup but you re-declare `column` each time.
To force a break *within* one column layout instead of a new section, use a column break — see
`references/layout.md`.

## Tables (essentials)

A table is rows of cells; each cell holds **paragraphs** (so a cell can contain several paragraphs
or list items). Header rows get `tableHeader: true` so they repeat if the table spans pages.

```js
const P = (t, opts = {}) => new docx.Paragraph({ children: [new docx.TextRun(t)], ...opts });
const B = (t) => new docx.Paragraph({ children: [new docx.TextRun({ text: t, bold: true })] });

const table = new docx.Table({
  width: { size: 100, type: docx.WidthType.PERCENTAGE },
  columnWidths: [2200, 3400, 3400],           // twips; keeps columns stable
  rows: [
    new docx.TableRow({
      tableHeader: true,
      children: [
        new docx.TableCell({ shading: { fill: "EFEFEF" }, children: [B("Mode of Heat Transfer")] }),
        new docx.TableCell({ children: [B("Source")] }),
        new docx.TableCell({ children: [B("role")] }),
      ],
    }),
    new docx.TableRow({
      children: [
        new docx.TableCell({ verticalAlign: docx.VerticalAlign.TOP, children: [P("Conduction")] }),
        new docx.TableCell({ children: [                    // multi-paragraph / numbered cell
          P("(i)\tDirect contact between the product and the carrier medium (baking tray, mould, or conveyor band)"),
          P("(ii)\tHeat transfer within the product matrix from the surface toward the core"),
        ]}),
        new docx.TableCell({ children: [P("Heat is transferred through direct contact ...")] }),
      ],
    }),
  ],
});
```

- **Merge across columns:** `new docx.TableCell({ columnSpan: 3, children: [...] })`.
- **Merge down rows:** `rowSpan: 2` on the top cell; **omit** the covered cell in the row below.
- **Borders:** by default cells have thin borders. To style or remove them, set `borders` on the
  table or per cell — see `references/tables.md`.
- **Cell padding:** `margins: { top, bottom, left, right }` (twips) on the table or cell.

**Numeric columns (finance).** Money and percentages read correctly only when right-aligned and
formatted with separators. Use a right-aligned cell helper and JS number formatting:

```js
const fmtMoney = (n, ccy = "USD") =>
  new Intl.NumberFormat("en-US", { style: "currency", currency: ccy, maximumFractionDigits: 0 }).format(n);
const fmtPct = (n) => `${(n >= 0 ? "+" : "")}${n.toFixed(1)}%`;

// right-aligned numeric cell
const numCell = (text, opts = {}) => new docx.TableCell({
  ...opts,
  children: [ new docx.Paragraph({ alignment: docx.AlignmentType.RIGHT, children: [new docx.TextRun(text)] }) ],
});

numCell(fmtMoney(1250000));   // "$1,250,000"
numCell(fmtPct(-3.4));         // "-3.4%"  (right-aligned)
```

Right-alignment keeps decimal points visually stacked for equal-decimal columns; for ragged
decimals use a DECIMAL tab stop instead (see `references/tables.md` → Finance). For a totals row,
shade it and bold the figures. For anything past a plain grid, read `references/tables.md`.

## Images and figure captions

```js
// data: a Buffer/Uint8Array/base64 of the image; type: "png" | "jpg" | "gif" | "bmp".
new docx.Paragraph({
  alignment: docx.AlignmentType.CENTER,
  children: [ new docx.ImageRun({ type: "png", data: imgBytes, transformation: { width: 460, height: 300 } }) ],
});
// then:
caption("Fig. 2: ", "Mechanism of heat transfer during baking process inside the baking oven");
```

In the browser, get `imgBytes` from a `File`/`fetch` as an `ArrayBuffer`/`Uint8Array`; in Node use
`fs.readFileSync("fig.png")`. Set `width`/`height` in pixels at the display size you want.

## Headers, footers, and page numbers

```js
sections: [{
  headers: {
    default: new docx.Header({ children: [ new docx.Paragraph({
      alignment: docx.AlignmentType.RIGHT,
      children: [ new docx.TextRun("My Title  "),
                  new docx.TextRun({ children: ["Page ", docx.PageNumber.CURRENT] }) ] }) ] }),
    first: new docx.Header({ children: [ /* different header on page 1 */ ] }),
  },
  footers: { default: new docx.Footer({ children: [ /* ... */ ] }) },
  properties: { page: { pageNumbers: { start: 1 } } },
  children: [ /* ... */ ],
}]
```

More (even/odd headers, restarting page numbers per section, TOC) is in `references/layout.md`.

## Common pitfalls

- **Wrong units** — `size` is half-points (double the pt you want); margins/indents are twips
  (multiply inches by 1440). See the units table.
- **`italics`, not `italic`** — the property is spelled `italics`.
- **A cell needs a `Paragraph`, not a raw `TextRun`** — `children` of a `TableCell` must be
  paragraphs (or nested tables).
- **Mixing column counts needs separate sections** — you can't switch from 2 columns to full-width
  and back inside one section; emit multiple sections with `type: docx.SectionType.CONTINUOUS`.
- **Missing covered cell after a `rowSpan`** — after a `rowSpan: 2`, the next row must have one
  *fewer* cell (don't add a placeholder), or the grid shifts.
- **Forgetting `tableHeader: true`** — long tables then lose their header row after a page break.
- **Trying to render docx.js inside an artifact sandbox** — generate/download works best in a real
  browser page or in Node; if you build an HTML harness, tell the user to open the file locally.

## Workflow

1. Draft the body as `generateDocument()` returning `new docx.Document(...)`, using the helpers.
2. For multi-column + full-width tables/figures, split into continuous sections as shown.
3. If you have Node available, **verify it packs**: run the Node harness to write a `.docx` and
   confirm it opens (a quick `unzip -p out.docx word/document.xml | grep` check catches structural
   issues). `scripts/generate-example.js` is a ready template to copy from.
4. Deliver the JS (a `generate()` function, or a full HTML page with the CDN tags and a button).
