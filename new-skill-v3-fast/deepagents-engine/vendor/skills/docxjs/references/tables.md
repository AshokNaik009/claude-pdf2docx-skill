# docx.js tables — deep reference

Contents: [Anatomy](#anatomy) · [Widths & fixed layout](#widths) · [Borders](#borders) ·
[Cell padding & shading](#padding-shading) · [Vertical alignment](#valign) ·
[Column spans](#colspan) · [Row spans / vertical merge](#rowspan) ·
[Multi-paragraph & list cells](#multi) · [Header rows & page breaks](#headers) ·
[Text direction / rotated cells](#rotate) · [Nested tables](#nested) · [Full recipe](#recipe) ·
[Finance tables](#finance)

<a id="anatomy"></a>
## Anatomy

```
Table
└─ rows: [ TableRow ]
     └─ children: [ TableCell ]
          └─ children: [ Paragraph | Table ]   // cells hold paragraphs (or nested tables), never bare runs
```

Every cell must contain at least one `Paragraph`. An empty cell still needs
`children: [ new docx.Paragraph({}) ]`.

<a id="widths"></a>
## Widths & fixed layout

Percentage width (relative to page/containing width):

```js
new docx.Table({ width: { size: 100, type: docx.WidthType.PERCENTAGE }, rows: [...] });
```

Fixed column widths (twips) — pair `columnWidths` with per-cell `width` for stable layout:

```js
new docx.Table({
  columnWidths: [2200, 3400, 3400],                    // must sum to the table width
  layout: docx.TableLayoutType.FIXED,                  // honor the widths exactly (don't autofit)
  rows: [ new docx.TableRow({ children: [
    new docx.TableCell({ width: { size: 2200, type: docx.WidthType.DXA }, children: [P("...")] }),
    // ...
  ]})],
});
```

`WidthType`: `PERCENTAGE` (size = percent number, e.g. 100), `DXA` (size = twips), `AUTO`, `NIL`.

<a id="borders"></a>
## Borders

Border style enum: `docx.BorderStyle.SINGLE | DOUBLE | DASHED | DOTTED | THICK | NONE` etc.
A border spec is `{ style, size, color }` where `size` is in eighths of a point (`size: 4` = 0.5pt).

Whole-table borders:

```js
new docx.Table({
  borders: {
    top:    { style: docx.BorderStyle.SINGLE, size: 4, color: "000000" },
    bottom: { style: docx.BorderStyle.SINGLE, size: 4, color: "000000" },
    left:   { style: docx.BorderStyle.SINGLE, size: 4, color: "000000" },
    right:  { style: docx.BorderStyle.SINGLE, size: 4, color: "000000" },
    insideHorizontal: { style: docx.BorderStyle.SINGLE, size: 2, color: "999999" },
    insideVertical:   { style: docx.BorderStyle.SINGLE, size: 2, color: "999999" },
  },
  rows: [...],
});
```

Borderless table (e.g. an "editorial" table like the Table 1 example, which shows only faint or no
gridlines): set every border to `NONE`, or per cell:

```js
const noBorders = {
  top: { style: docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
  bottom: { style: docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
  left: { style: docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
  right: { style: docx.BorderStyle.NONE, size: 0, color: "FFFFFF" },
};
new docx.TableCell({ borders: noBorders, children: [P("...")] });
```

A common academic look is a rule under the header row only: give header cells a `bottom` border and
all others `NONE`.

<a id="padding-shading"></a>
## Cell padding & shading

Padding (a.k.a. cell margins), in twips, on the table (applies to all cells) or per cell:

```js
new docx.Table({ margins: { top: 60, bottom: 60, left: 100, right: 100 }, rows: [...] });
new docx.TableCell({ margins: { top: 60, bottom: 60, left: 100, right: 100 }, children: [...] });
```

Background shading (hex fill, no `#`):

```js
new docx.TableCell({ shading: { fill: "EFEFEF" }, children: [B("Header")] });
```

<a id="valign"></a>
## Vertical alignment

```js
new docx.TableCell({ verticalAlign: docx.VerticalAlign.CENTER, children: [P("...")] });
// TOP | CENTER | BOTTOM
```

<a id="colspan"></a>
## Column spans (merge across)

```js
new docx.TableRow({ children: [
  new docx.TableCell({ columnSpan: 3, children: [P("spans all three columns")] }),
]});
```

The row with a span still needs only the cells you actually draw — a `columnSpan: 3` cell counts as
three grid columns, so don't add the two it covers.

<a id="rowspan"></a>
## Row spans / vertical merge (merge down)

Put `rowSpan: N` on the **top** cell and **omit** the covered cells in the following rows:

```js
new docx.Table({ rows: [
  new docx.TableRow({ children: [
    new docx.TableCell({ rowSpan: 2, children: [P("Conduction")] }),  // spans this row + next
    new docx.TableCell({ children: [P("(i) Direct contact ...")] }),
  ]}),
  new docx.TableRow({ children: [
    // NO cell for the merged column here — it is covered by the rowSpan above
    new docx.TableCell({ children: [P("(ii) Heat transfer within the product matrix ...")] }),
  ]}),
]});
```

This emits `vMerge` (restart on the top cell, continue below). Getting the cell count wrong in the
covered row is the usual cause of a mangled grid.

<a id="multi"></a>
## Multi-paragraph and list-style cells

Because a cell's `children` is an array of paragraphs, put several paragraphs in one cell for
sub-points. For the "(i) … (ii) …" style seen in academic tables, either fake it with a tab:

```js
new docx.TableCell({ children: [
  new docx.Paragraph({ children: [new docx.TextRun("(i)\tDirect contact between the product and the carrier medium")] }),
  new docx.Paragraph({ children: [new docx.TextRun("(ii)\tHeat transfer within the product matrix from the surface toward the core")] }),
]});
```

…or use a real numbered list by declaring `numbering` on the document and referencing it with
`numbering: { reference: "cell-list", level: 0 }` on each paragraph (see `layout.md` for the
`numbering` config block). Tabs are simpler when you want literal "(i)/(ii)" labels.

<a id="headers"></a>
## Header rows & page breaks

- `tableHeader: true` on a `TableRow` makes it **repeat** at the top of each page the table spans.
- `cantSplit: true` on a `TableRow` keeps that row from breaking across a page.

```js
new docx.TableRow({ tableHeader: true, cantSplit: true, children: [...] });
```

<a id="rotate"></a>
## Rotated / vertical text cells

```js
new docx.TableCell({ textDirection: docx.TextDirection.BOTTOM_TO_TOP_LEFT_TO_RIGHT, children: [P("Mode")] });
```

<a id="nested"></a>
## Nested tables

A cell can contain another `Table` in its `children` alongside or instead of paragraphs. Keep the
outer column width generous so the inner table has room.

<a id="recipe"></a>
## Full recipe: the "Mechanism of heat transfer" table

Reproduces the three-column academic table (Mode / Source / role) with a shaded, bold, repeating
header, a rule under the header, top-aligned cells, and a multi-paragraph Conduction cell.

```js
const P = (t) => new docx.Paragraph({ children: [new docx.TextRun(t)], spacing: { after: 40 } });
const B = (t) => new docx.Paragraph({ children: [new docx.TextRun({ text: t, bold: true })] });
const NONE = { style: docx.BorderStyle.NONE, size: 0, color: "FFFFFF" };
const cellNoBorder = { top: NONE, bottom: NONE, left: NONE, right: NONE };
const headRule = { ...cellNoBorder, bottom: { style: docx.BorderStyle.SINGLE, size: 4, color: "000000" } };

const table = new docx.Table({
  width: { size: 100, type: docx.WidthType.PERCENTAGE },
  columnWidths: [2000, 4000, 4000],
  layout: docx.TableLayoutType.FIXED,
  margins: { top: 40, bottom: 40, left: 80, right: 80 },
  rows: [
    new docx.TableRow({
      tableHeader: true,
      children: [
        new docx.TableCell({ borders: headRule, children: [B("Mode of Heat Transfer")] }),
        new docx.TableCell({ borders: headRule, children: [B("Source")] }),
        new docx.TableCell({ borders: headRule, children: [B("role")] }),
      ],
    }),
    new docx.TableRow({ children: [
      new docx.TableCell({ borders: cellNoBorder, verticalAlign: docx.VerticalAlign.TOP, children: [P("Radiation")] }),
      new docx.TableCell({ borders: cellNoBorder, children: [P("Emission of thermal energy through oven walls, heating elements, and surrounding hot surfaces.")] }),
      new docx.TableCell({ borders: cellNoBorder, children: [P("Heat is transferred by thermal radiation from hot oven walls and heating elements to the product surface. This mechanism dominates at high temperatures and contributes significantly to rapid surface heating and crust formation.")] }),
    ]}),
    new docx.TableRow({ children: [
      new docx.TableCell({ borders: cellNoBorder, verticalAlign: docx.VerticalAlign.TOP, children: [P("Convection")] }),
      new docx.TableCell({ borders: cellNoBorder, children: [P("Circulation of hot, humid air within the oven chamber.")] }),
      new docx.TableCell({ borders: cellNoBorder, children: [P("Heat transfer occurs through the movement of hot air around the product. Regulates heat exchange at the product surface, controls moisture evaporation, and ensures uniform temperature distribution for uniform baking.")] }),
    ]}),
    new docx.TableRow({ children: [
      new docx.TableCell({ borders: cellNoBorder, verticalAlign: docx.VerticalAlign.TOP, children: [P("Conduction")] }),
      new docx.TableCell({ borders: cellNoBorder, children: [
        new docx.Paragraph({ children: [new docx.TextRun("(i)\tDirect contact between the product and the carrier medium (baking tray, mould, or conveyor band)")], spacing: { after: 80 } }),
        new docx.Paragraph({ children: [new docx.TextRun("(ii)\tHeat transfer within the product matrix from the surface toward the core")] }),
      ]}),
      new docx.TableCell({ borders: cellNoBorder, children: [P("Heat is transferred through direct contact between the product and heated carrier surfaces and further conducted internally within the product")] }),
    ]}),
  ],
});
```

<a id="finance"></a>
## Finance tables (banking / asset management)

Financial tables live or die on **alignment and formatting**: labels left, numbers right, decimals
lined up, totals set apart. The mechanics below cover holdings tables, periodic-returns grids, fee
schedules, and statement-style layouts.

### Format numbers in JS, align in the cell

docx.js stores plain text — do the formatting with `Intl.NumberFormat` (or your own), then
right-align the cell so columns of figures line up.

```js
const money = (n, ccy = "USD", dp = 0) =>
  new Intl.NumberFormat("en-US", { style: "currency", currency: ccy, minimumFractionDigits: dp, maximumFractionDigits: dp }).format(n);
const num = (n, dp = 0) => new Intl.NumberFormat("en-US", { minimumFractionDigits: dp, maximumFractionDigits: dp }).format(n);
const pct = (n, dp = 1) => `${n >= 0 ? "+" : ""}${n.toFixed(dp)}%`;

const R = (t, opts = {}) => new docx.TableCell({ ...opts,
  children: [new docx.Paragraph({ alignment: docx.AlignmentType.RIGHT, children: [new docx.TextRun(t)] })] });
const L = (t, opts = {}) => new docx.TableCell({ ...opts,
  children: [new docx.Paragraph({ children: [new docx.TextRun(t)] })] });
```

### Decimal alignment for ragged decimals

If figures have different decimal lengths (12.5 vs 12.53), right-align alone misaligns the point.
Use a DECIMAL tab stop so the decimal point anchors:

```js
new docx.Paragraph({
  tabStops: [{ type: docx.TabStopType.DECIMAL, position: docx.TabStopPosition.MAX }],
  children: [new docx.TextRun("\t" + num(12.5, 2))],   // leading \t sends it to the decimal stop
});
```

### Negative numbers in red / parentheses

Convention varies; support both. Colour is a run property; parentheses are just formatting:

```js
function signed(n, fmt = (v) => num(v, 0)) {
  const neg = n < 0;
  const text = neg ? `(${fmt(Math.abs(n))})` : fmt(n);          // accounting style
  return new docx.Paragraph({ alignment: docx.AlignmentType.RIGHT,
    children: [new docx.TextRun({ text, color: neg ? "C00000" : "000000" })] });
}
```

### Totals / subtotal rows

Set the totals row apart with shading, a top border, and bold figures:

```js
const totalTop = { top: { style: docx.BorderStyle.SINGLE, size: 6, color: "000000" } };
const BR = (t) => new docx.TableCell({ borders: totalTop, shading: { fill: "F2F2F2" },
  children: [new docx.Paragraph({ alignment: docx.AlignmentType.RIGHT, children: [new docx.TextRun({ text: t, bold: true })] })] });
```

### RAG / status shading

Colour a status cell by value (breach = red, watch = amber, ok = green tint):

```js
const rag = { OK: "E6F4EA", WATCH: "FEF7E0", BREACH: "FCE8E6" };
const statusCell = (label) => new docx.TableCell({ shading: { fill: rag[label] || "FFFFFF" },
  children: [new docx.Paragraph({ alignment: docx.AlignmentType.CENTER, children: [new docx.TextRun(label)] })] });
```

### Recipe A — Portfolio holdings table with a totals row

Columns: Holding (left) · Sector (left) · Weight % (right) · Market Value (right). Shaded bold
header, thousands-separated values, and a shaded/bold totals row.

```js
function holdingsTable(rows /* [{name, sector, weight, value}] */, ccy = "USD") {
  const HB = (t, right) => new docx.TableCell({ shading: { fill: "1F3864" },
    children: [new docx.Paragraph({ alignment: right ? docx.AlignmentType.RIGHT : docx.AlignmentType.LEFT,
      children: [new docx.TextRun({ text: t, bold: true, color: "FFFFFF" })] })] });
  const totalW = rows.reduce((s, r) => s + r.weight, 0);
  const totalV = rows.reduce((s, r) => s + r.value, 0);
  const top = { top: { style: docx.BorderStyle.SINGLE, size: 6, color: "000000" } };
  return new docx.Table({
    width: { size: 100, type: docx.WidthType.PERCENTAGE },
    columnWidths: [3800, 2600, 1600, 2000], layout: docx.TableLayoutType.FIXED,
    margins: { top: 40, bottom: 40, left: 100, right: 100 },
    rows: [
      new docx.TableRow({ tableHeader: true, children: [ HB("Holding"), HB("Sector"), HB("Weight", true), HB("Market Value", true) ] }),
      ...rows.map((r) => new docx.TableRow({ children: [
        L(r.name), L(r.sector), R(pct(r.weight).replace("+", "")), R(money(r.value, ccy)),
      ]})),
      new docx.TableRow({ children: [
        new docx.TableCell({ borders: top, shading: { fill: "F2F2F2" }, columnSpan: 2,
          children: [new docx.Paragraph({ children: [new docx.TextRun({ text: "Total", bold: true })] })] }),
        new docx.TableCell({ borders: top, shading: { fill: "F2F2F2" },
          children: [new docx.Paragraph({ alignment: docx.AlignmentType.RIGHT, children: [new docx.TextRun({ text: `${totalW.toFixed(1)}%`, bold: true })] })] }),
        new docx.TableCell({ borders: top, shading: { fill: "F2F2F2" },
          children: [new docx.Paragraph({ alignment: docx.AlignmentType.RIGHT, children: [new docx.TextRun({ text: money(totalV, ccy), bold: true })] })] }),
      ]}),
    ],
  });
}
```

### Recipe B — Periodic returns table (fund vs benchmark)

A compact grid: row labels left, period columns right-aligned, negatives in red. Good for
factsheets and board packs.

```js
function returnsTable(periods /* ["1M","3M","YTD","1Y","3Y p.a."] */, fund, bench /* number[] */) {
  const H = (t, right) => new docx.TableCell({ shading: { fill: "EFEFEF" },
    children: [new docx.Paragraph({ alignment: right ? docx.AlignmentType.RIGHT : docx.AlignmentType.LEFT,
      children: [new docx.TextRun({ text: t, bold: true })] })] });
  const cell = (v) => new docx.TableCell({ children: [new docx.Paragraph({ alignment: docx.AlignmentType.RIGHT,
    children: [new docx.TextRun({ text: pct(v), color: v < 0 ? "C00000" : "000000" })] })] });
  const rowFor = (label, vals) => new docx.TableRow({ children: [ L(label), ...vals.map(cell) ] });
  return new docx.Table({
    width: { size: 100, type: docx.WidthType.PERCENTAGE },
    layout: docx.TableLayoutType.FIXED,
    rows: [
      new docx.TableRow({ tableHeader: true, children: [ H(""), ...periods.map((p) => H(p, true)) ] }),
      rowFor("Fund (net)", fund),
      rowFor("Benchmark", bench),
      rowFor("Excess", fund.map((v, i) => +(v - bench[i]).toFixed(1))),
    ],
  });
}
```

### Disclosures / footnotes

Regulatory pages need small-print disclosures. Use a smaller size and muted colour:

```js
new docx.Paragraph({
  spacing: { before: 200 },
  children: [new docx.TextRun({ size: 14, color: "666666",
    text: "Past performance is not a reliable indicator of future results. Figures are net of fees in the fund's base currency. Source: internal records." })],
});
```
