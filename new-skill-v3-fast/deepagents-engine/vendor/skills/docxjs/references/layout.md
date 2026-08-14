# docx.js layout — deep reference

Contents: [Sections](#sections) · [Columns & mixing counts](#columns) · [Column/page breaks](#breaks) ·
[Page size, orientation, margins](#page) · [Headers & footers](#headers) · [Page numbers](#pagenumbers) ·
[Numbered/bulleted lists](#lists) · [Tab stops](#tabs) · [Line & paragraph spacing](#spacing) ·
[Table of contents](#toc) · [Units cheat-sheet](#units)

<a id="sections"></a>
## Sections

Each object in the document's `sections` array is a Word *section*. A section owns its page setup,
columns, and headers/footers. New sections start a new page **unless** you set a continuous type:

```js
{ properties: { type: docx.SectionType.CONTINUOUS }, children: [ ... ] }
```

`SectionType`: `CONTINUOUS` (no page break — flows on the same page), `NEXT_PAGE` (default),
`EVEN_PAGE`, `ODD_PAGE`.

Put page size/margins and document `styles` on the first section; continuation sections inherit page
setup but you must re-declare `column` on each one that needs columns.

<a id="columns"></a>
## Columns & mixing column counts on one page

```js
properties: { column: { count: 2, space: 560, separate: false } }
// count: number of columns; space: gutter in twips; separate: draw a vertical rule between columns
```

Unequal columns:

```js
column: { count: 2, separate: false, equalWidth: false,
  children: [ new docx.Column({ width: 3600, space: 300 }), new docx.Column({ width: 5400 }) ] }
```

**Two columns → full-width table/figure → two columns** (the canonical academic pattern) is three
continuous sections:

```js
sections: [
  { properties: { page: { margin: MARG }, column: { count: 2, space: 560 } },
    children: [ /* intro body */ ] },
  { properties: { type: docx.SectionType.CONTINUOUS },
    children: [ captionParagraph, table, imageParagraph, figCaption ] },   // single column, full width
  { properties: { type: docx.SectionType.CONTINUOUS, column: { count: 2, space: 560 } },
    children: [ /* remaining body */ ] },
]
```

<a id="breaks"></a>
## Column and page breaks

- **Column break** (jump to the next column without a new section):

  ```js
  new docx.Paragraph({ children: [ new docx.TextRun({ break: 1 }) ] });
  // or the dedicated break:
  new docx.Paragraph({ children: [ new docx.ColumnBreak() ] });
  ```

- **Page break:**

  ```js
  new docx.Paragraph({ children: [ new docx.PageBreak() ] });
  // or end a run's paragraph then start fresh; PageBreak is cleanest.
  ```

- **Keep a heading with the next paragraph:** `keepNext: true` on the heading paragraph.
  **Keep a paragraph's lines together:** `keepLines: true`.

<a id="page"></a>
## Page size, orientation, margins

```js
properties: {
  page: {
    size: { width: 12240, height: 15840, orientation: docx.PageOrientation.PORTRAIT }, // Letter, twips
    margin: { top: 1440, bottom: 1440, left: 1440, right: 1440, header: 720, footer: 720, gutter: 0 },
  },
}
```

A4 is `{ width: 11906, height: 16838 }`. For landscape set
`orientation: docx.PageOrientation.LANDSCAPE` (docx.js swaps width/height for you).

<a id="headers"></a>
## Headers & footers

```js
{
  headers: {
    default: new docx.Header({ children: [ /* paragraphs */ ] }),
    first:   new docx.Header({ children: [ /* page 1 only */ ] }),
    even:    new docx.Header({ children: [ /* even pages */ ] }),
  },
  footers: { default: new docx.Footer({ children: [ /* ... */ ] }) },
  properties: { titlePage: true },   // required for `first` header to apply
}
```

To use distinct even/odd headers, also set `evenAndOddHeaderAndFooters: true` in the Document
options:

```js
new docx.Document({ evenAndOddHeaderAndFooters: true, sections: [...] });
```

<a id="pagenumbers"></a>
## Page numbers

```js
new docx.Paragraph({ children: [
  new docx.TextRun("Page "),
  new docx.TextRun({ children: [ docx.PageNumber.CURRENT ] }),
  new docx.TextRun(" of "),
  new docx.TextRun({ children: [ docx.PageNumber.TOTAL_PAGES ] }),
]});
```

Start or restart numbering per section, and choose a separator:

```js
properties: { page: { pageNumbers: { start: 1, separator: docx.PageNumberSeparator.EM_DASH } } }
```

<a id="lists"></a>
## Numbered & bulleted lists

Declare numbering config on the Document, then reference it from paragraphs:

```js
new docx.Document({
  numbering: {
    config: [{
      reference: "my-list",
      levels: [
        { level: 0, format: docx.LevelFormat.DECIMAL, text: "%1.", alignment: docx.AlignmentType.START,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
        { level: 1, format: docx.LevelFormat.LOWER_ROMAN, text: "(%2)", alignment: docx.AlignmentType.START,
          style: { paragraph: { indent: { left: 1440, hanging: 360 } } } },
      ],
    }],
  },
  sections: [{ children: [
    new docx.Paragraph({ numbering: { reference: "my-list", level: 0 }, children: [new docx.TextRun("First")] }),
    new docx.Paragraph({ numbering: { reference: "my-list", level: 1 }, children: [new docx.TextRun("Sub")] }),
  ]}],
});
```

For bullets use `new docx.Paragraph({ bullet: { level: 0 }, children: [...] })`.

<a id="tabs"></a>
## Tab stops

```js
new docx.Paragraph({
  tabStops: [ { type: docx.TabStopType.RIGHT, position: docx.TabStopPosition.MAX } ],
  children: [ new docx.TextRun("Left side"), new docx.TextRun("\tRight side") ],
});
```

`TabStopType`: `LEFT | CENTER | RIGHT | DECIMAL`. `position` is twips (or `TabStopPosition.MAX` for
the right margin) — handy for header lines with a title on the left and page number on the right.

<a id="spacing"></a>
## Line & paragraph spacing

```js
spacing: {
  before: 200,          // twips before the paragraph
  after: 180,           // twips after
  line: 288,            // 240ths of a line: 240 single, 360 1.5x, 480 double
  lineRule: docx.LineRuleType.AUTO,   // AUTO (multiple), EXACT/AT_LEAST use line in twips
}
```

First-line indent: `indent: { firstLine: 360 }`. Hanging indent: `indent: { left: 720, hanging: 360 }`.

<a id="toc"></a>
## Table of contents

```js
new docx.TableOfContents("Contents", { hyperlink: true, headingStyleRange: "1-3" });
```

Word shows a "right-click → Update Field" placeholder until the reader updates it; the entries come
from paragraphs styled with heading styles (`heading: docx.HeadingLevel.HEADING_1`, etc.).

<a id="units"></a>
## Units cheat-sheet

| Concept | Unit | Conversion |
|---|---|---|
| Font `size` | half-points | pt × 2 |
| `spacing.before/after`, margins, indents, column `space`, tab `position` | twips | inch × 1440, pt × 20 |
| `spacing.line` (AUTO) | 240ths of a line | 240 = single |
| Border `size` | eighths of a point | pt × 8 |
| Table/cell width `DXA` | twips | inch × 1440 |
| Table/cell width `PERCENTAGE` | percent number | e.g. 100 |
| `ImageRun.transformation` width/height | pixels | display px |
| Page size | twips | Letter 12240×15840, A4 11906×16838 |
