# Images and figures in docx.js

Read this whenever a page contains a figure, diagram, chart, logo, or screenshot. Image handling is
where generated page code most often goes subtly wrong, because **every image mistake packs
successfully** — a squashed diagram, a figure orphaned from its caption, and a figure jammed into a
narrow column all produce a valid `.docx` that simply looks wrong.

## Contents
[What the model decides vs what the host decides](#split) · [Sizing](#sizing) ·
[Placement decision table](#placement) · [Captions](#captions) · [Missing assets](#missing) ·
[Inline vs floating](#floating) · [Recipes](#recipes)

<a id="split"></a>
## 1. Split the work: the host sizes, the model places

The single biggest reliability win. `ImageRun` demands explicit pixel `width` **and** `height`, and
a model looking at a rendered page cannot know an asset's intrinsic dimensions — so any number it
writes is a guess, and a wrong ratio is invisible (a 1600×900 diagram emitted as `600×50` packs
fine and renders as a smear).

So **the generated code should never contain image dimensions at all**. The host reads each asset's
real size from its file header and scales it; the model only says *which* image goes *where*:

```js
// In generated code — no numbers, no aspect-ratio guessing:
new docx.Paragraph({ alignment: docx.AlignmentType.CENTER, children: [ ctx.image("fig-2.png") ] })

// Fit to a narrower band when the figure sits inside a multi-column section:
ctx.image("fig-2.png", { fit: "column" })
```

`scripts/image-fit.mjs` implements the host side: it parses PNG/JPEG/GIF/BMP headers with no
dependencies, scales to the available width, caps height so a tall figure can't push off the page,
and never upscales a small image. Verified against packed output: a 1600×900 asset places at
624×351 px (full content width), a 400×1200 asset at 207×620 (height-capped), and a 120×80 asset
stays 120×80 — ratios preserved in every case.

If the host does **not** provide `ctx.image`, and you must pass dimensions yourself, ask for the
intrinsic size rather than inventing one; failing that, set the width to the available width and
derive the height from the aspect ratio you can measure on the page image.

<a id="sizing"></a>
## 2. Sizing arithmetic (when you must do it yourself)

docx.js "pixels" are 96 dpi; Word measures in twips at 1440/inch, so `px = twips / 1440 * 96`.

| Band | Twips | Pixels |
|---|---|---|
| Letter content width (8.5" − 1" margins) | 9360 | **624** |
| A4 content width (210 mm − 1" margins) | 9026 | **602** |
| One of 2 columns (Letter, 560 gutter) | 4400 | **293** |
| Half content width | 4680 | 312 |

Rules that keep figures from breaking a page:

- Width must not exceed the band it sits in, or Word clips or reflows it.
- Cap height around **620 px** (Letter, 1" margins) so a portrait figure plus its caption still fits.
- **Never upscale.** A 120 px-wide logo stretched to 624 px looks terrible; leave small images native.
- Preserve the ratio: `height = round(intrinsicHeight × width / intrinsicWidth)`.

<a id="placement"></a>
## 3. Placement decision table

Decide from what you see on the page, then place accordingly:

| What the page shows | How to place it |
|---|---|
| Figure spans the **full page width**, text is two-column above/below | Its own section with `type: docx.SectionType.CONTINUOUS` and **no** `column` property. Never put a full-width figure inside a two-column section. |
| Figure sits **inside one column**, text flows above and below it in that column | Keep it in the multi-column section, `ctx.image(f, { fit: "column" })`. |
| Figure at the **top of the page**, body text follows | First child of the first section. |
| Figure with **text wrapping around it** | Inline is usually the faithful-enough choice; use floating only if wrapping is unmistakable (see below). |
| Two small figures **side by side** | A 2-column borderless table, one image per cell, `fit: "half"`. |
| Logo/mark in the **header** | Put the `ImageRun` in a `Header` paragraph, not the body. |

The full-width case is by far the most common in academic and report pages, and it is the one that
breaks silently: a wide diagram placed inside a two-column section is squeezed to ~293 px and
becomes unreadable. **When in doubt about a wide figure, give it its own `CONTINUOUS` section.**

<a id="captions"></a>
## 4. Captions

A caption is a separate centered paragraph immediately after (or before) the image paragraph, with
the label bold and the rest regular:

```js
new docx.Paragraph({
  alignment: docx.AlignmentType.CENTER,
  spacing: { before: 120, after: 120 },
  children: [
    new docx.TextRun({ text: "Fig. 2: ", bold: true }),
    new docx.TextRun("Mechanism of heat transfer during baking process inside the baking oven"),
  ],
});
```

Add `keepNext: true` to the **image** paragraph so Word never splits a figure from its caption
across a page break. Same for a table's caption paragraph when the caption sits above the table.
Transcribe caption text verbatim, including the label and its punctuation ("Fig. 2:", "Table 1:").

<a id="missing"></a>
## 5. When the asset is missing

If the page shows a figure but no corresponding file was extracted (common for vector diagrams that
the extractor could not rasterize), emit **one centered placeholder paragraph** naming it:

```js
new docx.Paragraph({
  alignment: docx.AlignmentType.CENTER,
  children: [ new docx.TextRun({ text: "[ Figure: Mechanism of heat transfer during baking ]", italics: true, color: "888888" }) ],
});
```

Then still emit the real caption paragraph beneath it.

**Do not fabricate the diagram.** Never rebuild it as a table of labels, a bulleted list of the
words visible in it, or ASCII art. A placeholder is honest and easy to fix later; a fabricated
table silently corrupts the document and is very hard to spot in a batch. Likewise, never invent a
filename that was not in the asset list — `ctx.image` on a nonexistent file yields a broken figure.

<a id="floating"></a>
## 6. Inline vs floating

Inline (the default) is right almost always: the image behaves like a big character in its
paragraph, and centering the paragraph centers the figure.

Floating with text wrap is available via `floating` on the `ImageRun`, but it needs absolute
positioning and is easy to get wrong; prefer inline unless the page unmistakably wraps text around
a figure.

```js
new docx.ImageRun({
  type: "png", data, transformation: { width: 200, height: 150 },
  floating: {
    horizontalPosition: { relative: docx.HorizontalPositionRelativeFrom.COLUMN, offset: 0 },
    verticalPosition: { relative: docx.VerticalPositionRelativeFrom.PARAGRAPH, offset: 0 },
    wrap: { type: docx.TextWrappingType.SQUARE, side: docx.TextWrappingSide.BOTH_SIDES },
  },
});
```

Offsets are EMUs (914400 per inch), not twips or pixels — a frequent source of images landing off
the page.

<a id="recipes"></a>
## 7. Recipes

**Full-width figure between two-column text** (the standard academic/report page):

```js
return { sections: [
  { properties: { page: { margin: M }, column: { count: 2, space: 560 } },
    children: [ bodyPara("...text above..."), ] },
  { properties: { type: docx.SectionType.CONTINUOUS },      // full width, no column property
    children: [
      new docx.Paragraph({ alignment: docx.AlignmentType.CENTER, keepNext: true,
        children: [ ctx.image("fig-2.png") ] }),
      caption("Fig. 2: ", "Mechanism of heat transfer during baking process inside the baking oven"),
    ] },
  { properties: { type: docx.SectionType.CONTINUOUS, column: { count: 2, space: 560 } },
    children: [ bodyPara("...text below...") ] },
] };
```

**Two figures side by side** (borderless table, so they stay aligned):

```js
const NONE = { style: docx.BorderStyle.NONE, size: 0, color: "FFFFFF" };
const bare = { top: NONE, bottom: NONE, left: NONE, right: NONE };
const imgCell = (file, cap) => new docx.TableCell({ borders: bare, children: [
  new docx.Paragraph({ alignment: docx.AlignmentType.CENTER, keepNext: true, children: [ ctx.image(file, { fit: "half" }) ] }),
  new docx.Paragraph({ alignment: docx.AlignmentType.CENTER, children: [ new docx.TextRun({ text: cap, size: 18 }) ] }),
]});
new docx.Table({ width: { size: 100, type: docx.WidthType.PERCENTAGE },
  rows: [ new docx.TableRow({ children: [ imgCell("fig-3a.png", "(a) before"), imgCell("fig-3b.png", "(b) after") ] }) ] });
```

**Logo in the header:**

```js
headers: { default: new docx.Header({ children: [
  new docx.Paragraph({ alignment: docx.AlignmentType.RIGHT, children: [ ctx.image("logo.png", { widthPx: 90 }) ] }),
]}) }
```
