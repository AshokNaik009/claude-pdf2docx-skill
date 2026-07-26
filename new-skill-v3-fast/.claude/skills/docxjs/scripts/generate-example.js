/*
 * Worked docx.js example: an academic excerpt on heat transfer during baking.
 * Demonstrates, on one page: two-column justified body with italic "et al." citations,
 * a bold sub-heading, a continuous section break to a full-width captioned editorial
 * table (bold repeating header, rule under header, top-aligned cells, a multi-paragraph
 * "(i)/(ii)" Conduction cell), a centered figure placeholder + caption, then a continuous
 * break back to two columns.
 *
 * Run in Node:   node scripts/generate-example.js   ->   writes example.docx
 * Use in browser: delete the Node block at the bottom, keep generateDocument(), and call
 *                 docx.Packer.toBlob(generateDocument()).then(b => saveAs(b, "example.docx"));
 */

// In the browser, `docx` is a global from the UMD <script>; in Node, require it:
const docx = (typeof window !== "undefined" && window.docx) ? window.docx : require("docx");

// ---- helpers ---------------------------------------------------------------
function runsFromText(text) {
  return text.split(/(et al\.)/g)
    .filter((s) => s !== "")
    .map((s) => s === "et al."
      ? new docx.TextRun({ text: s, italics: true })
      : new docx.TextRun(s));
}
const bodyPara = (text) => new docx.Paragraph({
  alignment: docx.AlignmentType.JUSTIFIED,
  spacing: { after: 180, line: 288 },
  children: runsFromText(text),
});
const heading = (text) => new docx.Paragraph({
  spacing: { before: 200, after: 120 },
  children: [new docx.TextRun({ text, bold: true })],
});
const caption = (label, rest) => new docx.Paragraph({
  alignment: docx.AlignmentType.CENTER,
  spacing: { before: 120, after: 120 },
  children: [new docx.TextRun({ text: label, bold: true }), new docx.TextRun(rest)],
});

// ---- table -----------------------------------------------------------------
const P = (t) => new docx.Paragraph({ children: [new docx.TextRun(t)], spacing: { after: 40 } });
const B = (t) => new docx.Paragraph({ children: [new docx.TextRun({ text: t, bold: true })] });
const NONE = { style: docx.BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorder = { top: NONE, bottom: NONE, left: NONE, right: NONE };
const headRule = { ...noBorder, bottom: { style: docx.BorderStyle.SINGLE, size: 4, color: "000000" } };

function heatTransferTable() {
  const dataRow = (mode, source, role) => new docx.TableRow({
    children: [
      new docx.TableCell({ borders: noBorder, verticalAlign: docx.VerticalAlign.TOP, children: [P(mode)] }),
      new docx.TableCell({ borders: noBorder, children: source }),
      new docx.TableCell({ borders: noBorder, children: [P(role)] }),
    ],
  });
  return new docx.Table({
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
      dataRow("Radiation",
        [P("Emission of thermal energy through oven walls, heating elements, and surrounding hot surfaces.")],
        "Heat is transferred by thermal radiation from hot oven walls and heating elements to the product surface. This mechanism dominates at high temperatures and contributes significantly to rapid surface heating and crust formation."),
      dataRow("Convection",
        [P("Circulation of hot, humid air within the oven chamber.")],
        "Heat transfer occurs through the movement of hot air around the product. Regulates heat exchange at the product surface, controls moisture evaporation, and ensures uniform temperature distribution for uniform baking."),
      dataRow("Conduction",
        [
          new docx.Paragraph({ children: [new docx.TextRun("(i)\tDirect contact between the product and the carrier medium (baking tray, mould, or conveyor band)")], spacing: { after: 80 } }),
          new docx.Paragraph({ children: [new docx.TextRun("(ii)\tHeat transfer within the product matrix from the surface toward the core")] }),
        ],
        "Heat is transferred through direct contact between the product and heated carrier surfaces and further conducted internally within the product"),
    ],
  });
}

// ---- document --------------------------------------------------------------
function generateDocument() {
  const MARG = { top: 1440, bottom: 1440, left: 1440, right: 1440 };
  return new docx.Document({
    styles: { default: { document: { run: { font: "Palatino Linotype", size: 22 } } } },
    sections: [
      // 1) two-column body
      {
        properties: { page: { margin: MARG }, column: { count: 2, space: 560, separate: false } },
        children: [
          bodyPara("heat, thermal conductivity, thermal diffusivity and moisture diffusivity, influence temperature and moisture simulation of the product during baking (Chakraborty and Dash, 2023)."),
          bodyPara("Using heat in food preparation allowed humans to unlock helpful nutrients and kill harmful bacteria. As soon as product kept in the oven moisture evaporates and baking started. An oven used for baking is a heated enclosure that cooks food evenly. Industrial baking oven conditions are usually generated using the three modes of heat transfer: radiation, convection and conduction. During baking, heat is transferred mainly by convection from the heating media, and by radiation from oven walls to the product surface followed by conduction to the geometric center (Sablani et al. 2002). Table 1 shows the various mechanisms for the modes of heat transfer. Baking is governed by simultaneous heat and mass transfer processes occurring within the product and between the product and the oven environment. Heat is transferred to the product through multiple mechanisms acting concurrently. Fig. 2 shows the Mechanism of heat transfer during baking process inside the baking oven."),
        ],
      },
      // 2) full-width, single column: caption + table + figure + figure caption
      {
        properties: { type: docx.SectionType.CONTINUOUS },
        children: [
          caption("Table 1: ", "Mechanism of heat transfer"),
          heatTransferTable(),
          new docx.Paragraph({
            alignment: docx.AlignmentType.CENTER,
            spacing: { before: 240, after: 60 },
            border: { top: { style: docx.BorderStyle.SINGLE, size: 4, color: "000000" },
                      bottom: { style: docx.BorderStyle.SINGLE, size: 4, color: "000000" },
                      left: { style: docx.BorderStyle.SINGLE, size: 4, color: "000000" },
                      right: { style: docx.BorderStyle.SINGLE, size: 4, color: "000000" } },
            children: [new docx.TextRun({ text: "[ Fig. 2 diagram: Radiation / Hot Air convection / Heat Conduction / Baking Tray — insert image with docx.ImageRun ]", italics: true, color: "888888" })],
          }),
          caption("Fig. 2: ", "Mechanism of heat transfer during baking process inside the baking oven"),
        ],
      },
      // 3) back to two columns
      {
        properties: { type: docx.SectionType.CONTINUOUS, column: { count: 2, space: 560, separate: false } },
        children: [
          bodyPara("The relative contribution of each heat transfer mechanism depends on the product characteristics and oven configuration. These mechanisms can be used to regulate moisture loss and achieve the desired quality attributes in the final baked product (Marcotte, 2007)."),
        ],
      },
    ],
  });
}

// ---- Node output (delete this block for browser use) -----------------------
if (typeof module !== "undefined" && require.main === module) {
  const fs = require("fs");
  docx.Packer.toBuffer(generateDocument()).then((buf) => {
    fs.writeFileSync("example.docx", buf);
    console.log("Wrote example.docx (" + buf.length + " bytes)");
  });
}

if (typeof module !== "undefined") module.exports = { generateDocument };
