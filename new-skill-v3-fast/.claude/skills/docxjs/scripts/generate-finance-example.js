/*
 * Worked docx.js example: a one-page fund factsheet / portfolio appendix.
 * Demonstrates finance-flavoured constructs: a bold section heading, a two-column commentary
 * block, a right-aligned periodic-returns table (fund vs benchmark, negatives in red), a
 * full-width holdings table with thousands separators and a shaded/bold totals row, and a
 * small-print disclosures footnote.
 *
 * Run in Node:    node scripts/generate-finance-example.js   ->   writes factsheet.docx
 * Use in browser: keep generateDocument(), then
 *                 docx.Packer.toBlob(generateDocument()).then(b => saveAs(b, "factsheet.docx"));
 */

const docx = (typeof window !== "undefined" && window.docx) ? window.docx : require("docx");

// ---- number formatting (do it in JS, align in the cell) --------------------
const money = (n, ccy = "USD", dp = 0) =>
  new Intl.NumberFormat("en-US", { style: "currency", currency: ccy, minimumFractionDigits: dp, maximumFractionDigits: dp }).format(n);
const pct = (n, dp = 1) => `${n >= 0 ? "+" : ""}${n.toFixed(dp)}%`;

// ---- cell + paragraph helpers ----------------------------------------------
const L = (t, opts = {}) => new docx.TableCell({ ...opts, children: [new docx.Paragraph({ children: [new docx.TextRun(t)] })] });
const R = (t, opts = {}) => new docx.TableCell({ ...opts, children: [new docx.Paragraph({ alignment: docx.AlignmentType.RIGHT, children: [new docx.TextRun(t)] })] });
const heading = (t) => new docx.Paragraph({ spacing: { before: 120, after: 120 }, children: [new docx.TextRun({ text: t, bold: true, size: 26 })] });
const body = (t) => new docx.Paragraph({ alignment: docx.AlignmentType.JUSTIFIED, spacing: { after: 140, line: 276 }, children: [new docx.TextRun(t)] });

// ---- periodic returns table ------------------------------------------------
function returnsTable(periods, fund, bench) {
  const H = (t, right) => new docx.TableCell({ shading: { fill: "EFEFEF" },
    children: [new docx.Paragraph({ alignment: right ? docx.AlignmentType.RIGHT : docx.AlignmentType.LEFT, children: [new docx.TextRun({ text: t, bold: true })] })] });
  const cell = (v) => new docx.TableCell({ children: [new docx.Paragraph({ alignment: docx.AlignmentType.RIGHT,
    children: [new docx.TextRun({ text: pct(v), color: v < 0 ? "C00000" : "000000" })] })] });
  const rowFor = (label, vals) => new docx.TableRow({ children: [L(label), ...vals.map(cell)] });
  return new docx.Table({
    width: { size: 100, type: docx.WidthType.PERCENTAGE }, layout: docx.TableLayoutType.FIXED,
    margins: { top: 30, bottom: 30, left: 80, right: 80 },
    rows: [
      new docx.TableRow({ tableHeader: true, children: [H(""), ...periods.map((p) => H(p, true))] }),
      rowFor("Fund (net)", fund),
      rowFor("Benchmark", bench),
      rowFor("Excess", fund.map((v, i) => +(v - bench[i]).toFixed(1))),
    ],
  });
}

// ---- holdings table with totals row ----------------------------------------
function holdingsTable(rows, ccy = "USD") {
  const HB = (t, right) => new docx.TableCell({ shading: { fill: "1F3864" },
    children: [new docx.Paragraph({ alignment: right ? docx.AlignmentType.RIGHT : docx.AlignmentType.LEFT, children: [new docx.TextRun({ text: t, bold: true, color: "FFFFFF" })] })] });
  const top = { top: { style: docx.BorderStyle.SINGLE, size: 6, color: "000000" } };
  const totalW = rows.reduce((s, r) => s + r.weight, 0);
  const totalV = rows.reduce((s, r) => s + r.value, 0);
  const totalCell = (t) => new docx.TableCell({ borders: top, shading: { fill: "F2F2F2" },
    children: [new docx.Paragraph({ alignment: docx.AlignmentType.RIGHT, children: [new docx.TextRun({ text: t, bold: true })] })] });
  return new docx.Table({
    width: { size: 100, type: docx.WidthType.PERCENTAGE },
    columnWidths: [3800, 2600, 1600, 2000], layout: docx.TableLayoutType.FIXED,
    margins: { top: 40, bottom: 40, left: 100, right: 100 },
    rows: [
      new docx.TableRow({ tableHeader: true, children: [HB("Holding"), HB("Sector"), HB("Weight", true), HB("Market Value", true)] }),
      ...rows.map((r) => new docx.TableRow({ children: [L(r.name), L(r.sector), R(`${r.weight.toFixed(1)}%`), R(money(r.value, ccy))] })),
      new docx.TableRow({ children: [
        new docx.TableCell({ borders: top, shading: { fill: "F2F2F2" }, columnSpan: 2,
          children: [new docx.Paragraph({ children: [new docx.TextRun({ text: "Total", bold: true })] })] }),
        totalCell(`${totalW.toFixed(1)}%`),
        totalCell(money(totalV, ccy)),
      ]}),
    ],
  });
}

// ---- document --------------------------------------------------------------
function generateDocument() {
  const MARG = { top: 1440, bottom: 1440, left: 1440, right: 1440 };
  const holdings = [
    { name: "Aurora Semiconductor", sector: "Information Technology", weight: 6.4, value: 3210000 },
    { name: "Meridian Health Group", sector: "Health Care", weight: 5.1, value: 2560000 },
    { name: "Northwind Energy", sector: "Energy", weight: 4.8, value: 2410000 },
    { name: "Coastal Financial", sector: "Financials", weight: 4.2, value: 2100000 },
    { name: "Vertex Industrials", sector: "Industrials", weight: 3.9, value: 1950000 },
  ];
  return new docx.Document({
    styles: { default: { document: { run: { font: "Calibri", size: 22 } } } },
    sections: [
      // Title + two-column manager commentary
      {
        properties: { page: { margin: MARG } },
        children: [
          new docx.Paragraph({ children: [new docx.TextRun({ text: "Global Equity Fund — Monthly Factsheet", bold: true, size: 32 })] }),
          new docx.Paragraph({ spacing: { after: 200 }, children: [new docx.TextRun({ text: "Share class I (Acc) · USD · Data as at 30 June 2026", color: "666666" })] }),
        ],
      },
      {
        properties: { type: docx.SectionType.CONTINUOUS, column: { count: 2, space: 560 } },
        children: [
          heading("Manager Commentary"),
          body("The Fund advanced over the month, supported by holdings in information technology and health care, while energy detracted modestly. Positioning remains tilted toward quality companies with durable cash flows and pricing power, consistent with the strategy's long-term mandate."),
          body("We added to two industrials names on weakness and trimmed a financials position after strong performance. Cash was held near the lower end of the policy range. Risk metrics remain within limits, and the portfolio's active share is unchanged versus the prior month."),
        ],
      },
      // Full-width performance + holdings
      {
        properties: { type: docx.SectionType.CONTINUOUS },
        children: [
          heading("Performance (%)"),
          returnsTable(["1M", "3M", "YTD", "1Y", "3Y p.a."], [2.1, 5.4, 9.8, 14.2, 11.5], [1.8, 4.9, 8.6, 12.7, 10.4]),
          heading("Top Holdings"),
          holdingsTable(holdings, "USD"),
          new docx.Paragraph({ spacing: { before: 200 }, children: [new docx.TextRun({ size: 14, color: "666666",
            text: "Past performance is not a reliable indicator of future results. Returns are net of fees in the fund's base currency (USD). Holdings are subject to change. Source: internal records." })] }),
        ],
      },
    ],
  });
}

// ---- Node output (delete for browser use) ----------------------------------
if (typeof module !== "undefined" && require.main === module) {
  const fs = require("fs");
  docx.Packer.toBuffer(generateDocument()).then((buf) => {
    fs.writeFileSync("factsheet.docx", buf);
    console.log("Wrote factsheet.docx (" + buf.length + " bytes)");
  });
}
if (typeof module !== "undefined") module.exports = { generateDocument };
