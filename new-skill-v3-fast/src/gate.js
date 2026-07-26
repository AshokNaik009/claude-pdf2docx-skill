// SOFFICE-FREE VERIFICATION GATE. After the merge, inspect the OUTPUT .docx
// deterministically with extractandVerify.py (pure stdlib — no LibreOffice/Word)
// and confirm the facts we INJECTED actually landed:
//   - the package is valid OOXML,
//   - every injected fill hex appears in the document's shading palette,
//   - the reconstructed text still covers the page's original words.
// This is the gate the HANDOFF calls for — a check that catches "the fill/frame
// silently didn't render" or "half the page's words are missing" without ever
// launching a renderer. It REPORTS mismatches (drives targeted fixes); it does
// not itself re-run pages.
import { execFileSync } from "node:child_process";
import { readFileSync, rmSync } from "node:fs";
import { join } from "node:path";

const normHex = (h) => String(h || "").replace(/^#/, "").toUpperCase();
const WORD = /[\p{L}\p{N}]{3,}/gu;
const words = (s) => new Set((String(s || "").toLowerCase().match(WORD) || []));

/**
 * Gather what SHOULD be true of the output from the page manifests: the set of
 * injected fill colors and the union of the pages' original words.
 * @returns {{fills:Set<string>, words:Set<string>}}
 */
export function expectedFromModel(model) {
  const fills = new Set();
  const wordSet = new Set();
  let images = 0;
  for (const pg of model.pages) {
    for (const f of pg.injected?.fills || []) if (f.fill) fills.add(normHex(f.fill));
    if (pg.injected?.frame?.color) fills.add(normHex(pg.injected.frame.color));
    for (const w of words((pg.fallbackText || []).join(" "))) wordSet.add(w);
    images += (pg.assets || []).length;
  }
  return { fills, words: wordSet, images };
}

/** Run extractandVerify.py on the docx and return its parsed JSON report. */
export function inspectDocx(docxPath, workDir, cfg) {
  const reportPath = join(workDir, "verify-report.json");
  try {
    execFileSync(cfg.gateBin || "python3", [cfg.gateScript || "extractandVerify.py", docxPath, "--json", reportPath],
      { stdio: ["ignore", "ignore", "pipe"], maxBuffer: 64 * 1024 * 1024 });
  } catch (e) {
    // --strict is not set, so a non-zero exit means the interpreter/script
    // itself failed (missing python, bad path) — surface that, don't swallow it.
    throw new Error(`extractandVerify.py failed: ${e.stderr ? e.stderr.toString().slice(0, 300) : e.message}`);
  }
  const report = JSON.parse(readFileSync(reportPath, "utf8"));
  rmSync(reportPath, { force: true });
  return report;
}

/**
 * Verify the output .docx against the injected facts. Returns a structured
 * result the caller can print; `ok` is true when every check passes.
 */
export function runGate(docxPath, model, cfg) {
  const expected = expectedFromModel(model);
  const report = inspectDocx(docxPath, workDir(docxPath, cfg), cfg);
  const checks = [];

  checks.push({ name: "package valid", ok: !!report.validation?.ok,
    detail: report.validation?.ok ? "" : (report.validation?.issues || []).map((i) => i.msg).slice(0, 3).join("; ") });

  // Injected fills must appear in the document palette.
  const palette = new Set(Object.keys(report.shading?.palette_counts || {}).map(normHex));
  const missingFills = [...expected.fills].filter((f) => !palette.has(f));
  checks.push({ name: "injected fills present",
    ok: missingFills.length === 0,
    detail: missingFills.length ? `missing from output: ${missingFills.map((f) => "#" + f).join(", ")}` : `${expected.fills.size} fill(s) confirmed` });

  // Geometry: no table may be wider than its container, or it clips off the
  // right edge (the "cropped right border" class of bug). Checks top-level
  // tables against the section content width and nested tables against the
  // width of the cell that holds them.
  const geo = report.sections?.[0]?.geometry || {};
  const contentPt = (geo.page_w_pt || 612) - (geo.margins_pt?.left || 72) - (geo.margins_pt?.right || 72);
  const overflows = [];
  const tableWidthPt = (t) => (t.grid_pt || []).reduce((s, w) => s + w, 0);
  const checkTable = (t, containerPt, path) => {
    const w = tableWidthPt(t);
    if (containerPt && w > containerPt * 1.03) {
      overflows.push(`${path} is ${w.toFixed(0)}pt wide but its container is ${containerPt.toFixed(0)}pt`);
    }
    (t.rows || []).forEach((row, ri) => row.forEach((cell, ci) => {
      const cw = cell.width_pt || (t.grid_pt || [])[ci];
      (cell.tables || []).forEach((nt, ni) => checkTable(nt, cw, `${path}>r${ri}c${ci}>t${ni}`));
    }));
  };
  (report.blocks || []).forEach((b, bi) => { if (b.type === "table") checkTable(b, contentPt, `table[${bi}]`); });
  checks.push({ name: "no table exceeds its width", ok: overflows.length === 0,
    detail: overflows.length ? overflows.slice(0, 3).join("; ") : "all tables fit their container" });

  // Word coverage: how many of the original words survived into the output.
  // Text lives in paragraph blocks AND in (possibly nested) table cells.
  const outWords = new Set();
  const addCells = (tbl) => {
    for (const row of tbl.rows || []) for (const cell of row) {
      for (const w of words(cell.text)) outWords.add(w);
      for (const nt of cell.tables || []) addCells(nt);
    }
  };
  for (const b of report.blocks || []) {
    if (b.type === "table") addCells(b);
    else for (const w of words(b.text)) outWords.add(w);
  }
  let hits = 0;
  for (const w of expected.words) if (outWords.has(w)) hits++;
  const cov = expected.words.size ? hits / expected.words.size : 1;
  checks.push({ name: "word coverage", ok: cov >= (cfg.coverageMin ?? 0.6),
    detail: `${(cov * 100).toFixed(0)}% of original words present (need ≥${((cfg.coverageMin ?? 0.6) * 100).toFixed(0)}%)` });

  // Images: every extracted figure/photo must be embedded in the output.
  const mediaOut = report.counts?.media_files ?? 0;
  checks.push({ name: "images embedded", ok: mediaOut >= expected.images,
    detail: `${mediaOut}/${expected.images} extracted image(s) embedded` });

  return {
    ok: checks.every((c) => c.ok),
    checks,
    counts: report.counts,
    palette: report.shading?.distinct_colors || [],
  };
}

// docxPath lives in outputsDir; the work dir (for the temp report) is derived
// from the caller — but we only need a writable dir, so reuse the docx's folder.
function workDir(docxPath, _cfg) {
  return join(docxPath, "..");
}
