/**
 * Validate a generated page module BEFORE merging it into the final document.
 *
 * Catches the failures that make batch page-reconstruction unreliable:
 *   - hard errors: SyntaxError, missing/incorrect export, wrong return shape, throw during
 *     buildPage, throw during Packer
 *   - structural warnings that pack "successfully" but render wrong: empty table cells, ragged
 *     row widths, undefined entries in children, full-width tables inside a multi-column section
 *
 * Feed `result.errors` back into a retry prompt — models reliably fix their own syntax error when
 * they can see the message.
 *
 * CLI:   node validate-page-module.mjs ./page-003.mjs [./assets]
 * API:   import { validatePageModule } from "./validate-page-module.mjs";
 *        const r = await validatePageModule(codeString, { imageDir });
 *        if (!r.ok) retry(r.errors);
 */
import * as docx from "docx";
import { readFile } from "node:fs/promises";
import { existsSync, readFileSync } from "node:fs";
import { join, extname } from "node:path";
import { makeImageCtx } from "./image-fit.mjs";

const IMG_TYPES = { ".png": "png", ".jpg": "jpg", ".jpeg": "jpg", ".gif": "gif", ".bmp": "bmp" };
// 1x1 transparent PNG, used when an asset is missing so validation still exercises the layout.
const STUB_PNG = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
  "base64"
);

function makeCtx(imageDir, warnings) {
  // Use the real host-side fitter so validation exercises actual sizing; fall back to a stub
  // image only when the asset is genuinely absent.
  const fitted = makeImageCtx(docx, { imageDir: imageDir || ".", onWarn: (m) => warnings.push(m) });
  return {
    widths: fitted.widths,
    image(file, opts = {}) {
      const p = imageDir ? join(imageDir, file) : file;
      if (existsSync(p)) return fitted.image(file, opts);
      warnings.push(`asset not found, used stub: ${file}`);
      const type = IMG_TYPES[extname(file).toLowerCase()] || "png";
      return new docx.ImageRun({ type, data: STUB_PNG, transformation: { width: 400, height: 260 } });
    },
  };
}

function inspectChildren(children, path, errors, warnings, inMultiColumn) {
  if (!Array.isArray(children)) {
    errors.push(`${path}.children must be an array`);
    return;
  }
  children.forEach((child, i) => {
    const at = `${path}.children[${i}]`;
    if (child === undefined || child === null) {
      errors.push(`${at} is ${child} — undefined entries are silently dropped from the document`);
      return;
    }
    if (Array.isArray(child)) {
      errors.push(`${at} is a nested array — flatten it with spread (…) instead`);
      return;
    }
    if (child instanceof docx.Table) {
      if (inMultiColumn) {
        warnings.push(`${at}: a table sits inside a multi-column section — full-width tables belong in their own CONTINUOUS single-column section`);
      }
      inspectTable(child, at, errors, warnings);
    }
  });
}

function inspectTable(table, path, errors, warnings) {
  // docx keeps rows on the instance; guard in case internals differ across versions.
  const rows = table.root?.filter?.((n) => n instanceof docx.TableRow) ?? [];
  if (!rows.length) return;
  const widths = rows.map((row) => {
    const cells = row.root?.filter?.((n) => n instanceof docx.TableCell) ?? [];
    cells.forEach((cell, ci) => {
      const kids = cell.root?.filter?.((n) => n instanceof docx.Paragraph || n instanceof docx.Table) ?? [];
      if (!kids.length) warnings.push(`${path}: a cell (col ${ci}) has no Paragraph — it will render empty`);
    });
    return cells.reduce((s, c) => s + (c.options?.columnSpan || 1), 0);
  });
  const max = Math.max(...widths);
  widths.forEach((w, ri) => {
    if (w !== max) {
      warnings.push(`${path}: row ${ri} spans ${w} columns but the widest row spans ${max} — check rowSpan/columnSpan (covered cells must be omitted)`);
    }
  });
}

/**
 * @param {string} code       ES module source exporting buildPage(docx, ctx)
 * @param {object} [opts]
 * @param {string} [opts.imageDir]    directory holding extracted page assets
 * @param {string} [opts.exportName]  default "buildPage"
 * @returns {Promise<{ok:boolean, errors:string[], warnings:string[], bytes:number|null}>}
 */
export async function validatePageModule(code, opts = {}) {
  const { imageDir, exportName = "buildPage" } = opts;
  const errors = [];
  const warnings = [];
  let bytes = null;

  if (code.includes("```")) errors.push("source contains markdown fences — extraction left wrapper text in the code");
  if (/new\s+docx\.ImageRun\s*\(/.test(code)) {
    warnings.push("code constructs docx.ImageRun directly with hand-written dimensions — prefer ctx.image(file) so the host sizes it from the asset's real aspect ratio");
  }

  let mod;
  try {
    const url = "data:text/javascript;base64," + Buffer.from(code, "utf8").toString("base64");
    mod = await import(url);
  } catch (e) {
    errors.push(`${e.constructor?.name || "Error"} while parsing/loading module: ${e.message}`);
    return { ok: false, errors, warnings, bytes };
  }

  const build = mod[exportName];
  if (typeof build !== "function") {
    errors.push(`module does not export a function named \`${exportName}\``);
    return { ok: false, errors, warnings, bytes };
  }

  let out;
  try {
    out = build(docx, makeCtx(imageDir, warnings));
  } catch (e) {
    errors.push(`${exportName}() threw: ${e.message}`);
    return { ok: false, errors, warnings, bytes };
  }

  // Accept { sections, styles } | single section | bare children array.
  let sections;
  let styles;
  if (Array.isArray(out)) sections = [{ children: out }];
  else if (out && Array.isArray(out.sections)) { sections = out.sections; styles = out.styles; }
  else if (out && Array.isArray(out.children)) sections = [out];
  else {
    errors.push("return value must be { sections: [...] } (or a single section, or a children array)");
    return { ok: false, errors, warnings, bytes };
  }
  if (!sections.length) errors.push("returned zero sections — the page would be blank");

  sections.forEach((s, i) => {
    const path = `sections[${i}]`;
    if (!s || typeof s !== "object") { errors.push(`${path} is not an object`); return; }
    const count = s.properties?.column?.count;
    inspectChildren(s.children, path, errors, warnings, typeof count === "number" && count > 1);
    if (i > 0 && s.properties && s.properties.type !== docx.SectionType.CONTINUOUS) {
      warnings.push(`${path} is not CONTINUOUS — it will start a new page`);
    }
  });

  if (errors.length) return { ok: false, errors, warnings, bytes };

  try {
    const doc = new docx.Document({ ...(styles ? { styles } : {}), sections });
    const buf = await docx.Packer.toBuffer(doc);
    bytes = buf.length;
    if (bytes < 2000) warnings.push(`packed document is only ${bytes} bytes — likely near-empty`);
  } catch (e) {
    errors.push(`Packer failed: ${e.message}`);
  }

  return { ok: errors.length === 0, errors, warnings, bytes };
}

// ---- CLI -------------------------------------------------------------------
if (import.meta.url === `file://${process.argv[1]}`) {
  const [file, imageDir] = process.argv.slice(2);
  if (!file) {
    console.error("usage: node validate-page-module.mjs <module.mjs> [imageDir]");
    process.exit(2);
  }
  const code = await readFile(file, "utf8");
  const r = await validatePageModule(code, { imageDir });
  for (const w of r.warnings) console.warn("  warn: " + w);
  for (const e of r.errors) console.error("  ERROR: " + e);
  console.log(r.ok ? `OK — packed ${r.bytes} bytes` : "FAILED");
  process.exit(r.ok ? 0 : 1);
}
