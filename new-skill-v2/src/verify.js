// Zero-LLM structural verifier (VERIFY box in the deterministic-first pipeline).
// Ported from .claude/skills/docxjs/scripts/validate-page-module.mjs so the runtime
// gets the same checks that script already performs, without spawning a subprocess:
// empty table cells, ragged rows/columnSpan mismatches, undefined/nested children,
// full-width tables inside multi-column sections, non-CONTINUOUS later sections, and
// a real Packer.toBuffer pack test. Feed `errors` back into a retry prompt — models
// reliably fix their own structural mistake when they can see the message.
import * as docx from "docx";

/** @param {string} code  raw model stdout/extracted module source */
export function verifyModuleCode(code) {
  const errors = [];
  const warnings = [];
  if (code.includes("```")) {
    errors.push("source contains markdown fences — extraction left wrapper text in the code");
  }
  if (/new\s+docx\.ImageRun\s*\(/.test(code)) {
    warnings.push(
      "code constructs docx.ImageRun directly with hand-written dimensions — prefer ctx.image(file) so the host sizes it from the asset's real aspect ratio"
    );
  }
  return { ok: errors.length === 0, errors, warnings };
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
        warnings.push(
          `${at}: a table sits inside a multi-column section — full-width tables belong in their own CONTINUOUS single-column section`
        );
      }
      inspectTable(child, at, errors, warnings);
    }
  });
}

function inspectTable(table, path, errors, warnings) {
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
      warnings.push(
        `${path}: row ${ri} spans ${w} columns but the widest row spans ${max} — check rowSpan/columnSpan (covered cells must be omitted)`
      );
    }
  });
}

/**
 * Verify an already-executed page result before it's merged into the final document.
 * @param {{sections: object[], styles?: object}} result  the shape returned by executeModule
 * @returns {Promise<{ok:boolean, errors:string[], warnings:string[], bytes:number|null}>}
 */
export async function verifySections({ sections, styles }) {
  const errors = [];
  const warnings = [];
  let bytes = null;

  if (!Array.isArray(sections) || !sections.length) {
    errors.push("returned zero sections — the page would be blank");
    return { ok: false, errors, warnings, bytes };
  }

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

const WORD = /[\p{L}\p{N}]{3,}/gu;

function flattenBlockText(blocks) {
  const parts = [];
  const visit = (b) => {
    if (!b) return;
    if (b.text) parts.push(b.text);
    if (b.items) parts.push(...b.items);
    if (b.rows) for (const row of b.rows) for (const cell of row) if (cell?.text) parts.push(cell.text);
  };
  for (const b of blocks) visit(b);
  return parts.join(" ");
}

/**
 * Zero-LLM fidelity gate for the deterministic path: what fraction of the
 * ORIGINAL page's distinct words (>=3 chars, so punctuation/short connectors
 * don't dilute the signal) also appear somewhere in the reconstructed blocks.
 * A cheap bag-of-words containment check rather than a true diff/LCS — good
 * enough to catch "detection went wrong and dropped/garbled most of the
 * page" (a badly mis-clustered table, a column-band misread that swallowed
 * content) without the cost of proper alignment.
 * @param {string} originalText     the PDF page's plain text layer
 * @param {object[]} blocks         the reconstructed doc-blocks (heading/paragraph/list/table)
 * @returns {{ok:boolean, score:number, reasons:string[]}}
 */
export function coverage(originalText, blocks, cfg = {}) {
  const min = cfg.coverageMin ?? 0.6;
  const originalWords = new Set((originalText || "").toLowerCase().match(WORD) || []);
  if (!originalWords.size) return { ok: true, score: 1, reasons: [] };

  const reconstructed = flattenBlockText(blocks).toLowerCase();
  const reconstructedWords = new Set(reconstructed.match(WORD) || []);

  let hits = 0;
  for (const w of originalWords) if (reconstructedWords.has(w)) hits++;
  const score = hits / originalWords.size;

  const reasons = [];
  if (score < min) reasons.push(`only ${(score * 100).toFixed(0)}% of the page's words reconstructed (need >=${(min * 100).toFixed(0)}%)`);
  return { ok: score >= min, score, reasons };
}
