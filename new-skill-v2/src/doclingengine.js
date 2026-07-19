// --docling engine: whole-document path that bypasses the per-page vision-LLM
// loop entirely. Runs Docling (Python, in the project venv) to parse the PDF
// into doc.json, then generates per-page docx.js `buildPage` modules from it
// (src/doclingcode.js) so the NORMAL merge (src/docx.js) builds the .docx.
//
//   PDF ──▶ [venv python: scripts/docling_extract.py] ──▶ doc.json + images
//        ──▶ generateDoclingCode() ──▶ work/<name>/pages/*.mjs (+ session.json)
//        ──▶ node index.js merge ──▶ outputs/<name>.docx
import { spawn, execFileSync } from "node:child_process";
import { mkdirSync } from "node:fs";
import { join, resolve, basename } from "node:path";
import { workDirFor } from "./pipeline.js";
import { generateDoclingCode } from "./doclingcode.js";

/** True if the venv python can import docling (the parser is installed). */
export function doclingAvailable(cfg) {
  try {
    execFileSync(resolve(cfg.pythonBin), ["-c", "import docling"], { stdio: "ignore" });
    return true;
  } catch {
    return false;
  }
}

/** Run a python subprocess, streaming its output so model downloads are visible. */
function runPython(pythonBin, args) {
  return new Promise((res, rej) => {
    const p = spawn(pythonBin, args, { stdio: ["ignore", "inherit", "inherit"] });
    p.on("error", rej);
    p.on("close", (code) => (code === 0 ? res() : rej(new Error(`docling extractor exited with code ${code}`))));
  });
}

/**
 * Extract + code-generate for one PDF. Returns { workDir } shaped like
 * runPipeline so index.js can hand it straight to the merge step.
 */
export async function runDoclingEngine(pdfPath, cfg, log = console.log) {
  const workDir = workDirFor(pdfPath, cfg);
  const doclingOut = join(workDir, "_docling");
  mkdirSync(doclingOut, { recursive: true });

  log(`Docling: parsing ${basename(pdfPath)} → doc.json (first run downloads model weights)…`);
  await runPython(resolve(cfg.pythonBin), [resolve(cfg.doclingScript), pdfPath, doclingOut]);

  log("Docling: generating docx.js page modules from doc.json…");
  const { pages } = generateDoclingCode(doclingOut, workDir, basename(pdfPath));
  log(`Docling: wrote ${pages} page module(s) → ${workDir}`);
  return { workDir };
}
