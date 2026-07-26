#!/usr/bin/env node
// deterministic-fast PDF → editable DOCX.
//
//   node cli.js run   <input.pdf>   # extract + per-page processing (writes work/<name>/)
//   node cli.js merge <work/<name>> # order + merge + generate DOCX + verify gate
//   node cli.js all   <input.pdf>   # run then merge in one shot
//   node cli.js <input.pdf>         # shorthand for `all`
//
// Flags (override config.js):
//   --work DIR --out NAME --dpi N --concurrency N --model M --retries N
//   --llm auto|all|none        which pages go to the model
//   --no-compact-json          disable the compact-JSON fast path (use docx.js code-gen)
//   --no-py-extract            use the Node mupdf binding instead of pdfextract.py
//   --no-gate                  skip the extractandVerify.py output check
//   --fresh                    ignore any prior session and reprocess every page
import { mkdirSync, writeFileSync, rmSync, existsSync } from "node:fs";
import { resolve, basename } from "node:path";
import { CONFIG } from "./config.js";
import { runPipeline, workDirFor } from "./src/pipeline.js";
import { buildDocumentModel } from "./src/merge.js";
import { renderDocx } from "./src/docx.js";
import { runGate } from "./src/gate.js";

function parseFlags(argv) {
  const cfg = { ...CONFIG };
  const positional = [];
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    const next = () => argv[++i];
    switch (a) {
      case "--work": cfg.workRoot = next(); break;
      case "--out": cfg.outName = next(); break;
      case "--dpi": cfg.dpi = Number(next()); break;
      case "--concurrency": cfg.concurrency = Number(next()); break;
      case "--model": cfg.model = next(); break;
      case "--retries": cfg.retries = Number(next()); break;
      case "--llm": cfg.llmMode = next(); break;
      case "--no-compact-json": cfg.compactJson = false; break;
      case "--no-py-extract": cfg.pyExtract = false; break;
      case "--no-gate": cfg.gate = false; break;
      case "--fresh": cfg.fresh = true; break;
      default: positional.push(a);
    }
  }
  return { cfg, positional };
}

async function doMerge(workDir, cfg) {
  const model = buildDocumentModel(workDir);
  if (model.missing.length) {
    console.warn(
      `⚠ ${model.missing.length} page(s) not completed and will be skipped: ` +
        model.missing.map((p) => `${p.page}(${p.status})`).join(", ")
    );
  }
  mkdirSync(resolve(cfg.outputsDir), { recursive: true });
  const stem = cfg.outName || basename(workDir);

  const buffer = await renderDocx(workDir, model, cfg);
  const docxPath = resolve(cfg.outputsDir, `${stem}.docx`);
  writeFileSync(docxPath, buffer);

  console.log(`\nMerged ${model.pages.length} page(s).`);
  console.log(`  DOCX: ${docxPath}`);

  if (cfg.gate) {
    try {
      const gate = runGate(docxPath, model, cfg);
      console.log(`\nVerification gate (soffice-free): ${gate.ok ? "PASS" : "ISSUES"}`);
      for (const c of gate.checks) console.log(`  ${c.ok ? "✓" : "✗"} ${c.name}${c.detail ? ` — ${c.detail}` : ""}`);
      if (gate.counts) console.log(`  doc: ${gate.counts.paragraphs} paragraph(s), ${gate.counts.tables} table(s), ${gate.counts.words} words, ${gate.counts.media_files} image(s)`);
    } catch (e) {
      console.warn(`\nVerification gate skipped: ${e.message}`);
    }
  }
}

async function main() {
  const [, , cmdRaw, ...rest] = process.argv;
  const { cfg, positional } = parseFlags(rest);

  let cmd = cmdRaw;
  let target = positional[0];
  if (cmdRaw && cmdRaw.toLowerCase().endsWith(".pdf")) { cmd = "all"; target = cmdRaw; }

  if (!cmd || !["run", "merge", "all"].includes(cmd) || !target) {
    console.error("Usage:");
    console.error("  node cli.js run   <input.pdf>");
    console.error("  node cli.js merge <work/<name>>");
    console.error("  node cli.js all   <input.pdf>");
    process.exit(1);
  }

  if (cmd === "merge") {
    await doMerge(resolve(target), cfg);
    return;
  }

  const pdfPath = resolve(target);
  if (!existsSync(pdfPath)) {
    console.error(`Input PDF not found: ${pdfPath}`);
    process.exit(1);
  }
  if (cfg.fresh) {
    rmSync(workDirFor(pdfPath, cfg), { recursive: true, force: true });
    console.log("(--fresh) cleared prior session; reprocessing all pages");
  }
  const { workDir } = await runPipeline(pdfPath, cfg);

  if (cmd === "all") {
    await doMerge(workDir, cfg);
  } else {
    console.log(`\nExtraction complete → ${workDir}`);
    console.log(`Run the final merge with:\n  node cli.js merge ${workDir}`);
  }
}

main().catch((e) => {
  console.error("Fatal:", e.stack || e.message);
  process.exit(1);
});
