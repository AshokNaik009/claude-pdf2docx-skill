#!/usr/bin/env node
// PDF → structured model → editable DOCX.
//
//   node index.js run   <input.pdf>   # extract + per-page processing (writes work/<name>/)
//   node index.js merge <work/<name>> # order + merge + generate DOCX  (the "final merge")
//   node index.js all   <input.pdf>   # run then merge in one shot
//
// Flags (override config.js): --work DIR --out NAME --dpi N --concurrency N
//   --model M --retries N --llm auto|all|none
import { mkdirSync, writeFileSync, rmSync, existsSync } from "node:fs";
import { resolve, join, basename } from "node:path";
import { CONFIG } from "./config.js";
import { runPipeline, workDirFor } from "./src/pipeline.js";
import { buildDocumentModel, documentToMarkdown } from "./src/merge.js";
import { renderDocx } from "./src/docx.js";

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
      case "--fresh": cfg.fresh = true; break; // ignore any prior session, reprocess every page
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

  const mdPath = resolve(cfg.outputsDir, `${stem}.md`);
  writeFileSync(mdPath, documentToMarkdown(model));

  const buffer = await renderDocx(workDir, model, cfg);
  const docxPath = resolve(cfg.outputsDir, `${stem}.docx`);
  writeFileSync(docxPath, buffer);

  const stats = model.pages.length;
  console.log(`\nMerged ${stats} page(s).`);
  console.log(`  Markdown: ${mdPath}`);
  console.log(`  DOCX:     ${docxPath}`);
}

async function main() {
  const [, , cmdRaw, ...rest] = process.argv;
  const { cfg, positional } = parseFlags(rest);

  // Allow `node index.js <file.pdf>` as shorthand for `all`.
  let cmd = cmdRaw;
  let target = positional[0];
  if (cmdRaw && cmdRaw.toLowerCase().endsWith(".pdf")) { cmd = "all"; target = cmdRaw; }

  if (!cmd || !["run", "merge", "all"].includes(cmd) || !target) {
    console.error("Usage:");
    console.error("  node index.js run   <input.pdf>");
    console.error("  node index.js merge <work/<name>>");
    console.error("  node index.js all   <input.pdf>");
    process.exit(1);
  }

  if (cmd === "merge") {
    await doMerge(resolve(target), cfg);
    return;
  }

  const pdfPath = resolve(target);
  // Validate the input BEFORE any destructive --fresh cleanup, so a wrong path
  // never wipes a prior good run's work dir.
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
    console.log(`Run the final merge with:\n  node index.js merge ${workDir}`);
  }
}

main().catch((e) => {
  console.error("Fatal:", e.stack || e.message);
  process.exit(1);
});
