#!/usr/bin/env npx tsx
// deepagents-engine CLI (spec §7.8).
//
//   npx tsx cli.ts all <input.pdf> [--out NAME] [--llm auto|all|none] [--concurrency N]
//                      [--fresh] [--no-agent] [--force-vision] [--model-index role=N]
//   npx tsx cli.ts merge <workDir> --out NAME     # re-merge only
//   npx tsx cli.ts gate  <docxPath> [--work DIR]  # re-gate only
import "dotenv/config";
import { existsSync, rmSync } from "node:fs";
import { basename, join, resolve } from "node:path";
import { CONFIG, type EngineConfig } from "./src/config.js";
import { setModelIndex } from "./src/model-factory.js";
import { deterministicPass } from "./src/tools/deterministic-pass.js";
import { mergeDocument, runGateOn, sessionSummary, markPageImageFallback, type GateResult } from "./src/tools/merge-gate.js";
// @ts-ignore vendored JS
import { workDirFor } from "./vendor/src/pipeline.js";

function usage(): never {
  console.error("Usage:");
  console.error("  npx tsx cli.ts all <input.pdf> [--out NAME] [--llm auto|all|none] [--concurrency N]");
  console.error("                     [--fresh] [--no-agent] [--force-vision] [--model-index role=N]");
  console.error("  npx tsx cli.ts merge <workDir> --out NAME");
  console.error("  npx tsx cli.ts gate  <docxPath> [--work DIR]");
  process.exit(1);
}

function parseFlags(argv: string[]): { cfg: EngineConfig; positional: string[]; workOverride?: string } {
  const cfg: EngineConfig = { ...CONFIG };
  const positional: string[] = [];
  let workOverride: string | undefined;
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    const next = () => argv[++i];
    switch (a) {
      case "--out": cfg.outName = next(); break;
      case "--llm": {
        const v = next();
        if (!["auto", "all", "none"].includes(v)) { console.error(`--llm must be auto|all|none (got "${v}")`); process.exit(1); }
        cfg.llmMode = v as EngineConfig["llmMode"];
        break;
      }
      case "--concurrency": cfg.concurrency = Number(next()); break;
      case "--dpi": cfg.dpi = Number(next()); break;
      case "--fresh": cfg.fresh = true; break;
      case "--no-agent": cfg.noAgent = true; break;
      case "--force-vision": cfg.forceVision = true; break;
      case "--no-gate": cfg.gate = false; break;
      case "--work": workOverride = next(); break;
      case "--model-index": {
        const v = next();
        const m = v?.match(/^([\w-]+)=(\d+)$/);
        if (!m) { console.error(`--model-index expects role=N (got "${v}")`); process.exit(1); }
        setModelIndex(m[1], Number(m[2]));
        break;
      }
      default: positional.push(a);
    }
  }
  return { cfg, positional, workOverride };
}

function printGate(gate: GateResult) {
  console.log(`\nVerification gate (soffice-free): ${gate.ok ? "PASS" : "ISSUES"}`);
  for (const c of gate.checks) console.log(`  ${c.ok ? "✓" : "✗"} ${c.name}${c.detail ? ` — ${c.detail}` : ""}`);
  if (gate.counts) {
    console.log(
      `  doc: ${gate.counts.paragraphs} paragraph(s), ${gate.counts.tables} table(s), ` +
      `${gate.counts.words} words, ${gate.counts.media_files} image(s)`
    );
  }
  if (!gate.ok && gate.perPageHints.length) {
    console.log("  per-page hints:");
    for (const h of gate.perPageHints.slice(0, 12)) console.log(`    page ${h.page || "?"}: ${h.hint}`);
  }
}

async function doMergeAndGate(workDir: string, cfg: EngineConfig) {
  const outName = cfg.outName || basename(workDir);
  const merged = await mergeDocument(workDir, outName, cfg);
  if (merged.missingPages.length) {
    console.warn(
      `⚠ ${merged.missingPages.length} page(s) not completed and skipped: ` +
        merged.missingPages.map((p) => `${p.page}(${p.status})`).join(", ")
    );
  }
  console.log(`\nMerged ${merged.mergedPages} page(s).`);
  console.log(`  DOCX: ${merged.docxPath}`);
  if (cfg.gate) {
    try {
      printGate(runGateOn(merged.docxPath, workDir, cfg));
    } catch (e) {
      console.warn(`\nVerification gate skipped: ${(e as Error).message}`);
    }
  }
  return merged;
}

/** M1 deterministic spine: complex pages go straight to the image fallback. */
async function runNoAgent(pdfPath: string, cfg: EngineConfig) {
  const result = await deterministicPass(pdfPath, cfg, console.log);
  console.log(
    `Deterministic pass: ${result.deterministicDone.length}/${result.totalPages} done, ` +
      `${result.complexPages.length} complex, ${result.failed.length} failed`
  );
  for (const cp of result.complexPages) {
    markPageImageFallback(result.workDir, cp.page, `image fallback (--no-agent): ${cp.reason}`, cfg);
    console.log(`  ✓ page ${cp.page}  [image: --no-agent fallback]`);
  }
  for (const f of result.failed) {
    try {
      markPageImageFallback(result.workDir, f.page, `image fallback (--no-agent, page failed: ${f.error})`, cfg);
      console.log(`  ✓ page ${f.page}  [image: fallback after failure]`);
    } catch (e) {
      console.warn(`  ✗ page ${f.page} unrecoverable: ${(e as Error).message}`);
    }
  }
  await doMergeAndGate(result.workDir, cfg);
}

async function main() {
  const [, , cmdRaw, ...rest] = process.argv;
  const { cfg, positional, workOverride } = parseFlags(rest);

  let cmd = cmdRaw;
  let target = positional[0];
  if (cmdRaw && cmdRaw.toLowerCase().endsWith(".pdf")) { cmd = "all"; target = cmdRaw; }
  if (!cmd || !["all", "merge", "gate"].includes(cmd) || !target) usage();

  if (cmd === "merge") {
    const workDir = resolve(target);
    if (!existsSync(join(workDir, "session.json"))) {
      console.error(`No session.json in ${workDir}`);
      process.exit(1);
    }
    await doMergeAndGate(workDir, cfg);
    return;
  }

  if (cmd === "gate") {
    const docxPath = resolve(target);
    if (!existsSync(docxPath)) { console.error(`DOCX not found: ${docxPath}`); process.exit(1); }
    const stem = basename(docxPath).replace(/\.docx$/i, "");
    const workDir = workOverride ? resolve(workOverride) : join(cfg.workRoot, stem);
    if (!existsSync(join(workDir, "session.json"))) {
      console.error(`No session.json in ${workDir} — pass --work <workDir> (the gate needs the page manifests)`);
      process.exit(1);
    }
    printGate(runGateOn(docxPath, workDir, cfg));
    return;
  }

  // cmd === "all"
  const pdfPath = resolve(target);
  if (!existsSync(pdfPath)) { console.error(`Input PDF not found: ${pdfPath}`); process.exit(1); }
  if (cfg.fresh) {
    rmSync(workDirFor(pdfPath, cfg), { recursive: true, force: true });
    console.log("(--fresh) cleared prior session; reprocessing all pages");
  }

  if (cfg.noAgent) {
    await runNoAgent(pdfPath, cfg);
    return;
  }

  const { runAgent } = await import("./src/agent.js");
  await runAgent(pdfPath, cfg, console.log);
  const workDir = workDirFor(pdfPath, cfg);
  const summary = sessionSummary(workDir);
  const byMode = new Map<string, number>();
  for (const p of summary.pages) byMode.set(p.mode ?? p.status, (byMode.get(p.mode ?? p.status) ?? 0) + 1);
  console.log(
    `\nSession: ${summary.pages.filter((p) => p.status === "completed").length}/${summary.totalPages} completed ` +
      `(${[...byMode.entries()].map(([m, n]) => `${m}: ${n}`).join(", ")})`
  );
}

main().catch((e) => {
  console.error("Fatal:", (e as Error).stack || (e as Error).message);
  process.exit(1);
});
