// Engine defaults — port of ../config.js (spec §7.2) plus the agent-layer
// additions. Every value is overridable from cli.ts flags.
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = join(dirname(fileURLToPath(import.meta.url)), "..");

export interface EngineConfig {
  workRoot: string;
  outputsDir: string;
  dpi: number;

  // Quality check (which pages go to the LLM)
  minCharsForText: number;
  columnGapPts: number;
  llmMode: "auto" | "all" | "none";

  concurrency: number;
  retries: number;
  coverageMin: number;
  headingRatio: number;
  pageBreakBetween: boolean;
  compactJson: boolean;

  // Deterministic Python extraction + gate
  pyExtract: boolean;
  pyExtractBin: string;
  pyExtractScript: string;
  gate: boolean;
  gateBin: string;
  gateScript: string;

  // Agent layer (spec §7.2 additions)
  maxBuilderAttempts: number; // in-context fix retries inside validate loop (was verifyRetries)
  maxParallelPages: number;   // parallel page_builder task() dispatches
  maxGateFixLoops: number;    // gate_fixer → re-merge → re-gate loops
  forceVision: boolean;       // --force-vision: route one complex page through vision_reconstruct
  noAgent: boolean;           // --no-agent: M1 deterministic spine only
  fresh: boolean;

  outName?: string;
  // Kept for session.json compatibility with the vendored initSession (it
  // records cfg.model); the deep agent resolves real models per role instead.
  model: string;
}

export const CONFIG: EngineConfig = {
  workRoot: join(ROOT, "work"),
  outputsDir: join(ROOT, "outputs"),
  dpi: 200,

  minCharsForText: 90,
  columnGapPts: 40,
  llmMode: "auto",

  concurrency: 4,
  retries: 1,
  coverageMin: 0.6,
  headingRatio: 1.15,
  pageBreakBetween: true,
  compactJson: true,

  pyExtract: true,
  pyExtractBin: "python3",
  pyExtractScript: join(ROOT, "vendor", "pdfextract.py"),
  gate: true,
  gateBin: "python3",
  gateScript: join(ROOT, "vendor", "extractandVerify.py"),

  maxBuilderAttempts: 2,
  maxParallelPages: 2,
  maxGateFixLoops: 2,
  forceVision: false,
  noAgent: false,
  fresh: false,

  model: "deepagents",
};

export const VENDOR_SKILL_DIR = join(ROOT, "vendor", "skills", "docxjs");
export const ROOT_DIR = ROOT;
