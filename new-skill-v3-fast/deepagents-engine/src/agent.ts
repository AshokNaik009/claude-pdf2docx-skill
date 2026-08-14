// Deep-agent wiring (spec §7.7): orchestrator deep agent + page_builder /
// gate_fixer subagents over the closure tools from makeTools. The orchestrator
// itself calls deterministic_pass (it is NOT pre-run here) and drives the
// build → merge → gate → fix loop end to end.
import { mkdirSync } from "node:fs";
import { basename, resolve } from "node:path";
// @ts-ignore vendored JS
import { workDirFor } from "../vendor/src/pipeline.js";
import { createDeepAgent, FilesystemBackend } from "deepagents";
import { modelForRole } from "./model-factory.js";
import { makeTools } from "./tools/page-io.js";
import { orchestratorPrompt } from "./prompts/orchestrator.js";
import { pageBuilderPrompt } from "./prompts/page-builder.js";
import { gateFixerPrompt } from "./prompts/gate-fixer.js";
import { CONFIG, VENDOR_SKILL_DIR, type EngineConfig } from "./config.js";
import type { Log } from "./tools/deterministic-pass.js";

export async function runAgent(pdfPath: string, cfg: EngineConfig = CONFIG, log: Log = console.log): Promise<void> {
  const absPdf = resolve(pdfPath);
  const workDir = workDirFor(absPdf, cfg);
  mkdirSync(workDir, { recursive: true });

  const outName = basename(cfg.outName || basename(absPdf)).replace(/\.(pdf|docx)$/i, "");
  const tools = makeTools(absPdf, workDir, cfg, log);

  const subagents = [
    {
      name: "page_builder",
      description:
        "Reconstructs ONE complex PDF page as compact layout JSON (with a single docx.js codegen fallback). Call with the page number, e.g. 'Build page 3.' Include prior errors when retrying.",
      systemPrompt: pageBuilderPrompt(cfg, VENDOR_SKILL_DIR),
      tools: tools.pageBuilderTools,
      model: modelForRole("page_builder"),
    },
    {
      name: "gate_fixer",
      description:
        "Repairs the specific pages the document gate flagged. Call with the failing checks and perPageHints verbatim; it fixes pages via the validators and reports (the orchestrator re-merges).",
      systemPrompt: gateFixerPrompt(cfg, VENDOR_SKILL_DIR),
      tools: tools.gateFixerTools,
      model: modelForRole("gate_fixer"),
    },
  ];

  const agent = createDeepAgent({
    model: modelForRole("orchestrator"),
    tools: tools.orchestratorTools,
    systemPrompt: orchestratorPrompt(cfg, outName),
    subagents,
    backend: new FilesystemBackend({ rootDir: workDir }),
  });

  log(`Agent run: ${absPdf} → outputs/${outName}.docx  (work: ${workDir})`);
  const result = await agent.invoke(
    {
      messages: [
        {
          role: "user",
          content: `Convert ${absPdf} to DOCX, output name "${outName}". Follow your procedure.`,
        },
      ],
    },
    { recursionLimit: 150 }
  );

  const final = result.messages?.[result.messages.length - 1];
  const content =
    typeof final?.content === "string"
      ? final.content
      : (final?.content ?? [])
          .map((c: any) => (typeof c === "string" ? c : c?.text ?? ""))
          .join("");
  log(`\n===== Orchestrator report =====\n${content || "(no final message)"}`);
}
