// M0 provider smoke test (spec §6 M0.4): for each role in models.json,
// instantiate candidate #0 AND candidate #1 and invoke a trivial prompt.
//   npx tsx src/smoke.ts            all roles, candidates 0 and 1
//   npx tsx src/smoke.ts --fake     also prove a fake provider entry resolves through the factory (M4c)
import "dotenv/config";
import { loadRoles, candidatesForRole, modelForRole } from "./model-factory.js";

async function tryCandidate(role: string, idx: number): Promise<void> {
  const list = candidatesForRole(role);
  if (idx >= list.length) return;
  const c = list[idx];
  const label = `${role} → ${c.provider}:${c.model}`;
  try {
    const model = modelForRole(role, idx);
    const res = await model.invoke("Reply with the single word: ok");
    const text = typeof res.content === "string" ? res.content.trim().slice(0, 40) : JSON.stringify(res.content).slice(0, 40);
    console.log(`  ok   ${label}  ("${text}")`);
  } catch (e) {
    console.log(`  FAIL ${label}  — ${(e as Error).message.slice(0, 160)}`);
  }
}

async function main() {
  const roles = Object.keys(loadRoles());
  console.log("candidate #0 (primary):");
  for (const role of roles) await tryCandidate(role, 0);
  console.log("candidate #1 (fallback):");
  for (const role of roles) await tryCandidate(role, 1);

  if (process.argv.includes("--fake")) {
    // M4(c): a made-up provider entry must fail with the factory's clear error,
    // proving resolution goes through the factory rather than silently defaulting.
    const rolesMap = loadRoles() as Record<string, unknown[]>;
    rolesMap["__fake_role"] = [{ provider: "bedrock", model: "anthropic.claude-opus-fake", params: { region: "us-east-1" } }];
    const { modelForRoleAsync } = await import("./model-factory.js");
    try {
      await modelForRoleAsync("__fake_role");
      console.log("  FAIL fake bedrock entry unexpectedly instantiated");
    } catch (e) {
      console.log(`  ok   fake bedrock entry resolved through factory and reported: ${(e as Error).message.slice(0, 120)}`);
    }
  }
}

main().catch((e) => { console.error(e); process.exit(1); });
