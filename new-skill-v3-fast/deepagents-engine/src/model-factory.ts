// Provider adapters + role → model resolution (spec §7.4).
// models.json maps each role to an ORDERED candidate list; swapping providers
// (OpenRouter/Groq today, Bedrock/OpenAI later) is a models.json edit only.
import { readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import type { BaseChatModel } from "@langchain/core/language_models/chat_models";
import { ChatOpenAI } from "@langchain/openai";
import { ChatGroq } from "@langchain/groq";

export type Provider = "openrouter" | "groq" | "openai" | "anthropic" | "bedrock";

export interface ModelCandidate {
  provider: Provider;
  model: string;
  params?: Record<string, unknown>;
}

const HERE = dirname(fileURLToPath(import.meta.url));
const MODELS_PATH = join(HERE, "..", "models.json");

// @langchain/core's default failed-attempt handler misclassifies Groq's
// TRANSIENT per-minute TPM 429s as permanent quota exhaustion (their message
// matches /billing/i — "Upgrade to Dev Tier … /billing") and throws without a
// single retry, killing the whole agent run. Free-tier TPM windows clear in
// <60s, so: retry every 429 with the caller's exponential backoff; only
// genuinely fatal statuses / aborts propagate.
const STATUS_NO_RETRY = [400, 401, 402, 403, 404, 405, 406, 407, 409];
function retryTransientRateLimits(error: any): void {
  const msg: unknown = error?.message;
  if (typeof msg === "string" && (msg.startsWith("Cancel") || msg.startsWith("AbortError"))) throw error;
  if (error?.name === "AbortError" || error?.code === "ECONNABORTED") throw error;
  const code = error?.code ?? error?.error?.code ?? error?.error?.error?.code;
  // Groq returns 400 `tool_use_failed` when the model emits malformed tool-call
  // JSON. Generations are stochastic — a resend usually parses. Retry it.
  if (code === "tool_use_failed" || (typeof msg === "string" && msg.includes("tool_use_failed"))) return;
  const status = error?.status ?? error?.statusCode ?? error?.response?.status;
  if (status != null && STATUS_NO_RETRY.includes(+status)) throw error;
  // anything else (429s, 5xx, network) → return void = retry with backoff
}
const RETRY_OPTS = { maxRetries: 8, onFailedAttempt: retryTransientRateLimits };

let _roles: Record<string, ModelCandidate[]> | null = null;
export function loadRoles(): Record<string, ModelCandidate[]> {
  if (!_roles) _roles = JSON.parse(readFileSync(MODELS_PATH, "utf8"));
  return _roles!;
}

// CLI `--model-index role=N` overrides which candidate a role starts from.
const _indexOverride = new Map<string, number>();
export function setModelIndex(role: string, index: number) {
  _indexOverride.set(role, index);
}

export function candidatesForRole(role: string): ModelCandidate[] {
  const roles = loadRoles();
  const list = roles[role];
  if (!list || !list.length) throw new Error(`models.json has no candidates for role "${role}"`);
  return list;
}

function instantiate(c: ModelCandidate): BaseChatModel {
  const params = c.params ?? {};
  switch (c.provider) {
    case "openrouter": {
      if (!process.env.OPENROUTER_API_KEY) throw new Error("OPENROUTER_API_KEY is not set (.env)");
      return new ChatOpenAI({
        model: c.model,
        apiKey: process.env.OPENROUTER_API_KEY,
        configuration: { baseURL: "https://openrouter.ai/api/v1" },
        ...RETRY_OPTS,
        ...params,
      });
    }
    case "groq": {
      if (!process.env.GROQ_API_KEY) throw new Error("GROQ_API_KEY is not set (.env)");
      return new ChatGroq({ model: c.model, ...RETRY_OPTS, ...params });
    }
    case "openai": {
      if (!process.env.OPENAI_API_KEY) throw new Error("OPENAI_API_KEY is not set (.env)");
      return new ChatOpenAI({ model: c.model, ...RETRY_OPTS, ...params });
    }
    case "bedrock":
      throw new BedrockNeedsImport(c);
    case "anthropic":
      throw new Error(
        'provider "anthropic" needs `npm i @langchain/anthropic` and a factory entry — not installed in this project'
      );
    default:
      throw new Error(`unknown provider "${(c as ModelCandidate).provider}" in models.json`);
  }
}

// Bedrock is a documented future swap (spec §9/§10): dynamic import so the
// package is only required when models.json actually points at bedrock:*.
class BedrockNeedsImport extends Error {
  constructor(public candidate: ModelCandidate) {
    super("bedrock candidate requires async instantiation");
  }
}

async function instantiateBedrock(c: ModelCandidate): Promise<BaseChatModel> {
  let mod: any;
  try {
    mod = await import("@langchain/aws" as string);
  } catch {
    throw new Error(
      'provider "bedrock" requires `npm i @langchain/aws` plus AWS_ACCESS_KEY_ID / AWS_SECRET_ACCESS_KEY / AWS_REGION in the env (spec §10). Package is not installed.'
    );
  }
  const { region, ...rest } = (c.params ?? {}) as Record<string, unknown>;
  return new mod.ChatBedrockConverse({
    model: c.model,
    ...(region ? { region } : {}),
    ...RETRY_OPTS,
    ...rest,
  });
}

/** Synchronous resolution (non-bedrock providers). */
export function modelForRole(role: string, candidateIndex = 0): BaseChatModel {
  const list = candidatesForRole(role);
  const base = _indexOverride.get(role) ?? 0;
  const idx = Math.min(base + candidateIndex, list.length - 1);
  const c = list[idx];
  try {
    return instantiate(c);
  } catch (e) {
    if (e instanceof BedrockNeedsImport) {
      throw new Error(
        `role "${role}" resolves to a bedrock model — use modelForRoleAsync() for bedrock candidates`
      );
    }
    throw e;
  }
}

/**
 * Async resolution with construction-level fallback: walk the candidate list,
 * skipping any whose instantiation throws (missing key / missing package),
 * optionally probing the first working one with a trivial invoke.
 */
export async function modelForRoleAsync(
  role: string,
  opts: { probe?: boolean; log?: (msg: string) => void } = {}
): Promise<{ model: BaseChatModel; candidate: ModelCandidate; index: number }> {
  const list = candidatesForRole(role);
  const start = _indexOverride.get(role) ?? 0;
  const log = opts.log ?? (() => {});
  let lastErr: unknown;
  for (let i = start; i < list.length; i++) {
    const c = list[i];
    try {
      const model = c.provider === "bedrock" ? await instantiateBedrock(c) : instantiate(c);
      if (opts.probe) await model.invoke("Reply with the single word: ok");
      return { model, candidate: c, index: i };
    } catch (e) {
      lastErr = e;
      log(`  role ${role}: candidate ${i} (${c.provider}:${c.model}) failed — ${(e as Error).message}`);
    }
  }
  throw new Error(
    `no working model for role "${role}": ${(lastErr as Error | undefined)?.message ?? "no candidates"}`
  );
}
