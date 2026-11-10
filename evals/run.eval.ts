/**
 * Live eval: sends every case to the real Gemini API and scores which tools the model calls.
 * Costs a small amount of API quota, so it is NOT part of `npm test`. Run it with `npm run eval`.
 *
 * Env:
 *   GEMINI_API_KEY   required (the suite is skipped without it)
 *   EVAL_MODEL       model id (default: the add-in's default)
 *   EVAL_ONLY        comma-separated case ids to run, e.g. "freeze-header,autofit"
 *   EVAL_MIN_PASS    pass rate needed for the run to succeed, 0-1 (default 0.8; LLM output varies)
 */
import { describe, expect, it } from "vitest";
import { createGeminiClient } from "../src/agent/geminiClient";
import { DEFAULT_MODEL_ID } from "../src/agent/models";
import { SYSTEM_PROMPT } from "../src/agent/prompt";
import { CASES } from "./cases";
import { runCase, summarize, type CaseRun } from "./runCase";

const apiKey = process.env.GEMINI_API_KEY;
const modelId = process.env.EVAL_MODEL || DEFAULT_MODEL_ID;
const only = new Set((process.env.EVAL_ONLY ?? "").split(",").map((s) => s.trim()).filter(Boolean));
const minPass = Number(process.env.EVAL_MIN_PASS ?? "0.8");

const out = (line: string) => process.stdout.write(`${line}\n`);

const describeRun = (r: CaseRun): string => {
  const tools = r.calls.map((c) => (c.ok ? c.name : `${c.name}(rejected)`)).join(", ") || "(no tools)";
  const status = r.score.pass ? "PASS" : "FAIL";
  const why = r.score.failures.length > 0 ? `\n      - ${r.score.failures.join("\n      - ")}` : "";
  return `  ${status}  ${r.id.padEnd(22)} ${tools}${why}`;
};

describe.skipIf(!apiKey)("live eval", () => {
  it(`meets the pass-rate bar (${minPass}) on ${modelId}`, async () => {
    const model = createGeminiClient({ apiKeys: [apiKey as string], model: modelId, systemInstruction: SYSTEM_PROMPT });
    const selected = CASES.filter((c) => only.size === 0 || only.has(c.id));
    expect(selected.length, "EVAL_ONLY matched no cases").toBeGreaterThan(0);

    const runs: CaseRun[] = [];
    for (const c of selected) {
      const run = await runCase(model, c);
      runs.push(run);
      out(describeRun(run));
    }

    const s = summarize(runs);
    const rejected = runs.reduce((n, r) => n + r.score.rejectedCalls, 0);
    out(`\n  ${s.passed}/${s.total} passed (${Math.round(s.rate * 100)}%) on ${modelId}; ${rejected} tool call(s) rejected by validation\n`);
    expect(s.rate).toBeGreaterThanOrEqual(minPass);
  });
});
