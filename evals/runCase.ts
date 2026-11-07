import { runAgent, type ModelClient } from "../src/agent/loop";
import { summarizeRange } from "../src/excel/context";
import { formatLiveContext } from "../src/excel/liveContext";
import type { EvalCase } from "./cases";
import { createFakeExecutor } from "./fakeExecutor";
import { scoreCase, type CaseScore, type RecordedCall } from "./score";

const MAX_STEPS = 8;

export interface CaseRun {
  readonly id: string;
  readonly score: CaseScore;
  readonly calls: readonly RecordedCall[];
  readonly reply: string;
  /** Set when the model call itself failed (network, quota, bad key). Counts as a failure. */
  readonly error?: string;
}

/** The same context text the add-in sends for this sheet. */
export const contextFor = (c: EvalCase): string =>
  formatLiveContext(summarizeRange({ sheetName: c.sheet.name, address: c.sheet.address, values: c.sheet.values }), c.workbook ?? null);

/** Runs one case through the real agent loop against `model`, with a fake workbook behind the tools. */
export const runCase = async (model: ModelClient, c: EvalCase): Promise<CaseRun> => {
  const fake = createFakeExecutor(c.sheet);
  try {
    const result = await runAgent({ model, execute: fake.execute, userMessage: c.prompt, context: contextFor(c), maxSteps: MAX_STEPS });
    return { id: c.id, score: scoreCase(c, fake.calls(), result.finalText), calls: fake.calls(), reply: result.finalText };
  } catch (err) {
    const error = err instanceof Error ? err.message : String(err);
    return { id: c.id, score: { pass: false, failures: [`model error: ${error}`], rejectedCalls: 0 }, calls: fake.calls(), reply: "", error };
  }
};

export interface Summary {
  readonly total: number;
  readonly passed: number;
  readonly rate: number;
  readonly failed: readonly CaseRun[];
}

export const summarize = (runs: readonly CaseRun[]): Summary => {
  const failed = runs.filter((r) => !r.score.pass);
  const passed = runs.length - failed.length;
  return { total: runs.length, passed, rate: runs.length === 0 ? 0 : passed / runs.length, failed };
};
