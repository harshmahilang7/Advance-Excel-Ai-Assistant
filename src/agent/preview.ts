import { friendlyError } from "../lib/friendlyError";
import type { ToolExecutor, ToolResult } from "./loop";
import { parseToolCall } from "./tools";

/** Tools that only read the workbook. In preview mode they still run, so the model can look around. */
const READ_ONLY_TOOLS: ReadonlySet<string> = new Set(["get_workbook_info", "read_range", "audit_formulas", "propose_plan"]);

export const isReadOnlyTool = (name: string): boolean => READ_ONLY_TOOLS.has(name);

export interface PendingCall {
  readonly name: string;
  readonly args: unknown;
  /** Short human description shown in the preview list. */
  readonly summary: string;
}

export interface AppliedStep {
  readonly ok: boolean;
  readonly text: string;
}

export interface ApplyOutcome {
  readonly results: readonly AppliedStep[];
  /** Calls not run because an earlier one failed. */
  readonly skipped: number;
}

const targetOf = (args: unknown): string => {
  const a = (args ?? {}) as { address?: unknown; sourceAddress?: unknown };
  const where = a.address ?? a.sourceAddress;
  return typeof where === "string" ? where : "";
};

/** One-line description of a tool call, e.g. "write_cells: 3 row(s) at D1". */
export const describeCall = (name: string, args: unknown): string => {
  const where = targetOf(args);
  if (name === "write_cells") {
    const rows = (args as { data?: unknown[] } | null)?.data?.length ?? 0;
    return `write_cells: ${rows} row(s) at ${where}`;
  }
  return where ? `${name}: ${where}` : name;
};

/**
 * Wraps the real executor for preview mode: reads run for real, valid writes are queued instead of
 * executed. Invalid calls come back as errors so the model can correct them, exactly as in a live run.
 */
export const createPreviewExecutor =
  (real: ToolExecutor, onQueue: (call: PendingCall) => void): ToolExecutor =>
  async (name, args) => {
    if (READ_ONLY_TOOLS.has(name)) return real(name, args);
    const parsed = parseToolCall(name, args);
    if (!parsed.ok) return { ok: false, error: parsed.error };
    const summary = describeCall(parsed.call.name, parsed.call.args);
    onQueue({ name: parsed.call.name, args: parsed.call.args, summary });
    return { ok: true, summary: `Queued for approval: ${summary}`, data: null };
  };

const runSafely = async (real: ToolExecutor, call: PendingCall): Promise<ToolResult> => {
  try {
    return await real(call.name, call.args);
  } catch (err) {
    return { ok: false, error: friendlyError(err) };
  }
};

/** Runs queued calls in order. Stops at the first failure so later steps do not build on a broken one. */
export const applyPending = async (real: ToolExecutor, calls: readonly PendingCall[]): Promise<ApplyOutcome> => {
  const results: AppliedStep[] = [];
  for (const call of calls) {
    const r = await runSafely(real, call);
    results.push({ ok: r.ok, text: r.ok ? r.summary : r.error });
    if (!r.ok) return { results, skipped: calls.length - results.length };
  }
  return { results, skipped: 0 };
};
