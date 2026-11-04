import { parseToolCall } from "./tools";

export interface RawToolCall {
  readonly name: string;
  readonly args: unknown;
}

export interface ModelTurn {
  readonly text: string;
  readonly calls: readonly RawToolCall[];
  /** Provider-native turn, replayed verbatim (keeps e.g. Gemini thought signatures intact). */
  readonly raw?: unknown;
}

export interface ToolResponse {
  readonly name: string;
  readonly response: Record<string, unknown>;
}

export type ChatMessage =
  | { readonly role: "user"; readonly text: string }
  | { readonly role: "model"; readonly text: string; readonly calls: readonly RawToolCall[]; readonly raw?: unknown }
  | { readonly role: "tool"; readonly results: readonly ToolResponse[] };

export interface ModelClient {
  generate(history: readonly ChatMessage[], signal?: AbortSignal): Promise<ModelTurn>;
}

export type ToolResult =
  | { readonly ok: true; readonly summary: string; readonly data: unknown }
  | { readonly ok: false; readonly error: string };

export type ToolExecutor = (name: string, args: unknown) => Promise<ToolResult>;

export type AgentEvent =
  | { readonly type: "tool"; readonly name: string; readonly args: unknown; readonly ok: boolean; readonly summary: string }
  | { readonly type: "text"; readonly text: string };

export interface AgentRun {
  readonly model: ModelClient;
  readonly execute: ToolExecutor;
  readonly userMessage: string;
  /** Fresh workbook context (selection summary etc.) prepended to this user turn. */
  readonly context: string;
  readonly history?: readonly ChatMessage[];
  readonly maxSteps?: number;
  readonly signal?: AbortSignal;
  readonly onEvent?: (event: AgentEvent) => void;
}

export interface AgentResult {
  readonly finalText: string;
  readonly toolCalls: number;
  readonly hitStepLimit: boolean;
  readonly aborted: boolean;
  readonly history: readonly ChatMessage[];
}

const DEFAULT_MAX_STEPS = 12;

const errorMessage = (err: unknown): string => (err instanceof Error ? err.message : String(err));

const runOne = async (execute: ToolExecutor, call: RawToolCall): Promise<ToolResult> => {
  const parsed = parseToolCall(call.name, call.args);
  if (!parsed.ok) return { ok: false, error: parsed.error };
  try {
    return await execute(parsed.call.name, parsed.call.args);
  } catch (err) {
    return { ok: false, error: errorMessage(err) };
  }
};

const toResponse = (name: string, result: ToolResult): ToolResponse => ({
  name,
  response: result.ok ? { ok: true, summary: result.summary, data: result.data } : { ok: false, error: result.error },
});

/** Model <-> tool loop. Pure orchestration: no Excel or network code lives here. */
export const runAgent = async (run: AgentRun): Promise<AgentResult> => {
  const maxSteps = run.maxSteps ?? DEFAULT_MAX_STEPS;
  let history: readonly ChatMessage[] = [
    ...(run.history ?? []),
    { role: "user", text: `${run.context}\n\nUser request: ${run.userMessage}` },
  ];
  let toolCalls = 0;
  let lastText = "";

  for (let step = 0; step < maxSteps; step++) {
    if (run.signal?.aborted) return { finalText: lastText, toolCalls, hitStepLimit: false, aborted: true, history };

    const turn = await run.model.generate(history, run.signal);
    history = [...history, { role: "model", text: turn.text, calls: turn.calls, raw: turn.raw }];
    if (turn.text) lastText = turn.text;

    if (turn.calls.length === 0) {
      if (turn.text) run.onEvent?.({ type: "text", text: turn.text });
      return { finalText: turn.text, toolCalls, hitStepLimit: false, aborted: false, history };
    }

    const results: ToolResponse[] = [];
    for (const call of turn.calls) {
      const result = await runOne(run.execute, call);
      toolCalls += 1;
      results.push(toResponse(call.name, result));
      run.onEvent?.({
        type: "tool",
        name: call.name,
        args: call.args,
        ok: result.ok,
        summary: result.ok ? result.summary : result.error,
      });
    }
    history = [...history, { role: "tool", results }];

    if (run.signal?.aborted) return { finalText: lastText, toolCalls, hitStepLimit: false, aborted: true, history };
  }

  return { finalText: lastText, toolCalls, hitStepLimit: true, aborted: false, history };
};
