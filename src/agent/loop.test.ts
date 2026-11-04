import { describe, expect, it } from "vitest";
import { runAgent, type AgentEvent, type ChatMessage, type ModelClient, type ModelTurn } from "./loop";
import type { ToolExecutor } from "./loop";

const scriptedModel = (turns: ModelTurn[]): ModelClient & { seen: ChatMessage[][] } => {
  const seen: ChatMessage[][] = [];
  let i = 0;
  return {
    seen,
    async generate(history) {
      seen.push([...history]);
      const turn = turns[i++];
      if (!turn) throw new Error("model script exhausted");
      return turn;
    },
  };
};

const okExecutor: ToolExecutor = async (name) => ({ ok: true, summary: `ran ${name}`, data: { done: true } });

describe("runAgent", () => {
  it("returns plain text when the model makes no tool calls", async () => {
    const model = scriptedModel([{ text: "Hello", calls: [] }]);
    const r = await runAgent({ model, execute: okExecutor, userMessage: "hi", context: "ctx" });
    expect(r.finalText).toBe("Hello");
    expect(r.hitStepLimit).toBe(false);
    expect(r.toolCalls).toBe(0);
  });

  it("executes tool calls, feeds results back, then finishes", async () => {
    const model = scriptedModel([
      { text: "", calls: [{ name: "write_cells", args: { address: "A1", data: [[1]] } }] },
      { text: "Wrote it.", calls: [] },
    ]);
    const events: AgentEvent[] = [];
    const r = await runAgent({
      model,
      execute: okExecutor,
      userMessage: "put 1 in A1",
      context: "ctx",
      onEvent: (e) => events.push(e),
    });
    expect(r.finalText).toBe("Wrote it.");
    expect(r.toolCalls).toBe(1);
    expect(events.map((e) => e.type)).toEqual(["tool", "text"]);
    const second = model.seen[1] ?? [];
    expect(second.at(-1)?.role).toBe("tool");
  });

  it("includes context and user message in the first turn", async () => {
    const model = scriptedModel([{ text: "ok", calls: [] }]);
    await runAgent({ model, execute: okExecutor, userMessage: "do X", context: "SHEET-INFO" });
    const first = model.seen[0]?.[0];
    expect(first?.role).toBe("user");
    expect(JSON.stringify(first)).toContain("SHEET-INFO");
    expect(JSON.stringify(first)).toContain("do X");
  });

  it("feeds executor failures back to the model instead of crashing", async () => {
    const model = scriptedModel([
      { text: "", calls: [{ name: "read_range", args: { address: "A1" } }] },
      { text: "Sorry, that failed.", calls: [] },
    ]);
    const failing: ToolExecutor = async () => {
      throw new Error("range not found");
    };
    const events: AgentEvent[] = [];
    const r = await runAgent({
      model,
      execute: failing,
      userMessage: "x",
      context: "",
      onEvent: (e) => events.push(e),
    });
    expect(r.finalText).toBe("Sorry, that failed.");
    const tool = events.find((e) => e.type === "tool");
    expect(tool && tool.type === "tool" && tool.ok).toBe(false);
    expect(JSON.stringify(model.seen[1]?.at(-1))).toContain("range not found");
  });

  it("stops at the step limit", async () => {
    const looping: ModelTurn = { text: "", calls: [{ name: "read_range", args: { address: "A1" } }] };
    const model = scriptedModel(Array.from({ length: 10 }, () => looping));
    const r = await runAgent({ model, execute: okExecutor, userMessage: "x", context: "", maxSteps: 3 });
    expect(r.hitStepLimit).toBe(true);
    expect(r.toolCalls).toBe(3);
  });

  it("stops early when aborted", async () => {
    const controller = new AbortController();
    const model = scriptedModel([
      { text: "", calls: [{ name: "read_range", args: { address: "A1" } }] },
      { text: "never", calls: [] },
    ]);
    const exec: ToolExecutor = async () => {
      controller.abort();
      return { ok: true, summary: "x", data: null };
    };
    const r = await runAgent({ model, execute: exec, userMessage: "x", context: "", signal: controller.signal });
    expect(r.aborted).toBe(true);
  });

  it("preserves the raw model turn for replay", async () => {
    const raw = { parts: [{ functionCall: { name: "read_range" }, thoughtSignature: "sig" }] };
    const model = scriptedModel([
      { text: "", calls: [{ name: "read_range", args: { address: "A1" } }], raw },
      { text: "done", calls: [] },
    ]);
    await runAgent({ model, execute: okExecutor, userMessage: "x", context: "" });
    const replayed = model.seen[1]?.find((m) => m.role === "model");
    expect(replayed && replayed.role === "model" && replayed.raw).toBe(raw);
  });
});
