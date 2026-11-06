import { describe, expect, it } from "vitest";
import { runAgent, type ModelClient, type StreamEvent } from "./loop";

const okExec = async () => ({ ok: true as const, summary: "done", data: null });

describe("runAgent streaming", () => {
  it("forwards text deltas and signals the end of every model turn", async () => {
    let call = 0;
    const model: ModelClient = {
      async generate(_history, _signal, onText) {
        call += 1;
        if (call === 1) {
          onText?.("Let me ");
          onText?.("check.");
          return { text: "Let me check.", calls: [{ name: "get_workbook_info", args: {} }] };
        }
        onText?.("All ");
        onText?.("set.");
        return { text: "All set.", calls: [] };
      },
    };
    const events: StreamEvent[] = [];
    await runAgent({ model, execute: okExec, userMessage: "go", context: "ctx", onStream: (e) => events.push(e) });
    expect(events).toEqual([
      { type: "delta", text: "Let me " },
      { type: "delta", text: "check." },
      { type: "end" },
      { type: "delta", text: "All " },
      { type: "delta", text: "set." },
      { type: "end" },
    ]);
  });

  it("still ends the turn when the model throws", async () => {
    const model: ModelClient = {
      async generate(_h, _s, onText) {
        onText?.("partial");
        throw new Error("network");
      },
    };
    const events: StreamEvent[] = [];
    await expect(runAgent({ model, execute: okExec, userMessage: "x", context: "c", onStream: (e) => events.push(e) })).rejects.toThrow("network");
    expect(events.at(-1)).toEqual({ type: "end" });
  });

  it("works with models that ignore the streaming callback", async () => {
    const model: ModelClient = { generate: async () => ({ text: "hi", calls: [] }) };
    const r = await runAgent({ model, execute: okExec, userMessage: "x", context: "c" });
    expect(r.finalText).toBe("hi");
  });
});
