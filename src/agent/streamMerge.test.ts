import type { Content } from "@google/genai";
import { describe, expect, it } from "vitest";
import { deltaText, mergeChunks } from "./streamMerge";

const model = (...parts: Content["parts"] & object): Content => ({ role: "model", parts });

describe("deltaText", () => {
  it("returns visible text and skips thoughts and function calls", () => {
    const c = model({ text: "hidden", thought: true }, { text: "Hello " }, { functionCall: { name: "x", args: {} } }, { text: "world" });
    expect(deltaText(c)).toBe("Hello world");
  });

  it("is empty for undefined or part-less content", () => {
    expect(deltaText(undefined)).toBe("");
    expect(deltaText({ role: "model" })).toBe("");
  });
});

describe("mergeChunks", () => {
  it("returns undefined when there is nothing", () => {
    expect(mergeChunks([])).toBeUndefined();
    expect(mergeChunks([undefined, { role: "model", parts: [] }])).toBeUndefined();
  });

  it("joins adjacent plain text parts", () => {
    const merged = mergeChunks([model({ text: "Hel" }), model({ text: "lo" })]);
    expect(merged).toEqual({ role: "model", parts: [{ text: "Hello" }] });
  });

  it("keeps thought parts and signatures verbatim (never merged)", () => {
    const merged = mergeChunks([
      model({ text: "thinking", thought: true }),
      model({ text: "a", thoughtSignature: "SIG1" }),
      model({ text: "b" }),
    ]);
    expect(merged?.parts).toEqual([
      { text: "thinking", thought: true },
      { text: "a", thoughtSignature: "SIG1" },
      { text: "b" },
    ]);
  });

  it("keeps function calls in order between text", () => {
    const call = { functionCall: { name: "write_cells", args: { address: "A1" } }, thoughtSignature: "S" };
    const merged = mergeChunks([model({ text: "ok " }), model(call), model({ text: "done" })]);
    expect(merged?.parts).toEqual([{ text: "ok " }, call, { text: "done" }]);
  });

  it("does not mutate the input chunks", () => {
    const a = model({ text: "x" });
    const b = model({ text: "y" });
    const before = JSON.stringify([a, b]);
    mergeChunks([a, b]);
    expect(JSON.stringify([a, b])).toBe(before);
  });
});
