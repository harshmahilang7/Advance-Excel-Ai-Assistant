import { describe, expect, it } from "vitest";
import type { ChatMessage } from "./loop";
import { estimateHistoryTokens, estimateTokens, formatTokens, LARGE_CONTEXT_TOKENS } from "./tokens";

describe("estimateTokens", () => {
  it("is about one token per four characters, rounded up", () => {
    expect(estimateTokens("")).toBe(0);
    expect(estimateTokens("abcd")).toBe(1);
    expect(estimateTokens("abcde")).toBe(2);
  });
});

describe("estimateHistoryTokens", () => {
  it("counts user text, model text and calls, and tool results; ignores raw provider payloads", () => {
    const history: readonly ChatMessage[] = [
      { role: "user", text: "a".repeat(400) },
      { role: "model", text: "b".repeat(40), calls: [{ name: "read_range", args: { address: "A1" } }], raw: { huge: "x".repeat(100000) } },
      { role: "tool", results: [{ name: "read_range", response: { ok: true } }] },
    ];
    const n = estimateHistoryTokens(history);
    expect(n).toBeGreaterThan(110);
    expect(n).toBeLessThan(200);
  });

  it("is zero for an empty history", () => {
    expect(estimateHistoryTokens([])).toBe(0);
  });
});

describe("formatTokens", () => {
  it("shows small numbers as-is and larger ones as k", () => {
    expect(formatTokens(850)).toBe("850");
    expect(formatTokens(4200)).toBe("4.2k");
    expect(formatTokens(12000)).toBe("12k");
  });

  it("exposes a sensible warning threshold", () => {
    expect(LARGE_CONTEXT_TOKENS).toBeGreaterThan(5000);
  });
});
