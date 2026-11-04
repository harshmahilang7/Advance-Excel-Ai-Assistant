import { describe, expect, it } from "vitest";
import { trimHistory } from "./history";
import type { ChatMessage } from "./loop";

const user = (text: string): ChatMessage => ({ role: "user", text });
const model = (text: string): ChatMessage => ({ role: "model", text, calls: [{ name: "read_range", args: {} }] });
const tool = (): ChatMessage => ({ role: "tool", results: [{ name: "read_range", response: { ok: true } }] });

describe("trimHistory", () => {
  it("returns history unchanged when under the limit", () => {
    const h = [user("a"), model("b"), tool()];
    expect(trimHistory(h, 4)).toEqual(h);
  });

  it("keeps only the last N user turns, starting on a user message", () => {
    const h = [user("1"), model("m1"), tool(), user("2"), model("m2"), tool(), user("3"), model("m3")];
    const trimmed = trimHistory(h, 2);
    expect(trimmed[0]).toEqual(user("2"));
    expect(trimmed).toHaveLength(5);
  });

  it("never starts with a tool or model message", () => {
    const h = [user("1"), model("m"), tool(), user("2"), model("m"), tool()];
    expect(trimHistory(h, 1)[0]?.role).toBe("user");
  });

  it("does not mutate the input", () => {
    const h = Object.freeze([user("1"), user("2"), user("3")]) as ChatMessage[];
    expect(() => trimHistory(h, 1)).not.toThrow();
  });
});
