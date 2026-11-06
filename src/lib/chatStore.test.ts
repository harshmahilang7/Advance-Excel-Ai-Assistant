import { describe, expect, it } from "vitest";
import { createChatStore, MAX_STORED_LINES, type Line } from "./chatStore";
import type { KeyStorage } from "./keyStore";

const memory = (): KeyStorage & { readonly data: Map<string, string> } => {
  const data = new Map<string, string>();
  return {
    data,
    getItem: (k) => data.get(k) ?? null,
    setItem: (k, v) => void data.set(k, v),
    removeItem: (k) => void data.delete(k),
  };
};

const lines: readonly Line[] = [
  { id: 1, kind: "user", text: "hi" },
  { id: 2, kind: "action", text: "✓ done", ok: true },
  { id: 3, kind: "error", text: "bad" },
];

describe("chatStore", () => {
  it("round-trips lines per workbook key", () => {
    const storage = memory();
    const store = createChatStore(storage);
    store.save("book-a", lines);
    expect(store.load("book-a")).toEqual(lines);
    expect(store.load("book-b")).toEqual([]);
  });

  it("returns [] for corrupt or invalid data", () => {
    const storage = memory();
    const store = createChatStore(storage);
    storage.setItem("griddy.chat.x", "{not json");
    expect(store.load("x")).toEqual([]);
    storage.setItem("griddy.chat.x", JSON.stringify([{ id: 1, kind: "hacker", text: "x" }]));
    expect(store.load("x")).toEqual([]);
  });

  it("keeps only the newest lines", () => {
    const store = createChatStore(memory());
    const many: Line[] = Array.from({ length: MAX_STORED_LINES + 5 }, (_, i) => ({ id: i + 1, kind: "note", text: `n${i}` }));
    store.save("k", many);
    const loaded = store.load("k");
    expect(loaded).toHaveLength(MAX_STORED_LINES);
    expect(loaded.at(-1)?.id).toBe(many.at(-1)?.id);
  });

  it("swallows storage failures", () => {
    const broken: KeyStorage = {
      getItem: () => {
        throw new Error("denied");
      },
      setItem: () => {
        throw new Error("quota");
      },
      removeItem: () => {
        throw new Error("denied");
      },
    };
    const store = createChatStore(broken);
    expect(() => store.save("k", lines)).not.toThrow();
    expect(store.load("k")).toEqual([]);
    expect(() => store.clear("k")).not.toThrow();
  });

  it("clears a workbook's chat", () => {
    const store = createChatStore(memory());
    store.save("k", lines);
    store.clear("k");
    expect(store.load("k")).toEqual([]);
  });
});
