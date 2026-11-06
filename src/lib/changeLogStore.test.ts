import { describe, expect, it } from "vitest";
import type { ChangeEntry } from "../agent/changeLog";
import { createChangeLogStore } from "./changeLogStore";
import type { KeyStorage } from "./keyStore";

const memory = (): KeyStorage => {
  const data = new Map<string, string>();
  return { getItem: (k) => data.get(k) ?? null, setItem: (k, v) => void data.set(k, v), removeItem: (k) => void data.delete(k) };
};

const log: readonly ChangeEntry[] = [{ command: "add total", steps: ["Wrote D1"] }];

describe("changeLogStore", () => {
  it("round-trips per workbook", () => {
    const store = createChangeLogStore(memory());
    store.save("a", log);
    expect(store.load("a")).toEqual(log);
    expect(store.load("b")).toEqual([]);
  });

  it("returns [] for corrupt or wrongly shaped data", () => {
    const storage = memory();
    const store = createChangeLogStore(storage);
    storage.setItem("griddy.changes.a", "{oops");
    expect(store.load("a")).toEqual([]);
    storage.setItem("griddy.changes.a", JSON.stringify([{ command: 1, steps: "no" }]));
    expect(store.load("a")).toEqual([]);
  });

  it("clears and never throws on broken storage", () => {
    const store = createChangeLogStore(memory());
    store.save("a", log);
    store.clear("a");
    expect(store.load("a")).toEqual([]);

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
    const safe = createChangeLogStore(broken);
    expect(() => safe.save("a", log)).not.toThrow();
    expect(safe.load("a")).toEqual([]);
    expect(() => safe.clear("a")).not.toThrow();
  });
});
