import { describe, expect, it } from "vitest";
import { createKeyStore, type KeyStorage } from "./keyStore";

const memoryStorage = (): KeyStorage => {
  const data = new Map<string, string>();
  return {
    getItem: (k) => data.get(k) ?? null,
    setItem: (k, v) => void data.set(k, v),
    removeItem: (k) => void data.delete(k),
  };
};

describe("keyStore", () => {
  it("returns null when no key saved", () => {
    expect(createKeyStore(memoryStorage()).get()).toBeNull();
  });

  it("saves and reads a trimmed key", () => {
    const store = createKeyStore(memoryStorage());
    store.set("  AIzaTestKey123  ");
    expect(store.get()).toBe("AIzaTestKey123");
  });

  it("rejects empty or whitespace-only keys", () => {
    const store = createKeyStore(memoryStorage());
    expect(() => store.set("   ")).toThrow(/empty/i);
    expect(store.get()).toBeNull();
  });

  it("clears the key", () => {
    const store = createKeyStore(memoryStorage());
    store.set("abc");
    store.clear();
    expect(store.get()).toBeNull();
  });

  it("survives a storage that throws (private mode / blocked)", () => {
    const broken: KeyStorage = {
      getItem: () => {
        throw new Error("blocked");
      },
      setItem: () => {
        throw new Error("blocked");
      },
      removeItem: () => {
        throw new Error("blocked");
      },
    };
    const store = createKeyStore(broken);
    expect(store.get()).toBeNull();
    expect(() => store.set("abc")).toThrow(/could not save/i);
  });
});
