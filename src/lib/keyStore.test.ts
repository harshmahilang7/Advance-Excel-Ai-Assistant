import { describe, expect, it } from "vitest";
import { createKeyStore, parseKeys, type KeyStorage } from "./keyStore";

const memoryStorage = (): KeyStorage => {
  const data = new Map<string, string>();
  return {
    getItem: (k) => data.get(k) ?? null,
    setItem: (k, v) => void data.set(k, v),
    removeItem: (k) => void data.delete(k),
  };
};

describe("keyStore", () => {
  it("returns an empty list when nothing saved", () => {
    expect(createKeyStore(memoryStorage()).list()).toEqual([]);
  });

  it("saves and reads trimmed, de-duplicated keys in order", () => {
    const store = createKeyStore(memoryStorage());
    store.save(["  AIzaOne  ", "AIzaTwo", "AIzaOne"]);
    expect(store.list()).toEqual(["AIzaOne", "AIzaTwo"]);
  });

  it("rejects an empty set of keys", () => {
    const store = createKeyStore(memoryStorage());
    expect(() => store.save(["   "])).toThrow(/empty/i);
    expect(() => store.save([])).toThrow(/empty/i);
    expect(store.list()).toEqual([]);
  });

  it("clears all keys", () => {
    const store = createKeyStore(memoryStorage());
    store.save(["abc", "def"]);
    store.clear();
    expect(store.list()).toEqual([]);
  });

  it("migrates the legacy single key", () => {
    const storage = memoryStorage();
    storage.setItem("griddy.geminiApiKey", "legacyKey");
    const store = createKeyStore(storage);
    expect(store.list()).toEqual(["legacyKey"]);
    store.save(["legacyKey", "newKey"]);
    expect(storage.getItem("griddy.geminiApiKey")).toBeNull();
    expect(store.list()).toEqual(["legacyKey", "newKey"]);
  });

  it("ignores corrupt stored data", () => {
    const storage = memoryStorage();
    storage.setItem("griddy.geminiApiKeys", "{not json");
    expect(createKeyStore(storage).list()).toEqual([]);
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
    expect(store.list()).toEqual([]);
    expect(() => store.save(["abc"])).toThrow(/could not save/i);
  });
});

describe("parseKeys", () => {
  it("splits on newlines, commas and spaces and drops blanks and duplicates", () => {
    expect(parseKeys("a\n b ,c;\n\n a")).toEqual(["a", "b", "c"]);
  });
});
