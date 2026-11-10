import { describe, expect, it } from "vitest";
import { addTemplate, createTemplateStore, MAX_TEMPLATES, nameFor, removeTemplate, type Template } from "./templateStore";

const memory = () => {
  const data = new Map<string, string>();
  return {
    getItem: (k: string) => data.get(k) ?? null,
    setItem: (k: string, v: string) => void data.set(k, v),
    removeItem: (k: string) => void data.delete(k),
  };
};
let n = 0;
const nextId = () => `t${++n}`;

describe("nameFor", () => {
  it("uses the prompt, clipped with an ellipsis", () => {
    expect(nameFor("Add a Total column")).toBe("Add a Total column");
    expect(nameFor("x".repeat(60))).toHaveLength(40);
    expect(nameFor("x".repeat(60)).endsWith("…")).toBe(true);
  });

  it("collapses whitespace", () => {
    expect(nameFor("  Add   a\ntotal ")).toBe("Add a total");
  });
});

describe("addTemplate", () => {
  it("adds without mutating", () => {
    const base: readonly Template[] = [];
    const r = addTemplate(base, "Bold the header", nextId);
    expect(r.status).toBe("added");
    expect(r.list).toHaveLength(1);
    expect(base).toHaveLength(0);
  });

  it("ignores empty prompts and duplicates", () => {
    expect(addTemplate([], "   ", nextId).status).toBe("empty");
    const once = addTemplate([], "Bold the header", nextId).list;
    const again = addTemplate(once, "  Bold the header ", nextId);
    expect(again.status).toBe("duplicate");
    expect(again.list).toBe(once);
  });

  it("refuses new templates when full", () => {
    let list: readonly Template[] = [];
    for (let i = 0; i < MAX_TEMPLATES; i++) list = addTemplate(list, `prompt ${i}`, nextId).list;
    const r = addTemplate(list, "one more", nextId);
    expect(r.status).toBe("full");
    expect(r.list).toBe(list);
  });
});

describe("removeTemplate", () => {
  it("removes by id without mutating", () => {
    const list = addTemplate(addTemplate([], "a", nextId).list, "b", nextId).list;
    const id = list[0]?.id ?? "";
    expect(removeTemplate(list, id).map((t) => t.prompt)).toEqual(["b"]);
    expect(list).toHaveLength(2);
  });
});

describe("createTemplateStore", () => {
  it("round-trips templates", () => {
    const store = createTemplateStore(memory());
    const list = addTemplate([], "Bold the header", nextId).list;
    store.save(list);
    expect(store.load()).toEqual(list);
  });

  it("returns an empty list for missing, corrupt or wrong-shaped data", () => {
    const storage = memory();
    const store = createTemplateStore(storage);
    expect(store.load()).toEqual([]);
    storage.setItem("griddy.templates", "{not json");
    expect(store.load()).toEqual([]);
    storage.setItem("griddy.templates", JSON.stringify([{ id: 1 }]));
    expect(store.load()).toEqual([]);
  });

  it("never throws when storage fails", () => {
    const fail = () => {
      throw new Error("blocked");
    };
    const store = createTemplateStore({ getItem: fail, setItem: fail, removeItem: fail });
    expect(store.load()).toEqual([]);
    expect(() => store.save([])).not.toThrow();
  });
});
