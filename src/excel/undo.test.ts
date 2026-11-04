import { describe, expect, it } from "vitest";
import { createUndoLog } from "./undo";

describe("undoLog", () => {
  it("starts empty", () => {
    expect(createUndoLog().canUndo()).toBe(false);
  });

  it("does not count empty groups", () => {
    const log = createUndoLog();
    log.beginGroup("nothing happened");
    expect(log.canUndo()).toBe(false);
  });

  it("restores the latest group in reverse order", async () => {
    const calls: string[] = [];
    const log = createUndoLog();
    log.beginGroup("cmd 1");
    log.record("a", async () => void calls.push("a"));
    log.record("b", async () => void calls.push("b"));
    const r = await log.undoLast();
    expect(calls).toEqual(["b", "a"]);
    expect(r).toEqual({ label: "cmd 1", restored: 2, failed: [] });
    expect(log.canUndo()).toBe(false);
  });

  it("keeps older groups until undone", async () => {
    const calls: string[] = [];
    const log = createUndoLog();
    log.beginGroup("first");
    log.record("x", async () => void calls.push("x"));
    log.beginGroup("second");
    log.record("y", async () => void calls.push("y"));
    await log.undoLast();
    expect(calls).toEqual(["y"]);
    expect(log.canUndo()).toBe(true);
    await log.undoLast();
    expect(calls).toEqual(["y", "x"]);
  });

  it("continues past a failing restore and reports it", async () => {
    const calls: string[] = [];
    const log = createUndoLog();
    log.beginGroup("cmd");
    log.record("ok", async () => void calls.push("ok"));
    log.record("bad", async () => {
      throw new Error("boom");
    });
    const r = await log.undoLast();
    expect(calls).toEqual(["ok"]);
    expect(r?.restored).toBe(1);
    expect(r?.failed).toEqual(["bad: boom"]);
  });

  it("returns null when there is nothing to undo", async () => {
    expect(await createUndoLog().undoLast()).toBeNull();
  });

  it("caps stored groups", async () => {
    const log = createUndoLog(2);
    for (const n of ["1", "2", "3"]) {
      log.beginGroup(n);
      log.record(n, async () => {});
    }
    expect((await log.undoLast())?.label).toBe("3");
    expect((await log.undoLast())?.label).toBe("2");
    expect(log.canUndo()).toBe(false);
  });
});
