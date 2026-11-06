import { describe, expect, it } from "vitest";
import {
  addEntry,
  dropLastMatching,
  formatChangeLog,
  MAX_ENTRIES,
  MAX_PROMPT_CHARS,
  type ChangeEntry,
} from "./changeLog";

const entry = (command: string, ...steps: string[]): ChangeEntry => ({ command, steps });

describe("addEntry", () => {
  it("appends entries with steps and returns a new array", () => {
    const log: readonly ChangeEntry[] = [];
    const next = addEntry(log, entry("add total", "Wrote 3x1 cells at D1"));
    expect(next).toHaveLength(1);
    expect(log).toHaveLength(0);
  });

  it("ignores entries with no steps", () => {
    expect(addEntry([], entry("just a question"))).toEqual([]);
  });

  it("keeps only the newest entries", () => {
    const log = Array.from({ length: MAX_ENTRIES + 3 }, (_, i) => entry(`c${i}`, `s${i}`)).reduce<readonly ChangeEntry[]>(
      (acc, e) => addEntry(acc, e),
      [],
    );
    expect(log).toHaveLength(MAX_ENTRIES);
    expect(log.at(-1)?.command).toBe(`c${MAX_ENTRIES + 2}`);
  });

  it("clips very long commands and steps", () => {
    const next = addEntry([], entry("x".repeat(500), "y".repeat(500)));
    expect(next[0]?.command.length).toBeLessThanOrEqual(100);
    expect(next[0]?.steps[0]?.length).toBeLessThanOrEqual(140);
  });

  it("caps steps per entry", () => {
    const steps = Array.from({ length: 20 }, (_, i) => `step ${i}`);
    expect(addEntry([], entry("many", ...steps))[0]?.steps.length).toBeLessThanOrEqual(8);
  });
});

describe("dropLastMatching", () => {
  it("removes the newest entry whose command matches the undo label", () => {
    const log = [entry("sort by qty", "Sorted"), entry("bold header", "Formatted"), entry("sort by qty", "Sorted again")];
    const next = dropLastMatching(log, "sort by qty");
    expect(next.map((e) => e.steps[0])).toEqual(["Sorted", "Formatted"]);
  });

  it("matches on the first 60 characters, like the undo label", () => {
    const long = "z".repeat(90);
    const next = dropLastMatching([entry(long, "s")], long.slice(0, 60));
    expect(next).toEqual([]);
  });

  it("leaves the log alone when nothing matches", () => {
    const log = [entry("a", "s")];
    expect(dropLastMatching(log, "other")).toEqual(log);
  });
});

describe("formatChangeLog", () => {
  it("is empty for an empty log", () => {
    expect(formatChangeLog([])).toBe("");
  });

  it("lists commands and steps oldest first", () => {
    const text = formatChangeLog([entry("add total", "Wrote D1"), entry("bold header", "Formatted A1:D1")]);
    expect(text).toContain("add total");
    expect(text.indexOf("add total")).toBeLessThan(text.indexOf("bold header"));
    expect(text).toMatch(/data|not instructions/i);
  });

  it("stays within the size budget by dropping the oldest entries", () => {
    const log = Array.from({ length: MAX_ENTRIES }, (_, i) => entry(`command ${i} ${"c".repeat(80)}`, "s".repeat(130), "t".repeat(130), "u".repeat(130)));
    const text = formatChangeLog(log);
    expect(text.length).toBeLessThanOrEqual(MAX_PROMPT_CHARS);
    expect(text).toContain(`command ${MAX_ENTRIES - 1}`);
  });
});
