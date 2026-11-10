import { describe, expect, it } from "vitest";
import {
  buildSelfCheckMessage,
  findingsFrom,
  MAX_CHECK_RANGES,
  MAX_LISTED_FINDINGS,
  rangesToCheck,
  shouldSelfCheck,
} from "./selfCheck";

const write = (address: string, ok = true) => ({ name: "write_cells", args: { address }, ok });

describe("shouldSelfCheck", () => {
  it("needs at least two changes", () => {
    expect(shouldSelfCheck(0)).toBe(false);
    expect(shouldSelfCheck(1)).toBe(false);
    expect(shouldSelfCheck(2)).toBe(true);
  });
});

describe("rangesToCheck", () => {
  it("lists distinct addresses from successful write_cells calls", () => {
    const calls = [write("E1:E6"), write("E1:E6"), write("G1"), write("H1", false), { name: "format_range", args: { address: "A1" }, ok: true }];
    expect(rangesToCheck(calls)).toEqual(["E1:E6", "G1"]);
  });

  it("caps the number of ranges", () => {
    const calls = Array.from({ length: MAX_CHECK_RANGES + 3 }, (_, i) => write(`A${i + 1}`));
    expect(rangesToCheck(calls)).toHaveLength(MAX_CHECK_RANGES);
  });

  it("ignores calls without a string address", () => {
    expect(rangesToCheck([{ name: "write_cells", args: null, ok: true }, { name: "write_cells", args: { address: 5 }, ok: true }])).toEqual([]);
  });
});

describe("findingsFrom", () => {
  it("reads findings from audit data", () => {
    const data = { findings: [{ kind: "error_value", address: "E3", detail: "Cell shows #REF!" }] };
    expect(findingsFrom(data)).toEqual([{ kind: "error_value", address: "E3", detail: "Cell shows #REF!" }]);
  });

  it("returns nothing for missing or malformed data", () => {
    expect(findingsFrom(null)).toEqual([]);
    expect(findingsFrom({ findings: "no" })).toEqual([]);
    expect(findingsFrom({ findings: [{ kind: 1 }] })).toEqual([]);
  });
});

describe("buildSelfCheckMessage", () => {
  const f = (n: number) => ({ kind: "error_value", address: `E${n}`, detail: "Cell shows #REF!" });

  it("lists findings and frames them as hints", () => {
    const msg = buildSelfCheckMessage([f(3)]);
    expect(msg).toContain("E3");
    expect(msg).toMatch(/hints/i);
    expect(msg).toMatch(/final summary/i);
  });

  it("truncates long lists", () => {
    const msg = buildSelfCheckMessage(Array.from({ length: MAX_LISTED_FINDINGS + 4 }, (_, i) => f(i + 1)));
    expect(msg).toContain("and 4 more");
  });
});
