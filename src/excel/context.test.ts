import { describe, expect, it } from "vitest";
import { summarizeRange } from "./context";

const input = (values: (string | number | boolean | null)[][]) => ({
  sheetName: "Sales",
  address: "Sales!A1:C4",
  values,
});

describe("summarizeRange", () => {
  const values = [
    ["Region", "Units", "Active"],
    ["North", 10, true],
    ["South", 20, false],
    ["East", 30, true],
  ];

  it("extracts headers and dimensions", () => {
    const s = summarizeRange(input(values));
    expect(s.headers).toEqual(["Region", "Units", "Active"]);
    expect(s.dataRowCount).toBe(3);
    expect(s.columnCount).toBe(3);
    expect(s.sheetName).toBe("Sales");
    expect(s.address).toBe("Sales!A1:C4");
  });

  it("infers column types from data rows", () => {
    expect(summarizeRange(input(values)).columnTypes).toEqual(["text", "number", "boolean"]);
  });

  it("flags mixed and empty columns", () => {
    const s = summarizeRange(
      input([
        ["A", "B", "C"],
        ["x", null, 1],
        [5, null, "y"],
      ]),
    );
    expect(s.columnTypes).toEqual(["mixed", "empty", "mixed"]);
  });

  it("limits sample rows without mutating input", () => {
    const frozen = Object.freeze(values.map((r) => Object.freeze([...r])));
    const s = summarizeRange(input(frozen as never), { maxSampleRows: 2 });
    expect(s.sampleRows).toEqual([
      ["North", 10, true],
      ["South", 20, false],
    ]);
    expect(s.truncated).toBe(true);
  });

  it("handles an empty range", () => {
    const s = summarizeRange(input([]));
    expect(s.headers).toEqual([]);
    expect(s.dataRowCount).toBe(0);
    expect(s.sampleRows).toEqual([]);
  });

  it("stringifies blank header cells as column letters", () => {
    const s = summarizeRange(input([["Name", null], ["a", 1]]));
    expect(s.headers).toEqual(["Name", "Column B"]);
  });
});
