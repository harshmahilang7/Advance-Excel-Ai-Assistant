import { describe, expect, it } from "vitest";
import { cleanGrid, parseLooseNumber } from "./cleaning";

describe("parseLooseNumber", () => {
  it.each([
    ["1,234.50", 1234.5],
    ["$ 12", 12],
    ["(45)", -45],
    ["-7", -7],
    ["12%", 0.12],
    ["  3.5 ", 3.5],
  ])("parses %s", (input, expected) => {
    expect(parseLooseNumber(input)).toBeCloseTo(expected);
  });

  it.each(["", "abc", "12abc", "1,2,3x", "--5", "1.2.3"])("rejects %s", (input) => {
    expect(parseLooseNumber(input)).toBeNull();
  });
});

describe("cleanGrid", () => {
  it("trims and collapses whitespace including non-breaking spaces", () => {
    const r = cleanGrid([["  a   b  ", "ok"]], ["trim"]);
    expect(r.values).toEqual([["a b", "ok"]]);
    expect(r.changed).toBe(1);
  });

  it("changes case", () => {
    expect(cleanGrid([["hello WORLD"]], ["upper"]).values).toEqual([["HELLO WORLD"]]);
    expect(cleanGrid([["hello WORLD"]], ["lower"]).values).toEqual([["hello world"]]);
    expect(cleanGrid([["hello WORLD o'neil"]], ["proper"]).values).toEqual([["Hello World O'neil"]]);
  });

  it("strips non-printing characters", () => {
    expect(cleanGrid([["a\u0007b​c"]], ["strip_nonprinting"]).values).toEqual([["abc"]]);
  });

  it("converts numeric text to numbers and leaves other text alone", () => {
    const r = cleanGrid([["1,200", "n/a", 5]], ["text_to_number"]);
    expect(r.values).toEqual([[1200, "n/a", 5]]);
    expect(r.changed).toBe(1);
  });

  it("fills blanks down from the cell above", () => {
    const r = cleanGrid([["East"], [""], [""], ["West"], [""]], ["fill_down"]);
    expect(r.values).toEqual([["East"], ["East"], ["East"], ["West"], ["West"]]);
    expect(r.changed).toBe(3);
  });

  it("does not fill leading blanks", () => {
    expect(cleanGrid([[""], ["A"]], ["fill_down"]).values).toEqual([[""], ["A"]]);
  });

  it("never touches formulas", () => {
    const r = cleanGrid([["=A1 "], [" x "]], ["trim", "upper"]);
    expect(r.values).toEqual([["=A1 "], ["X"]]);
    expect(r.changed).toBe(1);
  });

  it("leaves numbers and booleans untouched by text operations", () => {
    expect(cleanGrid([[1, true]], ["trim", "upper"]).values).toEqual([[1, true]]);
  });

  it("applies operations in a fixed order regardless of request order", () => {
    const a = cleanGrid([["  1,000 "]], ["text_to_number", "trim"]);
    expect(a.values).toEqual([[1000]]);
  });

  it("does not mutate input and counts per-operation changes", () => {
    const input = [[" a "]] as const;
    const r = cleanGrid(input, ["trim", "upper"]);
    expect(input).toEqual([[" a "]]);
    expect(r.perOperation.trim).toBe(1);
    expect(r.perOperation.upper).toBe(1);
  });

  it("reports zero changes for clean data", () => {
    expect(cleanGrid([["a", 1]], ["strip_nonprinting", "trim", "fill_down", "text_to_number"]).changed).toBe(0);
  });
});
