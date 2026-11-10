import { describe, expect, it } from "vitest";
import { describeVerification, verifyWrite, MAX_REPORTED } from "./writeVerify";

describe("verifyWrite", () => {
  it("matches identical constants and formulas", () => {
    const r = verifyWrite([["a", 1, "=SUM(A1:A3)"]], [["a", 1, "=SUM(A1:A3)"]], [["a", 1, 6]]);
    expect(r).toMatchObject({ total: 3, matched: 3, differed: [], errors: [] });
  });

  it("ignores case and spaces inside formulas", () => {
    const r = verifyWrite([["=sum( A1:A3 )"]], [["=SUM(A1:A3)"]], [[6]]);
    expect(r.matched).toBe(1);
  });

  it("accepts numeric text that Excel stored as a number", () => {
    const r = verifyWrite([["123", "1,5"]], [[123, "1,5"]], [[123, "1,5"]]);
    expect(r.matched).toBe(2);
  });

  it("treats empty text as a blank cell", () => {
    const r = verifyWrite([["", ""]], [["", ""]], [["", ""]]);
    expect(r.matched).toBe(2);
  });

  it("skips null cells: they were left unchanged, so whatever is there is fine", () => {
    const r = verifyWrite([["a", null]], [["a", "old"]], [["a", "#N/A"]]);
    expect(r).toMatchObject({ total: 1, matched: 1, differed: [], errors: [] });
  });

  it("reports cells that read back differently with A1-style positions", () => {
    const r = verifyWrite([["2024-01-05", "ok"]], [[45296, "ok"]], [[45296, "ok"]], { row: 2, col: 1 });
    expect(r.matched).toBe(1);
    expect(r.differed).toEqual([{ cell: "B3", wrote: "2024-01-05", got: "45296" }]);
  });

  it("reports error values in written cells", () => {
    const r = verifyWrite([["=A1/0", "=B9"]], [["=A1/0", "=B9"]], [["#DIV/0!", "#REF!"]], { row: 0, col: 3 });
    expect(r.errors).toEqual([
      { cell: "D1", value: "#DIV/0!" },
      { cell: "E1", value: "#REF!" },
    ]);
  });

  it("does not mutate its inputs", () => {
    const wrote = [["a"]];
    verifyWrite(wrote, [["a"]], [["a"]]);
    expect(wrote).toEqual([["a"]]);
  });
});

describe("describeVerification", () => {
  it("says verified when all match and no errors", () => {
    expect(describeVerification({ total: 4, matched: 4, differed: [], errors: [] })).toBe(" (verified 4/4 cells)");
  });

  it("warns about differences with examples", () => {
    const out = describeVerification({
      total: 10,
      matched: 9,
      differed: [{ cell: "B2", wrote: "2024-01-05", got: "45296" }],
      errors: [],
    });
    expect(out).toBe(' ⚠ 1 of 10 cells read back differently (B2: wrote "2024-01-05", Excel has "45296")');
  });

  it("warns about errors", () => {
    const out = describeVerification({ total: 2, matched: 2, differed: [], errors: [{ cell: "D1", value: "#DIV/0!" }] });
    expect(out).toBe(" ⚠ 1 cell shows an error (D1: #DIV/0!)");
  });

  it("limits listed examples", () => {
    const errors = Array.from({ length: MAX_REPORTED + 2 }, (_, i) => ({ cell: `A${i + 1}`, value: "#N/A" }));
    const out = describeVerification({ total: 20, matched: 20, differed: [], errors });
    expect(out).toContain("and 2 more");
  });
});
