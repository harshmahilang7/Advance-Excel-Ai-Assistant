import { describe, expect, it } from "vitest";
import { addTouched, EMPTY_TOUCHED, formatChanged, totalCells } from "./touched";

describe("touched ranges", () => {
  it("totals cells when every write reported a range", () => {
    const s = addTouched(addTouched(EMPTY_TOUCHED, { address: "Sheet1!A1:B2", cells: 4 }), { address: "Sheet1!D1", cells: 1 });
    expect(totalCells(s)).toBe(5);
  });

  it("gives no total when a write did not report its size", () => {
    const s = addTouched(addTouched(EMPTY_TOUCHED, { address: "Sheet1!A1", cells: 1 }), undefined);
    expect(totalCells(s)).toBeNull();
    expect(totalCells(EMPTY_TOUCHED)).toBeNull();
  });

  it("does not double count the same range", () => {
    const t = { address: "Sheet1!A1:A3", cells: 3 };
    expect(addTouched(addTouched(EMPTY_TOUCHED, t), t).ranges).toHaveLength(1);
  });

  it("lists ranges as backticked addresses, capped", () => {
    expect(formatChanged(EMPTY_TOUCHED)).toBe("");
    const s = addTouched(EMPTY_TOUCHED, { address: "Sheet1!A1:B2", cells: 4 });
    expect(formatChanged(s)).toBe("Changed: `Sheet1!A1:B2`");
    const many = Array.from({ length: 10 }, (_, i) => ({ address: `Sheet1!A${i + 1}`, cells: 1 })).reduce(addTouched, EMPTY_TOUCHED);
    expect(formatChanged(many)).toMatch(/and 2 more$/);
  });
});
