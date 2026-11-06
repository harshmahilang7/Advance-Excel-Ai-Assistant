import { describe, expect, it } from "vitest";
import { keepVisibleRows } from "./visibility";

describe("keepVisibleRows", () => {
  const rows = [["Type", "Qty"], ["A", 1], ["B", 2], ["A", 3], ["C", 4]];

  it("drops rows flagged hidden and counts them", () => {
    const r = keepVisibleRows(rows, [false, false, true, false, true]);
    expect(r.rows).toEqual([["Type", "Qty"], ["A", 1], ["A", 3]]);
    expect(r.hiddenCount).toBe(2);
    expect(r.totalCount).toBe(5);
  });

  it("always keeps the header row even if flagged hidden", () => {
    const r = keepVisibleRows(rows, [true, false, false, false, false]);
    expect(r.rows).toHaveLength(5);
    expect(r.hiddenCount).toBe(0);
  });

  it("treats missing flags as visible", () => {
    const r = keepVisibleRows(rows, []);
    expect(r.rows).toEqual(rows);
    expect(r.hiddenCount).toBe(0);
  });

  it("does not mutate its input", () => {
    const frozen = Object.freeze(rows.map((x) => Object.freeze([...x])));
    expect(() => keepVisibleRows(frozen, [false, true, false, false, false])).not.toThrow();
  });
});
