import { describe, expect, it } from "vitest";
import { checkPivotDestination, rectContains, type Rect } from "./pivotChecks";

const source: Rect = { sheet: "Sheet1", top: 0, left: 0, bottom: 9, right: 3 };

describe("rectContains", () => {
  it("includes the edges", () => {
    expect(rectContains(source, "Sheet1", 0, 0)).toBe(true);
    expect(rectContains(source, "Sheet1", 9, 3)).toBe(true);
  });

  it("excludes cells outside or on another sheet", () => {
    expect(rectContains(source, "Sheet1", 10, 0)).toBe(false);
    expect(rectContains(source, "Sheet2", 0, 0)).toBe(false);
  });

  it("compares sheet names case-insensitively", () => {
    expect(rectContains(source, "sheet1", 1, 1)).toBe(true);
  });
});

describe("checkPivotDestination", () => {
  it("allows an empty cell away from the source", () => {
    expect(checkPivotDestination({ sheet: "Sheet1", row: 0, col: 6, address: "G1", hasData: false, source })).toBeNull();
  });

  it("rejects a destination inside the source data", () => {
    const err = checkPivotDestination({ sheet: "Sheet1", row: 2, col: 1, address: "B3", hasData: true, source });
    expect(err).toMatch(/^DESTINATION_INSIDE_SOURCE:/);
  });

  it("rejects a non-empty destination outside the source", () => {
    const err = checkPivotDestination({ sheet: "Sheet1", row: 0, col: 6, address: "G1", hasData: true, source });
    expect(err).toMatch(/^DESTINATION_OCCUPIED:/);
    expect(err).toContain("G1");
  });

  it("allows an empty destination on another sheet", () => {
    expect(checkPivotDestination({ sheet: "Report", row: 0, col: 0, address: "Report!A1", hasData: false, source })).toBeNull();
  });
});
