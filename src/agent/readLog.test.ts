import { describe, expect, it } from "vitest";
import { addRead, describeRead, formatReadLog, MAX_LISTED } from "./readLog";

describe("describeRead", () => {
  it("returns the address for read_range", () => {
    expect(describeRead("read_range", { address: "Sheet1!A1:D20" })).toBe("Sheet1!A1:D20");
  });

  it("uses the given address for audit_formulas", () => {
    expect(describeRead("audit_formulas", { address: "B2:B50" })).toBe("B2:B50");
  });

  it("falls back to the whole sheet for audit_formulas without an address", () => {
    expect(describeRead("audit_formulas", {})).toBe("active sheet (formula audit)");
  });

  it("describes get_workbook_info", () => {
    expect(describeRead("get_workbook_info", {})).toBe("workbook overview");
  });

  it("returns null for write tools and bad args", () => {
    expect(describeRead("write_cells", { address: "A1" })).toBeNull();
    expect(describeRead("read_range", null)).toBeNull();
    expect(describeRead("read_range", { address: 5 })).toBeNull();
  });
});

describe("addRead", () => {
  it("appends without mutating and skips duplicates", () => {
    const base: readonly string[] = ["A1:B2"];
    const next = addRead(base, "C1:D2");
    expect(next).toEqual(["A1:B2", "C1:D2"]);
    expect(base).toEqual(["A1:B2"]);
    expect(addRead(next, "A1:B2")).toBe(next);
  });
});

describe("formatReadLog", () => {
  it("always mentions the automatic selection summary", () => {
    expect(formatReadLog([])).toBe("Sent to the AI: your selection summary only (no extra ranges read).");
  });

  it("lists ranges the model read", () => {
    expect(formatReadLog(["A1:B2", "Sheet2!C1:C9"])).toBe(
      "Sent to the AI: your selection summary, plus ranges it read: A1:B2, Sheet2!C1:C9.",
    );
  });

  it("truncates long lists", () => {
    const many = Array.from({ length: MAX_LISTED + 3 }, (_, i) => `A${i + 1}`);
    const out = formatReadLog(many);
    expect(out).toContain("and 3 more");
    expect(out).not.toContain(`A${MAX_LISTED + 1},`);
  });
});
