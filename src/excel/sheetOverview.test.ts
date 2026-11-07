import { describe, expect, it } from "vitest";
import { summarizeSheets, MAX_OVERVIEW_HEADERS, MAX_OVERVIEW_SHEETS } from "./sheetOverview";

describe("summarizeSheets", () => {
  it("lists each sheet with its used range and headers", () => {
    const out = summarizeSheets([
      { name: "Sales", usedRange: "Sales!A1:D6", firstRow: ["Region", "Qty", "", null] },
      { name: "Empty", usedRange: null, firstRow: [] },
    ]);
    expect(out.sheets).toEqual([
      { name: "Sales", usedRange: "Sales!A1:D6", headers: ["Region", "Qty", "Column C", "Column D"] },
      { name: "Empty", usedRange: null, headers: [] },
    ]);
    expect(out.omittedSheets).toBe(0);
  });

  it("caps headers per sheet and number of sheets", () => {
    const wide = Array.from({ length: MAX_OVERVIEW_HEADERS + 5 }, (_, i) => `H${i}`);
    const many = Array.from({ length: MAX_OVERVIEW_SHEETS + 3 }, (_, i) => ({ name: `S${i}`, usedRange: `S${i}!A1:A1`, firstRow: wide }));
    const out = summarizeSheets(many);
    expect(out.sheets).toHaveLength(MAX_OVERVIEW_SHEETS);
    expect(out.omittedSheets).toBe(3);
    expect(out.sheets[0]?.headers).toHaveLength(MAX_OVERVIEW_HEADERS);
  });

  it("truncates long header text and stringifies numbers", () => {
    const out = summarizeSheets([{ name: "A", usedRange: "A!A1:B1", firstRow: ["x".repeat(100), 2024] }]);
    expect(out.sheets[0]?.headers[0]?.length).toBeLessThanOrEqual(40);
    expect(out.sheets[0]?.headers[1]).toBe("2024");
  });
});
