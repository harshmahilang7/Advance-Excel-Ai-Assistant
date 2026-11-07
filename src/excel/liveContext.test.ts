import { describe, expect, it } from "vitest";
import { formatLiveContext } from "./liveContext";
import type { SheetOverview } from "./sheetOverview";

const two: SheetOverview = {
  sheets: [
    { name: "Sales", usedRange: "Sales!A1:B3", headers: ["Product", "Qty"] },
    { name: "Prices", usedRange: "Prices!A1:B3", headers: ["Product", "Price"] },
  ],
  omittedSheets: 0,
};

describe("formatLiveContext", () => {
  it("omits the sheet list for single-sheet workbooks and when unavailable", () => {
    const one: SheetOverview = { sheets: [two.sheets[0]!], omittedSheets: 0 };
    expect(formatLiveContext(null, one)).not.toMatch(/Sheets in this workbook/);
    expect(formatLiveContext(null)).not.toMatch(/Sheets in this workbook/);
  });

  it("lists every sheet when there are several", () => {
    const text = formatLiveContext(null, two);
    expect(text).toMatch(/Sheets in this workbook/);
    expect(text).toContain("Prices!A1:B3");
    expect(text).toContain('"Price"');
  });

  it("mentions sheets that were left out", () => {
    expect(formatLiveContext(null, { ...two, omittedSheets: 4 })).toMatch(/4 more sheet/);
  });
});
