import { describe, expect, it } from "vitest";
import { splitSheetAddress } from "./address";

describe("splitSheetAddress", () => {
  it("returns null sheet for unqualified addresses", () => {
    expect(splitSheetAddress("A1:B2")).toEqual({ sheet: null, cells: "A1:B2" });
  });

  it("splits a plain sheet name", () => {
    expect(splitSheetAddress("Sheet1!A1")).toEqual({ sheet: "Sheet1", cells: "A1" });
  });

  it("unquotes sheet names with spaces", () => {
    expect(splitSheetAddress("'My Sheet'!A1:C3")).toEqual({ sheet: "My Sheet", cells: "A1:C3" });
  });

  it("unescapes doubled apostrophes", () => {
    expect(splitSheetAddress("'Bob''s Data'!B2")).toEqual({ sheet: "Bob's Data", cells: "B2" });
  });

  it("handles whole-column references", () => {
    expect(splitSheetAddress("Sales!A:A")).toEqual({ sheet: "Sales", cells: "A:A" });
  });
});
