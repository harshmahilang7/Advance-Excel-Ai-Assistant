import { describe, expect, it } from "vitest";
import { cleanNote, localCell, MAX_NOTE_CHARS } from "./sourceNote";

describe("localCell", () => {
  it("strips the sheet name", () => {
    expect(localCell("Sheet1!D2")).toBe("D2");
  });

  it("handles quoted sheet names containing spaces and bangs", () => {
    expect(localCell("'Q1 Sales'!B7")).toBe("B7");
    expect(localCell("'Wow!'!C3")).toBe("C3");
  });

  it("strips absolute markers and keeps only the first cell of a range", () => {
    expect(localCell("Sheet1!$A$1:$B$2")).toBe("A1");
  });

  it("accepts an address without a sheet", () => {
    expect(localCell("E5")).toBe("E5");
  });

  it("returns null for anything that is not a cell address", () => {
    expect(localCell("")).toBeNull();
    expect(localCell("Sheet1!")).toBeNull();
    expect(localCell("Sales")).toBeNull();
  });
});

describe("cleanNote", () => {
  it("trims whitespace", () => {
    expect(cleanNote("  Total = B2 x C2  ")).toBe("Total = B2 x C2");
  });

  it("returns null for empty or non-string input", () => {
    expect(cleanNote("   ")).toBeNull();
    expect(cleanNote(undefined)).toBeNull();
  });

  it("clips long notes with an ellipsis", () => {
    const out = cleanNote("x".repeat(MAX_NOTE_CHARS + 50));
    expect(out).toHaveLength(MAX_NOTE_CHARS);
    expect(out?.endsWith("…")).toBe(true);
  });
});
