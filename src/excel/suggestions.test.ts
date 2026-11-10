import { describe, expect, it } from "vitest";
import type { SheetSummary } from "./context";
import { DEFAULT_SUGGESTIONS, REVIEW_SUGGESTIONS, suggestPrompts } from "./suggestions";

const summary = (over: Partial<SheetSummary> = {}): SheetSummary => ({
  hiddenRowCount: 0,
  sheetName: "Sheet1",
  address: "A1:C10",
  columnCount: 3,
  dataRowCount: 9,
  headers: ["Region", "Qty", "Price"],
  columnTypes: ["text", "number", "number"],
  sampleRows: [],
  truncated: false,
  ...over,
});

describe("suggestPrompts", () => {
  it("falls back to defaults without a selection", () => {
    expect(suggestPrompts(null)).toEqual(DEFAULT_SUGGESTIONS);
  });

  it("falls back to defaults for an empty selection", () => {
    expect(suggestPrompts(summary({ columnCount: 0, dataRowCount: 0, headers: [], columnTypes: [] }))).toEqual(
      DEFAULT_SUGGESTIONS,
    );
  });

  it("uses real header names for numeric and text columns", () => {
    const out = suggestPrompts(summary());
    expect(out.some((s) => s.includes("Qty"))).toBe(true);
    expect(out.some((s) => s.includes("Region"))).toBe(true);
  });

  it("offers a pivot only when there is a text and a number column", () => {
    const numericOnly = suggestPrompts(summary({ headers: ["A", "B"], columnTypes: ["number", "number"] }));
    expect(numericOnly.some((s) => /pivot/i.test(s))).toBe(false);
    expect(suggestPrompts(summary()).some((s) => /pivot/i.test(s))).toBe(true);
  });

  it("caps suggestions at 5 and never mutates input", () => {
    const input = summary();
    const frozen = JSON.stringify(input);
    expect(suggestPrompts(input).length).toBeLessThanOrEqual(5);
    expect(JSON.stringify(input)).toBe(frozen);
  });
});

describe("REVIEW_SUGGESTIONS", () => {
  it("has a few distinct prompts", () => {
    expect(REVIEW_SUGGESTIONS.length).toBeGreaterThanOrEqual(3);
    expect(new Set(REVIEW_SUGGESTIONS).size).toBe(REVIEW_SUGGESTIONS.length);
  });

  it("every prompt tells the model not to change anything", () => {
    for (const prompt of REVIEW_SUGGESTIONS) expect(prompt).toMatch(/do not change/i);
  });
});

