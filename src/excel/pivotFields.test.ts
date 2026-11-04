import { describe, expect, it } from "vitest";
import { resolvePivotFields } from "./pivotFields";

const headers = ["Region", "Product", "Units", "Revenue"];

describe("resolvePivotFields", () => {
  it("maps requested names to exact header text, case-insensitively", () => {
    const r = resolvePivotFields(headers, {
      rows: ["region"],
      columns: ["PRODUCT"],
      filters: [],
      values: [{ field: "revenue", summarizeBy: "sum" }],
    });
    expect(r).toEqual({
      ok: true,
      rows: ["Region"],
      columns: ["Product"],
      filters: [],
      values: [{ field: "Revenue", summarizeBy: "sum" }],
    });
  });

  it("trims whitespace in requested names", () => {
    const r = resolvePivotFields(headers, { rows: [" Region "], columns: [], filters: [], values: [{ field: "Units", summarizeBy: "count" }] });
    expect(r.ok).toBe(true);
  });

  it("lists available headers when a field is missing", () => {
    const r = resolvePivotFields(headers, { rows: ["Country"], columns: [], filters: [], values: [{ field: "Units", summarizeBy: "sum" }] });
    expect(r.ok).toBe(false);
    if (!r.ok) {
      expect(r.error).toContain('"Country"');
      expect(r.error).toContain("Region, Product, Units, Revenue");
    }
  });

  it("requires at least one value field", () => {
    const r = resolvePivotFields(headers, { rows: ["Region"], columns: [], filters: [], values: [] });
    expect(r.ok).toBe(false);
  });

  it("requires at least one row or column field", () => {
    const r = resolvePivotFields(headers, { rows: [], columns: [], filters: [], values: [{ field: "Units", summarizeBy: "sum" }] });
    expect(r.ok).toBe(false);
  });

  it("rejects ambiguous duplicate headers", () => {
    const r = resolvePivotFields(["Name", "Name", "Qty"], { rows: ["Name"], columns: [], filters: [], values: [{ field: "Qty", summarizeBy: "sum" }] });
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.error).toMatch(/duplicate/i);
  });

  it("rejects a pivot source with blank headers", () => {
    const r = resolvePivotFields(["Region", "", "Qty"], { rows: ["Region"], columns: [], filters: [], values: [{ field: "Qty", summarizeBy: "sum" }] });
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.error).toMatch(/blank header/i);
  });
});
