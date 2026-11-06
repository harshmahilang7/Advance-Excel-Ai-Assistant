import { describe, expect, it } from "vitest";
import { parseToolCall, toolDeclarations } from "./tools";

describe("parseToolCall", () => {
  it("accepts a valid write_cells call", () => {
    const r = parseToolCall("write_cells", { address: "Sheet1!D1", data: [["Total"], ["=SUM(B2:B4)"]] });
    expect(r.ok).toBe(true);
  });

  it("accepts quoted sheet names and plain ranges", () => {
    expect(parseToolCall("read_range", { address: "'My Sheet'!A1:C10" }).ok).toBe(true);
    expect(parseToolCall("read_range", { address: "A:A" }).ok).toBe(true);
    expect(parseToolCall("read_range", { address: "$A$1:$B$2" }).ok).toBe(true);
  });

  it("rejects malformed addresses", () => {
    const r = parseToolCall("read_range", { address: "not an address!!" });
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.error).toMatch(/address/i);
  });

  it("rejects non-2D data for write_cells", () => {
    expect(parseToolCall("write_cells", { address: "A1", data: ["x"] }).ok).toBe(false);
    expect(parseToolCall("write_cells", { address: "A1", data: [] }).ok).toBe(false);
  });

  it("rejects oversized writes", () => {
    const big = Array.from({ length: 2001 }, () => [1]);
    expect(parseToolCall("write_cells", { address: "A1", data: big }).ok).toBe(false);
  });

  it("rejects unknown tools with a clear error", () => {
    const r = parseToolCall("delete_workbook", {});
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.error).toMatch(/unknown tool/i);
  });

  it("validates chart type enum and sort options", () => {
    expect(parseToolCall("create_chart", { sourceAddress: "A1:B5", chartType: "Pie3000" }).ok).toBe(false);
    expect(parseToolCall("create_chart", { sourceAddress: "A1:B5", chartType: "ColumnClustered" }).ok).toBe(true);
    expect(
      parseToolCall("sort_range", { address: "A1:C9", columnIndex: 1, ascending: false, hasHeaders: true }).ok,
    ).toBe(true);
    expect(
      parseToolCall("sort_range", { address: "A1:C9", columnIndex: -1, ascending: true, hasHeaders: true }).ok,
    ).toBe(false);
  });

  it("rejects bad colours", () => {
    expect(parseToolCall("format_range", { address: "A1", fillColor: "red; drop" }).ok).toBe(false);
    expect(parseToolCall("format_range", { address: "A1", fillColor: "#FFEE00", bold: true }).ok).toBe(true);
  });
});

describe("create_pivot_table", () => {
  const valid = {
    sourceAddress: "Sales!A1:D100",
    rows: ["Region"],
    values: [{ field: "Revenue", summarizeBy: "sum" }],
  };

  it("accepts a minimal pivot with defaults", () => {
    const r = parseToolCall("create_pivot_table", valid);
    expect(r.ok).toBe(true);
    if (r.ok && r.call.name === "create_pivot_table") {
      expect(r.call.args.columns).toEqual([]);
      expect(r.call.args.filters).toEqual([]);
    }
  });

  it("accepts destination, columns and filters", () => {
    expect(
      parseToolCall("create_pivot_table", {
        ...valid,
        destinationAddress: "Pivot!A3",
        columns: ["Product"],
        filters: ["Year"],
        name: "SalesPivot",
      }).ok,
    ).toBe(true);
  });

  it("rejects unknown aggregation functions", () => {
    expect(parseToolCall("create_pivot_table", { ...valid, values: [{ field: "Revenue", summarizeBy: "median" }] }).ok).toBe(false);
  });

  it("rejects empty values and bad names", () => {
    expect(parseToolCall("create_pivot_table", { ...valid, values: [] }).ok).toBe(false);
    expect(parseToolCall("create_pivot_table", { ...valid, name: "bad name!" }).ok).toBe(false);
  });
});

describe("conditional_format", () => {
  it("accepts a cellValue rule with a fill", () => {
    expect(
      parseToolCall("conditional_format", {
        address: "B2:B50",
        rule: "cellValue",
        operator: "GreaterThan",
        value1: 100,
        fillColor: "#FFC7CE",
      }).ok,
    ).toBe(true);
  });

  it("requires value2 for Between and an operator for cellValue", () => {
    const base = { address: "B2:B50", rule: "cellValue", value1: 1, fillColor: "#FFC7CE" };
    expect(parseToolCall("conditional_format", { ...base, operator: "Between" }).ok).toBe(false);
    expect(parseToolCall("conditional_format", { ...base, operator: "Between", value2: 5 }).ok).toBe(true);
    expect(parseToolCall("conditional_format", base).ok).toBe(false);
  });

  it("requires some formatting for cellValue and text rules", () => {
    expect(
      parseToolCall("conditional_format", { address: "A1:A9", rule: "cellValue", operator: "EqualTo", value1: 1 }).ok,
    ).toBe(false);
    expect(parseToolCall("conditional_format", { address: "A1:A9", rule: "text", text: "Late" }).ok).toBe(false);
    expect(parseToolCall("conditional_format", { address: "A1:A9", rule: "text", text: "Late", fontColor: "#9C0006" }).ok).toBe(true);
  });

  it("validates colour scales", () => {
    expect(parseToolCall("conditional_format", { address: "A1:A9", rule: "colorScale", minColor: "#FFFFFF" }).ok).toBe(false);
    expect(
      parseToolCall("conditional_format", { address: "A1:A9", rule: "colorScale", minColor: "#FFFFFF", maxColor: "#63BE7B" }).ok,
    ).toBe(true);
  });

  it("rejects unknown rules and bad colours", () => {
    expect(parseToolCall("conditional_format", { address: "A1", rule: "iconSet" }).ok).toBe(false);
    expect(
      parseToolCall("conditional_format", { address: "A1", rule: "text", text: "x", fillColor: "red" }).ok,
    ).toBe(false);
  });
});

describe("find_replace", () => {
  it("accepts a minimal call and applies defaults", () => {
    const r = parseToolCall("find_replace", { find: "N/A", replace: "" });
    expect(r.ok).toBe(true);
    if (r.ok && r.call.name === "find_replace") {
      expect(r.call.args.matchCase).toBe(false);
      expect(r.call.args.matchEntireCell).toBe(false);
      expect(r.call.args.address).toBeUndefined();
    }
  });

  it("accepts a scoped address and options", () => {
    expect(
      parseToolCall("find_replace", { address: "Sheet1!B:B", find: "USA", replace: "United States", matchCase: true, matchEntireCell: true }).ok,
    ).toBe(true);
  });

  it("rejects an empty find string and bad addresses", () => {
    expect(parseToolCall("find_replace", { find: "", replace: "x" }).ok).toBe(false);
    expect(parseToolCall("find_replace", { address: "nope!!", find: "a", replace: "b" }).ok).toBe(false);
  });
});

describe("toolDeclarations", () => {
  it("exposes every tool with name, description and object schema", () => {
    expect(toolDeclarations.length).toBeGreaterThanOrEqual(8);
    for (const d of toolDeclarations) {
      expect(d.name).toMatch(/^[a-z_]+$/);
      expect(d.description.length).toBeGreaterThan(10);
      expect((d.parametersJsonSchema as { type?: string }).type).toBe("object");
    }
  });

  it("has unique names that all parse as known tools", () => {
    const names = toolDeclarations.map((d) => d.name);
    expect(new Set(names).size).toBe(names.length);
  });
});
