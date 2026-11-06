import { describe, expect, it } from "vitest";
import type { CellValue } from "./context";
import { auditFormulas, MAX_FINDINGS_PER_KIND } from "./formulaAudit";

const grid = (formulasR1C1: CellValue[][], values?: CellValue[][], startRow = 0, startCol = 0) => ({
  startRow,
  startCol,
  formulasR1C1,
  values: values ?? formulasR1C1.map((r) => r.map((c) => (typeof c === "string" && c.startsWith("=") ? 1 : c))),
});

describe("auditFormulas", () => {
  it("reports nothing for a clean consistent column", () => {
    const g = grid([["=RC[-2]*RC[-1]"], ["=RC[-2]*RC[-1]"], ["=RC[-2]*RC[-1]"], ["=RC[-2]*RC[-1]"]]);
    const r = auditFormulas(g);
    expect(r.findings).toEqual([]);
    expect(r.formulaCount).toBe(4);
  });

  it("flags error values with their A1 address", () => {
    const g = grid([[1, "=RC[-1]"], [2, "=RC[-1]"]], [[1, 1], [2, "#DIV/0!"]], 4, 2);
    const r = auditFormulas(g);
    expect(r.findings).toContainEqual(expect.objectContaining({ kind: "error_value", address: "D6" }));
  });

  it("flags a formula that breaks its column's pattern", () => {
    const same = "=RC[-2]*RC[-1]";
    const g = grid([[same], [same], ["=RC[-2]+RC[-1]"], [same], [same]]);
    const r = auditFormulas(g);
    const hit = r.findings.find((f) => f.kind === "inconsistent_formula");
    expect(hit?.address).toBe("A3");
    expect(r.findings.filter((f) => f.kind === "inconsistent_formula")).toHaveLength(1);
  });

  it("flags a hard-coded number inside a column of formulas", () => {
    const f = "=RC[-1]*2";
    const g = grid([[f], [f], [99], [f], [f]]);
    const r = auditFormulas(g);
    expect(r.findings).toContainEqual(expect.objectContaining({ kind: "hardcoded_in_formulas", address: "A3" }));
  });

  it("does not flag columns with too few formulas to establish a pattern", () => {
    const g = grid([["=RC[-1]"], ["=RC[-1]+1"]]);
    expect(auditFormulas(g).findings).toEqual([]);
  });

  it("does not flag a column with no dominant pattern", () => {
    const g = grid([["=R1C1"], ["=R2C1"], ["=R3C1"], ["=R4C1"]]);
    expect(auditFormulas(g).findings.filter((f) => f.kind === "inconsistent_formula")).toEqual([]);
  });

  it("ignores blanks and text around the formulas", () => {
    const f = "=RC[-1]";
    const g = grid([["Header"], [f], [f], [f], [null]]);
    expect(auditFormulas(g).findings).toEqual([]);
  });

  it("caps findings per kind and says it truncated", () => {
    const rows: CellValue[][] = Array.from({ length: MAX_FINDINGS_PER_KIND + 20 }, () => [1, "=RC[-1]"]);
    const values: CellValue[][] = rows.map(() => [1, "#REF!"]);
    const r = auditFormulas(grid(rows, values));
    expect(r.findings.filter((f) => f.kind === "error_value")).toHaveLength(MAX_FINDINGS_PER_KIND);
    expect(r.truncated).toBe(true);
    expect(r.totals.error_value).toBe(MAX_FINDINGS_PER_KIND + 20);
  });

  it("does not mutate its input", () => {
    const g = grid([["=RC[-1]"], ["=RC[-1]"], ["=RC[-1]+1"]]);
    const before = JSON.stringify(g);
    auditFormulas(g);
    expect(JSON.stringify(g)).toBe(before);
  });
});
