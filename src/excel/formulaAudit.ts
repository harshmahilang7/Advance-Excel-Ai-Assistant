import { columnLetter, type CellValue } from "./context";

export const MAX_FINDINGS_PER_KIND = 50;

/** A column needs this many formulas, and this share on one pattern, before outliers are flagged. */
const MIN_FORMULAS_FOR_PATTERN = 3;
const MIN_DOMINANT_SHARE = 0.6;

const ERROR_VALUE = /^#(REF!|DIV\/0!|N\/A|VALUE!|NAME\?|NULL!|NUM!|SPILL!|CALC!)$/;

export type FindingKind = "error_value" | "inconsistent_formula" | "hardcoded_in_formulas";

export interface AuditFinding {
  readonly kind: FindingKind;
  /** A1 address on the audited sheet, e.g. "C14". */
  readonly address: string;
  readonly detail: string;
}

export interface AuditGrid {
  /** 0-based sheet row/column of the grid's top-left cell. */
  readonly startRow: number;
  readonly startCol: number;
  /** R1C1 formulas; non-formula cells hold their constant value. */
  readonly formulasR1C1: readonly (readonly CellValue[])[];
  readonly values: readonly (readonly CellValue[])[];
}

export interface AuditResult {
  readonly scannedCells: number;
  readonly formulaCount: number;
  readonly findings: readonly AuditFinding[];
  /** Real counts per kind, before the per-kind cap. */
  readonly totals: Readonly<Record<FindingKind, number>>;
  readonly truncated: boolean;
}

const isFormula = (v: CellValue | undefined): v is string => typeof v === "string" && v.startsWith("=");

const addressOf = (grid: AuditGrid, row: number, col: number): string =>
  `${columnLetter(grid.startCol + col)}${grid.startRow + row + 1}`;

const errorFindings = (grid: AuditGrid): AuditFinding[] =>
  grid.values.flatMap((row, r) =>
    row.flatMap((v, c) =>
      typeof v === "string" && ERROR_VALUE.test(v)
        ? [{ kind: "error_value" as const, address: addressOf(grid, r, c), detail: `Cell shows ${v}` }]
        : [],
    ),
  );

interface ColumnFormula {
  readonly row: number;
  readonly formula: string;
}

/** The most common formula in a column and how many cells use it (first one wins a tie). */
const dominantOf = (formulas: readonly ColumnFormula[]): { readonly formula: string; readonly count: number } => {
  const counts = formulas.reduce((m, f) => m.set(f.formula, (m.get(f.formula) ?? 0) + 1), new Map<string, number>());
  const [formula, count] = [...counts.entries()].reduce((best, e) => (e[1] > best[1] ? e : best));
  return { formula, count };
};

const columnFindings = (grid: AuditGrid, col: number): AuditFinding[] => {
  const formulas: ColumnFormula[] = grid.formulasR1C1.flatMap((row, r) => {
    const f = row[col];
    return isFormula(f) ? [{ row: r, formula: f }] : [];
  });
  if (formulas.length < MIN_FORMULAS_FOR_PATTERN) return [];
  const dominant = dominantOf(formulas);
  if (dominant.count < 2 || dominant.count / formulas.length < MIN_DOMINANT_SHARE) return [];

  const inconsistent: AuditFinding[] = formulas
    .filter((f) => f.formula !== dominant.formula)
    .map((f) => ({
      kind: "inconsistent_formula" as const,
      address: addressOf(grid, f.row, col),
      detail: `Differs from the ${dominant.count} other formulas in this column (R1C1 ${f.formula}, expected ${dominant.formula}). May be intentional, e.g. a total.`,
    }));

  const first = formulas[0]?.row ?? 0;
  const last = formulas.at(-1)?.row ?? 0;
  const hardcoded: AuditFinding[] = grid.formulasR1C1.flatMap((row, r) => {
    const v = row[col];
    return r > first && r < last && typeof v === "number"
      ? [{ kind: "hardcoded_in_formulas" as const, address: addressOf(grid, r, col), detail: `Typed number ${v} inside a column of formulas` }]
      : [];
  });

  return [...inconsistent, ...hardcoded];
};

const KINDS: readonly FindingKind[] = ["error_value", "inconsistent_formula", "hardcoded_in_formulas"];

/** Scans a grid for error values, formulas that break their column's pattern, and typed-over numbers. Pure. */
export const auditFormulas = (grid: AuditGrid): AuditResult => {
  const cols = grid.formulasR1C1[0]?.length ?? 0;
  const all = [...errorFindings(grid), ...Array.from({ length: cols }, (_, c) => columnFindings(grid, c)).flat()];

  const byKind = (kind: FindingKind) => all.filter((f) => f.kind === kind);
  const totals = Object.fromEntries(KINDS.map((k) => [k, byKind(k).length])) as Record<FindingKind, number>;
  const findings = KINDS.flatMap((k) => byKind(k).slice(0, MAX_FINDINGS_PER_KIND));

  return {
    scannedCells: grid.formulasR1C1.reduce((n, row) => n + row.length, 0),
    formulaCount: grid.formulasR1C1.reduce((n, row) => n + row.filter(isFormula).length, 0),
    findings,
    totals,
    truncated: KINDS.some((k) => totals[k] > MAX_FINDINGS_PER_KIND),
  };
};
