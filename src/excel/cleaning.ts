/** Pure helpers for cleaning a grid of cell values (no Excel globals). */

export const CLEAN_OPERATIONS = [
  "strip_nonprinting",
  "trim",
  "fill_down",
  "text_to_number",
  "upper",
  "lower",
  "proper",
] as const;

export type CleanOperation = (typeof CLEAN_OPERATIONS)[number];
export type CleanCell = string | number | boolean | null;

export interface CleanResult {
  readonly values: readonly (readonly CleanCell[])[];
  /** Cells whose final value differs from the input. */
  readonly changed: number;
  /** Cells each operation changed when it ran. */
  readonly perOperation: Readonly<Record<CleanOperation, number>>;
}

// C0/C1 control chars (except tab/newline handled by trim), zero-width and BOM characters.
const NONPRINTING = /[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f-\u009f​-‍⁠﻿]/g;
const GROUPED_NUMBER = /^-?\d{1,3}(,\d{3})+(\.\d+)?$/;
const PLAIN_NUMBER = /^-?(\d+(\.\d+)?|\.\d+)$/;

const isFormula = (v: CleanCell): boolean => typeof v === "string" && v.startsWith("=");

/** "1,234.50", "$12", "(45)" and "12%" become numbers. Anything ambiguous returns null. */
export const parseLooseNumber = (input: string): number | null => {
  let s = input.replace(/[\s ]/g, "").replace(/^[$€£¥]/, "");
  let negative = false;
  if (/^\(.*\)$/.test(s)) {
    negative = true;
    s = s.slice(1, -1);
  }
  let percent = false;
  if (s.endsWith("%")) {
    percent = true;
    s = s.slice(0, -1);
  }
  if (!GROUPED_NUMBER.test(s) && !PLAIN_NUMBER.test(s)) return null;
  const n = Number(s.replaceAll(",", ""));
  if (!Number.isFinite(n)) return null;
  const signed = negative ? -n : n;
  return percent ? signed / 100 : signed;
};

const toProper = (s: string): string => s.toLowerCase().replace(/(^|\s)(\S)/g, (_m, sp: string, ch: string) => sp + ch.toUpperCase());

const textOps: Partial<Record<CleanOperation, (s: string) => string>> = {
  strip_nonprinting: (s) => s.replace(NONPRINTING, ""),
  trim: (s) => s.replace(/[\s ]+/g, " ").trim(),
  upper: (s) => s.toUpperCase(),
  lower: (s) => s.toLowerCase(),
  proper: toProper,
};

type Grid = readonly (readonly CleanCell[])[];

const mapText = (grid: Grid, fn: (s: string) => string): { grid: Grid; changed: number } => {
  let changed = 0;
  const next = grid.map((row) =>
    row.map((v) => {
      if (typeof v !== "string" || isFormula(v)) return v;
      const out = fn(v);
      if (out !== v) changed += 1;
      return out;
    }),
  );
  return { grid: next, changed };
};

const toNumbers = (grid: Grid): { grid: Grid; changed: number } => {
  let changed = 0;
  const next = grid.map((row) =>
    row.map((v) => {
      if (typeof v !== "string" || isFormula(v)) return v;
      const n = parseLooseNumber(v);
      if (n === null) return v;
      changed += 1;
      return n;
    }),
  );
  return { grid: next, changed };
};

const fillDown = (grid: Grid): { grid: Grid; changed: number } => {
  let changed = 0;
  const width = grid.reduce((m, r) => Math.max(m, r.length), 0);
  const last: (CleanCell | undefined)[] = Array.from({ length: width }, () => undefined);
  const next = grid.map((row) =>
    row.map((v, c) => {
      if (v === "" || v === null) {
        const above = last[c];
        if (above === undefined) return v;
        changed += 1;
        return above;
      }
      if (!isFormula(v)) last[c] = v;
      return v;
    }),
  );
  return { grid: next, changed };
};

/**
 * Applies the requested operations in a fixed, safe order (see CLEAN_OPERATIONS) and returns a new grid.
 * Formulas are never modified. The input is not mutated.
 */
export const cleanGrid = (input: Grid, operations: readonly CleanOperation[]): CleanResult => {
  const wanted = new Set(operations);
  const perOperation = Object.fromEntries(CLEAN_OPERATIONS.map((op) => [op, 0])) as Record<CleanOperation, number>;
  let grid: Grid = input.map((row) => [...row]);

  for (const op of CLEAN_OPERATIONS) {
    if (!wanted.has(op)) continue;
    const step =
      op === "fill_down" ? fillDown(grid) : op === "text_to_number" ? toNumbers(grid) : mapText(grid, textOps[op] as (s: string) => string);
    grid = step.grid;
    perOperation[op] = step.changed;
  }

  let changed = 0;
  grid.forEach((row, r) =>
    row.forEach((v, c) => {
      if (v !== input[r]?.[c]) changed += 1;
    }),
  );
  return { values: grid, changed, perOperation };
};
