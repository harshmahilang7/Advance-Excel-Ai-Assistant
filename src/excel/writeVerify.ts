/** Pure read-back checks for write_cells: did Excel store what we wrote, and do any written cells show errors? */

export const MAX_REPORTED = 3;

const ERROR_VALUE = /^#(REF!|DIV\/0!|N\/A|NAME\?|VALUE!|NULL!|NUM!|SPILL!|CALC!|GETTING_DATA)/;

export interface DifferedCell {
  readonly cell: string;
  readonly wrote: string;
  readonly got: string;
}
export interface ErrorCell {
  readonly cell: string;
  readonly value: string;
}
export interface Verification {
  readonly total: number;
  readonly matched: number;
  readonly differed: readonly DifferedCell[];
  readonly errors: readonly ErrorCell[];
}

type Grid = readonly (readonly unknown[])[];

const text = (v: unknown): string => (v === null || v === undefined ? "" : String(v));
const normalize = (v: unknown): string => {
  const s = text(v).trim();
  return s.startsWith("=") ? s.replace(/\s+/g, "").toUpperCase() : s;
};
const sameNumber = (a: string, b: string): boolean => a !== "" && b !== "" && Number.isFinite(Number(a)) && Number(a) === Number(b);

const columnLetters = (index: number): string => {
  let n = index;
  let out = "";
  do {
    out = String.fromCharCode(65 + (n % 26)) + out;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return out;
};

const cellName = (row: number, col: number): string => `${columnLetters(col)}${row + 1}`;

/**
 * Compare what was written with what Excel reports back. `origin` is the zero-based position of the grid's
 * top-left cell, used only to name cells. Never mutates its inputs.
 */
export const verifyWrite = (
  wrote: Grid,
  readFormulas: Grid,
  readValues: Grid,
  origin: { readonly row: number; readonly col: number } = { row: 0, col: 0 },
): Verification => {
  let total = 0;
  let matched = 0;
  const differed: DifferedCell[] = [];
  const errors: ErrorCell[] = [];
  wrote.forEach((rowCells, r) => {
    rowCells.forEach((expected, c) => {
      total += 1;
      const cell = cellName(origin.row + r, origin.col + c);
      const got = readFormulas[r]?.[c];
      const want = normalize(expected);
      const have = normalize(got);
      if (want === have || sameNumber(want, have)) matched += 1;
      else differed.push({ cell, wrote: text(expected), got: text(got) });
      const shown = text(readValues[r]?.[c]);
      if (ERROR_VALUE.test(shown)) errors.push({ cell, value: shown });
    });
  });
  return { total, matched, differed, errors };
};

const moreTail = (count: number): string => (count > MAX_REPORTED ? `, and ${count - MAX_REPORTED} more` : "");

/** Suffix for the step line: " (verified n/n cells)" or one or two ⚠ notes. */
export const describeVerification = (v: Verification): string => {
  const parts: string[] = [];
  if (v.differed.length > 0) {
    const shown = v.differed
      .slice(0, MAX_REPORTED)
      .map((d) => `${d.cell}: wrote "${d.wrote}", Excel has "${d.got}"`)
      .join("; ");
    parts.push(` ⚠ ${v.differed.length} of ${v.total} cells read back differently (${shown}${moreTail(v.differed.length)})`);
  }
  if (v.errors.length > 0) {
    const shown = v.errors
      .slice(0, MAX_REPORTED)
      .map((e) => `${e.cell}: ${e.value}`)
      .join("; ");
    const noun = v.errors.length === 1 ? "cell shows an error" : "cells show errors";
    parts.push(` ⚠ ${v.errors.length} ${noun} (${shown}${moreTail(v.errors.length)})`);
  }
  return parts.length > 0 ? parts.join("") : ` (verified ${v.matched}/${v.total} cells)`;
};
