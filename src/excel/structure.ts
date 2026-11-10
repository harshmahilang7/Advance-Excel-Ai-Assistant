/** Pure helpers for the structure tools (insert, delete, copy, merge). No Excel globals. */

/** The most existing cells a delete may remove, because they must all be snapshotted so the delete can be undone. */
export const MAX_DELETE_UNDO_CELLS = 50_000;

export type AddressKind = "rows" | "columns" | "cells";
export type InsertShift = "Down" | "Right";
export type DeleteShift = "Up" | "Left";

const ROWS = /^\$?\d+:\$?\d+$/;
const COLUMNS = /^\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}$/;

const withoutSheet = (address: string): string => address.slice(address.lastIndexOf("!") + 1);

/** Whole rows ("3:5"), whole columns ("C:D") or anything else (cells). Sheet names are ignored. */
export const addressKind = (address: string): AddressKind => {
  const cells = withoutSheet(address);
  if (ROWS.test(cells)) return "rows";
  if (COLUMNS.test(cells)) return "columns";
  return "cells";
};

/** Whole rows always shift down and whole columns right; only cell ranges honour the request. */
export const insertShift = (kind: AddressKind, requested: InsertShift | undefined): InsertShift => {
  if (kind === "rows") return "Down";
  if (kind === "columns") return "Right";
  return requested ?? "Down";
};

export const deleteShift = (kind: AddressKind, requested: DeleteShift | undefined): DeleteShift => {
  if (kind === "rows") return "Up";
  if (kind === "columns") return "Left";
  return requested ?? "Up";
};

export const undoInsertShift = (shift: InsertShift): DeleteShift => (shift === "Down" ? "Up" : "Left");
export const undoDeleteShift = (shift: DeleteShift): InsertShift => (shift === "Up" ? "Down" : "Right");

const isBlank = (v: unknown): boolean => v === "" || v === null || v === undefined;

export const hasAnyData = (values: readonly (readonly unknown[])[]): boolean => values.some((row) => row.some((v) => !isBlank(v)));

/** True when any cell other than the top-left one holds data (merging would discard it). */
export const otherCellsHaveData = (values: readonly (readonly unknown[])[]): boolean =>
  values.some((row, r) => row.some((v, c) => !(r === 0 && c === 0) && !isBlank(v)));
