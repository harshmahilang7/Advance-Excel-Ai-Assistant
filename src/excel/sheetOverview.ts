import { columnLetter, type CellValue } from "./context";

export const MAX_OVERVIEW_SHEETS = 20;
export const MAX_OVERVIEW_HEADERS = 12;
const MAX_HEADER_CHARS = 40;

export interface RawSheet {
  readonly name: string;
  /** Sheet-qualified used range, or null for an empty sheet. */
  readonly usedRange: string | null;
  readonly firstRow: readonly CellValue[];
}

export interface SheetInfo {
  readonly name: string;
  readonly usedRange: string | null;
  readonly headers: readonly string[];
}

export interface SheetOverview {
  readonly sheets: readonly SheetInfo[];
  /** Sheets left out because the workbook has more than MAX_OVERVIEW_SHEETS. */
  readonly omittedSheets: number;
}

const headerText = (cell: CellValue, index: number): string =>
  cell === null || cell === "" ? `Column ${columnLetter(index)}` : String(cell).slice(0, MAX_HEADER_CHARS);

/** Compact, bounded description of every sheet so the model can plan cross-sheet work. Pure. */
export const summarizeSheets = (raw: readonly RawSheet[]): SheetOverview => {
  const sheets = raw.slice(0, MAX_OVERVIEW_SHEETS).map((s) => ({
    name: s.name,
    usedRange: s.usedRange,
    headers: s.firstRow.slice(0, MAX_OVERVIEW_HEADERS).map(headerText),
  }));
  return { sheets, omittedSheets: Math.max(0, raw.length - MAX_OVERVIEW_SHEETS) };
};
