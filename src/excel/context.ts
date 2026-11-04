export type CellValue = string | number | boolean | null;
export type ColumnType = "text" | "number" | "boolean" | "mixed" | "empty";

export interface RangeSnapshot {
  readonly sheetName: string;
  readonly address: string;
  readonly values: readonly (readonly CellValue[])[];
}

export interface SheetSummary {
  readonly sheetName: string;
  readonly address: string;
  readonly columnCount: number;
  readonly dataRowCount: number;
  readonly headers: readonly string[];
  readonly columnTypes: readonly ColumnType[];
  readonly sampleRows: readonly (readonly CellValue[])[];
  readonly truncated: boolean;
}

interface SummaryOptions {
  readonly maxSampleRows?: number;
}

const DEFAULT_SAMPLE_ROWS = 5;

const isBlank = (v: CellValue): boolean => v === null || v === "";

/** 0 -> "A", 25 -> "Z", 26 -> "AA" */
const columnLetter = (index: number): string => {
  let n = index;
  let out = "";
  do {
    out = String.fromCharCode(65 + (n % 26)) + out;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return out;
};

const typeOf = (v: CellValue): Exclude<ColumnType, "mixed" | "empty"> => {
  if (typeof v === "number") return "number";
  if (typeof v === "boolean") return "boolean";
  return "text";
};

const inferColumnType = (column: readonly CellValue[]): ColumnType => {
  const kinds = new Set(column.filter((v) => !isBlank(v)).map(typeOf));
  if (kinds.size === 0) return "empty";
  if (kinds.size > 1) return "mixed";
  return [...kinds][0] as ColumnType;
};

/**
 * Compact description of a range for the LLM: headers, column types, a few sample rows.
 * Never sends the whole sheet. Pure; does not mutate its input.
 */
export const summarizeRange = (
  snapshot: RangeSnapshot,
  options: SummaryOptions = {},
): SheetSummary => {
  const maxSampleRows = options.maxSampleRows ?? DEFAULT_SAMPLE_ROWS;
  const [headerRow, ...dataRows] = snapshot.values;

  if (!headerRow) {
    return {
      sheetName: snapshot.sheetName,
      address: snapshot.address,
      columnCount: 0,
      dataRowCount: 0,
      headers: [],
      columnTypes: [],
      sampleRows: [],
      truncated: false,
    };
  }

  const headers = headerRow.map((cell, i) =>
    isBlank(cell) ? `Column ${columnLetter(i)}` : String(cell),
  );
  const columnTypes = headers.map((_, i) => inferColumnType(dataRows.map((r) => r[i] ?? null)));

  return {
    sheetName: snapshot.sheetName,
    address: snapshot.address,
    columnCount: headers.length,
    dataRowCount: dataRows.length,
    headers,
    columnTypes,
    sampleRows: dataRows.slice(0, maxSampleRows).map((r) => [...r]),
    truncated: dataRows.length > maxSampleRows,
  };
};
