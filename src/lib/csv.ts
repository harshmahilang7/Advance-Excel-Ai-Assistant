/** Small CSV/TSV parser for importing a file into a new sheet. Pure. */

export const MAX_ROWS = 20000;
export const MAX_COLS = 200;

const DELIMITERS = [",", "\t", ";", "|"] as const;
type Delimiter = (typeof DELIMITERS)[number];

export interface ParsedTable {
  readonly rows: readonly (readonly string[])[];
  readonly delimiter: Delimiter;
  /** True when rows or columns beyond the limits were dropped. */
  readonly truncated: boolean;
}

/** The delimiter that appears most in the first line, ignoring quoted parts. Comma wins ties. */
const detectDelimiter = (text: string): Delimiter => {
  const firstLine = text.split(/\r?\n/, 1)[0] ?? "";
  const unquoted = firstLine.replace(/"[^"]*"/g, "");
  let best: Delimiter = ",";
  let bestCount = 0;
  for (const d of DELIMITERS) {
    const count = unquoted.split(d).length - 1;
    if (count > bestCount) {
      best = d;
      bestCount = count;
    }
  }
  return best;
};

const splitRecords = (text: string, delimiter: Delimiter): string[][] => {
  const records: string[][] = [];
  let record: string[] = [];
  let field = "";
  let inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i] as string;
    if (inQuotes) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i += 1;
        } else inQuotes = false;
      } else field += ch;
    } else if (ch === '"' && field === "") inQuotes = true;
    else if (ch === delimiter) {
      record.push(field);
      field = "";
    } else if (ch === "\n" || ch === "\r") {
      if (ch === "\r" && text[i + 1] === "\n") i += 1;
      record.push(field);
      records.push(record);
      record = [];
      field = "";
    } else field += ch;
  }
  record.push(field);
  records.push(record);
  return records;
};

const isBlank = (record: readonly string[]): boolean => record.every((c) => c === "");

export const parseDelimited = (input: string): ParsedTable => {
  const text = input.replace(/^﻿/, "");
  const delimiter = detectDelimiter(text);
  const records = splitRecords(text, delimiter).filter((r) => !isBlank(r));
  const widest = records.reduce((max, r) => Math.max(max, r.length), 0);
  const width = Math.min(widest, MAX_COLS);
  const kept = records.slice(0, MAX_ROWS);
  const rows = kept.map((r) => Array.from({ length: width }, (_, i) => r[i] ?? ""));
  return { rows, delimiter, truncated: records.length > MAX_ROWS || widest > MAX_COLS };
};

/** Text starting with "=" would become a live formula in Excel; keep imported data as plain text. */
export const safeCell = (value: string): string => (value.startsWith("=") ? "'" + value : value);
