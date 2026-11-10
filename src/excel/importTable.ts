import { safeCell } from "../lib/csv";
import type { UndoLog } from "./undo";

const CHUNK_ROWS = 2000;
const MAX_SHEET_NAME = 31;
const BAD_SHEET_CHARS = /[\\/?*[\]:]/g;

export interface ImportOutcome {
  readonly sheet: string;
  readonly rows: number;
  readonly columns: number;
}

/** A sheet name Excel accepts, based on the file name ("sales 2024.csv" -> "sales 2024"). */
export const sheetNameFor = (fileName: string, taken: readonly string[]): string => {
  const base = fileName.replace(/\.[^.]*$/, "").replace(BAD_SHEET_CHARS, " ").trim() || "Import";
  const lower = new Set(taken.map((t) => t.toLowerCase()));
  let candidate = base.slice(0, MAX_SHEET_NAME);
  for (let i = 2; lower.has(candidate.toLowerCase()); i++) {
    const suffix = ` (${i})`;
    candidate = base.slice(0, MAX_SHEET_NAME - suffix.length) + suffix;
  }
  return candidate;
};

/**
 * Write parsed rows into a new sheet: header row bold and frozen, columns fitted. Never touches existing sheets.
 * Registers an undo step that deletes the new sheet.
 */
export const importTable = async (
  rows: readonly (readonly string[])[],
  fileName: string,
  undo: UndoLog,
): Promise<ImportOutcome> => {
  const columns = rows[0]?.length ?? 0;
  if (rows.length === 0 || columns === 0) throw new Error("There is no data to import.");

  return Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    await ctx.sync();
    const name = sheetNameFor(
      fileName,
      sheets.items.map((s) => s.name),
    );
    const sheet = sheets.add(name);
    undo.record(`delete sheet ${name}`, () =>
      Excel.run(async (c) => {
        c.workbook.worksheets.getItem(name).delete();
        await c.sync();
      }),
    );

    for (let start = 0; start < rows.length; start += CHUNK_ROWS) {
      const chunk = rows.slice(start, start + CHUNK_ROWS).map((r) => r.map(safeCell));
      sheet.getRangeByIndexes(start, 0, chunk.length, columns).values = chunk;
      await ctx.sync();
    }

    sheet.getRangeByIndexes(0, 0, 1, columns).format.font.bold = true;
    sheet.freezePanes.freezeRows(1);
    sheet.getRangeByIndexes(0, 0, Math.min(rows.length, 500), columns).format.autofitColumns();
    sheet.activate();
    await ctx.sync();
    return { sheet: name, rows: rows.length, columns };
  });
};
