import type { UndoLog } from "./undo";
import { keepVisibleRows } from "./visibility";
import { loadRowHidden, MAX_SCAN_ROWS } from "./visibleRange";

const MAX_COPY_CELLS = 500_000;

export type PivotSource =
  | { readonly ok: true; readonly range: Excel.Range; readonly headers: readonly string[]; readonly note: string }
  | { readonly ok: false; readonly error: string };

/**
 * Resolves the range a PivotTable should read. PivotTables ignore worksheet filters, so when rows are
 * hidden by a filter (and the caller wants visible rows only) the visible rows are copied to a new
 * sheet and the pivot reads that copy. The copy is a static snapshot.
 */
export const preparePivotSource = async (
  ctx: Excel.RequestContext,
  requested: Excel.Range,
  visibleOnly: boolean,
  suffix: string,
  undo: UndoLog,
): Promise<PivotSource> => {
  const used = requested.getUsedRangeOrNullObject(true);
  used.load(["rowCount", "columnCount"]);
  await ctx.sync();
  if (used.isNullObject) return { ok: false, error: "The pivot source range is empty." };

  const header = used.getRow(0);
  header.load("values");
  await ctx.sync();
  const headers = (header.values[0] ?? []).map((h) => String(h ?? ""));

  const hidden = visibleOnly ? await loadRowHidden(ctx, used, used.rowCount) : [];
  if (!hidden.some((isHidden, i) => i > 0 && isHidden)) return { ok: true, range: used, headers, note: "" };

  used.load(["values", "numberFormat"]);
  await ctx.sync();
  const paired = used.values.map((values, i) => ({ values, formats: used.numberFormat[i] ?? [] }));
  const kept = keepVisibleRows(paired, hidden);
  if (kept.rows.length * used.columnCount > MAX_COPY_CELLS) {
    return { ok: false, error: "The visible rows are too many to copy for a PivotTable (limit 500,000 cells)." };
  }

  const dataSheetName = `Pivot data ${suffix}`;
  const sheet = ctx.workbook.worksheets.add(dataSheetName);
  undo.record(`delete sheet ${dataSheetName}`, () =>
    Excel.run(async (c) => {
      c.workbook.worksheets.getItem(dataSheetName).delete();
      await c.sync();
    }),
  );
  const target = sheet.getRangeByIndexes(0, 0, kept.rows.length, used.columnCount);
  target.numberFormat = kept.rows.map((r) => r.formats);
  target.values = kept.rows.map((r) => r.values);
  await ctx.sync();

  const checked = used.rowCount > MAX_SCAN_ROWS ? ` (only the first ${MAX_SCAN_ROWS} rows were checked for filtering)` : "";
  const note = `Used only the ${kept.rows.length - 1} visible rows (${kept.hiddenCount} filtered-out rows skipped) from a snapshot copy on "${dataSheetName}"${checked}. It does not update when the filter or data changes.`;
  return { ok: true, range: target, headers, note };
};
