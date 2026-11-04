import { parseToolCall, type ToolArgs, type ToolName } from "../agent/tools";
import type { ToolExecutor, ToolResult } from "../agent/loop";
import { splitSheetAddress } from "./address";
import { resolvePivotFields, type PivotAggregation } from "./pivotFields";
import type { UndoLog } from "./undo";

const DEFAULT_READ_ROWS = 100;

// String literals (not Excel.AggregationFunction.*) so nothing touches the Excel global at import time.
const AGGREGATIONS: Record<PivotAggregation, "Sum" | "Count" | "Average" | "Max" | "Min"> = {
  sum: "Sum",
  count: "Count",
  average: "Average",
  max: "Max",
  min: "Min",
};

type Ctx = Excel.RequestContext;

const resolveRange = (ctx: Ctx, address: string): Excel.Range => {
  const { sheet, cells } = splitSheetAddress(address);
  const ws = sheet ? ctx.workbook.worksheets.getItem(sheet) : ctx.workbook.worksheets.getActiveWorksheet();
  return ws.getRange(cells);
};

const ok = (summary: string, data: unknown = null): ToolResult => ({ ok: true, summary, data });

/** Remember the current formulas of a range so the command can be undone. */
const snapshotFormulas = async (ctx: Ctx, range: Excel.Range, undo: UndoLog): Promise<string> => {
  range.load(["address", "formulas"]);
  await ctx.sync();
  const fullAddress = range.address;
  const saved = range.formulas.map((row) => [...row]);
  undo.record(`restore ${fullAddress}`, () =>
    Excel.run(async (c) => {
      resolveRange(c, fullAddress).formulas = saved;
      await c.sync();
    }),
  );
  return fullAddress;
};

type Handlers = { [N in ToolName]: (ctx: Ctx, args: ToolArgs<N>, undo: UndoLog) => Promise<ToolResult> };

const handlers: Handlers = {
  async get_workbook_info(ctx) {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    await ctx.sync();
    const used = sheets.items.map((ws) => {
      const r = ws.getUsedRangeOrNullObject(true);
      r.load("address");
      return { ws, r };
    });
    await ctx.sync();
    const info = used.map(({ ws, r }) => ({ name: ws.name, usedRange: r.isNullObject ? null : r.address }));
    return ok(`${info.length} sheet(s)`, info);
  },

  async read_range(ctx, { address, maxRows }) {
    const full = resolveRange(ctx, address);
    full.load(["rowCount", "columnCount"]);
    await ctx.sync();
    const limit = Math.min(maxRows ?? DEFAULT_READ_ROWS, full.rowCount);
    const range = full.getRow(0).getResizedRange(limit - 1, 0);
    range.load(["address", "values", "formulas"]);
    await ctx.sync();
    const formulas = range.formulas.map((row) => row.map((f) => (typeof f === "string" && f.startsWith("=") ? f : null)));
    return ok(`Read ${range.address} (${limit} of ${full.rowCount} rows)`, {
      address: range.address,
      totalRows: full.rowCount,
      columns: full.columnCount,
      values: range.values,
      formulas,
    });
  },

  async write_cells(ctx, { address, data }, undo) {
    const rows = data.length;
    const cols = data[0]?.length ?? 1;
    const target = resolveRange(ctx, address).getCell(0, 0).getResizedRange(rows - 1, cols - 1);
    const where = await snapshotFormulas(ctx, target, undo);
    target.formulas = data.map((row) => row.map((c) => c ?? ""));
    await ctx.sync();
    return ok(`Wrote ${rows}x${cols} cells at ${where}`);
  },

  async format_range(ctx, a) {
    const range = resolveRange(ctx, a.address);
    range.load(["rowCount", "columnCount", "address"]);
    await ctx.sync();
    if (a.bold !== undefined) range.format.font.bold = a.bold;
    if (a.italic !== undefined) range.format.font.italic = a.italic;
    if (a.fontColor) range.format.font.color = a.fontColor;
    if (a.fillColor) range.format.fill.color = a.fillColor;
    if (a.horizontalAlignment) range.format.horizontalAlignment = a.horizontalAlignment;
    if (a.wrapText !== undefined) range.format.wrapText = a.wrapText;
    if (a.numberFormat) {
      range.numberFormat = Array.from({ length: range.rowCount }, () => Array(range.columnCount).fill(a.numberFormat));
    }
    await ctx.sync();
    return ok(`Formatted ${range.address}`);
  },

  async sort_range(ctx, { address, columnIndex, ascending, hasHeaders }, undo) {
    const range = resolveRange(ctx, address);
    const where = await snapshotFormulas(ctx, range, undo);
    range.sort.apply([{ key: columnIndex, ascending }], false, hasHeaders);
    await ctx.sync();
    return ok(`Sorted ${where} by column ${columnIndex + 1} ${ascending ? "ascending" : "descending"}`);
  },

  async clear_range(ctx, { address, what }, undo) {
    const range = resolveRange(ctx, address);
    const where = await snapshotFormulas(ctx, range, undo);
    const applyTo = { contents: Excel.ClearApplyTo.contents, formats: Excel.ClearApplyTo.formats, all: Excel.ClearApplyTo.all }[what];
    range.clear(applyTo);
    await ctx.sync();
    return ok(`Cleared ${what} of ${where}`);
  },

  async create_table(ctx, { address, hasHeaders, name }, undo) {
    const range = resolveRange(ctx, address);
    range.load("address");
    await ctx.sync();
    const sheetName = splitSheetAddress(range.address).sheet;
    const table = range.worksheet.tables.add(range, hasHeaders);
    if (name) table.name = name;
    table.load("name");
    await ctx.sync();
    const tableName = table.name;
    undo.record(`remove table ${tableName}`, () =>
      Excel.run(async (c) => {
        c.workbook.tables.getItem(tableName).convertToRange();
        await c.sync();
      }),
    );
    return ok(`Created table ${tableName} on ${sheetName ?? "sheet"} (${range.address})`);
  },

  async create_chart(ctx, { sourceAddress, chartType, title }, undo) {
    const range = resolveRange(ctx, sourceAddress);
    const chart = range.worksheet.charts.add(chartType, range, Excel.ChartSeriesBy.auto);
    if (title) chart.title.text = title;
    chart.load("name");
    range.worksheet.load("name");
    await ctx.sync();
    const sheetName = range.worksheet.name;
    const chartName = chart.name;
    undo.record(`delete chart ${chartName}`, () =>
      Excel.run(async (c) => {
        c.workbook.worksheets.getItem(sheetName).charts.getItem(chartName).delete();
        await c.sync();
      }),
    );
    return ok(`Created ${chartType} chart "${title ?? chartName}"`);
  },

  async create_pivot_table(ctx, args, undo) {
    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.8")) {
      return { ok: false, error: "PivotTables need Excel API 1.8 or newer. Update Excel." };
    }
    const source = resolveRange(ctx, args.sourceAddress);
    source.load(["address", "values"]);
    await ctx.sync();
    const headers = (source.values[0] ?? []).map((h) => String(h ?? ""));
    const fields = resolvePivotFields(headers, args);
    if (!fields.ok) return { ok: false, error: fields.error };

    const name = args.name ?? `Pivot_${Date.now().toString(36).slice(-5)}`;
    let destination: Excel.Range;
    let newSheetName: string | null = null;
    if (args.destinationAddress) {
      destination = resolveRange(ctx, args.destinationAddress).getCell(0, 0);
    } else {
      newSheetName = `Pivot ${Date.now().toString(36).slice(-4)}`;
      const sheet = ctx.workbook.worksheets.add(newSheetName);
      destination = sheet.getRange("A3");
      undo.record(`delete sheet ${newSheetName}`, () =>
        Excel.run(async (c) => {
          c.workbook.worksheets.getItem(newSheetName as string).delete();
          await c.sync();
        }),
      );
    }

    const pivot = destination.worksheet.pivotTables.add(name, source, destination);
    pivot.worksheet.load("name");
    await ctx.sync();
    const sheetName = pivot.worksheet.name;
    if (newSheetName === null) {
      undo.record(`delete pivot ${name}`, () =>
        Excel.run(async (c) => {
          c.workbook.worksheets.getItem(sheetName).pivotTables.getItem(name).delete();
          await c.sync();
        }),
      );
    }

    for (const f of fields.rows) pivot.rowHierarchies.add(pivot.hierarchies.getItem(f));
    for (const f of fields.columns) pivot.columnHierarchies.add(pivot.hierarchies.getItem(f));
    for (const f of fields.filters) pivot.filterHierarchies.add(pivot.hierarchies.getItem(f));
    for (const v of fields.values) {
      const dh = pivot.dataHierarchies.add(pivot.hierarchies.getItem(v.field));
      dh.summarizeBy = AGGREGATIONS[v.summarizeBy];
    }
    await ctx.sync();

    const describe = fields.values.map((v) => `${v.summarizeBy} of ${v.field}`).join(", ");
    return ok(`Created PivotTable "${name}" on ${sheetName}: ${describe}${fields.rows.length ? ` by ${fields.rows.join(", ")}` : ""}`);
  },

  async autofit_columns(ctx, { address }) {
    const range = resolveRange(ctx, address);
    range.format.autofitColumns();
    await ctx.sync();
    return ok(`Auto-fitted columns for ${address}`);
  },
};

/** Real executor: validates, then runs the tool against the open workbook. */
export const createExcelExecutor =
  (undo: UndoLog): ToolExecutor =>
  async (name, args) => {
    const parsed = parseToolCall(name, args);
    if (!parsed.ok) return { ok: false, error: parsed.error };
    const call = parsed.call;
    const handler = handlers[call.name] as (ctx: Ctx, args: unknown, undo: UndoLog) => Promise<ToolResult>;
    return Excel.run((ctx) => handler(ctx, call.args, undo));
  };
