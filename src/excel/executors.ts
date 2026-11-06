import { parseToolCall, type ToolArgs, type ToolName } from "../agent/tools";
import type { ToolExecutor, ToolResult } from "../agent/loop";
import { splitSheetAddress } from "./address";
import { toRuleFormula } from "./conditionalFormat";
import { canSnapshotFormats, FORMAT_PROPS_REQUEST, toRestoreProps, type CellFormat } from "./formatSnapshot";
import { resolvePivotFields, type PivotAggregation } from "./pivotFields";
import { preparePivotSource } from "./pivotSource";
import type { UndoLog } from "./undo";
import { keepVisibleRows } from "./visibility";
import { loadRowHidden, MAX_SCAN_ROWS } from "./visibleRange";

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

/**
 * Remember cell formatting (font, fill, alignment, wrap, number format) so the command can be undone.
 * Returns false when it cannot (old Excel, or a range too large to snapshot cell by cell).
 */
const snapshotFormats = async (ctx: Ctx, range: Excel.Range, undo: UndoLog): Promise<boolean> => {
  if (!Office.context.requirements.isSetSupported("ExcelApi", "1.9")) return false;
  range.load(["address", "rowCount", "columnCount", "numberFormat"]);
  await ctx.sync();
  if (!canSnapshotFormats(range.rowCount, range.columnCount)) return false;
  const props = range.getCellProperties(FORMAT_PROPS_REQUEST as unknown as Excel.CellPropertiesLoadOptions);
  await ctx.sync();
  const fullAddress = range.address;
  const saved = toRestoreProps(props.value as unknown as CellFormat[][]) as unknown as Excel.CellProperties[][];
  const numberFormat = range.numberFormat.map((row) => [...row]);
  undo.record(`restore formatting ${fullAddress}`, () =>
    Excel.run(async (c) => {
      const r = resolveRange(c, fullAddress);
      r.format.fill.clear();
      r.setCellProperties(saved);
      r.numberFormat = numberFormat;
      await c.sync();
    }),
  );
  return true;
};

const NOT_UNDOABLE = " (formatting too large to undo)";

type Handlers ={ [N in ToolName]: (ctx: Ctx, args: ToolArgs<N>, undo: UndoLog) => Promise<ToolResult> };

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
    const full = resolveRange(ctx, address).getUsedRangeOrNullObject(true);
    full.load(["rowCount", "columnCount"]);
    await ctx.sync();
    if (full.isNullObject) return ok(`${address} is empty`, { address, totalRows: 0, columns: 0, values: [], formulas: [] });

    const scanRows = Math.min(full.rowCount, MAX_SCAN_ROWS);
    const range = full.getRow(0).getResizedRange(scanRows - 1, 0);
    range.load(["address", "values", "formulas"]);
    await ctx.sync();
    const hidden = await loadRowHidden(ctx, range, scanRows);
    const visible = keepVisibleRows(range.values.map((values, i) => ({ values, formulas: range.formulas[i] ?? [] })), hidden);
    const limit = Math.min(maxRows ?? DEFAULT_READ_ROWS, visible.rows.length);
    const shown = visible.rows.slice(0, limit);
    const formulas = shown.map((r) => r.formulas.map((f) => (typeof f === "string" && f.startsWith("=") ? f : null)));
    const skipped = visible.hiddenCount > 0 ? `, ${visible.hiddenCount} filtered-out rows skipped` : "";
    return ok(`Read ${range.address} (${limit} of ${visible.rows.length} visible rows${skipped})`, {
      address: range.address,
      totalRows: full.rowCount,
      hiddenRowsSkipped: visible.hiddenCount,
      columns: full.columnCount,
      values: shown.map((r) => r.values),
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

  async format_range(ctx, a, undo) {
    const range = resolveRange(ctx, a.address);
    const undoable = await snapshotFormats(ctx, range, undo);
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
    return ok(`Formatted ${range.address}${undoable ? "" : NOT_UNDOABLE}`);
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
    const undoable = what === "contents" || (await snapshotFormats(ctx, range, undo));
    const applyTo = { contents: Excel.ClearApplyTo.contents, formats: Excel.ClearApplyTo.formats, all: Excel.ClearApplyTo.all }[what];
    range.clear(applyTo);
    await ctx.sync();
    return ok(`Cleared ${what} of ${where}${undoable ? "" : NOT_UNDOABLE}`);
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
    const suffix = Date.now().toString(36).slice(-4);
    const prepared = await preparePivotSource(ctx, resolveRange(ctx, args.sourceAddress), args.visibleRowsOnly, suffix, undo);
    if (!prepared.ok) return { ok: false, error: prepared.error };
    const source = prepared.range;
    const fields = resolvePivotFields(prepared.headers, args);
    if (!fields.ok) return { ok: false, error: fields.error };

    const name = args.name ?? `Pivot_${Date.now().toString(36).slice(-5)}`;
    let destination: Excel.Range;
    let newSheetName: string | null = null;
    if (args.destinationAddress) {
      destination = resolveRange(ctx, args.destinationAddress).getCell(0, 0);
    } else {
      newSheetName = `Pivot ${suffix}`;
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
    const by = fields.rows.length ? ` by ${fields.rows.join(", ")}` : "";
    return ok(`Created PivotTable "${name}" on ${sheetName}: ${describe}${by}. ${prepared.note}`.trim());
  },

  async add_sheet(ctx, { name }, undo) {
    const sheet = ctx.workbook.worksheets.add(name);
    sheet.load("name");
    await ctx.sync();
    const created = sheet.name;
    undo.record(`delete sheet ${created}`, () =>
      Excel.run(async (c) => {
        c.workbook.worksheets.getItem(created).delete();
        await c.sync();
      }),
    );
    return ok(`Added sheet "${created}"`);
  },

  async rename_sheet(ctx, { currentName, newName }, undo) {
    const sheet = ctx.workbook.worksheets.getItem(currentName);
    sheet.name = newName;
    await ctx.sync();
    undo.record(`rename sheet back to ${currentName}`, () =>
      Excel.run(async (c) => {
        c.workbook.worksheets.getItem(newName).name = currentName;
        await c.sync();
      }),
    );
    return ok(`Renamed sheet "${currentName}" to "${newName}"`);
  },

  async freeze_panes(ctx, { rows, columns }, undo) {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    const before = sheet.freezePanes.getLocationOrNullObject();
    before.load("address");
    sheet.load("name");
    await ctx.sync();
    const sheetName = sheet.name;
    const previous = before.isNullObject ? null : before.address;
    undo.record("restore frozen panes", () =>
      Excel.run(async (c) => {
        const ws = c.workbook.worksheets.getItem(sheetName);
        if (previous === null) ws.freezePanes.unfreeze();
        else ws.freezePanes.freezeAt(resolveRange(c, previous));
        await c.sync();
      }),
    );
    if (rows === 0 && columns === 0) sheet.freezePanes.unfreeze();
    else if (columns === 0) sheet.freezePanes.freezeRows(rows);
    else if (rows === 0) sheet.freezePanes.freezeColumns(columns);
    else sheet.freezePanes.freezeAt(sheet.getRangeByIndexes(0, 0, rows, columns));
    await ctx.sync();
    return ok(rows === 0 && columns === 0 ? "Unfroze panes" : `Froze ${rows} row(s) and ${columns} column(s) on ${sheetName}`);
  },

  async data_validation(ctx, a, undo) {
    const range = resolveRange(ctx, a.address);
    const dv = range.dataValidation;
    dv.load(["type", "rule"]);
    range.load("address");
    await ctx.sync();
    const fullAddress = range.address;
    const priorType = dv.type;
    const priorRule = dv.rule;
    const restorable = priorType !== "MixedCriteria";
    if (restorable) {
      undo.record(`restore validation on ${fullAddress}`, () =>
        Excel.run(async (c) => {
          const target = resolveRange(c, fullAddress).dataValidation;
          target.clear();
          if (priorType !== "None") target.rule = priorRule;
          await c.sync();
        }),
      );
    }
    const note = restorable ? "" : " (previous mixed validation cannot be restored by undo)";

    if (a.rule === "clear") {
      dv.clear();
      await ctx.sync();
      return ok(`Cleared validation on ${fullAddress}${note}`);
    }
    dv.clear();
    if (a.rule === "list") {
      dv.rule = { list: { inCellDropDown: true, source: (a.listValues ?? []).join(",") } };
    } else {
      const numeric = {
        formula1: a.value1 as number,
        ...(a.value2 !== undefined ? { formula2: a.value2 } : {}),
        operator: a.operator as Excel.DataValidationOperator,
      };
      dv.rule = { [a.rule]: numeric } as Excel.DataValidationRule;
    }
    if (a.errorMessage) {
      dv.errorAlert = { message: a.errorMessage, showAlert: true, style: "Stop", title: "Invalid entry" };
    }
    await ctx.sync();
    return ok(`Set ${a.rule} validation on ${fullAddress}${note}`);
  },

  async remove_duplicates(ctx, { address, columnIndexes, hasHeaders }, undo) {
    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.9")) {
      return { ok: false, error: "Removing duplicates needs Excel API 1.9 or newer. Update Excel." };
    }
    const range = resolveRange(ctx, address);
    range.load("columnCount");
    await ctx.sync();
    if (columnIndexes.some((i) => i >= range.columnCount)) {
      return { ok: false, error: `columnIndexes must be between 0 and ${range.columnCount - 1} for ${address}.` };
    }
    const where = await snapshotFormulas(ctx, range, undo);
    const result = range.removeDuplicates([...columnIndexes], hasHeaders);
    await ctx.sync();
    return ok(`Removed ${result.removed} duplicate row(s) from ${where}; ${result.uniqueRemaining} remain`);
  },

  async filter_range(ctx, { address, action, columnIndex, values }, undo) {
    const range = resolveRange(ctx, address);
    const sheet = range.worksheet;
    sheet.autoFilter.load("enabled");
    sheet.load("name");
    await ctx.sync();
    const sheetName = sheet.name;
    const hadFilter = sheet.autoFilter.enabled;

    if (action === "remove") {
      sheet.autoFilter.remove();
      await ctx.sync();
      return ok(`Removed the filter on ${sheetName}${hadFilter ? " (undo cannot bring the old criteria back)" : ""}`);
    }
    if (hadFilter) sheet.autoFilter.remove();
    if (columnIndex !== undefined && values) {
      sheet.autoFilter.apply(range, columnIndex, { filterOn: "Values", values: [...values] });
    } else {
      sheet.autoFilter.apply(range);
    }
    await ctx.sync();
    if (!hadFilter) {
      undo.record(`remove filter on ${sheetName}`, () =>
        Excel.run(async (c) => {
          c.workbook.worksheets.getItem(sheetName).autoFilter.remove();
          await c.sync();
        }),
      );
    }
    const by = values ? ` where column ${(columnIndex ?? 0) + 1} is ${values.join(" / ")}` : "";
    return ok(`Filtered ${sheetName}${by}${hadFilter ? " (replaced the existing filter; undo cannot restore it)" : ""}`);
  },

  async find_replace(ctx, a, undo) {
    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.9")) {
      return { ok: false, error: "Find & replace needs Excel API 1.9 or newer. Update Excel." };
    }
    const scope = a.address
      ? resolveRange(ctx, a.address)
      : ctx.workbook.worksheets.getActiveWorksheet().getUsedRangeOrNullObject(true);
    if (!a.address) {
      (scope as Excel.Range).load("isNullObject");
      await ctx.sync();
      if ((scope as Excel.Range).isNullObject) return ok("The sheet is empty; nothing to replace.");
    }
    const where = await snapshotFormulas(ctx, scope, undo);
    const result = scope.replaceAll(a.find, a.replace, { completeMatch: a.matchEntireCell, matchCase: a.matchCase });
    await ctx.sync();
    return ok(`Replaced "${a.find}" with "${a.replace}" in ${result.value} cell(s) of ${where}`);
  },

  async conditional_format(ctx, a, undo) {
    const range = resolveRange(ctx, a.address);
    let cf: Excel.ConditionalFormat;
    let described: string;

    if (a.rule === "cellValue") {
      cf = range.conditionalFormats.add("CellValue");
      if (a.fillColor) cf.cellValue.format.fill.color = a.fillColor;
      if (a.fontColor) cf.cellValue.format.font.color = a.fontColor;
      cf.cellValue.rule = {
        formula1: toRuleFormula(a.value1 as string | number),
        ...(a.value2 !== undefined ? { formula2: toRuleFormula(a.value2) } : {}),
        operator: a.operator as Excel.ConditionalCellValueOperator,
      };
      described = `cells ${a.operator} ${a.value1}${a.value2 !== undefined ? ` and ${a.value2}` : ""}`;
    } else if (a.rule === "text") {
      cf = range.conditionalFormats.add("ContainsText");
      if (a.fillColor) cf.textComparison.format.fill.color = a.fillColor;
      if (a.fontColor) cf.textComparison.format.font.color = a.fontColor;
      cf.textComparison.rule = { operator: "Contains", text: a.text as string };
      described = `cells containing "${a.text}"`;
    } else {
      cf = range.conditionalFormats.add("ColorScale");
      cf.colorScale.criteria = {
        minimum: { type: "LowestValue", color: a.minColor as string },
        ...(a.midColor ? { midpoint: { type: "Percentile", formula: "50", color: a.midColor } } : {}),
        maximum: { type: "HighestValue", color: a.maxColor as string },
      } as Excel.ConditionalColorScaleCriteria;
      described = "colour scale";
    }

    cf.load("id");
    range.load("address");
    await ctx.sync();
    const id = cf.id;
    const fullAddress = range.address;
    undo.record(`remove conditional format on ${fullAddress}`, () =>
      Excel.run(async (c) => {
        resolveRange(c, fullAddress).conditionalFormats.getItem(id).delete();
        await c.sync();
      }),
    );
    return ok(`Added conditional format on ${fullAddress}: ${described}`);
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
