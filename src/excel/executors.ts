import { parseToolCall, type ToolArgs, type ToolName } from "../agent/tools";
import type { ToolExecutor, ToolResult } from "../agent/loop";
import { splitSheetAddress } from "./address";
import { cleanGrid, type CleanCell } from "./cleaning";
import { toRuleFormula } from "./conditionalFormat";
import { auditFormulas, MAX_FINDINGS_PER_KIND } from "./formulaAudit";
import {
  backgroundFormat,
  canSnapshotFormats,
  FORMAT_PROPS_REQUEST,
  MAX_FORMAT_SNAPSHOT_CELLS,
  planSnapshotBlocks,
  snapshotMode,
  toRestoreProps,
  type BackgroundFormat,
  type CellFormat,
} from "./formatSnapshot";
import { readSheetOverviewIn } from "./readSheets";
import {
  addressKind,
  deleteShift,
  hasAnyData,
  insertShift,
  MAX_DELETE_UNDO_CELLS,
  otherCellsHaveData,
  undoDeleteShift,
  undoInsertShift,
} from "./structure";
import { cleanNote, localCell } from "./sourceNote";
import { formatPlan } from "../agent/plan";
import { describeVerification, verifyWrite } from "./writeVerify";
import { formatToolError } from "../agent/toolError";
import { checkPivotDestination } from "./pivotChecks";
import { resolvePivotFields, type PivotAggregation } from "./pivotFields";
import { preparePivotSource } from "./pivotSource";
import type { UndoLog } from "./undo";
import { keepVisibleRows } from "./visibility";
import { loadRowHidden, MAX_SCAN_ROWS } from "./visibleRange";

const DEFAULT_READ_ROWS = 100;
const MAX_AUDIT_CELLS = 50000;
const MAX_CLEAN_CELLS = 50000;
const KINDS_ORDER = ["error_value", "inconsistent_formula", "hardcoded_in_formulas"] as const;

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
 * Read the written cells back and report cells Excel stored differently or that show errors.
 * Never fails the write: any problem reading back just returns "".
 */
const verifyWritten = async (ctx: Ctx, target: Excel.Range, wrote: readonly (readonly unknown[])[]): Promise<string> => {
  try {
    target.load(["formulas", "values", "rowIndex", "columnIndex"]);
    await ctx.sync();
    const result = verifyWrite(wrote, target.formulas, target.values, { row: target.rowIndex, col: target.columnIndex });
    return describeVerification(result);
  } catch {
    return "";
  }
};

/** Pre-flight for a PivotTable destination: null when safe, else a tool error. Changes nothing. */
const pivotDestinationError = async (ctx: Ctx, sourceAddress: string, destinationAddress: string): Promise<string | null> => {
  const src = resolveRange(ctx, sourceAddress);
  const dest = resolveRange(ctx, destinationAddress).getCell(0, 0);
  src.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);
  src.worksheet.load("name");
  dest.load(["rowIndex", "columnIndex", "values"]);
  dest.worksheet.load("name");
  await ctx.sync();
  const cell = dest.values[0]?.[0];
  return checkPivotDestination({
    sheet: dest.worksheet.name,
    row: dest.rowIndex,
    col: dest.columnIndex,
    address: destinationAddress,
    hasData: cell !== "" && cell !== null && cell !== undefined,
    source: {
      sheet: src.worksheet.name,
      top: src.rowIndex,
      left: src.columnIndex,
      bottom: src.rowIndex + src.rowCount - 1,
      right: src.columnIndex + src.columnCount - 1,
    },
  });
};

const NOTES_API = "1.18";

/**
 * Save the model's explanation as a cell note on the top-left cell of a write. Never fails the write:
 * returns a short suffix for the step line ("" when no note was asked for).
 */
const attachSourceNote = async (ctx: Ctx, target: Excel.Range, raw: string | undefined, undo: UndoLog): Promise<string> => {
  const text = cleanNote(raw);
  if (!text) return "";
  if (!Office.context.requirements.isSetSupported("ExcelApi", NOTES_API)) {
    return ` (note skipped: cell notes need Excel API ${NOTES_API})`;
  }
  try {
    const first = target.getCell(0, 0);
    first.load("address");
    await ctx.sync();
    const fullAddress = first.address;
    const cell = localCell(fullAddress);
    if (!cell) return " (note skipped: no single cell to attach to)";
    const sheetNotes = first.worksheet.notes;
    const existing = sheetNotes.getItemOrNullObject(cell);
    existing.load(["isNullObject", "content"]);
    await ctx.sync();
    if (existing.isNullObject) {
      sheetNotes.add(cell, text);
      undo.record(`remove note ${fullAddress}`, () =>
        Excel.run(async (c) => {
          const found = resolveRange(c, fullAddress).worksheet.notes.getItemOrNullObject(cell);
          found.load("isNullObject");
          await c.sync();
          if (!found.isNullObject) found.delete();
          await c.sync();
        }),
      );
    } else {
      const previous = existing.content;
      existing.content = text;
      undo.record(`restore note ${fullAddress}`, () =>
        Excel.run(async (c) => {
          const found = resolveRange(c, fullAddress).worksheet.notes.getItemOrNullObject(cell);
          found.load("isNullObject");
          await c.sync();
          if (!found.isNullObject) found.content = previous;
          await c.sync();
        }),
      );
    }
    await ctx.sync();
    return ` (note added on ${cell})`;
  } catch {
    return " (note could not be added)";
  }
};

interface FormatChunk {
  readonly address: string;
  readonly props: Excel.CellProperties[][];
  readonly numberFormat: string[][];
}

/** Put a sample cell's formatting back on a whole range (works for whole columns and rows). */
const applyBackground = (range: Excel.Range, bg: BackgroundFormat, numberFormat: string): void => {
  range.format.fill.clear();
  if (bg.fillColor) range.format.fill.color = bg.fillColor;
  if (bg.bold !== undefined) range.format.font.bold = bg.bold;
  if (bg.italic !== undefined) range.format.font.italic = bg.italic;
  if (bg.fontColor) range.format.font.color = bg.fontColor;
  if (bg.horizontalAlignment) range.format.horizontalAlignment = bg.horizontalAlignment as Excel.HorizontalAlignment;
  if (bg.wrapText !== undefined) range.format.wrapText = bg.wrapText;
  // Excel applies a single value to every cell in the range, which avoids building a huge grid.
  range.numberFormat = numberFormat as unknown as string[][];
};

/** Snapshot one block of cells: their per-cell formatting and number formats. */
const snapshotBlock = async (ctx: Ctx, block: Excel.Range): Promise<FormatChunk> => {
  block.load(["address", "numberFormat"]);
  const props = block.getCellProperties(FORMAT_PROPS_REQUEST as unknown as Excel.CellPropertiesLoadOptions);
  await ctx.sync();
  return {
    address: block.address,
    props: toRestoreProps(props.value as unknown as CellFormat[][]) as unknown as Excel.CellProperties[][],
    numberFormat: block.numberFormat.map((row) => [...row]),
  };
};

/**
 * Large range (e.g. a whole column): snapshot only the cells that hold real formatting (the part inside the
 * sheet's used range, in blocks) plus one sample cell outside it for the empty remainder.
 * Returns the restore plan, or null when the used part is too large to remember.
 */
const snapshotLargeRange = async (
  ctx: Ctx,
  range: Excel.Range,
): Promise<{ chunks: readonly FormatChunk[]; background: { fmt: BackgroundFormat; numberFormat: string } | null } | null> => {
  const used = range.worksheet.getUsedRangeOrNullObject(false);
  used.load("isNullObject");
  await ctx.sync();

  let inter: Excel.Range | null = null;
  if (!used.isNullObject) {
    const found = range.getIntersectionOrNullObject(used);
    found.load(["isNullObject", "rowCount", "columnCount"]);
    await ctx.sync();
    if (!found.isNullObject) inter = found;
  }

  const usedCells = inter ? inter.rowCount * inter.columnCount : 0;
  const mode = snapshotMode(Number.MAX_SAFE_INTEGER, usedCells);
  if (mode === "too-large") return null;

  const chunks: FormatChunk[] = [];
  if (inter && mode === "tiled") {
    const blocks = planSnapshotBlocks(inter.rowCount, inter.columnCount, MAX_FORMAT_SNAPSHOT_CELLS);
    if (blocks.length === 0) return null;
    for (const b of blocks) {
      chunks.push(await snapshotBlock(ctx, inter.getCell(b.rowOffset, 0).getResizedRange(b.rows - 1, inter.columnCount - 1)));
    }
  }

  // Background sample: a cell of the target that lies outside the used range.
  const candidates = [range.getLastCell(), range.getCell(0, 0)];
  const outside: Excel.Range[] = [];
  if (used.isNullObject) outside.push(candidates[1] as Excel.Range);
  else {
    const hits = candidates.map((c) => c.getIntersectionOrNullObject(used));
    hits.forEach((h) => h.load("isNullObject"));
    await ctx.sync();
    hits.forEach((h, i) => h.isNullObject && outside.push(candidates[i] as Excel.Range));
  }
  const probe = outside[0];
  if (!probe) return { chunks, background: null };

  probe.load("numberFormat");
  const sample = probe.getCellProperties(FORMAT_PROPS_REQUEST as unknown as Excel.CellPropertiesLoadOptions);
  await ctx.sync();
  const cell = (sample.value as unknown as CellFormat[][])[0]?.[0] ?? {};
  return { chunks, background: { fmt: backgroundFormat(cell), numberFormat: String(probe.numberFormat[0]?.[0] ?? "General") } };
};

/**
 * Remember cell formatting (font, fill, alignment, wrap, number format) so the command can be undone.
 * Small ranges are snapshotted cell by cell; big ones (whole columns) by the used part plus a background sample.
 * Returns false when it cannot (old Excel, or a range whose used part is too large).
 */
const snapshotFormats = async (ctx: Ctx, range: Excel.Range, undo: UndoLog): Promise<boolean> => {
  if (!Office.context.requirements.isSetSupported("ExcelApi", "1.9")) return false;
  range.load(["address", "rowCount", "columnCount"]);
  await ctx.sync();
  const fullAddress = range.address;

  if (canSnapshotFormats(range.rowCount, range.columnCount)) {
    range.load("numberFormat");
    const props = range.getCellProperties(FORMAT_PROPS_REQUEST as unknown as Excel.CellPropertiesLoadOptions);
    await ctx.sync();
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
  }

  const plan = await snapshotLargeRange(ctx, range);
  if (!plan) return false;
  undo.record(`restore formatting ${fullAddress}`, () =>
    Excel.run(async (c) => {
      if (plan.background) applyBackground(resolveRange(c, fullAddress), plan.background.fmt, plan.background.numberFormat);
      for (const chunk of plan.chunks) {
        const r = resolveRange(c, chunk.address);
        r.format.fill.clear();
        r.setCellProperties(chunk.props);
        r.numberFormat = chunk.numberFormat;
      }
      await c.sync();
    }),
  );
  return true;
};

const NOT_UNDOABLE = " (formatting too large to undo)";

type Handlers ={ [N in ToolName]: (ctx: Ctx, args: ToolArgs<N>, undo: UndoLog) => Promise<ToolResult> };

const handlers: Handlers = {
  async propose_plan(_ctx, { steps }) {
    return ok(formatPlan(steps));
  },

  async get_workbook_info(ctx) {
    const { sheets, omittedSheets } = await readSheetOverviewIn(ctx);
    const more = omittedSheets > 0 ? `, ${omittedSheets} more not listed` : "";
    return ok(`Looked at the workbook: ${sheets.length} sheet${sheets.length === 1 ? "" : "s"}${more}`, sheets.map((s) => ({ name: s.name, usedRange: s.usedRange, headers: s.headers })));
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

  async write_cells(ctx, { address, data, note }, undo) {
    const rows = data.length;
    const cols = data[0]?.length ?? 1;
    const target = resolveRange(ctx, address).getCell(0, 0).getResizedRange(rows - 1, cols - 1);
    const where = await snapshotFormulas(ctx, target, undo);
    // null cells stay null: Office.js leaves those cells unchanged. Use "" to blank a cell.
    target.formulas = data.map((row) => [...row]);
    await ctx.sync();
    const verified = await verifyWritten(ctx, target, data);
    const noteStatus = await attachSourceNote(ctx, target, note, undo);
    return ok(`Wrote ${rows}x${cols} cells at ${where}${verified}${noteStatus}`);
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
    const used = range.getUsedRangeOrNullObject(true);
    used.load("isNullObject");
    await ctx.sync();
    if (used.isNullObject) {
      return {
        ok: false,
        error: formatToolError("SOURCE_EMPTY", `There is no data in ${sourceAddress} to chart.`, "read the sheet and pass the range that holds the data, headers included."),
      };
    }
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
      return { ok: false, error: formatToolError("EXCEL_TOO_OLD", "PivotTables need Excel API 1.8 or newer.", "tell the user to update Excel, or summarize with SUMIF formulas instead.") };
    }
    if (args.destinationAddress) {
      const blocked = await pivotDestinationError(ctx, args.sourceAddress, args.destinationAddress);
      if (blocked) return { ok: false, error: blocked };
    }
    const suffix = Date.now().toString(36).slice(-4);
    const prepared = await preparePivotSource(ctx, resolveRange(ctx, args.sourceAddress), args.visibleRowsOnly, suffix, undo);
    if (!prepared.ok) {
      return { ok: false, error: formatToolError("PIVOT_SOURCE_INVALID", prepared.error, "read the source range and fix its header row, then try again.") };
    }
    const source = prepared.range;
    const fields = resolvePivotFields(prepared.headers, args);
    if (!fields.ok) {
      return { ok: false, error: formatToolError("PIVOT_SOURCE_INVALID", fields.error, "use the header names exactly as they appear in the source.") };
    }

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

  async audit_formulas(ctx, { address }) {
    const scope = address ? resolveRange(ctx, address) : ctx.workbook.worksheets.getActiveWorksheet().getUsedRangeOrNullObject(true);
    const used = address ? scope.getUsedRangeOrNullObject(true) : (scope as Excel.Range);
    used.load(["address", "rowCount", "columnCount", "rowIndex", "columnIndex"]);
    await ctx.sync();
    if (used.isNullObject) return ok("Nothing to audit: the range is empty.", { findings: [] });
    if (used.rowCount * used.columnCount > MAX_AUDIT_CELLS) {
      return { ok: false, error: formatToolError("RANGE_TOO_LARGE", `The range is too large to audit (over ${MAX_AUDIT_CELLS} cells).`, "pass a smaller address, such as one block of columns.") };
    }
    used.load(["formulasR1C1", "values"]);
    await ctx.sync();
    const result = auditFormulas({
      startRow: used.rowIndex,
      startCol: used.columnIndex,
      formulasR1C1: used.formulasR1C1,
      values: used.values,
    });
    const counts = KINDS_ORDER.map((k) => `${result.totals[k]} ${k.replaceAll("_", " ")}`).join(", ");
    const cap = result.truncated ? ` (showing first ${MAX_FINDINGS_PER_KIND} of each kind)` : "";
    return ok(`Audited ${used.address}: ${result.formulaCount} formulas, ${counts}${cap}`, { range: used.address, ...result });
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

  async insert_range(ctx, { address, shift }, undo) {
    const range = resolveRange(ctx, address);
    range.load("address");
    await ctx.sync();
    const fullAddress = range.address;
    const direction = insertShift(addressKind(address), shift);
    range.insert(direction);
    await ctx.sync();
    const reverse = undoInsertShift(direction);
    undo.record(`remove inserted ${fullAddress}`, () =>
      Excel.run(async (c) => {
        resolveRange(c, fullAddress).delete(reverse);
        await c.sync();
      }),
    );
    return ok(`Inserted blank cells at ${fullAddress}, shifting existing content ${direction.toLowerCase()}`);
  },

  async delete_range(ctx, { address, shift }, undo) {
    const range = resolveRange(ctx, address);
    range.load("address");
    const used = range.worksheet.getUsedRangeOrNullObject(true).getIntersectionOrNullObject(range);
    used.load(["rowCount", "columnCount"]);
    await ctx.sync();
    const fullAddress = range.address;
    const direction = deleteShift(addressKind(address), shift);
    if (!used.isNullObject) {
      if (used.rowCount * used.columnCount > MAX_DELETE_UNDO_CELLS) {
        return { ok: false, error: formatToolError("RANGE_TOO_LARGE", `${fullAddress} holds more than ${MAX_DELETE_UNDO_CELLS} used cells, too many to restore on undo.`, "delete smaller pieces, or ask the user to do it by hand.") };
      }
      await snapshotFormulas(ctx, used, undo);
      await snapshotFormats(ctx, used, undo);
    }
    range.delete(direction);
    await ctx.sync();
    const back = undoDeleteShift(direction);
    // Recorded last so undo runs it first: the space returns, then the saved contents are written back.
    undo.record(`re-insert ${fullAddress}`, () =>
      Excel.run(async (c) => {
        resolveRange(c, fullAddress).insert(back);
        await c.sync();
      }),
    );
    return ok(`Deleted ${fullAddress}, shifting remaining content ${direction.toLowerCase()}. Formulas elsewhere that used it may now show #REF! (undo does not repair those)`);
  },

  async copy_range(ctx, { sourceAddress, destinationAddress, what, overwrite }, undo) {
    const src = resolveRange(ctx, sourceAddress);
    src.load(["address", "rowCount", "columnCount"]);
    await ctx.sync();
    if (src.rowCount * src.columnCount > MAX_DELETE_UNDO_CELLS) {
      return { ok: false, error: formatToolError("RANGE_TOO_LARGE", `${src.address} is too large to copy with undo.`, "copy a smaller range.") };
    }
    const dest = resolveRange(ctx, destinationAddress).getCell(0, 0).getResizedRange(src.rowCount - 1, src.columnCount - 1);
    dest.load(["address", "values"]);
    await ctx.sync();
    if (what !== "formats" && !overwrite && hasAnyData(dest.values)) {
      return { ok: false, error: formatToolError("DESTINATION_OCCUPIED", `${dest.address} already contains data.`, "pick an empty destination, or set overwrite true if the user asked to replace it.") };
    }
    let undoable = true;
    if (what !== "formats") await snapshotFormulas(ctx, dest, undo);
    if (what !== "values") undoable = await snapshotFormats(ctx, dest, undo);
    const copyType = { all: Excel.RangeCopyType.all, values: Excel.RangeCopyType.values, formats: Excel.RangeCopyType.formats }[what];
    dest.copyFrom(src, copyType);
    await ctx.sync();
    return ok(`Copied ${what} of ${src.address} to ${dest.address}${undoable ? "" : NOT_UNDOABLE}`);
  },

  async merge_cells(ctx, { address, across, discardOthers }, undo) {
    const range = resolveRange(ctx, address);
    range.load(["address", "values"]);
    await ctx.sync();
    const loses = across ? range.values.some((row) => otherCellsHaveData([row])) : otherCellsHaveData(range.values);
    if (loses && !discardOthers) {
      return { ok: false, error: formatToolError("DATA_WOULD_BE_LOST", `Merging ${range.address} would discard values in all but the top-left cell.`, "move or combine the data first, or set discardOthers true if the user accepts losing it.") };
    }
    const where = await snapshotFormulas(ctx, range, undo);
    range.merge(across);
    await ctx.sync();
    undo.record(`unmerge ${where}`, () =>
      Excel.run(async (c) => {
        resolveRange(c, where).unmerge();
        await c.sync();
      }),
    );
    return ok(`Merged ${where}${across ? " row by row" : ""}${loses ? " (other values discarded)" : ""}`);
  },

  async remove_duplicates(ctx, { address, columnIndexes, hasHeaders }, undo) {
    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.9")) {
      return { ok: false, error: formatToolError("EXCEL_TOO_OLD", "Removing duplicates needs Excel API 1.9 or newer.", "tell the user to update Excel.") };
    }
    const range = resolveRange(ctx, address);
    range.load("columnCount");
    await ctx.sync();
    if (columnIndexes.some((i) => i >= range.columnCount)) {
      return { ok: false, error: formatToolError("COLUMN_OUT_OF_RANGE", `columnIndexes must be between 0 and ${range.columnCount - 1} for ${address}.`, "use 0-based column positions within the range.") };
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
      return { ok: false, error: formatToolError("EXCEL_TOO_OLD", "Find & replace needs Excel API 1.9 or newer.", "tell the user to update Excel.") };
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

  async clean_data(ctx, { address, operations }, undo) {
    const range = resolveRange(ctx, address);
    range.load(["rowCount", "columnCount"]);
    await ctx.sync();
    if (range.rowCount * range.columnCount > MAX_CLEAN_CELLS) {
      return { ok: false, error: formatToolError("RANGE_TOO_LARGE", `The range is too large to clean (max ${MAX_CLEAN_CELLS} cells).`, "clean one block of columns at a time.") };
    }
    range.load(["address", "formulas"]);
    await ctx.sync();
    const result = cleanGrid(range.formulas as CleanCell[][], operations);
    if (result.changed === 0) return ok(`Nothing to clean in ${range.address}`);
    const where = await snapshotFormulas(ctx, range, undo);
    range.formulas = result.values.map((row) => row.map((c) => c ?? "")) as (string | number | boolean)[][];
    await ctx.sync();
    const parts = operations.flatMap((op) => (result.perOperation[op] > 0 ? [`${op}: ${result.perOperation[op]}`] : []));
    return ok(`Cleaned ${result.changed} cell(s) in ${where} (${parts.join(", ")})`);
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
