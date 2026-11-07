import type { CellValue } from "./context";
import { summarizeSheets, MAX_OVERVIEW_SHEETS, type RawSheet, type SheetOverview } from "./sheetOverview";

/** Reads every sheet's used range and header row inside an existing Excel.run context. */
export const readSheetOverviewIn = async (ctx: Excel.RequestContext): Promise<SheetOverview> => {
  const sheets = ctx.workbook.worksheets;
  sheets.load("items/name");
  await ctx.sync();
  const probed = sheets.items.slice(0, MAX_OVERVIEW_SHEETS).map((ws) => {
    const used = ws.getUsedRangeOrNullObject(true);
    used.load("address");
    return { ws, used };
  });
  await ctx.sync();
  // Only touch non-empty sheets: calling members on a null object fails at the next sync.
  const withHeaders = probed.map(({ ws, used }) => {
    if (used.isNullObject) return { ws, used, header: null };
    const header = used.getRow(0);
    header.load("values");
    return { ws, used, header };
  });
  await ctx.sync();
  const raw: RawSheet[] = withHeaders.map(({ ws, used, header }) => ({
    name: ws.name,
    usedRange: used.isNullObject ? null : used.address,
    firstRow: header ? ((header.values[0] ?? []) as CellValue[]) : [],
  }));
  const overview = summarizeSheets(raw);
  return { ...overview, omittedSheets: overview.omittedSheets + Math.max(0, sheets.items.length - raw.length) };
};

export const readSheetOverview = async (): Promise<SheetOverview> => Excel.run(readSheetOverviewIn);
