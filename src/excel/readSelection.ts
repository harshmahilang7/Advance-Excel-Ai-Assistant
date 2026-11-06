import type { CellValue, RangeSnapshot } from "./context";
import { keepVisibleRows } from "./visibility";
import { loadRowHidden, MAX_SCAN_ROWS } from "./visibleRange";

const MAX_ROWS_READ = 200;

/**
 * Reads the current selection from Excel: only the used part of it (so selecting a whole sheet or
 * column works) and only rows that are visible (so a filtered table gives the filtered rows).
 */
export const readSelection = async (): Promise<RangeSnapshot> =>
  Excel.run(async (context) => {
    const selected = context.workbook.getSelectedRange();
    selected.load("address");
    const used = selected.getUsedRangeOrNullObject(true);
    used.load(["rowCount", "address"]);
    const sheet = selected.worksheet;
    sheet.load("name");
    await context.sync();

    if (used.isNullObject) return { sheetName: sheet.name, address: selected.address, values: [] };

    const scanRows = Math.min(used.rowCount, MAX_SCAN_ROWS);
    const scanned = used.getRow(0).getResizedRange(scanRows - 1, 0);
    scanned.load("values");
    await context.sync();
    const hidden = await loadRowHidden(context, scanned, scanRows);

    const visible = keepVisibleRows(scanned.values as CellValue[][], hidden);
    return {
      sheetName: sheet.name,
      address: used.address,
      values: visible.rows.slice(0, MAX_ROWS_READ + 1),
      hiddenRowCount: visible.hiddenCount,
    };
  });

export const isExcelHost = (): boolean =>
  typeof Office !== "undefined" && Office.context?.host === Office.HostType.Excel;
