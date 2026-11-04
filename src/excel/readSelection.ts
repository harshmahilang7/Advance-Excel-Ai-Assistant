import type { CellValue, RangeSnapshot } from "./context";

const MAX_ROWS_READ = 200;

/** Reads the current selection (capped to MAX_ROWS_READ rows) from Excel. */
export const readSelection = async (): Promise<RangeSnapshot> =>
  Excel.run(async (context) => {
    const selected = context.workbook.getSelectedRange();
    selected.load(["rowCount", "address"]);
    const sheet = selected.worksheet;
    sheet.load("name");
    await context.sync();

    const rows = Math.min(selected.rowCount, MAX_ROWS_READ);
    const capped = selected.getRow(0).getResizedRange(rows - 1, 0);
    capped.load(["values", "address"]);
    await context.sync();

    return {
      sheetName: sheet.name,
      address: capped.address,
      values: capped.values as CellValue[][],
    };
  });

export const isExcelHost = (): boolean =>
  typeof Office !== "undefined" && Office.context?.host === Office.HostType.Excel;
