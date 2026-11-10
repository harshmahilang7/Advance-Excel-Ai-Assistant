import { splitSheetAddress } from "./address";

/** Activate the sheet (if the address names one) and select the range, so a reply's cell link jumps there. */
export const selectRange = async (address: string): Promise<void> => {
  const { sheet, cells } = splitSheetAddress(address);
  await Excel.run(async (ctx) => {
    const ws = sheet ? ctx.workbook.worksheets.getItem(sheet) : ctx.workbook.worksheets.getActiveWorksheet();
    if (sheet) ws.activate();
    ws.getRange(cells).select();
    await ctx.sync();
  });
};
