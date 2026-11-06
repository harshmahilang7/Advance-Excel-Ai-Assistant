/** Excel-facing helpers for filtered / hidden rows. Pure logic lives in visibility.ts. */

// Per-row probes are batched into one sync; beyond this many rows we assume the rest are visible.
export const MAX_SCAN_ROWS = 5000;

/**
 * For each row of `range` (up to MAX_SCAN_ROWS) reports whether Excel hides it
 * (filtered out by AutoFilter or hidden manually).
 */
export const loadRowHidden = async (ctx: Excel.RequestContext, range: Excel.Range, rowCount: number): Promise<boolean[]> => {
  const probes = Array.from({ length: Math.min(rowCount, MAX_SCAN_ROWS) }, (_, i) => {
    const row = range.getRow(i);
    row.load("rowHidden");
    return row;
  });
  await ctx.sync();
  return probes.map((p) => p.rowHidden === true);
};
