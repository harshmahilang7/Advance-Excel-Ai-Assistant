export interface VisibleRows<T> {
  readonly rows: readonly T[];
  readonly hiddenCount: number;
  readonly totalCount: number;
}

/**
 * Keeps the header row (index 0) and every row not flagged hidden. A missing flag means visible.
 * Pure: used for ranges filtered with AutoFilter or hidden by hand.
 */
export const keepVisibleRows = <T>(rows: readonly T[], hidden: readonly boolean[]): VisibleRows<T> => {
  const kept = rows.filter((_, i) => i === 0 || hidden[i] !== true);
  return { rows: kept, hiddenCount: rows.length - kept.length, totalCount: rows.length };
};
