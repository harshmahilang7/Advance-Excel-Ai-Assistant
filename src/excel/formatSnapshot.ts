/** Pure helpers for snapshotting and restoring cell formatting (no Excel globals). */

export const MAX_FORMAT_SNAPSHOT_CELLS = 5000;

type Maybe<T> = T | null | undefined;

export interface CellFormat {
  readonly format?: {
    readonly font?: { readonly bold?: Maybe<boolean>; readonly italic?: Maybe<boolean>; readonly color?: Maybe<string> };
    readonly fill?: { readonly color?: Maybe<string> };
    readonly horizontalAlignment?: Maybe<string>;
    readonly wrapText?: Maybe<boolean>;
  };
}

/** Properties requested from Range.getCellProperties. */
export const FORMAT_PROPS_REQUEST = {
  format: {
    font: { bold: true, italic: true, color: true },
    fill: { color: true },
    horizontalAlignment: true,
    wrapText: true,
  },
} as const;

/** Per-cell snapshots are slow on huge ranges (e.g. whole columns), so cap them. */
export const canSnapshotFormats = (rows: number, columns: number): boolean =>
  Number.isInteger(rows) && Number.isInteger(columns) && rows > 0 && columns > 0 && rows * columns <= MAX_FORMAT_SNAPSHOT_CELLS;

/** The most used cells (inside a big range) we will snapshot, in blocks of MAX_FORMAT_SNAPSHOT_CELLS. */
export const MAX_TILED_SNAPSHOT_CELLS = 50_000;

export type SnapshotMode = "small" | "background-only" | "tiled" | "too-large";

/**
 * How to remember formatting before changing `targetCells` cells. `usedCells` is how many of them lie inside the
 * sheet's used range (null when unknown). Cells outside the used range hold no per-cell formatting, so only the
 * used part needs a cell-by-cell snapshot; the rest is restored from one background sample.
 */
export const snapshotMode = (targetCells: number, usedCells: number | null): SnapshotMode => {
  if (targetCells <= MAX_FORMAT_SNAPSHOT_CELLS) return "small";
  if (usedCells === null || usedCells <= 0) return "background-only";
  return usedCells <= MAX_TILED_SNAPSHOT_CELLS ? "tiled" : "too-large";
};

export interface SnapshotBlock {
  readonly rowOffset: number;
  readonly rows: number;
}

/** Row blocks of a rows x columns grid, each at most `maxCells` cells. Empty when even one row will not fit. */
export const planSnapshotBlocks = (rows: number, columns: number, maxCells: number): readonly SnapshotBlock[] => {
  if (!Number.isInteger(rows) || !Number.isInteger(columns) || rows <= 0 || columns <= 0) return [];
  const perBlock = Math.floor(maxCells / columns);
  if (perBlock < 1) return [];
  const blocks: SnapshotBlock[] = [];
  for (let rowOffset = 0; rowOffset < rows; rowOffset += perBlock) {
    blocks.push({ rowOffset, rows: Math.min(perBlock, rows - rowOffset) });
  }
  return blocks;
};

export interface BackgroundFormat {
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly fontColor?: string;
  readonly fillColor?: string;
  readonly horizontalAlignment?: string;
  readonly wrapText?: boolean;
}

/** Range-level settings that put a sample cell's formatting back on a whole range. No fillColor means "clear fill". */
export const backgroundFormat = (probe: CellFormat): BackgroundFormat => {
  const f = probe.format ?? {};
  const fill = f.fill?.color;
  return compact({
    bold: f.font?.bold,
    italic: f.font?.italic,
    fontColor: f.font?.color,
    fillColor: isNoFill(fill) ? undefined : fill,
    horizontalAlignment: f.horizontalAlignment,
    wrapText: f.wrapText,
  }) as BackgroundFormat;
};

const isNoFill = (color: Maybe<string>): boolean => !color || color.toUpperCase() === "#FFFFFF";

const compact = <T extends object>(obj: T): Partial<T> =>
  Object.fromEntries(Object.entries(obj).filter(([, v]) => v !== null && v !== undefined)) as Partial<T>;

const nonEmpty = <T extends object>(obj: T): T | undefined => (Object.keys(obj).length > 0 ? obj : undefined);

/**
 * Excel reports "no fill" as white, and writing white back would paint cells.
 * Drop it (the caller clears fills first) and strip nulls so setCellProperties accepts the payload.
 */
export const toRestoreProps = (rows: readonly (readonly CellFormat[])[]): CellFormat[][] =>
  rows.map((row) =>
    row.map((cell) => {
      const f = cell.format ?? {};
      const font = nonEmpty(compact({ bold: f.font?.bold, italic: f.font?.italic, color: f.font?.color }));
      const fill = isNoFill(f.fill?.color) ? undefined : { color: f.fill?.color as string };
      const rest = compact({ horizontalAlignment: f.horizontalAlignment, wrapText: f.wrapText });
      return { format: { ...(font ? { font } : {}), ...(fill ? { fill } : {}), ...rest } } as CellFormat;
    }),
  );
