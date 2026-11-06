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
