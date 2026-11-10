import { formatToolError } from "../agent/toolError";

/** Pure checks run before a PivotTable is created, so existing data is never overwritten (no Excel globals). */

export interface Rect {
  readonly sheet: string;
  readonly top: number;
  readonly left: number;
  readonly bottom: number;
  readonly right: number;
}

export const rectContains = (rect: Rect, sheet: string, row: number, col: number): boolean =>
  rect.sheet.toLowerCase() === sheet.toLowerCase() && row >= rect.top && row <= rect.bottom && col >= rect.left && col <= rect.right;

export interface PivotDestination {
  readonly sheet: string;
  readonly row: number;
  readonly col: number;
  /** Address as the user or model wrote it, for the message. */
  readonly address: string;
  /** True when the destination cell already holds a value or formula. */
  readonly hasData: boolean;
  readonly source: Rect;
}

/** A tool error string when the destination would overwrite data, otherwise null. */
export const checkPivotDestination = (d: PivotDestination): string | null => {
  if (rectContains(d.source, d.sheet, d.row, d.col)) {
    return formatToolError(
      "DESTINATION_INSIDE_SOURCE",
      `The PivotTable destination ${d.address} is inside its own source data.`,
      "choose an empty cell outside the source range, or leave destinationAddress out to use a new sheet.",
    );
  }
  if (d.hasData) {
    return formatToolError(
      "DESTINATION_OCCUPIED",
      `The PivotTable destination ${d.address} already contains data, and the PivotTable would overwrite it.`,
      "choose an empty cell with room to grow, or leave destinationAddress out to use a new sheet.",
    );
  }
  return null;
};
