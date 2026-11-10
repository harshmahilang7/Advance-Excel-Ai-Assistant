/** Pure helpers for the cell notes that explain where written formulas get their values (no Excel globals). */

export const MAX_NOTE_CHARS = 500;

const CELL = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/;

/** First cell of an address, without sheet name or `$` markers ("'Q1 Sales'!$A$1:$B$2" -> "A1"). Null if not a cell. */
export const localCell = (address: string): string | null => {
  const afterSheet = address.slice(address.lastIndexOf("!") + 1);
  const first = afterSheet.split(":")[0] ?? "";
  const match = CELL.exec(first.trim());
  return match ? `${match[1]!.toUpperCase()}${match[2]}` : null;
};

/** Trimmed note text, clipped to the limit; null when there is nothing to write. */
export const cleanNote = (text: string | undefined): string | null => {
  if (typeof text !== "string") return null;
  const trimmed = text.trim();
  if (trimmed.length === 0) return null;
  return trimmed.length > MAX_NOTE_CHARS ? `${trimmed.slice(0, MAX_NOTE_CHARS - 1)}…` : trimmed;
};
