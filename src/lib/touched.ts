import type { Touched } from "../agent/loop";

/** What a request changed, built up event by event. */
export interface TouchedState {
  readonly ranges: readonly Touched[];
  /** Number of successful writes that reported no range (so the cell total would be incomplete). */
  readonly unreported: number;
}

export const EMPTY_TOUCHED: TouchedState = { ranges: [], unreported: 0 };
export const MAX_LISTED_RANGES = 8;

/** New state after one successful write. `touched` is undefined for tools that do not report a range. */
export const addTouched = (state: TouchedState, touched: Touched | undefined): TouchedState => {
  if (!touched) return { ...state, unreported: state.unreported + 1 };
  const dup = state.ranges.some((r) => r.address === touched.address);
  return dup ? state : { ...state, ranges: [...state.ranges, touched] };
};

/** Total cells, or null when some writes did not report a size (a partial total would mislead). */
export const totalCells = (state: TouchedState): number | null =>
  state.unreported > 0 || state.ranges.length === 0 ? null : state.ranges.reduce((n, r) => n + r.cells, 0);

/** Clickable list of changed ranges for the chat (a backticked address becomes a link). Empty when none. */
export const formatChanged = (state: TouchedState): string => {
  if (state.ranges.length === 0) return "";
  const shown = state.ranges.slice(0, MAX_LISTED_RANGES).map((r) => `\`${r.address}\``);
  const more = state.ranges.length - shown.length;
  return `Changed: ${shown.join(", ")}${more > 0 ? ` and ${more} more` : ""}`;
};
