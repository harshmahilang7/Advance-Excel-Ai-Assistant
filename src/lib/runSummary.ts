/** One-line summary shown after each request. */

export interface RunStats {
  readonly ms: number;
  readonly toolCalls: number;
  readonly changes: number;
  readonly queued: boolean;
  /** Total cells changed, when every write reported its size. */
  readonly cells?: number | null;
}

export const formatDuration = (ms: number): string => {
  if (!Number.isFinite(ms) || ms < 1000) return "under 1s";
  const total = Math.round(ms / 1000);
  const minutes = Math.floor(total / 60);
  const seconds = total % 60;
  if (minutes === 0) return `${seconds}s`;
  return seconds === 0 ? `${minutes}m` : `${minutes}m ${seconds}s`;
};

const plural = (n: number, word: string): string => `${n} ${word}${n === 1 ? "" : "s"}`;

export const formatRunSummary = ({ ms, toolCalls, changes, queued, cells }: RunStats): string => {
  const parts = [`Worked for ${formatDuration(ms)}`];
  if (toolCalls > 0) parts.push(plural(toolCalls, "step"));
  if (changes === 0) parts.push("no changes");
  else parts.push(`${plural(changes, "change")}${queued ? " queued" : ""}`);
  if (cells && !queued) parts.push(plural(cells, "cell"));
  return parts.join(" · ");
};
