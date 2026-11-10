/** Tracks which workbook ranges the model read during a request, so the user can see what left their machine. */

export const MAX_LISTED = 10;

const addressOf = (args: unknown): string | null => {
  if (typeof args !== "object" || args === null) return null;
  const value = (args as { address?: unknown }).address;
  return typeof value === "string" && value.length > 0 ? value : null;
};

/** Short label for what a tool call read, or null when it reads nothing from the workbook. */
export const describeRead = (name: string, args: unknown): string | null => {
  switch (name) {
    case "read_range":
      return addressOf(args);
    case "audit_formulas":
      return addressOf(args) ?? "active sheet (formula audit)";
    case "get_workbook_info":
      return "workbook overview";
    default:
      return null;
  }
};

/** Returns a new list with `read` appended; the same list when it is already present. */
export const addRead = (reads: readonly string[], read: string): readonly string[] =>
  reads.includes(read) ? reads : [...reads, read];

export const formatReadLog = (reads: readonly string[]): string => {
  if (reads.length === 0) return "Sent to the AI: your selection summary only (no extra ranges read).";
  const shown = reads.slice(0, MAX_LISTED).join(", ");
  const extra = reads.length - MAX_LISTED;
  const tail = extra > 0 ? `, and ${extra} more` : "";
  return `Sent to the AI: your selection summary, plus ranges it read: ${shown}${tail}.`;
};
