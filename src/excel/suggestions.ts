import type { SheetSummary } from "./context";

export const DEFAULT_SUGGESTIONS: readonly string[] = [
  "Add a Total column that multiplies Qty by Price",
  "Bold the header row and add a light blue fill",
  "Sort by the second column, largest first",
  "Make a column chart from my selection",
];

/** Read-only prompts: the model may look and explain but must not change the workbook. */
export const REVIEW_SUGGESTIONS: readonly string[] = [
  "Audit this sheet for inconsistent formulas and hardcoded numbers. Do not change anything; rank the issues by impact.",
  "Find the largest or most unusual values in my selection and explain them. Do not change the workbook.",
  "Look at this workbook and suggest the next tasks worth automating. Do not change anything.",
];

const MAX_SUGGESTIONS = 5;

/** Selection-aware starter prompts built from real header names. Pure. */
export const suggestPrompts = (summary: SheetSummary | null): readonly string[] => {
  if (!summary || summary.columnCount === 0 || summary.dataRowCount === 0) return DEFAULT_SUGGESTIONS;

  const firstOf = (type: "text" | "number"): string | undefined =>
    summary.headers.find((_, i) => summary.columnTypes[i] === type);
  const text = firstOf("text");
  const num = firstOf("number");

  const out: string[] = ["Bold the header row and add a light blue fill"];
  if (num) out.push(`Sort by ${num}, largest first`);
  if (num) out.push(`Add a Total row that sums ${num}`);
  if (text && num) out.push(`Create a pivot table showing total ${num} by ${text}`);
  out.push("Make a column chart from my selection");
  out.push("Turn my selection into a table");

  return out.slice(0, MAX_SUGGESTIONS);
};
