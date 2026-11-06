/** Pure helpers for conditional formatting rules (no Excel globals). */

/**
 * Excel rule operands are formula strings. Numbers become `=5`, plain text becomes a quoted
 * string literal, and anything already starting with "=" (cell reference or formula) is kept.
 */
export const toRuleFormula = (value: string | number): string => {
  if (typeof value === "number") return `=${value}`;
  if (value.startsWith("=")) return value;
  return `="${value.replaceAll('"', '""')}"`;
};
