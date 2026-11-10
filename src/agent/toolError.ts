/**
 * Tool errors the model can act on: a stable CODE, what went wrong, and what to try next.
 * Format: "CODE: message Try: hint". The code is the leading ALL_CAPS word.
 */

export type ToolErrorCode =
  | "UNKNOWN_TOOL"
  | "INVALID_ARGUMENTS"
  | "EXCEL_TOO_OLD"
  | "DESTINATION_OCCUPIED"
  | "DESTINATION_INSIDE_SOURCE"
  | "SOURCE_EMPTY"
  | "RANGE_TOO_LARGE"
  | "COLUMN_OUT_OF_RANGE"
  | "PIVOT_SOURCE_INVALID";

const CODE_PATTERN = /^([A-Z][A-Z0-9_]{2,}):\s/;

export const formatToolError = (code: ToolErrorCode, message: string, hint?: string): string =>
  hint ? `${code}: ${message} Try: ${hint}` : `${code}: ${message}`;

/** The leading code of a formatted tool error, or null for free text. */
export const errorCodeOf = (error: string): string | null => CODE_PATTERN.exec(error)?.[1] ?? null;
