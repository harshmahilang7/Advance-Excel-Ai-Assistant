interface Rule {
  readonly code?: RegExp;
  readonly message?: RegExp;
  readonly text: string;
}

// First match wins, so keep the more specific rules first.
const RULES: readonly Rule[] = [
  { message: /merged/i, text: "That range overlaps merged cells. Unmerge them or pick a range that avoids them." },
  { code: /^AccessDenied$/i, text: "Excel blocked the change. The sheet or workbook may be protected; unprotect it and try again." },
  { message: /protected/i, text: "The sheet or workbook is protected. Unprotect it and try again." },
  { code: /^ItemNotFound$/i, text: "Could not find that sheet, table, chart or range. Check the name and try again." },
  { code: /^ItemAlreadyExists$/i, text: "That name is already used in the workbook. Pick a different name." },
  { code: /^(ApiNotFound|RichApiMessageNotSupported|NotImplemented)$/i, text: "This Excel version does not support that feature. Update Excel to a newer version." },
  { message: /failed to fetch|networkerror|network request failed/i, text: "Network problem: could not reach the AI service. Check your connection and try again." },
];

const codeOf = (err: unknown): string => {
  const code = (err as { code?: unknown } | null)?.code;
  return typeof code === "string" ? code : "";
};

const rawMessage = (err: unknown): string => {
  if (err instanceof Error) return err.message;
  if (typeof err === "string" && err.length > 0) return err;
  return "";
};

/** Turns Excel/Office/network errors into plain English. Unknown errors pass through unchanged. */
export const friendlyError = (err: unknown): string => {
  const message = rawMessage(err);
  const code = codeOf(err);
  const rule = RULES.find((r) => (r.code?.test(code) ?? false) || (r.message?.test(message) ?? false));
  if (rule) return rule.text;
  return message || "Something went wrong.";
};
