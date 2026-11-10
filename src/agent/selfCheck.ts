/** End-of-task self-check: audit the ranges the model wrote and, if problems show up, let it fix them. */

export const MIN_WRITES_FOR_CHECK = 2;
export const MAX_CHECK_RANGES = 3;
export const MAX_LISTED_FINDINGS = 8;

export interface CheckableCall {
  readonly name: string;
  readonly args: unknown;
  readonly ok: boolean;
}

export interface CheckFinding {
  readonly kind: string;
  readonly address: string;
  readonly detail: string;
}

/** Only multi-step tasks are worth the extra audit. */
export const shouldSelfCheck = (changeCount: number): boolean => changeCount >= MIN_WRITES_FOR_CHECK;

/** Distinct addresses of successful write_cells calls, in order, capped. */
export const rangesToCheck = (calls: readonly CheckableCall[]): readonly string[] => {
  const seen: string[] = [];
  for (const call of calls) {
    if (!call.ok || call.name !== "write_cells") continue;
    const address = (call.args as { address?: unknown } | null)?.address;
    if (typeof address === "string" && address.length > 0 && !seen.includes(address)) seen.push(address);
  }
  return seen.slice(0, MAX_CHECK_RANGES);
};

const isFinding = (v: unknown): v is CheckFinding => {
  const f = v as Partial<Record<keyof CheckFinding, unknown>> | null;
  return typeof f === "object" && f !== null && typeof f.kind === "string" && typeof f.address === "string" && typeof f.detail === "string";
};

/** Findings from an audit_formulas result's data; empty when the data is not in that shape. */
export const findingsFrom = (data: unknown): readonly CheckFinding[] => {
  const findings = (data as { findings?: unknown } | null)?.findings;
  return Array.isArray(findings) && findings.every(isFinding) ? findings : [];
};

export const buildSelfCheckMessage = (findings: readonly CheckFinding[]): string => {
  const shown = findings
    .slice(0, MAX_LISTED_FINDINGS)
    .map((f) => `- ${f.address}: ${f.detail}`)
    .join("\n");
  const extra = findings.length - MAX_LISTED_FINDINGS;
  const more = extra > 0 ? `\n- and ${extra} more` : "";
  return `[Automatic self-check after your changes. These are hints, not proof.]
${shown}${more}
Fix real mistakes you made (read the cells first if unsure). Leave anything that is intended, such as a totals row that differs on purpose. Then give your final summary.`;
};
