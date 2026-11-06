/** Pure scoring for eval cases: which tools the model called, with what arguments. No network, no Excel. */

export interface RecordedCall {
  readonly name: string;
  readonly args: unknown;
  /** False when the tool rejected the call (bad arguments etc.). */
  readonly ok: boolean;
}

export type Args = Readonly<Record<string, unknown>>;

export interface Expectations {
  /** Each tool must be called successfully at least once. */
  readonly mustCall?: readonly string[];
  /** None of these may be called at all (even unsuccessfully). */
  readonly mustNotCall?: readonly string[];
  /** Extra check on the successful calls and the model's final reply. Return a failure message, or null if fine. */
  readonly check?: (calls: readonly RecordedCall[], finalText: string) => string | null;
}

export interface CaseScore {
  readonly pass: boolean;
  readonly failures: readonly string[];
  /** Calls the tools rejected. Not a failure by itself, but a sign the prompt or schemas drifted. */
  readonly rejectedCalls: number;
}

export const argsOf = (call: RecordedCall): Args => (typeof call.args === "object" && call.args !== null ? (call.args as Args) : {});

/** True if some successful call to `name` has arguments satisfying `predicate`. */
export const someCall = (calls: readonly RecordedCall[], name: string, predicate: (args: Args) => boolean = () => true): boolean =>
  calls.some((c) => c.ok && c.name === name && predicate(argsOf(c)));

/** Column letters of the first cell in an A1 address: "Sheet1!$E2:E9" -> "E". Empty string if not an address. */
export const colOf = (address: unknown): string => {
  if (typeof address !== "string") return "";
  return /^\$?([A-Za-z]{1,3})/.exec(address.replace(/^.*!/, ""))?.[1]?.toUpperCase() ?? "";
};

export const scoreCase = (expect: Expectations, calls: readonly RecordedCall[], finalText = ""): CaseScore => {
  const okCalls = calls.filter((c) => c.ok);
  const missing = (expect.mustCall ?? []).filter((name) => !okCalls.some((c) => c.name === name));
  const forbidden = (expect.mustNotCall ?? []).filter((name) => calls.some((c) => c.name === name));
  const custom = expect.check?.(okCalls, finalText) ?? null;

  const failures = [
    ...missing.map((n) => `expected a successful ${n} call`),
    ...forbidden.map((n) => `must not call ${n}`),
    ...(custom ? [custom] : []),
  ];
  return { pass: failures.length === 0, failures, rejectedCalls: calls.filter((c) => !c.ok).length };
};
