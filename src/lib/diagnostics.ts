export interface DiagnosticsInput {
  readonly appVersion: string;
  readonly modelId: string;
  readonly keyCount: number;
  readonly host: string | undefined;
  readonly platform: string | undefined;
  /** Excel API set version -> supported. */
  readonly apiSets: Readonly<Record<string, boolean>>;
  readonly userAgent: string;
  readonly now: Date;
}

/** Plain-text environment report for bug reports. Never includes keys or workbook data. */
export const buildDiagnostics = (d: DiagnosticsInput): string =>
  [
    "Excel-Ai-Assistant diagnostics",
    `Time: ${d.now.toISOString()}`,
    `App version: ${d.appVersion}`,
    `Model: ${d.modelId || "none"}`,
    `API keys saved: ${d.keyCount}`,
    `Host: ${d.host ?? "unknown"}`,
    `Platform: ${d.platform ?? "unknown"}`,
    ...Object.entries(d.apiSets).map(([v, yes]) => `Excel API ${v}: ${yes ? "yes" : "no"}`),
    `User agent: ${d.userAgent}`,
    "No workbook data or API keys are included.",
  ].join("\n");
