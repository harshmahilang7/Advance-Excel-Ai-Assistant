import { describe, expect, it } from "vitest";
import { buildDiagnostics } from "./diagnostics";

const base = {
  appVersion: "0.1.0",
  modelId: "gemini-x",
  keyCount: 2,
  host: "Excel",
  platform: "PC",
  apiSets: { "1.9": true, "1.18": false },
  userAgent: "UA/1.0",
  now: new Date("2026-11-11T10:00:00Z"),
};

describe("buildDiagnostics", () => {
  it("lists versions, model and API support", () => {
    const text = buildDiagnostics(base);
    expect(text).toContain("App version: 0.1.0");
    expect(text).toContain("Model: gemini-x");
    expect(text).toContain("Excel API 1.9: yes");
    expect(text).toContain("Excel API 1.18: no");
    expect(text).toContain("2026-11-11T10:00:00.000Z");
  });

  it("reports only the number of keys, never key text or workbook data", () => {
    const text = buildDiagnostics({ ...base, keyCount: 3 });
    expect(text).toContain("API keys saved: 3");
    expect(text).not.toMatch(/AIza/);
    expect(text).toContain("No workbook data or API keys are included.");
  });

  it("handles missing Office info", () => {
    const text = buildDiagnostics({ ...base, host: undefined, platform: undefined, apiSets: {} });
    expect(text).toContain("Host: unknown");
    expect(text).toContain("Platform: unknown");
  });
});
