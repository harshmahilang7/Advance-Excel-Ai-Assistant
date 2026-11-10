import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";

interface Header {
  readonly key: string;
  readonly value: string;
}
interface Rule {
  readonly source: string;
  readonly headers: readonly Header[];
}

const config = JSON.parse(readFileSync(new URL("../../vercel.json", import.meta.url), "utf8")) as { headers: readonly Rule[] };
const all = config.headers.find((r) => r.source === "/(.*)")?.headers ?? [];
const get = (key: string): string | undefined => all.find((h) => h.key === key)?.value;

describe("vercel.json security headers", () => {
  it("sets the basic hardening headers on every route", () => {
    expect(get("X-Content-Type-Options")).toBe("nosniff");
    expect(get("Referrer-Policy")).toBe("no-referrer");
    expect(get("Permissions-Policy")).toContain("camera=()");
    expect(get("Strict-Transport-Security")).toMatch(/max-age=\d+/);
  });

  it("does not block Office from framing the pane", () => {
    const keys = all.map((h) => h.key.toLowerCase());
    expect(keys).not.toContain("x-frame-options");
    expect(get("Content-Security-Policy-Report-Only")).not.toContain("frame-ancestors");
    expect(get("Content-Security-Policy")).toBeUndefined();
  });

  it("keeps Office.js and the Gemini API reachable under the CSP", () => {
    const csp = get("Content-Security-Policy-Report-Only") ?? "";
    expect(csp).toContain("script-src 'self' https://appsforoffice.microsoft.com");
    expect(csp).toContain("connect-src 'self' https://generativelanguage.googleapis.com");
    expect(csp).not.toContain("unsafe-eval");
    expect(csp).toContain("object-src 'none'");
  });
});
