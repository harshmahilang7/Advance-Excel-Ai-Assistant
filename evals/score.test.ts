import { describe, expect, it } from "vitest";
import { colOf, scoreCase, someCall, type RecordedCall } from "./score";

const call = (name: string, args: unknown = {}, ok = true): RecordedCall => ({ name, args, ok });

describe("scoreCase", () => {
  it("passes when required tools were called and forbidden ones were not", () => {
    const r = scoreCase({ mustCall: ["format_range"], mustNotCall: ["clear_range"] }, [call("read_range"), call("format_range")]);
    expect(r).toEqual({ pass: true, failures: [], rejectedCalls: 0 });
  });

  it("fails when a required tool is missing or only failed", () => {
    expect(scoreCase({ mustCall: ["sort_range"] }, []).failures).toEqual(["expected a successful sort_range call"]);
    expect(scoreCase({ mustCall: ["sort_range"] }, [call("sort_range", {}, false)]).pass).toBe(false);
  });

  it("fails on a forbidden tool even if that call was rejected", () => {
    const r = scoreCase({ mustNotCall: ["clear_range"] }, [call("clear_range", {}, false)]);
    expect(r.pass).toBe(false);
    expect(r.failures[0]).toMatch(/clear_range/);
  });

  it("runs the custom check on successful calls and the final text only", () => {
    const seen: { names: string[]; text: string }[] = [];
    scoreCase({ check: (calls, text) => (seen.push({ names: calls.map((c) => c.name), text }), null) }, [call("a"), call("b", {}, false)], "reply");
    expect(seen).toEqual([{ names: ["a"], text: "reply" }]);
  });

  it("reports the custom failure message and counts rejected calls", () => {
    const r = scoreCase({ check: () => "wrong column" }, [call("a", {}, false), call("b")]);
    expect(r).toEqual({ pass: false, failures: ["wrong column"], rejectedCalls: 1 });
  });
});

describe("someCall", () => {
  it("matches name and argument predicate on successful calls", () => {
    const calls = [call("freeze_panes", { rows: 1 }), call("freeze_panes", { rows: 5 }, false)];
    expect(someCall(calls, "freeze_panes", (a) => a.rows === 1)).toBe(true);
    expect(someCall(calls, "freeze_panes", (a) => a.rows === 5)).toBe(false);
  });

  it("tolerates non-object arguments", () => {
    expect(someCall([call("x", null)], "x", (a) => a.anything === undefined)).toBe(true);
  });
});

describe("colOf", () => {
  it("extracts the first column from plain, absolute and sheet-qualified addresses", () => {
    expect(colOf("E2")).toBe("E");
    expect(colOf("$e$2:E9")).toBe("E");
    expect(colOf("Sheet1!AB3")).toBe("AB");
    expect(colOf("'My Sheet'!C1:C9")).toBe("C");
  });

  it("returns empty for non-addresses", () => {
    expect(colOf(undefined)).toBe("");
    expect(colOf("???")).toBe("");
  });
});
