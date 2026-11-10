import { describe, expect, it } from "vitest";
import { formatDuration, formatRunSummary } from "./runSummary";

describe("formatDuration", () => {
  it("formats short, second and minute durations", () => {
    expect(formatDuration(400)).toBe("under 1s");
    expect(formatDuration(57_300)).toBe("57s");
    expect(formatDuration(75_000)).toBe("1m 15s");
    expect(formatDuration(120_000)).toBe("2m");
  });

  it("treats bad input as under 1s", () => {
    expect(formatDuration(-5)).toBe("under 1s");
    expect(formatDuration(Number.NaN)).toBe("under 1s");
  });
});

describe("formatRunSummary", () => {
  it("counts changes and reads", () => {
    expect(formatRunSummary({ ms: 12_000, toolCalls: 5, changes: 3, queued: false })).toBe(
      "Worked for 12s · 5 steps · 3 changes",
    );
  });

  it("uses singular forms", () => {
    expect(formatRunSummary({ ms: 2_000, toolCalls: 1, changes: 1, queued: false })).toBe("Worked for 2s · 1 step · 1 change");
  });

  it("says queued in preview mode", () => {
    expect(formatRunSummary({ ms: 3_000, toolCalls: 2, changes: 2, queued: true })).toBe(
      "Worked for 3s · 2 steps · 2 changes queued",
    );
  });

  it("omits the change count when nothing changed", () => {
    expect(formatRunSummary({ ms: 3_000, toolCalls: 1, changes: 0, queued: false })).toBe("Worked for 3s · 1 step · no changes");
  });

  it("handles a reply with no tool calls", () => {
    expect(formatRunSummary({ ms: 1_400, toolCalls: 0, changes: 0, queued: false })).toBe("Worked for 1s · no changes");
  });
});
