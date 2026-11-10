import { describe, expect, it } from "vitest";
import { formatPlan, isPlanTool, PLAN_TOOL } from "./plan";
import { parseToolCall } from "./tools";

describe("formatPlan", () => {
  it("numbers the steps and trims them", () => {
    expect(formatPlan(["  Add Total column ", "Bold header"])).toBe("Plan:\n1. Add Total column\n2. Bold header");
  });

  it("drops blank steps", () => {
    expect(formatPlan(["A", "  ", "B"])).toBe("Plan:\n1. A\n2. B");
  });
});

describe("isPlanTool", () => {
  it("matches only the plan tool", () => {
    expect(isPlanTool(PLAN_TOOL)).toBe(true);
    expect(isPlanTool("write_cells")).toBe(false);
  });
});

describe("propose_plan schema", () => {
  it("accepts 2-8 short steps", () => {
    expect(parseToolCall(PLAN_TOOL, { steps: ["a", "b"] }).ok).toBe(true);
  });

  it("rejects a single step, too many steps and empty steps", () => {
    expect(parseToolCall(PLAN_TOOL, { steps: ["a"] }).ok).toBe(false);
    expect(parseToolCall(PLAN_TOOL, { steps: Array(9).fill("x") }).ok).toBe(false);
    expect(parseToolCall(PLAN_TOOL, { steps: ["a", ""] }).ok).toBe(false);
  });
});
