import { describe, expect, it } from "vitest";
import { toRuleFormula } from "./conditionalFormat";

describe("toRuleFormula", () => {
  it("turns numbers into formulas", () => {
    expect(toRuleFormula(100)).toBe("=100");
    expect(toRuleFormula(-2.5)).toBe("=-2.5");
  });

  it("quotes plain text and escapes embedded quotes", () => {
    expect(toRuleFormula("Late")).toBe('="Late"');
    expect(toRuleFormula('say "hi"')).toBe('="say ""hi"""');
  });

  it("passes through formulas and cell references", () => {
    expect(toRuleFormula("=$B$1")).toBe("=$B$1");
    expect(toRuleFormula("=AVERAGE(B:B)")).toBe("=AVERAGE(B:B)");
  });
});
