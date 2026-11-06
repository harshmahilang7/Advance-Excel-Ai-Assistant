import { describe, expect, it } from "vitest";
import { parseToolCall } from "./tools";

const base = { sourceAddress: "A1:D100", rows: ["Type"], values: [{ field: "Qty", summarizeBy: "sum" }] };

describe("create_pivot_table visibleRowsOnly", () => {
  it("defaults to visible rows only", () => {
    const r = parseToolCall("create_pivot_table", base);
    expect(r.ok && r.call.name === "create_pivot_table" && r.call.args.visibleRowsOnly).toBe(true);
  });

  it("can be switched off", () => {
    const r = parseToolCall("create_pivot_table", { ...base, visibleRowsOnly: false });
    expect(r.ok && r.call.name === "create_pivot_table" && r.call.args.visibleRowsOnly).toBe(false);
  });
});
