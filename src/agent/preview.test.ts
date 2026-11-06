import { describe, expect, it, vi } from "vitest";
import type { ToolExecutor } from "./loop";
import { applyPending, createPreviewExecutor, describeCall, type PendingCall } from "./preview";

const okResult = (summary: string) => ({ ok: true as const, summary, data: null });

describe("createPreviewExecutor", () => {
  it("runs read-only tools for real and does not queue them", async () => {
    const real = vi.fn<ToolExecutor>(async () => okResult("read"));
    const queued: PendingCall[] = [];
    const exec = createPreviewExecutor(real, (c) => queued.push(c));
    const r = await exec("read_range", { address: "A1:B2" });
    expect(real).toHaveBeenCalledWith("read_range", { address: "A1:B2" });
    expect(r).toEqual(okResult("read"));
    expect(queued).toEqual([]);
  });

  it("treats audit_formulas as read-only", async () => {
    const real = vi.fn<ToolExecutor>(async () => okResult("audited"));
    const queued: PendingCall[] = [];
    const exec = createPreviewExecutor(real, (c) => queued.push(c));
    await exec("audit_formulas", {});
    expect(real).toHaveBeenCalledOnce();
    expect(queued).toEqual([]);
  });

  it("queues write tools without running them", async () => {
    const real = vi.fn<ToolExecutor>(async () => okResult("done"));
    const queued: PendingCall[] = [];
    const exec = createPreviewExecutor(real, (c) => queued.push(c));
    const r = await exec("write_cells", { address: "C1", data: [["Total"], [1]] });
    expect(real).not.toHaveBeenCalled();
    expect(r.ok).toBe(true);
    expect(queued).toHaveLength(1);
    expect(queued[0]).toMatchObject({ name: "write_cells" });
    expect(queued[0]?.summary).toMatch(/C1/);
  });

  it("rejects invalid calls so the model can fix them, and does not queue them", async () => {
    const real = vi.fn<ToolExecutor>(async () => okResult("x"));
    const queued: PendingCall[] = [];
    const exec = createPreviewExecutor(real, (c) => queued.push(c));
    const r = await exec("write_cells", { address: "not valid!!", data: [[1]] });
    expect(r.ok).toBe(false);
    expect(queued).toEqual([]);
    expect((await exec("nope", {})).ok).toBe(false);
  });
});

describe("applyPending", () => {
  const calls: readonly PendingCall[] = [
    { name: "write_cells", args: { address: "A1", data: [[1]] }, summary: "one" },
    { name: "autofit_columns", args: { address: "A:A" }, summary: "two" },
    { name: "clear_range", args: { address: "B1", what: "all" }, summary: "three" },
  ];

  it("runs calls in order and reports each result", async () => {
    const seen: string[] = [];
    const real: ToolExecutor = async (name) => {
      seen.push(name);
      return okResult(`did ${name}`);
    };
    const out = await applyPending(real, calls);
    expect(seen).toEqual(["write_cells", "autofit_columns", "clear_range"]);
    expect(out.results.map((r) => r.ok)).toEqual([true, true, true]);
    expect(out.skipped).toBe(0);
  });

  it("stops at the first failure and counts the skipped calls", async () => {
    const real: ToolExecutor = async (name) => (name === "autofit_columns" ? { ok: false, error: "boom" } : okResult("ok"));
    const out = await applyPending(real, calls);
    expect(out.results).toHaveLength(2);
    expect(out.results[1]).toMatchObject({ ok: false, text: "boom" });
    expect(out.skipped).toBe(1);
  });

  it("turns a thrown error into a failed result", async () => {
    const real: ToolExecutor = async () => {
      throw new Error("kaput");
    };
    const out = await applyPending(real, calls.slice(0, 1));
    expect(out.results[0]).toMatchObject({ ok: false });
  });
});

describe("describeCall", () => {
  it("mentions the target address and size for writes", () => {
    expect(describeCall("write_cells", { address: "D1", data: [["a"], ["b"], ["c"]] })).toMatch(/3 row.*D1/);
  });

  it("falls back to the tool name", () => {
    expect(describeCall("get_workbook_info", {})).toMatch(/get_workbook_info/);
  });
});
