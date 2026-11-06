import { describe, expect, it } from "vitest";
import { createFakeExecutor, type FakeSheet } from "./fakeExecutor";

const sheet: FakeSheet = {
  name: "Sheet1",
  address: "Sheet1!A1:B3",
  values: [["Item", "Qty"], ["a", 1], ["b", 2]],
};

describe("createFakeExecutor", () => {
  it("answers reads from the fake sheet", async () => {
    const fake = createFakeExecutor(sheet);
    const r = await fake.execute("read_range", { address: "A1:B3" });
    expect(r.ok).toBe(true);
    if (r.ok) expect((r.data as { values: unknown[] }).values).toHaveLength(3);
  });

  it("accepts valid writes without side effects and records them", async () => {
    const fake = createFakeExecutor(sheet);
    const r = await fake.execute("freeze_panes", { rows: 1 });
    expect(r.ok).toBe(true);
    expect(fake.calls()).toEqual([{ name: "freeze_panes", args: { rows: 1, columns: 0 }, ok: true }]);
  });

  it("rejects invalid arguments with the real schema error and records the rejection", async () => {
    const fake = createFakeExecutor(sheet);
    const r = await fake.execute("sort_range", { address: "A1:B3", columnIndex: -1, ascending: true, hasHeaders: true });
    expect(r.ok).toBe(false);
    expect(fake.calls()[0]).toMatchObject({ name: "sort_range", ok: false });
    expect((await fake.execute("delete_everything", {})).ok).toBe(false);
  });

  it("keeps separate call logs per executor", async () => {
    const a = createFakeExecutor(sheet);
    const b = createFakeExecutor(sheet);
    await a.execute("get_workbook_info", {});
    expect(a.calls()).toHaveLength(1);
    expect(b.calls()).toHaveLength(0);
  });
});
