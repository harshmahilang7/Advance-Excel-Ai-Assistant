import { describe, expect, it } from "vitest";
import { parseToolCall } from "../agent/tools";
import {
  addressKind,
  deleteShift,
  hasAnyData,
  insertShift,
  otherCellsHaveData,
  undoDeleteShift,
  undoInsertShift,
} from "./structure";

describe("addressKind", () => {
  it("recognizes whole rows, whole columns and cell ranges", () => {
    expect(addressKind("3:5")).toBe("rows");
    expect(addressKind("$3:$5")).toBe("rows");
    expect(addressKind("C:D")).toBe("columns");
    expect(addressKind("Sheet1!C:C")).toBe("columns");
    expect(addressKind("'Q1 Sales'!2:2")).toBe("rows");
    expect(addressKind("B2:D9")).toBe("cells");
    expect(addressKind("A1")).toBe("cells");
  });
});

describe("insertShift / deleteShift", () => {
  it("forces the shift for whole rows and columns", () => {
    expect(insertShift("rows", "Right")).toBe("Down");
    expect(insertShift("columns", "Down")).toBe("Right");
    expect(deleteShift("rows", "Left")).toBe("Up");
    expect(deleteShift("columns", "Up")).toBe("Left");
  });

  it("uses the requested shift for cell ranges, with a default", () => {
    expect(insertShift("cells", "Right")).toBe("Right");
    expect(insertShift("cells", undefined)).toBe("Down");
    expect(deleteShift("cells", "Left")).toBe("Left");
    expect(deleteShift("cells", undefined)).toBe("Up");
  });

  it("maps each shift to the one that reverses it", () => {
    expect(undoInsertShift("Down")).toBe("Up");
    expect(undoInsertShift("Right")).toBe("Left");
    expect(undoDeleteShift("Up")).toBe("Down");
    expect(undoDeleteShift("Left")).toBe("Right");
  });
});

describe("data checks", () => {
  it("hasAnyData ignores blanks", () => {
    expect(hasAnyData([["", null], [undefined, ""]])).toBe(false);
    expect(hasAnyData([["", 0]])).toBe(true);
    expect(hasAnyData([[false]])).toBe(true);
    expect(hasAnyData([])).toBe(false);
  });

  it("otherCellsHaveData ignores the top-left cell", () => {
    expect(otherCellsHaveData([["keep", ""], ["", ""]])).toBe(false);
    expect(otherCellsHaveData([["keep", "x"]])).toBe(true);
    expect(otherCellsHaveData([["", ""], ["y", ""]])).toBe(true);
  });
});

describe("structure tool schemas", () => {
  it("accepts whole-row and whole-column addresses", () => {
    expect(parseToolCall("insert_range", { address: "3:5" }).ok).toBe(true);
    expect(parseToolCall("insert_range", { address: "C:D" }).ok).toBe(true);
    expect(parseToolCall("delete_range", { address: "Sheet1!2:2" }).ok).toBe(true);
    expect(parseToolCall("format_range", { address: "1:1", bold: true }).ok).toBe(true);
  });

  it("still rejects bad addresses and shifts", () => {
    expect(parseToolCall("insert_range", { address: "hello world" }).ok).toBe(false);
    expect(parseToolCall("insert_range", { address: "B2", shift: "Sideways" }).ok).toBe(false);
    expect(parseToolCall("delete_range", { address: "B2", shift: "Down" }).ok).toBe(false);
  });

  it("copy_range defaults to copying everything without overwriting", () => {
    const r = parseToolCall("copy_range", { sourceAddress: "A1:D6", destinationAddress: "G1" });
    expect(r.ok).toBe(true);
    if (r.ok && r.call.name === "copy_range") {
      expect(r.call.args.what).toBe("all");
      expect(r.call.args.overwrite).toBe(false);
    }
    expect(parseToolCall("copy_range", { sourceAddress: "A1", destinationAddress: "B1", what: "everything" }).ok).toBe(false);
  });

  it("merge_cells defaults to not discarding data", () => {
    const r = parseToolCall("merge_cells", { address: "A1:D1" });
    expect(r.ok).toBe(true);
    if (r.ok && r.call.name === "merge_cells") {
      expect(r.call.args.across).toBe(false);
      expect(r.call.args.discardOthers).toBe(false);
    }
  });
});
