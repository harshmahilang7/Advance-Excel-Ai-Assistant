import { describe, expect, it } from "vitest";
import { canSnapshotFormats, MAX_FORMAT_SNAPSHOT_CELLS, toRestoreProps, type CellFormat } from "./formatSnapshot";

describe("canSnapshotFormats", () => {
  it("allows small ranges", () => {
    expect(canSnapshotFormats(10, 5)).toBe(true);
  });
  it("allows exactly the cap", () => {
    expect(canSnapshotFormats(MAX_FORMAT_SNAPSHOT_CELLS, 1)).toBe(true);
  });
  it("rejects whole-column sized ranges", () => {
    expect(canSnapshotFormats(1_048_576, 1)).toBe(false);
  });
  it("rejects invalid dimensions", () => {
    expect(canSnapshotFormats(0, 5)).toBe(false);
    expect(canSnapshotFormats(Number.NaN, 5)).toBe(false);
  });
});

describe("toRestoreProps", () => {
  it("drops the default white fill so restore leaves cells unfilled", () => {
    const input: CellFormat[][] = [[{ format: { fill: { color: "#FFFFFF" }, font: { bold: false } } }]];
    expect(toRestoreProps(input)).toEqual([[{ format: { font: { bold: false } } }]]);
  });

  it("treats white fill case-insensitively and empty string as no fill", () => {
    const input: CellFormat[][] = [[{ format: { fill: { color: "#ffffff" } } }, { format: { fill: { color: "" } } }]];
    expect(toRestoreProps(input)).toEqual([[{ format: {} }, { format: {} }]]);
  });

  it("keeps a real fill colour", () => {
    const input: CellFormat[][] = [[{ format: { fill: { color: "#DDEEFF" } } }]];
    expect(toRestoreProps(input)).toEqual([[{ format: { fill: { color: "#DDEEFF" } } }]]);
  });

  it("strips null and undefined values", () => {
    const input: CellFormat[][] = [[{ format: { font: { bold: null, italic: true }, wrapText: null } }]];
    expect(toRestoreProps(input)).toEqual([[{ format: { font: { italic: true } } }]]);
  });

  it("does not mutate its input", () => {
    const input: CellFormat[][] = [[{ format: { fill: { color: "#FFFFFF" }, font: { bold: true } } }]];
    const before = JSON.stringify(input);
    toRestoreProps(input);
    expect(JSON.stringify(input)).toBe(before);
  });
});
