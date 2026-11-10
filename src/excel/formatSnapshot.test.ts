import { describe, expect, it } from "vitest";
import {
  backgroundFormat,
  canSnapshotFormats,
  MAX_FORMAT_SNAPSHOT_CELLS,
  MAX_TILED_SNAPSHOT_CELLS,
  planSnapshotBlocks,
  snapshotMode,
  toRestoreProps,
  type CellFormat,
} from "./formatSnapshot";

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

describe("snapshotMode", () => {
  it("uses the simple path for small ranges, without looking at the used range", () => {
    expect(snapshotMode(MAX_FORMAT_SNAPSHOT_CELLS, null)).toBe("small");
  });

  it("restores only the background when a big range holds no used cells", () => {
    expect(snapshotMode(1_048_576, 0)).toBe("background-only");
    expect(snapshotMode(1_048_576, null)).toBe("background-only");
  });

  it("tiles the used part of a whole-column range", () => {
    expect(snapshotMode(1_048_576, 400)).toBe("tiled");
    expect(snapshotMode(1_048_576, MAX_TILED_SNAPSHOT_CELLS)).toBe("tiled");
  });

  it("gives up when the used part is itself too large", () => {
    expect(snapshotMode(5_000_000, MAX_TILED_SNAPSHOT_CELLS + 1)).toBe("too-large");
  });
});

describe("planSnapshotBlocks", () => {
  it("splits rows into blocks that stay under the cell limit", () => {
    expect(planSnapshotBlocks(10, 1000, 5000)).toEqual([
      { rowOffset: 0, rows: 5 },
      { rowOffset: 5, rows: 5 },
    ]);
  });

  it("returns one block when everything fits", () => {
    expect(planSnapshotBlocks(100, 4, 5000)).toEqual([{ rowOffset: 0, rows: 100 }]);
  });

  it("covers every row exactly once", () => {
    const blocks = planSnapshotBlocks(23, 7, 50);
    expect(blocks.reduce((n, b) => n + b.rows, 0)).toBe(23);
    expect(blocks[1]?.rowOffset).toBe(blocks[0]?.rows);
  });

  it("returns nothing when a single row is wider than the limit", () => {
    expect(planSnapshotBlocks(10, 6000, 5000)).toEqual([]);
  });

  it("returns nothing for invalid sizes", () => {
    expect(planSnapshotBlocks(0, 5, 5000)).toEqual([]);
    expect(planSnapshotBlocks(5, 0, 5000)).toEqual([]);
  });
});

describe("backgroundFormat", () => {
  it("turns a probe cell into range-level settings", () => {
    const probe: CellFormat = {
      format: { font: { bold: true, italic: false, color: "#FF0000" }, fill: { color: "#FFEE00" }, horizontalAlignment: "Center", wrapText: true },
    };
    expect(backgroundFormat(probe)).toEqual({
      bold: true,
      italic: false,
      fontColor: "#FF0000",
      fillColor: "#FFEE00",
      horizontalAlignment: "Center",
      wrapText: true,
    });
  });

  it("leaves fillColor out for no fill, so the caller clears it", () => {
    expect(backgroundFormat({ format: { fill: { color: "#FFFFFF" } } })).toEqual({});
  });

  it("skips null and missing properties", () => {
    expect(backgroundFormat({ format: { font: { bold: null, color: undefined } } })).toEqual({});
    expect(backgroundFormat({})).toEqual({});
  });
});

