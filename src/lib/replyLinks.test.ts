import { describe, expect, it } from "vitest";
import { linkifyReply } from "./replyLinks";

describe("linkifyReply", () => {
  it("returns plain text as one segment", () => {
    expect(linkifyReply("All done.")).toEqual([{ kind: "text", text: "All done." }]);
  });

  it("turns backticked cell addresses into refs without the backticks", () => {
    expect(linkifyReply("Totals are in `D2:D20` now.")).toEqual([
      { kind: "text", text: "Totals are in " },
      { kind: "ref", text: "D2:D20", address: "D2:D20" },
      { kind: "text", text: " now." },
    ]);
  });

  it("handles sheet-qualified and quoted sheet addresses", () => {
    const segs = linkifyReply("See `Sheet1!A1` and `'Q1 Sales'!B2:C9`.");
    expect(segs.filter((s) => s.kind === "ref").map((s) => (s as { address: string }).address)).toEqual([
      "Sheet1!A1",
      "'Q1 Sales'!B2:C9",
    ]);
  });

  it("allows absolute markers", () => {
    expect(linkifyReply("`$A$1:$B$2`")).toEqual([{ kind: "ref", text: "$A$1:$B$2", address: "$A$1:$B$2" }]);
  });

  it("leaves other backticked text and bare words alone", () => {
    expect(linkifyReply("Use `SUM` on Q3 and FY26.")).toEqual([{ kind: "text", text: "Use `SUM` on Q3 and FY26." }]);
  });

  it("does not link whole-column or row-less references", () => {
    expect(linkifyReply("`A:A` and `Sheet1!`")).toEqual([{ kind: "text", text: "`A:A` and `Sheet1!`" }]);
  });

  it("returns an empty list for empty text", () => {
    expect(linkifyReply("")).toEqual([]);
  });
});
