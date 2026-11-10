import { describe, expect, it } from "vitest";
import { MAX_COLS, MAX_ROWS, parseDelimited, safeCell } from "./csv";

describe("parseDelimited", () => {
  it("parses simple comma data and pads ragged rows", () => {
    const r = parseDelimited("a,b,c\n1,2\n3,4,5\n");
    expect(r.rows).toEqual([["a", "b", "c"], ["1", "2", ""], ["3", "4", "5"]]);
    expect(r.delimiter).toBe(",");
  });

  it("handles CRLF and a trailing blank line", () => {
    expect(parseDelimited("a,b\r\n1,2\r\n\r\n").rows).toEqual([["a", "b"], ["1", "2"]]);
  });

  it("handles quotes, embedded commas, newlines and doubled quotes", () => {
    const r = parseDelimited('name,note\n"Smith, J","said ""hi""\nthen left"\n');
    expect(r.rows).toEqual([["name", "note"], ["Smith, J", 'said "hi"\nthen left']]);
  });

  it("detects tab and semicolon delimiters", () => {
    expect(parseDelimited("a\tb\n1\t2").delimiter).toBe("\t");
    expect(parseDelimited("a;b\n1;2").delimiter).toBe(";");
  });

  it("ignores a delimiter that only appears inside quotes when detecting", () => {
    expect(parseDelimited('"a,b"\tc\n1\t2').delimiter).toBe("\t");
  });

  it("strips a byte order mark", () => {
    expect(parseDelimited("﻿a,b\n1,2").rows[0]).toEqual(["a", "b"]);
  });

  it("returns no rows for empty or blank input", () => {
    expect(parseDelimited("").rows).toEqual([]);
    expect(parseDelimited("\n\n").rows).toEqual([]);
  });

  it("limits rows and columns and says so", () => {
    const wide = Array.from({ length: MAX_COLS + 5 }, (_, i) => `c${i}`).join(",");
    const r = parseDelimited(wide);
    expect(r.rows[0]).toHaveLength(MAX_COLS);
    expect(r.truncated).toBe(true);

    const tall = Array.from({ length: MAX_ROWS + 5 }, (_, i) => `r${i}`).join("\n");
    const t = parseDelimited(tall);
    expect(t.rows).toHaveLength(MAX_ROWS);
    expect(t.truncated).toBe(true);
  });

  it("does not report truncation for small input", () => {
    expect(parseDelimited("a,b\n1,2").truncated).toBe(false);
  });
});

describe("safeCell", () => {
  it("prefixes text starting with = so it cannot run as a formula", () => {
    expect(safeCell("=HYPERLINK(1)")).toBe("'=HYPERLINK(1)");
  });

  it("leaves everything else alone", () => {
    expect(safeCell("hello")).toBe("hello");
    expect(safeCell("-5")).toBe("-5");
    expect(safeCell("")).toBe("");
  });
});
