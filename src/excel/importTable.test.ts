import { describe, expect, it } from "vitest";
import { sheetNameFor } from "./importTable";

describe("sheetNameFor", () => {
  it("uses the file name without its extension", () => {
    expect(sheetNameFor("sales 2024.csv", [])).toBe("sales 2024");
  });

  it("replaces characters Excel forbids in sheet names", () => {
    expect(sheetNameFor("a/b:c[1]?.csv", [])).toBe("a b c 1");
  });

  it("falls back to Import for an empty name", () => {
    expect(sheetNameFor(".csv", [])).toBe("Import");
  });

  it("avoids existing sheet names, ignoring case", () => {
    expect(sheetNameFor("Sales.csv", ["sales"])).toBe("Sales (2)");
    expect(sheetNameFor("Sales.csv", ["sales", "Sales (2)"])).toBe("Sales (3)");
  });

  it("keeps names within 31 characters", () => {
    const long = sheetNameFor("x".repeat(60) + ".csv", []);
    expect(long).toHaveLength(31);
    expect(sheetNameFor("x".repeat(60) + ".csv", [long])).toHaveLength(31);
  });
});
