import { describe, expect, it } from "vitest";
import { errorCodeOf, formatToolError } from "./toolError";

describe("formatToolError", () => {
  it("puts the code first and the hint last", () => {
    expect(formatToolError("DESTINATION_OCCUPIED", "D2 already has data.", "Pick an empty cell.")).toBe(
      "DESTINATION_OCCUPIED: D2 already has data. Try: Pick an empty cell.",
    );
  });

  it("omits the hint when none is given", () => {
    expect(formatToolError("EXCEL_TOO_OLD", "Needs Excel API 1.8.")).toBe("EXCEL_TOO_OLD: Needs Excel API 1.8.");
  });
});

describe("errorCodeOf", () => {
  it("reads the code back from a formatted error", () => {
    expect(errorCodeOf(formatToolError("SOURCE_EMPTY", "Nothing there."))).toBe("SOURCE_EMPTY");
  });

  it("returns null for free text", () => {
    expect(errorCodeOf("Something went wrong.")).toBeNull();
    expect(errorCodeOf("")).toBeNull();
  });
});
