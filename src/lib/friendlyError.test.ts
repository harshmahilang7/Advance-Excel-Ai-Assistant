import { describe, expect, it } from "vitest";
import { friendlyError } from "./friendlyError";

const officeError = (code: string, message: string): Error => Object.assign(new Error(message), { code });

describe("friendlyError", () => {
  it("explains protected sheets", () => {
    expect(friendlyError(officeError("AccessDenied", "Access denied"))).toMatch(/protected/i);
    expect(friendlyError(new Error("The sheet is protected and cannot be changed"))).toMatch(/protected/i);
  });

  it("explains missing items", () => {
    expect(friendlyError(officeError("ItemNotFound", "The requested resource doesn't exist."))).toMatch(/could not find/i);
  });

  it("explains merged cells", () => {
    expect(friendlyError(officeError("InvalidOperation", "This operation is not allowed on merged cells"))).toMatch(/merged/i);
  });

  it("explains duplicate names", () => {
    expect(friendlyError(officeError("ItemAlreadyExists", "A table with that name already exists"))).toMatch(/already/i);
  });

  it("explains old Excel versions", () => {
    expect(friendlyError(officeError("ApiNotFound", "x"))).toMatch(/update excel|newer/i);
  });

  it("explains network failures", () => {
    expect(friendlyError(new TypeError("Failed to fetch"))).toMatch(/network|connection/i);
  });

  it("keeps unknown messages unchanged", () => {
    expect(friendlyError(new Error("Something odd"))).toBe("Something odd");
  });

  it("handles non-Error values", () => {
    expect(friendlyError("boom")).toBe("boom");
    expect(friendlyError(undefined)).toBe("Something went wrong.");
  });
});
