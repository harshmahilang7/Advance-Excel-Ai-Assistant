// @vitest-environment jsdom
import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";

const selectRange = vi.fn<(address: string) => Promise<void>>();
vi.mock("../excel/selectRange", () => ({ selectRange: (address: string) => selectRange(address) }));

import { ReplyText } from "./ReplyText";

afterEach(() => {
  cleanup();
  selectRange.mockReset();
});

describe("ReplyText", () => {
  it("shows plain text unchanged", () => {
    render(<ReplyText text="All done." />);
    expect(screen.getByText("All done.")).toBeTruthy();
    expect(screen.queryAllByRole("button")).toHaveLength(0);
  });

  it("turns backticked cell addresses into buttons without the backticks", () => {
    selectRange.mockResolvedValue();
    render(<ReplyText text="Totals are in `D2:D20` now." />);
    const link = screen.getByRole("button", { name: "D2:D20" });
    expect(link.textContent).toBe("D2:D20");
    expect(document.body.textContent).not.toContain("`");
  });

  it("selects the range when a link is clicked", () => {
    selectRange.mockResolvedValue();
    render(<ReplyText text="See `'Q1 Sales'!B2:C9`." />);
    fireEvent.click(screen.getByRole("button", { name: "'Q1 Sales'!B2:C9" }));
    expect(selectRange).toHaveBeenCalledWith("'Q1 Sales'!B2:C9");
  });

  it("does not crash when selecting fails", async () => {
    selectRange.mockRejectedValue(new Error("sheet deleted"));
    render(<ReplyText text="See `A1`." />);
    fireEvent.click(screen.getByRole("button", { name: "A1" }));
    await Promise.resolve();
    expect(screen.getByRole("button", { name: "A1" })).toBeTruthy();
  });

  it("does not link function names or words like Q3", () => {
    render(<ReplyText text="Use `SUM` for Q3 and FY26." />);
    expect(screen.queryAllByRole("button")).toHaveLength(0);
  });
});
