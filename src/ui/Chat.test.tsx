// @vitest-environment jsdom
import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { ChatMessage, ModelTurn, ToolResult } from "../agent/loop";

const state = vi.hoisted(() => ({
  excel: true,
  generate: vi.fn(),
  execute: vi.fn(),
  importTable: vi.fn(),
}));

vi.mock("../excel/readSelection", () => ({
  isExcelHost: () => state.excel,
  readSelection: () => Promise.reject(new Error("no selection in tests")),
}));
vi.mock("../excel/liveContext", () => ({ buildLiveContext: async () => "CTX" }));
vi.mock("../excel/executors", () => ({ createExcelExecutor: () => state.execute }));
vi.mock("../agent/geminiClient", () => ({ createGeminiClient: () => ({ generate: state.generate }) }));
vi.mock("../excel/importTable", () => ({ importTable: (...args: unknown[]) => state.importTable(...args) }));

import { Chat } from "./Chat";

const PLACEHOLDER = "e.g. Add a column that totals each row";

const okTool: ToolResult = { ok: true, summary: "Wrote 1x1 cells at Sheet1!E1", data: null };
const writeCall = { name: "write_cells", args: { address: "E1", data: [["Total"]] } };

const renderChat = () => render(<Chat apiKeys={["key-1"]} modelId="model-1" />);
const textbox = () => screen.getByPlaceholderText(PLACEHOLDER) as HTMLTextAreaElement;
const type = (value: string) => fireEvent.change(textbox(), { target: { value } });
const send = () => fireEvent.click(screen.getByRole("button", { name: "Send" }));

/** A model that makes one write call, then replies with `reply`. */
const writeThenReply = (reply: string) => {
  let turn = 0;
  state.generate.mockImplementation(async (): Promise<ModelTurn> => {
    turn += 1;
    return turn === 1 ? { text: "", calls: [writeCall] } : { text: reply, calls: [] };
  });
};

beforeEach(() => {
  state.excel = true;
  state.execute.mockResolvedValue(okTool);
});

afterEach(() => {
  cleanup();
  window.localStorage.clear();
  state.generate.mockReset();
  state.execute.mockReset();
  state.importTable.mockReset();
});

describe("Chat start screen", () => {
  it("offers starter, review and (once saved) template prompts that fill the text box", async () => {
    renderChat();
    const starter = await screen.findByText("Add a Total column that multiplies Qty by Price");
    fireEvent.click(starter);
    expect(textbox().value).toBe("Add a Total column that multiplies Qty by Price");

    expect(screen.getByText("Or check your sheet without changing it:")).toBeTruthy();
    fireEvent.click(screen.getByText(/^Audit this sheet for inconsistent formulas/));
    expect(textbox().value).toMatch(/Do not change anything/);
  });

  it("keeps Send disabled until there is text", () => {
    renderChat();
    const sendButton = screen.getByRole("button", { name: "Send" }) as HTMLButtonElement;
    expect(sendButton.disabled).toBe(true);
    type("hello");
    expect(sendButton.disabled).toBe(false);
  });
});

describe("Chat layout", () => {
  it("lets the chat area shrink and scroll while the footer keeps its size", () => {
    const { container } = renderChat();
    const root = container.firstElementChild as HTMLElement;
    expect(root.style.overflow).toBe("hidden");
    const scroll = screen.getByTestId("chat-scroll");
    expect(scroll.style.overflowY).toBe("auto");
    expect(scroll.style.minHeight).toBe("60px");
    expect(screen.getByTestId("chat-footer").style.flexShrink).toBe("0");
  });
});

describe("Chat sending", () => {
  it("explains when it is not running inside Excel", async () => {
    state.excel = false;
    renderChat();
    type("add a total");
    send();
    expect(await screen.findByText(/Not running inside Excel/)).toBeTruthy();
    expect(state.generate).not.toHaveBeenCalled();
  });

  it("runs a request: tool step, clickable reply addresses, run summary and read log", async () => {
    writeThenReply("Added a Total column in `E1:E6`.");
    renderChat();
    type("add a total");
    send();

    expect(await screen.findByText("✓ Wrote 1x1 cells at Sheet1!E1")).toBeTruthy();
    expect(await screen.findByRole("button", { name: "E1:E6" })).toBeTruthy();
    const note = await screen.findByText(/Worked for/);
    expect(note.textContent).toMatch(/1 step · 1 change/);
    expect(note.textContent).toMatch(/Sent to the AI: your selection summary only/);
    expect(state.execute).toHaveBeenCalledWith("write_cells", expect.objectContaining({ address: "E1" }));
  });

  it("lists the ranges the model read in the read log", async () => {
    let turn = 0;
    state.generate.mockImplementation(async (): Promise<ModelTurn> => {
      turn += 1;
      return turn === 1 ? { text: "", calls: [{ name: "read_range", args: { address: "Sheet1!A1:D6" } }] } : { text: "Looks fine.", calls: [] };
    });
    renderChat();
    type("what is here");
    send();
    const note = await screen.findByText(/Sent to the AI/);
    expect(note.textContent).toContain("Sheet1!A1:D6");
  });

  it("shows the Stop button while running and reports Stopped", async () => {
    state.generate.mockImplementation(
      (_h: readonly ChatMessage[], signal?: AbortSignal) =>
        new Promise<ModelTurn>((_resolve, reject) => {
          signal?.addEventListener("abort", () => reject(new Error("aborted")));
        }),
    );
    renderChat();
    type("add a total");
    send();
    fireEvent.click(await screen.findByRole("button", { name: "Stop" }));
    expect(await screen.findByText("Stopped.")).toBeTruthy();
  });

  it("offers Continue after the step limit and resumes with the continue prompt", async () => {
    state.generate.mockImplementation(async (history: readonly ChatMessage[]): Promise<ModelTurn> => {
      const last = history.at(-1);
      if (last?.role === "user" && last.text.includes("Continue the previous task")) return { text: "Finished.", calls: [] };
      return { text: "", calls: [{ name: "read_range", args: { address: "A1" } }] };
    });
    renderChat();
    type("do a big task");
    send();
    expect(await screen.findByText(/Stopped after the step limit/)).toBeTruthy();
    fireEvent.click(await screen.findByRole("button", { name: "Continue" }));
    expect(await screen.findByText("Finished.")).toBeTruthy();
    await waitFor(() => expect(screen.queryByRole("button", { name: "Continue" })).toBeNull());
  });
});

describe("Chat preview mode", () => {
  const enablePreview = () => fireEvent.click(screen.getByRole("switch", { name: /Preview before applying/ }));

  it("queues writes without running them, and Discard drops them", async () => {
    writeThenReply("I would add a Total column.");
    renderChat();
    enablePreview();
    type("add a total");
    send();
    expect(await screen.findByText(/Review 1 queued change\(s\)/)).toBeTruthy();
    expect(state.execute).not.toHaveBeenCalled();
    fireEvent.click(screen.getByRole("button", { name: "Discard" }));
    expect(await screen.findByText(/Discarded 1 queued change/)).toBeTruthy();
    expect(state.execute).not.toHaveBeenCalled();
  });

  it("Apply runs the queued writes", async () => {
    writeThenReply("I would add a Total column.");
    renderChat();
    enablePreview();
    type("add a total");
    send();
    fireEvent.click(await screen.findByRole("button", { name: "Apply" }));
    await waitFor(() => expect(state.execute).toHaveBeenCalledWith("write_cells", expect.objectContaining({ address: "E1" })));
  });
});

describe("Chat templates", () => {
  it("saves the typed command, then shows it on the start screen where it can be reused and deleted", async () => {
    renderChat();
    type("Bold the header row");
    fireEvent.click(screen.getByRole("button", { name: "Save template" }));
    expect(await screen.findByText(/Saved as a template/)).toBeTruthy();
    expect(window.localStorage.getItem("griddy.templates")).toContain("Bold the header row");

    fireEvent.click(screen.getByRole("button", { name: "Clear chat" }));
    expect(await screen.findByText("Your templates:")).toBeTruthy();
    fireEvent.click(screen.getByRole("button", { name: "Bold the header row" }));
    expect(textbox().value).toBe("Bold the header row");

    fireEvent.click(screen.getByRole("button", { name: "Delete template Bold the header row" }));
    await waitFor(() => expect(screen.queryByText("Your templates:")).toBeNull());
  });

  it("tells you when there is nothing to save or it is already saved", async () => {
    renderChat();
    const save = screen.getByRole("button", { name: "Save template" }) as HTMLButtonElement;
    expect(save.disabled).toBe(true);
    type("Sort by Qty");
    fireEvent.click(save);
    await screen.findByText(/Saved as a template/);
    fireEvent.click(save);
    expect(await screen.findByText(/already saved/)).toBeTruthy();
  });
});

describe("Chat CSV import", () => {
  const chooseFile = (name: string, content: string) => {
    const input = document.querySelector('input[type="file"]') as HTMLInputElement;
    const file = new File([content], name, { type: "text/csv" });
    Object.defineProperty(file, "text", { value: () => Promise.resolve(content) });
    fireEvent.change(input, { target: { files: [file] } });
  };

  it("imports a file into a new sheet and reports it", async () => {
    state.importTable.mockResolvedValue({ sheet: "sales", rows: 2, columns: 2 });
    renderChat();
    chooseFile("sales.csv", "a,b\n1,2");
    expect(await screen.findByText(/Imported 2 rows × 2 columns into a new sheet "sales"/)).toBeTruthy();
    expect(state.importTable).toHaveBeenCalledWith([["a", "b"], ["1", "2"]], "sales.csv", expect.anything());
  });

  it("reports an empty file instead of importing", async () => {
    renderChat();
    chooseFile("empty.csv", "\n\n");
    expect(await screen.findByText(/no rows to import/)).toBeTruthy();
    expect(state.importTable).not.toHaveBeenCalled();
  });

  it("reports import failures in plain words", async () => {
    state.importTable.mockRejectedValue(new Error("Something broke"));
    renderChat();
    chooseFile("bad.csv", "a,b\n1,2");
    expect(await screen.findByText("Something broke")).toBeTruthy();
  });
});
