// @vitest-environment jsdom
import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";

const validateKeys = vi.fn();
vi.mock("../agent/gemini", () => ({ validateKeys: (keys: readonly string[]) => validateKeys(keys) }));
vi.mock("./Chat", () => ({ Chat: () => <div>chat-stub</div> }));
vi.mock("./SelectionPanel", () => ({ SelectionPanel: () => <div>selection-stub</div> }));

import { browserKeyStore } from "../lib/keyStore";
import { App } from "./App";

const KEY = "AIzaSyAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA";
const MODELS = [
  { id: "model-one", label: "Model One" },
  { id: "model-two", label: "Model Two" },
];

afterEach(() => {
  cleanup();
  window.localStorage.clear();
  validateKeys.mockReset();
});

const withSavedKey = () => {
  browserKeyStore().save([KEY]);
  validateKeys.mockResolvedValue({ keep: [KEY], rejected: [], models: MODELS });
};

const openMenu = async () => {
  fireEvent.click(await screen.findByRole("button", { name: "Settings" }));
};

describe("App", () => {
  it("asks for API keys when none are saved", () => {
    render(<App />);
    expect(screen.getByText("Enter your Gemini API keys")).toBeTruthy();
    expect(validateKeys).not.toHaveBeenCalled();
  });

  it("restores saved keys and shows the chat with a model picker", async () => {
    withSavedKey();
    render(<App />);
    expect(await screen.findByText("chat-stub")).toBeTruthy();
    expect(validateKeys).toHaveBeenCalledWith([KEY]);
    expect(screen.getByRole("combobox").getAttribute("aria-label")).toBe("Model");
  });

  it("does not repeat the pane title, and keeps the page from overflowing", async () => {
    withSavedKey();
    const { container } = render(<App />);
    await screen.findByText("chat-stub");
    expect(screen.queryByText("Excel-Ai-Assistant")).toBeNull();
    const root = container.firstElementChild as HTMLElement;
    expect(root.style.overflow).toBe("hidden");
    expect(root.style.height).toBe("100vh");
  });

  it("keeps key management in the settings menu", async () => {
    withSavedKey();
    render(<App />);
    await openMenu();
    expect(await screen.findByRole("menuitem", { name: "Manage API keys (1)" })).toBeTruthy();
    expect(screen.getByRole("menuitem", { name: "Show selection summary" })).toBeTruthy();
    expect(screen.getByRole("menuitem", { name: "Copy diagnostics" })).toBeTruthy();
    expect(screen.getByRole("menuitem", { name: "Remove all API keys" })).toBeTruthy();
    expect(screen.queryByText("selection-stub")).toBeNull();
  });

  it("toggles the selection summary", async () => {
    withSavedKey();
    render(<App />);
    await openMenu();
    fireEvent.click(await screen.findByRole("menuitem", { name: "Show selection summary" }));
    expect(await screen.findByText("selection-stub")).toBeTruthy();
    await openMenu();
    expect(await screen.findByRole("menuitem", { name: "Hide selection summary" })).toBeTruthy();
  });

  it("opens the key screen with the saved key from Manage API keys", async () => {
    withSavedKey();
    render(<App />);
    await openMenu();
    fireEvent.click(await screen.findByRole("menuitem", { name: "Manage API keys (1)" }));
    expect(await screen.findByText("Enter your Gemini API keys")).toBeTruthy();
    expect((screen.getByPlaceholderText(/AIza/) as HTMLTextAreaElement).value).toBe(KEY);
  });

  it("forgets the keys and returns to the key screen from Remove all API keys", async () => {
    withSavedKey();
    render(<App />);
    await openMenu();
    fireEvent.click(await screen.findByRole("menuitem", { name: "Remove all API keys" }));
    expect(await screen.findByText("Enter your Gemini API keys")).toBeTruthy();
    await waitFor(() => expect(browserKeyStore().list()).toEqual([]));
  });
});
