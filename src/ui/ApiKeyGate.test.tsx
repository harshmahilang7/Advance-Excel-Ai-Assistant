// @vitest-environment jsdom
import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";

const validateKeys = vi.fn();
vi.mock("../agent/gemini", () => ({ validateKeys: (keys: readonly string[]) => validateKeys(keys) }));

import { ApiKeyGate } from "./ApiKeyGate";

const KEY_A = "AIzaSyAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA";
const KEY_B = "AIzaSyBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBB";

afterEach(() => {
  cleanup();
  validateKeys.mockReset();
});

const typeKeys = (value: string) => fireEvent.change(screen.getByPlaceholderText(/AIza/), { target: { value } });

describe("ApiKeyGate", () => {
  it("disables Save until a key is entered", () => {
    render(<ApiKeyGate onReady={vi.fn()} saveKeys={vi.fn()} />);
    const save = screen.getByRole("button", { name: "Save and continue" }) as HTMLButtonElement;
    expect(save.disabled).toBe(true);
    typeKeys(KEY_A);
    expect(save.disabled).toBe(false);
  });

  it("counts one key per line", () => {
    render(<ApiKeyGate onReady={vi.fn()} saveKeys={vi.fn()} />);
    typeKeys(`${KEY_A}\n${KEY_B}`);
    expect(screen.getByText("API keys (2)")).toBeTruthy();
  });

  it("validates, saves and reports ready", async () => {
    validateKeys.mockResolvedValue({ keep: [KEY_A], rejected: [], models: [{ id: "m1", label: "Model 1" }] });
    const onReady = vi.fn();
    const saveKeys = vi.fn();
    render(<ApiKeyGate onReady={onReady} saveKeys={saveKeys} />);
    typeKeys(KEY_A);
    fireEvent.click(screen.getByRole("button", { name: "Save and continue" }));
    await waitFor(() => expect(onReady).toHaveBeenCalled());
    expect(saveKeys).toHaveBeenCalledWith([KEY_A]);
    expect(onReady).toHaveBeenCalledWith([KEY_A], [{ id: "m1", label: "Model 1" }]);
  });

  it("shows a masked error and does not continue when a key is rejected", async () => {
    validateKeys.mockResolvedValue({ keep: [], rejected: [KEY_A], models: [] });
    const onReady = vi.fn();
    render(<ApiKeyGate onReady={onReady} saveKeys={vi.fn()} />);
    typeKeys(KEY_A);
    fireEvent.click(screen.getByRole("button", { name: "Save and continue" }));
    const message = await screen.findByText(/Not valid:/);
    expect(message.textContent).toContain("AIzaSy...AAAA");
    expect(message.textContent).not.toContain(KEY_A);
    expect(onReady).not.toHaveBeenCalled();
  });

  it("shows the error when validation fails outright", async () => {
    validateKeys.mockRejectedValue(new Error("Network problem"));
    render(<ApiKeyGate onReady={vi.fn()} saveKeys={vi.fn()} />);
    typeKeys(KEY_A);
    fireEvent.click(screen.getByRole("button", { name: "Save and continue" }));
    expect(await screen.findByText("Network problem")).toBeTruthy();
  });
});
