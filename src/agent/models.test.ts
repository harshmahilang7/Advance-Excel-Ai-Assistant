import { describe, expect, it } from "vitest";
import { DEFAULT_MODEL_ID, pickDefaultModel, selectExcelModels } from "./models";
import { describeGeminiError } from "./gemini";

const m = (id: string) => ({ id, label: id });

describe("selectExcelModels", () => {
  const everything = [
    m("gemini-2.5-flash"),
    m("gemini-3.1-flash-image"),
    m("gemini-3.8-flash-tts"),
    m("gemini-3.8-live"),
    m("gemini-3.5-flash-lite"),
    m("gemini-embedding-2"),
    m("gemini-4-argon"),
    m("gemini-3.1-pro-preview"),
    m("gemini-3.6-flash"),
    m("gemini-3.8-flash"),
  ];

  it("keeps only the supported Excel models, in preferred order", () => {
    expect(selectExcelModels(everything).map((x) => x.id)).toEqual([
      "gemini-3.8-flash",
      "gemini-3.6-flash",
      "gemini-3.1-pro-preview",
    ]);
  });

  it("drops image, speech, live, lite, embedding and deprecated models", () => {
    const ids = selectExcelModels(everything).map((x) => x.id);
    for (const bad of ["gemini-3.1-flash-image", "gemini-3.8-flash-tts", "gemini-3.8-live", "gemini-3.5-flash-lite", "gemini-embedding-2", "gemini-2.5-flash", "gemini-4-argon"]) {
      expect(ids).not.toContain(bad);
    }
  });

  it("applies friendly labels", () => {
    const first = selectExcelModels(everything)[0];
    expect(first?.label).toMatch(/3\.8 Flash/);
    expect(first?.label).toMatch(/recommended/i);
  });

  it("only returns models the key can actually see", () => {
    expect(selectExcelModels([m("gemini-3.6-flash")]).map((x) => x.id)).toEqual(["gemini-3.6-flash"]);
  });

  it("returns an empty list when none are available", () => {
    expect(selectExcelModels([m("gemini-2.5-flash")])).toEqual([]);
  });

  it("accepts ids with the models/ prefix", () => {
    expect(selectExcelModels([m("models/gemini-3.8-flash")]).map((x) => x.id)).toEqual(["gemini-3.8-flash"]);
  });
});

describe("pickDefaultModel", () => {
  it("defaults to Gemini 3.8 Flash", () => {
    expect(DEFAULT_MODEL_ID).toBe("gemini-3.8-flash");
    expect(pickDefaultModel([m("gemini-3.6-flash"), m("gemini-3.8-flash")])).toBe("gemini-3.8-flash");
  });

  it("falls back to the first available model", () => {
    expect(pickDefaultModel([m("gemini-3.6-flash")])).toBe("gemini-3.6-flash");
  });

  it("returns an empty string for an empty list", () => {
    expect(pickDefaultModel([])).toBe("");
  });
});

describe("describeGeminiError", () => {
  it("reports an invalid key only for real key errors", () => {
    expect(describeGeminiError(new Error("API key not valid. Please pass a valid API key."))).toMatch(/key is not valid/i);
    expect(describeGeminiError(new Error('{"error":{"status":"INVALID_ARGUMENT","details":[{"reason":"API_KEY_INVALID"}]}}'))).toMatch(/key is not valid/i);
  });

  it("does not blame the key for other 400 errors", () => {
    const msg = describeGeminiError(new Error('400 INVALID_ARGUMENT: Function calling is not supported for this model'));
    expect(msg).not.toMatch(/key is not valid/i);
    expect(msg).toMatch(/Function calling is not supported/);
  });

  it("explains an unavailable model", () => {
    expect(describeGeminiError(new Error("404 NOT_FOUND: models/gemini-x is not found"))).toMatch(/model/i);
  });

  it("keeps rate-limit and network messages", () => {
    expect(describeGeminiError(new Error("429 RESOURCE_EXHAUSTED"))).toMatch(/quota|rate/i);
    expect(describeGeminiError(new Error("Failed to fetch"))).toMatch(/network/i);
  });
});
