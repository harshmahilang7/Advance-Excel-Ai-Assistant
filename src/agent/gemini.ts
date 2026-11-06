import { GoogleGenAI } from "@google/genai";
import { selectExcelModels, type ModelInfo } from "./models";

export type { ModelInfo } from "./models";

const stripPrefix = (name: string): string => name.replace(/^models\//, "");

export const describeGeminiError = (err: unknown): string => {
  const message = err instanceof Error ? err.message : String(err);
  if (/API key not valid|API_KEY_INVALID/i.test(message)) return "That API key is not valid.";
  if (/\b403\b|PERMISSION_DENIED/i.test(message)) return "Key rejected (permission denied).";
  if (/\b429\b|RESOURCE_EXHAUSTED/i.test(message)) return "Rate limit or quota reached. Try again shortly.";
  if (/\b404\b|NOT_FOUND/i.test(message)) return "That model is not available to your key. Pick another model from the list.";
  if (/Failed to fetch|NetworkError/i.test(message)) return "Network error reaching Google. Check your connection.";
  if (/\b400\b|INVALID_ARGUMENT/i.test(message)) return `Gemini rejected the request: ${message}`;
  return `Gemini error: ${message}`;
};

/** True when an error means the key itself is bad (safe to forget the saved key). */
export const isInvalidKeyError = (err: unknown): boolean => {
  const message = err instanceof Error ? err.message : String(err);
  return /API key not valid|API_KEY_INVALID|\b403\b|PERMISSION_DENIED|not valid/i.test(message);
};

/**
 * Validates a Gemini key by listing models, and returns the supported Excel models it can use.
 * Throws a user-friendly Error on failure.
 */
export const listChatModels = async (apiKey: string): Promise<readonly ModelInfo[]> => {
  const ai = new GoogleGenAI({ apiKey });
  const seen: ModelInfo[] = [];
  try {
    const pager = await ai.models.list({ config: { pageSize: 100 } });
    for await (const m of pager) {
      if (!m.name) continue;
      const id = stripPrefix(m.name);
      seen.push({ id, label: m.displayName ?? id });
    }
  } catch (err) {
    throw new Error(describeGeminiError(err));
  }
  const models = selectExcelModels(seen);
  if (models.length === 0) {
    throw new Error("Key accepted, but none of the supported models (Gemini 3.8 Flash, 3.6 Flash, 3.1 Pro) are available to it.");
  }
  return models;
};
