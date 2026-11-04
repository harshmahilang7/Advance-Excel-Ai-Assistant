import { GoogleGenAI } from "@google/genai";

export interface ModelInfo {
  readonly id: string;
  readonly label: string;
}

const stripPrefix = (name: string): string => name.replace(/^models\//, "");

export const describeGeminiError = (err: unknown): string => {
  const message = err instanceof Error ? err.message : String(err);
  if (/API key not valid|API_KEY_INVALID|\b400\b/i.test(message)) return "That API key is not valid.";
  if (/\b403\b|PERMISSION_DENIED/i.test(message)) return "Key rejected (permission denied).";
  if (/\b429\b|RESOURCE_EXHAUSTED/i.test(message)) return "Rate limit or quota reached. Try again shortly.";
  if (/Failed to fetch|NetworkError/i.test(message)) return "Network error reaching Google. Check your connection.";
  return `Gemini error: ${message}`;
};

/**
 * Validates a Gemini key by listing models, and returns the models usable for chat.
 * Throws a user-friendly Error on failure.
 */
export const listChatModels = async (apiKey: string): Promise<readonly ModelInfo[]> => {
  const ai = new GoogleGenAI({ apiKey });
  const models: ModelInfo[] = [];
  try {
    const pager = await ai.models.list({ config: { pageSize: 100 } });
    for await (const m of pager) {
      const supportsChat = m.supportedActions?.includes("generateContent") ?? false;
      if (!m.name || !supportsChat) continue;
      const id = stripPrefix(m.name);
      if (!id.startsWith("gemini")) continue;
      models.push({ id, label: m.displayName ?? id });
    }
  } catch (err) {
    throw new Error(describeGeminiError(err));
  }
  if (models.length === 0) throw new Error("Key accepted but no Gemini chat models are available.");
  return models;
};
