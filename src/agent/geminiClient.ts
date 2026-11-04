import { GoogleGenAI, type Content, type FunctionDeclaration, type Part } from "@google/genai";
import type { ChatMessage, ModelClient, ModelTurn, RawToolCall } from "./loop";
import { describeGeminiError } from "./gemini";
import { toolDeclarations } from "./tools";

interface ClientOptions {
  readonly apiKey: string;
  readonly model: string;
  readonly systemInstruction: string;
}

const toContent = (message: ChatMessage): Content => {
  switch (message.role) {
    case "user":
      return { role: "user", parts: [{ text: message.text }] };
    case "model": {
      // Replay the provider-native turn verbatim so thought signatures stay valid.
      if (message.raw) return message.raw as Content;
      const parts: Part[] = [
        ...(message.text ? [{ text: message.text }] : []),
        ...message.calls.map((c) => ({ functionCall: { name: c.name, args: c.args as Record<string, unknown> } })),
      ];
      return { role: "model", parts };
    }
    case "tool":
      return {
        role: "user",
        parts: message.results.map((r) => ({ functionResponse: { name: r.name, response: r.response } })),
      };
  }
};

const readTurn = (content: Content | undefined): ModelTurn => {
  const parts = content?.parts ?? [];
  const text = parts
    .filter((p) => typeof p.text === "string" && !p.thought)
    .map((p) => p.text)
    .join("");
  const calls: RawToolCall[] = parts.flatMap((p) =>
    p.functionCall?.name ? [{ name: p.functionCall.name, args: p.functionCall.args ?? {} }] : [],
  );
  return { text, calls, raw: content };
};

export const createGeminiClient = ({ apiKey, model, systemInstruction }: ClientOptions): ModelClient => {
  const ai = new GoogleGenAI({ apiKey });
  const functionDeclarations = toolDeclarations as unknown as FunctionDeclaration[];

  return {
    async generate(history, signal) {
      try {
        const response = await ai.models.generateContent({
          model,
          contents: history.map(toContent),
          config: {
            systemInstruction,
            tools: [{ functionDeclarations }],
            abortSignal: signal,
          },
        });
        const candidate = response.candidates?.[0];
        if (!candidate?.content) {
          const reason = candidate?.finishReason ?? response.promptFeedback?.blockReason ?? "no response";
          throw new Error(`Gemini returned nothing (${reason}). Try rephrasing.`);
        }
        return readTurn(candidate.content);
      } catch (err) {
        if (signal?.aborted) throw err;
        throw new Error(describeGeminiError(err));
      }
    },
  };
};
