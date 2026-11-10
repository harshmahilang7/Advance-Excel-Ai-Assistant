import { GoogleGenAI, type Content, type FunctionDeclaration, type Part } from "@google/genai";
import type { ChatMessage, ModelClient, ModelTurn, RawToolCall } from "./loop";
import { classifyGeminiError, describeGeminiError } from "./gemini";
import { createKeyPool, KeysExhaustedError } from "./keyPool";
import { deltaText, mergeChunks } from "./streamMerge";
import { toolDeclarations } from "./tools";

interface ClientOptions {
  readonly apiKeys: readonly string[];
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

export const createGeminiClient = ({ apiKeys, model, systemInstruction }: ClientOptions): ModelClient => {
  const pool = createKeyPool({ keys: apiKeys, classify: classifyGeminiError });
  const functionDeclarations = toolDeclarations as unknown as FunctionDeclaration[];

  const generateOnce = async (
    apiKey: string,
    history: readonly ChatMessage[],
    signal: AbortSignal | undefined,
    onText: (delta: string) => void,
  ): Promise<ModelTurn> => {
    const ai = new GoogleGenAI({ apiKey });
    const stream = await ai.models.generateContentStream({
      model,
      contents: history.map(toContent),
      config: {
        systemInstruction,
        tools: [{ functionDeclarations }],
        abortSignal: signal,
      },
    });
    const chunks: Content[] = [];
    let reason: string | undefined;
    for await (const chunk of stream) {
      const candidate = chunk.candidates?.[0];
      reason = candidate?.finishReason ?? chunk.promptFeedback?.blockReason ?? reason;
      if (!candidate?.content) continue;
      chunks.push(candidate.content);
      const delta = deltaText(candidate.content);
      if (delta) onText(delta);
    }
    const content = mergeChunks(chunks);
    if (!content) throw new Error(`Gemini returned nothing (${reason ?? "no response"}). Try rephrasing.`);
    return readTurn(content);
  };

  return {
    async generate(history, signal, onText) {
      // Failing over to another key is only safe while nothing has been shown to the user yet.
      let started = false;
      try {
        return await pool.run(
          (apiKey) =>
            generateOnce(apiKey, history, signal, (delta) => {
              started = true;
              onText?.(delta);
            }),
          () => !started && !signal?.aborted,
        );
      } catch (err) {
        if (signal?.aborted || err instanceof KeysExhaustedError) throw err;
        throw new Error(describeGeminiError(err));
      }
    },
  };
};
