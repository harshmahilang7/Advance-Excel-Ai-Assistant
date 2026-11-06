import type { Content, Part } from "@google/genai";

/** Visible (non-thought) text carried by one streamed chunk. */
export const deltaText = (chunk: Content | undefined): string =>
  (chunk?.parts ?? [])
    .filter((p) => typeof p.text === "string" && !p.thought)
    .map((p) => p.text)
    .join("");

/** A part that is only plain text: safe to join with a neighbour without losing metadata. */
const isPlainText = (p: Part): p is Part & { text: string } => {
  const keys = Object.keys(p);
  return keys.length === 1 && keys[0] === "text" && typeof p.text === "string";
};

/**
 * Rebuilds one model turn from streamed chunks. Parts keep their order and metadata (thought
 * signatures, function calls); only adjacent plain-text fragments are joined. Pure.
 */
export const mergeChunks = (chunks: readonly (Content | undefined)[]): Content | undefined => {
  const parts = chunks.flatMap((c) => c?.parts ?? []);
  if (parts.length === 0) return undefined;
  const merged = parts.reduce<readonly Part[]>((acc, part) => {
    const last = acc.at(-1);
    if (last && isPlainText(last) && isPlainText(part)) return [...acc.slice(0, -1), { text: last.text + part.text }];
    return [...acc, part];
  }, []);
  return { role: "model", parts: [...merged] };
};
