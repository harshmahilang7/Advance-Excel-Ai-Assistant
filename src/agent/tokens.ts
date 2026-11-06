import type { ChatMessage } from "./loop";

/** Above this, suggest clearing the chat: every request re-sends the whole history. */
export const LARGE_CONTEXT_TOKENS = 20000;

const CHARS_PER_TOKEN = 4;

/** Rough estimate (about 4 characters per token). Good enough for a cost hint, not for billing. */
export const estimateTokens = (text: string): number => Math.ceil(text.length / CHARS_PER_TOKEN);

const messageChars = (m: ChatMessage): string => {
  switch (m.role) {
    case "user":
      return m.text;
    case "model":
      // `raw` is the provider payload (thought signatures etc.); it is not what the user reads or pays for here.
      return m.text + JSON.stringify(m.calls);
    case "tool":
      return JSON.stringify(m.results);
  }
};

export const estimateHistoryTokens = (history: readonly ChatMessage[]): number =>
  history.reduce((sum, m) => sum + estimateTokens(messageChars(m)), 0);

export const formatTokens = (n: number): string => {
  if (n < 1000) return String(n);
  const k = n / 1000;
  return `${k >= 10 ? Math.round(k) : Math.round(k * 10) / 10}k`;
};
