import type { ChatMessage } from "./loop";

/**
 * Keep only the last `keepUserTurns` user requests (and everything after them).
 * Cutting on a user message keeps model tool-calls paired with their tool results.
 */
export const trimHistory = (history: readonly ChatMessage[], keepUserTurns: number): readonly ChatMessage[] => {
  const userIndexes = history.flatMap((m, i) => (m.role === "user" ? [i] : []));
  if (userIndexes.length <= keepUserTurns) return history;
  const start = userIndexes[userIndexes.length - keepUserTurns] ?? 0;
  return history.slice(start);
};
