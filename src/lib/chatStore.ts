import { z } from "zod";
import type { KeyStorage } from "./keyStore";

export const MAX_STORED_LINES = 100;
const MAX_TEXT = 4000;
const PREFIX = "griddy.chat.";

const text = z.string().max(MAX_TEXT);

const lineSchema = z.discriminatedUnion("kind", [
  z.object({ id: z.number().int(), kind: z.enum(["user", "assistant", "note"]), text }),
  z.object({ id: z.number().int(), kind: z.literal("action"), text, ok: z.boolean() }),
  z.object({ id: z.number().int(), kind: z.literal("error"), text }),
]);

export type Line = z.infer<typeof lineSchema>;

export interface ChatStore {
  load(workbookKey: string): readonly Line[];
  save(workbookKey: string, lines: readonly Line[]): void;
  clear(workbookKey: string): void;
}

const clip = (line: Line): Line => (line.text.length > MAX_TEXT ? { ...line, text: line.text.slice(0, MAX_TEXT) } : line);

/** Per-workbook chat display history in local browser storage. Never throws; storage may be unavailable. */
export const createChatStore = (storage: KeyStorage): ChatStore => ({
  load(workbookKey) {
    try {
      const raw = storage.getItem(PREFIX + workbookKey);
      if (!raw) return [];
      const parsed = z.array(lineSchema).safeParse(JSON.parse(raw));
      return parsed.success ? parsed.data.slice(-MAX_STORED_LINES) : [];
    } catch {
      return [];
    }
  },
  save(workbookKey, lines) {
    try {
      storage.setItem(PREFIX + workbookKey, JSON.stringify(lines.slice(-MAX_STORED_LINES).map(clip)));
    } catch {
      /* quota or blocked storage: persistence is best effort */
    }
  },
  clear(workbookKey) {
    try {
      storage.removeItem(PREFIX + workbookKey);
    } catch {
      /* nothing to clear */
    }
  },
});

export const browserChatStore = (): ChatStore => createChatStore(window.localStorage);
