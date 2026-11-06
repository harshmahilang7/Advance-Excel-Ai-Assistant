import { z } from "zod";
import { MAX_ENTRIES, type ChangeEntry } from "../agent/changeLog";
import type { KeyStorage } from "./keyStore";

const PREFIX = "griddy.changes.";

const schema = z.array(z.object({ command: z.string().max(200), steps: z.array(z.string().max(300)).max(20) })).max(100);

export interface ChangeLogStore {
  load(workbookKey: string): readonly ChangeEntry[];
  save(workbookKey: string, log: readonly ChangeEntry[]): void;
  clear(workbookKey: string): void;
}

/** Per-workbook change log in local browser storage. Never throws; storage may be unavailable. */
export const createChangeLogStore = (storage: KeyStorage): ChangeLogStore => ({
  load(workbookKey) {
    try {
      const raw = storage.getItem(PREFIX + workbookKey);
      if (!raw) return [];
      const parsed = schema.safeParse(JSON.parse(raw));
      return parsed.success ? parsed.data.slice(-MAX_ENTRIES) : [];
    } catch {
      return [];
    }
  },
  save(workbookKey, log) {
    try {
      storage.setItem(PREFIX + workbookKey, JSON.stringify(log));
    } catch {
      /* quota or blocked storage: memory across reloads is best effort */
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

export const browserChangeLogStore = (): ChangeLogStore => createChangeLogStore(window.localStorage);
