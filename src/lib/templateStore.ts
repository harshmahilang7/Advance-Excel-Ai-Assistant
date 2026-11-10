import { z } from "zod";
import type { KeyStorage } from "./keyStore";

/** Saved commands the user can rerun on any workbook. Kept in local browser storage. */

export const MAX_TEMPLATES = 20;
const MAX_PROMPT_CHARS = 2000;
const NAME_CHARS = 40;
const STORAGE_KEY = "griddy.templates";

const templateSchema = z.object({
  id: z.string().min(1).max(40),
  name: z.string().min(1).max(NAME_CHARS),
  prompt: z.string().min(1).max(MAX_PROMPT_CHARS),
});

export type Template = z.infer<typeof templateSchema>;

export interface TemplateStore {
  load(): readonly Template[];
  save(templates: readonly Template[]): void;
}

export type AddStatus = "added" | "duplicate" | "empty" | "full";

export interface AddResult {
  readonly status: AddStatus;
  readonly list: readonly Template[];
}

const tidy = (text: string): string => text.trim().replace(/\s+/g, " ");

export const nameFor = (prompt: string): string => {
  const clean = tidy(prompt);
  return clean.length > NAME_CHARS ? `${clean.slice(0, NAME_CHARS - 1)}…` : clean;
};

let counter = 0;
const defaultId = (): string => {
  counter += 1;
  return `t${Date.now().toString(36)}${counter}`;
};

export const addTemplate = (list: readonly Template[], prompt: string, makeId: () => string = defaultId): AddResult => {
  const clean = tidy(prompt).slice(0, MAX_PROMPT_CHARS);
  if (clean.length === 0) return { status: "empty", list };
  if (list.some((t) => tidy(t.prompt) === clean)) return { status: "duplicate", list };
  if (list.length >= MAX_TEMPLATES) return { status: "full", list };
  return { status: "added", list: [...list, { id: makeId(), name: nameFor(clean), prompt: clean }] };
};

export const removeTemplate = (list: readonly Template[], id: string): readonly Template[] => list.filter((t) => t.id !== id);

export const createTemplateStore = (storage: KeyStorage): TemplateStore => ({
  load() {
    try {
      const raw = storage.getItem(STORAGE_KEY);
      if (!raw) return [];
      const parsed = z.array(templateSchema).safeParse(JSON.parse(raw));
      return parsed.success ? parsed.data.slice(0, MAX_TEMPLATES) : [];
    } catch {
      return [];
    }
  },
  save(templates) {
    try {
      storage.setItem(STORAGE_KEY, JSON.stringify(templates.slice(0, MAX_TEMPLATES)));
    } catch {
      /* storage blocked or full: templates are a convenience, so fail quietly */
    }
  },
});

export const browserTemplateStore = (): TemplateStore => createTemplateStore(window.localStorage);
