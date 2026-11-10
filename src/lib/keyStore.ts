const STORAGE_KEY = "griddy.geminiApiKeys";
const LEGACY_STORAGE_KEY = "griddy.geminiApiKey";

export interface KeyStorage {
  getItem(key: string): string | null;
  setItem(key: string, value: string): void;
  removeItem(key: string): void;
}

export interface KeyStore {
  list(): readonly string[];
  save(keys: readonly string[]): void;
  clear(): void;
}

/** Splits pasted text (one per line, or comma/space separated) into trimmed, de-duplicated keys. */
export const parseKeys = (text: string): readonly string[] => [
  ...new Set(text.split(/[\s,;]+/).map((k) => k.trim()).filter((k) => k.length > 0)),
];

const readList = (storage: KeyStorage): readonly string[] => {
  const raw = storage.getItem(STORAGE_KEY);
  if (raw) {
    const parsed: unknown = JSON.parse(raw);
    if (Array.isArray(parsed)) return parsed.filter((k): k is string => typeof k === "string" && k.length > 0);
  }
  const legacy = storage.getItem(LEGACY_STORAGE_KEY);
  return legacy ? [legacy] : [];
};

/**
 * Stores the user's own Gemini keys in local browser storage only.
 * Keys never leave this machine except in requests to Google.
 */
export const createKeyStore = (storage: KeyStorage): KeyStore => ({
  list() {
    try {
      return readList(storage);
    } catch {
      return [];
    }
  },
  save(keys) {
    const cleaned = parseKeys(keys.join("\n"));
    if (cleaned.length === 0) throw new Error("API key is empty.");
    try {
      storage.setItem(STORAGE_KEY, JSON.stringify(cleaned));
      storage.removeItem(LEGACY_STORAGE_KEY);
    } catch {
      throw new Error("Could not save the API keys: browser storage is unavailable.");
    }
  },
  clear() {
    try {
      storage.removeItem(STORAGE_KEY);
      storage.removeItem(LEGACY_STORAGE_KEY);
    } catch {
      /* nothing stored we can remove; treat as cleared */
    }
  },
});

export const browserKeyStore = (): KeyStore => createKeyStore(window.localStorage);
