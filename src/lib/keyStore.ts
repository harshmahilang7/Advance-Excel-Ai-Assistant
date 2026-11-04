const STORAGE_KEY = "griddy.geminiApiKey";

export interface KeyStorage {
  getItem(key: string): string | null;
  setItem(key: string, value: string): void;
  removeItem(key: string): void;
}

export interface KeyStore {
  get(): string | null;
  set(key: string): void;
  clear(): void;
}

/**
 * Stores the user's own Gemini key in local browser storage only.
 * The key never leaves this machine except in requests to Google.
 */
export const createKeyStore = (storage: KeyStorage): KeyStore => ({
  get() {
    try {
      const value = storage.getItem(STORAGE_KEY);
      return value && value.length > 0 ? value : null;
    } catch {
      return null;
    }
  },
  set(key) {
    const trimmed = key.trim();
    if (trimmed.length === 0) throw new Error("API key is empty.");
    try {
      storage.setItem(STORAGE_KEY, trimmed);
    } catch {
      throw new Error("Could not save the API key: browser storage is unavailable.");
    }
  },
  clear() {
    try {
      storage.removeItem(STORAGE_KEY);
    } catch {
      /* nothing stored we can remove; treat as cleared */
    }
  },
});

export const browserKeyStore = (): KeyStore => createKeyStore(window.localStorage);
