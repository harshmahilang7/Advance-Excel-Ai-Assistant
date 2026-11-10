export type FailureKind = "quota" | "invalid" | "other";

export class KeysExhaustedError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "KeysExhaustedError";
  }
}

interface PoolOptions {
  readonly keys: readonly string[];
  /** Decides whether an error is key-specific (quota / invalid) or should just be thrown. */
  readonly classify: (err: unknown) => FailureKind;
  readonly now?: () => number;
  readonly quotaCooldownMs?: number;
}

export interface KeyPool {
  readonly size: number;
  /**
   * Runs `op` with the next usable key (round-robin). On a quota or invalid-key failure the key
   * is benched and the next one is tried. `canRetry` can veto failover (e.g. output already streamed).
   */
  run<T>(op: (key: string) => Promise<T>, canRetry?: () => boolean): Promise<T>;
}

const DEFAULT_QUOTA_COOLDOWN_MS = 60_000;

export const createKeyPool = ({
  keys,
  classify,
  now = Date.now,
  quotaCooldownMs = DEFAULT_QUOTA_COOLDOWN_MS,
}: PoolOptions): KeyPool => {
  const unique = [...new Set(keys)];
  const blockedUntil = new Map<string, number>();
  let cursor = 0;

  const pick = (tried: ReadonlySet<string>): string | null => {
    for (let i = 0; i < unique.length; i++) {
      const idx = (cursor + i) % unique.length;
      const key = unique[idx];
      if (key === undefined || tried.has(key) || (blockedUntil.get(key) ?? 0) > now()) continue;
      cursor = (idx + 1) % unique.length;
      return key;
    }
    return null;
  };

  const exhausted = (): KeysExhaustedError => {
    const waits = unique.map((k) => blockedUntil.get(k) ?? 0).filter((t) => Number.isFinite(t) && t > now());
    const hint = waits.length > 0 ? ` Next key frees up in about ${Math.ceil((Math.min(...waits) - now()) / 1000)}s.` : "";
    return new KeysExhaustedError(
      `All ${unique.length} API keys are rate-limited or invalid.${hint} Add another key or try again shortly.`,
    );
  };

  return {
    size: unique.length,
    async run(op, canRetry = () => true) {
      const tried = new Set<string>();
      let lastError: unknown;
      for (;;) {
        const key = pick(tried);
        if (key === null) break;
        tried.add(key);
        try {
          return await op(key);
        } catch (err) {
          const kind = classify(err);
          if (kind === "other") throw err;
          blockedUntil.set(key, kind === "invalid" ? Infinity : now() + quotaCooldownMs);
          lastError = err;
          if (!canRetry()) throw err;
        }
      }
      if (unique.length === 1 && lastError !== undefined) throw lastError;
      throw exhausted();
    },
  };
};
