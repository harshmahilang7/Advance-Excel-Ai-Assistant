import { describe, expect, it } from "vitest";
import { createKeyPool, KeysExhaustedError, type FailureKind } from "./keyPool";

const classify = (err: unknown): FailureKind => {
  const m = err instanceof Error ? err.message : "";
  if (m === "quota") return "quota";
  if (m === "invalid") return "invalid";
  return "other";
};

const clock = () => {
  let t = 1_000;
  return { now: () => t, advance: (ms: number) => void (t += ms) };
};

const pool = (keys: string[], c = clock()) => ({ c, pool: createKeyPool({ keys, classify, now: c.now }) });

describe("keyPool", () => {
  it("round-robins across keys on successive calls", async () => {
    const { pool: p } = pool(["a", "b", "c"]);
    const used: string[] = [];
    for (let i = 0; i < 4; i++) await p.run(async (k) => void used.push(k));
    expect(used).toEqual(["a", "b", "c", "a"]);
  });

  it("switches to the next key when one hits quota", async () => {
    const { pool: p } = pool(["a", "b"]);
    const used: string[] = [];
    const result = await p.run(async (k) => {
      used.push(k);
      if (k === "a") throw new Error("quota");
      return "ok";
    });
    expect(result).toBe("ok");
    expect(used).toEqual(["a", "b"]);
  });

  it("skips a rate-limited key until its cooldown passes", async () => {
    const { pool: p, c } = pool(["a", "b"]);
    await p.run(async (k) => {
      if (k === "a") throw new Error("quota");
    });
    const used: string[] = [];
    for (let i = 0; i < 2; i++) await p.run(async (k) => void used.push(k));
    expect(used).toEqual(["b", "b"]);
    c.advance(61_000);
    const later: string[] = [];
    for (let i = 0; i < 2; i++) await p.run(async (k) => void later.push(k));
    expect(later).toContain("a");
  });

  it("never retries a key marked invalid", async () => {
    const { pool: p, c } = pool(["a", "b"]);
    await p.run(async (k) => {
      if (k === "a") throw new Error("invalid");
    });
    c.advance(10 * 60 * 60 * 1000);
    const used: string[] = [];
    for (let i = 0; i < 3; i++) await p.run(async (k) => void used.push(k));
    expect(used).toEqual(["b", "b", "b"]);
  });

  it("does not fail over on other errors", async () => {
    const { pool: p } = pool(["a", "b"]);
    const used: string[] = [];
    await expect(
      p.run(async (k) => {
        used.push(k);
        throw new Error("boom");
      }),
    ).rejects.toThrow("boom");
    expect(used).toEqual(["a"]);
  });

  it("throws KeysExhaustedError when every key fails", async () => {
    const { pool: p } = pool(["a", "b"]);
    await expect(
      p.run(async () => {
        throw new Error("quota");
      }),
    ).rejects.toBeInstanceOf(KeysExhaustedError);
  });

  it("reports exhaustion without calling op when all keys are cooling down", async () => {
    const { pool: p } = pool(["a", "b"]);
    await p.run(async () => {}).catch(() => {});
    await expect(p.run(async () => { throw new Error("quota"); })).rejects.toBeInstanceOf(KeysExhaustedError);
    let called = false;
    await expect(
      p.run(async () => {
        called = true;
      }),
    ).rejects.toBeInstanceOf(KeysExhaustedError);
    expect(called).toBe(false);
  });

  it("rethrows the original error when there is only one key", async () => {
    const { pool: p } = pool(["a"]);
    await expect(
      p.run(async () => {
        throw new Error("quota");
      }),
    ).rejects.toThrow("quota");
  });

  it("stops failing over once canRetry says output already started", async () => {
    const { pool: p } = pool(["a", "b"]);
    const used: string[] = [];
    await expect(
      p.run(
        async (k) => {
          used.push(k);
          throw new Error("quota");
        },
        () => false,
      ),
    ).rejects.toThrow("quota");
    expect(used).toEqual(["a"]);
  });

  it("dedupes keys", async () => {
    const { pool: p } = pool(["a", "a", "b"]);
    expect(p.size).toBe(2);
  });
});
