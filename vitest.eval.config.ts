import { defineConfig } from "vitest/config";

// Live evals call the real model API, so they live in their own config and are never part of `npm test`.
export default defineConfig({
  test: {
    environment: "node",
    include: ["evals/**/*.eval.ts"],
    testTimeout: 15 * 60 * 1000,
    fileParallelism: false,
  },
});
