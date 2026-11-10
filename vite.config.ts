import { defineConfig } from "vitest/config";
import react from "@vitejs/plugin-react";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";

// Office add-ins must be served over HTTPS. Certs come from `npm run certs`.
const certDir = path.join(os.homedir(), ".office-addin-dev-certs");
const keyPath = path.join(certDir, "localhost.key");
const certPath = path.join(certDir, "localhost.crt");
const https =
  fs.existsSync(keyPath) && fs.existsSync(certPath)
    ? { key: fs.readFileSync(keyPath), cert: fs.readFileSync(certPath) }
    : undefined;

export default defineConfig({
  plugins: [react()],
  server: { port: 3100, strictPort: true, https },
  test: {
    environment: "node",
    // UI tests opt into jsdom with a "@vitest-environment jsdom" comment at the top of the file.
    include: ["src/**/*.test.{ts,tsx}", "evals/**/*.test.ts"],
    setupFiles: ["src/ui/testSetup.ts"],
  },
});
