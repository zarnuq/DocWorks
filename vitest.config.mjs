import { fileURLToPath } from "url";
import { defineConfig } from "vitest/config";

export default defineConfig({
  resolve: { alias: { obsidian: fileURLToPath(new URL("tests/obsidian-stub.ts", import.meta.url)) } },
  // Saving zips is slow when every test file runs at once: a generous timeout avoids false failures.
  test: { environment: "jsdom", include: ["tests/**/*.test.ts"], setupFiles: ["tests/setup.ts"], testTimeout: 30000 },
});
