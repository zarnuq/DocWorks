#!/usr/bin/env node
import { spawnSync } from "node:child_process";
import { resolve } from "node:path";
const file = process.argv[2];
const result = spawnSync("npx", ["vitest", "run", "tests/performance.test.ts", "--maxWorkers=1"], {
  stdio: "inherit", env: {...process.env, DOC_WORKS_BENCH: "1", ...(file ? {DOC_WORKS_BENCH_FILE: resolve(file)} : {})}
});
process.exit(result.status ?? 1);
