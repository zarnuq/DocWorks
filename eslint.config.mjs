import { defineConfig } from "eslint/config";
import obsidianmd from "eslint-plugin-obsidianmd";

export default defineConfig([
  { ignores: ["main.js", "node_modules/", "tests/", "*.mjs", "src/core/shims/", "src/presentation/pptx/font-metrics.ts", "src/domain/pptx/preset-shapes.ts"] },
  ...obsidianmd.configs.recommended,
  {
    languageOptions: {
      parserOptions: { projectService: { allowDefaultProject: ["eslint.config.*"] } },
    },
  },
]);
