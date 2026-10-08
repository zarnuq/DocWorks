import esbuild from "esbuild";
import process from "process";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const prod = process.argv[2] === "production";

const context = await esbuild.context({
  banner: {
    js: "/* Doc Works for Obsidian (GPL-3.0) */",
  },
  entryPoints: ["src/main.ts"],
  bundle: true,
  alias: {
    immediate: path.resolve(__dirname, "src/core/shims/immediate.js"),
    setimmediate: path.resolve(__dirname, "src/core/shims/setimmediate.js"),
    lie: path.resolve(__dirname, "src/core/shims/lie.js"),
    jszip: path.resolve(__dirname, "node_modules/jszip/lib/index.js"),
    "readable-stream": path.resolve(__dirname, "node_modules/jszip/lib/readable-stream-browser.js"),
    stream: path.resolve(__dirname, "src/core/shims/stream.js"),
  },
  external: [
    "obsidian",
    "electron",
    "@codemirror/autocomplete",
    "@codemirror/collab",
    "@codemirror/commands",
    "@codemirror/language",
    "@codemirror/lint",
    "@codemirror/search",
    "@codemirror/state",
    "@codemirror/view",
    "@lezer/common",
    "@lezer/highlight",
    "@lezer/lr",
  ],
  format: "cjs",
  target: "es2020",
  logLevel: "info",
  sourcemap: prod ? false : "inline",
  treeShaking: true,
  outfile: "main.js",
  minify: prod,
});

if (prod) {
  await context.rebuild();
  process.exit(0);
} else {
  await context.watch();
}
