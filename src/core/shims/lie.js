// Clean, native Promise shim for browser & obsidian environments
// Replaces legacy lie promise polyfill
"use strict";

module.exports = typeof Promise !== "undefined" ? Promise : globalThis.Promise;
