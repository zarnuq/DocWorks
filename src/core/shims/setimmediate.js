// Clean, safe setImmediate / clearImmediate scheduler for browser & obsidian environments
// Replaces legacy setimmediate polyfill that dynamically injected <script> elements and eval/new Function
"use strict";

const g =
  typeof globalThis !== "undefined"
    ? globalThis
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : this;

const si =
  typeof g.setImmediate === "function"
    ? g.setImmediate.bind(g)
    : function (fn, ...args) {
        return setTimeout(() => fn(...args), 0);
      };

const ci =
  typeof g.clearImmediate === "function"
    ? g.clearImmediate.bind(g)
    : function (id) {
        clearTimeout(id);
      };

if (typeof g.setImmediate !== "function") {
  g.setImmediate = si;
  g.clearImmediate = ci;
}

module.exports = si;
module.exports.setImmediate = si;
module.exports.clearImmediate = ci;
