// Clean, safe microtask scheduler for browser & obsidian environments
// Replaces legacy immediate polyfill that dynamically injected <script> elements
"use strict";

module.exports = function immediate(task) {
  if (typeof queueMicrotask === "function") {
    queueMicrotask(task);
  } else if (typeof Promise !== "undefined") {
    Promise.resolve().then(task);
  } else {
    setTimeout(task, 0);
  }
};
