/**
 * The find and replace bar the three editors share (Ctrl+F / Ctrl+H): find and replace fields, option checkboxes,
 * "n of m", a preview line for matches in comments and notes, and ↑ ↓ Replace / Replace all / ×. Each editor does
 * its own searching through the callbacks. Google's options (case, whole words, regular expressions, accents) come
 * first; the last search of each file is kept while Obsidian runs.
 */
import type { SearchOptions } from "../../core/utils/text-matches";

/** Google's options, before each editor's own: [label, accessible name, key]. */
const COMMON: [string, string, keyof SearchOptions][] = [
  ["Aa", "Match case", "matchCase"],
  ["Whole words", "Match whole words", "wholeWord"],
  [".*", "Match using regular expressions", "regex"],
  ["Accents", "Ignore accents", "ignoreAccents"],
];

/** Last search per file: query, replacement and every checkbox. */
const remembered = new Map<string, { query: string; replace: string; checks: boolean[] }>();
export interface FindBarOptions {
  /** Remembers the search under this key (the file's path). */
  key?: string;
  /** The editor's own checkboxes, after Google's: [label, accessible name, initially checked]. */
  options: [text: string, title: string, checked: boolean][];
  /** Accessible names of Replace and Replace all. */
  replaceTitles: [one: string, all: string];
  /** The search text (`text`) or an option changed. */
  changed(text: boolean): void;
  /** Enter / ↓ (1), Shift+Enter / ↑ (-1). */
  step(dir: number): void;
  replace(): void;
  replaceAll(): void;
  close(): void;
}

export interface FindBar {
  el: HTMLElement;
  input: HTMLInputElement;
  replaceInput: HTMLInputElement;
  /** The editor's own checkboxes. */
  options: HTMLInputElement[];
  /** Google's options as checked. */
  search(): SearchOptions;
  count: HTMLElement;
  preview: HTMLElement;
  replaceButton: HTMLButtonElement;
  replaceAllButton: HTMLButtonElement;
  /** "2 of 5", "No results", or nothing while the search is empty. */
  setCount(query: string, index: number, total: number): void;
}

/** Start a file's next find bar on `query` (a vault search result being opened). */
export function rememberSearch(key: string, query: string, matchCase: boolean): void {
  const last = remembered.get(key);
  const checks = last?.checks.slice() ?? [];
  checks[0] = matchCase;
  remembered.set(key, { query, replace: last?.replace ?? "", checks: Array.from(checks, (c) => !!c) });
}

/** Forget every remembered search (tests). */
export function forgetSearches(): void {
  remembered.clear();
}

/** A find bar inserted into `parent` before `before` (the document area). */
export function findBar(parent: HTMLElement, before: Element | null, o: FindBarOptions): FindBar {
  const el = createDiv({ cls: "pptx-find" });
  parent.insertBefore(el, before);
  const input = el.createEl("input", { type: "search", attr: { placeholder: "Find", "aria-label": "Find" } });
  const replaceInput = el.createEl("input", { type: "text", attr: { placeholder: "Replace with", "aria-label": "Replace with" } });
  const boxes = [...COMMON.map(([text, title]): [string, string, boolean] => [text, title, false]), ...o.options].map(([text, title, checked]) => {
    const label = el.createEl("label", { cls: "pptx-find-case" });
    const box = label.createEl("input", { type: "checkbox", attr: { "aria-label": title } });
    box.checked = checked;
    label.appendText(text);
    box.addEventListener("change", () => {
      remember();
      o.changed(false);
    });
    return box;
  });
  const options = boxes.slice(COMMON.length);
  const search = () => Object.fromEntries(COMMON.map(([, , key], i) => [key, boxes[i].checked])) as SearchOptions;
  const remember = () => {
    if (o.key) remembered.set(o.key, { query: input.value, replace: replaceInput.value, checks: boxes.map((b) => b.checked) });
  };
  const last = o.key ? remembered.get(o.key) : undefined;
  if (last) {
    input.value = last.query;
    replaceInput.value = last.replace;
    last.checks.forEach((c, i) => boxes[i] && (boxes[i].checked = c));
    input.select();
  }
  replaceInput.addEventListener("input", remember);
  const count = el.createSpan({ cls: "pptx-find-count" });
  const preview = createDiv({ cls: "dw-find-preview", attr: { "aria-live": "polite" } });
  const button = (text: string, title: string, fn: () => void) => {
    const b = el.createEl("button", { text, cls: "dw-btn", attr: { "aria-label": title } });
    b.addEventListener("click", fn);
    return b;
  };
  button("↑", "Previous match (Shift+Enter)", () => o.step(-1));
  button("↓", "Next match (Enter)", () => o.step(1));
  const replaceButton = button("Replace", o.replaceTitles[0], () => o.replace());
  const replaceAllButton = button("Replace all", o.replaceTitles[1], () => o.replaceAll());
  button("×", "Close (Esc)", () => o.close());
  el.append(preview);
  input.addEventListener("input", () => {
    remember();
    o.changed(true);
  });
  el.addEventListener("keydown", (e) => {
    if (e.key === "Escape") o.close();
    else if (e.key === "Enter" && e.target === input) {
      e.preventDefault();
      o.step(e.shiftKey ? -1 : 1);
    }
  });
  const setCount = (query: string, index: number, total: number) => count.setText(query ? (total ? `${index + 1} of ${total}` : "No results") : "");
  return { el, input, replaceInput, options, search, count, preview, replaceButton, replaceAllButton, setCount };
}
