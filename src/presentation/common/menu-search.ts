/*
 * Google's "Search the menus" box (Alt+/): every item of the menu row, found by its name or its menu path and run
 * from the keyboard. The menus are built into a recorder each time the box gets focus, so items show the current
 * state (ticks, disabled) and read-only files list only what their menus keep.
 */
import { setIcon, type Menu } from "obsidian";

export interface MenuEntry {
  title: string;
  /** The menus above the item: ["View", "Zoom"]. */
  path: string[];
  run: (evt: MouseEvent | KeyboardEvent) => void;
  checked?: boolean;
  disabled?: boolean;
  icon?: string;
}

/** Results shown at once. */
const LIMIT = 15;

interface Recorded {
  title: string;
  run?: (evt: MouseEvent | KeyboardEvent) => void;
  checked?: boolean;
  disabled?: boolean;
  icon?: string;
  sub?: Recorded[];
}

/** Stands in for an Obsidian Menu or MenuItem: `own` methods record, every other method is a no-op returning itself. */
function stand<T extends object>(own: Record<string, (...args: never[]) => unknown>): T {
  const proxy: T = new Proxy(own, {
    get: (target, key: string) => target[key] ?? (() => proxy),
  }) as unknown as T;
  return proxy;
}

function recorder(items: Recorded[]): Menu {
  const menu: Menu = stand<Menu>({
    addItem: (cb: (item: unknown) => unknown) => {
      const rec: Recorded = { title: "" };
      const item: unknown = stand({
        setTitle: (t: string | DocumentFragment) => ((rec.title = typeof t === "string" ? t : t.textContent ?? ""), item),
        onClick: (fn: Recorded["run"]) => ((rec.run = fn), item),
        setChecked: (c: boolean | null) => ((rec.checked = c ?? undefined), item),
        setDisabled: (d: boolean) => ((rec.disabled = d), item),
        setIcon: (icon: string | null) => ((rec.icon = icon ?? undefined), item),
        setSubmenu: () => recorder((rec.sub ??= [])),
      });
      items.push(rec);
      cb(item);
      return menu;
    },
  });
  return menu;
}

function flatten(items: Recorded[], path: string[], out: MenuEntry[]): MenuEntry[] {
  for (const rec of items) {
    if (rec.sub) flatten(rec.sub, [...path, rec.title], out);
    else if (rec.title && rec.run) {
      const e: MenuEntry = { title: rec.title, path, run: rec.run };
      if (rec.checked !== undefined) e.checked = rec.checked;
      if (rec.disabled) e.disabled = true;
      if (rec.icon) e.icon = rec.icon;
      out.push(e);
    }
  }
  return out;
}

/** The items one menu builds, its submenus' items under their path. A builder that throws gives what it had. */
export function recordMenu(name: string, build: (m: Menu) => void): MenuEntry[] {
  const items: Recorded[] = [];
  try {
    build(recorder(items));
  } catch {
    // The menu would have failed to open too; keep what it added.
  }
  return flatten(items, [name], []);
}

const norm = (s: string) => s.toLowerCase().replace(/…|\.\.\.|&/g, " ").replace(/\s+/g, " ").trim();

/** A query word is found at a word start, or anywhere once it's three letters long (one letter finds too much). */
const finds = (text: string, word: string) => (` ${text}`.includes(` ${word}`) || (word.length >= 3 && text.includes(word)));

/**
 * The entries whose title or path holds every word of `query`: the exact title first, then titles starting with
 * it, with a word starting with it, holding it, and last the ones found through their path. Menu order otherwise.
 */
export function matchMenus(entries: MenuEntry[], query: string): MenuEntry[] {
  const q = norm(query);
  if (!q) return [];
  const words = q.split(" ");
  const seen = new Set<string>();
  const scored: [number, MenuEntry][] = [];
  for (const e of entries) {
    const key = [...e.path, e.title].join("\u0000");
    if (seen.has(key)) continue;
    const t = norm(e.title), p = norm(e.path.join(" "));
    if (!words.every((w) => finds(t, w) || finds(p, w))) continue;
    seen.add(key);
    const score = t === q ? 0 : t.startsWith(q) ? 1 : t.includes(` ${q}`) ? 2 : t.includes(q) ? 3 : words.every((w) => finds(t, w)) ? 4 : 5;
    scored.push([score, e]);
  }
  // Array sort is stable: menu order among equals.
  return scored.sort((a, b) => a[0] - b[0]).map(([, e]) => e);
}

let ids = 0;

/** The box and its result list. `collect` gives every menu item as things stand now. */
export class MenuSearch {
  readonly el: HTMLElement;
  readonly input: HTMLInputElement;
  private list: HTMLElement | null = null;
  private entries: MenuEntry[] = [];
  private hits: MenuEntry[] = [];
  private selected = 0;
  /** Where focus was before the box: it goes back there after a pick or Escape. */
  private returnTo: HTMLElement | null = null;
  private readonly id = `dw-menu-search-${++ids}`;

  constructor(private collect: () => MenuEntry[]) {
    this.el = createDiv({ cls: "dw-menu-search" });
    setIcon(this.el.createSpan({ cls: "dw-menu-search-icon" }), "search");
    this.input = this.el.createEl("input", {
      type: "search",
      placeholder: "Menus",
      attr: { "aria-label": "Search the menus", "aria-keyshortcuts": "Alt+/", role: "combobox", "aria-expanded": "false", "aria-autocomplete": "list", "aria-controls": this.id, spellcheck: "false" },
    });
    const input = this.input;
    // Clicked: remember the editor before focus moves.
    input.addEventListener("mousedown", () => this.remember(document.activeElement));
    input.addEventListener("focus", (e) => {
      if (!this.returnTo) this.remember(e.relatedTarget);
      this.entries = this.collect();
    });
    input.addEventListener("blur", () => {
      this.close();
      this.returnTo = null;
    });
    input.addEventListener("input", () => this.show());
    input.addEventListener("keydown", (e) => this.key(e));
  }

  /** Alt+/: focus the box, with what was typed selected. */
  focus(): void {
    this.remember(document.activeElement);
    this.input.focus();
    this.input.select();
  }

  private remember(el: EventTarget | null): void {
    this.returnTo = el instanceof HTMLElement && !this.el.contains(el) && el !== document.body ? el : null;
  }

  private key(e: KeyboardEvent): void {
    // Plain keys are the box's own; app hotkeys (Ctrl / Alt) still go through.
    if (!e.ctrlKey && !e.metaKey && !e.altKey) e.stopPropagation();
    if (e.isComposing) return;
    if (e.key === "ArrowDown" || e.key === "ArrowUp") {
      e.preventDefault();
      if (!this.hits.length) return;
      this.select((this.selected + (e.key === "ArrowDown" ? 1 : -1) + this.hits.length) % this.hits.length);
    } else if (e.key === "Enter") {
      e.preventDefault();
      const hit = this.hits[this.selected];
      if (hit) this.choose(hit, e);
    } else if (e.key === "Escape") {
      e.preventDefault();
      this.input.value = "";
      this.leave();
    }
  }

  private choose(hit: MenuEntry, evt: MouseEvent | KeyboardEvent): void {
    if (hit.disabled) return;
    this.input.value = "";
    this.leave();
    hit.run(evt);
  }

  /** Close and give focus back. */
  private leave(): void {
    const to = this.returnTo;
    this.close();
    if (to?.isConnected) to.focus();
    else this.input.blur();
  }

  private close(): void {
    this.list?.remove();
    this.list = null;
    this.hits = [];
    this.input.setAttribute("aria-expanded", "false");
    this.input.removeAttribute("aria-activedescendant");
  }

  private show(): void {
    const q = this.input.value;
    if (!q.trim()) return this.close();
    this.hits = matchMenus(this.entries, q).slice(0, LIMIT);
    this.selected = 0;
    if (!this.list) {
      const doc = this.input.ownerDocument;
      this.list = doc.body.createDiv({ cls: "suggestion-container dw-menu-search-results" });
      // Clicking a result mustn't blur the box (which closes the list) before the click lands.
      this.list.addEventListener("mousedown", (e) => e.preventDefault());
      const r = this.input.getBoundingClientRect();
      this.list.setCssProps({ "--dw-menu-search-left": `${r.left}px`, "--dw-menu-search-top": `${r.bottom + 4}px` });
    }
    this.list.empty();
    const box = this.list.createDiv({ cls: "suggestion", attr: { id: this.id, role: "listbox", "aria-label": "Menu items" } });
    this.input.setAttribute("aria-expanded", "true");
    if (!this.hits.length) {
      box.createDiv({ cls: "suggestion-empty", text: "No results" });
      this.input.removeAttribute("aria-activedescendant");
      return;
    }
    this.hits.forEach((hit, n) => {
      const item = box.createDiv({ cls: "suggestion-item dw-menu-search-item", attr: { id: `${this.id}-${n}`, role: "option" } });
      item.toggleClass("is-disabled", !!hit.disabled);
      item.toggleClass("is-checked", !!hit.checked);
      if (hit.disabled) item.setAttribute("aria-disabled", "true");
      const icon = item.createSpan({ cls: "dw-menu-search-item-icon" });
      if (hit.checked) setIcon(icon, "check");
      else if (hit.icon) setIcon(icon, hit.icon);
      item.createSpan({ cls: "dw-menu-search-title", text: hit.title });
      item.createSpan({ cls: "dw-menu-search-path", text: hit.path.join(" › ") });
      item.addEventListener("mousemove", () => this.selected !== n && this.select(n));
      item.addEventListener("click", (e) => this.choose(hit, e));
    });
    this.select(0);
  }

  private select(n: number): void {
    this.selected = n;
    const items = Array.from(this.list?.querySelectorAll<HTMLElement>(".suggestion-item") ?? []);
    items.forEach((el, i) => {
      el.toggleClass("is-selected", i === n);
      el.setAttribute("aria-selected", String(i === n));
    });
    const current = items[n];
    if (!current) return;
    this.input.setAttribute("aria-activedescendant", current.id);
    current.scrollIntoView?.({ block: "nearest" });
  }
}
