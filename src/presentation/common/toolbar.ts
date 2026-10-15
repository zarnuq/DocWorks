import { SaveStatus } from "./save-status";
import { rovingToolbar } from "./keyboard";
import { App, FileView, Menu, MenuItem, Modal, setIcon, Setting, Modifier, Scope } from "obsidian";

export const FONTS = ["Aptos", "Calibri", "Cambria", "Arial", "Times New Roman", "Georgia", "Verdana", "Segoe UI", "Trebuchet MS", "Courier New"];

export type PaperTheme = "light" | "dark";

/** Open a toolbar menu under its button. */
export function showBelow(menu: Menu, anchor: HTMLElement): void {
  const r = anchor.getBoundingClientRect();
  lastAt = { x: r.left, y: r.bottom }; // FALLBACK(submenu)
  menu.showAtPosition(lastAt);
}

/** Menu bar items' icons, by title, as Google's menus show them. */
const MENU_ICONS: Record<string, string> = {
  "Export…": "file-output", Undo: "undo-2", Redo: "redo-2", Cut: "scissors", Copy: "copy", Paste: "clipboard-paste",
  "Paste without formatting": "clipboard-type", "Paste values only": "clipboard-type", "Select all": "text-select", Delete: "trash-2",
  "Find and replace": "replace", Editing: "pencil", Suggesting: "file-pen", "Page view": "file-text", Ruler: "ruler", "Linked from": "links-coming-in",
  Outline: "list-tree", "Review panel": "messages-square", Zoom: "zoom-in", Light: "sun", Dark: "moon", Present: "play",
  "Present from start": "presentation", "Grid view": "layout-grid", "Format panel": "panel-right", Freeze: "snowflake",
  Image: "image", Table: "table", Drawing: "pen-tool", Chart: "chart-column", Equation: "sigma", "Special characters": "omega",
  "Horizontal line": "minus", Break: "separator-horizontal", Footnote: "superscript", Comment: "message-square-plus", Link: "link",
  "Table of contents": "book-open-text", "Header & footer": "panel-top", Watermark: "droplet", "Text box": "square-dashed-text",
  "Audio or video": "film", Shape: "shapes", Field: "braces", "New slide": "square-plus", Sheet: "sheet", "Pivot table": "table-2",
  Function: "sigma", Note: "sticky-note", Dropdown: "circle-chevron-down", Text: "type", "Paragraph style": "pilcrow",
  Align: "align-left", "Line spacing": "unfold-vertical", "Paragraph…": "pilcrow", "Columns…": "columns-2", "Bulleted list": "list",
  "Numbered list": "list-ordered", "Clear formatting": "remove-formatting", "Character spacing": "move-horizontal", Number: "hash",
  Alignment: "align-left", "Wrap text": "wrap-text", "Merge cells": "table-cells-merge", "Conditional formatting": "palette",
  "Duplicate slide": "copy", "Copy slide": "copy", "Cut slide": "scissors", "Delete slide": "trash-2", "Hide slide": "eye-off",
  "Move up": "arrow-up", "Move down": "arrow-down", "Add section here": "between-horizontal-start", "Theme…": "palette", "Motion…": "sparkles",
  Order: "layers", Distribute: "align-horizontal-space-around", Group: "group", Ungroup: "ungroup", "Paint format": "paintbrush",
  "Word count": "whole-word", "Sort A → Z": "arrow-down-az", "Sort Z → A": "arrow-down-za", "Create a filter": "filter",
  "Remove filter": "filter-x", "Data validation": "list-checks", "Remove duplicates…": "list-x", "Split text to columns…": "columns-2",
  "Named ranges…": "tag", "Import CSV as new sheet…": "file-input", Bold: "bold", Italic: "italic", Underline: "underline",
  "Keyboard shortcuts": "keyboard", "Copy as Markdown table": "table", "Copy as Markdown outline": "list-tree", "Copy link to slide": "link", "Copy link to cells": "link", "Copy link to heading": "link", Comments: "messages-square", Checkbox: "square-check", Bar: "chart-bar", Column: "chart-column", Line: "chart-line", Pie: "chart-pie", Strikethrough: "strikethrough", Subscript: "subscript", Superscript: "superscript", Code: "code", "Remove highlight": "highlighter",
};
/** Set while a menu bar menu is being built: its items get MENU_ICONS. */
let menuBarIcons = false;

/** Read-only files: the menus kept, and the items (or whole submenus) in them that only look at the file. */
const VIEW_MENUS = new Set(["File", "Edit", "View", "Tools", "Help"]);
const VIEW_ITEMS = new Set(["Export…", "Copy", "Copy as Markdown table", "Copy as Markdown outline", "Copy link to slide", "Copy link to cells", "Copy link to heading", "Select all", "Find and replace", "Page view", "Outline", "Zoom", "Light", "Dark", "Present", "Present from start", "Grid view", "Word count", "Keyboard shortcuts", "Linked from"]);
/** Set while a read-only file's menu is being built: other items are left out. */
let viewOnly = false;
const hidden = (title: string) => viewOnly && !VIEW_ITEMS.has(title);
const menuIcon = (title: string) => (menuBarIcons ? MENU_ICONS[title] ?? (/^Ungroup /.test(title) ? "ungroup" : /^Group /.test(title) ? "group" : undefined) : undefined);

export function menuItem(m: Menu, title: string, run: () => void, opts: { checked?: boolean; disabled?: boolean; icon?: string } = {}): void {
  if (hidden(title)) return;
  m.addItem((i) => {
    i.setTitle(title).onClick(run);
    if (opts.checked !== undefined) i.setChecked(opts.checked);
    if (opts.disabled) i.setDisabled(true);
    const icon = opts.icon ?? menuIcon(title);
    if (icon) i.setIcon(icon);
  });
}

/** Obsidian's MenuItem.setSubmenu (in the app since 1.4, not in its typings): a hover submenu with a › arrow. */
type SubmenuItem = MenuItem & { setSubmenu?: () => Menu };

/** An item that opens a submenu. */
export function subMenu(m: Menu, title: string, build: (m: Menu) => void): void {
  if (hidden(title)) return;
  // A submenu kept for a read-only file is kept whole (Zoom's levels).
  const outer = build;
  build = (sub) => {
    const was = viewOnly;
    viewOnly = false;
    try {
      outer(sub);
    } finally {
      viewOnly = was;
    }
  };
  m.addItem((i: SubmenuItem) => {
    i.setTitle(title);
    const icon = menuIcon(title);
    if (icon) i.setIcon(icon);
    if (typeof i.setSubmenu === "function") return build(i.setSubmenu());
    // FALLBACK(submenu): only for an Obsidian without setSubmenu. To drop it, delete from here to the end of this
    // function, `lastAt` below, its assignment in showBelow and the fallback test in tests/toolbar.test.ts.
    i.setTitle(`${title} ▸`).onClick((e) => {
      const sub = new Menu();
      build(sub);
      sub.showAtPosition("clientX" in e && e.clientX ? { x: e.clientX, y: e.clientY } : lastAt);
    });
  });
}

/** Separators only between items: leftovers of removed items (read-only menus) are dropped. */
function tidySeparators(menu: Menu): void {
  const addItem = menu.addItem.bind(menu), addSeparator = menu.addSeparator.bind(menu);
  let pending = false, any = false;
  menu.addSeparator = () => ((pending = any), menu);
  menu.addItem = (cb) => {
    if (pending) addSeparator();
    pending = false;
    any = true;
    return addItem(cb);
  };
}

// FALLBACK(submenu): where the last toolbar menu opened, for fallback submenus opened from the keyboard.
let lastAt = { x: 0, y: 0 };

/**
 * The editor toolbars, as Google's: a menu row holding every command above one row of everyday controls, boxed in
 * groups split by gap(). Groups that don't fit fold into a ⋮ menu; contextual groups show only for a selection.
 */
export class Toolbar {
  private group: HTMLElement;
  menubar: HTMLElement | null = null;
  private more: HTMLButtonElement | null = null;
  private tests = new Map<HTMLElement, () => boolean>();
  /** Whether the row is wider than the bar; replaced in tests (jsdom has no layout). */
  overflows = (): boolean => this.el.scrollWidth > this.el.clientWidth;

  /** `readOnly`: only the menus and groups (`keep`) that look at the file, never change it. */
  constructor(
    readonly el: HTMLElement,
    readonly readOnly = false
  ) {
    this.group = el.createDiv({ cls: "dw-group" });
  }

  button(label: string, title: string, onClick: () => void, cls = ""): HTMLButtonElement {
    const b = this.group.createEl("button", { text: label, cls: `dw-btn ${cls}`.trim(), attr: { "aria-label": title } });
    b.addEventListener("click", (e) => {
      e.preventDefault();
      onClick();
    });
    return b;
  }

  /** A button showing a Lucide icon instead of a label. */
  icon(icon: string, title: string, onClick: () => void): HTMLButtonElement {
    const b = this.button("", title, onClick, "dw-icon");
    setIcon(b, icon);
    return b;
  }

  /**
   * A dropdown showing its value (Font, Size, Zoom…): a button opening the app's menu with the value ticked. The
   * options live in a hidden select, which callers get back and set (`value`, `showValue`); the button follows it.
   */
  select(title: string, placeholder: string, options: [label: string, value: string][], onChange: (value: string) => void): HTMLSelectElement {
    const btn = this.group.createEl("button", { cls: "dw-btn dw-select", attr: { "aria-label": title, "aria-haspopup": "menu" } });
    const label = btn.createSpan({ cls: "dw-select-label" });
    btn.createSpan({ cls: "dw-caret", text: "▾" });
    const select = optionsMenu(btn, fillSelect(btn.createEl("select", { attr: { "aria-label": title } }), placeholder, options, onChange), true);
    const sync = () => label.setText(select.selectedOptions[0]?.text ?? placeholder);
    const value = Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, "value")!;
    Object.defineProperty(select, "value", {
      get: () => value.get!.call(select) as string,
      set: (v: string) => {
        value.set!.call(select, v);
        sync();
      },
    });
    select.addEventListener("change", sync);
    sync();
    return select;
  }

  /** An icon button opening a menu of actions, the same kind of menu as the menu bar's (see `select`). */
  iconSelect(icon: string, title: string, options: [label: string, value: string][], onChange: (value: string) => void): HTMLSelectElement {
    const btn = this.group.createEl("button", { cls: "dw-btn dw-icon-select", attr: { "aria-label": title, "aria-haspopup": "menu" } });
    const glyph = btn.createSpan();
    setIcon(glyph, icon);
    if (!glyph.querySelector("svg")) glyph.setText(title);
    btn.createSpan({ cls: "dw-caret", text: "▾" });
    // A list of actions, never showing a value: picking the same one again must work.
    const select = optionsMenu(btn, fillSelect(btn.createEl("select", { attr: { "aria-label": title } }), title, options, onChange), false);
    select.addEventListener("change", () => (select.value = ""));
    return select;
  }

  /** A button-like label over a hidden color input; the last color picked shows as an underline. */
  color(label: string, title: string, initial: string, onChange: (hex: string) => void): void {
    const wrap = this.group.createEl("label", { text: label, cls: "dw-btn dw-color", attr: { "aria-label": title } });
    wrap.setCssProps({ "--dw-swatch": initial });
    const input = wrap.createEl("input", { type: "color", value: initial, attr: { "aria-label": title } });
    input.addEventListener("change", () => {
      wrap.setCssProps({ "--dw-swatch": input.value });
      onChange(input.value.toUpperCase());
    });
  }

  /**
   * File: the same menu as the view's ⋮ (More options) button, opened under File. Obsidian builds it in
   * ItemView.onMoreOptions, which the app has but its typings don't; without it, `build` makes the menu.
   */
  fileMenu(view: object, build: (m: Menu) => void): HTMLButtonElement {
    const more = (view as { onMoreOptions?: (e: MouseEvent) => void }).onMoreOptions;
    return this.menu("File", build, more ? (e) => (more.call(view, e), true) : undefined);
  }

  /** A menu in the menu row above the toolbar; `build` runs each time it opens, so items show the current state. */
  menu(name: string, build: (m: Menu) => void, open?: (e: MouseEvent) => boolean): HTMLButtonElement {
    if (this.readOnly && !VIEW_MENUS.has(name)) return createEl("button");
    const b = this.menuRow().createEl("button", { text: name, cls: "dw-menu-btn", attr: { "aria-haspopup": "menu" } });
    b.addEventListener("mousedown", (e) => e.preventDefault());
    b.addEventListener("click", (e) => {
      if (open?.(e)) return;
      const menu = new Menu();
      if (this.readOnly) tidySeparators(menu);
      menuBarIcons = true;
      viewOnly = this.readOnly;
      try {
        build(menu);
      } finally {
        menuBarIcons = false;
        viewOnly = false;
      }
      showBelow(menu, b);
    });
    b.addEventListener("keydown", (e) => {
      if (e.key !== "ArrowDown") return;
      e.preventDefault();
      b.click();
    });
    return b;
  }

  private menuRow(): HTMLElement {
    if (!this.menubar) {
      this.menubar = createDiv({ cls: "dw-menubar" });
      this.el.parentElement?.insertBefore(this.menubar, this.el);
    }
    return this.menubar;
  }

  private sideEl: HTMLElement | null = null;

  /** An icon button at the right end of the menu row, as Google's version history and comments buttons. */
  side(icon: string, title: string, onClick: () => void): HTMLButtonElement {
    this.sideEl ??= this.menuRow().createDiv({ cls: "dw-menubar-end" });
    const b = this.sideEl.createEl("button", { cls: "clickable-icon dw-menubar-icon", attr: { "aria-label": title } });
    setIcon(b, icon);
    b.addEventListener("mousedown", (e) => e.preventDefault());
    b.addEventListener("click", onClick);
    return b;
  }

  /** Start the next group. */
  gap(): void {
    this.group = this.el.createDiv({ cls: "dw-group" });
  }

  /** The current group stays on a read-only file's toolbar. */
  keep(): void {
    this.group.dataset.view = "true";
  }

  /** The current group folds into ⋮ only after every unpinned one. */
  pin(): void {
    this.group.dataset.pinned = "true";
  }

  /** The current group shows only while `test()` holds (checked again by `sync`). */
  context(test: () => boolean): void {
    this.group.addClass("dw-context");
    this.tests.set(this.group, test);
  }

  /** Check the contextual groups again after the selection changed; refit only when one appeared or went (fit reads layout). */
  sync(): void {
    let changed = false;
    for (const [group, test] of this.tests) {
      const off = !test();
      if (off === group.classList.contains("is-off")) continue;
      group.toggleClass("is-off", off);
      changed = true;
    }
    if (changed) this.fit();
  }

  private groups(): HTMLElement[] {
    return Array.from(this.el.querySelectorAll<HTMLElement>(":scope > .dw-group")).filter((g) => g.childElementCount && !g.classList.contains("is-off"));
  }

  /** Fold unpinned groups from the right into ⋮ until the row fits on one line; on a phone, pinned ones too (never the first). */
  fit(): void {
    for (const g of Array.from(this.el.querySelectorAll<HTMLElement>(":scope > .dw-group"))) {
      g.hidden = g.classList.contains("is-off");
      delete g.dataset.folded;
    }
    if (this.more) this.more.hidden = true;
    if (!this.overflows()) return;
    if (this.more) this.more.hidden = false;
    // Right to left, the selection's own tools last: they're what the user is working with.
    const foldable = this.groups().filter((g) => !g.dataset.pinned).reverse();
    const pinned = this.groups().slice(1).filter((g) => g.dataset.pinned).reverse();
    const order = [...foldable.filter((g) => !this.tests.has(g)), ...foldable.filter((g) => this.tests.has(g)), ...pinned];
    for (const g of order) {
      if (!this.overflows()) break;
      g.hidden = true;
      g.dataset.folded = "true";
    }
  }

  /** The folded groups' controls as menu items; each one works the real (hidden) control. */
  overflowItems(m: Menu): void {
    this.groups()
      .filter((g) => g.dataset.folded)
      .forEach((g, n) => {
        if (n) m.addSeparator();
        for (const el of Array.from(g.children)) {
          const select = el.instanceOf(HTMLSelectElement) ? el : el.querySelector("select");
          const color = el.querySelector<HTMLInputElement>('input[type="color"]');
          // "Bold (Ctrl+B)" -> "Bold": the hint is for the tooltip.
          const label = (el.getAttribute("aria-label") ?? el.textContent ?? "").replace(/\s*\([^)]*\)$/, "");
          if (color) {
            menuItem(m, `${label}…`, () => pickColor(color.value, (hex) => {
              color.value = hex;
              color.dispatchEvent(new Event("change"));
            }));
          } else if (select) {
            subMenu(m, label, (sub) => {
              for (const o of Array.from(select.options).filter((o) => o.value !== "")) {
                menuItem(sub, o.text, () => {
                  select.value = o.value;
                  select.dispatchEvent(new Event("change"));
                }, { checked: select.value === o.value });
              }
            });
          } else if (el.instanceOf(HTMLButtonElement)) {
            menuItem(m, label, () => el.click(), { checked: el.classList.contains("is-active") || undefined, disabled: el.disabled });
          }
        }
      });
  }

  /** Save status (and whatever the view adds to the returned element) pushed to the end of the bar. */
  end(): { status: SaveStatus; el: HTMLElement } {
    if (this.readOnly) {
      for (const g of Array.from(this.el.querySelectorAll<HTMLElement>(":scope > .dw-group")).filter((g) => !g.dataset.view)) {
        this.tests.delete(g);
        g.remove();
      }
    }
    const more = (this.more = this.el.createEl("button", { cls: "dw-btn dw-icon dw-more", attr: { "aria-label": "More tools" } }));
    setIcon(more, "more-vertical");
    more.hidden = true;
    more.addEventListener("click", () => {
      const menu = new Menu();
      this.overflowItems(menu);
      showBelow(menu, more);
    });
    if (typeof ResizeObserver !== "undefined") new ResizeObserver(() => this.fit()).observe(this.el);
    const end = this.el.createDiv({ cls: "dw-end" });
    if (this.readOnly) end.createSpan({ cls: "dw-status", text: "View only" });
    // The end area grows on its own ("Couldn't save" + Retry, the word count): the bar's size doesn't change, so refit.
    new MutationObserver(() => this.fit()).observe(end, { childList: true, subtree: true, characterData: true });
    // The bar is built: make it (and the menu row) one Tab stop with arrow keys inside.
    rovingToolbar(this.el);
    if (this.menubar) {
      rovingToolbar(this.menubar);
      this.menubar.setAttribute("role", "menubar");
    }
    const status = new SaveStatus(end);
    // Nothing is ever saved: no "Saved".
    if (this.readOnly) end.querySelector<HTMLElement>('[role="status"]')?.hide();
    return { status, el: end };
  }
}

/** The system colour picker from a menu item (Chromium won't open a hidden input's): a temporary input passes on the pick. */
export function pickColor(initial: string, onPick: (hex: string) => void): void {
  const temp = document.body.createEl("input", { type: "color", cls: "dw-color-temp", value: initial });
  temp.addEventListener("change", () => {
    onPick(temp.value.toUpperCase());
    temp.remove();
  });
  temp.addEventListener("blur", () => window.setTimeout(() => temp.remove(), 1000));
  try {
    temp.showPicker();
  } catch {
    temp.click();
  }
}

/** Show a select's current value: tick the matching option, or name a value that isn't listed in its placeholder. */
export function showValue(s: HTMLSelectElement, value: string | null): void {
  const listed = value !== null && Array.from(s.options).some((o) => o.value === value && o.value !== "");
  s.options[0].text = listed || value === null ? (s.dataset.placeholder ?? "") : value;
  s.value = listed ? value : "";
}

/** `btn` opens the app's menu of `select`'s options (hidden); picking one sets it and fires its change. */
function optionsMenu(btn: HTMLElement, select: HTMLSelectElement, tick: boolean): HTMLSelectElement {
  select.hidden = true;
  btn.addEventListener("mousedown", (e) => e.preventDefault());
  btn.addEventListener("click", (e) => {
    e.preventDefault();
    const menu = new Menu();
    for (const o of Array.from(select.options).filter((o) => o.value !== "")) {
      menuItem(menu, o.text, () => {
        select.value = o.value;
        select.dispatchEvent(new Event("change"));
      }, tick ? { checked: select.value === o.value } : {});
    }
    showBelow(menu, btn);
  });
  return select;
}

function fillSelect(s: HTMLSelectElement, placeholder: string, options: [label: string, value: string][], onChange: (value: string) => void): HTMLSelectElement {
  s.dataset.placeholder = placeholder;
  // The placeholder labels the closed control only; macOS would otherwise list it with a check mark and anchor the menu on it.
  s.createEl("option", { text: placeholder, value: "" }).hidden = true;
  for (const [text, value] of options) s.createEl("option", { text, value });
  s.addEventListener("change", () => s.value && onChange(s.value));
  return s;
}

/** Ask for one line of text; resolves null when cancelled. */
export function promptText(app: App, title: string, value: string, password = false): Promise<string | null> {
  return new Promise((resolve) => {
    let result: string | null = null;
    const modal = new Modal(app);
    modal.titleEl.setText(title);
    const submit = () => {
      result = value;
      modal.close();
    };
    new Setting(modal.contentEl).addText((text) => {
      text.setValue(value).onChange((v) => (value = v));
      if (password) text.inputEl.type = "password";
      // Cancel the key too: focus returns to the editor before the key's input, which would add a line break there.
      text.inputEl.addEventListener("keydown", (e) => {
        if (e.key !== "Enter" || e.isComposing) return;
        e.preventDefault();
        submit();
      });
      window.setTimeout(() => text.inputEl.select());
    });
    new Setting(modal.contentEl)
      .addButton((b) => b.setButtonText("Save").setCta().onClick(submit))
      .addButton((b) => b.setButtonText("Cancel").onClick(() => modal.close()));
    modal.onClose = () => resolve(result);
    modal.open();
  });
}

/**
 * Keys the editors need that Obsidian's hotkeys (or the user's own, e.g. Shift+Tab for the sidebar) may take:
 * while an editable element has focus they reach it; otherwise they're left alone.
 */
export const EDITING_KEYS: [Modifier[], string][] = [
  [["Mod"], "b"], [["Mod"], "i"], [["Mod"], "u"], [["Mod"], "k"], [["Mod"], "e"], [["Mod"], "l"], [["Mod"], "a"],
  [["Shift"], "Tab"], [[], "Tab"],
];
export function shieldEditingKeys(scope: Scope, keys: [Modifier[], string][] = EDITING_KEYS): void {
  for (const [mods, key] of keys)
    scope.register(mods, key, () => {
      const a = document.activeElement as HTMLElement | null;
      const editable = a?.isContentEditable || !!a?.closest('[contenteditable=""], [contenteditable="true"]') || a?.tagName === "INPUT" || a?.tagName === "TEXTAREA";
      return editable ? true : undefined;
    });
}

/**
 * The file items Obsidian puts in its own editors' ⋮ menu (rename, delete, and what core and other plugins add on
 * "file-menu": move, bookmark, copy path, open in default app…), which a plain FileView doesn't get.
 */
export function addFileItems(view: FileView, menu: Menu, source: string): void {
  const file = view.file;
  if (!file) return;
  const files = view.app.fileManager as { promptForFileRename?: (f: typeof file) => Promise<void> } & typeof view.app.fileManager;
  if (files.promptForFileRename) menu.addItem((i) => i.setSection("action").setTitle("Rename…").setIcon("lucide-edit-3").onClick(() => void files.promptForFileRename!(file)));
  menu.addItem((i) => i.setSection("danger").setTitle("Delete file").setIcon("lucide-trash-2").setWarning(true).onClick(() => void view.app.fileManager.promptForDeletion(file)));
  view.app.workspace.trigger("file-menu", menu, file, source, view.leaf);
}
