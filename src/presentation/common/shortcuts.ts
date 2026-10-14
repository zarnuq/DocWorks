import { App, Modal, Platform, Scope } from "obsidian";

/** Sections of [action, keys]; keys are written with Ctrl / Alt and shown as ⌘ / ⌥ on macOS. */
export type ShortcutSection = [title: string, rows: [action: string, keys: string][]];

/** What every editor shares. */
export const COMMON_SHORTCUTS: ShortcutSection[] = [
  ["Common actions", [
    ["Undo", "Ctrl+Z"],
    ["Redo", "Ctrl+Y, Ctrl+Shift+Z"],
    ["Cut", "Ctrl+X"],
    ["Copy", "Ctrl+C"],
    ["Paste", "Ctrl+V"],
    ["Select all", "Ctrl+A"],
    ["Find and replace", "Ctrl+F"],
    ["Keyboard shortcuts", "Ctrl+/"],
  ]],
  ["Moving around", [
    ["Next area (menus, toolbar, document, panels)", "F6"],
    ["Previous area", "Shift+F6"],
    ["Back to the document from a panel", "Esc"],
  ]],
];

export const DOCS_SHORTCUTS: ShortcutSection[] = [
  ["Text formatting", [
    ["Bold", "Ctrl+B"],
    ["Italic", "Ctrl+I"],
    ["Underline", "Ctrl+U"],
    ["Strikethrough", "Ctrl+Shift+X"],
    ["Code", "Ctrl+`"],
  ]],
  ["Paragraphs", [
    ["Bulleted list", "Ctrl+Shift+8"],
    ["Numbered list", "Ctrl+Shift+7"],
    ["Increase indent / list level", "Tab"],
    ["Decrease indent / list level", "Shift+Tab"],
  ]],
  ["Insert", [
    ["Link", "Ctrl+K"],
    ["Follow a link", "Ctrl+click"],
    ["Comment", "Ctrl+Alt+M"],
    ["Page break", "Ctrl+Enter"],
    ["Find and replace", "Ctrl+H"],
  ]],
  ["Tables", [
    ["Next cell", "Tab"],
    ["Previous cell", "Shift+Tab"],
  ]],
];

export const SLIDES_SHORTCUTS: ShortcutSection[] = [
  ["Objects", [
    ["Duplicate", "Ctrl+D"],
    ["Group", "Ctrl+G"],
    ["Ungroup", "Ctrl+Shift+G"],
    ["Nudge", "Arrow keys"],
    ["Nudge by a pixel", "Alt+Arrow keys"],
    ["Delete", "Delete"],
    ["Select none", "Esc"],
  ]],
  ["Text", [
    ["Bold", "Ctrl+B"],
    ["Italic", "Ctrl+I"],
    ["Underline", "Ctrl+U"],
    ["Link", "Ctrl+K"],
    ["Comment", "Ctrl+Alt+M"],
    ["Bulleted list", "Ctrl+Shift+8"],
    ["Numbered list", "Ctrl+Shift+7"],
    ["Increase / decrease list level", "Tab, Shift+Tab"],
  ]],
  ["Slides", [
    ["Next / previous slide", "Page Down, Page Up"],
    ["Grid view", "Ctrl+Alt+1"],
    ["Present from the beginning", "F5"],
    ["Present from this slide", "Shift+F5"],
  ]],
  ["Presenting", [
    ["Next", "→, Space, Enter, N"],
    ["Previous", "←, Backspace"],
    ["First / last slide", "Home, End"],
    ["Go to slide", "Number, then Enter"],
    ["Presenter view", "P"],
    ["Stop presenting", "Esc"],
  ]],
];

export const SHEETS_SHORTCUTS: ShortcutSection[] = [
  ["Moving and selecting", [
    ["Move", "Arrow keys"],
    ["To the edge of the data", "Ctrl+Arrow keys"],
    ["Extend the selection", "Shift+Arrow keys"],
    ["Start / end of the row", "Home, End"],
    ["Start / end of the sheet", "Ctrl+Home, Ctrl+End"],
    ["Page down / up", "Page Down, Page Up"],
    ["Cell menu", "Shift+F10"],
  ]],
  ["Editing", [
    ["Edit cell", "F2, Enter"],
    ["Cancel editing", "Esc"],
    ["New line in a cell", "Alt+Enter"],
    ["Clear", "Delete"],
    ["Fill down", "Ctrl+D"],
    ["Fill right", "Ctrl+R"],
    ["Insert date", "Ctrl+;"],
    ["Insert time", "Ctrl+Shift+;"],
    ["Check / uncheck checkboxes", "Space"],
  ]],
  ["Formatting", [
    ["Bold", "Ctrl+B"],
    ["Italic", "Ctrl+I"],
    ["Underline", "Ctrl+U"],
    ["Strikethrough", "Ctrl+Shift+X"],
  ]],
  ["Insert", [
    ["Link", "Ctrl+K"],
    ["Open a link", "Ctrl+click"],
    ["Comment", "Ctrl+Alt+M"],
    ["Note", "Shift+F2"],
  ]],
];

const shown = (keys: string) => (Platform.isMacOS ? keys.replace(/Ctrl\+/g, "⌘").replace(/Alt\+/g, "⌥").replace(/Shift\+/g, "⇧") : keys);

/** Help › Keyboard shortcuts: every section, filtered as you type. */
export function shortcutsDialog(app: App, sections: ShortcutSection[]): void {
  const modal = new Modal(app);
  modal.titleEl.setText("Keyboard shortcuts");
  const search = modal.contentEl.createEl("input", { type: "search", attr: { placeholder: "Search shortcuts", "aria-label": "Search shortcuts" } });
  const list = modal.contentEl.createDiv({ cls: "dw-shortcuts" });
  const draw = () => {
    list.empty();
    const q = search.value.trim().toLowerCase();
    for (const [title, rows] of sections) {
      const hits = rows.filter(([action, keys]) => !q || action.toLowerCase().includes(q) || keys.toLowerCase().includes(q));
      if (!hits.length) continue;
      list.createEl("h4", { text: title });
      for (const [action, keys] of hits) {
        const row = list.createDiv({ cls: "dw-shortcut" });
        row.createSpan({ text: action });
        row.createEl("kbd", { text: shown(keys) });
      }
    }
    if (!list.childElementCount) list.createDiv({ cls: "dw-chars-empty", text: "No matches" });
  };
  search.addEventListener("input", draw);
  draw();
  modal.open();
  search.focus();
}

/** Ctrl+/ opens the list, as in Google's editors. */
export function registerShortcutsKey(scope: Scope, open: () => void): void {
  scope.register(["Mod"], "/", () => {
    open();
    return false;
  });
}
