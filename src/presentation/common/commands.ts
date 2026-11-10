/** A view's commands by id, shared by the Office-style ribbon, Quick Access Toolbar, right-click menu and status bar. */
import type { Menu } from "obsidian";

export interface Cmd {
  /** Caption, as the Office app words it. */
  label: string;
  /** Lucide icon name. */
  icon: string;
  run?: () => void;
  /** A dropdown's items (or a split button's ▾ part, when there's also `run`). */
  menu?: (m: Menu) => void;
  /** On, for toggles (Bold, Ruler…). */
  active?: () => boolean;
  enabled?: () => boolean;
  /** Only looks at the file: kept for read-only files. */
  view?: boolean;
  /** Shown in the tooltip, e.g. "Ctrl+B". */
  shortcut?: string;
  /** A double-click on its button (Format Painter: keep painting). */
  dblclick?: () => void;
}

export type Commands = Record<string, Cmd>;

/** Obsidian's per-vault local storage (App.loadLocalStorage / saveLocalStorage). */
export interface AppStorage {
  loadLocalStorage(key: string): unknown;
  saveLocalStorage(key: string, value: unknown): void;
}

/** "Bold (Ctrl+B)": the tooltip of a command's button. */
export const tooltip = (c: Cmd): string => (c.shortcut ? `${c.label} (${c.shortcut})` : c.label);

/** "Mod-Shift-l" → "Ctrl+Shift+L", as the keyboard shortcuts list writes keys. */
export function keyName(key: string): string {
  return key.split("-").map((p) => (p === "Mod" ? "Ctrl" : p.length === 1 ? p.toUpperCase() : p)).join("+");
}

/** One shortcuts-list section of an app's keys ([ProseMirror-style key, command id]), captions from its commands. */
export function shortcutRows(keys: [key: string, id: string][], label: (id: string) => string): [action: string, keys: string][] {
  const rows = new Map<string, string[]>();
  for (const [key, id] of keys) rows.set(id, [...(rows.get(id) ?? []), keyName(key)]);
  return [...rows].map(([id, k]) => [label(id), k.join(", ")]);
}
