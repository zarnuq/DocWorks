/*
 * The app's own spell checker (Chromium's, the one drawing the red underlines), for menus that replace the native
 * right-click menu and so must offer its fixes themselves. Desktop only.
 */
import { Platform } from "obsidian";

export interface Speller {
  misspelled(word: string): boolean;
  suggestions(word: string): string[];
  addToDictionary(word: string): void;
}

interface Electron {
  webFrame?: { isWordMisspelled?(word: string): boolean; getWordSuggestions?(word: string): string[] };
  remote?: { getCurrentWebContents(): { session: { addWordToSpellCheckerDictionary(word: string): boolean } } };
}

export function electronSpeller(): Speller | null {
  if (!Platform.isDesktopApp) return null;
  let electron: Electron;
  try {
    electron = window.require("electron") as Electron;
  } catch {
    return null;
  }
  const frame = electron.webFrame;
  if (!frame?.isWordMisspelled || !frame.getWordSuggestions) return null;
  return {
    misspelled: (w) => frame.isWordMisspelled!(w),
    suggestions: (w) => frame.getWordSuggestions!(w),
    addToDictionary: (w) => void electron.remote?.getCurrentWebContents().session.addWordToSpellCheckerDictionary(w),
  };
}

/** Cut / copy / paste as the app's Edit menu does it (Electron runs the real clipboard command). */
export function runEditCommand(command: "cut" | "copy" | "paste" | "pasteAndMatchStyle"): boolean {
  if (!Platform.isDesktopApp) return command !== "paste" && command !== "pasteAndMatchStyle" && activeDocument.execCommand(command);
  const contents = (window.require("electron") as { remote: { getCurrentWebContents(): Record<typeof command, () => void> } }).remote.getCurrentWebContents();
  contents[command]();
  return true;
}
