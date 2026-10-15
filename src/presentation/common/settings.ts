/** Plugin settings, shared by the views (no Obsidian imports). */
import type { PageSize } from "../../domain/docx/pages";

export type PaperSetting = "default" | "light" | "dark" | "obsidian";

export interface DocWorksSettings {
  /** "document": pages use the document's own size; "letter": always 8.5 x 11 in. */
  pageSize: PageSize;
  /** Name shown on new comments; asked for on the first comment when empty. */
  author: string;
  /** Paper / canvas colours a file opens with; "default": light documents, dark slides and sheets. */
  paper: PaperSetting;
  /** Which file types Doc Works opens (applies after a reload). */
  openDocx: boolean;
  openPptx: boolean;
  openXlsx: boolean;
  /** Open every file as a viewer: no editing controls, nothing is ever saved. */
  readOnly: boolean;
  /** Keep earlier versions of files before saving over them (File › Version history), for `historyDays` days. */
  history: boolean;
  historyDays: number;
}

export const DEFAULT_SETTINGS: DocWorksSettings = { pageSize: "document", author: "", paper: "default", openDocx: true, openPptx: true, openXlsx: true, readOnly: false, history: true, historyDays: 14 };

/** The theme a view opens with: the setting, or the view's own default. */
export function paperFor(settings: DocWorksSettings, viewDefault: "light" | "dark"): "light" | "dark" {
  if (settings.paper === "light" || settings.paper === "dark") return settings.paper;
  if (settings.paper === "obsidian") return document.body.classList.contains("theme-dark") ? "dark" : "light";
  return viewDefault;
}
