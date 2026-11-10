/** Plugin settings, shared by the views (no Obsidian imports). */
import type { PageSize } from "../../domain/docx/pages";

export type PaperSetting = "default" | "light" | "dark" | "obsidian";
export type FileKind = "docx" | "pptx" | "xlsx";

export interface DocWorksSettings {
  /** "document": pages use the document's own size; "letter": always 8.5 x 11 in. */
  pageSize: PageSize;
  /** Name shown on new comments; asked for on the first comment when empty. */
  author: string;
  /** Paper / canvas colours each file type opens with; "default": light documents, dark slides and sheets. */
  paperDocx: PaperSetting;
  paperPptx: PaperSetting;
  paperXlsx: PaperSetting;
  /** Word documents open with desktop Word's ribbon, status bar and menus instead of the Google Docs layout (beta). */
  officeDocx: boolean;
  /** The same for presentations (PowerPoint's ribbon) and spreadsheets (Excel's). */
  officePptx: boolean;
  officeXlsx: boolean;
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

export const DEFAULT_SETTINGS: DocWorksSettings = {
  pageSize: "document", author: "", paperDocx: "default", paperPptx: "default", paperXlsx: "default", officeDocx: false, officePptx: false, officeXlsx: false,
  openDocx: true, openPptx: true, openXlsx: true, readOnly: false, history: true, historyDays: 14,
};

const PAPERS: readonly unknown[] = ["default", "light", "dark", "obsidian"];
const paperOr = (value: unknown, fallback: PaperSetting): PaperSetting => (PAPERS.includes(value) ? (value as PaperSetting) : fallback);

/** Settings as saved (by this or an earlier version) with the defaults filled in; the one `paper` setting of earlier versions applies to every file type. */
export function migrateSettings(raw: Record<string, unknown> | null): DocWorksSettings {
  const { paper, ...rest } = raw ?? {};
  const old = paperOr(paper, "default");
  const s = { ...DEFAULT_SETTINGS, ...(rest as Partial<DocWorksSettings>) };
  s.paperDocx = paperOr(rest.paperDocx, old);
  s.paperPptx = paperOr(rest.paperPptx, old);
  s.paperXlsx = paperOr(rest.paperXlsx, old);
  s.officeDocx = rest.officeDocx === true;
  s.officePptx = rest.officePptx === true;
  s.officeXlsx = rest.officeXlsx === true;
  return s;
}

/** The theme a view of `kind` opens with: its setting, or the view's own default. */
export function paperFor(settings: DocWorksSettings, kind: FileKind, viewDefault: "light" | "dark"): "light" | "dark" {
  const paper = kind === "docx" ? settings.paperDocx : kind === "pptx" ? settings.paperPptx : settings.paperXlsx;
  if (paper === "light" || paper === "dark") return paper;
  if (paper === "obsidian") return document.body.classList.contains("theme-dark") ? "dark" : "light";
  return viewDefault;
}
