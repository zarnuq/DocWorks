/**
 * Search in documents, presentations and spreadsheets: a side panel that searches the text of every Office file in
 * the vault (Obsidian's own search skips them). Files are read once and kept until they change. A result opens the
 * file with its find bar already on the search.
 */
import { ItemView, Keymap, TFile, WorkspaceLeaf } from "obsidian";
import { officeText, searchText, TextPart } from "../../infrastructure/office-text";
import { rememberSearch } from "./find-bar";

export const VIEW_TYPE_SEARCH = "doc-works-search";
/** Results shown per file before "n more". */
const PER_FILE = 5;

export class OfficeSearchView extends ItemView {
  private cache = new Map<string, { mtime: number; parts: TextPart[] }>();
  private runs = 0;
  private input!: HTMLInputElement;
  private matchCase!: HTMLInputElement;
  private status!: HTMLElement;
  private results!: HTMLElement;
  private timer: number | null = null;

  /** `kinds`: the file types Doc Works opens. */
  constructor(
    leaf: WorkspaceLeaf,
    private kinds: () => string[]
  ) {
    super(leaf);
  }

  getViewType(): string {
    return VIEW_TYPE_SEARCH;
  }

  getDisplayText(): string {
    return "Search documents";
  }

  getIcon(): string {
    return "file-search";
  }

  async onOpen(): Promise<void> {
    const root = this.contentEl;
    root.empty();
    root.addClass("dw-search");
    this.input = root.createEl("input", { type: "search", attr: { placeholder: "Search documents, slides and sheets", "aria-label": "Search documents, slides and sheets" } });
    const label = root.createEl("label", { cls: "dw-search-option" });
    this.matchCase = label.createEl("input", { type: "checkbox", attr: { "aria-label": "Match case" } });
    label.appendText("Match case");
    this.status = root.createDiv({ cls: "dw-search-status", attr: { "aria-live": "polite" } });
    this.results = root.createDiv({ cls: "dw-search-results" });
    const later = () => {
      if (this.timer) window.clearTimeout(this.timer);
      this.timer = window.setTimeout(() => void this.runSearch(this.input.value), 250);
    };
    this.input.addEventListener("input", later);
    this.matchCase.addEventListener("change", later);
    this.input.focus();
  }

  async onClose(): Promise<void> {
    if (this.timer) window.clearTimeout(this.timer);
    this.runs++;
  }

  /** Search every Office file; a newer search stops this one. */
  async runSearch(query: string): Promise<void> {
    const run = ++this.runs;
    this.input.value = query;
    this.results.empty();
    this.status.setText("");
    if (!query.trim()) return;
    const kinds = this.kinds();
    const files = this.app.vault.getFiles().filter((f) => kinds.includes(f.extension.toLowerCase())).sort((a, b) => a.path.localeCompare(b.path));
    let total = 0, found = 0;
    for (const [i, file] of files.entries()) {
      if (files.length > 5) this.status.setText(`Searching ${i + 1} of ${files.length}…`);
      const parts = await this.textOf(file);
      if (run !== this.runs) return;
      const hits = searchText(parts, query, { matchCase: this.matchCase.checked });
      if (!hits.length) continue;
      total += hits.length;
      found++;
      const group = this.results.createDiv({ cls: "dw-search-file" });
      const head = group.createDiv({ cls: "dw-search-head" });
      head.createSpan({ cls: "dw-search-name", text: file.name });
      head.createSpan({ cls: "dw-search-count", text: String(hits.length) });
      if (file.parent && !file.parent.isRoot()) group.createDiv({ cls: "dw-search-folder", text: file.parent.path });
      head.addEventListener("click", (e) => void this.openResult(file, query, e));
      for (const hit of hits.slice(0, PER_FILE)) {
        const row = group.createDiv({ cls: "dw-search-hit", attr: { tabindex: "0" } });
        if (hit.label) row.createDiv({ cls: "dw-search-where", text: hit.label });
        const line = row.createDiv({ cls: "dw-search-line" });
        line.appendText(hit.before);
        line.createEl("mark", { text: hit.match });
        line.appendText(hit.after);
        row.addEventListener("click", (e) => void this.openResult(file, query, e));
        row.addEventListener("keydown", (e) => e.key === "Enter" && void this.openResult(file, query, e));
      }
      if (hits.length > PER_FILE) group.createDiv({ cls: "dw-search-more", text: `${hits.length - PER_FILE} more` });
    }
    this.status.setText(total ? `${total} ${total === 1 ? "result" : "results"} in ${found} ${found === 1 ? "file" : "files"}` : "No results");
  }

  /** The file's text, read again only after it changed. A file that can't be read has none. */
  private async textOf(file: TFile): Promise<TextPart[]> {
    const cached = this.cache.get(file.path);
    if (cached && cached.mtime === file.stat.mtime) return cached.parts;
    let parts: TextPart[] = [];
    try {
      parts = await officeText(file.extension.toLowerCase() as "docx", await this.app.vault.readBinary(file));
    } catch (err) {
      console.warn(`[doc-works] couldn't read ${file.path} for search`, err);
    }
    this.cache.set(file.path, { mtime: file.stat.mtime, parts });
    return parts;
  }

  /** Open the file (a new tab with Ctrl / Cmd) with its find bar on the search. */
  private async openResult(file: TFile, query: string, e: MouseEvent | KeyboardEvent): Promise<void> {
    rememberSearch(file.path, query, this.matchCase.checked);
    const leaf = this.app.workspace.getLeaf(Keymap.isModEvent(e));
    await leaf.openFile(file);
    (leaf.view as { openFind?: () => void }).openFind?.();
  }
}
