import { addExportToMenu, addNoteExportToMenu, exportNoteToDocx } from "./presentation/common/export-note";
import { exportDocxFile } from "./presentation/common/export-markdown";
import { registerEmbeds } from "./presentation/common/embeds";
import { canExportPdf } from "./presentation/common/pdf";
import { OfficeSearchView, VIEW_TYPE_SEARCH } from "./presentation/common/office-search";
import { registerCellValues } from "./presentation/common/cell-values";
import { IndexedDbVersionStore, setVersionHistory, VersionHistory, versions } from "./presentation/common/version-history";
import { App, FileView, Notice, Plugin, PluginSettingTab, Setting, SettingDefinitionItem, TFile, TFolder, WorkspaceLeaf } from "obsidian";
import { blankDocx, blankPptx, blankXlsx } from "./infrastructure/blank";
import { DEFAULT_SETTINGS, DocWorksSettings, DocxView, VIEW_TYPE_DOCX } from "./presentation/docx/docx-view";
import { PptxView, VIEW_TYPE_PPTX } from "./presentation/pptx/pptx-view";
import { XlsxView, VIEW_TYPE_XLSX } from "./presentation/xlsx/xlsx-view";

const NEW_FILES = [
  ["new-presentation", "New presentation", "presentation", "pptx", blankPptx],
  ["new-document", "New document", "file-text", "docx", blankDocx],
  ["new-spreadsheet", "New spreadsheet", "table", "xlsx", blankXlsx],
] as const;

export default class DocWorksPlugin extends Plugin {
  settings: DocWorksSettings = { ...DEFAULT_SETTINGS };

  private async exportWordNote(file: TFile): Promise<void> {
    const active = this.app.workspace.getActiveViewOfType(DocxView);
    const open = active?.file?.path === file.path ? active : this.app.workspace.getLeavesOfType(VIEW_TYPE_DOCX)
      .map(leaf => leaf.view).find((view): view is DocxView => view instanceof DocxView && view.file?.path === file.path);
    if (open) await open.exportAs("md");
    else await exportDocxFile(this.app, file);
  }

  /**
   * Background tabs restored after the plugin (re)loads keep the placeholder title and icon Obsidian gave them while
   * it was off (the view type, a ghost); name them after their file until they're opened.
   */
  private nameBackgroundTabs(): void {
    const icons: Record<string, string> = { [VIEW_TYPE_DOCX]: "file-text", [VIEW_TYPE_PPTX]: "presentation", [VIEW_TYPE_XLSX]: "table" };
    this.app.workspace.iterateAllLeaves((leaf) => {
      const { type, state } = leaf.getViewState();
      const file = state?.file;
      if (!icons[type] || !leaf.isDeferred || typeof file !== "string") return;
      // ponytail: a deferred view's title / icon and updateHeader() are private API; if they change this does nothing.
      const view = leaf.view as unknown as { title?: string; icon?: string };
      view.title = file.split("/").pop()!.replace(/\.[^.]+$/, "");
      view.icon = icons[type];
      (leaf as unknown as { updateHeader?: () => void }).updateHeader?.();
    });
  }

  async onload() {
    this.settings = { ...DEFAULT_SETTINGS, ...((await this.loadData()) as Partial<DocWorksSettings> | null) };
    this.addSettingTab(new DocWorksSettingTab(this.app, this));
    // Earlier versions of files, kept in the app's storage (not the vault) before Doc Works saves over them.
    setVersionHistory(new VersionHistory(new IndexedDbVersionStore(`doc-works-history/${this.app.vault.getName()}`), () => ({ enabled: this.settings.history, days: this.settings.historyDays })));
    const office = (f: unknown): f is TFile => f instanceof TFile && ["docx", "pptx", "xlsx"].includes(f.extension.toLowerCase());
    this.registerEvent(this.app.vault.on("rename", (f, old) => office(f) && void versions?.renamed(old, f.path)));
    this.registerEvent(this.app.vault.on("delete", (f) => office(f) && void versions?.deleted(f.path)));
    this.registerView(VIEW_TYPE_DOCX, (leaf: WorkspaceLeaf) => new DocxView(leaf, this.settings, () => this.saveSettings()));
    this.registerView(VIEW_TYPE_PPTX, (leaf: WorkspaceLeaf) => new PptxView(leaf, this.settings));
    this.registerView(VIEW_TYPE_XLSX, (leaf: WorkspaceLeaf) => new XlsxView(leaf, this.settings));
    // File types switched off in the settings are left to Obsidian or another plugin.
    if (this.settings.openDocx) this.claimExtension("docx", VIEW_TYPE_DOCX);
    if (this.settings.openPptx) this.claimExtension("pptx", VIEW_TYPE_PPTX);
    if (this.settings.openXlsx) this.claimExtension("xlsx", VIEW_TYPE_XLSX);
    // Previews in notes (![[deck.pptx#3]]) for the same file types.
    const kinds = (["docx", "pptx", "xlsx"] as const).filter((k) => this.settings[k === "docx" ? "openDocx" : k === "pptx" ? "openPptx" : "openXlsx"]);
    registerEmbeds(this, kinds);
    // `cell: Budget.xlsx#Sheet!B4` in notes shows the cell's value.
    if (this.settings.openXlsx) registerCellValues(this);
    // Search inside the same file types, in a side panel.
    this.registerView(VIEW_TYPE_SEARCH, (leaf: WorkspaceLeaf) => new OfficeSearchView(leaf, () => kinds));
    this.addCommand({ id: "search-office-files", name: "Search in documents, presentations and spreadsheets", callback: () => void this.openSearch() });
    this.app.workspace.onLayoutReady(() => this.nameBackgroundTabs());

    for (const [id, name, fromStart] of [
      ["start-presentation", "Start presentation from the beginning", true],
      ["start-presentation-current", "Start presentation from the current slide", false],
    ] as const) {
      this.addCommand({
        id,
        name,
        checkCallback: (checking) => {
          const view = this.app.workspace.getActiveViewOfType(PptxView);
          if (view && !checking) view.startPresentation(fromStart);
          return !!view;
        },
      });
    }

    this.addCommand({
      id: "export-note-docx",
      name: "Export current note to .docx",
      checkCallback: (checking) => {
        const file = this.app.workspace.getActiveFile();
        if (file?.extension !== "md") return false;
        if (!checking) void exportNoteToDocx(this.app, file);
        return true;
      },
    });
    this.registerEvent(
      this.app.workspace.on("file-menu", (menu, file, source) => {
        if (file instanceof TFile && file.extension === "md")
          addNoteExportToMenu(menu, this.app, file);
        if (file instanceof TFile && file.extension === "docx" && source !== "more-options")
          addExportToMenu(menu, "Export to Markdown", () => void this.exportWordNote(file));
      })
    );

    this.addCommand({
      id: "export-docx-markdown",
      name: "Export current document to Markdown",
      checkCallback: (checking) => {
        const file = this.app.workspace.getActiveFile();
        if (file?.extension !== "docx") return false;
        if (!checking) void this.exportWordNote(file);
        return true;
      },
    });

    this.addCommand({
      id: "export-pdf",
      name: "Export to PDF",
      checkCallback: (checking) => {
        const view = this.app.workspace.getActiveViewOfType(DocxView) ?? this.app.workspace.getActiveViewOfType(PptxView) ?? this.app.workspace.getActiveViewOfType(XlsxView);
        if (!view || !canExportPdf()) return false;
        if (!checking) void view.exportAs("pdf");
        return true;
      },
    });

    for (const [id, name, , ext, make] of NEW_FILES) {
      this.addCommand({
        id,
        name,
        callback: () => void this.createFile(ext, make, this.app.fileManager.getNewFileParent(this.app.workspace.getActiveFile()?.path ?? "")),
      });
    }
    this.registerEvent(
      this.app.workspace.on("file-menu", (menu, file) => {
        if (!(file instanceof TFolder)) return;
        for (const [, name, icon, ext, make] of NEW_FILES) {
          menu.addItem((item) => item.setTitle(name).setIcon(icon).onClick(() => void this.createFile(ext, make, file)));
        }
      })
    );

    // Flush any pending debounced save when Obsidian is closing/reloading.
    this.registerEvent(
      (this.app.workspace as { on(name: string, cb: () => Promise<void>): import("obsidian").EventRef }).on(
        "quit",
        async () => {
          for (const type of [VIEW_TYPE_DOCX, VIEW_TYPE_PPTX, VIEW_TYPE_XLSX]) {
            for (const leaf of this.app.workspace.getLeavesOfType(type)) {
              await (leaf.view as DocxView | PptxView | XlsxView).requestSave?.();
            }
          }
        }
      )
    );
  }

  onunload(): void {
    setVersionHistory(null);
  }

  async saveSettings(): Promise<void> {
    await this.saveData(this.settings);
    // Tabs not shown since startup hold a placeholder view until they're opened.
    for (const leaf of this.app.workspace.getLeavesOfType(VIEW_TYPE_DOCX)) if (leaf.view instanceof DocxView) leaf.view.applyPageSize();
  }

  /** The search panel in the right sidebar (the open one, or a new one). */
  private async openSearch(): Promise<void> {
    const leaf = this.app.workspace.getLeavesOfType(VIEW_TYPE_SEARCH)[0] ?? this.app.workspace.getRightLeaf(false);
    if (!leaf) return;
    await leaf.setViewState({ type: VIEW_TYPE_SEARCH, active: true });
    await this.app.workspace.revealLeaf(leaf);
  }

  /** Open files again with the current settings (read-only on or off); pending edits are saved first. */
  async reopenViews(): Promise<void> {
    for (const type of [VIEW_TYPE_DOCX, VIEW_TYPE_PPTX, VIEW_TYPE_XLSX]) {
      for (const leaf of this.app.workspace.getLeavesOfType(type)) {
        const view = leaf.view;
        if (!(view instanceof FileView) || !view.file) continue;
        await view.onUnloadFile(view.file);
        await view.onLoadFile(view.file);
      }
    }
  }

  /** Create "Untitled.<ext>" (or "Untitled 1"...) in `folder` from a blank template and open it. */
  private async createFile(ext: string, make: () => Promise<Uint8Array>, folder: TFolder): Promise<void> {
    const dir = folder.isRoot() ? "" : `${folder.path}/`;
    let path = `${dir}Untitled.${ext}`;
    for (let i = 1; this.app.vault.getAbstractFileByPath(path); i++) path = `${dir}Untitled ${i}.${ext}`;
    const bytes = await make();
    const file = await this.app.vault.createBinary(path, bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer);
    await this.app.workspace.getLeaf("tab").openFile(file);
  }

  /** Another plugin may already own the extension; keep loading instead of failing. */
  private claimExtension(ext: string, viewType: string): void {
    try {
      this.registerExtensions([ext], viewType);
    } catch (err) {
      console.warn(`[doc-works] .${ext} is already handled by another plugin`, err);
      new Notice(`Doc Works: .${ext} is already handled by another plugin; disable it to use this editor.`);
    }
  }
}

const READ_ONLY = { name: "Read only", desc: "Open documents, presentations and spreadsheets for viewing only: no editing tools, and files are never changed." };

const HISTORY = { name: "Version history", desc: "Before saving over a file, keep the version it replaces (File › Version history). Kept in the app, not in the vault." };
const HISTORY_DAYS = { name: "Keep versions for", desc: "Older versions are deleted. At most 10 per file and 200 MB in all.", options: { "7": "1 week", "14": "2 weeks", "30": "30 days" } };

const AUTHOR = { name: "Comment author", desc: "Your name on comments you add to Word documents." };

const PAPER = {
  name: "Paper colour",
  desc: "How files open: light or dark pages, slides' surroundings and sheets. Each file's toolbar can still switch.",
  options: { default: "Light documents, dark slides and sheets", light: "Light", dark: "Dark", obsidian: "Match Obsidian's theme" },
};

const OPEN = [
  ["openDocx", "Open Word documents (.docx)"],
  ["openPptx", "Open PowerPoint presentations (.pptx)"],
  ["openXlsx", "Open Excel spreadsheets (.xlsx)"],
] as const;
const OPEN_DESC = "Turn off to leave this file type to Obsidian or another plugin. Applies after reloading Obsidian.";

const PAGE_SIZE = {
  name: "Word page size",
  desc: "Page view uses each document's own page size, or always US Letter (8.5 × 11 in).",
  options: { document: "From the document", letter: "Always Letter" },
};

class DocWorksSettingTab extends PluginSettingTab {
  constructor(
    app: App,
    private plugin: DocWorksPlugin
  ) {
    super(app, plugin);
  }

  // Obsidian 1.13+: rendered declaratively and searchable.
  getSettingDefinitions(): SettingDefinitionItem[] {
    return [
      { name: READ_ONLY.name, desc: READ_ONLY.desc, control: { type: "toggle", key: "readOnly", defaultValue: false } },
      { name: HISTORY.name, desc: HISTORY.desc, control: { type: "toggle", key: "history", defaultValue: true } },
      { name: HISTORY_DAYS.name, desc: HISTORY_DAYS.desc, control: { type: "dropdown", key: "historyDays", options: HISTORY_DAYS.options, defaultValue: "14" } },
      { name: PAGE_SIZE.name, desc: PAGE_SIZE.desc, control: { type: "dropdown", key: "pageSize", options: PAGE_SIZE.options, defaultValue: "document" } },
      { name: AUTHOR.name, desc: AUTHOR.desc, control: { type: "text", key: "author", defaultValue: "" } },
      { name: PAPER.name, desc: PAPER.desc, control: { type: "dropdown", key: "paper", options: PAPER.options, defaultValue: "default" } },
      ...OPEN.map(([key, name]) => ({ name, desc: OPEN_DESC, control: { type: "toggle" as const, key, defaultValue: true } })),
    ];
  }

  // The days dropdown's options are text.
  getControlValue(key: string): unknown {
    const value = (this.plugin.settings as unknown as Record<string, unknown>)[key];
    return key === "historyDays" ? String(value) : value;
  }

  async setControlValue(key: string, value: unknown): Promise<void> {
    if (key === "pageSize") this.plugin.settings.pageSize = value === "letter" ? "letter" : "document";
    if (key === "author") this.plugin.settings.author = typeof value === "string" ? value.trim() : "";
    if (key === "paper" && typeof value === "string" && value in PAPER.options) this.plugin.settings.paper = value as keyof typeof PAPER.options;
    for (const [open] of OPEN) if (key === open) this.plugin.settings[open] = value !== false;
    const reopen = key === "readOnly" && this.plugin.settings.readOnly !== (value === true);
    if (key === "readOnly") this.plugin.settings.readOnly = value === true;
    if (key === "history") this.plugin.settings.history = value !== false;
    if (key === "historyDays") this.plugin.settings.historyDays = Number(value) || 14;
    await this.plugin.saveSettings();
    if (reopen) await this.plugin.reopenViews();
  }

  // Older Obsidian versions.
  display(): void {
    this.containerEl.empty();
    new Setting(this.containerEl)
      .setName(READ_ONLY.name)
      .setDesc(READ_ONLY.desc)
      .addToggle((t) => t.setValue(this.plugin.settings.readOnly).onChange((v) => void this.setControlValue("readOnly", v)));
    new Setting(this.containerEl)
      .setName(HISTORY.name)
      .setDesc(HISTORY.desc)
      .addToggle((t) => t.setValue(this.plugin.settings.history).onChange((v) => void this.setControlValue("history", v)));
    new Setting(this.containerEl)
      .setName(HISTORY_DAYS.name)
      .setDesc(HISTORY_DAYS.desc)
      .addDropdown((d) => d.addOptions(HISTORY_DAYS.options).setValue(String(this.plugin.settings.historyDays)).onChange((v) => void this.setControlValue("historyDays", v)));
    new Setting(this.containerEl)
      .setName(PAGE_SIZE.name)
      .setDesc(PAGE_SIZE.desc)
      .addDropdown((d) =>
        d
          .addOptions(PAGE_SIZE.options)
          .setValue(this.plugin.settings.pageSize)
          .onChange((v) => void this.setControlValue("pageSize", v))
      );
    new Setting(this.containerEl)
      .setName(AUTHOR.name)
      .setDesc(AUTHOR.desc)
      .addText((t) => t.setValue(this.plugin.settings.author).onChange((v) => void this.setControlValue("author", v)));
    new Setting(this.containerEl)
      .setName(PAPER.name)
      .setDesc(PAPER.desc)
      .addDropdown((d) => d.addOptions(PAPER.options).setValue(this.plugin.settings.paper).onChange((v) => void this.setControlValue("paper", v)));
    for (const [key, name] of OPEN) {
      new Setting(this.containerEl)
        .setName(name)
        .setDesc(OPEN_DESC)
        .addToggle((t) => t.setValue(this.plugin.settings[key]).onChange((v) => void this.setControlValue(key, v)));
    }
  }
}
