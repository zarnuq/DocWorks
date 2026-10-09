import { App, Notice, Plugin, PluginSettingTab, Setting, SettingDefinitionItem, TFolder, WorkspaceLeaf } from "obsidian";
import { blankDocx, blankPptx } from "./infrastructure/blank";
import { DEFAULT_SETTINGS, DocWorksSettings, DocxView, VIEW_TYPE_DOCX } from "./presentation/docx/docx-view";
import { PptxView, VIEW_TYPE_PPTX } from "./presentation/pptx/pptx-view";

const NEW_FILES = [
  ["new-presentation", "New presentation", "presentation", "pptx", blankPptx],
  ["new-document", "New document", "file-text", "docx", blankDocx],
] as const;

export default class DocWorksPlugin extends Plugin {
  settings: DocWorksSettings = { ...DEFAULT_SETTINGS };

  async onload() {
    this.settings = { ...DEFAULT_SETTINGS, ...((await this.loadData()) as Partial<DocWorksSettings> | null) };
    this.addSettingTab(new DocWorksSettingTab(this.app, this));
    this.registerView(VIEW_TYPE_DOCX, (leaf: WorkspaceLeaf) => new DocxView(leaf, this.settings, () => this.saveSettings()));
    this.registerView(VIEW_TYPE_PPTX, (leaf: WorkspaceLeaf) => new PptxView(leaf));
    this.claimExtension("docx", VIEW_TYPE_DOCX);
    this.claimExtension("pptx", VIEW_TYPE_PPTX);

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
          for (const type of [VIEW_TYPE_DOCX, VIEW_TYPE_PPTX]) {
            for (const leaf of this.app.workspace.getLeavesOfType(type)) {
              await (leaf.view as DocxView | PptxView).requestSave?.();
            }
          }
        }
      )
    );
  }

  async saveSettings(): Promise<void> {
    await this.saveData(this.settings);
    for (const leaf of this.app.workspace.getLeavesOfType(VIEW_TYPE_DOCX)) (leaf.view as DocxView).applyPageSize();
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

const AUTHOR = { name: "Comment author", desc: "Your name on comments you add to Word documents." };

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
      { name: PAGE_SIZE.name, desc: PAGE_SIZE.desc, control: { type: "dropdown", key: "pageSize", options: PAGE_SIZE.options, defaultValue: "document" } },
      { name: AUTHOR.name, desc: AUTHOR.desc, control: { type: "text", key: "author", defaultValue: "" } },
    ];
  }

  async setControlValue(key: string, value: unknown): Promise<void> {
    if (key === "pageSize") this.plugin.settings.pageSize = value === "letter" ? "letter" : "document";
    if (key === "author") this.plugin.settings.author = typeof value === "string" ? value.trim() : "";
    await this.plugin.saveSettings();
  }

  // Older Obsidian versions.
  display(): void {
    this.containerEl.empty();
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
  }
}
