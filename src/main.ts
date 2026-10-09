import { Notice, Plugin, TFolder, WorkspaceLeaf } from "obsidian";
import { blankDocx, blankPptx } from "./infrastructure/blank";
import { DocxView, VIEW_TYPE_DOCX } from "./presentation/docx/docx-view";
import { PptxView, VIEW_TYPE_PPTX } from "./presentation/pptx/pptx-view";

const NEW_FILES = [
  ["new-presentation", "New presentation", "presentation", "pptx", blankPptx],
  ["new-document", "New document", "file-text", "docx", blankDocx],
] as const;

export default class DocWorksPlugin extends Plugin {
  async onload() {
    this.registerView(VIEW_TYPE_DOCX, (leaf: WorkspaceLeaf) => new DocxView(leaf));
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
