import { Notice, Plugin, WorkspaceLeaf } from "obsidian";
import { DocxView, VIEW_TYPE_DOCX } from "./presentation/docx/docx-view";
import { PptxView, VIEW_TYPE_PPTX } from "./presentation/pptx/pptx-view";

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
