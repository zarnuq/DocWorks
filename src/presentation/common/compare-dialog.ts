/*
 * Google Docs' Tools › Compare documents: pick the comparison document and who the differences are attributed to.
 */
import { App, ButtonComponent, Modal, Setting, TFile } from "obsidian";
import { FilePicker } from "./file-picker";

/** Resolves the chosen file and author, or null when cancelled. */
export function compareDialog(app: App, files: TFile[], author: string): Promise<{ file: TFile; author: string } | null> {
  return new Promise((resolve) => {
    let result: { file: TFile; author: string } | null = null;
    let chosen: TFile | null = null;
    const modal = new Modal(app);
    modal.titleEl.setText("Compare documents");
    let compareBtn: ButtonComponent | null = null;
    const pick = new Setting(modal.contentEl).setName("Select the comparison document").setDesc("No document chosen.");
    pick.addButton((b) =>
      b.setButtonText("Choose…").onClick(() =>
        new FilePicker(app, files, (f) => {
          chosen = f;
          pick.setDesc(f.path);
          compareBtn?.setDisabled(false);
        }, "Compare with…").open()
      )
    );
    new Setting(modal.contentEl).setName("Attribute differences to").addText((t) => t.setValue(author).onChange((v) => (author = v)));
    new Setting(modal.contentEl)
      .addButton((b) => b.setButtonText("Cancel").onClick(() => modal.close()))
      .addButton((b) => {
        b.setButtonText("Compare")
          .setCta()
          .onClick(() => {
            if (!chosen) return;
            result = { file: chosen, author: author.trim() || "Doc Works" };
            modal.close();
          });
        compareBtn = b.setDisabled(true);
      });
    modal.onClose = () => resolve(result);
    modal.open();
  });
}
