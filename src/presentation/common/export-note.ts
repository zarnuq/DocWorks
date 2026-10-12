/*
 * Export a markdown note as a Word document next to it: Obsidian renders the note (callouts, wikilinks, embeds),
 * and the rendered HTML is turned into the document. Pictures come from the vault; web images are left out.
 */
import { App, Component, Menu, Modal, Notice, Setting, TFile } from "obsidian";
import { IMAGE_MIME } from "../../core/utils/images";
import { DEFAULT_NOTE_EXPORT_OPTIONS, NoteExportOptions, noteToDocx } from "../../domain/docx/from-html";
import { blankDocx } from "../../infrastructure/blank";
import { renderNoteForExport } from "./render-note-export";

type Picture = { bytes: Uint8Array; ext: string; width: number; height: number };

/** Obsidian adds PDF after firing file-menu; append DOCX just before showing to keep them together. */
export function addNoteExportToMenu(menu: Menu, app: App, file: TFile): void {
  addExportToMenu(menu, "Export to .docx", () => void exportNoteToDocx(app, file));
}

/** Append after built-in export actions using either supported menu display method. */
export function addExportToMenu(menu: Menu, title: string, run: () => void): void {
  let added = false;
  const add = () => {
    if (added) return;
    added = true;
    menu.addItem((i) => i.setTitle(title).setIcon("file-output").setSection("action").onClick(run));
  };
  const showAtPosition = menu.showAtPosition.bind(menu);
  const showAtMouseEvent = menu.showAtMouseEvent.bind(menu);
  menu.showAtPosition = (...args) => { add(); return showAtPosition(...args); };
  menu.showAtMouseEvent = (...args) => { add(); return showAtMouseEvent(...args); };
}

/** The vault picture an <img> in the rendered note shows (an ![[embed]] or a relative ![](path)), if any. */
function pictureFile(app: App, img: HTMLImageElement, source: string): TFile | null {
  const link = img.closest(".internal-embed")?.getAttribute("src") ?? img.getAttribute("src") ?? "";
  if (/^[a-z]+:/i.test(link) && !img.closest(".internal-embed")) return null; // web or app:// URL
  let path = link.split("#")[0].split("|")[0];
  try {
    path = decodeURIComponent(path);
  } catch {
    // keep as written
  }
  const file = app.metadataCache.getFirstLinkpathDest(path, source);
  return file && IMAGE_MIME[file.extension.toLowerCase()] ? file : null;
}

async function loadPicture(app: App, file: TFile): Promise<Picture | null> {
  try {
    const bytes = new Uint8Array(await app.vault.readBinary(file));
    const bitmap = await createImageBitmap(new Blob([bytes], { type: IMAGE_MIME[file.extension.toLowerCase()] }));
    const out = { bytes, ext: file.extension.toLowerCase(), width: bitmap.width, height: bitmap.height };
    bitmap.close();
    return out;
  } catch {
    return null;
  }
}

export async function exportNoteToDocx(app: App, file: TFile): Promise<void> {
  const options = await promptNoteExport(app, file.basename);
  if (!options) return;
  const notice = new Notice("Exporting to .docx…", 0);
  const component = new Component();
  component.load();
  try {
    const { el, warnings } = await renderNoteForExport(app, file, component);
    const pictures = new Map<HTMLImageElement, Picture>();
    for (const img of Array.from(el.querySelectorAll("img"))) {
      const target = pictureFile(app, img, img.dataset.exportSource ?? file.path);
      const pic = target && (await loadPicture(app, target));
      if (pic) pictures.set(img, pic);
    }
    const bytes = await noteToDocx(await blankDocx(), el, (img) => pictures.get(img) ?? null, options);
    const dir = file.parent && !file.parent.isRoot() ? `${file.parent.path}/` : "";
    let path = `${dir}${file.basename}.docx`;
    for (let i = 1; app.vault.getAbstractFileByPath(path); i++) path = `${dir}${file.basename} ${i}.docx`;
    const created = await app.vault.createBinary(path, bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer);
    new Notice(`Exported to ${path}${warnings.length ? `\n${warnings.join("\n")}` : ""}`);
    await app.workspace.getLeaf("tab").openFile(created);
  } catch (err) {
    new Notice(`Couldn't export to .docx: ${(err as Error).message}`);
  } finally {
    component.unload();
    notice.hide();
  }
}

/** Closing the dialog without exporting leaves the vault untouched. */
export function promptNoteExport(app: App, basename: string): Promise<NoteExportOptions | null> {
  return new Promise((resolve) => {
    const options = { ...DEFAULT_NOTE_EXPORT_OPTIONS };
    let result: NoteExportOptions | null = null;
    const modal = new Modal(app);
    modal.titleEl.setText("Export to .docx");
    modal.contentEl.createEl("p", { text: `Export “${basename}” to DOCX with the settings below.` });
    new Setting(modal.contentEl).setName("Include file name as title").addToggle((toggle) =>
      toggle.setValue(false).onChange((value) => { options.title = value ? basename : ""; })
    );
    new Setting(modal.contentEl).setName("Page size").addDropdown((select) =>
      select.addOptions({ letter: "Letter", legal: "Legal", a4: "A4", a5: "A5" }).setValue(options.pageSize).onChange((value) => { options.pageSize = value as NoteExportOptions["pageSize"]; })
    );
    new Setting(modal.contentEl).setName("Landscape").addToggle((toggle) =>
      toggle.setValue(options.landscape).onChange((value) => { options.landscape = value; })
    );
    new Setting(modal.contentEl).setName("Margin").addDropdown((select) =>
      select.addOptions({ default: "Default (1 inch)", narrow: "Narrow (½ inch)", wide: "Wide (2 inches)", none: "None" }).setValue(options.margin).onChange((value) => { options.margin = value as NoteExportOptions["margin"]; })
    );
    new Setting(modal.contentEl).setName("Downscale percent").setDesc("Reduce text and picture sizes.").addSlider((slider) =>
      slider.setLimits(10, 100, 1).setValue(options.scale).onChange((value) => { options.scale = value; })
    );
    new Setting(modal.contentEl)
      .addButton((button) => button.setButtonText("Export to .docx").setCta().onClick(() => { result = { ...options }; modal.close(); }))
      .addButton((button) => button.setButtonText("Cancel").onClick(() => modal.close()));
    modal.onClose = () => resolve(result);
    modal.open();
  });
}
