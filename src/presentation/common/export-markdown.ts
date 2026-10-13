import { App, Notice, TAbstractFile, TFile } from "obsidian";
import { Node as PMNode } from "prosemirror-model";
import { docxToMarkdown } from "../../domain/docx/to-markdown";
import { DocxDocument } from "../../infrastructure/docx/docx-document";
import { exportDialog } from "./export-dialog";

/**
 * Write `file` as an editable note at vault path `path` (default: beside it, numbered if taken), with its pictures in
 * a "<name> attachments" folder next to the note. Replacing a note replaces same-named pictures in that folder.
 */
export async function exportDocxToMarkdown(app: App, file: TFile, current?: { document: DocxDocument; content: PMNode }, path?: string, open = true): Promise<TFile | null> {
  const notice = new Notice("Exporting to Markdown…", 0);
  const created: TAbstractFile[] = [];
  let note: TFile | null = null;
  try {
    const docx = current?.document ?? await DocxDocument.load(await app.vault.readBinary(file));
    let directory = file.path.slice(0, file.path.lastIndexOf("/") + 1);
    let name = file.basename;
    if (path) {
      directory = path.slice(0, path.lastIndexOf("/") + 1);
      name = path.slice(directory.length).replace(/\.md$/, "");
    } else {
      for (let i = 1; app.vault.getAbstractFileByPath(`${directory}${name}.md`) || app.vault.getAbstractFileByPath(`${directory}${name} attachments`); i++) name = `${file.basename} ${i}`;
    }
    const folder = `${directory}${name} attachments`;
    const result = await docxToMarkdown(docx, `${name} attachments`, current?.content);
    if (result.images.length) {
      if (!app.vault.getAbstractFileByPath(folder)) created.push(await app.vault.createFolder(folder));
      for (const image of result.images) {
        const data = image.bytes.buffer.slice(image.bytes.byteOffset, image.bytes.byteOffset + image.bytes.byteLength) as ArrayBuffer;
        const at = `${folder}/${image.name}`;
        const old = app.vault.getAbstractFileByPath(at);
        if (old instanceof TFile) await app.vault.modifyBinary(old, data);
        else created.push(await app.vault.createBinary(at, data));
      }
    }
    const existing = app.vault.getAbstractFileByPath(`${directory}${name}.md`);
    if (existing instanceof TFile) await app.vault.modify((note = existing), result.markdown);
    else note = await app.vault.create(`${directory}${name}.md`, result.markdown);
    new Notice(`Exported to ${note.path}${result.warnings.length ? `\n${result.warnings.join("\n")}` : ""}`);
    if (open) await app.workspace.getLeaf("tab").openFile(note);
    return note;
  } catch (err) {
    // Roll back only this export's new files; an existing note, folder or picture is never removed.
    if (!note) for (const asset of created.reverse()) {
      try { await app.fileManager.trashFile(asset); } catch { /* Keep the original export error. */ }
    }
    new Notice(`Couldn't export to Markdown: ${(err as Error).message}`);
    return note;
  } finally { notice.hide(); }
}

/** Export a Word file (not open in a view) to Markdown through the export dialog. */
export async function exportDocxFile(app: App, file: TFile): Promise<void> {
  const choice = await exportDialog(app, file, [MARKDOWN_FORMAT]);
  if (choice && "vault" in choice.target) await exportDocxToMarkdown(app, file, undefined, choice.target.vault, choice.open);
}

export const MARKDOWN_FORMAT = { id: "md", label: "Markdown note (.md)", ext: "md" };
