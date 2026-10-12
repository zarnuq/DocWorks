import { App, Notice, TAbstractFile, TFile } from "obsidian";
import { Node as PMNode } from "prosemirror-model";
import { docxToMarkdown } from "../../domain/docx/to-markdown";
import { DocxDocument } from "../../infrastructure/docx/docx-document";

/** Create an editable note beside a Word file, with a dedicated folder for its pictures. */
export async function exportDocxToMarkdown(app: App, file: TFile, current?: { document: DocxDocument; content: PMNode }): Promise<TFile | null> {
  const notice = new Notice("Exporting to Markdown…", 0);
  const created: TAbstractFile[] = [];
  let note: TFile | null = null;
  try {
    const docx = current?.document ?? await DocxDocument.load(await app.vault.readBinary(file));
    const directory = file.path.slice(0, file.path.lastIndexOf("/") + 1);
    let name = file.basename;
    for (let i = 1; app.vault.getAbstractFileByPath(`${directory}${name}.md`) || app.vault.getAbstractFileByPath(`${directory}${name} attachments`); i++) name = `${file.basename} ${i}`;
    const folder = `${directory}${name} attachments`;
    const result = await docxToMarkdown(docx, `${name} attachments`, current?.content);
    if (result.images.length) {
      created.push(await app.vault.createFolder(folder));
      for (const image of result.images) {
        const data = image.bytes.buffer.slice(image.bytes.byteOffset, image.bytes.byteOffset + image.bytes.byteLength) as ArrayBuffer;
        created.push(await app.vault.createBinary(`${folder}/${image.name}`, data));
      }
    }
    note = await app.vault.create(`${directory}${name}.md`, result.markdown);
    new Notice(`Exported to ${note.path}${result.warnings.length ? `\n${result.warnings.join("\n")}` : ""}`);
    await app.workspace.getLeaf("tab").openFile(note);
    return note;
  } catch (err) {
    // Roll back only this export's attachments; an existing note or folder is never removed.
    if (!note) for (const asset of created.reverse()) {
      try { await app.fileManager.trashFile(asset); } catch { /* Keep the original export error. */ }
    }
    new Notice(`Couldn't export to Markdown: ${(err as Error).message}`);
    return note;
  } finally { notice.hide(); }
}
