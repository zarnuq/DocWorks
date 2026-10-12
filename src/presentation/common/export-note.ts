/*
 * Export a markdown note as a Word document next to it: Obsidian renders the note (callouts, wikilinks, embeds),
 * and the rendered HTML is turned into the document. Pictures come from the vault; web images are left out.
 */
import { App, Component, MarkdownRenderer, Notice, TFile } from "obsidian";
import { IMAGE_MIME } from "../../core/utils/images";
import { noteToDocx } from "../../domain/docx/from-html";
import { blankDocx } from "../../infrastructure/blank";

type Picture = { bytes: Uint8Array; ext: string; width: number; height: number };

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
  const notice = new Notice("Exporting to .docx…", 0);
  const component = new Component();
  component.load();
  try {
    // Properties stay in Obsidian; the document starts at the note's text.
    const markdown = (await app.vault.cachedRead(file)).replace(/^---\r?\n[\s\S]*?\r?\n---\r?\n?/, "");
    const el = createDiv();
    await MarkdownRenderer.render(app, markdown, el, file.path, component);
    // ponytail: embeds fill in asynchronously; a short wait covers typical notes, not slow huge embeds.
    await new Promise((r) => window.setTimeout(r, 400));
    const pictures = new Map<HTMLImageElement, Picture>();
    for (const img of Array.from(el.querySelectorAll("img"))) {
      const target = pictureFile(app, img, file.path);
      const pic = target && (await loadPicture(app, target));
      if (pic) pictures.set(img, pic);
    }
    const bytes = await noteToDocx(await blankDocx(), el, (img) => pictures.get(img) ?? null);
    const dir = file.parent && !file.parent.isRoot() ? `${file.parent.path}/` : "";
    let path = `${dir}${file.basename}.docx`;
    for (let i = 1; app.vault.getAbstractFileByPath(path); i++) path = `${dir}${file.basename} ${i}.docx`;
    const created = await app.vault.createBinary(path, bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer);
    new Notice(`Exported to ${path}`);
    await app.workspace.getLeaf("tab").openFile(created);
  } catch (err) {
    new Notice(`Couldn't export to .docx: ${(err as Error).message}`);
  } finally {
    component.unload();
    notice.hide();
  }
}
