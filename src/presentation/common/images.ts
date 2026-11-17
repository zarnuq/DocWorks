import type { App, TFile } from "obsidian";
import { IMAGE_MIME } from "../../core/utils/images";

/**
 * Vault pictures by link path, as written in pasted Markdown (![[name.png]] or a path relative to `source`), as the
 * files the editors' image inserts take. Links that aren't a picture in the vault are left out.
 */
export async function vaultPictures(app: App, source: string, links: string[]): Promise<Map<string, File>> {
  const out = new Map<string, File>();
  for (const link of links) {
    const file = vaultPicture(app, source, link);
    // One that can't be read is left out like one that isn't there: the rest of the paste goes on.
    const bytes = file && (await app.vault.readBinary(file).catch(() => null));
    if (file && bytes) out.set(link, new File([bytes], file.name, { type: IMAGE_MIME[file.extension.toLowerCase()] }));
  }
  return out;
}

/** The vault picture a link path names, if there is one. */
export function vaultPicture(app: App, source: string, link: string): TFile | null {
  const file = app.metadataCache.getFirstLinkpathDest(link.split("#")[0], source);
  return file && IMAGE_MIME[file.extension.toLowerCase()] ? file : null;
}

/** Open the image picker and pass the selected file to the editor. */
export function pickImage(onPick: (file: File) => void): void {
  const input = document.body.createEl("input", { type: "file", attr: { accept: "image/*" } });
  input.setCssStyles({ display: "none" });
  input.onchange = () => {
    const file = input.files?.[0];
    if (file) onPick(file);
    input.remove();
  };
  input.click();
}
