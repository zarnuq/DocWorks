import { App, Notice, TFile } from "obsidian";

/** Copy a link to a place in `file` (#3, #Sheet!B4, #Heading) for pasting into a note, in the vault's link format. */
export function copyLink(app: App, file: TFile | null, subpath: string, what: string): void {
  if (!file) return;
  void navigator.clipboard.writeText(app.fileManager.generateMarkdownLink(file, "", subpath)).then(() => new Notice(`Copied a link to this ${what}.`));
}
