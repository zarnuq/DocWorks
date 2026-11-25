import { App, FuzzySuggestModal, TFile } from "obsidian";

/** Pick one of `files` by fuzzy search. */
export class FilePicker extends FuzzySuggestModal<TFile> {
  constructor(app: App, private files: TFile[], private choose: (f: TFile) => void, placeholder?: string) {
    super(app);
    if (placeholder) this.setPlaceholder(placeholder);
  }
  getItems(): TFile[] {
    return this.files;
  }
  getItemText(f: TFile): string {
    return f.path;
  }
  onChooseItem(f: TFile): void {
    this.choose(f);
  }
}
