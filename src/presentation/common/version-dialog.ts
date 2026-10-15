/** File › Version history: the kept versions of a file, newest first, each with Preview and Restore. */
import { App, Modal, Notice, TFile } from "obsidian";
import { renderDocxEmbed, renderPptxEmbed, renderXlsxEmbed } from "./embeds";
import { keepVersion, versions } from "./version-history";
import { diffWords } from "../../core/utils/diff";
import { compareText, officeText, type TextChange } from "../../infrastructure/office-text";

export interface VersionedView {
  file: TFile | null;
  requestSave(): Promise<void>;
  onUnloadFile(file: TFile): Promise<void>;
  onLoadFile(file: TFile): Promise<void>;
}

const PREVIEW = { docx: renderDocxEmbed, pptx: renderPptxEmbed, xlsx: renderXlsxEmbed } as const;
/** More changes than this are counted, not listed. */
const MAX_LISTED = 500;
const size = (n: number) => (n < 1024 * 1024 ? `${Math.max(1, Math.round(n / 1024))} KB` : `${(n / 1024 / 1024).toFixed(1)} MB`);

export async function versionHistoryDialog(app: App, view: VersionedView, readOnly: boolean): Promise<void> {
  const file = view.file;
  if (!file || !versions) return;
  const list = await versions.list(file.path);
  const modal = new Modal(app);
  modal.titleEl.setText(`Version history: ${file.name}`);
  modal.modalEl.addClass("dw-versions-modal");
  const root = modal.contentEl.createDiv({ cls: "dw-versions" });
  if (!list.length) root.createDiv({ cls: "dw-muted", text: "No versions yet. Doc Works keeps one before it first saves over this file, then every 10 minutes while you edit." });
  const preview = modal.contentEl.createDiv({ cls: "dw-version-preview" });
  let cleanup: (() => void) | null = null;
  const clear = () => {
    cleanup?.();
    cleanup = null;
    preview.empty();
  };
  modal.onClose = clear;
  for (const v of list) {
    const when = new Date(v.time).toLocaleString();
    const row = root.createDiv({ cls: "dw-version" });
    row.createSpan({ cls: "dw-version-when", text: when });
    row.createSpan({ cls: "dw-version-size", text: size(v.size) });
    row.createEl("button", { text: "Preview" }).addEventListener("click", () => void (async () => {
      clear();
      const bytes = await versions?.bytes(v.id);
      const render = PREVIEW[file.extension.toLowerCase() as keyof typeof PREVIEW];
      if (!bytes || !render) return void preview.setText("This version can't be shown.");
      preview.createDiv({ cls: "dw-version-preview-title", text: `As it was on ${when}` });
      cleanup = await render(preview.createDiv({ cls: `dw-embed-body dw-embed-${file.extension.toLowerCase()}` }), bytes, "");
    })());
    row.createEl("button", { text: "Compare" }).addEventListener("click", () => void (async () => {
      clear();
      const ext = file.extension.toLowerCase();
      const old = await versions?.bytes(v.id);
      if (!old || !(ext in PREVIEW)) return void preview.setText("This version can't be compared.");
      // The file as it is now includes edits still waiting to be saved.
      await view.requestSave();
      const kind = ext as keyof typeof PREVIEW;
      const changes = compareText(kind, await officeText(kind, old), await officeText(kind, await app.vault.readBinary(file)));
      showChanges(preview, changes, when);
    })());
    if (readOnly) continue;
    // The file as it is now is kept first, so a restore can itself be undone from this list.
    row.createEl("button", { text: "Restore", cls: "mod-cta" }).addEventListener("click", () => void (async () => {
      const bytes = await versions?.bytes(v.id);
      if (!bytes) return void new Notice("This version is no longer kept.");
      await view.requestSave();
      await keepVersion(app, file, true);
      await view.onUnloadFile(file);
      await app.vault.modifyBinary(file, bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer);
      await view.onLoadFile(file);
      modal.close();
      new Notice(`Restored ${file.name} as it was on ${when}.`);
    })());
  }
  modal.open();
}

/** The text changes since a version: where, then the words struck out and added (Google's "Show changes"). */
function showChanges(el: HTMLElement, changes: TextChange[], when: string): void {
  el.createDiv({ cls: "dw-version-preview-title", text: `Changes since ${when}` });
  if (!changes.length) return void el.createDiv({ cls: "dw-muted", text: "No text changes. Formatting, pictures or layout may still differ." });
  el.createDiv({ cls: "dw-muted", text: changes.length === 1 ? "1 change" : `${changes.length} changes${changes.length > MAX_LISTED ? `, the first ${MAX_LISTED} shown` : ""}` });
  const list = el.createDiv({ cls: "dw-changes" });
  for (const c of changes.slice(0, MAX_LISTED)) {
    const item = list.createDiv({ cls: `dw-change is-${c.kind}` });
    if (c.where) item.createSpan({ cls: "dw-change-where", text: c.where });
    const text = item.createSpan({ cls: "dw-change-text" });
    const parts = c.kind === "changed" ? diffWords(c.before ?? "", c.after ?? "") : [{ kind: c.kind, text: (c.kind === "added" ? c.after : c.before) ?? "" }];
    for (const p of parts) {
      if (p.kind === "same") text.appendText(p.text);
      else text.createEl(p.kind === "added" ? "ins" : "del", { text: p.text });
    }
  }
}
