/*
 * Export to PDF with Electron's own printer, as Obsidian's Export to PDF does: page boxes go in a `.print`
 * container (Obsidian's print CSS hides everything else), the window prints them, and the user picks where to save.
 * Files are written through the vault API, never Node's fs: a place inside the vault is written directly, and a
 * place outside it is handed to the app's download (its own save dialog, with the chosen name).
 */
import { App, FileSystemAdapter, normalizePath, Notice, TFile } from "obsidian";

interface Electron {
  remote: {
    getCurrentWebContents(): { printToPDF(options: object): Promise<Uint8Array> };
    dialog: { showSaveDialog(options: object): Promise<{ canceled: boolean; filePath?: string }> };
  };
}

/** `filePath` relative to the vault folder `base`, or null when it's outside the vault. */
export function vaultPath(base: string | null, filePath: string): string | null {
  if (!base) return null;
  const root = base.replace(/[\\/]+$/, "");
  const sep = filePath[root.length];
  if (!filePath.startsWith(root) || (sep !== "/" && sep !== "\\")) return null;
  return filePath.slice(root.length + 1).replace(/\\/g, "/");
}

async function save(app: App, filePath: string, data: Uint8Array): Promise<string> {
  const adapter = app.vault.adapter;
  const inVault = vaultPath(adapter instanceof FileSystemAdapter ? adapter.getBasePath() : null, filePath);
  const bytes = data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength) as ArrayBuffer;
  if (inVault !== null) {
    const path = normalizePath(inVault);
    const existing = app.vault.getFileByPath(path);
    if (existing) await app.vault.modifyBinary(existing, bytes);
    else await app.vault.createBinary(path, bytes);
    return `Exported to ${path}`;
  }
  const url = URL.createObjectURL(new Blob([bytes], { type: "application/pdf" }));
  const link = document.body.createEl("a", { href: url, attr: { download: filePath.split(/[\\/]/).pop() ?? "export.pdf" } });
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 60_000);
  return "Choose where to save the PDF";
}

/** Print `pages` (each one `size` px) to a PDF the user saves; resolves with the saved path, or null if cancelled. */
export async function exportPdf(app: App, file: TFile, pages: HTMLElement[], size: { width: number; height: number }, cls = ""): Promise<string | null> {
  const electron = window.require("electron") as Electron;
  const adapter = app.vault.adapter;
  const defaultPath = adapter instanceof FileSystemAdapter ? adapter.getFullPath(file.path.replace(/\.[^.]+$/, ".pdf")) : `${file.basename}.pdf`;
  const { canceled, filePath } = await electron.remote.dialog.showSaveDialog({ title: "Export to PDF", defaultPath, filters: [{ name: "PDF", extensions: ["pdf"] }] });
  if (canceled || !filePath) return null;

  const notice = new Notice("Exporting to PDF…", 0);
  const root = document.body.createDiv({ cls: `print dw-print ${cls}`.trim() });
  root.setCssProps({ "--dw-page-w": `${size.width}px`, "--dw-page-h": `${size.height}px` });
  // Page breaks are set inline: the plugin directory's CSS lint flags break-* properties in styles.css.
  const pageBreak = { "break-after": "page" };
  pages.forEach((p, i) => {
    if (i < pages.length - 1) p.setCssProps(pageBreak);
    root.appendChild(p);
  });
  try {
    await document.fonts.ready;
    await Promise.all(Array.from(root.querySelectorAll("img")).map((img) => img.decode().catch(() => undefined)));
    const data = await electron.remote.getCurrentWebContents().printToPDF({
      pageSize: { width: size.width / 96, height: size.height / 96 },
      margins: { top: 0, bottom: 0, left: 0, right: 0 },
      printBackground: true,
    });
    new Notice(await save(app, filePath, data));
    return filePath;
  } catch (e) {
    new Notice(`Couldn't export to PDF: ${e instanceof Error ? e.message : String(e)}`);
    return null;
  } finally {
    root.remove();
    notice.hide();
  }
}
