/*
 * Export to PDF with Electron's own printer, as Obsidian's Export to PDF does: page boxes go in a `.print`
 * container (Obsidian's print CSS hides everything else), the window prints them, and the user picks where to save.
 */
import { App, FileSystemAdapter, Notice, TFile } from "obsidian";

interface Electron {
  remote: {
    getCurrentWebContents(): { printToPDF(options: object): Promise<Uint8Array> };
    dialog: { showSaveDialog(options: object): Promise<{ canceled: boolean; filePath?: string }> };
  };
}

/** Print `pages` (each one `size` px) to a PDF the user saves; resolves with the saved path, or null if cancelled. */
export async function exportPdf(app: App, file: TFile, pages: HTMLElement[], size: { width: number; height: number }, cls = ""): Promise<string | null> {
  const electron = window.require("electron") as Electron;
  const fs = window.require("fs") as { promises: { writeFile(path: string, data: Uint8Array): Promise<void> } };
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
    await fs.promises.writeFile(filePath, data);
    new Notice(`Exported to ${filePath}`);
    return filePath;
  } catch (e) {
    new Notice(`Couldn't export to PDF: ${e instanceof Error ? e.message : String(e)}`);
    return null;
  } finally {
    root.remove();
    notice.hide();
  }
}
