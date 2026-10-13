/*
 * PDF with Electron's own printer, as Obsidian's Export to PDF does: page boxes go in a `.print` container
 * (Obsidian's print CSS hides everything else) and the window prints them. Where the file goes is the export
 * dialog's job (export-dialog.ts).
 */
import { Notice, Platform } from "obsidian";

interface Electron {
  remote: { getCurrentWebContents(): { printToPDF(options: object): Promise<Uint8Array> } };
}

/** PDF export prints through Electron, which only the desktop app has. */
export const canExportPdf = (): boolean => Platform.isDesktopApp;

/** Print `pages` (each one `size` px) to PDF bytes; null on mobile or when printing fails (with a notice). */
export async function renderPdf(pages: HTMLElement[], size: { width: number; height: number }, cls = ""): Promise<Uint8Array | null> {
  if (!canExportPdf()) {
    new Notice("Export to PDF needs the desktop app.");
    return null;
  }
  const electron = window.require("electron") as Electron;
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
    return await electron.remote.getCurrentWebContents().printToPDF({
      pageSize: { width: size.width / 96, height: size.height / 96 },
      margins: { top: 0, bottom: 0, left: 0, right: 0 },
      printBackground: true,
    });
  } catch (e) {
    new Notice(`Couldn't export to PDF: ${e instanceof Error ? e.message : String(e)}`);
    return null;
  } finally {
    root.remove();
    notice.hide();
  }
}
