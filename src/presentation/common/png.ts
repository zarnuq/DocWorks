/*
 * Download a slide as PNG, as Google Slides' File › Download › PNG: the slide is drawn over everything at the top left
 * of the window, Chromium's DevTools protocol renders that box of the page at PNG_WIDTH pixels wide (the window's
 * own capture returns stale frames while it isn't focused), and the picture is saved next to the deck. Desktop only.
 */
import { App, normalizePath, Notice, TFile } from "obsidian";
import { canExportPdf } from "./pdf";

interface Electron {
  remote: {
    getCurrentWebContents(): {
      getZoomFactor?(): number;
      debugger: {
        isAttached(): boolean;
        attach(version?: string): void;
        detach(): void;
        sendCommand(method: string, params?: object): Promise<{ data: string }>;
      };
    };
  };
}

/** Width of the saved picture in pixels (Full HD for a 16:9 deck). */
export const PNG_WIDTH = 1920;

/** "<folder>/<deck> - slide N.png". */
export function pngPath(file: TFile, slideNumber: number): string {
  const folder = file.path.includes("/") ? file.path.slice(0, file.path.lastIndexOf("/") + 1) : "";
  return normalizePath(`${folder}${file.basename} - slide ${slideNumber}.png`);
}

/** Draw a slide `size` px big with `draw`, capture it and save it; resolves with the vault path, or null. */
export async function exportPng(app: App, file: TFile, slideNumber: number, size: { width: number; height: number }, draw: (host: HTMLElement) => void): Promise<string | null> {
  if (!canExportPdf()) {
    new Notice("Downloading a PNG needs the desktop app.");
    return null;
  }
  const host = document.body.createDiv({ cls: "dw-png-capture" });
  host.setCssProps({ "--dw-png-w": `${size.width}px`, "--dw-png-h": `${size.height}px` });
  try {
    draw(host);
    await document.fonts.ready;
    await Promise.all(Array.from(host.querySelectorAll("img")).map((img) => img.decode().catch(() => undefined)));
    // No waiting for animation frames: they stop while the window is in the background, and the capture below lays
    // out and paints the page itself.
    const contents = (window.require("electron") as Electron).remote.getCurrentWebContents();
    // Clip rectangles are in window units: CSS pixels times the app's zoom.
    const zoom = contents.getZoomFactor?.() ?? 1;
    const width = Math.round(size.width * zoom), height = Math.round(size.height * zoom);
    // A hair over, so Chromium's rounding of width × scale doesn't land a pixel short (1919 × 1079).
    const clip = { x: 0, y: 0, width, height, scale: (PNG_WIDTH + 0.01) / width };
    const dbg = contents.debugger;
    const attached = dbg.isAttached();
    if (!attached) dbg.attach("1.3");
    let shot: { data: string };
    try {
      shot = await dbg.sendCommand("Page.captureScreenshot", { format: "png", clip, captureBeyondViewport: false });
    } finally {
      if (!attached) dbg.detach();
    }
    const bytes = Uint8Array.from(atob(shot.data), (c) => c.charCodeAt(0));
    const data = bytes.buffer;
    const path = pngPath(file, slideNumber);
    const existing = app.vault.getFileByPath(path);
    if (existing) await app.vault.modifyBinary(existing, data);
    else await app.vault.createBinary(path, data);
    new Notice(`Saved ${path}`);
    return path;
  } catch (e) {
    new Notice(`Couldn't save the PNG: ${e instanceof Error ? e.message : String(e)}`);
    return null;
  } finally {
    host.remove();
  }
}
