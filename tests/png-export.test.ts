import { afterEach, describe, expect, it, vi } from "vitest";
import { App, TFile } from "./obsidian-stub";
import { exportPng, pngPath } from "../src/presentation/common/png";

afterEach(() => {
  delete (window as any).require;
  document.body.empty();
});

describe("PNG export", () => {
  it("names the file after the deck and slide, next to the deck", () => {
    expect(pngPath(new TFile("talks/Deck.pptx") as any, 3)).toBe("talks/Deck - slide 3.png");
    expect(pngPath(new TFile("Deck.pptx") as any, 1)).toBe("Deck - slide 1.png");
  });

  it("captures the drawn slide's box from the window and saves the PNG in the vault, removing the drawing", async () => {
    Object.defineProperty(document, "fonts", { value: { ready: Promise.resolve() }, configurable: true });
    const calls: [string, any][] = [];
    let attached = false;
    (window as any).require = () => ({
      remote: {
        getCurrentWebContents: () => ({
          getZoomFactor: () => 0.5,
          debugger: {
            isAttached: () => attached,
            attach: () => (attached = true),
            detach: () => (attached = false),
            sendCommand: async (method: string, params: object) => (calls.push([method, params]), { data: btoa(String.fromCharCode(137, 80)) }),
          },
        }),
      },
    });
    const app = new App();
    const created: [string, ArrayBuffer][] = [];
    (app.vault as any).getFileByPath = () => null;
    (app.vault as any).createBinary = async (p: string, d: ArrayBuffer) => created.push([p, d]);
    let drawnInto: HTMLElement | null = null;
    const path = await exportPng(app as any, new TFile("Deck.pptx") as any, 2, { width: 800, height: 450 }, (host) => {
      drawnInto = host;
      host.createDiv({ text: "slide" });
    });
    const { Notice } = await import("./obsidian-stub");
    expect([path, Notice.messages.at(-1)]).toEqual(["Deck - slide 2.png", "Saved Deck - slide 2.png"]);
    expect(calls).toMatchObject([["Page.captureScreenshot", { format: "png", clip: { x: 0, y: 0, width: 400, height: 225 }, captureBeyondViewport: false }]]);
    expect(calls[0][1].clip.scale * 400).toBeCloseTo(1920, 1);
    expect(attached).toBe(false);
    expect(created[0][0]).toBe("Deck - slide 2.png");
    expect(new Uint8Array(created[0][1])).toEqual(new Uint8Array([137, 80]));
    expect(drawnInto!.isConnected).toBe(false);
  });
});
