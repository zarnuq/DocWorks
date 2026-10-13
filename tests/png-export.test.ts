import { afterEach, describe, expect, it } from "vitest";
import { renderPng } from "../src/presentation/common/png";

afterEach(() => {
  delete (window as any).require;
  document.body.empty();
});

describe("PNG capture", () => {
  it("captures the drawn slide's box at 1920 px wide with the DevTools protocol, removing the drawing", async () => {
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
    let drawnInto: HTMLElement | null = null;
    const bytes = await renderPng({ width: 800, height: 450 }, (host) => {
      drawnInto = host;
      host.createDiv({ text: "slide" });
    });
    expect(bytes).toEqual(new Uint8Array([137, 80]));
    expect(calls).toMatchObject([["Page.captureScreenshot", { format: "png", clip: { x: 0, y: 0, width: 400, height: 225 }, captureBeyondViewport: false }]]);
    expect(calls[0][1].clip.scale * 400).toBeCloseTo(1920, 1);
    expect(attached).toBe(false);
    expect(drawnInto!.isConnected).toBe(false);
  });
});
