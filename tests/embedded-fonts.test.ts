import { describe, it, expect, afterEach, vi } from "vitest";
import { loadEmbeddedFonts } from "../src/presentation/pptx/embedded-fonts";

class FakeFace {
  constructor(public family: string, public data: ArrayBuffer | Uint8Array, public desc: { weight: string; style: string }) {}
  load() {
    return Promise.resolve(this);
  }
}

describe("loadEmbeddedFonts", () => {
  const added: FakeFace[] = [];
  const fonts = { add: (f: FakeFace) => added.push(f), delete: (f: FakeFace) => added.splice(added.indexOf(f), 1) };
  afterEach(() => {
    added.length = 0;
    vi.unstubAllGlobals();
  });

  it("decodes each font, registers it under the deck's font name and style, and removes it again", async () => {
    vi.stubGlobal("FontFace", FakeFace);
    Object.defineProperty(document, "fonts", { value: fonts, configurable: true });
    const { loaded, unload } = await loadEmbeddedFonts(
      [
        { family: "Oswald", bold: false, italic: false, data: new Uint8Array([1]) },
        { family: "Oswald", bold: true, italic: false, data: new Uint8Array([2]) },
      ],
      (eot) => new Uint8Array([eot[0] * 10])
    );
    expect(added.map((f) => [f.family, f.desc.weight, f.desc.style, Array.from(new Uint8Array(f.data))])).toEqual([
      ["Oswald", "400", "normal", [10]],
      ["Oswald", "700", "normal", [20]],
    ]);
    expect(loaded).toBe(2);
    unload();
    expect(added).toEqual([]);
  });

  it("skips a font that fails to decode and keeps the others", async () => {
    vi.stubGlobal("FontFace", FakeFace);
    Object.defineProperty(document, "fonts", { value: fonts, configurable: true });
    vi.spyOn(console, "warn").mockImplementation(() => {});
    await loadEmbeddedFonts(
      [
        { family: "Bad", bold: false, italic: false, data: new Uint8Array([0]) },
        { family: "Good", bold: false, italic: true, data: new Uint8Array([1]) },
      ],
      (eot) => {
        if (!eot[0]) throw new Error("corrupt");
        return eot;
      }
    );
    expect(added.map((f) => [f.family, f.desc.style])).toEqual([["Good", "italic"]]);
  });

  it("does nothing where the browser has no FontFace", async () => {
    vi.stubGlobal("FontFace", undefined);
    const { loaded, unload } = await loadEmbeddedFonts([{ family: "X", bold: false, italic: false, data: new Uint8Array([1]) }], (e) => e);
    expect(loaded).toBe(0);
    expect(unload).toBeTypeOf("function");
  });
});
