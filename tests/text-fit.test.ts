import { describe, it, expect, afterEach, vi } from "vitest";
import { letterSpacing, textWidthEm } from "../src/presentation/pptx/text-fit";
import { renderParagraphs } from "../src/presentation/pptx/text-html";
import type { Para, RunProps } from "../src/domain/pptx/slide-model";

/** Fake canvas: every character of the fallback font is 0.6em wide. */
class FakeCanvas {
  getContext() {
    return {
      font: "",
      measureText(this: { font: string }, text: string) {
        return { width: [...text].length * 0.6 * parseFloat(/([\d.]+)px/.exec(this.font)![1]) };
      },
    };
  }
}

afterEach(() => vi.unstubAllGlobals());

describe("textWidthEm", () => {
  it("sums advance widths from the font's table", () => {
    // Reference values from Carlito/Arial hmtx via fontTools.
    expect(textWidthEm("Assumptions", "Calibri", true, false)).toBeCloseTo(5.356, 1);
    expect(textWidthEm("Assumptions", "Calibri", false, false)).toBeCloseTo(5.219, 1);
    expect(textWidthEm("Hello world", "Arial", false, false)).toBeCloseTo(4.946, 1);
  });

  it("matches names case-insensitively and through aliases", () => {
    expect(textWidthEm("Hello", "helvetica", false, false)).toBe(textWidthEm("Hello", "Arial", false, false));
    expect(textWidthEm("Hello", "CALIBRI", false, false)).toBe(textWidthEm("Hello", "Calibri", false, false));
  });

  it("falls back to the nearest style a font has", () => {
    expect(textWidthEm("Hello", "Impact", true, true)).toBe(textWidthEm("Hello", "Impact", false, false));
  });

  it("returns null for fonts without a table", () => {
    expect(textWidthEm("Hello", "Wingdings", false, false)).toBeNull();
  });
});

describe("letterSpacing", () => {
  it("spreads the width difference over every character", () => {
    vi.stubGlobal("OffscreenCanvas", FakeCanvas);
    const em = textWidthEm("Assumptions", "Calibri", true, false)!;
    expect(letterSpacing("Assumptions", { font: "Calibri", b: true, i: false }, 40)).toBeCloseTo(((em - 11 * 0.6) * 40) / 11, 5);
  });

  it("measures characters missing from the table in the fallback font", () => {
    vi.stubGlobal("OffscreenCanvas", FakeCanvas);
    const em = textWidthEm("ab", "Calibri", false, false)!;
    expect(letterSpacing("a→b", { font: "Calibri", b: false, i: false }, 10)).toBeCloseTo(((em + 0.6 - 3 * 0.6) * 10) / 3, 5);
  });

  it("is null without a canvas or a table", () => {
    expect(letterSpacing("Hello", { font: "Calibri", b: false, i: false }, 10)).toBeNull();
    vi.stubGlobal("OffscreenCanvas", FakeCanvas);
    expect(letterSpacing("Hello", { font: "Wingdings", b: false, i: false }, 10)).toBeNull();
  });
});

describe("renderParagraphs width fitting", () => {
  const defaults: RunProps = { b: false, i: false, u: false, strike: false, sz: 20, color: "#000000", font: "Calibri", spc: 0, link: null };
  const para = (font: string): Para => ({
    runs: [{ ...defaults, font, text: "Assumptions" }],
    align: "l",
    lvl: 0,
    bullet: null,
    marL: 0,
    indent: 0,
    defaults,
  });

  it("sets letter-spacing so runs are as wide as in PowerPoint", () => {
    vi.stubGlobal("OffscreenCanvas", FakeCanvas);
    const host = document.createElement("div");
    renderParagraphs(host, [para("Calibri")], 2);
    const ls = parseFloat(host.querySelector("span")!.style.letterSpacing);
    expect(ls).toBeCloseTo(letterSpacing("Assumptions", { font: "Calibri", b: false, i: false }, 40)!, 3);
  });

  it("leaves unknown fonts alone", () => {
    vi.stubGlobal("OffscreenCanvas", FakeCanvas);
    const host = document.createElement("div");
    renderParagraphs(host, [para("Wingdings")], 2);
    expect(host.querySelector("span")!.style.letterSpacing).toBe("");
  });
});
