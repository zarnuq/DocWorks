import { describe, it, expect, afterEach, vi } from "vitest";
import { kerningEm, letterSpacing, textWidthEm } from "../src/presentation/pptx/text-fit";
import { renderParagraphs } from "../src/presentation/pptx/text-html";
import type { Para, RunProps } from "../src/domain/pptx/slide-model";

/** Fake canvas: every character of the fallback font is 0.6em wide. */
class FakeCanvas {
  getContext() {
    return {
      font: "",
      measureText(this: { font: string }, text: string) {
        // A missing family resolves differently against the two availability probes.
        if (this.font.endsWith(', monospace')) return { width: text.length * 10 };
        if (this.font.endsWith(', serif')) return { width: text.length * 8 };
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
  it("preserves native glyph spacing when the requested font resolves in the browser", () => {
    class NativeCanvas {
      getContext() {
        return { font: "", measureText: (text: string) => ({ width: text.length * 7 }) };
      }
    }
    vi.stubGlobal("OffscreenCanvas", NativeCanvas);
    expect(letterSpacing("AVATAR", { font: "Arial", b: true, i: true, kern: true }, 20)).toBe(0);
  });

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
  const defaults: RunProps = { b: false, i: false, u: false, strike: false, sz: 20, color: "#000000", font: "Calibri", spc: 0, cap: "none", link: null };
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

describe("letter spacing for capitals", () => {
  it("measures all-caps text in capitals, as it is drawn", () => {
    vi.stubGlobal("OffscreenCanvas", FakeCanvas);
    const host = document.createElement("div");
    const run = { text: "Contingent", b: true, i: false, u: false, strike: false, sz: 20, color: "#000000", font: "Calibri", spc: 0, cap: "all" as const, link: null };
    renderParagraphs(host, [{ runs: [run], align: "l", lvl: 0, bullet: null, defaults: { ...run, text: undefined } as never } as never], 1, 1, 0, {});
    const span = host.querySelector("span")!;
    const caps = letterSpacing("CONTINGENT", run, 20)!;
    const lower = letterSpacing("Contingent", run, 20)!;
    expect(caps).not.toBeCloseTo(lower, 3);
    expect(parseFloat(span.style.letterSpacing)).toBeCloseTo(caps, 3);
  });
});

describe("kerning", () => {
  it("sums the font's pair adjustments (AV and To close up in every common font)", () => {
    for (const font of ["Calibri", "Arial", "Times New Roman", "Aptos"]) {
      expect(kerningEm("AV", font, false)).toBeLessThan(0);
      expect(kerningEm("To", font, true)).toBeLessThan(0);
    }
    expect(kerningEm("xx", "Calibri", false)).toBe(0);
    expect(kerningEm("AV", "Wingdings", false)).toBe(0);
    // Pairs are only between neighbours.
    expect(kerningEm("A V", "Arial", false)).toBe(kerningEm("A ", "Arial", false) + kerningEm(" V", "Arial", false));
  });

  it("narrows the width text is fitted to when kerning is on", () => {
    vi.stubGlobal("OffscreenCanvas", FakeCanvas);
    const plain = letterSpacing("AVATAR", { font: "Calibri", b: false, i: false }, 40)!;
    const kerned = letterSpacing("AVATAR", { font: "Calibri", b: false, i: false, kern: true }, 40)!;
    expect(kerned).toBeCloseTo(plain + (kerningEm("AVATAR", "Calibri", false) * 40) / 6, 5);
  });

  it("kerns slide text from PowerPoint's minimum size (a:rPr kern), not below it or without it", () => {
    vi.stubGlobal("OffscreenCanvas", FakeCanvas);
    const host = document.body.appendChild(document.createElement("div"));
    const defaults: RunProps = { b: false, i: false, u: false, strike: false, sz: 20, color: "#000000", font: "Calibri", spc: 0, cap: "none", link: null };
    const run = (extra: Partial<RunProps>) => ({ ...defaults, text: "AVATAR", ...extra });
    const para = (r: ReturnType<typeof run>): Para => ({ runs: [r], align: "l", lvl: 0, bullet: null, marL: 0, indent: 0, defaults });
    const spacing = (r: ReturnType<typeof run>) => {
      renderParagraphs(host, [para(r)], 1);
      return parseFloat(host.querySelector("span")!.style.letterSpacing);
    };
    const off = spacing(run({}));
    expect(spacing(run({ kern: 12 }))).toBeLessThan(off);
    expect(spacing(run({ kern: 24 }))).toBeCloseTo(off, 5);
    host.remove();
  });
});
