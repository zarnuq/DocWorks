import { Node as PMNode } from "prosemirror-model";
import { NS } from "../../core/utils/xml";
import { letterNumber, romanNumber } from "../../core/utils/numbering";
import { w, wVal } from "./docx-reader";
import { attrsOf, ParagraphAttrs } from "./schema";

/** Marker drawn before a list paragraph. */
export interface ListMark {
  text: string;
  indentPt: number;
  hangingPt: number;
}

const isOn = (el: Element | null) => !!el && !["0", "false", "off"].includes(wVal(el) ?? "");
/** Symbol / Wingdings private-use bullets and their Unicode look-alikes. */
const BULLET_GLYPHS: Record<string, string> = { "": "•", "": "▪", "": "➢", "": "✓", "": "❖", o: "◦", "": "•" };

function formatNumber(n: number, fmt: string | null): string {
  switch (fmt) {
    case "lowerLetter":
      return letterNumber(n);
    case "upperLetter":
      return letterNumber(n).toUpperCase();
    case "lowerRoman":
      return romanNumber(n);
    case "upperRoman":
      return romanNumber(n).toUpperCase();
    case "decimalZero":
      return String(n).padStart(2, "0");
    default:
      return String(n);
  }
}

/**
 * Display-only view of styles.xml and numbering.xml: CSS for paragraph styles
 * and list markers. Nothing here is written back.
 */
/** The w:lvl for `ilvl` of list `numId` in numbering.xml. */
export function numberingLevel(numbering: Document | null, numId: string, ilvl: number): Element | null {
  const nums = Array.from(numbering?.getElementsByTagNameNS(NS.w, "num") ?? []);
  const absId = wVal(w(nums.find((n) => wVal(n, "numId") === numId), "abstractNumId"));
  const abs = Array.from(numbering?.getElementsByTagNameNS(NS.w, "abstractNum") ?? []).find((a) => wVal(a, "abstractNumId") === absId);
  return Array.from(abs?.getElementsByTagNameNS(NS.w, "lvl") ?? []).find((l) => wVal(l, "ilvl") === String(ilvl)) ?? null;
}

/**
 * Word's "single" line height as a multiple of the font size. It comes from each font's own metrics, so CSS's
 * line-height: 1 draws Word text too tight (Garamond by a quarter).
 */
// ponytail: hand-measured table for common fonts; generate from font files (like font-metrics.ts) if more fonts matter.
const SINGLE_LINE: Record<string, number> = {
  calibri: 1.22, "calibri light": 1.22, cambria: 1.17, garamond: 1.33, "times new roman": 1.15, arial: 1.15, aptos: 1.2,
  georgia: 1.14, verdana: 1.22, "segoe ui": 1.33, tahoma: 1.21, "century gothic": 1.23, "book antiqua": 1.17,
  "palatino linotype": 1.14, "trebuchet ms": 1.16, "courier new": 1.13,
};
export const singleLine = (font: string | null): number => SINGLE_LINE[font?.toLowerCase() ?? ""] ?? 1.17;

/** Generic CSS family to fall back on when a document font isn't installed. */
export function genericFamily(font: string): string {
  if (/mono|courier|consolas|menlo/i.test(font)) return "monospace";
  return /garamond|times|cambria|georgia|palatino|book|baskerville|bodoni|didot|century|minion|serif|roman/i.test(font) && !/sans/i.test(font) ? "serif" : "sans-serif";
}

/** CSS line-height for `lines` of Word line spacing in `font`. */
export const lineHeight = (lines: number, font: string | null): number => Math.round(lines * singleLine(font) * 100) / 100;

export class DocxStyles {
  private styles = new Map<string, Element>();
  private defaultStyle: string | null = null;
  private cssCache = new Map<string, string>();

  /** Theme fonts that w:rFonts asciiTheme names refer to. */
  private themeFonts = { major: "", minor: "" };

  constructor(private stylesXml: Document | null, private numbering: Document | null, theme: Document | null = null) {
    const latin = (tag: string) => theme?.getElementsByTagNameNS(NS.a, tag)[0]?.getElementsByTagNameNS(NS.a, "latin")[0]?.getAttribute("typeface") ?? "";
    this.themeFonts = { major: latin("majorFont"), minor: latin("minorFont") };
    for (const s of Array.from(stylesXml?.getElementsByTagNameNS(NS.w, "style") ?? [])) {
      const id = wVal(s, "styleId");
      if (!id) continue;
      this.styles.set(id, s);
      if (wVal(s, "type") === "paragraph" && wVal(s, "default") === "1") this.defaultStyle = id;
    }
  }

  /** The Latin font a paragraph style resolves to (through the table style and defaults), or null. */
  fontFor(styleId: string | null, tableStyleId: string | null = null): string | null {
    const defaults = this.stylesXml?.getElementsByTagNameNS(NS.w, "docDefaults")[0];
    const sources = [...this.chain(styleId), ...(tableStyleId ? this.chain(tableStyleId) : []), defaults].filter(Boolean) as Element[];
    return sources.map((s) => this.fontOf(w(w(s, "rPr") ?? w(w(s, "rPrDefault"), "rPr"), "rFonts"))).find(Boolean) ?? null;
  }

  /** The font size (pt) a paragraph style resolves to, or null. */
  sizeFor(styleId: string | null, tableStyleId: string | null = null): number | null {
    const defaults = this.stylesXml?.getElementsByTagNameNS(NS.w, "docDefaults")[0];
    const sources = [...this.chain(styleId), ...(tableStyleId ? this.chain(tableStyleId) : []), defaults].filter(Boolean) as Element[];
    const sz = sources.map((s) => wVal(w(w(s, "rPr") ?? w(w(s, "rPrDefault"), "rPr"), "sz"))).find(Boolean);
    return sz ? Number(sz) / 2 : null;
  }

  /** A w:rFonts' Latin font, resolving theme references (asciiTheme="minorHAnsi" and the like). */
  private fontOf(rFonts: Element | null): string | null {
    const theme = wVal(rFonts, "asciiTheme");
    if (theme) return (theme.startsWith("major") ? this.themeFonts.major : this.themeFonts.minor) || null;
    return wVal(rFonts, "ascii");
  }

  /** The style and its basedOn ancestors, most specific first. */
  private chain(styleId: string | null): Element[] {
    const out: Element[] = [];
    for (let id = styleId ?? this.defaultStyle; id && out.length < 10; ) {
      const s = this.styles.get(id);
      if (!s) break;
      out.push(s);
      id = wVal(w(s, "basedOn"));
    }
    return out;
  }

  /** Display CSS for a paragraph style; inside a table, the table style's paragraph settings sit between the defaults and the paragraph style, as in Word. */
  paragraphCss(styleId: string | null, tableStyleId: string | null = null): string {
    const cacheKey = `${styleId}|${tableStyleId}`;
    const cached = this.cssCache.get(cacheKey);
    if (cached !== undefined) return cached;

    const defaults = this.stylesXml?.getElementsByTagNameNS(NS.w, "docDefaults")[0];
    const sources = [...this.chain(styleId), ...(tableStyleId ? this.chain(tableStyleId) : []), defaults].filter(Boolean) as Element[];
    const rPrs = sources.map((s) => w(s, "rPr") ?? w(w(s, "rPrDefault"), "rPr"));
    const pPrs = sources.map((s) => w(s, "pPr") ?? w(w(s, "pPrDefault"), "pPr"));
    const first = (props: (Element | null)[], local: string) => props.map((p) => w(p, local)).find(Boolean) ?? null;

    const css: string[] = [];
    const sz = wVal(first(rPrs, "sz"));
    if (sz) css.push(`font-size: ${Number(sz) / 2}pt`);
    const b = first(rPrs, "b");
    if (b) css.push(`font-weight: ${isOn(b) ? "bold" : "normal"}`);
    const i = first(rPrs, "i");
    if (i) css.push(`font-style: ${isOn(i) ? "italic" : "normal"}`);
    const color = wVal(first(rPrs, "color"));
    if (color && /^[0-9a-f]{6}$/i.test(color)) css.push(`color: #${color.toUpperCase()}`);
    const font = rPrs.map((r) => this.fontOf(w(r, "rFonts"))).find(Boolean) ?? null;
    if (font) css.push(`font-family: "${font}", ${genericFamily(font)}`);
    if (isOn(first(rPrs, "caps"))) css.push("text-transform: uppercase");

    const jc = wVal(first(pPrs, "jc"));
    if (jc) css.push(`text-align: ${jc === "both" ? "justify" : jc === "end" ? "right" : jc === "start" ? "left" : jc}`);
    const spacings = pPrs.map((p) => w(p, "spacing")).filter(Boolean) as Element[];
    const spacing = (attr: string) => spacings.map((s) => wVal(s, attr)).find((v) => v !== null) ?? null;
    const before = spacing("before");
    const after = spacing("after");
    if (before) css.push(`margin-top: ${Number(before) / 20}pt`);
    if (after) css.push(`margin-bottom: ${Number(after) / 20}pt`);
    const line = spacing("line");
    const rule = spacing("lineRule");
    if (line && rule === "exact") css.push(`line-height: ${Number(line) / 20}pt`);
    else css.push(`line-height: ${lineHeight(line && rule !== "atLeast" ? Number(line) / 240 : 1, font)}`);

    const out = css.join("; ");
    this.cssCache.set(cacheKey, out);
    return out;
  }

  /** numId / ilvl a paragraph gets from its style, if any. */
  private styleNumbering(styleId: string | null): { numId: string | null; ilvl: number | null } {
    for (const s of this.chain(styleId)) {
      const numPr = w(w(s, "pPr"), "numPr");
      if (numPr) return { numId: wVal(w(numPr, "numId")), ilvl: wVal(w(numPr, "ilvl")) === null ? null : Number(wVal(w(numPr, "ilvl"))) };
    }
    return { numId: null, ilvl: null };
  }

  private level(numId: string, ilvl: number): Element | null {
    return numberingLevel(this.numbering, numId, ilvl);
  }

  /** List markers for every list paragraph, keyed by the paragraph's document position. */
  listMarks(doc: PMNode): Map<number, ListMark> {
    const out = new Map<number, ListMark>();
    const counters = new Map<string, number[]>();

    doc.descendants((node, pos) => {
      if (node.type.name !== "paragraph") return true;
      const a = attrsOf<ParagraphAttrs>(node);
      const fromStyle = this.styleNumbering(a.style);
      const numId = a.numId ?? fromStyle.numId;
      if (!numId || numId === "0") return false;
      const ilvl = a.numId || a.ilvl ? a.ilvl : fromStyle.ilvl ?? 0;

      const lvl = this.level(numId, ilvl);
      const counts = counters.get(numId) ?? [];
      counters.set(numId, counts);
      const start = Number(wVal(w(lvl, "start")) ?? 1);
      counts[ilvl] = (counts[ilvl] ?? start - 1) + 1;
      counts.length = ilvl + 1;

      const fmt = wVal(w(lvl, "numFmt"));
      const lvlText = wVal(w(lvl, "lvlText")) ?? "";
      const text =
        !lvl || fmt === "bullet"
          ? BULLET_GLYPHS[lvlText] ?? lvlText
          : lvlText.replace(/%(\d)/g, (_, d) => {
              const k = Number(d) - 1;
              return formatNumber(counts[k] ?? 1, k === ilvl ? fmt : wVal(w(this.level(numId, k), "numFmt")));
            });
      const ind = w(w(lvl, "pPr"), "ind");
      out.set(pos, {
        text,
        indentPt: Number(wVal(ind, "left") ?? wVal(ind, "start") ?? 720 * (ilvl + 1)) / 20,
        hangingPt: Number(wVal(ind, "hanging") ?? 360) / 20,
      });
      return false;
    });
    return out;
  }
}
