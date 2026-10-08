import { Node as PMNode } from "prosemirror-model";
import { NS, childNS } from "../../core/utils/xml";
import { wVal } from "./docx-reader";
import { attrsOf, ParagraphAttrs } from "./schema";

/** Marker drawn before a list paragraph. */
export interface ListMark {
  text: string;
  indentPt: number;
  hangingPt: number;
}

const w = (el: Element | null | undefined, local: string) => (el ? childNS(el, NS.w, local) : null);
const isOn = (el: Element | null) => !!el && !["0", "false", "off"].includes(wVal(el) ?? "");
/** Symbol / Wingdings private-use bullets and their Unicode look-alikes. */
const BULLET_GLYPHS: Record<string, string> = { "": "•", "": "▪", "": "➢", "": "✓", "": "❖", o: "◦", "": "•" };
const ROMAN: [number, string][] = [[1000, "m"], [900, "cm"], [500, "d"], [400, "cd"], [100, "c"], [90, "xc"], [50, "l"], [40, "xl"], [10, "x"], [9, "ix"], [5, "v"], [4, "iv"], [1, "i"]];

function roman(n: number): string {
  let out = "";
  for (const [v, s] of ROMAN) {
    while (n >= v) {
      out += s;
      n -= v;
    }
  }
  return out;
}

function letters(n: number): string {
  let out = "";
  for (; n > 0; n = Math.floor((n - 1) / 26)) out = String.fromCharCode(97 + ((n - 1) % 26)) + out;
  return out;
}

function formatNumber(n: number, fmt: string | null): string {
  switch (fmt) {
    case "lowerLetter":
      return letters(n);
    case "upperLetter":
      return letters(n).toUpperCase();
    case "lowerRoman":
      return roman(n);
    case "upperRoman":
      return roman(n).toUpperCase();
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
export class DocxStyles {
  private styles = new Map<string, Element>();
  private defaultStyle: string | null = null;
  private cssCache = new Map<string | null, string>();

  constructor(private stylesXml: Document | null, private numbering: Document | null) {
    for (const s of Array.from(stylesXml?.getElementsByTagNameNS(NS.w, "style") ?? [])) {
      const id = wVal(s, "styleId");
      if (!id) continue;
      this.styles.set(id, s);
      if (wVal(s, "type") === "paragraph" && wVal(s, "default") === "1") this.defaultStyle = id;
    }
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

  paragraphCss(styleId: string | null): string {
    const cached = this.cssCache.get(styleId);
    if (cached !== undefined) return cached;

    const defaults = this.stylesXml?.getElementsByTagNameNS(NS.w, "docDefaults")[0];
    const sources = [...this.chain(styleId), defaults].filter(Boolean) as Element[];
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
    const font = rPrs.map((r) => wVal(w(r, "rFonts"), "ascii")).find(Boolean);
    if (font) css.push(`font-family: "${font}"`);
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
    if (line && (!rule || rule === "auto")) css.push(`line-height: ${Math.round((Number(line) / 240) * 100) / 100}`);

    const out = css.join("; ");
    this.cssCache.set(styleId, out);
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
    const nums = Array.from(this.numbering?.getElementsByTagNameNS(NS.w, "num") ?? []);
    const absId = wVal(w(nums.find((n) => wVal(n, "numId") === numId), "abstractNumId"));
    const abs = Array.from(this.numbering?.getElementsByTagNameNS(NS.w, "abstractNum") ?? []).find((a) => wVal(a, "abstractNumId") === absId);
    return Array.from(abs?.getElementsByTagNameNS(NS.w, "lvl") ?? []).find((l) => wVal(l, "ilvl") === String(ilvl)) ?? null;
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
