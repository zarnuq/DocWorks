/*
 * styles.xml: cell formats (cellXfs) as CSS for the grid, and new formats derived from an existing one for edits.
 * Records are shared by many cells, so an edit never changes one in place: it appends (or reuses) a record.
 */
import { fontStack } from "../../core/utils/fonts";
import { hex, hslToRgb, rgbToHsl } from "../pptx/color";
import { NS_S } from "../../infrastructure/xlsx/xlsx-workbook";
import { builtinFormat } from "./number-format";

const A = "http://schemas.openxmlformats.org/drawingml/2006/main";

// Legacy 64-color palette for indexed="n".
const INDEXED = (
  "000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF 000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF " +
  "800000 008000 000080 808000 800080 008080 C0C0C0 808080 9999FF 993366 FFFFCC CCFFFF 660066 FF8080 0066CC CCCCFF " +
  "000080 FF00FF FFFF00 00FFFF 800080 800000 008080 0000FF 00CCFF CCFFFF CCFFCC FFFF99 99CCFF FF99CC CC99FF FFCC99 " +
  "3366FF 33CCCC 99CC00 FFCC00 FF9900 FF6600 666699 969696 003366 339966 003300 333300 993300 993366 333399 333333"
).split(" ");
// Spreadsheet theme indexes list light 1 before dark 1, unlike the theme's own order.
const THEME_SLOTS = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];

/** Theme color scheme slot -> "#RRGGBB". */
export function themeColors(theme: Document | null): Record<string, string> {
  const out: Record<string, string> = {};
  const scheme = theme?.getElementsByTagNameNS(A, "clrScheme")[0];
  for (const slot of Array.from(scheme?.children ?? [])) {
    const c = slot.firstElementChild;
    const v = c?.getAttribute("lastClr") ?? c?.getAttribute("val");
    if (v && /^[0-9a-f]{6}$/i.test(v)) out[slot.localName] = `#${v.toUpperCase()}`;
  }
  return out;
}

export type BorderStyle = "thin" | "medium" | "thick" | "dashed" | "dotted" | "double" | null;
type Side = "left" | "right" | "top" | "bottom";
const SIDES: Side[] = ["left", "right", "top", "bottom"];

export interface StyleChange {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  fontName?: string;
  fontSize?: number;
  /** "#RRGGBB"; null for automatic. */
  color?: string | null;
  /** "#RRGGBB"; null for none. */
  fill?: string | null;
  numFmt?: string;
  hAlign?: "left" | "center" | "right" | "general";
  vAlign?: "top" | "center" | "bottom";
  wrap?: boolean;
  borders?: Partial<Record<Side, BorderStyle>>;
}

export interface Format {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  fontName: string;
  fontSize: number;
  color: string | null;
  fill: string | null;
  hAlign: string;
  vAlign: string;
  wrap: boolean;
  numFmt: string;
}

/** What a conditional format changes about a cell. */
export interface DxfStyle {
  fill?: string;
  color?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
}

/** CT_Stylesheet child order. */
const SHEET_ORDER = ["numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "cellStyles", "dxfs", "tableStyles", "colors", "extLst"];

export interface Xf {
  el: Element;
  numFmtId: number;
  fontId: number;
  fillId: number;
  borderId: number;
}

const FONT_ORDER = ["b", "i", "strike", "condense", "extend", "outline", "shadow", "u", "vertAlign", "sz", "color", "name", "family", "charset", "scheme"];
const BORDER_CSS: Record<string, string> = {
  thin: "1px solid", hair: "1px solid", medium: "2px solid", thick: "3px solid", dashed: "1px dashed", mediumDashed: "2px dashed",
  dotted: "1px dotted", double: "3px double", dashDot: "1px dashed", mediumDashDot: "2px dashed", dashDotDot: "1px dotted", slantDashDot: "2px dashed",
};

export class Styles {
  private cssCache = new Map<number, Record<string, string>>();

  constructor(
    readonly doc: Document,
    readonly theme: Record<string, string>
  ) {}

  // Records are only ever appended, never changed in place, so lookups can be cached; appending clears `lists`.
  private lists = new Map<string, Element[]>();
  private formats = new Map<number, Format>();
  private codes = new Map<number, string>();

  private list(tag: string): Element[] {
    let out = this.lists.get(tag);
    if (!out) {
      const parent = this.doc.getElementsByTagNameNS(NS_S, tag)[0];
      out = parent ? Array.from(parent.children).filter((c) => c.namespaceURI === NS_S) : [];
      this.lists.set(tag, out);
    }
    return out;
  }

  private child(el: Element | undefined, tag: string): Element | null {
    return Array.from(el?.children ?? []).find((c) => c.localName === tag) ?? null;
  }

  xf(index: number): Xf {
    const el = this.list("cellXfs")[index] ?? this.list("cellXfs")[0];
    const n = (a: string) => Number(el?.getAttribute(a) ?? 0);
    return { el, numFmtId: n("numFmtId"), fontId: n("fontId"), fillId: n("fillId"), borderId: n("borderId") };
  }

  numFmt(index: number): string {
    let code = this.codes.get(index);
    if (code === undefined) {
      const id = this.xf(index).numFmtId;
      const custom = this.list("numFmts").find((f) => Number(f.getAttribute("numFmtId")) === id);
      code = custom?.getAttribute("formatCode") ?? builtinFormat(id);
      this.codes.set(index, code);
    }
    return code;
  }

  color(el: Element | null): string | null {
    if (!el || el.getAttribute("auto") === "1") return null;
    let rgb: string | null = null;
    const argb = el.getAttribute("rgb");
    if (argb && /^[0-9a-f]{6,8}$/i.test(argb)) rgb = `#${argb.slice(-6).toUpperCase()}`;
    else if (el.hasAttribute("theme")) rgb = this.theme[THEME_SLOTS[Number(el.getAttribute("theme"))]] ?? null;
    else if (el.hasAttribute("indexed")) {
      const i = Number(el.getAttribute("indexed"));
      rgb = i < INDEXED.length ? `#${INDEXED[i]}` : null; // 64 = system foreground, 65 = background
    }
    const tint = Number(el.getAttribute("tint") ?? 0);
    if (!rgb || !tint) return rgb;
    const [h, s, l] = rgbToHsl([1, 3, 5].map((i) => parseInt(rgb.slice(i, i + 2), 16)));
    return hex(hslToRgb([h, s, tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint]));
  }

  private fillColor(fill: Element | undefined): string | null {
    const pattern = this.child(fill, "patternFill");
    if (pattern && pattern.getAttribute("patternType") === "solid") return this.color(this.child(pattern, "fgColor")) ?? "#000000";
    if (pattern?.getAttribute("patternType") && pattern.getAttribute("patternType") !== "none") return this.color(this.child(pattern, "fgColor"));
    const stop = this.child(fill, "gradientFill")?.getElementsByTagNameNS(NS_S, "color")[0];
    return stop ? this.color(stop) : null;
  }

  format(index: number): Format {
    const hit = this.formats.get(index);
    if (hit) return hit;
    const f = this.readFormat(index);
    this.formats.set(index, f);
    return f;
  }

  private readFormat(index: number): Format {
    const xf = this.xf(index);
    const font = this.list("fonts")[xf.fontId];
    const flag = (tag: string) => {
      const el = this.child(font, tag);
      return !!el && el.getAttribute("val") !== "0" && el.getAttribute("val") !== "false" && el.getAttribute("val") !== "none";
    };
    const align = this.child(xf.el, "alignment");
    return {
      bold: flag("b"),
      italic: flag("i"),
      underline: flag("u"),
      strike: flag("strike"),
      fontName: this.child(font, "name")?.getAttribute("val") ?? "Calibri",
      fontSize: Number(this.child(font, "sz")?.getAttribute("val") ?? 11),
      color: this.color(this.child(font, "color")),
      fill: this.fillColor(this.list("fills")[xf.fillId]),
      hAlign: align?.getAttribute("horizontal") ?? "general",
      vAlign: align?.getAttribute("vertical") ?? "bottom",
      wrap: align?.getAttribute("wrapText") === "1" || align?.getAttribute("wrapText") === "true",
      numFmt: this.numFmt(index),
    };
  }

  css(index: number): Record<string, string> {
    const hit = this.cssCache.get(index);
    if (hit) return hit;
    const f = this.format(index);
    const css: Record<string, string> = { "font-family": fontStack(f.fontName), "font-size": `${f.fontSize}pt` };
    if (f.bold) css["font-weight"] = "bold";
    if (f.italic) css["font-style"] = "italic";
    const deco = [f.underline && "underline", f.strike && "line-through"].filter(Boolean).join(" ");
    if (deco) css["text-decoration"] = deco;
    if (f.color) css.color = f.color;
    if (f.fill) css["background-color"] = f.fill;
    if (f.hAlign !== "general") css["text-align"] = f.hAlign === "centerContinuous" ? "center" : f.hAlign === "fill" || f.hAlign === "distributed" ? "justify" : f.hAlign;
    css["align-items"] = f.vAlign === "top" ? "flex-start" : f.vAlign === "center" ? "center" : "flex-end";
    if (f.wrap) css["white-space"] = "pre-wrap";
    const border = this.list("borders")[this.xf(index).borderId];
    for (const side of SIDES) {
      const el = this.child(border, side);
      const style = el?.getAttribute("style");
      if (style && BORDER_CSS[style]) css[`border-${side}`] = `${BORDER_CSS[style]} ${this.color(this.child(el ?? undefined, "color")) ?? "#000000"}`;
    }
    const indent = Number(this.child(this.xf(index).el, "alignment")?.getAttribute("indent") ?? 0);
    if (indent) css[f.hAlign === "right" ? "padding-right" : "padding-left"] = `${indent * 9 + 3}px`;
    this.cssCache.set(index, css);
    return css;
  }

  // ---------------------------------------------------------------------------
  // Deriving formats
  // ---------------------------------------------------------------------------

  private container(tag: string): Element {
    let el = this.doc.getElementsByTagNameNS(NS_S, tag)[0];
    if (!el) {
      // numFmts comes first; dxfs (often missing) after cellStyles.
      el = this.doc.createElementNS(NS_S, tag);
      const root = this.doc.documentElement;
      const rank = SHEET_ORDER.indexOf(tag);
      root.insertBefore(el, Array.from(root.children).find((c) => SHEET_ORDER.indexOf(c.localName) > rank) ?? null);
    }
    return el;
  }

  /** A differential format (a conditional format's style), as far as the grid draws one. */
  dxf(index: number): DxfStyle {
    const el = this.list("dxfs")[index];
    const out: DxfStyle = {};
    if (!el) return out;
    const font = this.child(el, "font");
    const flag = (tag: string) => {
      const f = this.child(font ?? undefined, tag);
      return f ? f.getAttribute("val") !== "0" && f.getAttribute("val") !== "false" && f.getAttribute("val") !== "none" : undefined;
    };
    for (const [key, tag] of [["bold", "b"], ["italic", "i"], ["underline", "u"], ["strike", "strike"]] as const) {
      const v = flag(tag);
      if (v !== undefined) out[key] = v;
    }
    const color = this.color(this.child(font ?? undefined, "color"));
    if (color) out.color = color;
    // A dxf's solid fill is its bgColor (fgColor in some writers).
    const pattern = this.child(this.child(el, "fill") ?? undefined, "patternFill");
    const fill = pattern && pattern.getAttribute("patternType") !== "none" ? this.color(this.child(pattern, "bgColor")) ?? this.color(this.child(pattern, "fgColor")) : null;
    if (fill) out.fill = fill;
    return out;
  }

  /** Append a differential format; returns its index. */
  addDxf(style: DxfStyle): number {
    const dxf = this.el("dxf");
    const font = this.el("font");
    for (const [key, tag] of [["bold", "b"], ["italic", "i"], ["strike", "strike"], ["underline", "u"]] as const) if (style[key]) this.setFontChild(font, tag, this.el(tag));
    if (style.color) this.setFontChild(font, "color", this.el("color", { rgb: `FF${style.color.slice(1).toUpperCase()}` }));
    if (font.children.length) dxf.appendChild(font);
    if (style.fill) dxf.appendChild(this.el("fill")).appendChild(this.el("patternFill", { patternType: "solid" })).appendChild(this.el("bgColor", { rgb: `FF${style.fill.slice(1).toUpperCase()}` }));
    const parent = this.container("dxfs");
    const index = this.list("dxfs").length;
    parent.appendChild(dxf);
    this.lists.delete("dxfs");
    parent.setAttribute("count", String(index + 1));
    return index;
  }

  /** Index of an equal record in `tag`, or of `el` appended to it. */
  private intern(tag: string, el: Element): number {
    const key = new XMLSerializer().serializeToString(el);
    const list = this.list(tag);
    const found = list.findIndex((e) => new XMLSerializer().serializeToString(e) === key);
    if (found >= 0) return found;
    const parent = this.container(tag);
    parent.appendChild(el);
    this.lists.delete(tag);
    parent.setAttribute("count", String(list.length + 1));
    return list.length;
  }

  private el(tag: string, attrs: Record<string, string> = {}): Element {
    const e = this.doc.createElementNS(NS_S, tag);
    for (const [k, v] of Object.entries(attrs)) e.setAttribute(k, v);
    return e;
  }

  private setFontChild(font: Element, tag: string, el: Element | null): void {
    this.child(font, tag)?.remove();
    if (!el) return;
    const rank = FONT_ORDER.indexOf(tag);
    const next = Array.from(font.children).find((c) => FONT_ORDER.indexOf(c.localName) > rank) ?? null;
    font.insertBefore(el, next);
  }

  private numFmtId(code: string): number {
    for (let id = 0; id < 50; id++) if (builtinFormat(id) === code && (id === 0 || code !== "General")) return id;
    const custom = this.list("numFmts").find((f) => f.getAttribute("formatCode") === code);
    if (custom) return Number(custom.getAttribute("numFmtId"));
    const id = Math.max(163, ...this.list("numFmts").map((f) => Number(f.getAttribute("numFmtId")))) + 1;
    const parent = this.container("numFmts");
    parent.appendChild(this.el("numFmt", { numFmtId: String(id), formatCode: code }));
    this.lists.delete("numFmts");
    parent.setAttribute("count", String(this.list("numFmts").length));
    return id;
  }

  /** The cell format `index` with `change` applied: an existing equal record, or a new one. */
  derive(index: number, change: StyleChange): number {
    const base = this.xf(index);
    const xf = base.el.cloneNode(true) as Element;
    xf.setAttribute("xfId", xf.getAttribute("xfId") ?? "0");

    const fontKeys: (keyof StyleChange)[] = ["bold", "italic", "underline", "strike", "fontName", "fontSize", "color"];
    if (fontKeys.some((k) => k in change)) {
      const font = (this.list("fonts")[base.fontId]?.cloneNode(true) as Element | undefined) ?? this.el("font");
      const flag = (tag: string, on: boolean | undefined) => on !== undefined && this.setFontChild(font, tag, on ? this.el(tag) : null);
      flag("b", change.bold);
      flag("i", change.italic);
      flag("strike", change.strike);
      flag("u", change.underline);
      if (change.fontSize !== undefined) this.setFontChild(font, "sz", this.el("sz", { val: String(change.fontSize) }));
      if ("color" in change) this.setFontChild(font, "color", change.color ? this.el("color", { rgb: `FF${change.color.slice(1).toUpperCase()}` }) : null);
      if (change.fontName !== undefined) {
        this.setFontChild(font, "name", this.el("name", { val: change.fontName }));
        this.setFontChild(font, "scheme", null); // a theme font scheme would override the name
        this.setFontChild(font, "family", null);
      }
      xf.setAttribute("fontId", String(this.intern("fonts", font)));
      xf.setAttribute("applyFont", "1");
    }
    if ("fill" in change) {
      const fill = this.el("fill");
      const pattern = fill.appendChild(this.el("patternFill", { patternType: change.fill ? "solid" : "none" }));
      if (change.fill) {
        pattern.appendChild(this.el("fgColor", { rgb: `FF${change.fill.slice(1).toUpperCase()}` }));
        pattern.appendChild(this.el("bgColor", { indexed: "64" }));
      }
      xf.setAttribute("fillId", String(change.fill ? this.intern("fills", fill) : 0));
      xf.setAttribute("applyFill", "1");
    }
    if (change.borders) {
      const border = (this.list("borders")[base.borderId]?.cloneNode(true) as Element | undefined) ?? this.el("border");
      for (const side of [...SIDES, "diagonal"]) if (!this.child(border, side)) border.appendChild(this.el(side));
      // Schema order: left, right, top, bottom, diagonal.
      for (const side of [...SIDES, "diagonal"]) border.appendChild(this.child(border, side)!);
      for (const [side, style] of Object.entries(change.borders)) {
        const el = this.el(side);
        if (style) {
          el.setAttribute("style", style);
          el.appendChild(this.el("color", { indexed: "64" }));
        }
        border.replaceChild(el, this.child(border, side)!);
      }
      xf.setAttribute("borderId", String(this.intern("borders", border)));
      xf.setAttribute("applyBorder", "1");
    }
    if (change.numFmt !== undefined) {
      xf.setAttribute("numFmtId", String(this.numFmtId(change.numFmt)));
      xf.setAttribute("applyNumberFormat", "1");
    }
    if (change.hAlign !== undefined || change.vAlign !== undefined || change.wrap !== undefined) {
      let align = this.child(xf, "alignment");
      if (!align) align = xf.insertBefore(this.el("alignment"), xf.firstElementChild);
      const set = (attr: string, v: string | undefined, dflt: string) => {
        if (v === undefined) return;
        if (v === dflt) align.removeAttribute(attr);
        else align.setAttribute(attr, v);
      };
      set("horizontal", change.hAlign, "general");
      set("vertical", change.vAlign, "bottom");
      set("wrapText", change.wrap === undefined ? undefined : change.wrap ? "1" : "0", "0");
      if (!align.attributes.length) align.remove();
      xf.setAttribute("applyAlignment", "1");
    }
    return this.intern("cellXfs", xf);
  }
}
