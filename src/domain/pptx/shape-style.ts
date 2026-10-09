/*
 * Shape fill, outline and shadow: read from a:spPr with the p:style theme
 * references as fallback, convert to CSS, and write back in schema order.
 */
import { NS } from "../../core/utils/xml";
import { ColorContext, resolveColor } from "./color";

export interface GradientStop {
  /** 0..1 */
  pos: number;
  color: string;
}

export type Fill =
  | { type: "solid"; color: string }
  /** `angle`: DrawingML degrees, clockwise from left-to-right. */
  | { type: "gradient"; radial: boolean; angle: number; stops: GradientStop[] };

export type Size = "sm" | "med" | "lg";
export interface Arrow {
  type: string;
  w: Size;
  len: Size;
}

export interface Line {
  color: string;
  /** EMU */
  width: number;
  /** a:prstDash value ("solid", "dash", "sysDot"...) */
  dash: string;
  cap: "flat" | "rnd" | "sq";
  head: Arrow | null;
  tail: Arrow | null;
}

export interface Shadow {
  color: string;
  /** EMU */
  blur: number;
  dist: number;
  /** degrees, clockwise from right */
  dir: number;
}

/** The theme's style matrix (a:fmtScheme), indexed by p:style references. */
export interface ThemeStyles {
  fills: Element[];
  bgFills: Element[];
  lines: Element[];
  effects: Element[];
}

const kid = (el: Element | null | undefined, local: string) =>
  Array.from(el?.children ?? []).find((c) => c.namespaceURI === NS.a && c.localName === local) ?? null;
const kids = (el: Element | null | undefined) => Array.from(el?.children ?? []);
const num = (el: Element | null | undefined, attr: string) => {
  const v = el?.getAttribute(attr);
  return v === null || v === undefined || v === "" ? null : Number(v);
};

export function themeStyles(theme: Document | null): ThemeStyles {
  const list = (tag: string) => kids(theme?.getElementsByTagNameNS(NS.a, tag)[0]);
  return { fills: list("fillStyleLst"), bgFills: list("bgFillStyleLst"), lines: list("lnStyleLst"), effects: list("effectStyleLst") };
}

const FILL_TAGS = new Set(["noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"]);

/** A theme style entry and the color its phClr stands for, from a style reference (a:fillRef, p:bgRef...). */
function styleRef(ref: Element | null, list: (idx: number) => Element | undefined, ctx: ColorContext) {
  const idx = num(ref, "idx") ?? 0;
  if (!ref || !idx) return null;
  const entry = list(idx);
  return entry ? { entry, ctx: { ...ctx, phClr: resolveColor(ref, ctx) ?? undefined } } : null;
}

function fillOf(el: Element, ctx: ColorContext): Fill | null | undefined {
  switch (el.localName) {
    case "noFill":
      return null;
    case "solidFill": {
      const color = resolveColor(el, ctx);
      return color ? { type: "solid", color } : null;
    }
    case "pattFill": {
      const color = resolveColor(kid(el, "fgClr"), ctx);
      return color ? { type: "solid", color } : null;
    }
    case "gradFill": {
      const stops = Array.from(el.getElementsByTagNameNS(NS.a, "gs"))
        .map((gs) => ({ pos: (num(gs, "pos") ?? 0) / 100000, color: resolveColor(gs, ctx) ?? "" }))
        .filter((s) => s.color)
        .sort((a, b) => a.pos - b.pos);
      if (!stops.length) return null;
      return { type: "gradient", radial: !!kid(el, "path"), angle: (num(kid(el, "lin"), "ang") ?? 0) / 60000, stops };
    }
    default:
      return undefined;
  }
}

/** The fill of a shape; null = transparent. `groupFill` is used for a:grpFill. */
export function readFill(spPr: Element | null, styleEl: Element | null, ctx: ColorContext, theme: ThemeStyles | null, groupFill: Fill | null = null): Fill | null {
  const own = kids(spPr).find((c) => c.namespaceURI === NS.a && FILL_TAGS.has(c.localName));
  if (own) {
    if (own.localName === "grpFill") return groupFill;
    const f = fillOf(own, ctx);
    if (f !== undefined) return f;
  }
  return readFillRef(kid(styleEl, "fillRef"), ctx, theme);
}

/** Fill from a theme reference: a:fillRef (idx 1-999 = fill styles) or p:bgRef (1001+ = background fills). */
export function readFillRef(ref: Element | null, ctx: ColorContext, theme: ThemeStyles | null): Fill | null {
  const r = styleRef(ref, (i) => (i > 1000 ? theme?.bgFills[i - 1001] : theme?.fills[i - 1]), ctx);
  return r ? (fillOf(r.entry, r.ctx) ?? null) : null;
}

export function fillCss(fill: Fill | null, flipH = false, flipV = false): string | null {
  if (!fill) return null;
  if (fill.type === "solid") return fill.color;
  const stops = fill.stops.map((s) => `${s.color} ${Math.round(s.pos * 1000) / 10}%`).join(", ");
  if (fill.radial) return `radial-gradient(${stops})`;
  // DrawingML 0deg runs left to right, clockwise; CSS 0deg points up. Flips mirror the gradient.
  let ang = fill.angle;
  if (flipH) ang = 180 - ang;
  if (flipV) ang = -ang;
  return `linear-gradient(${(((ang + 90) % 360) + 360) % 360}deg, ${stops})`;
}

const arrow = (el: Element | null): Arrow | null => {
  const type = el?.getAttribute("type");
  if (!el || !type || type === "none") return null;
  return { type, w: (el.getAttribute("w") as Size) ?? "med", len: (el.getAttribute("len") as Size) ?? "med" };
};

/** The outline of a shape; null = none. Unset parts come from the theme line style. */
export function readLine(spPr: Element | null, styleEl: Element | null, ctx: ColorContext, theme: ThemeStyles | null): Line | null {
  const ln = kid(spPr, "ln");
  const ref = styleRef(kid(styleEl, "lnRef"), (i) => theme?.lines[i - 1], ctx);
  // Look in the shape's a:ln first, then the theme's.
  const layers: [Element, ColorContext][] = [];
  if (ln) layers.push([ln, ctx]);
  if (ref) layers.push([ref.entry, ref.ctx]);
  return lineOf(layers, ln);
}

/** A line from one a:ln-like element (a:lnL, a table style border...); null for none. */
export function lineFrom(ln: Element, ctx: ColorContext): Line | null {
  return lineOf([[ln, ctx]], ln);
}

function lineOf(layers: [Element, ColorContext][], ln: Element | null): Line | null {
  if (!layers.length) return null;

  let color: string | null | undefined;
  for (const [l, c] of layers) {
    const f = kids(l).find((k) => FILL_TAGS.has(k.localName));
    if (!f) continue;
    const fill = fillOf(f, c);
    color = fill === null ? null : fill?.type === "solid" ? fill.color : fill?.stops[0].color;
    break;
  }
  if (!color) return null;
  const pick = <T>(fn: (l: Element) => T | null) => {
    for (const [l] of layers) {
      const v = fn(l);
      if (v !== null) return v;
    }
    return null;
  };
  return {
    color,
    width: pick((l) => num(l, "w")) ?? 12700,
    dash: pick((l) => kid(l, "prstDash")?.getAttribute("val") ?? null) ?? "solid",
    cap: (pick((l) => l.getAttribute("cap")) as Line["cap"] | null) ?? "flat",
    head: ln ? arrow(kid(ln, "headEnd")) : null,
    tail: ln ? arrow(kid(ln, "tailEnd")) : null,
  };
}

/** The outer shadow; an explicit a:effectLst on the shape overrides the theme effect style. */
export function readShadow(spPr: Element | null, styleEl: Element | null, ctx: ColorContext, theme: ThemeStyles | null): Shadow | null {
  let list = kid(spPr, "effectLst");
  let c = ctx;
  if (!list && !kid(spPr, "effectDag")) {
    const ref = styleRef(kid(styleEl, "effectRef"), (i) => theme?.effects[i - 1], ctx);
    list = ref ? kid(ref.entry, "effectLst") : null;
    if (ref) c = ref.ctx;
  }
  const sh = kid(list, "outerShdw");
  const color = sh && resolveColor(sh, c);
  if (!sh || !color) return null;
  return { color, blur: num(sh, "blurRad") ?? 0, dist: num(sh, "dist") ?? 0, dir: (num(sh, "dir") ?? 0) / 60000 };
}

// -----------------------------------------------------------------------------
// Writing
// -----------------------------------------------------------------------------

const SPPR_ORDER = ["xfrm", "custGeom", "prstGeom", "fill", "ln", "effectLst", "effectDag", "scene3d", "sp3d", "extLst"];
const LN_ORDER = ["fill", "prstDash", "custDash", "round", "bevel", "miter", "headEnd", "tailEnd", "extLst"];
const EFFECT_ORDER = ["blur", "fillOverlay", "glow", "innerShdw", "outerShdw", "prstShdw", "reflection", "softEdge"];

const slot = (el: Element) => (FILL_TAGS.has(el.localName) ? "fill" : el.localName);

/** Insert `child` into `parent` before the first sibling that comes later in `order`. */
function place(parent: Element, child: Element, order: string[]): Element {
  const at = order.indexOf(slot(child));
  const next = kids(parent).find((c) => order.indexOf(slot(c)) > at) ?? null;
  return parent.insertBefore(child, next);
}

const a = (doc: Document, local: string, attrs: Record<string, string | number> = {}) => {
  const el = doc.createElementNS(NS.a, `a:${local}`);
  for (const [k, v] of Object.entries(attrs)) el.setAttribute(k, String(v));
  return el;
};

/** a:srgbClr for "#RRGGBB" / "#RRGGBBAA". */
function colorEl(doc: Document, color: string): Element {
  const el = a(doc, "srgbClr", { val: color.slice(1, 7).toUpperCase() });
  if (color.length === 9) el.appendChild(a(doc, "alpha", { val: Math.round((parseInt(color.slice(7, 9), 16) / 255) * 100000) }));
  return el;
}

function fillEl(doc: Document, fill: Fill | null): Element {
  if (!fill) return a(doc, "noFill");
  if (fill.type === "solid") {
    const el = a(doc, "solidFill");
    el.appendChild(colorEl(doc, fill.color));
    return el;
  }
  const el = a(doc, "gradFill", { rotWithShape: 1 });
  const gsLst = el.appendChild(a(doc, "gsLst"));
  for (const s of fill.stops) gsLst.appendChild(a(doc, "gs", { pos: Math.round(s.pos * 100000) })).appendChild(colorEl(doc, s.color));
  if (fill.radial) {
    el.appendChild(a(doc, "path", { path: "circle" })).appendChild(a(doc, "fillToRect", { l: 50000, t: 50000, r: 50000, b: 50000 }));
  } else {
    el.appendChild(a(doc, "lin", { ang: Math.round(fill.angle * 60000), scaled: 0 }));
  }
  return el;
}

/** Set the shape's fill (null = no fill), replacing whatever fill it had. */
/** Replace the fill of `spPr` (or another fill holder, with its child `order`); undefined removes it. */
export function writeFill(spPr: Element, fill: Fill | null | undefined, order = SPPR_ORDER): void {
  kids(spPr)
    .filter((c) => FILL_TAGS.has(c.localName))
    .forEach((c) => c.remove());
  if (fill !== undefined) place(spPr, fillEl(spPr.ownerDocument, fill), order);
}

/** Change some outline properties (null = no outline), keeping everything else on a:ln. */
export function writeLine(spPr: Element, patch: Partial<Line> | null): void {
  const doc = spPr.ownerDocument;
  const ln = kid(spPr, "ln") ?? place(spPr, a(doc, "ln"), SPPR_ORDER);
  const setChild = (local: string, el: Element | null) => {
    kids(ln)
      .filter((c) => slot(c) === local)
      .forEach((c) => c.remove());
    if (el) place(ln, el, LN_ORDER);
  };
  if (patch === null) {
    setChild("fill", a(doc, "noFill"));
    return;
  }
  if (patch.color !== undefined) setChild("fill", fillEl(doc, { type: "solid", color: patch.color }));
  if (patch.width !== undefined) ln.setAttribute("w", String(Math.round(patch.width)));
  if (patch.cap !== undefined) ln.setAttribute("cap", patch.cap);
  if (patch.dash !== undefined) {
    setChild("custDash", null);
    setChild("prstDash", a(doc, "prstDash", { val: patch.dash }));
  }
  for (const end of ["head", "tail"] as const) {
    if (patch[end] === undefined) continue;
    const v = patch[end];
    setChild(`${end}End`, v ? a(doc, `${end}End`, { type: v.type, w: v.w, len: v.len }) : null);
  }
}

/** Set or remove the outer shadow. Removing leaves an empty a:effectLst so the theme effect stays off. */
export function writeShadow(spPr: Element, shadow: Shadow | null): void {
  const doc = spPr.ownerDocument;
  const list = kid(spPr, "effectLst") ?? place(spPr, a(doc, "effectLst"), SPPR_ORDER);
  kid(list, "outerShdw")?.remove();
  if (!shadow) return;
  const sh = a(doc, "outerShdw", {
    blurRad: Math.round(shadow.blur),
    dist: Math.round(shadow.dist),
    dir: Math.round(shadow.dir * 60000),
    algn: "ctr",
    rotWithShape: 0,
  });
  sh.appendChild(colorEl(doc, shadow.color));
  place(list, sh, EFFECT_ORDER);
}
