/*
 * Picture adjustments, as Google Slides' format options: transparency (a:alphaModFix), brightness and contrast
 * (a:lum) and recolor (a:grayscl, or a:duotone between black / white and a theme accent, PowerPoint's "Accent N
 * Dark / Light"). Drawn with an SVG filter so the editor, thumbnails, slideshow and PDF all match.
 */
import { NS, childNS } from "../../core/utils/xml";
import { ColorContext, resolveColor } from "./color";
import type { PptxDocument } from "../../infrastructure/pptx/pptx-document";

/** "none", "gray", or "dark:<accentN>" / "light:<accentN>". */
export type Recolor = string;

export interface PictureAdjust {
  /** 0 (opaque) .. 1. */
  transparency: number;
  /** -1 .. 1. */
  brightness: number;
  /** -1 .. 1. */
  contrast: number;
  recolor: Recolor;
  /** Resolved duotone colors (dark, light), for drawing. */
  duotone?: [string, string] | null;
}

export const NO_ADJUST: PictureAdjust = { transparency: 0, brightness: 0, contrast: 0, recolor: "none", duotone: null };
export const ACCENTS = ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"];
const OURS = new Set(["alphaModFix", "lum", "grayscl", "duotone"]);
const pct = (el: Element | null, attr: string) => Number(el?.getAttribute(attr) ?? 0) / 100000;

export function blipOf(pic: Element): Element | null {
  return pic.getElementsByTagNameNS(NS.a, "blip")[0] ?? null;
}

export function readAdjust(pic: Element, colors: ColorContext): PictureAdjust {
  const blip = blipOf(pic);
  if (!blip) return { ...NO_ADJUST };
  const alpha = childNS(blip, NS.a, "alphaModFix");
  const lum = childNS(blip, NS.a, "lum");
  const duo = childNS(blip, NS.a, "duotone");
  let recolor = childNS(blip, NS.a, "grayscl") ? "gray" : "none";
  let duotone: [string, string] | null = null;
  if (duo) {
    const [a, b] = Array.from(duo.children);
    duotone = [resolveColor(a ?? null, colors) ?? "#000000", resolveColor(b ?? null, colors) ?? "#FFFFFF"];
    const scheme = (el: Element | undefined) => (el?.localName === "schemeClr" ? el.getAttribute("val") : null);
    recolor = a?.localName === "prstClr" && scheme(b) ? `dark:${scheme(b)}` : scheme(a) ? `light:${scheme(a)}` : "duotone";
  }
  return {
    transparency: alpha ? +(1 - (alpha.hasAttribute("amt") ? pct(alpha, "amt") : 1)).toFixed(5) : 0,
    brightness: pct(lum, "bright"),
    contrast: pct(lum, "contrast"),
    recolor,
    duotone,
  };
}

const x = (doc: Document, tag: string, attrs: Record<string, string> = {}, kids: Element[] = []) => {
  const el = doc.createElementNS(NS.a, `a:${tag}`);
  for (const [k, v] of Object.entries(attrs)) el.setAttribute(k, v);
  for (const k of kids) el.appendChild(k);
  return el;
};

/** Rewrite the picture's adjustments; other blip children (and an unknown duotone, when unchanged) stay. */
export function setPictureAdjust(pic: Element, adj: PictureAdjust): void {
  const blip = blipOf(pic);
  if (!blip) return;
  const doc = blip.ownerDocument;
  const keepDuotone = adj.recolor === "duotone";
  for (const c of Array.from(blip.children)) if (OURS.has(c.localName) && !(keepDuotone && c.localName === "duotone")) c.remove();
  const add: Element[] = [];
  const round = (v: number) => String(Math.round(v * 100000));
  if (adj.transparency > 0) add.push(x(doc, "alphaModFix", { amt: round(1 - adj.transparency) }));
  if (adj.recolor === "gray") add.push(x(doc, "grayscl"));
  const [kind, accent] = adj.recolor.split(":");
  if (accent && (kind === "dark" || kind === "light")) {
    const scheme = (mods: [string, string][]) => x(doc, "schemeClr", { val: accent }, mods.map(([t, v]) => x(doc, t, { val: v })));
    add.push(
      x(doc, "duotone", {}, kind === "dark"
        ? [x(doc, "prstClr", { val: "black" }), scheme([["tint", "45000"], ["satMod", "400000"]])]
        : [scheme([["shade", "45000"], ["satMod", "135000"]]), x(doc, "prstClr", { val: "white" })])
    );
  }
  if (adj.brightness || adj.contrast) {
    const attrs: Record<string, string> = {};
    if (adj.brightness) attrs.bright = round(adj.brightness);
    if (adj.contrast) attrs.contrast = round(adj.contrast);
    add.push(x(doc, "lum", attrs));
  }
  const extLst = childNS(blip, NS.a, "extLst");
  for (const el of add) blip.insertBefore(el, extLst);
}

const n = (v: number) => String(+v.toFixed(3));

/** SVG filter primitives drawing `adj` (opacity aside), or null when the picture is drawn as is. */
export function pictureFilter(adj: PictureAdjust): string | null {
  const out: string[] = [];
  if (adj.recolor !== "none") out.push('<feColorMatrix type="saturate" values="0"/>');
  if (adj.brightness || adj.contrast) {
    const c = Math.max(-0.99, Math.min(0.99, adj.contrast));
    const slope = c >= 0 ? 1 / (1 - c) : 1 + c;
    const intercept = adj.brightness + 0.5 * (1 - slope);
    const f = (ch: string) => `<feFunc${ch} type="linear" slope="${n(slope)}" intercept="${n(intercept)}"/>`;
    out.push(`<feComponentTransfer>${f("R")}${f("G")}${f("B")}</feComponentTransfer>`);
  }
  if (adj.duotone) {
    const [a, b] = adj.duotone.map((h) => [1, 3, 5].map((i) => parseInt(h.slice(i, i + 2), 16) / 255));
    const f = (ch: string, i: number) => `<feFunc${ch} type="table" tableValues="${n(a[i])} ${n(b[i])}"/>`;
    out.push(`<feComponentTransfer>${f("R", 0)}${f("G", 1)}${f("B", 2)}</feComponentTransfer>`);
  }
  return out.length ? out.join("") : null;
}

/** Replace a picture's image, keeping its size, crop and adjustments (an SVG version of the old image goes). */
export function replacePicture(doc: PptxDocument, part: string, pic: Element, bytes: Uint8Array, ext: string): void {
  const blip = blipOf(pic);
  if (!blip) return;
  blip.setAttributeNS(NS.r, "r:embed", doc.addImage(part, bytes, ext));
  // Office 2016+ keeps an SVG version in a:extLst, which would still be drawn instead of the new image.
  for (const ext of Array.from(blip.getElementsByTagNameNS(NS.a, "ext"))) if (ext.getElementsByTagNameNS("*", "svgBlip").length || ext.getAttribute("uri") === "{96DAC541-7B7A-43D3-8B79-37D633B846F1}") ext.remove();
  const list = childNS(blip, NS.a, "extLst");
  if (list && !list.children.length) list.remove();
  doc.touch(part);
}
