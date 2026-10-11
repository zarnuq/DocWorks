export interface ColorContext {
  /** Theme color scheme: dk1, lt1, accent1... -> "#RRGGBB" */
  scheme: Record<string, string>;
  /** Master clrMap: tx1/bg1/tx2/bg2 -> scheme slot */
  map: Record<string, string>;
  /** Placeholder color (phClr) inside theme style entries: the p:style reference's color. */
  phClr?: string;
}

const PRESET: Record<string, string> = {
  black: "#000000",
  white: "#FFFFFF",
  red: "#FF0000",
  green: "#008000",
  blue: "#0000FF",
  yellow: "#FFFF00",
  gray: "#808080",
};

const COLOR_TAGS = new Set(["srgbClr", "schemeClr", "sysClr", "prstClr", "scrgbClr", "hslClr"]);

export const hex = (rgb: number[]) =>
  "#" + rgb.map((c) => Math.round(Math.max(0, Math.min(255, c))).toString(16).padStart(2, "0")).join("").toUpperCase();

export function rgbToHsl([r, g, b]: number[]): [number, number, number] {
  r /= 255;
  g /= 255;
  b /= 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  const h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
  return [h / 6, s, l];
}

export function hslToRgb([h, s, l]: number[]): number[] {
  if (s === 0) return [l * 255, l * 255, l * 255];
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const f = (t: number) => {
    t = (t + 1) % 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  return [f(h + 1 / 3) * 255, f(h) * 255, f(h - 1 / 3) * 255];
}

function baseColor(el: Element, ctx: ColorContext): string | null {
  const val = el.getAttribute("val") ?? "";
  switch (el.localName) {
    case "srgbClr":
      return /^[0-9a-f]{6}$/i.test(val) ? `#${val.toUpperCase()}` : null;
    case "sysClr":
      return el.getAttribute("lastClr") ? `#${el.getAttribute("lastClr")!.toUpperCase()}` : null;
    case "prstClr":
      return PRESET[val] ?? null;
    case "schemeClr":
      if (val === "phClr") return ctx.phClr?.slice(0, 7) ?? null;
      return ctx.scheme[ctx.map[val] ?? val] ?? null;
    default:
      return null;
  }
}

/**
 * Resolve the first DrawingML color element inside `container` (a:solidFill,
 * a:fontRef, ...) to "#RRGGBB" ("#RRGGBBAA" when translucent), applying
 * lumMod/lumOff/tint/shade/alpha.
 */
export function resolveColor(container: Element | null, ctx: ColorContext): string | null {
  const el = container && (COLOR_TAGS.has(container.localName) ? container : Array.from(container.children).find((c) => COLOR_TAGS.has(c.localName)));
  if (!el) return null;
  const base = baseColor(el, ctx);
  if (!base) return null;

  let rgb = [1, 3, 5].map((i) => parseInt(base.slice(i, i + 2), 16));
  let alpha = 1;
  for (const mod of Array.from(el.children)) {
    const v = Number(mod.getAttribute("val")) / 100000;
    if (mod.localName === "lumMod" || mod.localName === "lumOff") {
      const [h, s, l] = rgbToHsl(rgb);
      rgb = hslToRgb([h, s, mod.localName === "lumMod" ? l * v : Math.min(1, l + v)]);
    } else if (mod.localName === "tint") {
      rgb = rgb.map((c) => c * v + 255 * (1 - v));
    } else if (mod.localName === "shade") {
      rgb = rgb.map((c) => c * v);
    } else if (mod.localName === "alpha") {
      alpha = v;
    }
  }
  return alpha < 1 ? hex([...rgb, alpha * 255]) : hex(rgb);
}
