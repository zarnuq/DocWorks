/*
 * Make text as wide as PowerPoint would draw it, whatever font is installed:
 * font-metrics.ts holds each Office font's advance widths, and the gap to the
 * fallback font actually used is spread over the run as letter-spacing.
 */
import { CHARS, FONT_ALIASES, FONT_WIDTHS, UNITS_PER_EM } from "./font-metrics";

/** Metric-compatible stand-ins for Office fonts, used when the real font isn't installed. */
const FALLBACK: Record<string, string> = {
  calibri: '"Carlito", "Liberation Sans", Arial',
  cambria: '"Caladea", "Liberation Serif", serif',
  arial: '"Liberation Sans", Arimo',
  "times new roman": '"Liberation Serif", Tinos, serif',
  "courier new": '"Liberation Mono", Cousine, monospace',
};

export function fontStack(font: string): string {
  const extra = FALLBACK[font.toLowerCase()];
  return `"${font}", ${extra ? extra + ", " : ""}sans-serif`;
}

const INDEX = new Map([...CHARS].map((ch, i) => [ch, i]));
const decoded = new Map<string, Uint8Array | null>();
/** Style order in FONT_WIDTHS is [regular, bold, italic, bold italic]; nearest available wins. */
const STYLE_PREFERENCE = [[0], [1, 0], [2, 0], [3, 1, 2, 0]];

function widthTable(font: string, b: boolean, i: boolean): Uint8Array | null {
  const name = font.toLowerCase();
  const styles = FONT_WIDTHS[FONT_ALIASES[name] ?? name];
  if (!styles) return null;
  const style = STYLE_PREFERENCE[(b ? 1 : 0) + (i ? 2 : 0)].find((s) => styles[s])!;
  const key = `${name}/${style}`;
  if (!decoded.has(key)) decoded.set(key, Uint8Array.from(atob(styles[style]), (c) => c.charCodeAt(0)));
  return decoded.get(key)!;
}

/** Width of `text` in em, from the table (characters it lacks count as 0); null if the font has no table. */
export function textWidthEm(text: string, font: string, b: boolean, i: boolean): number | null {
  const table = widthTable(font, b, i);
  if (!table) return null;
  let units = 0;
  for (const ch of text) units += table[INDEX.get(ch) ?? -1] ?? 0;
  return units / UNITS_PER_EM;
}

let canvas: { ctor: unknown; ctx: OffscreenCanvasRenderingContext2D | null } | null = null;
function context(): OffscreenCanvasRenderingContext2D | null {
  if (typeof OffscreenCanvas === "undefined") return null;
  if (canvas?.ctor !== OffscreenCanvas) canvas = { ctor: OffscreenCanvas, ctx: new OffscreenCanvas(1, 1).getContext("2d") };
  return canvas.ctx;
}

/** Letter-spacing (px) that makes `text` as wide as in `font`; null when the font or a canvas is unavailable. */
export function letterSpacing(text: string, run: { font: string; b: boolean; i: boolean }, sizePx: number): number | null {
  const table = widthTable(run.font, run.b, run.i);
  const ctx = table && context();
  if (!table || !ctx || !text) return null;
  ctx.font = `${run.i ? "italic " : ""}${run.b ? "bold " : ""}${sizePx}px ${fontStack(run.font)}`;
  let units = 0;
  let missing = "";
  for (const ch of text) {
    const w = table[INDEX.get(ch) ?? -1];
    if (w) units += w;
    else missing += ch;
  }
  const target = (units / UNITS_PER_EM) * sizePx + (missing ? ctx.measureText(missing).width : 0);
  return (target - ctx.measureText(text).width) / [...text].length;
}
