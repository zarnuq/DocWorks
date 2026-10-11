/*
 * Approximate Office text widths when the requested font is unavailable:
 * font-metrics.ts holds each Office font's advance widths, and the gap to the
 * fallback font actually used is spread over the run as letter-spacing.
 */
import { CHARS, FONT_ALIASES, FONT_WIDTHS, UNITS_PER_EM } from "./font-metrics";
import { fontStack } from "../../core/utils/fonts";
import { KERN_CHARS, KERN_UNITS, KERNING } from "./font-kerning";

export { fontStack };

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
let metricsRevision = 0;
/** Font registration changes invalidate both canvas font resolution and document fit caches. */
export function invalidateTextMetrics(): void {
  canvas = null;
  metricsRevision++;
}
export function textMetricsRevision(): number { return metricsRevision; }
function context(): OffscreenCanvasRenderingContext2D | null {
  if (typeof OffscreenCanvas === "undefined") return null;
  if (canvas?.ctor !== OffscreenCanvas) canvas = { ctor: OffscreenCanvas, ctx: new OffscreenCanvas(1, 1).getContext("2d") };
  return canvas.ctx;
}

const KERN_INDEX = new Map([...KERN_CHARS].map((ch, i) => [ch, i]));
const kernTables = new Map<string, Map<number, number> | null>();

/** Pair adjustments of a font ([first index * 256 + second] -> 1/KERN_UNITS em), null if it has none. */
function kernTable(font: string, b: boolean): Map<number, number> | null {
  const name = font.toLowerCase();
  const key = `${name}/${b}`;
  if (!kernTables.has(key)) {
    const data = KERNING[FONT_ALIASES[name] ?? name]?.[b ? 1 : 0];
    let table: Map<number, number> | null = null;
    if (data) {
      const bytes = Uint8Array.from(atob(data), (c) => c.charCodeAt(0));
      table = new Map();
      for (let k = 0; k + 2 < bytes.length; k += 3) table.set(bytes[k] * 256 + bytes[k + 1], (bytes[k + 2] << 24) >> 24);
    }
    kernTables.set(key, table);
  }
  return kernTables.get(key)!;
}

/** Total kerning of `text` in `font`, in em (negative: the pairs close up). Italic uses the upright pairs. */
export function kerningEm(text: string, font: string, b: boolean): number {
  const table = kernTable(font, b);
  if (!table) return 0;
  const chars = [...text];
  let units = 0;
  for (let k = 1; k < chars.length; k++) {
    const x = KERN_INDEX.get(chars[k - 1]);
    const y = KERN_INDEX.get(chars[k]);
    if (x !== undefined && y !== undefined) units += table.get(x * 256 + y) ?? 0;
  }
  return units / KERN_UNITS;
}

/**
 * Letter-spacing (px) that makes `text` as wide as in `font` (with its kerning when `kern`); null when the font or
 * a canvas is unavailable.
 */
export function letterSpacing(text: string, run: { font: string; b: boolean; i: boolean; kern?: boolean }, sizePx: number): number | null {
  const table = widthTable(run.font, run.b, run.i);
  const ctx = table && context();
  if (!table || !ctx || !text) return null;
  // Probe with two different generic fallbacks: equal widths mean the requested family resolved.
  // Preserve that font's native shaping instead of overriding it with rounded reference widths.
  const style = `${run.i ? "italic " : ""}${run.b ? "bold " : ""}72px`;
  const family = JSON.stringify(run.font);
  const probe = "mmmmmmmmiiiiiiiiWWWWAVATAR0123456789";
  ctx.font = `${style} ${family}, monospace`;
  const mono = ctx.measureText(probe).width;
  ctx.font = `${style} ${family}, serif`;
  if (mono === ctx.measureText(probe).width) return 0;
  ctx.font = `${run.i ? "italic " : ""}${run.b ? "bold " : ""}${sizePx}px ${fontStack(run.font)}`;
  let units = 0;
  let missing = "";
  for (const ch of text) {
    const w = table[INDEX.get(ch) ?? -1];
    if (w) units += w;
    else missing += ch;
  }
  const kerning = run.kern ? kerningEm(text, run.font, run.b) * sizePx : 0;
  const target = (units / UNITS_PER_EM) * sizePx + kerning + (missing ? ctx.measureText(missing).width : 0);
  return (target - ctx.measureText(text).width) / [...text].length;
}
