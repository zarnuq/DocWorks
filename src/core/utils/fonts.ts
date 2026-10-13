/** CSS font stacks for document fonts, shared by the slide and Word views (and their text measuring). */

/** Metric-compatible stand-ins for Office fonts, used when the real font isn't installed. */
const FALLBACK: Record<string, string> = {
  calibri: '"Carlito", "Liberation Sans", Arial',
  cambria: '"Caladea", "Liberation Serif"',
  arial: '"Liberation Sans", Arimo',
  "times new roman": '"Liberation Serif", Tinos',
  "courier new": '"Liberation Mono", Cousine',
  // Both ballot boxes (☐ ☑) from one font, never an emoji font's ☑.
  "segoe ui symbol": '"Noto Sans Symbols 2", "DejaVu Sans"',
};

/** Generic CSS family to fall back on when a document font isn't installed. */
export function genericFamily(font: string): string {
  if (/mono|courier|consolas|menlo/i.test(font)) return "monospace";
  return /garamond|times|cambria|georgia|palatino|book|baskerville|bodoni|didot|century|minion|serif|roman/i.test(font) && !/sans/i.test(font) ? "serif" : "sans-serif";
}

export function fontStack(font: string): string {
  const extra = FALLBACK[font.toLowerCase()];
  return `"${font}", ${extra ? extra + ", " : ""}${genericFamily(font)}`;
}
