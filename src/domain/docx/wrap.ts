/**
 * Text wrapping on both sides of a floating picture (Word's "both sides" square wrap, Google Docs' "Wrap text" with
 * the picture mid-column). Browsers only wrap text on one side of a float, so the page view lays the picture over
 * the text and inserts a blank as wide as the picture into each line it crosses: the text before the blank stays
 * left of the picture, the text after it continues on its right.
 */

export interface Box {
  left: number;
  right: number;
  top: number;
  bottom: number;
}

/** A paragraph as laid out, which the planner measures and inserts blanks into one at a time. */
export interface TextLayout {
  /** The paragraph's text, one entry per character. */
  text: string;
  /** Box of character `k` as laid out now (with the blanks inserted so far); null when it has no box. */
  box(k: number): Box | null;
  /** Insert a blank `width` wide before character `k`. */
  insert(k: number, width: number): void;
}

const space = (c: string | undefined) => c === undefined || /\s/.test(c);
const sameLine = (a: Box, b: Box) => a.top < b.bottom && b.top < a.bottom;

/**
 * Blanks that keep the text out of `areas` (pictures plus their distance from the text), in the paragraph's
 * coordinates. A word that would cross a picture's left edge moves to its right side, whole.
 */
export function wrapAround(layout: TextLayout, areas: Box | Box[], limit = 1000): { at: number; width: number; area: number }[] {
  const list = Array.isArray(areas) ? areas : [areas];
  const gaps: { at: number; width: number; area: number }[] = [];
  const n = layout.text.length;
  const bottom = Math.max(...list.map((a) => a.bottom));
  for (let k = 0, steps = 0; k < n && steps < limit; steps++) {
    const box = layout.box(k);
    if (!box || box.right <= box.left) {
      k++;
      continue;
    }
    if (box.top >= bottom) break;
    const hit = space(layout.text[k]) ? -1 : list.findIndex((a) => box.bottom > a.top && box.top < a.bottom && box.right > a.left && box.left < a.right);
    if (hit < 0) {
      k++;
      continue;
    }
    const area = list[hit];
    // Back to the start of the word, on this line.
    let start = k;
    while (start > 0 && !space(layout.text[start - 1])) {
      const prev = layout.box(start - 1);
      if (!prev || !sameLine(prev, box)) break;
      start--;
    }
    const at = layout.box(start)!;
    if (gaps.some((g) => g.at === start) && at.left >= area.left) {
      // Already moved once and still crossing: the word is wider than the room beside the picture; leave it.
      k = start + 1;
      continue;
    }
    const width = area.right - at.left;
    if (width <= 0) {
      k++;
      continue;
    }
    layout.insert(start, width);
    gaps.push({ at: start, width, area: hit });
    k = start;
  }
  return gaps;
}

/**
 * Justified text beside pictures: the browser would justify a line across its blanks, moving them off the
 * pictures. Instead each side is spread by its own word spacing: text left of a picture to end at its left edge
 * (its blank narrowing to match), the last piece to end at the margin (`right`), or a blank (`area` -1) fills
 * after a last piece of one word. A paragraph's last line keeps
 * its natural spacing, as justified text does. Returns the blanks' widths and the spacing for character ranges.
 */
export function justify(layout: TextLayout, gaps: { at: number; width: number; area: number }[], areas: Box[], right: number): { gaps: { at: number; width: number; area: number }[]; spacing: { from: number; to: number; px: number }[] } {
  const out = gaps.map((g) => ({ ...g }));
  const spacing: { from: number; to: number; px: number }[] = [];
  const n = layout.text.length;
  const spaces = (from: number, to: number) => {
    let count = 0;
    for (let k = from; k < to; k++) if (layout.text[k] === " ") count++;
    return count;
  };
  const done = new Set<number>();
  const blanks = [...out];
  for (const g of blanks) {
    if (done.has(g.at)) continue;
    const box = layout.box(g.at);
    if (!box) continue;
    // The line: from its first character to its last.
    let start = g.at, end = g.at;
    while (start > 0 && layout.box(start - 1) && sameLine(layout.box(start - 1)!, box)) start--;
    while (end + 1 < n && layout.box(end + 1) && sameLine(layout.box(end + 1)!, box)) end++;
    let from = start;
    for (const h of blanks.filter((x) => x.at >= start && x.at <= end).sort((a, b) => a.at - b.at)) {
      done.add(h.at);
      const gapStart = layout.box(h.at)!.left - h.width;
      const extra = areas[h.area].left - gapStart;
      const count = spaces(from, h.at);
      if (extra > 0.5 && count) {
        spacing.push({ from, to: h.at, px: extra / count });
        h.width -= extra;
      }
      from = h.at;
    }
    if (end + 1 >= n) continue; // the paragraph's last line
    let last = end;
    while (last > from && space(layout.text[last])) last--;
    const extra = right - layout.box(last)!.right - 0.5;
    const count = spaces(from, last);
    if (extra > 0.5 && count) spacing.push({ from, to: last, px: extra / count });
    // One word: nothing to spread, so a blank after it fills the line (else the browser would justify the whole line).
    else if (extra > 0.5) out.push({ at: last + 1, width: extra, area: -1 });
  }
  return { gaps: out, spacing };
}

/** Whether a "wrap text" picture leaves room for text on both of its sides (else it floats to one side). */
export function wrapsBothSides(x: number, width: number, column: number, align: string | null): boolean {
  if (align === "center") return width < column - 2 * MIN_SIDE;
  if (align) return false;
  return x > MIN_SIDE && column - x - width > MIN_SIDE;
}

/** Narrowest strip of text beside a picture worth wrapping into (0.5 in). */
export const MIN_SIDE = 48;
