import { describe, expect, it } from "vitest";
import { Box, justify, TextLayout, wrapAround, wrapsBothSides } from "../src/domain/docx/wrap";

/**
 * A monospace paragraph: characters 10 px wide, lines 20 px tall in a column `width` px wide. Words and blanks
 * don't break; spaces at the end of a line hang (take no room on the next).
 */
function fake(text: string, width = 300): TextLayout & { lines(): string[] } {
  const gaps = new Map<number, number>();
  const layout = () => {
    const boxes: Box[] = [];
    let x = 0, y = 0;
    const items: { start: number; end: number; gap: number }[] = [];
    let k = 0;
    while (k < text.length) {
      const gap = gaps.get(k) ?? 0;
      let end = k + 1;
      if (text[k] !== " ") while (end < text.length && text[end] !== " " && !gaps.has(end)) end++;
      items.push({ start: k, end, gap });
      k = end;
    }
    for (const it of items) {
      const w = it.gap + (it.end - it.start) * 10;
      const isSpace = text[it.start] === " ";
      if (!isSpace && x > 0 && x + w > width) {
        x = 0;
        y += 20;
      }
      x += it.gap;
      for (let c = it.start; c < it.end; c++) {
        boxes[c] = { left: x, right: x + 10, top: y, bottom: y + 20 };
        x += 10;
      }
    }
    return boxes;
  };
  return {
    text,
    box: (k) => layout()[k] ?? null,
    insert: (k, w) => void gaps.set(k, w),
    lines() {
      const boxes = layout();
      const out: string[] = [];
      boxes.forEach((b, k) => {
        const line = b.top / 20;
        out[line] ??= "";
        const col = Math.round(b.left / 10);
        out[line] = out[line].padEnd(col, ".") + text[k];
      });
      return out.map((l) => l.trimEnd());
    },
  };
}

describe("text on both sides of a picture", () => {
  const words = "lorem ipsum dolor sit amet consectetur adipiscing elit sed do eiusmod tempor incididunt ut labore et dolore magna aliqua";

  it("leaves the picture's columns blank on the lines it covers, with text on both sides", () => {
    const t = fake(words);
    // Picture from x 100 to 200, over the second and third lines.
    const gaps = wrapAround(t, { left: 100, right: 200, top: 20, bottom: 60 });
    expect(gaps).toHaveLength(2);
    const lines = t.lines();
    for (const line of lines.slice(1, 3)) {
      expect(line.slice(10, 20).replace(/\./g, "").trim()).toBe("");
      expect(line.slice(0, 10).trim()).not.toBe("");
      expect(line.slice(20).trim()).not.toBe("");
    }
    // Lines above and below run across.
    expect(lines[0].length).toBeGreaterThan(20);
    expect(lines[3].slice(10, 20)).not.toMatch(/^\.+$/);
  });

  it("moves a word that would cross the picture's left edge to its right side whole", () => {
    // "bbbbbbb" would run from 90 to 160, over the picture's left edge at 100.
    const t = fake("aaaaaaaa bbbbbbb ccc dddd eeee");
    wrapAround(t, { left: 100, right: 200, top: 0, bottom: 20 });
    const first = t.lines()[0];
    expect(first.slice(0, 9)).toBe("aaaaaaaa ");
    expect(first.slice(9, 20)).toMatch(/^\.*$/);
    expect(first.slice(20, 27)).toBe("bbbbbbb");
  });

  it("does nothing to lines that end before the picture", () => {
    const t = fake("short");
    expect(wrapAround(t, { left: 100, right: 200, top: 0, bottom: 100 })).toEqual([]);
  });

  it("keeps going past a word too wide for either side instead of looping", () => {
    const t = fake("x ".repeat(5) + "y".repeat(25) + " z", 300);
    const gaps = wrapAround(t, { left: 100, right: 200, top: 0, bottom: 200 });
    expect(gaps.length).toBeLessThan(10);
  });

  it("wraps on both sides only when both sides have room", () => {
    expect(wrapsBothSides(200, 200, 624, null)).toBe(true);
    expect(wrapsBothSides(10, 200, 624, null)).toBe(false);
    expect(wrapsBothSides(400, 200, 624, null)).toBe(false);
    expect(wrapsBothSides(0, 200, 624, "center")).toBe(true);
    expect(wrapsBothSides(0, 600, 624, "center")).toBe(false);
    expect(wrapsBothSides(0, 200, 624, "left")).toBe(false);
  });
});

describe("two pictures beside the same lines", () => {
  it("keeps text out of both, with text between and around them", () => {
    const t = fake("lorem ipsum dolor sit amet consectetur adipiscing elit sed do eiusmod tempor incididunt ut labore et dolore magna aliqua ut enim ad minim veniam quis nostrud", 500);
    const gaps = wrapAround(t, [{ left: 100, right: 160, top: 0, bottom: 40 }, { left: 300, right: 360, top: 0, bottom: 40 }]);
    expect(gaps.map((g) => g.area).sort()).toEqual([0, 0, 1, 1]);
    for (const line of t.lines().slice(0, 2)) {
      expect(line.slice(10, 16)).toMatch(/^\.*$/);
      expect(line.slice(30, 36)).toMatch(/^\.*$/);
      expect(line.slice(16, 30).replace(/\./g, "")).not.toBe("");
    }
  });
});

describe("justified lines beside a picture", () => {
  it("spreads each side's words so the left side ends at the picture and the right side at the margin", () => {
    const t = fake("aa bb cc dd ee ff gg hh ii jj kk ll mm nn oo pp qq rr", 300);
    const area = { left: 125, right: 200, top: 0, bottom: 20 };
    const gaps = wrapAround(t, area);
    const { gaps: fitted, spacing } = justify(t, gaps, [area], 300);
    // First line: "aa bb cc dd " ends at 120, 5 px short of the picture: its 4 spaces widen by 1.25 px.
    const left = spacing.find((s) => s.from === 0)!;
    expect(left).toEqual({ from: 0, to: 12, px: 1.25 });
    const box = t.box(gaps[0].at)!;
    const spaces = t.text.slice(0, gaps[0].at).split(" ").length - 1;
    // The blank now starts where the left side ends (spread), and still ends at the picture's right edge.
    expect(box.left - fitted[0].width).toBeCloseTo(area.left);
    expect(left.px * spaces).toBeCloseTo(area.left - (box.left - gaps[0].width));
    const right = spacing.find((s) => s.from === gaps[0].at)!;
    expect(right.px).toBeGreaterThan(0);
  });

  it("fills out a line whose last piece is one word, so the browser doesn't spread the whole line", () => {
    // Right of the picture only "cccccccc" fits before the next word wraps.
    const t = fake("aa bb cccccccc dddddddddddddddd ee", 300);
    const area = { left: 50, right: 190, top: 0, bottom: 20 };
    const gaps = wrapAround(t, area);
    const { gaps: fitted } = justify(t, gaps, [area], 300);
    const filler = fitted.find((g) => g.area === -1)!;
    // "cccccccc" runs 190..270; 30 px are left to the margin, filled after the word.
    expect(filler).toEqual({ at: 14, width: 29.5, area: -1 });
  });

  it("leaves the paragraph's last line alone", () => {
    const t = fake("aa bb cc dd ee ff", 300);
    const area = { left: 50, right: 120, top: 0, bottom: 20 };
    const gaps = wrapAround(t, area);
    const { spacing } = justify(t, gaps, [area], 300);
    expect(spacing.every((s) => s.to <= gaps[0].at)).toBe(true);
  });
});
