import { describe, it, expect, afterEach } from "vitest";
import { EditorState } from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import { layoutPages, PageLayout, TopBlock, Unit } from "../src/domain/docx/pages";
import { docxSchema as schema } from "../src/domain/docx/schema";
import { pagesKey, pagesPlugin } from "../src/presentation/docx/docx-pages";

const GAP = 30;
const AVAIL = 500;

const block = (height: number, breakAfter = false, pos = 0): TopBlock => ({ unit: { kind: "block", height, pos }, breakAfter });
/** A paragraph of `n` lines of `lh` px; line k starts at text position pos + 100 * (k + 1). */
const text = (n: number, lh: number, pos = 0, extra = 0): Unit => ({
  kind: "text",
  height: n * lh + extra,
  pos,
  left: 96,
  lines: Array.from({ length: n }, (_, k) => ({ top: k * lh, bottom: (k + 1) * lh, pos: pos + 100 * (k + 1) })),
});
const para = (n: number, lh: number, pos = 0, extra = 0): TopBlock => ({ unit: text(n, lh, pos, extra), breakAfter: false });
const cell = (left: number, children: Unit[], padTop = 0, padBottom = 0) => ({ padTop, padBottom, left, children });
const table = (rows: { height: number; cantSplit?: boolean; cells: ReturnType<typeof cell>[] }[], pos = 0): TopBlock => ({
  unit: { kind: "table", height: rows.reduce((a, r) => a + r.height, 0), pos, rows },
  breakAfter: false,
});
const lay = (blocks: TopBlock[]): PageLayout => layoutPages(blocks, AVAIL, GAP);

describe("layoutPages: between blocks", () => {
  it("lays an empty document out as one empty page", () => {
    expect(lay([])).toEqual({ starts: [], pageNumbers: [1], fills: [AVAIL], spacers: [], pages: 1 });
  });

  it("keeps a block that exactly fills the page on it", () => {
    expect(lay([block(200), block(300)]).starts).toEqual([]);
    expect(lay([block(200), block(301)]).starts).toEqual([1]);
  });

  it("honours consecutive manual page breaks, each giving an empty-ish page", () => {
    const l = lay([block(10, true), block(10, true), block(10)]);
    expect(l.starts).toEqual([1, 2]);
    expect(l.pageNumbers).toEqual([1, 2, 3]);
    expect(l.fills).toEqual([490, 490, 490]);
  });

  it("ignores a manual break after the last block", () => {
    expect(lay([block(10), block(10, true)]).pages).toBe(1);
  });

  it("never reports negative blank space, even for an oversized block", () => {
    const l = lay([block(1200), block(10)]);
    expect(l.fills.every((f) => f >= 0)).toBe(true);
    expect(l.starts).toEqual([1]);
  });
});

describe("layoutPages: paragraphs", () => {
  it("splits a paragraph that runs over three pages twice", () => {
    const l = lay([para(30, 40, 1000)]); // 1200 px: 12 lines a page
    expect(l.spacers.map((s) => [s.pos, s.page])).toEqual([
      [1000 + 1300, 1],
      [1000 + 2500, 2],
    ]);
    expect(l.pages).toBe(3);
    expect(l.fills).toEqual([AVAIL - 6 * 40]);
  });

  it("moves a line back so two lines carry over (widow control)", () => {
    // 5 lines of 100 at 0: 5 would fit exactly, so use an offset: 4 fit, 1 would be alone on page 2.
    const l = lay([block(60), para(5, 100, 1000)]);
    // Only lines 1-3 stay; lines 4-5 go together.
    expect(l.spacers).toEqual([{ pos: 1400, height: AVAIL - 60 - 300 + GAP, gapAt: AVAIL - 60 - 300, left: 96, page: 1 }]);
  });

  it("moves the whole paragraph when fewer than two lines would stay (orphan control)", () => {
    expect(lay([block(450), para(6, 40)]).starts).toEqual([1]);
  });

  it("moves short paragraphs (under four lines) whole", () => {
    expect(lay([block(430), para(3, 40)]).starts).toEqual([1]);
  });

  it("does not split when only the spacing after the paragraph overflows", () => {
    const l = lay([block(100), para(10, 40, 0, 30)]); // 400 of lines + 30 spacing = 530 at 100
    expect(l.spacers).toEqual([]);
    expect(l.starts).toEqual([]);
  });

  it("places a line taller than a page without looping forever", () => {
    const l = lay([para(1, 1500), block(10)]);
    expect(l.pages).toBe(2);
    expect(l.starts).toEqual([1]);
  });

  it("splits a paragraph that starts a page and is taller than it", () => {
    const l = lay([para(20, 40, 0)]);
    expect(l.spacers).toHaveLength(1);
    expect(l.spacers[0].gapAt).toBe(AVAIL - 12 * 40);
  });

  it("numbers pages after an in-paragraph break correctly for later gaps", () => {
    const l = lay([para(20, 40), block(400)]); // page 2 has 8 lines (320), then 400 doesn't fit
    expect(l.starts).toEqual([1]);
    expect(l.pageNumbers).toEqual([1, 3]);
    expect(l.pages).toBe(3);
  });

  it("lays out thousands of paragraphs quickly", () => {
    const blocks = Array.from({ length: 3000 }, (_, i) => para(3 + (i % 5), 20, i * 1000));
    const t = performance.now();
    const l = lay(blocks);
    expect(performance.now() - t).toBeLessThan(200);
    expect(l.pages).toBeGreaterThan(100);
  });
});

describe("layoutPages: tables", () => {
  it("lets rows that fit stay, and splits the row that crosses the page bottom", () => {
    const l = lay([
      block(100),
      table([
        { height: 100, cells: [cell(100, [text(2, 50, 1000)])] },
        { height: 500, cells: [cell(100, [text(10, 50, 2000)]), cell(400, [text(2, 50, 3000)])] },
        { height: 100, cells: [cell(100, [text(2, 50, 4000)])] },
      ]),
    ]);
    // Row 2 starts at 200: 6 lines fit, 4 carry over.
    expect(l.spacers).toEqual([{ pos: 2700, height: GAP, gapAt: 0, left: 96, page: 1 }]);
    expect(l.pages).toBe(2);
    expect(l.fills).toEqual([AVAIL - 200 - 100]);
  });

  it("moves a table whose first row can't start on the page", () => {
    const l = lay([block(480), table([{ height: 100, cells: [cell(100, [text(2, 50, 1000)])] }])]);
    expect(l.starts).toEqual([1]);
    expect(l.spacers).toEqual([]);
  });

  it("splits a table at the top of a page whose first row is taller than the page", () => {
    const l = lay([table([{ height: 800, cells: [cell(100, [text(16, 50, 1000)])] }])]);
    expect(l.starts).toEqual([]);
    expect(l.spacers).toHaveLength(1);
    expect(l.pages).toBe(2);
  });

  it("counts cell padding above and below the content", () => {
    const l = lay([block(90), table([{ height: 20 + 450, cells: [cell(100, [text(9, 50, 1000)], 10, 10)] }])]);
    // Content starts at 100: 8 lines fit (500), line 9 alone would be a widow, so 7 + 2.
    expect(l.spacers.map((s) => s.pos)).toEqual([1800]);
  });

  it("breaks between two paragraphs in a cell when the second can't start", () => {
    const l = lay([table([{ height: 520, cells: [cell(100, [text(9, 50, 1000), text(2, 30, 5000)])] }])]);
    // First paragraph ends at 450; the 2-line one can't split or fit (450 + 60 > 500) and moves.
    expect(l.spacers).toEqual([{ pos: 5000, height: 50 + GAP, gapAt: 50, left: 100, page: 1 }]);
  });

  it("moves a table nested in a cell whole", () => {
    const nested: Unit = { kind: "block", height: 250, pos: 7000 };
    const l = lay([block(100), table([{ height: 500, cells: [cell(100, [text(4, 50, 1000), nested])] }])]);
    expect(l.spacers).toEqual([{ pos: 7000, height: 200 + GAP, gapAt: 200, left: 100, page: 1 }]);
  });

  it("draws the page gap once per break, even when several cells break there", () => {
    const l = lay([
      table([
        { height: 900, cells: [cell(100, [text(18, 50, 1000)]), cell(300, [text(18, 50, 3000)]), cell(500, [text(18, 50, 5000)])] },
      ]),
    ]);
    const byPage = (p: number) => l.spacers.filter((s) => s.page === p);
    expect(byPage(1)).toHaveLength(3);
    expect(byPage(1).filter((s) => s.gapAt !== null)).toHaveLength(1);
  });

  it("continues a row on the page its longest cell reaches", () => {
    const l = lay([
      table([{ height: 1200, cells: [cell(100, [text(24, 50, 1000)]), cell(400, [text(2, 50, 3000)])] }]),
      block(50),
    ]);
    expect(l.pages).toBe(3);
    // Page 3 holds the last 4 lines of the long cell, then the block.
    expect(l.fills).toEqual([AVAIL - 4 * 50 - 50]);
  });
});

describe("layoutPages: keeping things together (Word / Google Docs paragraph and row options)", () => {
  const keep = (b: TopBlock, opts: Partial<TopBlock>) => ({ ...b, ...opts });

  it("moves a heading kept with the next paragraph to the page that paragraph starts on", () => {
    // 440 used; the heading (30) fits, the 3-line paragraph after it doesn't, so both move.
    const l = lay([block(440), keep(block(30), { keepNext: true }), para(3, 20)]);
    expect(l.starts).toEqual([1]);
    expect(l.fills).toEqual([60, 410]);
  });

  it("moves a whole chain of kept paragraphs", () => {
    const l = lay([block(400), keep(block(30), { keepNext: true }), keep(block(30), { keepNext: true }), block(100)]);
    expect(l.starts).toEqual([1]);
  });

  it("ignores keep with next when the chain starts at the top of a page (it can't do better)", () => {
    const l = lay([keep(block(450), { keepNext: true }), block(100)]);
    expect(l.starts).toEqual([1]);
  });

  it("keeps the paragraph when the next one only splits (its first lines stay on the page)", () => {
    // 380 + 30 = 410 used; 10 lines of 20 split after 4 lines, which is enough.
    const l = lay([block(380), keep(block(30), { keepNext: true }), para(10, 20, 1000)]);
    expect(l.starts).toEqual([]);
    expect(l.spacers).toHaveLength(1);
  });

  it("does not move a heading kept with the next one across a manual page break", () => {
    const l = lay([block(440), keep(block(30), { keepNext: true, breakAfter: true }), block(100)]);
    expect(l.starts).toEqual([2]);
  });

  it("moves a paragraph whose lines are kept together instead of splitting it", () => {
    const l = lay([block(300), keep(para(10, 30, 1000), { keepLines: true })]);
    expect(l.spacers).toEqual([]);
    expect(l.starts).toEqual([1]);
  });

  it("still splits a kept-together paragraph that is taller than a page", () => {
    const l = lay([keep(para(30, 20, 1000), { keepLines: true })]);
    expect(l.pages).toBe(2);
  });

  it("moves a row that can't split to the next page whole", () => {
    // Row 1 fits (300); row 2 (300) would split, but can't, so a spacer opens each of its cells.
    const l = lay([
      table([
        { height: 300, cells: [cell(100, [text(15, 20, 1000)])] },
        { height: 300, cantSplit: true, cells: [cell(100, [text(15, 20, 2000)]), cell(400, [text(2, 20, 3000)])] },
      ]),
    ]);
    expect(l.spacers.map((x) => [x.pos, x.gapAt])).toEqual([
      [2000, 200],
      [3000, null],
    ]);
    expect(l.pages).toBe(2);
    expect(l.fills).toEqual([AVAIL - 300]);
  });

  it("moves a table whose first row can't split whole", () => {
    const l = lay([block(300), table([{ height: 300, cantSplit: true, cells: [cell(100, [text(15, 20, 1000)])] }])]);
    expect(l.starts).toEqual([1]);
    expect(l.spacers).toEqual([]);
  });
});

// -----------------------------------------------------------------------------
// The page view in an editor, with fake geometry (jsdom does no layout)
// -----------------------------------------------------------------------------

const LINE = 20;
const SETUP = { width: 816, height: 1056, top: 96, right: 96, bottom: 96, left: 96, header: 48, footer: 48 };
const TEXT_AREA = SETUP.height - SETUP.top - SETUP.bottom; // 864 px = 43.2 lines

/**
 * Give each paragraph `lines` lines of LINE px, stacked from y = 0, and make Range.getClientRects return
 * those lines. Paragraph line counts come from data in the text: "n:<lines>".
 */
function fakeLayout(view: EditorView) {
  const linesOf = (p: Element) => Number(/n:(\d+)/.exec(p.textContent ?? "")?.[1] ?? 1);
  const visible = (el: Element) => !el.closest(".docx-measuring .docx-page-spacer, .docx-measuring .docx-page-gap, .docx-measuring .docx-page-first, .docx-measuring .docx-page-last");
  const layout = () => {
    const rects = new Map<Element, DOMRect>();
    let y = 0;
    for (const p of Array.from(view.dom.querySelectorAll("p")).filter(visible)) {
      const h = linesOf(p) * LINE;
      rects.set(p, new DOMRect(96, y, 624, h));
      y += h;
    }
    return rects;
  };
  const ownRect = Element.prototype.getBoundingClientRect;
  Element.prototype.getBoundingClientRect = function (this: Element) {
    if (this === view.dom) return new DOMRect(0, 0, 816, 100000);
    return layout().get(this) ?? ownRect.call(this);
  };
  const ownRange = Range.prototype.getClientRects;
  Range.prototype.getClientRects = function (this: Range) {
    const p = this.startContainer as Element;
    const r = layout().get(p);
    if (!r) return [] as unknown as DOMRectList;
    return Array.from({ length: linesOf(p) }, (_, k) => new DOMRect(96, r.top + k * LINE, 600, LINE)) as unknown as DOMRectList;
  };
  // posAtCoords needs real layout; map a line's y back to the paragraph and line.
  view.posAtCoords = ({ top }) => {
    let pos: number | null = null;
    view.state.doc.descendants((node, at) => {
      if (pos !== null || node.type !== schema.nodes.paragraph) return pos === null;
      const dom = view.nodeDOM(at) as Element;
      const r = layout().get(dom);
      if (r && top >= r.top && top < r.bottom) pos = at + 1 + Math.floor((top - r.top) / LINE);
      return false;
    });
    return pos === null ? null : { pos, inside: -1 };
  };
  return () => {
    Element.prototype.getBoundingClientRect = ownRect;
    Range.prototype.getClientRects = ownRange;
  };
}

describe("page view in the editor", () => {
  let restore: (() => void) | null = null;
  let view: EditorView | null = null;
  afterEach(() => {
    restore?.();
    view?.destroy();
    document.body.empty();
  });

  const mount = async (blocks: ReturnType<typeof schema.nodes.paragraph.create>[], keep?: (node: ReturnType<typeof schema.nodes.paragraph.create>) => { next: boolean; lines: boolean }) => {
    const host = document.body.createDiv();
    view = new EditorView(host, {
      state: EditorState.create({
        doc: schema.nodes.doc.create(null, blocks),
        plugins: [pagesPlugin({ setup: () => SETUP, header: (p) => createSpan({ text: `H${p}` }), footer: (p) => createSpan({ text: `F${p}` }), keep })],
      }),
    });
    restore = fakeLayout(view);
    const { repaginate } = await import("../src/presentation/docx/docx-pages");
    repaginate(view);
    await new Promise((r) => setTimeout(r, 120));
    return view;
  };
  const p = (lines: number, label = "") => schema.nodes.paragraph.create(null, schema.text(`n:${lines} ${label}`.padEnd(lines + 6, "x")));

  it("puts a spacer with the page gap inside a paragraph that crosses the page bottom", async () => {
    const v = await mount([p(30, "a"), p(30, "b")]);
    // 60 lines of 20 px; 43 fit on page 1 -> break inside the second paragraph.
    const spacers = v.dom.querySelectorAll<HTMLElement>(".docx-page-spacer");
    expect(spacers).toHaveLength(1);
    expect(spacers[0].closest("p")!.textContent).toContain("n:30 b");
    const gap = spacers[0].querySelector(".docx-page-gap-float")!;
    expect(gap.textContent).toContain("F1");
    expect(gap.textContent).toContain("H2");
    expect(pagesKey.getState(v.state)!.pages).toBe(2);
  });

  it("removes the measuring class and keeps the old spacers from piling up on re-layout", async () => {
    const v = await mount([p(30), p(30)]);
    expect(v.dom.classList.contains("docx-measuring")).toBe(false);
    const { repaginate } = await import("../src/presentation/docx/docx-pages");
    repaginate(v);
    await new Promise((r) => setTimeout(r, 120));
    expect(v.dom.querySelectorAll(".docx-page-spacer")).toHaveLength(1);
  });

  it("uses a gap between blocks when a short paragraph moves whole", async () => {
    const v = await mount([p(42), p(3, "moves")]);
    expect(v.dom.querySelectorAll(".docx-page-spacer")).toHaveLength(0);
    const gap = v.dom.querySelector(".docx-page-gap")!;
    expect(gap.nextElementSibling!.textContent).toContain("moves");
  });

  it("keeps everything on one page when it fits", async () => {
    const v = await mount([p(10), p(10)]);
    expect(v.dom.querySelectorAll(".docx-page-spacer, .docx-page-gap")).toHaveLength(0);
    expect(pagesKey.getState(v.state)!.pages).toBe(1);
    expect(TEXT_AREA).toBeGreaterThan(400);
  });

  it("moves a heading kept with the next paragraph, using the style's setting when the paragraph has none", async () => {
    const heading = schema.nodes.paragraph.create({ style: "KeepStyle" }, schema.text("n:1 heading"));
    // (A plain paragraph style: heading styles render as h1, which the fake layout doesn't measure.)
    const v = await mount([p(42), heading, p(3)], (node) => ({ next: node.attrs.style === "KeepStyle", lines: false }));
    const gap = v.dom.querySelector(".docx-page-gap")!;
    expect(gap.nextElementSibling!.textContent).toContain("heading");
  });

  it("lets a paragraph's own keep with next override its style", async () => {
    const heading = schema.nodes.paragraph.create({ style: "KeepStyle", keepNext: false }, schema.text("n:1 heading"));
    const v = await mount([p(42), heading, p(3, "after")], (node) => ({ next: node.attrs.style === "KeepStyle", lines: false }));
    expect(v.dom.querySelector(".docx-page-gap")!.nextElementSibling!.textContent).toContain("after");
  });
});
