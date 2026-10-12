import { describe, it, expect, afterEach } from "vitest";
import { EditorState } from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import { layoutPages, PageLayout, pageSetup, sectionSetups, TopBlock, Unit } from "../src/domain/docx/pages";
import { docxSchema as schema } from "../src/domain/docx/schema";
import { pagesKey, pagesPlugin, printPages, sectionWidthsPlugin } from "../src/presentation/docx/docx-pages";

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

describe("mixed page sizes", () => {
  it("gives each page the text height of the section that starts it", () => {
    // Portrait section (500 px of text), then a landscape one (300 px) after a section break.
    const tall = (h: number) => ({ ...block(h), avail: 500 });
    const wide = (h: number) => ({ ...block(h), avail: 300 });
    const l = lay([tall(400), { ...tall(10), breakAfter: true }, wide(200), wide(200), wide(200)]);
    expect(l.starts).toEqual([2, 3, 4]);
    expect(l.fills).toEqual([90, 100, 100, 100]);
  });

  it("splits a paragraph at the bottom of a shorter landscape page", () => {
    const l = lay([{ ...block(10, true), avail: 500 }, { ...para(20, 20), avail: 300 }]);
    expect(l.spacers[0].pos).toBe(1600);
  });

  it("reads each section's page size and margins, landscape sections included", () => {
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const portrait = `<w:sectPr xmlns:w="${W}"><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr>`;
    const landscape = `<w:sectPr xmlns:w="${W}"><w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="360" w:footer="360"/></w:sectPr>`;
    const body = new DOMParser().parseFromString(`<w:body xmlns:w="${W}">${portrait}</w:body>`, "application/xml").documentElement;
    const final = pageSetup(body);
    const doc = schema.nodes.doc.create({ sectPr: portrait }, [
      schema.nodes.paragraph.create(null, schema.text("a")),
      schema.nodes.paragraph.create({ sectPr: landscape }, schema.text("b")),
      schema.nodes.paragraph.create(null, schema.text("c")),
    ]);
    const setups = sectionSetups(doc, final);
    expect(setups.map((s) => [s.width, s.height, s.left])).toEqual([[1056, 816, 48], [1056, 816, 48], [816, 1056, 96]]);
    expect(sectionSetups(doc, final, "letter").every((s) => s === final)).toBe(true);
    // In the editor (as wide as the landscape page) each section's blocks get their own text width.
    const wide = { ...final, width: 1056 };
    const view = new EditorView(document.body.createDiv(), { state: EditorState.create({ doc, plugins: [sectionWidthsPlugin(() => wide, () => final, () => "document")] }) });
    const styles = Array.from(view.dom.children).map((p) => (p as HTMLElement).getAttribute("style") ?? "");
    // Text widths: landscape 1056 - 96 = 960, portrait 624, editor 1056 - 192 = 864.
    expect(styles[0]).toContain("translate: -48px 0");
    expect(styles[0]).toMatch(/margin-right: calc\((0pt - )?-?96px\)/);
    expect(styles[2]).toMatch(/margin-right: calc\((0pt - -)?240px\)/);
    expect(styles[2]).not.toContain("translate");
    view.destroy();
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

describe("layoutPages: columns", () => {
  const inCols = (b: TopBlock, count = 2, group = 1, extra: Partial<TopBlock> = {}): TopBlock => ({ ...b, columns: { count, group }, ...extra });
  it("continues a long table across pages of columns instead of moving the whole table", () => {
    const rows = Array.from({ length: 12 }, (_, i) => ({ height: 100, cantSplit: true, cells: [cell(96, [text(2, 50, i * 1000)])] }));
    const l = lay([block(100), inCols(table(rows)), block(100)]);
    expect(l.pages).toBe(2);
    expect(l.starts).toEqual([]);
    expect(l.spacers.map(s => s.pos)).toEqual([8000]);
    expect(l.fills).toEqual([200]);
  });
  it("splits an oversized row inside columns and draws one gap for its cells", () => {
    const l = lay([inCols(table([{ height: 1400, cells: [cell(96, [text(28, 50, 1000)]), cell(300, [text(28, 50, 5000)])] }]))]);
    expect(l.pages).toBe(2);
    expect(l.spacers).toHaveLength(2);
    expect(l.spacers.filter(s => s.gapAt !== null)).toHaveLength(1);
    expect(l.fills).toEqual([300]);
  });

  it("fits N column-heights of text on a page, and continues below the balanced columns", () => {
    // 100 used, so each of the 2 columns has 400: room for 800 of text.
    const l = lay([block(100), inCols(block(300)), inCols(block(300)), inCols(block(300)), block(100)]);
    expect(l.starts).toEqual([3]);
    // Page 1: 100 + (300 + 300) / 2 = 400 used. Page 2: 300 / 2 = 150 of columns, then the 100 block.
    expect(l.fills).toEqual([100, AVAIL - 250]);
  });

  it("moves tables and other blocks in columns whole", () => {
    const l = lay([inCols(block(400)), inCols(block(400)), inCols(block(400))]);
    expect(l.spacers).toEqual([]);
    expect(l.starts).toEqual([2]);
  });

  it("continues a paragraph that runs past a page of columns on the next page, with a spacer carrying the gap", () => {
    // 60 lines of 20 = 1200; a page of 2 columns holds 1000: lines 0-49 there, 50-59 on page 2.
    const l = lay([inCols(para(60, 20, 1000))]);
    expect(l.spacers).toEqual([{ pos: 1000 + 100 * 51, height: GAP, gapAt: 0, left: 96, page: 1 }]);
    expect(l.pages).toBe(2);
    // Page 2: 200 of text in 2 columns.
    expect(l.fills).toEqual([AVAIL - 100]);
  });

  it("keeps two lines of a paragraph in columns on each side of a page break", () => {
    // 950 already in the columns: a 6-line paragraph splits 2 + 4; a 3-line one moves whole.
    expect(lay([inCols(block(950)), inCols(para(6, 20, 1000))]).spacers).toHaveLength(1);
    const l = lay([inCols(block(950)), inCols(para(3, 20, 1000))]);
    expect(l.spacers).toEqual([]);
    expect(l.starts).toEqual([1]);
  });

  it("jumps to the next column after a column break", () => {
    // Column break after the first block: it takes a whole column (500), the next two fill column 2 and spill.
    const l = lay([inCols(block(100), 2, 1, { columnBreak: true }), inCols(block(300)), inCols(block(300))]);
    expect(l.starts).toEqual([2]);
  });

  it("balances forced columns by their tallest content rather than unused column capacity", () => {
    const l = lay([block(50), inCols(block(100), 2, 1, { columnBreak: true }), inCols(block(200)), block(50)]);
    expect(l.pages).toBe(1);
    expect(l.fills).toEqual([200]);
  });

  it("a break at a filled column still advances, and a break in the last column starts a page", () => {
    const l = lay([inCols(block(500), 2, 1, { columnBreak: true }), inCols(block(100), 2, 1, { columnBreak: true }), inCols(block(100))]);
    expect(l.pages).toBe(2);
    expect(l.starts).toEqual([2]);
    expect(l.fills).toEqual([0,450]);
  });

  it("starts a new set of columns after a page break inside them", () => {
    const l = lay([inCols(block(100), 2, 1, { breakAfter: true }), inCols(block(100))]);
    expect(l.starts).toEqual([1]);
    expect(l.fills).toEqual([0, AVAIL - 50]);
    expect(l.columnPads).toEqual([{pos:0,height:900}]);
  });

  it("treats two column sections in a row separately", () => {
    const l = lay([inCols(block(400), 2, 1), inCols(block(400), 3, 2)]);
    // 200 then 400 / 3.
    expect(l.starts).toEqual([]);
    expect(l.fills[0]).toBeCloseTo(AVAIL - 200 - 400 / 3);
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

  it("starts a new column page for a manual page break without rebalancing the page it ends", async () => {
    const first = schema.nodes.paragraph.create(null, [schema.text("n:5"), schema.nodes.page_break.create()]);
    const cols = schema.nodes.columns.create({ count: 2, space: 720 }, [first, p(5)]);
    const v = await mount([cols, p(5)]);
    expect(pagesKey.getState(v.state)!.pages).toBe(2);
    expect(v.dom.querySelector<HTMLElement>(".docx-column-pad")!.style.height).toBe("1628px");
    const state = pagesKey.getState(v.state)!;
    const gap = state.decorations.find().find(d => d.spec.key?.startsWith("gap-"))!;
    const widget = (gap as unknown as {type:{toDOM:()=>HTMLElement}}).type.toDOM();
    expect(widget.querySelector<HTMLElement>(".docx-page-fill")!.style.getPropertyValue("--docx-fill")).toBe("0px");
  });

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

describe("printPages", () => {
  const rect = (el: Element, top: number, bottom: number) => (el.getBoundingClientRect = () => ({ top, bottom }) as DOMRect);

  it("cuts the paged document into one box per page, each showing its own stretch", () => {
    const host = document.createElement("div");
    host.className = "docx-editor-host is-paged";
    host.innerHTML = `<div class="ProseMirror"><p>one</p><div class="docx-page-gap"><div class="docx-page-space"></div></div><p>two</p><div class="docx-page-spacer"><div class="docx-page-gap-float"><div class="docx-page-space"></div></div></div><p>three</p></div>`;
    const pm = host.firstElementChild!;
    rect(pm, 50, 3300);
    const [a, b] = Array.from(pm.querySelectorAll(".docx-page-space"));
    rect(a, 1110, 1134);
    rect(b, 2164, 2188);
    const pages = printPages(host, 1056);
    expect(pages).toHaveLength(3);
    const view = (p: HTMLElement) => p.querySelector<HTMLElement>(".dw-print-clip")!;
    const copy = (p: HTMLElement) => p.querySelector<HTMLElement>(".docx-editor-host")!;
    expect(pages.map((p) => copy(p).style.getPropertyValue("--dw-print-top"))).toEqual(["0px", "-1084px", "-2138px"]);
    // Each page shows up to the next gap, and no more than a page.
    expect(pages.map((p) => view(p).style.getPropertyValue("--dw-print-height"))).toEqual(["1056px", "1030px", "1056px"]);
    expect(copy(pages[1]).textContent).toContain("two");
  });

  it("is empty without an editor", () => {
    expect(printPages(document.createElement("div"), 1056)).toEqual([]);
  });
  it("prints at full size when the editor is scaled to fit beside review panels", () => {
    const host = document.createElement("div");
    host.className = "docx-editor-host is-paged";
    host.innerHTML = `<div class="ProseMirror" style="zoom: 0.5"><p>one</p><div class="docx-page-gap"><div class="docx-page-space"></div></div><p>two</p></div>`;
    rect(host.firstElementChild!, 50, 1100);
    rect(host.querySelector(".docx-page-space")!, 578, 590);
    const pages = printPages(host, 1056);
    expect(pages).toHaveLength(2);
    expect(pages[1].querySelector<HTMLElement>(".docx-editor-host")!.style.getPropertyValue("--dw-print-top")).toBe("-1080px");
    expect(pages[0].querySelector<HTMLElement>(".ProseMirror")!.style.zoom).toBe("1");
  });
});

describe("layoutPages: keeping to Word's pages", () => {
  it("lets a page run a little long instead of starting another (squeeze), with a negative fill", () => {
    const blocks = [block(300), para(12, 20)]; // 540 px on a 500 px page
    expect(lay(blocks).pages).toBe(2);
    const squeezed = layoutPages(blocks, AVAIL, GAP, { squeeze: 40 });
    expect(squeezed.pages).toBe(1);
    expect(squeezed.fills).toEqual([-40]);
    // Not more than allowed.
    expect(layoutPages(blocks, AVAIL, GAP, { squeeze: 30 }).pages).toBe(2);
  });

  it("starts a page where Word did when this page is nearly full", () => {
    const flagged = { ...para(1, 20, 50), wordBreak: true };
    // 420 of 500 px used: Word's break is followed though the paragraph would fit.
    const full = layoutPages([block(420), flagged, block(20)], AVAIL, GAP);
    expect(full.starts).toEqual([1]);
    expect(full.fills[0]).toBe(80);
  });

  it("ignores Word's break when the page is far from full (text was deleted since)", () => {
    const flagged = { ...para(1, 20, 50), wordBreak: true };
    expect(layoutPages([block(150), flagged], AVAIL, GAP).starts).toEqual([]);
  });

  it("lets a paragraph broken across pages run long by the squeeze too, the gap drawn below its last line", () => {
    const squeezed = layoutPages([block(300), para(20, 12)], AVAIL, GAP, { squeeze: 24 });
    // 300 + 240 = 540: 16 lines fit in 500 + 24, so the break comes after line 18 (300 + 18 * 12 = 516).
    expect(squeezed.spacers).toHaveLength(1);
    expect(squeezed.spacers[0].pos).toBe(100 * 19);
    expect(squeezed.spacers[0].gapAt).toBe(-16);
    expect(squeezed.spacers[0].height).toBe(GAP - 16);
  });
});

describe("layoutPages: no squeeze", () => {
  it("never draws a gap above the end of a page's text when pages may not run long", () => {
    // Only the spacing after a paragraph overflows: the break comes after it, the gap at the page's end.
    const r = lay([block(300), para(20, 12)]);
    expect(r.spacers.every((s) => s.gapAt === null || s.gapAt >= 0)).toBe(true);
    expect(r.fills.every((f) => f >= 0)).toBe(true);
  });
});

describe("layoutPages: spacing at the bottom of a page", () => {
  it("doesn't make the page taller when only a paragraph's space after runs past it (Word drops it there)", () => {
    // 400 + 4 lines (80) = 480 of text, then 30 px of space after: 10 px past the page.
    const r = lay([block(400), para(4, 20, 0, 30), block(50)]);
    expect(r.starts).toEqual([2]);
    // The gap moves up over that spacing, so the page keeps its height.
    expect(r.fills[0]).toBe(-10);
  });

  it("still pads a page normally when nothing runs past it", () => {
    expect(lay([block(400), para(4, 20, 0, 10), block(50)]).fills[0]).toBe(10);
  });
});
