/**
 * Word-style page view for the docx editor: pages of the document's size with headers, footers and gaps
 * between them. Page breaks fall between blocks, between the lines of a paragraph and inside table rows;
 * a break inside a paragraph or cell is a blank spacer that carries the drawn gap across the whole page.
 * No Obsidian imports, so it runs under jsdom tests.
 */
import { Node as PMNode } from "prosemirror-model";
import { Plugin, PluginKey } from "prosemirror-state";
import { Decoration, DecorationSet, EditorView } from "prosemirror-view";
import { Cell, layoutPages, Line, PageLayout, PageSetup, Row, Spacer, TopBlock, Unit } from "../../domain/docx/pages";
import { docxSchema as schema } from "../../domain/docx/schema";

export interface PagesOptions {
  /** Current page geometry, or null for the continuous (pageless) view. */
  setup(): PageSetup | null;
  /** Header / footer content for a 1-based page number out of `total`. */
  header(page: number, total: number): Node | null;
  /** Keep with next / keep lines together for a paragraph, with its style applied; default: its own settings only. */
  keep?(node: PMNode): { next: boolean; lines: boolean };
  footer(page: number, total: number): Node | null;
}

interface PagesState {
  pages: number;
  decorations: DecorationSet;
  /** Header / footer band heights the layout assumed. */
  bands: [number, number];
}

export const pagesKey = new PluginKey<PagesState>("docx-pages");
const schedulers = new WeakMap<EditorView, () => void>();
/** Height of the grey space between two pages (.docx-page-space). */
const PAGE_SPACE = 24;

/** Re-measure the pages (after a resize or a header / footer change). */
export function repaginate(view: EditorView): void {
  schedulers.get(view)?.();
}

function hasPageBreak(node: PMNode): boolean {
  let found = !!node.attrs.sectionBreak;
  node.descendants((n) => {
    if (n.type === schema.nodes.page_break) found = true;
    return !found;
  });
  return found;
}

/** Heights of the drawn header and footer bands: a header or footer taller than its margin pushes the text in, as in Word. */
function bandHeights(view: EditorView, setup: PageSetup): [number, number] {
  const head = view.dom.querySelector<HTMLElement>(".docx-page-first")?.offsetHeight ?? 0;
  const foot = view.dom.querySelector<HTMLElement>(".docx-page-last .docx-page-foot")?.offsetHeight ?? 0;
  return [Math.max(setup.top, head), Math.max(setup.bottom, foot)];
}

// -----------------------------------------------------------------------------
// Measuring (with the page decorations hidden, so positions are the plain text flow)
// -----------------------------------------------------------------------------

const px = (v: string) => parseFloat(v) || 0;

/** Left edge of an element's content box, relative to the editor. */
function contentLeft(el: HTMLElement, origin: DOMRect): number {
  const cs = getComputedStyle(el);
  return el.getBoundingClientRect().left + px(cs.paddingLeft) + px(cs.borderLeftWidth) - origin.left;
}

/** A paragraph's lines, relative to its top. Each line's start position is looked up only if a break lands on it. */
function linesOf(view: EditorView, el: HTMLElement, from: number, to: number): Line[] {
  const range = document.createRange();
  range.selectNodeContents(el);
  const top = el.getBoundingClientRect().top;
  const lines: { top: number; bottom: number; left: number }[] = [];
  for (const r of Array.from(range.getClientRects())) {
    if (!r.height) continue;
    const line = lines.find((l) => Math.abs(l.top - r.top) < 2 || (r.top < l.bottom - 2 && r.bottom > l.top + 2));
    if (line) {
      line.top = Math.min(line.top, r.top);
      line.bottom = Math.max(line.bottom, r.bottom);
      line.left = Math.min(line.left, r.left);
    } else lines.push({ top: r.top, bottom: r.bottom, left: r.left });
  }
  lines.sort((a, b) => a.top - b.top);
  return lines.map((l) => ({
    top: l.top - top,
    bottom: l.bottom - top,
    get pos() {
      const hit = view.posAtCoords({ left: l.left + 1, top: (l.top + l.bottom) / 2 })?.pos ?? from;
      return Math.min(Math.max(hit, from), to);
    },
  }));
}

function measureUnit(view: EditorView, node: PMNode, pos: number, el: HTMLElement, height: number, origin: DOMRect): Unit {
  if (node.type === schema.nodes.paragraph) {
    return { kind: "text", height, pos, left: contentLeft(el, origin), lines: linesOf(view, el, pos + 1, pos + node.nodeSize - 1) };
  }
  if (node.type === schema.nodes.table) return { kind: "table", height, pos, rows: measureRows(view, node, pos, origin) };
  return { kind: "block", height, pos };
}

function measureRows(view: EditorView, table: PMNode, tablePos: number, origin: DOMRect): Row[] {
  const rows: Row[] = [];
  table.forEach((row, rowOffset) => {
    const rowPos = tablePos + 1 + rowOffset;
    const tr = view.nodeDOM(rowPos);
    if (!(tr instanceof HTMLElement)) return;
    const cells: Cell[] = [];
    row.forEach((cell, cellOffset) => {
      const cellPos = rowPos + 1 + cellOffset;
      const td = view.nodeDOM(cellPos);
      if (!(td instanceof HTMLElement) || td.offsetParent === null) return; // covered by a vertical merge
      const box = td.getBoundingClientRect();
      const cs = getComputedStyle(td);
      const kids: { node: PMNode; pos: number; el: HTMLElement }[] = [];
      cell.forEach((child, childOffset) => {
        const el = view.nodeDOM(cellPos + 1 + childOffset);
        if (el instanceof HTMLElement) kids.push({ node: child, pos: cellPos + 1 + childOffset, el });
      });
      if (!kids.length) return;
      const tops = kids.map((k) => k.el.getBoundingClientRect().top);
      const contentBottom = box.bottom - px(cs.paddingBottom) - px(cs.borderBottomWidth);
      const children = kids.map((k, i) => {
        const height = (i + 1 < kids.length ? tops[i + 1] : contentBottom) - tops[i];
        const unit = measureUnit(view, k.node, k.pos, k.el, height, origin);
        // A table inside a cell moves whole.
        return unit.kind === "table" ? { kind: "block" as const, height, pos: k.pos } : unit;
      });
      cells.push({ padTop: tops[0] - box.top, padBottom: box.bottom - contentBottom, left: contentLeft(td, origin), children });
    });
    rows.push({ height: tr.getBoundingClientRect().height, cantSplit: !!row.attrs.cantSplit, cells });
  });
  return rows;
}

/** The blocks pages break between: top-level blocks, with content controls opened up. */
function flowBlocks(doc: PMNode): { node: PMNode; pos: number }[] {
  const out: { node: PMNode; pos: number }[] = [];
  const walk = (parent: PMNode, start: number) =>
    parent.forEach((node, offset) => {
      if (node.type === schema.nodes.sdt) walk(node, start + offset + 1);
      else out.push({ node, pos: start + offset });
    });
  walk(doc, 0);
  return out;
}

function measure(view: EditorView, opts: PagesOptions): TopBlock[] {
  const origin = view.dom.getBoundingClientRect();
  const items: { node: PMNode; pos: number; el: HTMLElement | null; top: number }[] = [];
  flowBlocks(view.state.doc).forEach(({ node, pos }) => {
    const el = view.nodeDOM(pos);
    const dom = el instanceof HTMLElement ? el : null;
    items.push({ node, pos, el: dom, top: dom ? dom.getBoundingClientRect().top : NaN });
  });
  return items.map((it, i) => {
    let height = 0;
    if (it.el) {
      const next = items.slice(i + 1).find((n) => n.el);
      height = next ? next.top - it.top : it.el.getBoundingClientRect().height + px(getComputedStyle(it.el).marginBottom);
    }
    const unit = it.el ? measureUnit(view, it.node, it.pos, it.el, height, origin) : { kind: "block" as const, height, pos: it.pos };
    if (it.node.type !== schema.nodes.paragraph) return { unit, breakAfter: hasPageBreak(it.node) };
    // The paragraph's own setting wins over its style's.
    const style = opts.keep?.(it.node) ?? { next: false, lines: false };
    const a = it.node.attrs as { keepNext: boolean | null; keepLines: boolean | null };
    return { unit, breakAfter: hasPageBreak(it.node), keepNext: a.keepNext ?? style.next, keepLines: a.keepLines ?? style.lines };
  });
}

// -----------------------------------------------------------------------------
// Drawing
// -----------------------------------------------------------------------------

/** Bumped on every layout so header / footer widgets are rebuilt with fresh content. */
let generation = 0;

function pageDecorations(view: EditorView, opts: PagesOptions): PagesState {
  const gen = ++generation;
  const setup = opts.setup();
  const doc = view.state.doc;
  if (!setup) return { pages: 1, decorations: DecorationSet.empty, bands: [0, 0] };
  const bands = bandHeights(view, setup);

  view.dom.classList.add("docx-measuring");
  let layout: PageLayout;
  try {
    layout = layoutPages(measure(view, opts), setup.height - bands[0] - bands[1], bands[1] + PAGE_SPACE + bands[0]);
  } finally {
    view.dom.classList.remove("docx-measuring");
  }
  const { starts, pageNumbers, fills, spacers, pages } = layout;
  const offsets = flowBlocks(doc).map((b) => b.pos);

  const part = (cls: string, content: Node | null) => {
    const el = createDiv({ cls, attr: { contenteditable: "false" } });
    if (content) el.appendChild(content);
    return el;
  };
  /** Footer of `page`, the grey space, header of the next page. */
  const gapParts = (gap: HTMLElement, page: number) => {
    gap.appendChild(part("docx-page-foot", opts.footer(page, pages)));
    gap.createDiv({ cls: "docx-page-space" });
    gap.appendChild(part("docx-page-head", opts.header(page + 1, pages)));
  };
  // Blank space that pads a page to its full height.
  const fill = (k: number) => {
    const el = createDiv({ cls: "docx-page-fill" });
    el.setCssProps({ "--docx-fill": `${fills[k]}px` });
    return el;
  };
  const spacer = (s: Spacer) => {
    const el = createDiv({ cls: "docx-page-spacer", attr: { contenteditable: "false" } });
    el.setCssProps({ "--docx-spacer": `${s.height}px` });
    if (s.gapAt !== null) {
      const gap = el.createDiv({ cls: "docx-page-gap docx-page-gap-float" });
      gap.setCssProps({ "--docx-gap-top": `${s.gapAt}px`, "--docx-gap-left": `${-s.left}px` });
      gapParts(gap, s.page);
    }
    return el;
  };

  const decos = [
    Decoration.widget(0, () => part("docx-page-first", opts.header(1, pages)), { side: -1, key: `first-${gen}`, ignoreSelection: true }),
    ...starts.map((i, k) =>
      Decoration.widget(
        offsets[i],
        () => {
          const gap = createDiv({ cls: "docx-page-gap", attr: { contenteditable: "false" } });
          gap.appendChild(fill(k));
          gapParts(gap, pageNumbers[k + 1] - 1);
          return gap;
        },
        { side: -1, key: `gap-${gen}-${k}`, ignoreSelection: true }
      )
    ),
    ...spacers.map((s, k) => Decoration.widget(s.pos, () => spacer(s), { side: -1, key: `spacer-${gen}-${k}`, ignoreSelection: true })),
    Decoration.widget(
      doc.content.size,
      () => {
        const last = createDiv({ cls: "docx-page-last", attr: { contenteditable: "false" } });
        last.appendChild(fill(fills.length - 1));
        last.appendChild(part("docx-page-foot", opts.footer(pages, pages)));
        return last;
      },
      { side: 1, key: `last-${gen}`, ignoreSelection: true }
    ),
  ];
  return { pages, decorations: DecorationSet.create(doc, decos), bands };
}

/** Lays the document out on pages of the document's size, with headers, footers and page gaps. */
export function pagesPlugin(opts: PagesOptions): Plugin<PagesState> {
  return new Plugin<PagesState>({
    key: pagesKey,
    state: {
      init: () => ({ pages: 1, decorations: DecorationSet.empty, bands: [0, 0] }),
      apply(tr, old) {
        const meta = tr.getMeta(pagesKey) as PagesState | undefined;
        return meta ?? { ...old, decorations: old.decorations.map(tr.mapping, tr.doc) };
      },
    },
    props: {
      decorations: (state) => pagesKey.getState(state)?.decorations,
    },
    view(view) {
      let frame = 0;
      const measureNow = () => {
        frame = 0;
        if (!view.dom.isConnected) return;
        // Hiding the gaps to measure shortens the document for a moment; keep the scroll position.
        const scroller = view.dom.parentElement;
        const scrollTop = scroller?.scrollTop ?? 0;
        const next = pageDecorations(view, opts);
        view.dispatch(view.state.tr.setMeta(pagesKey, next).setMeta("addToHistory", false));
        if (scroller) scroller.scrollTop = scrollTop;
        // The header / footer just drawn may be taller than the margins assumed: lay out again once with their real size.
        const setup = opts.setup();
        if (setup && bandHeights(view, setup).some((h, i) => h !== next.bands[i])) schedule();
      };
      // A timer rather than requestAnimationFrame: frames stop while the window is in the background.
      const schedule = () => {
        if (!frame) frame = window.setTimeout(measureNow, 60);
      };
      schedulers.set(view, schedule);
      schedule();
      return {
        update: (v, prev) => {
          if (v.state.doc !== prev.doc) schedule();
        },
        destroy: () => window.clearTimeout(frame),
      };
    },
  });
}
