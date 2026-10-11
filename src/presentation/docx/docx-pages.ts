/**
 * Word-style page view for the docx editor: pages of the document's size with headers, footers and gaps
 * between them. Page breaks fall between blocks, between the lines of a paragraph and inside table rows;
 * a break inside a paragraph or cell is a blank spacer that carries the drawn gap across the whole page.
 * No Obsidian imports, so it runs under jsdom tests.
 */
import { Node as PMNode } from "prosemirror-model";
import { EditorState, Plugin, PluginKey } from "prosemirror-state";
import { Decoration, DecorationSet, EditorView } from "prosemirror-view";
import { Cell, layoutPages, Line, PageLayout, PageSetup, PageSize, Row, sectionSetups, Spacer, TopBlock, Unit } from "../../domain/docx/pages";
import { sectionBreaks } from "../../domain/docx/sections";
import { docxSchema as schema } from "../../domain/docx/schema";

export interface PagesOptions {
  /** Current page geometry, or null for the continuous (pageless) view. */
  setup(): PageSetup | null;
  /** Header / footer content for a 1-based page number out of `total`. */
  header(page: number, total: number): Node | null;
  /** Keep with next / keep lines together for a paragraph, with its style applied; default: its own settings only. */
  keep?(node: PMNode): { next: boolean; lines: boolean };
  footer(page: number, total: number): Node | null;
  /** The page size setting, for sections with their own page size (default: the document's). */
  pageSize?(): PageSize;
}

interface PagesState {
  pages: number;
  decorations: DecorationSet;
  /** Header / footer band heights the layout assumed. */
  bands: [number, number];
  /** Where each page after the first starts, in document order. */
  pageStarts: { pos: number; page: number }[];
}

/** Page number of document position `pos`, or null when the page view is off. */
export function pageAt(state: EditorState, pos: number): number | null {
  const pages = pagesKey.getState(state);
  if (!pages || !pages.decorations.find().length) return null;
  let page = 1;
  for (const s of pages.pageStarts) if (s.pos <= pos) page = s.page;
  return page;
}

export const pagesKey = new PluginKey<PagesState>("docx-pages");
const schedulers = new WeakMap<EditorView, () => void>();
/** Page gaps cross every column of a column section (set inline: the directory's CSS lint flags multi-column properties). */
const SPAN_ALL = { "column-span": "all" };
/** Height of the grey space between two pages (.docx-page-space). */
const PAGE_SPACE = 24;

/** Re-measure the pages (after a resize or a header / footer change). */
export function repaginate(view: EditorView): void {
  schedulers.get(view)?.();
}

function hasPageBreak(node: PMNode, sectionEnds: Set<PMNode>): boolean {
  let found = sectionEnds.has(node);
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
/** The blocks pages break between: top-level blocks, with content controls and column sections opened up. */
function flowBlocks(doc: PMNode): { node: PMNode; pos: number; columns?: { count: number; group: number } }[] {
  const out: { node: PMNode; pos: number; columns?: { count: number; group: number } }[] = [];
  let group = 0;
  const walk = (parent: PMNode, start: number, columns?: { count: number; group: number }) =>
    parent.forEach((node, offset) => {
      if (node.type === schema.nodes.sdt) walk(node, start + offset + 1, columns);
      else if (node.type === schema.nodes.columns) walk(node, start + offset + 1, { count: node.attrs.count as number, group: ++group });
      else out.push({ node, pos: start + offset, columns });
    });
  walk(doc, 0);
  return out;
}

const COLUMN_BREAK = /<w:br\b[^>]*w:type="column"/;
function hasColumnBreak(node: PMNode): boolean {
  let found = false;
  node.descendants((n) => {
    if (n.type === schema.nodes.opaque && COLUMN_BREAK.test(n.attrs.xml as string)) found = true;
    return !found;
  });
  return found;
}

function measure(view: EditorView, opts: PagesOptions): TopBlock[] {
  const origin = view.dom.getBoundingClientRect();
  const items: { node: PMNode; pos: number; el: HTMLElement | null; top: number; columns?: { count: number; group: number } }[] = [];
  flowBlocks(view.state.doc).forEach(({ node, pos, columns }) => {
    const el = view.nodeDOM(pos);
    const dom = el instanceof HTMLElement ? el : null;
    items.push({ node, pos, el: dom, top: dom ? dom.getBoundingClientRect().top : NaN, columns });
  });
  const sectionEnds = sectionBreaks(view.state.doc);
  return items.map((it, i) => {
    let height = 0;
    if (it.el) {
      const next = items.slice(i + 1).find((n) => n.el);
      height = next ? next.top - it.top : it.el.getBoundingClientRect().height + px(getComputedStyle(it.el).marginBottom);
    }
    const unit = it.el ? measureUnit(view, it.node, it.pos, it.el, height, origin) : { kind: "block" as const, height, pos: it.pos };
    const breakAfter = hasPageBreak(it.node, sectionEnds);
    if (it.columns) return { unit, breakAfter, columns: it.columns, columnBreak: hasColumnBreak(it.node) };
    if (it.node.type !== schema.nodes.paragraph) return { unit, breakAfter };
    // The paragraph's own setting wins over its style's.
    const style = opts.keep?.(it.node) ?? { next: false, lines: false };
    const a = it.node.attrs as { keepNext: boolean | null; keepLines: boolean | null };
    return { unit, breakAfter, keepNext: a.keepNext ?? style.next, keepLines: a.keepLines ?? style.lines };
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
  if (!setup) return { pages: 1, decorations: DecorationSet.empty, bands: [0, 0], pageStarts: [] };
  const zoom = view.dom.style.zoom;
  view.dom.setCssProps({ zoom: "1" });
  let bands: [number, number];
  view.dom.classList.add("docx-measuring");
  let layout: PageLayout;
  try {
    bands = bandHeights(view, setup);
    const blocks = measure(view, opts);
    // Sections with their own page size (a landscape section): pages they start get their own text height.
    const setups = sectionSetups(doc, setup, opts.pageSize?.() ?? "document");
    if (setups.some((st) => st !== setup)) {
      flowBlocks(doc).forEach(({ pos }, k) => {
        const st = setups[doc.resolve(pos).index(0)] ?? setup;
        blocks[k].avail = st.height - Math.max(st.top, bands[0]) - Math.max(st.bottom, bands[1]);
      });
    }
    layout = layoutPages(blocks, setup.height - bands[0] - bands[1], bands[1] + PAGE_SPACE + bands[0]);
  } finally {
    view.dom.classList.remove("docx-measuring");
    view.dom.style.zoom = zoom;
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
    el.setCssProps(SPAN_ALL);
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
          gap.setCssProps(SPAN_ALL);
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
  const pageStarts = [...starts.map((i, k) => ({ pos: offsets[i], page: pageNumbers[k + 1] })), ...spacers.map((s) => ({ pos: s.pos, page: s.page + 1 }))].sort((a, b) => a.pos - b.pos);
  return { pages, decorations: DecorationSet.create(doc, decos), bands, pageStarts };
}

/**
 * Text columns of sections with their own page size, in the page view: the editor is as wide as the widest page
 * (`setup`), and each block is moved and narrowed or widened to its section's margins. Set without touching the
 * paragraph's own margins (its indents): a shift with `translate` and a width change through the right margin.
 */
// ponytail: every page is drawn as wide as the widest one; a portrait page next to landscape ones shows extra white.
export function sectionWidthsPlugin(setup: () => PageSetup | null, final: () => PageSetup | null, size: () => PageSize): Plugin {
  const key = sectionWidthsKey;
  const build = (doc: PMNode) => {
    const page = setup(), base = final();
    if (!page || !base) return DecorationSet.empty;
    const setups = sectionSetups(doc, base, size());
    if (setups.every((st) => st === base) && page.width === base.width) return DecorationSet.empty;
    const textWidth = page.width - base.left - base.right;
    const decos: Decoration[] = [];
    doc.forEach((node, pos, i) => {
      const st = setups[i];
      const dx = st.left - base.left;
      const dw = st.width - st.left - st.right - textWidth;
      if (!dx && !dw) return;
      const right = node.attrs.rightIndent as number | null | undefined;
      const style = `${dx ? `translate: ${dx}px 0;` : ""}${dw ? `margin-right: calc(${right ?? 0}pt - ${dw}px);` : ""}`;
      decos.push(Decoration.node(pos, pos + node.nodeSize, { style }));
    });
    return DecorationSet.create(doc, decos);
  };
  return new Plugin<DecorationSet>({
    key,
    state: { init: (_, state) => build(state.doc), apply: (tr, old) => (tr.docChanged || tr.getMeta(key) ? build(tr.doc) : old) },
    props: { decorations: (state) => key.getState(state) },
  });
}

/** Rebuild the section widths after the page setup changed. */
export function refreshSectionWidths(view: EditorView): void {
  if (sectionWidthsKey.getState(view.state)) view.dispatch(view.state.tr.setMeta(sectionWidthsKey, true).setMeta("addToHistory", false));
}
const sectionWidthsKey = new PluginKey<DecorationSet>("docx-section-widths");

/**
 * The editor is as wide as the widest page; a narrower page (portrait among landscape ones) gets the workspace
 * colour painted over the strip to its right, as layered background gradients.
 */
function shadeNarrowPages(view: EditorView, opts: PagesOptions): void {
  const setup = opts.setup();
  const doc = view.state.doc;
  const setups = setup && opts.pageSize ? sectionSetups(doc, setup, opts.pageSize()) : [];
  const layers: { y: number; h: number; w: number }[] = [];
  if (setup && setups.some((st) => st.width < setup.width)) {
    const zoom = Number.parseFloat(view.dom.style.zoom) || 1;
    const box = view.dom.getBoundingClientRect();
    const gaps = Array.from(view.dom.querySelectorAll(".docx-page-space")).map((g) => g.getBoundingClientRect()).sort((a, b) => a.top - b.top);
    const tops = [0, ...gaps.map((g) => (g.bottom - box.top) / zoom)];
    const bottoms = [...gaps.map((g) => (g.top - box.top) / zoom), box.height / zoom];
    const starts = pagesKey.getState(view.state)?.pageStarts ?? [];
    tops.forEach((top, k) => {
      const pos = k === 0 ? 0 : starts.find((s) => s.page === k + 1)?.pos;
      if (pos === undefined) return;
      const st = setups[doc.resolve(Math.min(pos, doc.content.size)).index(0)] ?? setup;
      if (st.width < setup.width) layers.push({ y: top, h: bottoms[k] - top, w: setup.width - st.width });
    });
  }
  // The editor's shadow outlines the widest page; with narrower pages each gets a thin edge of its own instead.
  const fill = "linear-gradient(var(--background-secondary), var(--background-secondary))";
  const edge = "linear-gradient(rgba(0, 0, 0, 0.2), rgba(0, 0, 0, 0.2))";
  const total = setup?.width ?? 0;
  view.dom.setCssStyles({
    boxShadow: layers.length ? "none" : "",
    backgroundImage: layers.flatMap(() => [edge, fill]).join(", "),
    backgroundRepeat: "no-repeat",
    // The strip runs a little past the editor's edge (clipped there): a fractional width would leave a white sliver.
    backgroundSize: layers.flatMap((l) => [`1px ${l.h}px`, `${l.w + 4}px ${l.h}px`]).join(", "),
    backgroundPosition: layers.flatMap((l) => [`left ${total - l.w}px top ${l.y}px`, `left ${total - l.w + 1}px top ${l.y}px`]).join(", "),
  });
}

/** Lays the document out on pages of the document's size, with headers, footers and page gaps. */
export function pagesPlugin(opts: PagesOptions): Plugin<PagesState> {
  return new Plugin<PagesState>({
    key: pagesKey,
    state: {
      init: () => ({ pages: 1, decorations: DecorationSet.empty, bands: [0, 0], pageStarts: [] }),
      apply(tr, old) {
        const meta = tr.getMeta(pagesKey) as PagesState | undefined;
        return meta ?? { ...old, decorations: old.decorations.map(tr.mapping, tr.doc), pageStarts: old.pageStarts.map((s) => ({ ...s, pos: tr.mapping.map(s.pos) })) };
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
        shadeNarrowPages(view, opts);
        if (scroller) scroller.scrollTop = scrollTop;
        // The header / footer just drawn may be taller than the margins assumed: lay out again once with their real size.
        const setup = opts.setup();
        if (setup) {
          const zoom = view.dom.style.zoom;
          view.dom.setCssProps({ zoom: "1" });
          const bands = bandHeights(view, setup);
          view.dom.style.zoom = zoom;
          if (bands.some((h, i) => h !== next.bands[i])) schedule();
        }
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

/**
 * The paged document cut into page boxes for printing to PDF. Drawn pages can be a few px off the page height,
 * so each box shows a copy of the editor moved up to where its page starts, clipped where the next gap begins.
 */
export function printPages(host: HTMLElement, pageHeight: number): HTMLElement[] {
  const pm = host.querySelector<HTMLElement>(".ProseMirror");
  if (!pm) return [];
  const top = pm.getBoundingClientRect().top;
  const zoom = Number.parseFloat(pm.style.zoom) || 1;
  const gaps = Array.from(pm.querySelectorAll(".docx-page-space"))
    .map((e) => e.getBoundingClientRect())
    .sort((a, b) => a.top - b.top);
  const starts = [0, ...gaps.map((r) => (r.bottom - top) / zoom)];
  const ends = [...gaps.map((r) => (r.top - top) / zoom), Infinity];
  // ponytail: one copy of the whole document per page; fine for typical documents, heavy past a few hundred pages.
  return starts.map((start, k) => {
    const page = createDiv({ cls: "dw-print-page" });
    const clip = page.createDiv({ cls: "dw-print-clip" });
    clip.setCssProps({ "--dw-print-height": `${Math.min(pageHeight, ends[k] - start)}px` });
    const copy = clip.appendChild(host.cloneNode(true) as HTMLElement);
    const editor = copy.querySelector<HTMLElement>(".ProseMirror");
    if (editor) editor.setCssProps({ zoom: "1" });
    copy.setCssProps({ "--dw-print-top": `${-start}px` });
    return page;
  });
}
