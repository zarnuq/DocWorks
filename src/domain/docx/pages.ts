/** Page geometry and the editor's page layout (an estimate of Word's: breaks fall between lines and inside table rows). */

const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

export interface PageSetup {
  /** px at 96 dpi */
  width: number;
  height: number;
  top: number;
  right: number;
  bottom: number;
  left: number;
  /** Distance of the header / footer from the page edge. */
  header: number;
  footer: number;
}

export type PageSize = "document" | "letter";

/** Page size and margins from the body's final w:sectPr (Letter, 1" margins by default); `letter` forces 8.5 x 11 in pages. */
export function pageSetup(body: Element, size: PageSize = "document"): PageSetup {
  const sectPr = Array.from(body.children).filter((c) => c.localName === "sectPr").pop();
  const get = (local: string, name: string, fallback: number) => {
    const v = sectPr?.getElementsByTagNameNS(NS_W, local)[0]?.getAttributeNS(NS_W, name);
    return Math.round(Number(v ?? fallback) / 15);
  };
  return {
    width: size === "letter" ? 816 : get("pgSz", "w", 12240),
    height: size === "letter" ? 1056 : get("pgSz", "h", 15840),
    top: Math.abs(get("pgMar", "top", 1440)),
    right: get("pgMar", "right", 1440),
    bottom: Math.abs(get("pgMar", "bottom", 1440)),
    left: get("pgMar", "left", 1440),
    header: get("pgMar", "header", 720),
    footer: get("pgMar", "footer", 720),
  };
}

/** Which header / footer page `page` shows, given the section's "different first page" and "different odd & even" options. */
export function headerFooterType(page: number, firstPage: boolean, oddEven: boolean): "default" | "first" | "even" {
  if (firstPage && page === 1) return "first";
  return oddEven && page % 2 === 0 ? "even" : "default";
}

/** A line of a paragraph: its extent relative to the paragraph top, and the text position it starts at. */
export interface Line {
  top: number;
  bottom: number;
  pos: number;
}

/**
 * What the page layout needs to know about a block, measured from the editor. Heights include the spacing
 * to the next block. `pos` is where a spacer goes to push the unit (or a line of it) to the next page.
 */
export type Unit =
  | { kind: "text"; height: number; pos: number; left: number; lines: Line[] }
  | { kind: "table"; height: number; pos: number; rows: Row[] }
  | { kind: "block"; height: number; pos: number };

export interface Row {
  height: number;
  /** The row moves to the next page whole instead of breaking across it. */
  cantSplit?: boolean;
  cells: Cell[];
}

export interface Cell {
  padTop: number;
  padBottom: number;
  /** Left edge of the cell's content, for drawing the page gap across the whole page. */
  left: number;
  children: Unit[];
}

export interface TopBlock {
  unit: Unit;
  /** A manual page break ends the block. */
  breakAfter: boolean;
  /** Keep on the same page as the start of the next block. */
  keepNext?: boolean;
  /** Don't break inside this paragraph. */
  keepLines?: boolean;
  /** Part of a section laid out in `count` columns; consecutive blocks with the same `group` share them. */
  columns?: { count: number; group: number };
  /** Contains a column break: what follows starts the next column. */
  columnBreak?: boolean;
}

/** Blank space inserted at `pos` that ends a page inside a paragraph or table cell; `gapAt` is where the drawn gap starts (once per page break). */
export interface Spacer {
  pos: number;
  height: number;
  gapAt: number | null;
  left: number;
  /** The page it ends. */
  page: number;
}

export interface PageLayout {
  /** Top-level blocks that start a page after a gap between blocks. */
  starts: number[];
  /** Page number at the top of each page drawn between those gaps. */
  pageNumbers: number[];
  /** Blank space at the bottom of each of those pages. */
  fills: number[];
  /** Page breaks inside paragraphs and table rows. */
  spacers: Spacer[];
  pages: number;
}

/** Word keeps at least this many lines of a paragraph on each side of a page break (widow / orphan control). */
const MIN_LINES = 2;

type Flow = { used: number; page: number };

/**
 * Lay blocks out on pages of `avail` px of text, breaking between blocks, between the lines of a paragraph, and
 * inside table rows (each cell continues at the top of the next page). `gap` is the height of a drawn page gap.
 */
export function layoutPages(blocks: TopBlock[], avail: number, gap: number): PageLayout {
  const out: PageLayout = { starts: [], pageNumbers: [1], fills: [], spacers: [], pages: 1 };
  const shown = new Set<number>();
  let columnFlow: { count: number; base: number; firstPage: number } | null = null;
  const capacity = (page: number) => columnFlow ? columnFlow.count * (avail - (page === columnFlow.firstPage ? columnFlow.base : 0)) : avail;
  const addSpacer = (pos: number, used: number, left: number, page: number) => {
    const drawnUsed = columnFlow ? used / columnFlow.count + (page === columnFlow.firstPage ? columnFlow.base : 0) : used;
    const blank = Math.max(0, avail - drawnUsed);
    out.spacers.push({ pos, height: blank + gap, gapAt: shown.has(page) ? null : blank, left, page });
    shown.add(page);
  };

  /** Flow a paragraph from `used` on `page`; null when it has to move to the next page whole. */
  const flowText = (u: Extract<Unit, { kind: "text" }>, used: number, page: number, keepLines = false): Flow | null => {
    const lines = u.lines;
    if (keepLines && used > 0 && used + u.height > capacity(page) && lines.some((l) => used + l.bottom > capacity(page))) return null;
    let start = 0;
    let startTop = 0;
    let base = used;
    while (base + u.height - startTop > capacity(page)) {
      let k = start;
      while (k < lines.length && base + lines[k].bottom - startTop <= capacity(page)) k++;
      if (k >= lines.length) break; // only the spacing after it overflows
      const nice = lines.length - start >= 2 * MIN_LINES;
      if (nice) k = Math.min(k, lines.length - MIN_LINES);
      if (!nice || k - start < MIN_LINES) {
        if (start === 0 && used > 0) return null;
        k = Math.max(k, start + 1); // a line taller than the page: place it anyway
        if (k >= lines.length) break;
      }
      addSpacer(lines[k].pos, base + lines[k].top - startTop, u.left, page++);
      base = 0;
      startTop = lines[k].top;
      start = k;
    }
    return { used: base + u.height - startTop, page };
  };

  /** Flow one cell's content; `moved` when its first block couldn't start on this page. */
  const flowCell = (cell: Cell, used: number, page: number): Flow & { moved: boolean } => {
    let cu = used + cell.padTop;
    let moved = false;
    cell.children.forEach((child, i) => {
      // Tables inside a cell move whole.
      const r = child.kind === "text" ? flowText(child, cu, page) : cu > 0 && cu + child.height > capacity(page) ? null : { used: cu + child.height, page };
      if (r) return ({ used: cu, page } = r);
      if (i === 0) moved = true;
      addSpacer(child.pos, cu, cell.left, page++);
      cu = 0;
      const again = child.kind === "text" ? flowText(child, 0, page)! : { used: child.height, page };
      ({ used: cu, page } = again);
    });
    return { used: cu + cell.padBottom, page, moved };
  };

  const flowTable = (u: Extract<Unit, { kind: "table" }>, used: number, page: number): Flow | null => {
    for (const [i, row] of u.rows.entries()) {
      if (used + row.height <= capacity(page)) {
        used += row.height;
        continue;
      }
      if (row.cantSplit && used > 0 && row.height <= capacity(page)) {
        // The whole row moves: a spacer opens each of its cells.
        if (i === 0) return null;
        for (const c of row.cells) if (c.children.length) addSpacer(c.children[0].pos, used + c.padTop, c.left, page);
        page++;
        used = row.height;
        continue;
      }
      const mark = out.spacers.length;
      const marked = new Set(shown);
      const cells = row.cells.map((c) => flowCell(c, used, page));
      // A table whose first row can't start here moves whole, like any block.
      if (i === 0 && used > 0 && cells.every((c) => c.moved)) {
        out.spacers.length = mark;
        shown.clear();
        marked.forEach((p) => shown.add(p));
        return null;
      }
      const last = Math.max(...cells.map((c) => c.page));
      used = Math.max(...cells.map((c, k) => (c.page === last ? c.used : row.cells[k].padBottom)));
      page = last;
    }
    return { used, page };
  };

  let used = 0;
  let page = 1;
  const newPage = (i: number) => {
    out.fills.push(Math.max(0, avail - used));
    out.starts.push(i);
    out.pageNumbers.push(++page);
    used = 0;
  };
  const flow = (b: TopBlock): Flow | null => {
    const u = b.unit;
    if (u.kind === "text") return flowText(u, used, page, b.keepLines);
    if (u.kind === "table") return flowTable(u, used, page);
    return used > 0 && used + u.height > avail ? null : { used: used + u.height, page };
  };
  /**
   * A run of blocks in columns: a page holds `count` column-heights of them, drawn balanced (CSS columns), so the
   * page continues below at used + text / count. Blocks move whole between pages; the browser splits them between
   * columns. Returns the index after the run.
   */
  // ponytail: column balancing is the browser's; a page of columns can end a line or two off Word's / Google's.
  const flowColumns = (first: number): number => {
    const { count, group } = blocks[first].columns!;
    let column = avail - used;
    let text = 0;
    let j = first;
    for (; j < blocks.length && blocks[j].columns?.group === group; j++) {
      const b = blocks[j];
      const end = () => {
        used += text / count;
        newPage(j);
        column = avail;
        text = 0;
      };
      if (j > first && blocks[j - 1].breakAfter) end();
      const u = b.unit;
      const startTop = 0;
      if (u.kind === "text" || u.kind === "table") {
        columnFlow = { count, base: used, firstPage: page };
        const run = () => u.kind === "text" ? flowText(u, text, page, b.keepLines) : flowTable(u, text, page);
        let result = run();
        if (!result) {
          columnFlow = null;
          end();
          columnFlow = { count, base: 0, firstPage: page };
          result = run()!;
        }
        if (result.page !== page) { used = 0; column = avail; }
        page = result.page;
        text = result.used;
        columnFlow = null;
      } else {
        if (text + u.height > count * column && (text > 0 || used > 0)) end();
        text += u.height - startTop;
      }
      if (b.columnBreak && text % column) text = Math.min(count * column, Math.ceil(text / column) * column);
    }
    used += text / count;
    return j;
  };

  // Where each block started, to move a chain of "keep with next" blocks back onto a new page.
  const starts: { used: number; page: number; spacers: number }[] = [];
  for (let i = 0; i < blocks.length; i++) {
    const b = blocks[i];
    if (i > 0 && blocks[i - 1].breakAfter) newPage(i);
    if (b.columns) {
      i = flowColumns(i) - 1;
      continue;
    }
    starts[i] = { used, page, spacers: out.spacers.length };
    let r = flow(b);
    if (!r) {
      let j = i;
      while (j > 0 && blocks[j - 1].keepNext && !blocks[j - 1].breakAfter && starts[j - 1].page === page && starts[j - 1].spacers === out.spacers.length) j--;
      if (j < i && starts[j].used > 0) {
        used = starts[j].used;
        newPage(j);
        i = j - 1;
        continue;
      }
      newPage(i);
      starts[i] = { used, page, spacers: out.spacers.length };
      r = flow(b)!;
    }
    ({ used, page } = r);
  }
  out.fills.push(Math.max(0, avail - used));
  out.pages = page;
  return out;
}
