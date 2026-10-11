/*
 * The spreadsheet grid: a native scroll container whose content is a sticky "canvas" the size of the viewport.
 * Only visible rows and columns get DOM, drawn into four clipped quadrants (frozen corner, frozen rows, frozen
 * columns, scrolling body). One textarea is both the cell editor and the keyboard / IME sink while not editing.
 * Plain DOM, no Obsidian: the view supplies a host for everything that changes the workbook.
 */
import { Rect, cellName, colName, rangeName } from "../../domain/xlsx/references";
import { formulaRefs } from "../../domain/xlsx/formula";
import { formatValue, isDateFormat } from "../../domain/xlsx/number-format";
import { numText } from "../../domain/xlsx/formula-evaluator";
import { Cell, Sheet } from "../../domain/xlsx/workbook-model";
import { Picture } from "../../domain/xlsx/drawing";

export const HEADER_W = 46;
export const HEADER_H = 24;
const REF_COLORS = ["#1a73e8", "#e8710a", "#9334e6", "#188038", "#d93025", "#12b5cb"];

export interface Pos {
  row: number;
  col: number;
}

export interface Hit extends Pos {
  area: "cell" | "colHeader" | "rowHeader" | "corner";
  /** Pointer on a header boundary: resizing this column / row. */
  resize?: number;
}

export interface GridHost {
  sheet(): Sheet;
  isStale(cell: Cell): boolean;
  isCircular(cell: Cell): boolean;
  /** Write typed text; false keeps the editor open (e.g. a malformed formula). */
  commit(row: number, col: number, raw: string): boolean;
  selectionChanged(): void;
  /** Editor text while editing (null when editing ends), to mirror in the formula bar. */
  editing(text: string | null): void;
  clear(r: Rect): void;
  fill(src: Rect, target: Rect): void;
  resize(axis: "row" | "col", from: number, to: number, px: number): void;
  autofit(cols: number[]): void;
  contextMenu(e: MouseEvent, hit: Hit): void;
  filterMenu(col: number, anchor: HTMLElement): void;
  /** Pictures on the current sheet, and a URL for an image (null while it loads). */
  pictures?(): Picture[];
  imageUrl?(path: string): string | null;
}

/** Text to edit for a cell: "=formula", or the value written so that typing it back gives the same value. */
export function editText(cell: Cell | undefined, numFmt: string, date1904: boolean): string {
  if (!cell) return "";
  if (cell.formula !== undefined) return "=" + cell.formula;
  const v = cell.value;
  if (v === null) return "";
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  if (typeof v === "object") return (v).error;
  if (typeof v === "string") return v;
  if (isDateFormat(numFmt)) {
    const hasDate = Math.floor(v) !== 0, hasTime = v % 1 !== 0;
    const fmt = [hasDate ? "yyyy-mm-dd" : "", hasTime ? "h:mm:ss" : ""].filter(Boolean).join(" ");
    return formatValue(v, fmt || "yyyy-mm-dd", date1904).text;
  }
  if (/%/.test(numFmt.replace(/"[^"]*"/g, ""))) return numText(v * 100) + "%";
  return numText(v);
}

/** Relative luminance (0 black .. 1 white) of "#RRGGBB". */
function luminance(hex: string): number {
  const [r, g, b] = [1, 3, 5].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255);
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

interface Quadrant {
  el: HTMLElement;
  rows: number[];
  cols: number[];
  /** Offset from sheet coordinates to quadrant coordinates. */
  ox: number;
  oy: number;
}

export class Grid {
  readonly el: HTMLElement;
  private scroll: HTMLElement;
  private spacer: HTMLElement;
  private canvas: HTMLElement;
  private quads: HTMLElement[];
  private colHeads: HTMLElement[];
  private rowHeads: HTMLElement[];
  private corner: HTMLElement;
  readonly input: HTMLTextAreaElement;
  private guide: HTMLElement;
  private frozenLines: HTMLElement[];

  anchor: Pos = { row: 0, col: 0 };
  focus: Pos = { row: 0, col: 0 };
  active: Pos = { row: 0, col: 0 };
  /** Text being edited, or null. `typed`: editing started by typing (arrows then commit and move). */
  edit: { text: string; typed: boolean } | null = null;
  private nRows = 1000;
  private nCols = 26;
  private frame = 0;
  /** Selected text range of the formula being built by clicking cells. */
  private refRange: { start: number; end: number } | null = null;
  private tabStartCol: number | null = null;
  private textCache = new Map<number, Record<string, string>>();
  private boxCache = new Map<number, Record<string, string>>();
  /** Font properties last copied onto the editor. */
  private editorFont: string[] = [];
  /** Dark theme: white fills and black text on unfilled cells take the theme's colors instead. */
  dark = false;
  /** Viewport size for environments without layout (tests). */
  fallbackSize = { width: 1000, height: 600 };

  constructor(
    parent: HTMLElement,
    private host: GridHost
  ) {
    this.el = parent.createDiv({ cls: "xlsx-grid" });
    this.scroll = this.el.createDiv({ cls: "xlsx-scroll" });
    this.spacer = this.scroll.createDiv({ cls: "xlsx-spacer" });
    this.canvas = this.scroll.createDiv({ cls: "xlsx-canvas", attr: { role: "grid" } });
    this.quads = ["corner", "top", "left", "main"].map((q) => this.canvas.createDiv({ cls: `xlsx-quad xlsx-quad-${q}` }));
    this.colHeads = [0, 1].map(() => this.canvas.createDiv({ cls: "xlsx-heads xlsx-colheads" }));
    this.rowHeads = [0, 1].map(() => this.canvas.createDiv({ cls: "xlsx-heads xlsx-rowheads" }));
    this.corner = this.canvas.createDiv({ cls: "xlsx-corner", attr: { "aria-label": "Select all" } });
    this.frozenLines = [this.canvas.createDiv({ cls: "xlsx-frozen-line is-row" }), this.canvas.createDiv({ cls: "xlsx-frozen-line" })];
    this.guide = this.canvas.createDiv({ cls: "xlsx-guide" });
    this.guide.hide();
    // Outside the scroll container: typing would otherwise make the browser scroll it to show the caret.
    this.input = this.el.createEl("textarea", { cls: "xlsx-editor", attr: { spellcheck: "false", "aria-multiline": "false", autocapitalize: "off" } });

    this.scroll.addEventListener("scroll", () => this.schedule());
    this.canvas.addEventListener("pointerdown", (e) => this.onPointerDown(e));
    this.canvas.addEventListener("dblclick", (e) => this.onDoubleClick(e));
    this.canvas.addEventListener("pointermove", (e) => {
      if (e.buttons) return;
      const hit = this.hit(e.clientX, e.clientY);
      this.canvas.toggleClass("is-col-resize", hit?.area === "colHeader" && hit.resize !== undefined);
      this.canvas.toggleClass("is-row-resize", hit?.area === "rowHeader" && hit.resize !== undefined);
    });
    this.canvas.addEventListener("contextmenu", (e) => {
      const hit = this.hit(e.clientX, e.clientY);
      if (!hit) return;
      e.preventDefault();
      const sel = this.selection();
      if (hit.area === "cell" && !this.contains(sel, hit)) this.select(hit, hit);
      // A header outside the selected rows / columns selects its whole row / column first.
      if (hit.area === "rowHeader" && !(sel.c1 === 0 && sel.c2 === 16383 && hit.row >= sel.r1 && hit.row <= sel.r2)) this.select({ row: hit.row, col: 0 }, { row: hit.row, col: 16383 });
      if (hit.area === "colHeader" && !(sel.r1 === 0 && sel.r2 === 1048575 && hit.col >= sel.c1 && hit.col <= sel.c2)) this.select({ row: 0, col: hit.col }, { row: 1048575, col: hit.col });
      this.host.contextMenu(e, hit);
    });
    this.input.addEventListener("keydown", (e) => this.onKey(e));
    this.input.addEventListener("input", () => this.onInput());
    this.input.addEventListener("compositionstart", () => {
      if (!this.edit) this.beginEdit("", true);
    });
  }

  private get sheet(): Sheet {
    return this.host.sheet();
  }

  private viewport(): { width: number; height: number } {
    const w = this.scroll.clientWidth, h = this.scroll.clientHeight;
    return w && h ? { width: w, height: h } : this.fallbackSize;
  }

  /** Redraw on the next frame (scrolling), or now. */
  schedule(): void {
    if (this.frame) return;
    this.frame = window.requestAnimationFrame(() => {
      this.frame = 0;
      this.render();
    });
  }

  /** Forget cached styles (after a format change) and redraw. */
  invalidate(): void {
    this.textCache.clear();
    this.boxCache.clear();
    this.render();
  }

  destroy(): void {
    if (this.frame) window.cancelAnimationFrame(this.frame);
  }

  // ---------------------------------------------------------------------------
  // Selection
  // ---------------------------------------------------------------------------

  /** The selected rectangle, grown to cover any merged range it touches. */
  selection(): Rect {
    let r: Rect = {
      r1: Math.min(this.anchor.row, this.focus.row),
      c1: Math.min(this.anchor.col, this.focus.col),
      r2: Math.max(this.anchor.row, this.focus.row),
      c2: Math.max(this.anchor.col, this.focus.col),
    };
    for (let changed = true; changed; ) {
      changed = false;
      for (const m of this.sheet.merges) {
        if (m.r1 <= r.r2 && r.r1 <= m.r2 && m.c1 <= r.c2 && r.c1 <= m.c2 && (m.r1 < r.r1 || m.c1 < r.c1 || m.r2 > r.r2 || m.c2 > r.c2)) {
          r = { r1: Math.min(r.r1, m.r1), c1: Math.min(r.c1, m.c1), r2: Math.max(r.r2, m.r2), c2: Math.max(r.c2, m.c2) };
          changed = true;
        }
      }
    }
    return r;
  }

  private contains(r: Rect, p: Pos): boolean {
    return p.row >= r.r1 && p.row <= r.r2 && p.col >= r.c1 && p.col <= r.c2;
  }

  select(anchor: Pos, focus: Pos = anchor, active: Pos = anchor): void {
    const clamp = (p: Pos) => ({ row: Math.max(0, Math.min(1048575, p.row)), col: Math.max(0, Math.min(16383, p.col)) });
    this.anchor = clamp(anchor);
    this.focus = clamp(focus);
    this.active = clamp(active);
    const m = this.sheet.mergeAt(this.active.row, this.active.col);
    if (m) this.active = { row: m.r1, col: m.c1 };
    this.grow(Math.max(this.focus.row, this.anchor.row), Math.max(this.focus.col, this.anchor.col));
    this.host.selectionChanged();
    this.render();
  }

  selectRect(r: Rect): void {
    this.select({ row: r.r1, col: r.c1 }, { row: r.r2, col: r.c2 });
  }

  /** Make sure the grid extends past (row, col). */
  private grow(row: number, col: number): void {
    if (row + 50 > this.nRows) this.nRows = Math.min(1048576, row + 500);
    if (col + 5 > this.nCols) this.nCols = Math.min(16384, col + 10);
  }

  /** Scroll so (row, col) is visible (frozen panes are always visible). */
  reveal(p: Pos): void {
    const s = this.sheet;
    const { width, height } = this.viewport();
    const fw = s.colX(s.frozen.cols), fh = s.rowY(s.frozen.rows);
    const vw = width - HEADER_W - fw, vh = height - HEADER_H - fh;
    if (p.col >= s.frozen.cols) {
      const x = s.colX(p.col) - fw, w = s.colWidth(p.col);
      if (x < this.scroll.scrollLeft) this.scroll.scrollLeft = x;
      else if (x + w > this.scroll.scrollLeft + vw) this.scroll.scrollLeft = Math.max(0, x + w - vw);
    }
    if (p.row >= s.frozen.rows) {
      const y = s.rowY(p.row) - fh, h = s.rowHeight(p.row);
      if (y < this.scroll.scrollTop) this.scroll.scrollTop = y;
      else if (y + h > this.scroll.scrollTop + vh) this.scroll.scrollTop = Math.max(0, y + h - vh);
    }
  }

  /** Next visible index from i in direction d (skipping hidden rows / columns). */
  private step(axis: "row" | "col", i: number, d: number): number {
    const size = (k: number) => (axis === "row" ? this.sheet.rowHeight(k) : this.sheet.colWidth(k));
    const max = axis === "row" ? 1048575 : 16383;
    let k = i + d;
    while (k >= 0 && k <= max && size(k) === 0) k += d;
    return k < 0 || k > max ? i : k;
  }

  /** Ctrl+arrow: the edge of the current block of filled cells, or the next filled cell. */
  private jump(p: Pos, dr: number, dc: number): Pos {
    const filled = (r: number, c: number) => {
      const cell = this.sheet.at(r, c);
      return !!cell && (cell.value !== null || cell.formula !== undefined);
    };
    const used = this.sheet.used() ?? { r1: 0, c1: 0, r2: 0, c2: 0 };
    const limit = dr ? (dr > 0 ? Math.max(used.r2, p.row) : 0) : dc > 0 ? Math.max(used.c2, p.col) : 0;
    let { row, col } = p;
    const next = () => ({ row: row + dr, col: col + dc });
    const past = (q: Pos) => (dr ? (dr > 0 ? q.row > limit : q.row < 0) : dc > 0 ? q.col > limit : q.col < 0);
    if (filled(row, col) && !past(next()) && filled(next().row, next().col)) {
      while (!past(next()) && filled(next().row, next().col)) ({ row, col } = next());
      return { row, col };
    }
    while (!past(next())) {
      ({ row, col } = next());
      if (filled(row, col)) return { row, col };
    }
    // Nothing further: go to the sheet's edge in that direction.
    return dr ? { row: dr > 0 ? Math.max(row, this.nRows - 1) : 0, col } : { row, col: dc > 0 ? Math.max(col, this.nCols - 1) : 0 };
  }

  move(dr: number, dc: number, extend: boolean, jump = false): void {
    const from = extend ? this.focus : this.active;
    let to: Pos;
    if (jump) to = this.jump(from, dr, dc);
    else {
      // Step out of a merged range from its far side.
      const m = this.sheet.mergeAt(from.row, from.col);
      const base = m && !extend ? { row: dr > 0 ? m.r2 : m.r1, col: dc > 0 ? m.c2 : m.c1 } : from;
      to = { row: dr ? this.step("row", base.row, dr) : base.row, col: dc ? this.step("col", base.col, dc) : base.col };
    }
    if (extend) this.select(this.anchor, to, this.active);
    else this.select(to);
    this.reveal(to);
  }

  // ---------------------------------------------------------------------------
  // Rendering
  // ---------------------------------------------------------------------------

  /** A cell format's CSS: the box part (fill, borders) or the text part (font, alignment). */
  private css(index: number, box: boolean): Record<string, string> {
    const cache = box ? this.boxCache : this.textCache;
    const hit = cache.get(index);
    if (hit) return hit;
    const out = Object.fromEntries(Object.entries(this.sheet.wb.styles.css(index)).filter(([k]) => (k === "background-color" || k.startsWith("border-")) === box));
    cache.set(index, out);
    return out;
  }

  /** Visible indexes along an axis from offset `from` covering `span` pixels. */
  private visible(axis: "row" | "col", from: number, span: number, start: number, end: number): number[] {
    const s = this.sheet;
    const at = axis === "row" ? s.rowAt(from) : s.colAt(from);
    const pos = (i: number) => (axis === "row" ? s.rowY(i) : s.colX(i));
    const size = (i: number) => (axis === "row" ? s.rowHeight(i) : s.colWidth(i));
    const out: number[] = [];
    for (let i = Math.max(start, at); i < end && pos(i) < from + span; i++) if (size(i) > 0) out.push(i);
    return out;
  }

  render(): void {
    const s = this.sheet;
    const used = s.used();
    if (used) this.grow(used.r2, used.c2);
    this.grow(s.frozen.rows, s.frozen.cols);
    const { width, height } = this.viewport();
    const totalW = s.colX(this.nCols), totalH = s.rowY(this.nRows);
    this.spacer.setCssProps({ "--xlsx-w": `${HEADER_W + totalW}px`, "--xlsx-h": `${HEADER_H + totalH}px` });
    this.canvas.setCssProps({ "--xlsx-vw": `${width}px`, "--xlsx-vh": `${height}px` });
    const sl = this.scroll.scrollLeft, st = this.scroll.scrollTop;
    // Near the end: add rows / columns, like scrolling a long sheet.
    if (st + height > HEADER_H + totalH - 400) this.grow(this.nRows + 100, 0);
    if (sl + width > HEADER_W + totalW - 200) this.grow(0, this.nCols + 5);

    const F = s.frozen.rows, C = s.frozen.cols;
    const fw = s.colX(C), fh = s.rowY(F);
    const vw = Math.max(0, width - HEADER_W), vh = Math.max(0, height - HEADER_H);
    const frozenCols = this.visible("col", 0, fw, 0, C), frozenRows = this.visible("row", 0, fh, 0, F);
    const cols = this.visible("col", fw + sl, vw - fw, C, this.nCols), rows = this.visible("row", fh + st, vh - fh, F, this.nRows);
    const place = (el: HTMLElement, x: number, y: number, w: number, h: number) =>
      el.setCssProps({ "--x": `${x}px`, "--y": `${y}px`, "--w": `${Math.max(0, w)}px`, "--h": `${Math.max(0, h)}px` });

    const quads: Quadrant[] = [
      { el: this.quads[0], rows: frozenRows, cols: frozenCols, ox: 0, oy: 0 },
      { el: this.quads[1], rows: frozenRows, cols, ox: -(fw + sl) + fw, oy: 0 },
      { el: this.quads[2], rows, cols: frozenCols, ox: 0, oy: -(fh + st) + fh },
      { el: this.quads[3], rows, cols, ox: -(fw + sl) + fw, oy: -(fh + st) + fh },
    ];
    place(this.quads[0], HEADER_W, HEADER_H, fw, fh);
    place(this.quads[1], HEADER_W + fw, HEADER_H, vw - fw, fh);
    place(this.quads[2], HEADER_W, HEADER_H + fh, fw, vh - fh);
    place(this.quads[3], HEADER_W + fw, HEADER_H + fh, vw - fw, vh - fh);
    // Quadrant offsets are relative to the quadrant's own origin.
    quads[1].ox = -sl - fw;
    quads[3].ox = -sl - fw;
    quads[2].oy = -st - fh;
    quads[3].oy = -st - fh;
    for (const q of quads) this.drawQuadrant(q);

    // Thick dividers under frozen rows and right of frozen columns, like Sheets.
    this.frozenLines[0].toggle(F > 0);
    this.frozenLines[1].toggle(C > 0);
    place(this.frozenLines[0], 0, HEADER_H + fh - 2, width, 4);
    place(this.frozenLines[1], HEADER_W + fw - 2, 0, 4, height);

    // Headers.
    const sel = this.selection();
    place(this.corner, 0, 0, HEADER_W, HEADER_H);
    place(this.colHeads[0], HEADER_W, 0, fw, HEADER_H);
    place(this.colHeads[1], HEADER_W + fw, 0, vw - fw, HEADER_H);
    place(this.rowHeads[0], 0, HEADER_H, HEADER_W, fh);
    place(this.rowHeads[1], 0, HEADER_H + fh, HEADER_W, vh - fh);
    const head = (parent: HTMLElement, label: string, x: number, y: number, w: number, h: number, on: boolean) => {
      const el = parent.createDiv({ cls: on ? "xlsx-head is-selected" : "xlsx-head", text: label });
      place(el, x, y, w, h);
    };
    for (const [k, list, ox] of [[0, frozenCols, 0], [1, cols, -sl - fw]] as const) {
      const parent = this.colHeads[k];
      parent.empty();
      for (const c of list) head(parent, colName(c), s.colX(c) + ox, 0, s.colWidth(c), HEADER_H, c >= sel.c1 && c <= sel.c2);
    }
    for (const [k, list, oy] of [[0, frozenRows, 0], [1, rows, -st - fh]] as const) {
      const parent = this.rowHeads[k];
      parent.empty();
      for (const r of list) head(parent, String(r + 1), 0, s.rowY(r) + oy, HEADER_W, s.rowHeight(r), r >= sel.r1 && r <= sel.r2);
    }

    // The editor / keyboard sink sits on the active cell.
    const a = this.cellBox(this.active.row, this.active.col);
    const box = this.boxOf(a);
    place(this.input, box.x, box.y, box.w, box.h);
    this.input.toggleClass("is-editing", !!this.edit);
    this.input.setAttribute("aria-label", `${cellName(this.active.row, this.active.col)} ${s.display(this.active.row, this.active.col).text}`);
    // While editing, the editor shows the cell's font.
    const cell = this.edit ? s.at(this.active.row, this.active.col) : undefined;
    const font = Object.fromEntries(Object.entries(cell ? this.css(cell.style, false) : {}).filter(([k]) => /^(font|color)/.test(k)));
    this.input.setCssProps({ ...Object.fromEntries(this.editorFont.map((k) => [k, ""])), ...font });
    this.editorFont = Object.keys(font);
  }

  /** The rectangle (sheet coordinates) a cell occupies, merged ranges included. */
  private cellBox(row: number, col: number): Rect {
    return this.sheet.mergeAt(row, col) ?? { r1: row, c1: col, r2: row, c2: col };
  }

  /** Canvas position of a cell rectangle (clamped into the pane the top-left cell is in). */
  private boxOf(r: Rect): { x: number; y: number; w: number; h: number } {
    const s = this.sheet;
    const sl = r.c1 >= s.frozen.cols ? this.scroll.scrollLeft : 0, st = r.r1 >= s.frozen.rows ? this.scroll.scrollTop : 0;
    return { x: HEADER_W + s.colX(r.c1) - sl, y: HEADER_H + s.rowY(r.r1) - st, w: s.colX(r.c2 + 1) - s.colX(r.c1), h: s.rowY(r.r2 + 1) - s.rowY(r.r1) };
  }

  /** Draw part of the sheet into `el` without selection or editing marks (for printing). */
  drawRange(el: HTMLElement, r: Rect): void {
    const s = this.sheet;
    const rows: number[] = [], cols: number[] = [];
    for (let i = r.r1; i <= r.r2; i++) if (s.rowHeight(i) > 0) rows.push(i);
    for (let i = r.c1; i <= r.c2; i++) if (s.colWidth(i) > 0) cols.push(i);
    this.drawQuadrant({ el, rows, cols, ox: -s.colX(r.c1), oy: -s.rowY(r.r1) }, false);
  }

  private drawQuadrant(q: Quadrant, chrome = true): void {
    const s = this.sheet;
    const frag = createFragment();
    const make = (cls: string, x: number, y: number, w: number, h: number) => {
      const el = frag.createDiv({ cls });
      el.setCssProps({ "--x": `${x}px`, "--y": `${y}px`, "--w": `${Math.max(0, w)}px`, "--h": `${Math.max(0, h)}px` });
      return el;
    };
    if (!q.rows.length || !q.cols.length) {
      q.el.replaceChildren();
      return;
    }
    const x0 = s.colX(q.cols[0]) + q.ox, x1 = s.colX(q.cols[q.cols.length - 1] + 1) + q.ox;
    const y0 = s.rowY(q.rows[0]) + q.oy, y1 = s.rowY(q.rows[q.rows.length - 1] + 1) + q.oy;
    if (s.showGrid) {
      for (const c of q.cols) make("xlsx-vline", s.colX(c + 1) + q.ox - 1, y0, 1, y1 - y0);
      for (const r of q.rows) make("xlsx-hline", x0, s.rowY(r + 1) + q.oy - 1, x1 - x0, 1);
    }
    const view: Rect = { r1: q.rows[0], c1: q.cols[0], r2: q.rows[q.rows.length - 1], c2: q.cols[q.cols.length - 1] };
    const merges = s.merges.filter((m) => m.r1 <= view.r2 && view.r1 <= m.r2 && m.c1 <= view.c2 && view.c1 <= m.c2);
    const covered = (r: number, c: number) => merges.some((m) => r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2);
    const blank = (r: number, c: number) => {
      const cell = s.at(r, c);
      return (!cell || (cell.value === null && cell.formula === undefined) || cell.value === "") && !covered(r, c);
    };
    const drawCell = (row: number, col: number, r: Rect) => {
      const cell = s.at(row, col);
      if (!cell) return;
      const x = s.colX(r.c1) + q.ox, y = s.rowY(r.r1) + q.oy;
      const w = s.colX(r.c2 + 1) - s.colX(r.c1), h = s.rowY(r.r2 + 1) - s.rowY(r.r1);
      let boxCss = this.css(cell.style, true);
      if (this.dark && boxCss["background-color"] === "#FFFFFF") {
        boxCss = { ...boxCss };
        delete boxCss["background-color"];
      }
      if (Object.keys(boxCss).length) make("xlsx-box", x - 1, y - 1, w + 1, h + 1).setCssProps(boxCss);
      const shown = s.display(row, col);
      if (!shown.text && !this.host.isStale(cell)) return;
      const fmt = s.wb.styles.format(cell.style);
      const isNum = typeof cell.value === "number";
      const align = fmt.hAlign !== "general" ? fmt.hAlign : isNum ? "right" : typeof cell.value === "boolean" || typeof cell.value === "object" ? "center" : "left";
      // Unwrapped left-aligned text runs on over empty cells to its right.
      let tw = w;
      if (!fmt.wrap && align === "left" && !isNum && r.c1 === r.c2) {
        for (let c = col + 1; c <= view.c2 && blank(row, c); c++) tw += s.colWidth(c);
      }
      const t = make("xlsx-cell", x, y, tw, h);
      const props: Record<string, string> = { ...this.css(cell.style, false) };
      props["justify-content"] = align === "right" ? "flex-end" : align === "center" ? "center" : "flex-start";
      props["text-align"] = align;
      // Automatic text on a filled cell stays dark in the dark theme, as on paper; dark text on no fill turns light.
      const filled = "background-color" in boxCss;
      if (!props.color && filled) props.color = "#000000";
      if (this.dark && !filled && props.color && luminance(props.color) < 0.25) delete props.color;
      if (shown.color) props.color = shown.color;
      t.setCssProps(props);
      t.createSpan({ text: shown.text });
      if (fmt.wrap) t.addClass("is-wrapped");
      if (this.host.isStale(cell)) {
        t.addClass("is-stale");
        t.setAttribute("title", "This result may be out of date: its formula uses a function that can't be calculated here yet.");
      } else if (this.host.isCircular(cell)) t.setAttribute("title", "Circular reference: this formula depends on its own result.");
    };
    for (const m of merges) drawCell(m.r1, m.c1, m);
    for (const row of q.rows) for (const col of q.cols) if (!covered(row, col)) drawCell(row, col, { r1: row, c1: col, r2: row, c2: col });

    // Pictures float over the cells.
    const qx0 = x0, qx1 = x1, qy0 = y0, qy1 = y1;
    for (const p of this.host.pictures?.() ?? []) {
      const px = p.x + q.ox, py = p.y + q.oy;
      if (px > qx1 || px + p.w < qx0 || py > qy1 || py + p.h < qy0) continue;
      const url = this.host.imageUrl?.(p.image);
      const el = make("xlsx-picture", px, py, p.w, p.h);
      if (url) el.createEl("img", { attr: { src: url, alt: p.alt || p.name, draggable: "false" } });
    }
    if (!chrome) {
      q.el.replaceChildren(frag);
      return;
    }
    // Filter buttons on the header row.
    const f = s.filter;
    if (f && q.rows.includes(f.r1)) {
      for (const col of q.cols) {
        if (col < f.c1 || col > f.c2) continue;
        const size = Math.min(16, s.rowHeight(f.r1) - 2);
        const b = make("xlsx-filter-btn", s.colX(col + 1) + q.ox - size - 2, s.rowY(f.r1) + q.oy + (s.rowHeight(f.r1) - size) / 2, size, size);
        b.setText("▾");
        b.dataset.col = String(col);
        b.setAttribute("aria-label", `Filter ${colName(col)}`);
      }
    }

    // Formula references while editing.
    if (this.edit?.text.startsWith("=")) {
      formulaRefs(this.edit.text.slice(1))
        .filter((t) => t.sheet === undefined || t.sheet.toLowerCase() === s.name.toLowerCase())
        .forEach((t, i) => {
          const r = this.refRect(t.a, t.b);
          const b = make("xlsx-ref", s.colX(r.c1) + q.ox - 1, s.rowY(r.r1) + q.oy - 1, s.colX(r.c2 + 1) - s.colX(r.c1) + 1, s.rowY(r.r2 + 1) - s.rowY(r.r1) + 1);
          b.setCssProps({ "--xlsx-ref": REF_COLORS[i % REF_COLORS.length] });
        });
    }

    // Selection.
    const sel = this.selection();
    if (sel.r1 <= view.r2 + 1 && view.r1 <= sel.r2 + 1 && sel.c1 <= view.c2 + 1 && view.c1 <= sel.c2 + 1) {
      const x = s.colX(sel.c1) + q.ox, y = s.rowY(sel.r1) + q.oy;
      const w = s.colX(sel.c2 + 1) - s.colX(sel.c1), h = s.rowY(sel.r2 + 1) - s.rowY(sel.r1);
      const single = sel.r1 === sel.r2 && sel.c1 === sel.c2;
      const merged = s.mergeAt(sel.r1, sel.c1);
      const isOne = single || (merged && merged.r1 === sel.r1 && merged.c1 === sel.c1 && merged.r2 === sel.r2 && merged.c2 === sel.c2);
      if (!isOne) make("xlsx-sel", x - 1, y - 1, w + 1, h + 1);
      const a = this.cellBox(this.active.row, this.active.col);
      make("xlsx-active", s.colX(a.c1) + q.ox - 1, s.rowY(a.r1) + q.oy - 1, s.colX(a.c2 + 1) - s.colX(a.c1) + 1, s.rowY(a.r2 + 1) - s.rowY(a.r1) + 1);
      if (!this.edit) make("xlsx-handle", x + w - 4, y + h - 4, 7, 7);
    }
    q.el.replaceChildren(frag);
  }

  private refRect(a: { row?: number; col?: number }, b?: { row?: number; col?: number }): Rect {
    const e = b ?? a;
    const r1 = a.row ?? 0, r2 = e.row ?? (a.row === undefined ? 1048575 : r1);
    const c1 = a.col ?? 0, c2 = e.col ?? (a.col === undefined ? 16383 : c1);
    return { r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) };
  }

  // ---------------------------------------------------------------------------
  // Pointer
  // ---------------------------------------------------------------------------

  /** What's under a point (client coordinates). */
  hit(clientX: number, clientY: number): Hit | null {
    const s = this.sheet;
    const box = this.canvas.getBoundingClientRect();
    const x = clientX - box.left, y = clientY - box.top;
    const fw = s.colX(s.frozen.cols), fh = s.rowY(s.frozen.rows);
    const sx = x - HEADER_W < fw ? x - HEADER_W : x - HEADER_W + this.scroll.scrollLeft;
    const sy = y - HEADER_H < fh ? y - HEADER_H : y - HEADER_H + this.scroll.scrollTop;
    const col = Math.min(16383, s.colAt(Math.max(0, sx))), row = Math.min(1048575, s.rowAt(Math.max(0, sy)));
    if (x < HEADER_W && y < HEADER_H) return { row: 0, col: 0, area: "corner" };
    if (y < HEADER_H) {
      // Within 4px of a column's right edge (or the left edge of the next): resize that column.
      const right = s.colX(col + 1), left = s.colX(col);
      const resize = right - sx <= 4 ? col : sx - left <= 4 && col > 0 ? this.step("col", col, -1) : undefined;
      return { row, col, area: "colHeader", resize };
    }
    if (x < HEADER_W) {
      const bottom = s.rowY(row + 1), top = s.rowY(row);
      const resize = bottom - sy <= 3 ? row : sy - top <= 3 && row > 0 ? this.step("row", row, -1) : undefined;
      return { row, col, area: "rowHeader", resize };
    }
    return { row, col, area: "cell" };
  }

  private track(move: (e: PointerEvent) => void, up: (e: PointerEvent) => void, pointer: PointerEvent): void {
    this.canvas.setPointerCapture?.(pointer.pointerId);
    const onMove = (e: PointerEvent) => {
      // Scroll while dragging past the edge.
      const b = this.scroll.getBoundingClientRect();
      if (e.clientY > b.bottom - 10) this.scroll.scrollTop += 20;
      else if (e.clientY < b.top + HEADER_H) this.scroll.scrollTop -= 20;
      if (e.clientX > b.right - 10) this.scroll.scrollLeft += 20;
      else if (e.clientX < b.left + HEADER_W) this.scroll.scrollLeft -= 20;
      move(e);
    };
    const onUp = (e: PointerEvent) => {
      this.canvas.removeEventListener("pointermove", onMove);
      this.canvas.removeEventListener("pointerup", onUp);
      this.canvas.removeEventListener("pointercancel", onUp);
      up(e);
    };
    this.canvas.addEventListener("pointermove", onMove);
    this.canvas.addEventListener("pointerup", onUp);
    this.canvas.addEventListener("pointercancel", onUp);
  }

  /** Can a cell reference be inserted at the editor's caret (right after an operator, "(", "," or "=")? */
  private refInsertable(): boolean {
    if (!this.edit?.text.startsWith("=")) return false;
    if (this.refRange && this.input.selectionStart === this.refRange.end) return true;
    const before = this.input.value.slice(0, this.input.selectionStart).trimEnd();
    return /[=(,+\-*/^&<>:;]$/.test(before);
  }

  private onPointerDown(e: PointerEvent): void {
    if (e.button !== 0) return;
    const target = e.target as HTMLElement;
    if (target === this.input && this.edit) return;
    const filterBtn = target.closest<HTMLElement>(".xlsx-filter-btn");
    if (filterBtn) {
      e.preventDefault();
      this.host.filterMenu(Number(filterBtn.dataset.col), filterBtn);
      return;
    }
    const hit = this.hit(e.clientX, e.clientY);
    if (!hit) return;
    e.preventDefault();
    const s = this.sheet;

    if (hit.area === "cell" && this.refInsertable()) {
      // Building a formula: clicking (and dragging over) cells writes their reference into it.
      const start = this.refRange?.start ?? this.input.selectionStart;
      const write = (a: Pos, b: Pos) => {
        const r = { r1: Math.min(a.row, b.row), c1: Math.min(a.col, b.col), r2: Math.max(a.row, b.row), c2: Math.max(a.col, b.col) };
        const text = rangeName(r);
        const v = this.input.value;
        const end = this.refRange?.end ?? start;
        this.input.value = v.slice(0, start) + text + v.slice(end);
        this.refRange = { start, end: start + text.length };
        this.input.setSelectionRange(this.refRange.end, this.refRange.end);
        this.onInput(true);
      };
      write(hit, hit);
      this.track((m) => {
        const h = this.hit(m.clientX, m.clientY);
        if (h) write(hit, h);
      }, () => this.input.focus(), e);
      return;
    }
    if (this.edit && !this.commitEdit()) return;

    if (target.classList.contains("xlsx-handle")) {
      const src = this.selection();
      let to = src;
      this.track((m) => {
        const h = this.hit(m.clientX, m.clientY);
        if (!h) return;
        const dr = h.row > src.r2 ? h.row - src.r2 : h.row < src.r1 ? h.row - src.r1 : 0;
        const dc = h.col > src.c2 ? h.col - src.c2 : h.col < src.c1 ? h.col - src.c1 : 0;
        to = Math.abs(dr) >= Math.abs(dc)
          ? { ...src, r1: Math.min(src.r1, h.row), r2: Math.max(src.r2, h.row) }
          : { ...src, c1: Math.min(src.c1, h.col), c2: Math.max(src.c2, h.col) };
        this.select({ row: to.r1, col: to.c1 }, { row: to.r2, col: to.c2 }, this.active);
      }, () => {
        if (to !== src) this.host.fill(src, to);
        this.input.focus();
      }, e);
      return;
    }

    if ((hit.area === "colHeader" || hit.area === "rowHeader") && hit.resize !== undefined) {
      const axis = hit.area === "colHeader" ? "col" : "row";
      const i = hit.resize;
      const start = axis === "col" ? e.clientX : e.clientY;
      const size0 = axis === "col" ? s.colWidth(i) : s.rowHeight(i);
      const box = this.boxOf({ r1: axis === "row" ? i : 0, c1: axis === "col" ? i : 0, r2: axis === "row" ? i : 0, c2: axis === "col" ? i : 0 });
      let size = size0;
      this.guide.show();
      const showGuide = () => {
        const p = axis === "col" ? box.x + size : box.y + size;
        this.guide.toggleClass("is-row", axis === "row");
        this.guide.setCssProps({ "--x": `${axis === "col" ? p - 1 : 0}px`, "--y": `${axis === "row" ? p - 1 : 0}px` });
      };
      showGuide();
      this.track((m) => {
        size = Math.max(axis === "col" ? 8 : 6, size0 + (axis === "col" ? m.clientX : m.clientY) - start);
        showGuide();
      }, () => {
        this.guide.hide();
        if (size === size0) return;
        // Resizing one of several selected whole columns / rows resizes them all.
        const sel = this.selection();
        const whole = axis === "col" ? sel.r1 === 0 && sel.r2 === 1048575 : sel.c1 === 0 && sel.c2 === 16383;
        const [a, b] = axis === "col" ? [sel.c1, sel.c2] : [sel.r1, sel.r2];
        if (whole && i >= a && i <= b) this.host.resize(axis, a, b, size);
        else this.host.resize(axis, i, i, size);
      }, e);
      return;
    }

    const extend = e.shiftKey;
    if (hit.area === "corner") {
      this.select({ row: 0, col: 0 }, { row: 1048575, col: 16383 });
      this.input.focus();
      return;
    }
    const pin = (h: Hit): Pos => (hit.area === "colHeader" ? { row: 1048575, col: h.col } : hit.area === "rowHeader" ? { row: h.row, col: 16383 } : h);
    const origin: Pos = hit.area === "colHeader" ? { row: 0, col: hit.col } : hit.area === "rowHeader" ? { row: hit.row, col: 0 } : hit;
    if (extend) this.select(this.anchor, pin(hit), this.active);
    else this.select(origin, pin(hit), origin);
    this.track((m) => {
      const h = this.hit(m.clientX, m.clientY);
      if (h) this.select(this.anchor, pin({ ...h, area: hit.area }), this.active);
    }, () => undefined, e);
    this.input.focus({ preventScroll: true });
  }

  private onDoubleClick(e: MouseEvent): void {
    const hit = this.hit(e.clientX, e.clientY);
    if (!hit) return;
    if (hit.area === "colHeader" && hit.resize !== undefined) {
      const sel = this.selection();
      const whole = sel.r1 === 0 && sel.r2 === 1048575 && hit.resize >= sel.c1 && hit.resize <= sel.c2;
      this.host.autofit(whole ? Array.from({ length: sel.c2 - sel.c1 + 1 }, (_, k) => sel.c1 + k) : [hit.resize]);
      return;
    }
    if (hit.area === "cell") this.beginEdit(null, false);
  }

  // ---------------------------------------------------------------------------
  // Editing
  // ---------------------------------------------------------------------------

  /** Start editing the active cell with `text` (null: its current content). */
  beginEdit(text: string | null, typed: boolean): void {
    const s = this.sheet;
    const cell = s.at(this.active.row, this.active.col);
    const value = text ?? editText(cell, s.wb.styles.numFmt(cell?.style ?? 0), s.wb.date1904);
    this.edit = { text: value, typed };
    this.refRange = null;
    this.input.value = value;
    this.input.focus({ preventScroll: true });
    this.input.setSelectionRange(value.length, value.length);
    this.reveal(this.active);
    this.render();
    this.autosize();
    this.host.editing(value);
  }

  /** Replace the editor's text (from the formula bar). */
  setEditText(text: string): void {
    if (!this.edit) this.beginEdit(text, false);
    this.input.value = text;
    this.edit!.text = text;
    this.render();
  }

  private onInput(fromRef = false): void {
    if (!this.edit) {
      // Typing on a selected cell replaces its content.
      this.beginEdit(this.input.value, true);
      return;
    }
    if (!fromRef) this.refRange = null;
    this.edit.text = this.input.value;
    this.autosize();
    this.render();
    this.host.editing(this.edit.text);
  }

  private autosize(): void {
    this.input.setCssProps({ "--xlsx-edit-w": "0px", "--xlsx-edit-h": "0px" });
    this.input.setCssProps({ "--xlsx-edit-w": `${this.input.scrollWidth + 4}px`, "--xlsx-edit-h": `${this.input.scrollHeight}px` });
  }

  /** Write the edit; false if the host refused it (the editor stays open). */
  commitEdit(): boolean {
    if (!this.edit) return true;
    const text = this.input.value;
    if (!this.host.commit(this.active.row, this.active.col, text)) return false;
    this.endEdit();
    return true;
  }

  cancelEdit(): void {
    this.endEdit();
  }

  private endEdit(): void {
    this.edit = null;
    this.refRange = null;
    this.input.value = "";
    this.input.setCssProps({ "--xlsx-edit-w": "", "--xlsx-edit-h": "" });
    this.host.editing(null);
    this.render();
  }

  private onKey(e: KeyboardEvent): void {
    if (e.isComposing) return;
    const mod = e.ctrlKey || e.metaKey;
    if (this.edit) {
      if (e.key === "Escape") {
        e.preventDefault();
        this.cancelEdit();
        return;
      }
      if (e.key === "Enter" && (e.altKey || mod)) {
        e.preventDefault();
        const i = this.input.selectionStart, v = this.input.value;
        this.input.value = v.slice(0, i) + "\n" + v.slice(this.input.selectionEnd);
        this.input.setSelectionRange(i + 1, i + 1);
        this.onInput();
        return;
      }
      if (e.key === "Enter" || e.key === "Tab") {
        e.preventDefault();
        if (!this.commitEdit()) return;
        this.afterCommitMove(e.key, e.shiftKey);
        return;
      }
      const arrows: Record<string, [number, number]> = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
      if (arrows[e.key] && this.edit.typed && !this.edit.text.startsWith("=")) {
        e.preventDefault();
        if (!this.commitEdit()) return;
        this.move(...arrows[e.key], false);
      }
      return;
    }

    const arrows: Record<string, [number, number]> = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
    if (arrows[e.key]) {
      e.preventDefault();
      this.tabStartCol = null;
      this.move(...arrows[e.key], e.shiftKey, mod);
      return;
    }
    const { height } = this.viewport();
    const pageRows = Math.max(1, Math.floor((height - HEADER_H) / this.sheet.rows.size) - 1);
    switch (e.key) {
      case "Tab":
        e.preventDefault();
        if (this.tabStartCol === null) this.tabStartCol = this.active.col;
        this.move(0, e.shiftKey ? -1 : 1, false);
        return;
      case "Enter":
        e.preventDefault();
        if (e.shiftKey) this.move(-1, 0, false);
        else this.beginEdit(null, false);
        return;
      case "F2":
        e.preventDefault();
        this.beginEdit(null, false);
        return;
      case "Delete":
      case "Backspace":
        e.preventDefault();
        this.host.clear(this.selection());
        return;
      case "Home":
        e.preventDefault();
        this.select(mod ? { row: 0, col: 0 } : { row: this.active.row, col: 0 });
        this.reveal(this.active);
        return;
      case "End": {
        e.preventDefault();
        const used = this.sheet.used() ?? { r1: 0, c1: 0, r2: 0, c2: 0 };
        this.select(mod ? { row: used.r2, col: used.c2 } : { row: this.active.row, col: used.c2 });
        this.reveal(this.active);
        return;
      }
      case "PageDown":
      case "PageUp": {
        e.preventDefault();
        const d = e.key === "PageDown" ? pageRows : -pageRows;
        this.scroll.scrollTop += d * this.sheet.rows.size;
        this.select({ row: Math.max(0, this.active.row + d), col: this.active.col });
        this.reveal(this.active);
        return;
      }
    }
  }

  /** After Enter / Tab commits: move down (back to where a Tab run started) or sideways. */
  private afterCommitMove(key: string, shift: boolean): void {
    if (key === "Tab") {
      if (this.tabStartCol === null) this.tabStartCol = this.active.col;
      this.move(0, shift ? -1 : 1, false);
      return;
    }
    const col = this.tabStartCol ?? this.active.col;
    this.tabStartCol = null;
    const sel = this.selection();
    const m = this.sheet.mergeAt(this.active.row, this.active.col);
    const row = shift ? this.step("row", this.active.row, -1) : this.step("row", m ? m.r2 : this.active.row, 1);
    if (sel.r1 !== sel.r2 || sel.c1 !== sel.c2) {
      // Enter inside a selected range moves within it.
      const next = row > sel.r2 ? { row: sel.r1, col: col + 1 > sel.c2 ? sel.c1 : col + 1 } : { row, col };
      this.select(this.anchor, this.focus, next);
    } else this.select({ row, col });
    this.reveal(this.active);
  }

  scrollHome(): void {
    this.scroll.scrollTop = 0;
    this.scroll.scrollLeft = 0;
  }

  /** Put keyboard focus on the grid. */
  focusGrid(): void {
    this.input.focus({ preventScroll: true });
  }
}
