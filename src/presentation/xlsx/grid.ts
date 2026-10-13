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
import type { CfStyle } from "../../domain/xlsx/conditional";
import type { Sparkline } from "../../domain/xlsx/sparklines";
import { fontStack } from "../../core/utils/fonts";
import { OutlineAxis, OutlineGroup, outlineGroups, summaryIndex } from "../../domain/xlsx/outline";
import { Op, paint } from "./paint";

export const HEADER_W = 46;
export const HEADER_H = 24;
/** Width of one outline level in the group gutter left of the row headers (and above the column headers). */
const GUTTER = 16;
let accessibilityId = 0;
const REF_COLORS = ["#1a73e8", "#e8710a", "#9334e6", "#188038", "#d93025", "#12b5cb"];

export interface Pos {
  row: number;
  col: number;
}

export interface Hit extends Pos {
  area: "cell" | "colHeader" | "rowHeader" | "corner" | "outline";
  /** Pointer on a header boundary: resizing this column / row. */
  resize?: number;
  /** In the group gutter: the group whose +/- button is under the pointer, if any. */
  outline?: { axis: OutlineAxis; group: OutlineGroup };
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
  /** Conditional formatting of a cell, if any rule applies. */
  conditional?(row: number, col: number): CfStyle | null;
  /** Sparklines on the current sheet. */
  sparklines?(): Sparkline[];
  /** Whether a cell has a dropdown (list validation). */
  dropdown?(row: number, col: number): boolean;
  /** A cell's note, if it has one (marked in its corner, shown on hover). */
  note?(row: number, col: number): string | null;
  comment?(row: number, col: number): boolean;
  /** Whether a cell's value breaks its validation rule (marked, as Sheets does). */
  invalid?(row: number, col: number): boolean;
  /** Show a cell's choices below the point `at` (the arrow's bottom left, in client pixels). */
  openDropdown?(row: number, col: number, at: { x: number; y: number }): void;
  /** Charts on the current sheet: where each sits, and how to draw it into its box. */
  charts?(): { key: string; x: number; y: number; w: number; h: number; draw(el: HTMLElement): void }[];
  chartMenu?(key: string, e: MouseEvent): void;
  /** A chart dragged (or resized from its corner) to `box`, in sheet pixels. */
  moveChart?(key: string, box: { x: number; y: number; w: number; h: number }): void;
  /** The +/- of a row / column group was clicked. */
  toggleGroup?(axis: OutlineAxis, group: OutlineGroup): void;
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
  canvas: HTMLCanvasElement;
  /** DOM over the canvas: selection, reference boxes, filter buttons, pictures. */
  overlay: HTMLElement;
  /** Pane size in CSS pixels. */
  w: number;
  h: number;
  name?: string;
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
  private announcement: HTMLElement;
  private textCache = new Map<number, Record<string, string>>();
  private boxCache = new Map<number, Record<string, string>>();
  /** Font properties last copied onto the editor. */
  private editorFont: string[] = [];
  /** Obsidian's interface font for header labels (a canvas can't use CSS variables), read once. */
  private uiFont = getComputedStyle(document.body).getPropertyValue("--font-interface").trim() || "sans-serif";
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
    // Named through a hidden label: Obsidian shows aria-label as a hover tooltip, which a whole grid shouldn't have.
    const name = this.el.createDiv({ cls: "dw-sr-only", text: "Spreadsheet cells", attr: { id: `dw-grid-name-${++accessibilityId}` } });
    this.canvas = this.scroll.createDiv({ cls: "xlsx-canvas", attr: { role: "region", "aria-labelledby": name.id } });
    this.quads = ["corner", "top", "left", "main"].map((q) => {
      const el = this.canvas.createDiv({ cls: `xlsx-quad xlsx-quad-${q}` });
      el.createEl("canvas", { cls: "xlsx-paint", attr: { "aria-hidden": "true" } });
      el.createDiv({ cls: "xlsx-overlay" });
      return el;
    });
    const heads = (cls: string) => [0, 1].map(() => {
      const el = this.canvas.createDiv({ cls: `xlsx-heads ${cls}` });
      el.createEl("canvas", { cls: "xlsx-paint", attr: { "aria-hidden": "true" } });
      return el;
    });
    this.colHeads = heads("xlsx-colheads");
    this.rowHeads = heads("xlsx-rowheads");
    this.corner = this.canvas.createDiv({ cls: "xlsx-corner", attr: { "aria-label": "Select all" } });
    this.frozenLines = [this.canvas.createDiv({ cls: "xlsx-frozen-line is-row" }), this.canvas.createDiv({ cls: "xlsx-frozen-line" })];
    this.guide = this.canvas.createDiv({ cls: "xlsx-guide" });
    this.guide.hide();
    // Outside the scroll container: typing would otherwise make the browser scroll it to show the caret.
    this.input = this.el.createEl("textarea", { cls: "xlsx-editor", attr: { spellcheck: "false", "aria-multiline": "false", autocapitalize: "off" } });

    const hint = this.el.createDiv({ cls: "dw-sr-only", text: "Use arrow keys to navigate, Enter or F2 to edit, Escape to cancel, Shift+F10 for cell actions. Use F6 to move between cells and toolbar, then Tab to reach other controls. Use the name box to jump to a cell.", attr: { id: `dw-grid-help-${++accessibilityId}` } });
    this.input.setAttribute("aria-describedby", hint.id);
    this.announcement = this.el.createDiv({ cls: "dw-sr-only", attr: { role: "status", "aria-live": "polite", "aria-atomic": "true" } });

    this.scroll.addEventListener("scroll", () => this.schedule());
    this.canvas.addEventListener("pointerdown", (e) => this.onPointerDown(e));
    this.canvas.addEventListener("click", (e) => {
      const chart = (e.target as HTMLElement).closest<HTMLElement>(".xlsx-chart")?.dataset.chart;
      // A drag that just moved the chart isn't a click on it.
      if (this.chartDragged) this.chartDragged = false;
      else if (chart) this.host.chartMenu?.(chart, e);
    });
    this.canvas.addEventListener("dblclick", (e) => this.onDoubleClick(e));
    this.canvas.addEventListener("pointermove", (e) => {
      if (e.buttons) return;
      const hit = this.hit(e.clientX, e.clientY);
      this.canvas.toggleClass("is-col-resize", hit?.area === "colHeader" && hit.resize !== undefined);
      this.canvas.toggleClass("is-row-resize", hit?.area === "rowHeader" && hit.resize !== undefined);
      this.showNote(hit?.area === "cell" ? hit : null);
    });
    this.canvas.addEventListener("pointerleave", () => this.showNote(null));
    this.canvas.addEventListener("contextmenu", (e) => {
      const chart = (e.target as HTMLElement).closest<HTMLElement>(".xlsx-chart")?.dataset.chart;
      if (chart && this.host.chartMenu) {
        e.preventDefault();
        this.host.chartMenu(chart, e);
        return;
      }
      const hit = this.hit(e.clientX, e.clientY);
      if (!hit) return;
      e.preventDefault();
      if (hit.area === "outline") return;
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

  private outlineCache: { sheet: Sheet; row: OutlineGroup[]; col: OutlineGroup[] } | null = null;

  /** Row and column groups of the current sheet (sheets are rebuilt after every structural edit, so cached by sheet). */
  private outline(): { row: OutlineGroup[]; col: OutlineGroup[] } {
    const s = this.sheet;
    if (this.outlineCache?.sheet !== s) this.outlineCache = { sheet: s, row: outlineGroups(s, "row"), col: outlineGroups(s, "col") };
    return this.outlineCache;
  }

  /** Group gutter widths: one level's width per level, plus room for the outermost button. */
  private gutter(axis: OutlineAxis): number {
    const depth = Math.max(0, ...this.outline()[axis].map((g) => g.level));
    return depth ? depth * GUTTER + 4 : 0;
  }

  private get hw(): number {
    return HEADER_W + this.gutter("row");
  }

  private get hh(): number {
    return HEADER_H + this.gutter("col");
  }

  /** Header sizes including group gutters. */
  headerSize(): { hw: number; hh: number } {
    return { hw: this.hw, hh: this.hh };
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
    const vw = width - this.hw - fw, vh = height - this.hh - fh;
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
    // Read every layout value before writing to the DOM: a read after a write forces a full layout each frame.
    const { width, height } = this.viewport();
    const sl = this.scroll.scrollLeft, st = this.scroll.scrollTop;
    // Near the end: add rows / columns, like scrolling a long sheet.
    if (st + height > this.hh + s.rowY(this.nRows) - 400) this.grow(this.nRows + 100, 0);
    if (sl + width > this.hw + s.colX(this.nCols) - 200) this.grow(0, this.nCols + 5);
    const totalW = s.colX(this.nCols), totalH = s.rowY(this.nRows);
    this.spacer.setCssProps({ "--xlsx-w": `${this.hw + totalW}px`, "--xlsx-h": `${this.hh + totalH}px` });
    this.canvas.setCssProps({ "--xlsx-vw": `${width}px`, "--xlsx-vh": `${height}px` });

    const F = s.frozen.rows, C = s.frozen.cols;
    const fw = s.colX(C), fh = s.rowY(F);
    const vw = Math.max(0, width - this.hw), vh = Math.max(0, height - this.hh);
    const frozenCols = this.visible("col", 0, fw, 0, C), frozenRows = this.visible("row", 0, fh, 0, F);
    const cols = this.visible("col", fw + sl, vw - fw, C, this.nCols), rows = this.visible("row", fh + st, vh - fh, F, this.nRows);
    const place = (el: HTMLElement, x: number, y: number, w: number, h: number) =>
      el.setCssProps({ "--x": `${x}px`, "--y": `${y}px`, "--w": `${Math.max(0, w)}px`, "--h": `${Math.max(0, h)}px` });

    const pane = (i: number, rows: number[], cols: number[], ox: number, oy: number, w: number, h: number): Quadrant => ({
      el: this.quads[i],
      canvas: this.quads[i].querySelector("canvas")!,
      overlay: this.quads[i].querySelector<HTMLElement>(".xlsx-overlay")!,
      name: ["corner", "top", "left", "main"][i],
      rows,
      cols,
      ox,
      oy,
      w: Math.max(0, w),
      h: Math.max(0, h),
    });
    const quads: Quadrant[] = [
      pane(0, frozenRows, frozenCols, 0, 0, fw, fh),
      pane(1, frozenRows, cols, 0, 0, vw - fw, fh),
      pane(2, rows, frozenCols, 0, 0, fw, vh - fh),
      pane(3, rows, cols, 0, 0, vw - fw, vh - fh),
    ];
    place(this.quads[0], this.hw, this.hh, fw, fh);
    place(this.quads[1], this.hw + fw, this.hh, vw - fw, fh);
    place(this.quads[2], this.hw, this.hh + fh, fw, vh - fh);
    place(this.quads[3], this.hw + fw, this.hh + fh, vw - fw, vh - fh);
    // Quadrant offsets are relative to the quadrant's own origin.
    quads[1].ox = -sl - fw;
    quads[3].ox = -sl - fw;
    quads[2].oy = -st - fh;
    quads[3].oy = -st - fh;
    for (const q of quads) this.drawQuadrant(q);

    // Thick dividers under frozen rows and right of frozen columns, like Sheets.
    this.frozenLines[0].toggle(F > 0);
    this.frozenLines[1].toggle(C > 0);
    place(this.frozenLines[0], 0, this.hh + fh - 2, width, 4);
    place(this.frozenLines[1], this.hw + fw - 2, 0, 4, height);

    // Headers.
    const sel = this.selection();
    place(this.corner, 0, 0, this.hw, this.hh);
    place(this.colHeads[0], this.hw, 0, fw, this.hh);
    place(this.colHeads[1], this.hw + fw, 0, vw - fw, this.hh);
    place(this.rowHeads[0], 0, this.hh, this.hw, fh);
    place(this.rowHeads[1], 0, this.hh + fh, this.hw, vh - fh);
    const pal = this.palette();
    const font = `11px ${this.uiFont}`;
    const gw = this.gutter("row"), gh = this.gutter("col");
    const headOps = (labels: { text: string; x: number; y: number; w: number; h: number; on: boolean }[], w: number, h: number) => {
      const ops: Op[] = [{ kind: "rect", x: 0, y: 0, w, h, color: pal.head }];
      for (const l of labels) {
        if (l.on) ops.push({ kind: "rect", x: l.x, y: l.y, w: l.w, h: l.h, color: pal.headOn });
        ops.push({ kind: "line", x1: l.x + l.w - 1, y1: l.y, x2: l.x + l.w - 1, y2: l.y + l.h, color: pal.border, width: 1 });
        ops.push({ kind: "line", x1: l.x, y1: l.y + l.h - 1, x2: l.x + l.w, y2: l.y + l.h - 1, color: pal.border, width: 1 });
        ops.push({ kind: "text", x: l.x, y: l.y, w: l.w, h: l.h, text: l.text, font: (l.on ? "bold " : "") + font, size: 11, color: l.on ? pal.headOnInk : pal.headInk, align: "center", valign: "center", wrap: false, underline: false, strike: false, pad: 0 });
      }
      return ops;
    };
    for (const [k, list, ox, w] of [[0, frozenCols, 0, fw], [1, cols, -sl - fw, vw - fw]] as const) {
      const ops = headOps(list.map((c) => ({ text: colName(c), x: s.colX(c) + ox, y: gh, w: s.colWidth(c), h: HEADER_H, on: c >= sel.c1 && c <= sel.c2 })), Math.max(0, w), this.hh);
      ops.push(...this.gutterOps("col", ox, pal, font));
      this.ops.set(`colheads${k}`, ops);
      paint(this.colHeads[k].querySelector("canvas")!, Math.max(0, w), this.hh, ops);
    }
    for (const [k, list, oy, h] of [[0, frozenRows, 0, fh], [1, rows, -st - fh, vh - fh]] as const) {
      const ops = headOps(list.map((r) => ({ text: String(r + 1), x: gw, y: s.rowY(r) + oy, w: HEADER_W, h: s.rowHeight(r), on: r >= sel.r1 && r <= sel.r2 })), this.hw, Math.max(0, h));
      ops.push(...this.gutterOps("row", oy, pal, font));
      this.ops.set(`rowheads${k}`, ops);
      paint(this.rowHeads[k].querySelector("canvas")!, this.hw, Math.max(0, h), ops);
    }

    // The editor / keyboard sink sits on the active cell.
    const a = this.cellBox(this.active.row, this.active.col);
    const box = this.boxOf(a, { left: sl, top: st });
    place(this.input, box.x, box.y, box.w, box.h);
    this.input.toggleClass("is-editing", !!this.edit);
    const name = `${s.name}!${cellName(this.active.row, this.active.col)}`;
    const value = s.display(this.active.row, this.active.col).text || "Blank";
    this.input.setAttribute("aria-label", `${name}, ${value}`);
    const activeCell = s.at(this.active.row, this.active.col);
    const note = this.host.note?.(this.active.row, this.active.col);
    const details = [name, value, activeCell?.formula === undefined ? "" : `Formula: =${activeCell.formula}`,
      note ? `Note: ${note}` : "", this.host.comment?.(this.active.row, this.active.col) ? "Comment thread" : "",
      this.host.invalid?.(this.active.row, this.active.col) ? "Invalid value" : "",
      activeCell && this.host.isStale(activeCell) ? "Result is stale" : "",
      `Selected ${rangeName(sel)}`].filter(Boolean).join(". ");
    if (!this.edit && this.announcement.textContent !== details) this.announcement.setText(details);
    // While editing, the editor shows the cell's font.
    const cell = this.edit ? s.at(this.active.row, this.active.col) : undefined;
    const editFont = Object.fromEntries(Object.entries(cell ? this.css(cell.style, false) : {}).filter(([k]) => /^(font|color)/.test(k)));
    this.input.setCssProps({ ...Object.fromEntries(this.editorFont.map((k) => [k, ""])), ...editFont });
    this.editorFont = Object.keys(editFont);
  }

  /** Where a group's +/- button sits in its header strip: centre across (gutter) and along (sheet offset + `o`). */
  private groupButton(axis: OutlineAxis, g: OutlineGroup, o: number): { across: number; along: number } | null {
    const s = this.sheet;
    const at = summaryIndex(s, axis, g);
    if (at < 0) return null;
    const pos = axis === "row" ? s.rowY(at) : s.colX(at), size = axis === "row" ? s.rowHeight(at) : s.colWidth(at);
    return { across: (g.level - 1) * GUTTER + GUTTER / 2 + 2, along: pos + size / 2 + o };
  }

  /** Brackets and +/- buttons for the groups along one axis, in its header strip's coordinates. */
  private gutterOps(axis: OutlineAxis, o: number, pal: ReturnType<Grid["palette"]>, font: string): Op[] {
    const s = this.sheet;
    const ops: Op[] = [];
    const B = 11;
    for (const g of this.outline()[axis]) {
      const btn = this.groupButton(axis, g, o);
      if (!btn) continue;
      const a = btn.across;
      if (!g.collapsed) {
        const from = (axis === "row" ? s.rowY(g.start) : s.colX(g.start)) + o + 2;
        const to = (axis === "row" ? s.rowY(g.end + 1) : s.colX(g.end + 1)) + o - 2;
        const [lo, hi] = btn.along > to ? [from, btn.along] : [btn.along, to];
        ops.push(axis === "row" ? { kind: "line", x1: a, y1: lo, x2: a, y2: hi, color: pal.headInk, width: 1 } : { kind: "line", x1: lo, y1: a, x2: hi, y2: a, color: pal.headInk, width: 1 });
        // A tick at the far end, like a bracket.
        const end = btn.along > to ? lo : hi;
        ops.push(axis === "row" ? { kind: "line", x1: a, y1: end, x2: a + 4, y2: end, color: pal.headInk, width: 1 } : { kind: "line", x1: end, y1: a, x2: end, y2: a + 4, color: pal.headInk, width: 1 });
      }
      const [x, y] = axis === "row" ? [a - B / 2, btn.along - B / 2] : [btn.along - B / 2, a - B / 2];
      ops.push({ kind: "rect", x, y, w: B, h: B, color: pal.headInk });
      ops.push({ kind: "rect", x: x + 1, y: y + 1, w: B - 2, h: B - 2, color: pal.head });
      ops.push({ kind: "text", x, y, w: B, h: B, text: g.collapsed ? "+" : "−", font, size: 11, color: pal.headInk, align: "center", valign: "center", wrap: false, underline: false, strike: false, pad: 0 });
    }
    return ops;
  }

  /** The group whose +/- button is at gutter offset `across` on row / column `index`. */
  private groupAt(axis: OutlineAxis, across: number, index: number): OutlineGroup | undefined {
    const level = Math.floor(Math.max(0, across - 2) / GUTTER) + 1;
    return this.outline()[axis].find((g) => g.level === level && summaryIndex(this.sheet, axis, g) === index);
  }

  /** The rectangle (sheet coordinates) a cell occupies, merged ranges included. */
  private cellBox(row: number, col: number): Rect {
    return this.sheet.mergeAt(row, col) ?? { r1: row, c1: col, r2: row, c2: col };
  }

  /** Canvas position of a cell rectangle (clamped into the pane the top-left cell is in). */
  private boxOf(r: Rect, scroll = { left: this.scroll.scrollLeft, top: this.scroll.scrollTop }): { x: number; y: number; w: number; h: number } {
    const s = this.sheet;
    const sl = r.c1 >= s.frozen.cols ? scroll.left : 0, st = r.r1 >= s.frozen.rows ? scroll.top : 0;
    return { x: this.hw + s.colX(r.c1) - sl, y: this.hh + s.rowY(r.r1) - st, w: s.colX(r.c2 + 1) - s.colX(r.c1), h: s.rowY(r.r2 + 1) - s.rowY(r.r1) };
  }

  /** Draw part of the sheet into `el` without selection or editing marks (for printing), painted at `scale`. */
  drawRange(el: HTMLElement, r: Rect, scale = 3): void {
    const s = this.sheet;
    const rows: number[] = [], cols: number[] = [];
    for (let i = r.r1; i <= r.r2; i++) if (s.rowHeight(i) > 0) rows.push(i);
    for (let i = r.c1; i <= r.c2; i++) if (s.colWidth(i) > 0) cols.push(i);
    const canvas = el.createEl("canvas", { cls: "xlsx-paint", attr: { "aria-hidden": "true" } });
    const overlay = el.createDiv({ cls: "xlsx-overlay" });
    const w = s.colX(r.c2 + 1) - s.colX(r.c1), h = s.rowY(r.r2 + 1) - s.rowY(r.r1);
    this.drawQuadrant({ el, canvas, overlay, rows, cols, ox: -s.colX(r.c1), oy: -s.rowY(r.r1), w, h }, false, scale);
  }

  /** Colors the canvas can't take from CSS variables. */
  private palette() {
    return this.dark
      ? { ink: "#e3e3e3", line: "#3c4043", head: "#282a2c", headInk: "#9aa0a6", headOn: "#394457", headOnInk: "#a8c7fa", border: "#5f6368" }
      : { ink: "#000000", line: "#e1e3e1", head: "#f8f9fa", headInk: "#444746", headOn: "#d3e3fd", headOnInk: "#0b57d0", border: "#c4c7c5" };
  }

  /** Paint operations of the last frame per pane ("corner", "top", "left", "main", "colheads", "rowheads"). */
  readonly ops = new Map<string, Op[]>();

  private drawQuadrant(q: Quadrant, chrome = true, scale?: number): void {
    const s = this.sheet;
    const pal = this.palette();
    const frag = createFragment();
    const make = (cls: string, x: number, y: number, w: number, h: number) => {
      const el = frag.createDiv({ cls });
      el.setCssProps({ "--x": `${x}px`, "--y": `${y}px`, "--w": `${Math.max(0, w)}px`, "--h": `${Math.max(0, h)}px` });
      return el;
    };
    const ops: Op[] = [];
    const fills: Op[] = [], borders: Op[] = [], texts: Op[] = [];
    if (q.name) this.ops.set(q.name, ops);
    if (!q.rows.length || !q.cols.length) {
      paint(q.canvas, q.w, q.h, [], scale);
      q.overlay.replaceChildren();
      return;
    }
    const x0 = s.colX(q.cols[0]) + q.ox, x1 = s.colX(q.cols[q.cols.length - 1] + 1) + q.ox;
    const y0 = s.rowY(q.rows[0]) + q.oy, y1 = s.rowY(q.rows[q.rows.length - 1] + 1) + q.oy;
    if (chrome ? s.showGrid : this.printGridlines ?? s.showGrid) {
      for (const c of q.cols) ops.push({ kind: "line", x1: s.colX(c + 1) + q.ox - 1, y1: y0, x2: s.colX(c + 1) + q.ox - 1, y2: y1, color: pal.line, width: 1 });
      for (const r of q.rows) ops.push({ kind: "line", x1: x0, y1: s.rowY(r + 1) + q.oy - 1, x2: x1, y2: s.rowY(r + 1) + q.oy - 1, color: pal.line, width: 1 });
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
      const cf = this.host.conditional?.(row, col) ?? null;
      const fmt = cf ? { ...s.wb.styles.format(cell.style), ...cf, fill: cf.fill ?? s.wb.styles.format(cell.style).fill } : s.wb.styles.format(cell.style);
      const fill = this.dark && fmt.fill === "#FFFFFF" && !cf?.fill ? null : fmt.fill;
      if (fill) fills.push({ kind: "rect", x: x - 1, y: y - 1, w: w + 1, h: h + 1, color: fill });
      if (cf?.bar) fills.push({ kind: "rect", x: x + 1, y: y + 2, w: Math.max(0, (w - 3) * cf.bar.fraction), h: Math.max(0, h - 5), color: `${cf.bar.color}B3` });
      const box = this.css(cell.style, true);
      for (const side of ["left", "right", "top", "bottom"] as const) {
        const m = /^(\d+)px (\w+) (#[0-9A-Fa-f]{6})/.exec(box[`border-${side}`] ?? "");
        if (!m) continue;
        const bx = side === "left" ? x - 1 : side === "right" ? x + w - 1 : x - 1;
        const by = side === "top" ? y - 1 : side === "bottom" ? y + h - 1 : y - 1;
        const vertical = side === "left" || side === "right";
        const dash = m[2] === "dashed" ? [4, 2] : m[2] === "dotted" ? [1, 2] : undefined;
        borders.push({ kind: "line", x1: bx, y1: by, x2: vertical ? bx : x + w, y2: vertical ? y + h : by, color: m[3], width: Number(m[1]), dash });
      }
      const shown = s.display(row, col);
      const stale = this.host.isStale(cell);
      if (stale) texts.push({ kind: "corner", x: x + w - 1, y, color: "#e37400" });
      else if (this.host.invalid?.(row, col)) texts.push({ kind: "corner", x: x + w - 1, y, color: "#d93025" });
      const px = (fmt.fontSize * 4) / 3;
      if (cf?.icon) texts.push({ kind: "text", x, y, w, h, text: cf.icon.glyph, font: `${px}px sans-serif`, size: px, color: cf.icon.color, align: "left", valign: "center", wrap: false, underline: false, strike: false, pad: 0 });
      if (!shown.text || cf?.hideValue) return;
      const isNum = typeof cell.value === "number";
      const align = fmt.hAlign !== "general" ? fmt.hAlign : isNum ? "right" : typeof cell.value === "boolean" || typeof cell.value === "object" ? "center" : "left";
      // Unwrapped left-aligned text runs on over empty cells to its right.
      let tw = w;
      if (!fmt.wrap && align === "left" && !isNum && r.c1 === r.c2) {
        for (let c = col + 1; c <= view.c2 && blank(row, c); c++) tw += s.colWidth(c);
      }
      // Automatic text on a filled cell stays dark in the dark theme, as on paper; dark text on no fill turns light.
      let color = fmt.color ?? (fill ? "#000000" : pal.ink);
      if (this.dark && !fill && luminance(color) < 0.25) color = pal.ink;
      if (shown.color) color = shown.color;
      const indent = parseFloat(this.css(cell.style, false)["padding-left"] ?? this.css(cell.style, false)["padding-right"] ?? "0") || 0;
      texts.push({
        kind: "text",
        x,
        y,
        w: tw,
        h,
        text: shown.text,
        font: `${fmt.italic ? "italic " : ""}${fmt.bold ? "bold " : ""}${px}px ${fontStack(fmt.fontName)}`,
        size: px,
        color,
        align: align === "center" || align === "centerContinuous" ? "center" : align === "right" ? "right" : "left",
        valign: fmt.vAlign === "top" ? "top" : fmt.vAlign === "center" ? "center" : "bottom",
        wrap: fmt.wrap,
        underline: fmt.underline,
        strike: fmt.strike,
        pad: (indent ? indent - 3 : 0) + (cf?.icon && align === "left" ? px + 2 : 0),
      });
    };
    for (const m of merges) drawCell(m.r1, m.c1, m);
    for (const row of q.rows) for (const col of q.cols) if (!covered(row, col)) drawCell(row, col, { r1: row, c1: col, r2: row, c2: col });
    // Notes: a small triangle in the cell's top right corner, as Sheets marks them (empty cells too).
    const notes: Op[] = [];
    if (this.host.note || this.host.comment)
      for (const row of q.rows)
        for (const col of q.cols) {
          if (covered(row, col) && !merges.some((m) => m.r1 === row && m.c1 === col)) continue;
          const comment = this.host.comment?.(row, col);
          if (!this.host.note?.(row, col) && !comment) continue;
          const m = merges.find((x) => x.r1 === row && x.c1 === col);
          notes.push({ kind: "corner", x: s.colX((m?.c2 ?? col) + 1) + q.ox - 1, y: s.rowY(row) + q.oy, color: comment ? "#f9ab00" : this.dark ? "#e8eaed" : "#000000" });
        }
    const sparks: Op[] = [];
    for (const sp of this.host.sparklines?.() ?? []) {
      if (sp.row < view.r1 || sp.row > view.r2 || sp.col < view.c1 || sp.col > view.c2) continue;
      const m = merges.find((x) => x.r1 === sp.row && x.c1 === sp.col);
      const x = s.colX(sp.col) + q.ox, y = s.rowY(sp.row) + q.oy;
      sparks.push({ kind: "spark", x, y, w: s.colX((m?.c2 ?? sp.col) + 1) - s.colX(sp.col) - 1, h: s.rowY((m?.r2 ?? sp.row) + 1) - s.rowY(sp.row) - 1, s: sp });
    }
    ops.push(...fills, ...borders, ...texts, ...sparks, ...notes);
    // Painted at the screen's resolution times the zoom, so zoomed text stays sharp.
    paint(q.canvas, q.w, q.h, ops, scale ?? (window.devicePixelRatio || 1) * this.zoom);

    // Pictures float over the cells.
    for (const p of this.host.pictures?.() ?? []) {
      const px = p.x + q.ox, py = p.y + q.oy;
      if (px > x1 || px + p.w < x0 || py > y1 || py + p.h < y0) continue;
      const url = this.host.imageUrl?.(p.image);
      const el = make("xlsx-picture", px, py, p.w, p.h);
      if (url) el.createEl("img", { attr: { src: url, alt: p.alt || p.name, draggable: "false" } });
    }
    for (const ch of this.host.charts?.() ?? []) {
      const px = ch.x + q.ox, py = ch.y + q.oy;
      if (px > x1 || px + ch.w < x0 || py > y1 || py + ch.h < y0) continue;
      const el = make("xlsx-chart", px, py, ch.w, ch.h);
      el.dataset.chart = ch.key;
      ch.draw(el);
      if (this.host.moveChart) el.createDiv({ cls: "xlsx-chart-handle", attr: { "aria-label": "Resize chart" } });
    }
    if (!chrome) {
      q.overlay.replaceChildren(frag);
      return;
    }
    // Dropdown arrows on cells with a list of choices.
    if (this.host.dropdown) {
      for (const row of q.rows) {
        if (!s.rowHeight(row)) continue;
        for (const col of q.cols) {
          if (!s.colWidth(col) || covered(row, col) && !merges.some((m) => m.r1 === row && m.c1 === col) || !this.host.dropdown(row, col)) continue;
          const m = merges.find((x) => x.r1 === row && x.c1 === col);
          const right = s.colX((m?.c2 ?? col) + 1) + q.ox, top = s.rowY(row) + q.oy, h = s.rowY((m?.r2 ?? row) + 1) - s.rowY(row);
          const size = Math.min(16, h - 2);
          const b = make("xlsx-dropdown-btn", right - size - 2, top + (h - size) / 2, size, size);
          b.setText("▾");
          b.dataset.row = String(row);
          b.dataset.col = String(col);
          b.setAttribute("aria-label", `Choose a value for ${cellName(row, col)}`);
        }
      }
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
    q.overlay.replaceChildren(frag);
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
  private notePopup: HTMLElement | null = null;
  private noteAt = "";
  /** Show a cell's note beside it while the pointer is over the cell (null hides it). */
  private showNote(at: { row: number; col: number } | null): void {
    const text = at ? this.host.note?.(at.row, at.col) ?? null : null;
    const key = text && at ? `${at.row}:${at.col}` : "";
    if (key === this.noteAt) return;
    this.noteAt = key;
    this.notePopup?.remove();
    this.notePopup = null;
    if (!text || !at) return;
    const cell = this.boxOf({ r1: at.row, c1: at.col, r2: at.row, c2: at.col });
    const width = this.canvas.getBoundingClientRect().width / this.zoom;
    this.notePopup = this.canvas.createDiv({ cls: "xlsx-note-popup", text });
    this.notePopup.setCssProps({ "--xlsx-note-x": `${Math.min(cell.x + cell.w + 4, width - 220)}px`, "--xlsx-note-y": `${cell.y}px` });
  }

  /** Printing: draw gridlines (the print setting) instead of following the sheet's view. */
  printGridlines: boolean | null = null;

  /** Set while a chart drag ends, so the click that follows doesn't open its menu. */
  private chartDragged = false;

  /** Zoom (1 = 100%): the grid is drawn larger or smaller; pointer positions are scaled back. */
  zoom = 1;
  setZoom(zoom: number): void {
    this.zoom = zoom;
    this.el.setCssProps({ zoom: String(zoom) });
    this.invalidate();
  }

  hit(clientX: number, clientY: number): Hit | null {
    const s = this.sheet;
    const box = this.canvas.getBoundingClientRect();
    // Client positions are on screen; the grid's own pixels are that divided by the zoom.
    const x = (clientX - box.left) / this.zoom, y = (clientY - box.top) / this.zoom;
    const fw = s.colX(s.frozen.cols), fh = s.rowY(s.frozen.rows);
    const sx = x - this.hw < fw ? x - this.hw : x - this.hw + this.scroll.scrollLeft;
    const sy = y - this.hh < fh ? y - this.hh : y - this.hh + this.scroll.scrollTop;
    const col = Math.min(16383, s.colAt(Math.max(0, sx))), row = Math.min(1048575, s.rowAt(Math.max(0, sy)));
    if (x < this.hw && y < this.hh) return { row: 0, col: 0, area: "corner" };
    const gw = this.gutter("row"), gh = this.gutter("col");
    if (x < gw && y >= this.hh) {
      const group = this.groupAt("row", x, row);
      return { row, col, area: "outline", outline: group && { axis: "row", group } };
    }
    if (y < gh && x >= this.hw) {
      const group = this.groupAt("col", y, col);
      return { row, col, area: "outline", outline: group && { axis: "col", group } };
    }
    if (y < this.hh) {
      // Within 4px of a column's right edge (or the left edge of the next): resize that column.
      const right = s.colX(col + 1), left = s.colX(col);
      const resize = right - sx <= 4 ? col : sx - left <= 4 && col > 0 ? this.step("col", col, -1) : undefined;
      return { row, col, area: "colHeader", resize };
    }
    if (x < this.hw) {
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
      else if (e.clientY < b.top + this.hh) this.scroll.scrollTop -= 20;
      if (e.clientX > b.right - 10) this.scroll.scrollLeft += 20;
      else if (e.clientX < b.left + this.hw) this.scroll.scrollLeft -= 20;
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
    // A chart's menu opens on click (see the constructor): opened on press, the release would close it. Dragging
    // moves the chart; dragging its corner handle resizes it.
    const chartEl = target.closest<HTMLElement>(".xlsx-chart");
    if (chartEl?.dataset.chart) {
      e.preventDefault();
      const key = chartEl.dataset.chart;
      const ch = this.host.charts?.().find((c) => c.key === key);
      if (!ch || !this.host.moveChart) return;
      const resize = !!target.closest(".xlsx-chart-handle");
      const left = Number.parseFloat(chartEl.style.getPropertyValue("--x")) || 0, top = Number.parseFloat(chartEl.style.getPropertyValue("--y")) || 0;
      let box = { x: ch.x, y: ch.y, w: ch.w, h: ch.h };
      let moved = false;
      this.chartDragged = false;
      this.track((m) => {
        const dx = (m.clientX - e.clientX) / this.zoom, dy = (m.clientY - e.clientY) / this.zoom;
        if (!moved && Math.abs(dx) + Math.abs(dy) < 4) return;
        moved = true;
        box = resize ? { ...box, w: Math.max(60, ch.w + dx), h: Math.max(40, ch.h + dy) } : { ...box, x: Math.max(0, ch.x + dx), y: Math.max(0, ch.y + dy) };
        chartEl.setCssProps({ "--x": `${left + box.x - ch.x}px`, "--y": `${top + box.y - ch.y}px`, "--w": `${box.w}px`, "--h": `${box.h}px` });
      }, () => {
        if (!moved) return;
        this.chartDragged = true;
        this.host.moveChart?.(key, box);
      }, e);
      return;
    }
    const dropBtn = target.closest<HTMLElement>(".xlsx-dropdown-btn");
    if (dropBtn) {
      e.preventDefault();
      const row = Number(dropBtn.dataset.row), col = Number(dropBtn.dataset.col);
      // Measured first: selecting redraws the overlay, replacing the button.
      const box = dropBtn.getBoundingClientRect();
      this.select({ row, col });
      this.render();
      this.host.openDropdown?.(row, col, { x: box.left, y: box.bottom });
      return;
    }
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
        size = Math.max(axis === "col" ? 8 : 6, size0 + ((axis === "col" ? m.clientX : m.clientY) - start) / this.zoom);
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

    if (hit.area === "outline") {
      e.preventDefault();
      if (hit.outline) this.host.toggleGroup?.(hit.outline.axis, hit.outline.group);
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
    if (!this.edit && ((e.key === "F10" && e.shiftKey) || e.key === "ContextMenu")) {
      e.preventDefault();
      const box = this.input.getBoundingClientRect();
      this.host.contextMenu(new MouseEvent("contextmenu", {clientX: box.left, clientY: box.bottom}), {area: "cell", ...this.active});
      return;
    }
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
    const pageRows = Math.max(1, Math.floor((height - this.hh) / this.sheet.rows.size) - 1);
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
