import { assertAction, assertCell } from "./protection";
/*
 * The workbook as sparse cells over the live worksheet XML. Each cell keeps its <c> element, so a write patches
 * just that element (and the shared string table / styles it needs); every other byte of the sheet stays as it was.
 */
import { NS_S, SheetRef, XlsxWorkbook } from "../../infrastructure/xlsx/xlsx-workbook";
import { parseXml } from "../../infrastructure/ooxml/ooxml-package";
import { Rect, cellName, parseCellRef, parseRange, rangeName } from "./references";
import { translateFormula } from "./formula";
import { ErrorValue, Scalar, formatValue } from "./number-format";
import { Styles, themeColors } from "./styles";

export interface Cell {
  row: number;
  col: number;
  el: Element;
  /** Literal value, or a formula's cached / calculated result. */
  value: Scalar;
  formula?: string;
  /** The formula has no cached result in the file. */
  noCache?: boolean;
  /** cellXfs index. */
  style: number;
  /** Shared formula group id. */
  shared?: string;
  /** Array formula range (on every member; the master has the formula). */
  array?: Rect;
  dynamic?: boolean;
}

export interface CellInput {
  value?: Scalar;
  /** Formula text without "="; null removes the formula. */
  formula?: string | null;
  style?: number;
}

const key = (row: number, col: number) => row * 16384 + col;

/** Records state before changes, for undo: cell elements one by one, whole parts for structural edits. */
export interface Journal {
  cell(sheet: Sheet, row: number, col: number): void;
  part(path: string): void;
}
const isError = (v: Scalar): v is ErrorValue => typeof v === "object" && v !== null;
const DEFAULT_COL_PX = 64;

/** Sizes along one axis: a default plus overrides, with positions by prefix sums. */
export class Axis {
  private keys: number[] | null = null;
  private sums: number[] = [];

  constructor(
    public size: number,
    readonly overrides = new Map<number, number>()
  ) {}

  set(i: number, px: number | null): void {
    if (px === null) this.overrides.delete(i);
    else this.overrides.set(i, px);
    this.keys = null;
  }

  get(i: number): number {
    return this.overrides.get(i) ?? this.size;
  }

  private index(): number[] {
    if (!this.keys) {
      this.keys = [...this.overrides.keys()].sort((a, b) => a - b);
      let acc = 0;
      this.sums = this.keys.map((k) => (acc += this.overrides.get(k)! - this.size));
    }
    return this.keys;
  }

  /** Offset of index i's start. */
  pos(i: number): number {
    const keys = this.index();
    let lo = 0, hi = keys.length; // count of override keys < i
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (keys[mid] < i) lo = mid + 1;
      else hi = mid;
    }
    return i * this.size + (lo ? this.sums[lo - 1] : 0);
  }

  /** Index whose span contains offset px (skipping zero-size indexes). */
  at(px: number): number {
    let lo = 0, hi = 1 << 21;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (this.pos(mid) <= px) lo = mid;
      else hi = mid - 1;
    }
    return lo;
  }
}

export class Sheet {
  cells = new Map<number, Cell>();
  private rowEls = new Map<number, Element>();
  merges: Rect[] = [];
  frozen = { rows: 0, cols: 0 };
  showGrid = true;
  /** Filter range (autoFilter), header row included. */
  filter: Rect | null = null;
  cols: Axis;
  rows: Axis;
  /** Hidden rows / columns (also size 0 in the axes). */
  hiddenRows = new Set<number>();
  hiddenCols = new Set<number>();
  /** Outline (group) levels of rows / columns, and the indexes marked collapsed (the summary row / column). */
  rowLevels = new Map<number, number>();
  colLevels = new Map<number, number>();
  collapsedRows = new Set<number>();
  collapsedCols = new Set<number>();
  /** Group summaries (and their +/- buttons) sit below / right of a group unless sheetPr outlinePr says otherwise. */
  summaryBelow = true;
  summaryRight = true;

  constructor(
    readonly wb: Workbook,
    public ref: SheetRef,
    readonly doc: Document
  ) {
    const root = doc.documentElement;
    const fmt = this.first(root, "sheetFormatPr");
    const defRow = Number(fmt?.getAttribute("defaultRowHeight") ?? 15);
    const defCol = fmt?.getAttribute("defaultColWidth");
    this.rows = new Axis(Math.round((defRow * 4) / 3));
    this.cols = new Axis(defCol ? Math.round(Number(defCol) * 7) : DEFAULT_COL_PX);
    for (const col of this.all(this.first(root, "cols"), "col")) {
      const min = Number(col.getAttribute("min")) - 1, max = Math.min(Number(col.getAttribute("max")) - 1, 16383);
      const hidden = col.getAttribute("hidden") === "1" || col.getAttribute("hidden") === "true";
      const width = col.getAttribute("width");
      const level = Number(col.getAttribute("outlineLevel") ?? 0);
      const collapsed = col.getAttribute("collapsed") === "1" || col.getAttribute("collapsed") === "true";
      for (let c = min; c <= max; c++) {
        if (level) this.colLevels.set(c, level);
        if (collapsed) this.collapsedCols.add(c);
        if (hidden) this.hiddenCols.add(c);
        if (hidden) this.cols.set(c, 0);
        else if (width !== null) this.cols.set(c, Math.round(Number(width) * 7));
      }
    }
    const outlinePr = this.first(this.first(root, "sheetPr"), "outlinePr");
    this.summaryBelow = !/^(0|false)$/.test(outlinePr?.getAttribute("summaryBelow") ?? "");
    this.summaryRight = !/^(0|false)$/.test(outlinePr?.getAttribute("summaryRight") ?? "");
    const sheetData = this.sheetData();
    const shared = new Map<string, { formula: string; row: number; col: number }>();
    const members: Cell[] = [];
    const arrays: Cell[] = [];
    let rowIndex = -1;
    for (const rowEl of this.all(sheetData, "row")) {
      rowIndex = rowEl.hasAttribute("r") ? Number(rowEl.getAttribute("r")) - 1 : rowIndex + 1;
      this.rowEls.set(rowIndex, rowEl);
      const level = Number(rowEl.getAttribute("outlineLevel") ?? 0);
      if (level) this.rowLevels.set(rowIndex, level);
      if (rowEl.getAttribute("collapsed") === "1" || rowEl.getAttribute("collapsed") === "true") this.collapsedRows.add(rowIndex);
      if (rowEl.getAttribute("hidden") === "1" || rowEl.getAttribute("hidden") === "true") {
        this.hiddenRows.add(rowIndex);
        this.rows.set(rowIndex, 0);
      } else if (rowEl.hasAttribute("ht")) this.rows.set(rowIndex, Math.round((Number(rowEl.getAttribute("ht")) * 4) / 3));
      let colIndex = -1;
      for (const c of this.all(rowEl, "c")) {
        const at = parseCellRef(c.getAttribute("r") ?? "");
        colIndex = at ? at.col : colIndex + 1;
        const cell = this.readCell(c, rowIndex, colIndex);
        this.cells.set(key(rowIndex, colIndex), cell);
        const f = this.first(c, "f");
        if (!f) continue;
        const t = f.getAttribute("t");
        if (t === "shared") {
          cell.shared = f.getAttribute("si") ?? "";
          if (f.textContent) shared.set(cell.shared, { formula: f.textContent, row: rowIndex, col: colIndex });
          else members.push(cell);
        } else if (t === "array") {
          const ref = parseRange(f.getAttribute("ref") ?? "");
          if (ref) {
            cell.array = ref;
            arrays.push(cell);
          }
        }
      }
    }
    for (const cell of members) {
      const master = shared.get(cell.shared!);
      if (master) cell.formula = translateFormula(master.formula, cell.row - master.row, cell.col - master.col);
    }
    for (const master of arrays) {
      for (const cell of this.cellsIn(master.array!)) if (cell !== master) cell.array = master.array;
    }
    for (const m of this.all(this.first(root, "mergeCells"), "mergeCell")) {
      const r = parseRange(m.getAttribute("ref") ?? "");
      if (r) this.merges.push(r);
    }
    const view = this.first(this.first(root, "sheetViews"), "sheetView");
    const pane = this.first(view, "pane");
    if (pane && (pane.getAttribute("state") === "frozen" || pane.getAttribute("state") === "frozenSplit")) {
      this.frozen = { rows: Number(pane.getAttribute("ySplit") ?? 0), cols: Number(pane.getAttribute("xSplit") ?? 0) };
    }
    this.showGrid = view?.getAttribute("showGridLines") !== "0";
    this.filter = parseRange(this.first(root, "autoFilter")?.getAttribute("ref") ?? "") ?? null;
  }

  get name(): string {
    return this.ref.name;
  }

  first(parent: Element | null | undefined, tag: string): Element | null {
    for (const c of Array.from(parent?.children ?? [])) if (c.localName === tag && c.namespaceURI === NS_S) return c;
    return null;
  }

  all(parent: Element | null | undefined, tag: string): Element[] {
    return Array.from(parent?.children ?? []).filter((c) => c.localName === tag && c.namespaceURI === NS_S);
  }

  sheetData(): Element {
    return this.first(this.doc.documentElement, "sheetData")!;
  }

  private readCell(c: Element, row: number, col: number): Cell {
    const t = c.getAttribute("t") ?? "n";
    const v = this.first(c, "v")?.textContent ?? null;
    const f = this.first(c, "f");
    let value: Scalar = null;
    if (t === "inlineStr") value = textOf(this.first(c, "is"));
    else if (v !== null) {
      if (t === "s") value = this.wb.stringAt(Number(v));
      else if (t === "b") value = v === "1" || v === "true";
      else if (t === "e") value = { error: v };
      else if (t === "str") value = v;
      else if (t === "d") value = isoToSerial(v, this.wb.date1904);
      else value = Number(v);
    }
    const cell: Cell = { row, col, el: c, value, style: Number(c.getAttribute("s") ?? 0), dynamic: this.wb.pkg.isDynamicMetadata(Number(c.getAttribute("cm"))) };
    if (f && f.getAttribute("t") !== "dataTable") {
      cell.formula = f.textContent ?? "";
      if (v === null && t !== "inlineStr") cell.noCache = true;
    }
    return cell;
  }

  // ---------------------------------------------------------------------------
  // Reading
  // ---------------------------------------------------------------------------

  at(row: number, col: number): Cell | undefined {
    return this.cells.get(key(row, col));
  }

  cell(a1: string): Cell | undefined {
    const r = parseCellRef(a1);
    return r ? this.at(r.row, r.col) : undefined;
  }

  /** Cells inside a rectangle (sparse: only existing ones). */
  cellsIn(r: Rect): Cell[] {
    const out: Cell[] = [];
    const area = (r.r2 - r.r1 + 1) * (r.c2 - r.c1 + 1);
    if (area <= this.cells.size) {
      for (let row = r.r1; row <= r.r2; row++) for (let col = r.c1; col <= r.c2; col++) {
        const c = this.at(row, col);
        if (c) out.push(c);
      }
    } else for (const c of this.cells.values()) if (c.row >= r.r1 && c.row <= r.r2 && c.col >= r.c1 && c.col <= r.c2) out.push(c);
    return out;
  }

  /** Displayed text and color of a cell. */
  display(row: number, col: number): { text: string; color?: string } {
    const c = this.at(row, col);
    if (!c) return { text: "" };
    return formatValue(c.value, this.wb.styles.numFmt(c.style), this.wb.date1904);
  }

  text(a1: string): string {
    const r = parseCellRef(a1)!;
    return this.display(r.row, r.col).text;
  }

  /** Bounding box of cells with values or formulas (null for an empty sheet). */
  used(): Rect | null {
    let r: Rect | null = null;
    for (const c of this.cells.values()) {
      if (c.value === null && c.formula === undefined) continue;
      if (!r) r = { r1: c.row, c1: c.col, r2: c.row, c2: c.col };
      else {
        r.r1 = Math.min(r.r1, c.row);
        r.c1 = Math.min(r.c1, c.col);
        r.r2 = Math.max(r.r2, c.row);
        r.c2 = Math.max(r.c2, c.col);
      }
    }
    return r;
  }

  mergeAt(row: number, col: number): Rect | undefined {
    return this.merges.find((m) => row >= m.r1 && row <= m.r2 && col >= m.c1 && col <= m.c2);
  }

  colWidth(c: number): number {
    return this.cols.get(c);
  }
  rowHeight(r: number): number {
    return this.rows.get(r);
  }
  colX(c: number): number {
    return this.cols.pos(c);
  }
  rowY(r: number): number {
    return this.rows.pos(r);
  }
  colAt(x: number): number {
    return this.cols.at(x);
  }
  rowAt(y: number): number {
    return this.rows.at(y);
  }

  // ---------------------------------------------------------------------------
  // Writing
  // ---------------------------------------------------------------------------

  touch(): void {
    this.wb.pkg.touch(this.ref.path);
  }

  private node(tag: string, text?: string): Element {
    const el = this.doc.createElementNS(NS_S, tag);
    if (text !== undefined) el.textContent = text;
    return el;
  }

  rowEl(row: number, create: boolean): Element | null {
    let el = this.rowEls.get(row);
    if (el || !create) return el ?? null;
    el = this.node("row");
    el.setAttribute("r", String(row + 1));
    const data = this.sheetData();
    const next = this.all(data, "row").find((r) => Number(r.getAttribute("r")) - 1 > row) ?? null;
    data.insertBefore(el, next);
    this.rowEls.set(row, el);
    return el;
  }

  private cellEl(row: number, col: number): Cell {
    const existing = this.at(row, col);
    if (existing) return existing;
    const rowEl = this.rowEl(row, true)!;
    rowEl.removeAttribute("spans"); // an optimization hint that a new cell can invalidate
    const el = this.node("c");
    el.setAttribute("r", cellName(row, col));
    const next = this.all(rowEl, "c").find((c) => (parseCellRef(c.getAttribute("r") ?? "")?.col ?? -1) > col) ?? null;
    rowEl.insertBefore(el, next);
    // A row's own style applies to cells it doesn't list.
    const rowStyle = rowEl.getAttribute("customFormat") === "1" ? Number(rowEl.getAttribute("s") ?? 0) : 0;
    const cell: Cell = { row, col, el, value: null, style: rowStyle || this.colStyle(col) };
    if (cell.style) el.setAttribute("s", String(cell.style));
    this.cells.set(key(row, col), cell);
    return cell;
  }

  /** Effective style, including row and column defaults for empty cells. */
  styleAt(row: number, col: number): number {
    const cell = this.at(row, col);
    if (cell) return cell.style;
    const r = this.rowEl(row, false);
    return r?.getAttribute("customFormat") === "1" ? Number(r.getAttribute("s") ?? 0) : this.colStyle(col);
  }

  private colStyle(col: number): number {
    for (const c of this.all(this.first(this.doc.documentElement, "cols"), "col")) {
      if (col + 1 >= Number(c.getAttribute("min")) && col + 1 <= Number(c.getAttribute("max"))) return Number(c.getAttribute("style") ?? 0);
    }
    return 0;
  }

  /** Write <v> and the type attribute for `value` (keeps <f> and <is> handling to the caller). */
  private writeValue(el: Element, value: Scalar, formula: boolean): void {
    this.first(el, "v")?.remove();
    this.first(el, "is")?.remove();
    el.removeAttribute("t");
    el.removeAttribute("vm"); // rich value metadata no longer matches
    if (!this.wb.pkg.isDynamicMetadata(Number(el.getAttribute("cm")))) el.removeAttribute("cm");
    if (value === null) return;
    let t: string | null = null, v: string;
    if (typeof value === "number") v = String(value);
    else if (typeof value === "boolean") [t, v] = ["b", value ? "1" : "0"];
    else if (isError(value)) [t, v] = ["e", value.error];
    else if (formula) [t, v] = ["str", value];
    else [t, v] = ["s", String(this.wb.addString(value))];
    if (t) el.setAttribute("t", t);
    el.appendChild(this.node("v", v));
  }

  /** Write every shared formula out as plain formulas (before rewriting references structurally). */
  unshareAll(): void {
    for (const si of new Set([...this.cells.values()].map((c) => c.shared).filter((x) => x !== undefined))) this.unshare(si);
  }

  /** Turn every member of a shared formula group into its own plain formula. */
  private unshare(si: string): void {
    for (const c of this.cells.values()) {
      if (c.shared !== si) continue;
      this.wb.journal?.cell(this, c.row, c.col);
      const f = this.first(c.el, "f")!;
      for (const a of ["t", "si", "ref"]) f.removeAttribute(a);
      f.textContent = c.formula ?? "";
      delete c.shared;
    }
  }

  /** Set a cell's value / formula / style. Returns the cell (undefined when it was removed). */
  setCell(row: number, col: number, input: CellInput): Cell | undefined {
    if ("value" in input || "formula" in input) assertCell(this, row, col);
    if (input.style !== undefined) assertAction(this, "formatCells");
    const current = this.at(row, col);
    this.wb.journal?.cell(this, row, col);
    if (current?.array && !(current.array.r1 === row && current.array.c1 === col) && ("value" in input || "formula" in input)) {
      throw new Error("You can't change part of an array formula.");
    }
    if (current?.shared !== undefined && ("value" in input || "formula" in input)) this.unshare(current.shared);
    const cell = this.cellEl(row, col);
    if (input.style !== undefined) {
      cell.style = input.style;
      if (input.style) cell.el.setAttribute("s", String(input.style));
      else cell.el.removeAttribute("s");
    }
    if ("formula" in input || "value" in input) {
      if (input.formula === null || (input.formula === undefined && "value" in input)) {
        this.first(cell.el, "f")?.remove();
        if (cell.array) this.clearArray(cell);
        delete cell.formula;
        delete cell.dynamic;
        cell.el.removeAttribute("cm");
      }
      if (input.formula) {
        let f = this.first(cell.el, "f");
        if (!f || f.getAttribute("t") === "array") {
          f?.remove();
          f = cell.el.insertBefore(this.node("f"), cell.el.firstChild);
          if (cell.array) this.clearArray(cell);
        }
        f.textContent = input.formula;
        cell.formula = input.formula;
        cell.dynamic = true;
      }
      cell.value = input.value ?? null;
      delete cell.noCache;
      this.writeValue(cell.el, cell.value, cell.formula !== undefined);
    }
    this.touch();
    if (cell.value === null && cell.formula === undefined && !cell.style) {
      this.removeCell(cell);
      return undefined;
    }
    this.growDimension(row, col);
    return cell;
  }

  private clearArray(master: Cell): void {
    for (const c of [...this.cellsIn(master.array!)]) {
      if (c !== master && master.dynamic) {
        this.wb.journal?.cell(this,c.row,c.col);
        c.value=null;
        this.writeValue(c.el,null,false);
        if (!c.style) this.removeCell(c);
      }
      delete c.array;
    }
    delete master.array;
  }

  /** Write the calculated spill as cached cells. Opening uses the same values in memory without touching XML. */
  setSpill(master: Cell, rows: Scalar[][], write: boolean): Cell[] {
    const height=rows.length, width=rows[0]?.length ?? 0;
    const range: Rect={r1:master.row,c1:master.col,r2:master.row+height-1,c2:master.col+width-1};
    let blocked = !height || !width || height*width > 100000 || range.r2 > 1048575 || range.c2 > 16383;
    for(const c of this.cellsIn(range)) if(c!==master && !(master.array && c.array === master.array) && (c.value!==null || c.formula!==undefined)) blocked=true;
    if(this.merges.some(m => m.r1<=range.r2 && m.r2>=range.r1 && m.c1<=range.c2 && m.c2>=range.c1)) blocked=true;
    for(const rel of this.wb.pkg.rels(this.ref.path)) if(rel.type === "table") {
      const t=parseRange(this.wb.pkg.xml(rel.target)?.documentElement.getAttribute("ref") ?? "");
      if(t && t.r1<=master.row && t.r2>=master.row && t.c1<=master.col && t.c2>=master.col) blocked=true;
    }
    if(blocked) { rows=[[{error:"#SPILL!"}]];range.r2=master.row;range.c2=master.col; }
    const changed: Cell[]=[];
    if(write) {
      this.wb.journal?.cell(this,master.row,master.col);
      if(master.array) {
        for(const c of this.cellsIn(master.array)) if(c!==master) changed.push(c);
        this.clearArray(master);
      }
      const f=this.first(master.el,"f")!;
      f.setAttribute("t","array");f.setAttribute("ref",rangeName(range));f.setAttribute("aca","1");
      master.el.setAttribute("cm",String(this.wb.pkg.ensureDynamicMetadata()));
    } else if(master.array) {
      for(const c of this.cellsIn(master.array)) if(c!==master) {c.value=null;delete c.array;}
    }
    master.array=range;
    for(let r=0;r<rows.length;r++) for(let c=0;c<rows[r].length;c++) {
      const row=master.row+r,col=master.col+c;
      if(write) this.wb.journal?.cell(this,row,col);
      const cell = write ? this.cellEl(row,col) : this.at(row,col) ?? {row,col,el:this.node("c"),value:null,style:this.styleAt(row,col)};
      if(!write) this.cells.set(key(row,col),cell);
      cell.value=rows[r][c];cell.array=range;delete cell.noCache;
      if(write) {this.writeValue(cell.el,cell.value,cell===master);this.growDimension(row,col);}
      changed.push(cell);
    }
    if(write) this.touch();
    return changed;
  }

  private removeCell(cell: Cell): void {
    const rowEl = cell.el.parentElement;
    cell.el.remove();
    this.cells.delete(key(cell.row, cell.col));
    if (rowEl && !rowEl.children.length && Array.from(rowEl.attributes).every((a) => /^(r|spans|x14ac:dyDescent)$/.test(a.name))) {
      rowEl.remove();
      this.rowEls.delete(cell.row);
    }
  }

  /** Store a formula's newly calculated result as its cache. */
  setCached(cell: Cell, value: Scalar): void {
    this.wb.journal?.cell(this, cell.row, cell.col);
    cell.value = value;
    delete cell.noCache;
    this.writeValue(cell.el, value, true);
    this.touch();
  }

  /** Put back a cell element saved by a journal (null: no cell). */
  restoreCell(row: number, col: number, saved: Element | null): void {
    const current = this.at(row, col);
    if (current) {
      const rowEl = current.el.parentElement;
      current.el.remove();
      this.cells.delete(key(row, col));
      if (rowEl && !saved && !rowEl.children.length && Array.from(rowEl.attributes).every((a) => /^(r|spans|x14ac:dyDescent)$/.test(a.name))) {
        rowEl.remove();
        this.rowEls.delete(row);
      }
    }
    if (saved) {
      const rowEl = this.rowEl(row, true)!;
      const el = saved.cloneNode(true) as Element;
      const next = this.all(rowEl, "c").find((c) => (parseCellRef(c.getAttribute("r") ?? "")?.col ?? -1) > col) ?? null;
      rowEl.insertBefore(el, next);
      const cell = this.readCell(el, row, col);
      if (this.first(el, "f")?.getAttribute("t") === "shared") cell.shared = this.first(el, "f")!.getAttribute("si") ?? "";
      this.cells.set(key(row, col), cell);
    }
    this.touch();
  }

  /** Formulas of shared-formula members, from their group's master (after cells were restored). */
  resolveShared(): void {
    for (const c of this.cells.values()) delete c.array;
    for (const master of this.cells.values()) {
      const f=this.first(master.el,"f");
      const range=f?.getAttribute("t") === "array" ? parseRange(f.getAttribute("ref") ?? "") : null;
      if(range) for(const c of this.cellsIn(range)) c.array=range;
    }
    const masters = new Map<string, Cell>();
    for (const c of this.cells.values()) if (c.shared !== undefined && this.first(c.el, "f")?.textContent) masters.set(c.shared, c);
    for (const c of this.cells.values()) {
      const m = c.shared !== undefined && masters.get(c.shared);
      if (m && m !== c) c.formula = translateFormula(m.formula ?? "", c.row - m.row, c.col - m.col);
    }
  }

  private growDimension(row: number, col: number): void {
    const dim = this.first(this.doc.documentElement, "dimension");
    if (!dim) return;
    const r = parseRange(dim.getAttribute("ref") ?? "") ?? { r1: row, c1: col, r2: row, c2: col };
    const grown = { r1: Math.min(r.r1, row), c1: Math.min(r.c1, col), r2: Math.max(r.r2, row), c2: Math.max(r.c2, col) };
    dim.setAttribute("ref", rangeName(grown));
  }
}

function textOf(el: Element | null): string {
  if (!el) return "";
  // Phonetic runs (rPh) are reading hints, not part of the text.
  return Array.from(el.getElementsByTagNameNS(NS_S, "t"))
    .filter((t) => t.parentElement?.localName !== "rPh")
    .map((t) => t.textContent ?? "")
    .join("");
}

function isoToSerial(iso: string, date1904: boolean): number {
  const ms = Date.parse(iso.endsWith("Z") || /T.*[+-]\d\d:/.test(iso) ? iso : iso + (iso.includes("T") ? "Z" : "T00:00:00Z"));
  const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
  return (ms - epoch) / 86400000;
}

export interface DefinedName {
  name: string;
  formula: string;
  /** Sheet index for sheet-scoped names. */
  local?: number;
  el: Element;
}

export class Workbook {
  sheets: Sheet[] = [];
  journal: Journal | null = null;
  styles: Styles;
  date1904: boolean;
  private strings: string[] = [];
  private stringIndex = new Map<string, number>();

  private constructor(readonly pkg: XlsxWorkbook) {
    const pr = pkg.workbook.getElementsByTagNameNS(NS_S, "workbookPr")[0];
    this.date1904 = pr?.getAttribute("date1904") === "1" || pr?.getAttribute("date1904") === "true";
    const themePath = pkg.relOfType(pkg.workbookPath, "theme");
    const stylesDoc = (pkg.stylesPath && pkg.xml(pkg.stylesPath)) || parseXml(`<styleSheet xmlns="${NS_S}"/>`)!;
    this.styles = new Styles(stylesDoc, themeColors(themePath ? pkg.xml(themePath) : null));
    const sst = pkg.stringsPath ? pkg.xml(pkg.stringsPath) : null;
    Array.from(sst?.documentElement.children ?? []).forEach((si, i) => {
      if (si.localName !== "si") return;
      const text = textOf(si);
      this.strings.push(text);
      const plain = si.children.length === 1 && si.firstElementChild!.localName === "t";
      if (plain && !this.stringIndex.has(text)) this.stringIndex.set(text, i);
    });
    this.sheets = pkg.sheets().map((ref) => new Sheet(this, ref, pkg.xml(ref.path)!));
  }

  static open(pkg: XlsxWorkbook): Workbook {
    return new Workbook(pkg);
  }

  stringAt(i: number): string {
    return this.strings[i] ?? "";
  }

  /** Index of `text` in the shared string table, appending it if new. */
  addString(text: string): number {
    const found = this.stringIndex.get(text);
    const sst = this.pkg.strings().documentElement;
    sst.setAttribute("count", String(Number(sst.getAttribute("count") ?? 0) + 1));
    if (found !== undefined) return found;
    const si = sst.ownerDocument.createElementNS(NS_S, "si");
    const t = si.appendChild(sst.ownerDocument.createElementNS(NS_S, "t"));
    t.textContent = text;
    if (/^\s|\s$/.test(text)) t.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve");
    sst.appendChild(si);
    const i = this.strings.push(text) - 1;
    this.stringIndex.set(text, i);
    sst.setAttribute("uniqueCount", String(this.strings.length));
    this.pkg.touch(this.pkg.stringsPath!);
    return i;
  }

  touchStyles(): void {
    if (this.pkg.stylesPath) this.pkg.touch(this.pkg.stylesPath);
  }

  sheet(name: string): Sheet | undefined {
    const lower = name.toLowerCase();
    return this.sheets.find((s) => s.name.toLowerCase() === lower);
  }

  names(): DefinedName[] {
    return Array.from(this.pkg.workbook.getElementsByTagNameNS(NS_S, "definedName")).map((el) => {
      const local = el.getAttribute("localSheetId");
      return { name: el.getAttribute("name") ?? "", formula: el.textContent ?? "", local: local === null ? undefined : Number(local), el };
    });
  }
}
