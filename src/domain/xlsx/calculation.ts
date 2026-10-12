import { tableReference } from "./tables";
/*
 * Recalculation. Formula cells index the rectangles they read (bucketed by column, so whole-column references
 * aren't expanded); an edit marks the formulas reading the changed cells, transitively, and only those are
 * recalculated, in dependency order (a formula reading a pending cell calculates it first).
 *
 * Formulas using something not implemented keep the file's cached result; when their inputs change they're marked
 * stale (and so is whatever reads them), and the workbook asks Excel to recalculate fully on open.
 */
import { NS_S } from "../../infrastructure/xlsx/xlsx-workbook";
import { Node, parseFormula } from "./formula";
import { EvalContext, Unsupported, evaluate, evaluateArray, isVolatile } from "./formula-evaluator";
import { Rect } from "./references";
import { ErrorValue, Scalar } from "./number-format";
import { Cell, Sheet, Workbook } from "./workbook-model";

interface Dep {
  sheet: Sheet;
  rect: Rect;
}

interface Info {
  sheet: Sheet;
  cell: Cell;
  node: Node | null;
  deps: Dep[];
  /** Recalculate on every change (volatile functions, unparseable formulas). */
  always: boolean;
}

/** Columns a rectangle may span before it goes in the "wide" list instead of per-column buckets. */
const WIDE = 64;

class Circular extends Error {
  constructor(readonly cell: Cell) {
    super("circular reference");
  }
}

const same = (a: Scalar, b: Scalar) => a === b || (typeof a === "object" && typeof b === "object" && a !== null && b !== null && a.error === b.error);

export interface Change {
  sheet: Sheet;
  row: number;
  col: number;
}

export class Calc {
  static readonly UNSUPPORTED: ErrorValue = { error: "#UNSUPPORTED" };

  private infos = new Map<Cell, Info>();
  /** Indexed formula cell per sheet position (row * 16384 + col), to drop entries of cells that were replaced. */
  private byPos = new Map<Sheet, Map<number, Cell>>();
  private byCol = new Map<Sheet, Map<number, Set<Info>>>();
  private wide = new Map<Sheet, Set<Info>>();
  private always = new Set<Info>();
  private parsed = new Map<string, Node | null>();
  private names = new Map<string, Node | null>();
  private stale = new Set<Cell>();
  private circular = new Set<Cell>();

  /** Cells waiting to be recalculated in the current pass, and the ones in progress (cycle detection). */
  private pending = new Set<Cell>();
  private computing = new Set<Cell>();
  /** Cells without a cached result already worked out in memory during this pass. */
  private done = new Set<Cell>();
  private readStale = false;
  private write = true;

  constructor(private wb: Workbook) {
    for (const sheet of wb.sheets) for (const cell of sheet.cells.values()) if (cell.formula !== undefined) this.index(sheet, cell);
  }

  isStale(cell: Cell): boolean {
    return this.stale.has(cell);
  }
  isCircular(cell: Cell): boolean {
    return this.circular.has(cell);
  }

  private parse(formula: string): Node | null {
    if (!this.parsed.has(formula)) {
      let node: Node | null = null;
      try {
        node = parseFormula(formula);
      } catch {
        node = null;
      }
      this.parsed.set(formula, node);
    }
    return this.parsed.get(formula)!;
  }

  /** A defined name's parsed formula: sheet-scoped names win over workbook ones. */
  private name(name: string, sheet: Sheet): Node | null {
    const idx = this.wb.sheets.indexOf(sheet);
    const lower = name.toLowerCase();
    const key = `${idx}|${lower}`;
    if (!this.names.has(key)) {
      const all = this.wb.names().filter((n) => n.name.toLowerCase() === lower);
      const def = all.find((n) => n.local === idx) ?? all.find((n) => n.local === undefined);
      this.names.set(key, def ? this.parse(def.formula) : null);
    }
    return this.names.get(key)!;
  }

  /** Forget cached name definitions (after names change). */
  resetNames(): void {
    this.names.clear();
  }

  private collect(node: Node, sheet: Sheet, info: Info, seen = new Set<string>()): void {
    switch (node.type) {
      case "spill":
        info.always = true;
        this.collect(node.arg, sheet, info, seen);
        break;
      case "struct":
        this.collect(tableReference(this.wb, sheet, info.cell.row, info.cell.col, node.text), sheet, info, seen);
        break;
      case "ref": {
        const target = node.sheet === undefined ? sheet : this.wb.sheet(node.sheet);
        if (target) info.deps.push({ sheet: target, rect: node.range });
        break;
      }
      case "name": {
        const def = this.name(node.name, sheet) ?? tableReference(this.wb, sheet, info.cell.row, info.cell.col, node.name);
        if (def && !seen.has(node.name.toLowerCase())) {
          seen.add(node.name.toLowerCase());
          this.collect(def, sheet, info, seen);
        }
        break;
      }
      case "call":
        if (isVolatile(node.name) || node.name === "ANCHORARRAY") info.always = true;
        node.args.forEach((a) => this.collect(a, sheet, info, seen));
        break;
      case "un":
        this.collect(node.arg, sheet, info, seen);
        break;
      case "bin":
        this.collect(node.left, sheet, info, seen);
        this.collect(node.right, sheet, info, seen);
        break;
    }
  }

  private unindex(cell: Cell): void {
    const info = this.infos.get(cell);
    if (!info) return;
    this.infos.delete(cell);
    const pos = this.byPos.get(info.sheet);
    if (pos?.get(cell.row * 16384 + cell.col) === cell) pos.delete(cell.row * 16384 + cell.col);
    this.always.delete(info);
    this.wide.get(info.sheet)?.delete(info);
    for (const d of info.deps) {
      if (d.rect.c2 - d.rect.c1 >= WIDE) this.wide.get(d.sheet)?.delete(info);
      else for (let c = d.rect.c1; c <= d.rect.c2; c++) this.byCol.get(d.sheet)?.get(c)?.delete(info);
    }
  }

  private index(sheet: Sheet, cell: Cell): void {
    this.unindex(cell);
    if (cell.formula === undefined) return;
    const node = this.parse(cell.formula);
    const info: Info = { sheet, cell, node, deps: [], always: !node || !!cell.dynamic };
    if (node) this.collect(node, sheet, info);
    this.infos.set(cell, info);
    if (!this.byPos.has(sheet)) this.byPos.set(sheet, new Map());
    this.byPos.get(sheet)!.set(cell.row * 16384 + cell.col, cell);
    if (info.always) this.always.add(info);
    for (const d of info.deps) {
      if (d.rect.c2 - d.rect.c1 >= WIDE) {
        if (!this.wide.has(d.sheet)) this.wide.set(d.sheet, new Set());
        this.wide.get(d.sheet)!.add(info);
      } else {
        if (!this.byCol.has(d.sheet)) this.byCol.set(d.sheet, new Map());
        const cols = this.byCol.get(d.sheet)!;
        for (let c = d.rect.c1; c <= d.rect.c2; c++) {
          if (!cols.has(c)) cols.set(c, new Set());
          cols.get(c)!.add(info);
        }
      }
    }
  }

  /** Formulas reading cell (sheet, row, col). */
  private readers(sheet: Sheet, row: number, col: number): Info[] {
    const out: Info[] = [];
    for (const set of [this.byCol.get(sheet)?.get(col), this.wide.get(sheet)]) {
      for (const info of set ?? []) {
        if (info.deps.some((d) => d.sheet === sheet && row >= d.rect.r1 && row <= d.rect.r2 && col >= d.rect.c1 && col <= d.rect.c2)) out.push(info);
      }
    }
    return out;
  }

  private context(sheet: Sheet, row: number, col: number): EvalContext {
    return {
      wb: this.wb,
      sheet,
      row,
      col,
      value: (s, r, c) => {
        let cell = s.at(r, c);
        const anchor = cell?.array && s.at(cell.array.r1,cell.array.c1);
        if(anchor && anchor !== cell && this.pending.has(anchor)) {this.compute(anchor);cell=s.at(r,c);}
        if (!cell) return null;
        if (this.computing.has(cell)) throw new Circular(cell);
        if (this.pending.has(cell) || (cell.noCache && !this.done.has(cell))) this.compute(cell);
        if (this.stale.has(cell)) this.readStale = true;
        return cell.value;
      },
      name: (n, s) => this.name(n, s),
    };
  }

  /** Evaluate `formula` as if it were in cell (row, col) of `sheet`. */
  evaluate(sheet: Sheet, row: number, col: number, formula: string): Scalar {
    const node = this.parse(formula);
    if (!node) return Calc.UNSUPPORTED;
    this.write = false;
    try {
      return evaluate(node, this.context(sheet, row, col));
    } catch (e) {
      if (e instanceof Unsupported) return Calc.UNSUPPORTED;
      throw e;
    } finally {
      this.done.clear();
    }
  }

  private changed: Cell[] = [];

  private compute(cell: Cell): void {
    this.pending.delete(cell);
    this.done.add(cell);
    const info = this.infos.get(cell);
    if (!info) return;
    this.computing.add(cell);
    const outerStale = this.readStale;
    this.readStale = false;
    let value: Scalar;
    let unsupported = !info.node;
    try {
      if (info.node && cell.dynamic) {
        const rows=evaluateArray(info.node,this.context(info.sheet,cell.row,cell.col));
        const wantsSpill=!!cell.array || rows.length > 1 || (rows[0]?.length ?? 0) > 1 || /(?:^|\W)(?:_xlfn\.)?(?:_xlws\.)?(?:SEQUENCE|FILTER|SORT|UNIQUE|TRANSPOSE)\s*\(/i.test(cell.formula ?? "");
        if(wantsSpill) {
          const outputs=info.sheet.setSpill(cell,rows,this.write);
          this.changed.push(...outputs);
          if(this.write) for(const output of outputs) for(const reader of this.readers(info.sheet,output.row,output.col)) {
            if(!this.computing.has(reader.cell) && !this.done.has(reader.cell)) this.pending.add(reader.cell);
          }
          value=cell.value;
        } else value=rows[0]?.[0] ?? 0;
      } else value = info.node ? evaluate(info.node, this.context(info.sheet, cell.row, cell.col)) : null;
      this.circular.delete(cell);
    } catch (e) {
      if (e instanceof Unsupported) {
        unsupported = true;
        value = null;
      } else if (e instanceof Circular) {
        this.circular.add(cell);
        value = { error: "#REF!" };
        if (e.cell !== cell) {
          this.store(cell, value);
          this.computing.delete(cell);
          this.readStale = outerStale;
          throw e;
        }
      } else {
        this.computing.delete(cell);
        throw e;
      }
    }
    this.computing.delete(cell);
    if (unsupported || this.readStale) this.stale.add(cell);
    else this.stale.delete(cell);
    if (!unsupported) this.store(cell, value);
    this.readStale = outerStale;
  }

  private store(cell: Cell, value: Scalar): void {
    if (same(cell.value, value) && !cell.noCache) return;
    if (this.write) {
      const sheet = this.infos.get(cell)!.sheet;
      sheet.setCached(cell, value);
      this.changed.push(cell);
    } else cell.value = value;
  }

  private run(cells: Iterable<Cell>, write: boolean): Cell[] {
    this.write = write;
    this.changed = [];
    for (const c of cells) this.pending.add(c);
    while(this.pending.size) this.compute(this.pending.values().next().value!);
    this.pending.clear();
    this.done.clear();
    return this.changed;
  }

  /** Every formula cell, as changes (to recalculate everything after a structural edit). */
  formulaCells(): Change[] {
    return [...this.infos.values()].map((i) => ({ sheet: i.sheet, row: i.cell.row, col: i.cell.col }));
  }

  /** Show results for formulas the file has no cached value for, in memory only (opening must not change the file). */
  computeMissing(): void {
    const missing = [...this.infos.values()].filter((i) => i.cell.noCache).map((i) => i.cell);
    this.run(missing, false);
  }

  /** Recalculate after cells changed; returns the formula cells whose value changed (and were written). */
  recalc(changes: Change[]): Cell[] {
    let formulasChanged = false;
    for (const ch of changes) {
      const cell = ch.sheet.at(ch.row, ch.col);
      // Drop index entries of cells that were cleared or replaced.
      const old = this.byPos.get(ch.sheet)?.get(ch.row * 16384 + ch.col);
      if (old && old !== cell) {
        this.unindex(old);
        formulasChanged = true;
      }
      if (cell && (cell.formula !== undefined || this.infos.has(cell))) {
        formulasChanged = true;
        this.index(ch.sheet, cell);
      }
    }
    if (formulasChanged) this.wb.pkg.dropCalcChain();

    // Everything reading a changed cell, transitively.
    const dirty = new Set<Cell>();
    const queue: Change[] = [...changes];
    for (const ch of changes) {
      const cell = ch.sheet.at(ch.row, ch.col);
      if (cell && this.infos.has(cell)) dirty.add(cell);
    }
    for (const info of this.always) {
      dirty.add(info.cell);
      queue.push({ sheet: info.sheet, row: info.cell.row, col: info.cell.col });
    }
    while (queue.length) {
      const ch = queue.pop()!;
      for (const info of this.readers(ch.sheet, ch.row, ch.col)) {
        if (dirty.has(info.cell)) continue;
        dirty.add(info.cell);
        queue.push({ sheet: info.sheet, row: info.cell.row, col: info.cell.col });
      }
    }
    const changed = this.run(dirty, true);
    if (this.stale.size) this.requestFullCalc();
    return changed;
  }

  /** Ask Excel to recalculate everything when it opens the file (our results for some formulas are missing). */
  private requestFullCalc(): void {
    const doc = this.wb.pkg.workbook;
    let calcPr = doc.getElementsByTagNameNS(NS_S, "calcPr")[0];
    if (!calcPr) {
      calcPr = doc.createElementNS(NS_S, "calcPr");
      // calcPr follows definedNames / sheets; before oleSize, customWorkbookViews, pivotCaches, ...
      const after = doc.getElementsByTagNameNS(NS_S, "definedNames")[0] ?? doc.getElementsByTagNameNS(NS_S, "sheets")[0];
      after.parentNode!.insertBefore(calcPr, after.nextSibling);
    }
    if (calcPr.getAttribute("fullCalcOnLoad") === "1") return;
    calcPr.setAttribute("fullCalcOnLoad", "1");
    this.wb.pkg.touch(this.wb.pkg.workbookPath);
  }
}
