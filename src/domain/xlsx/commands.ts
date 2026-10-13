import { assertAction, isProtected } from "./protection";
/*
 * Edits as undoable transactions. While one runs, the workbook journals each cell element before it changes (and
 * whole parts for structural edits); undo puts those back and recalculates. Shared strings and style records are
 * append-only, so restored cells can still point at them.
 */
import { XlsxWorkbook } from "../../infrastructure/xlsx/xlsx-workbook";
import { parseXml } from "../../infrastructure/ooxml/ooxml-package";
import { Calc, Change } from "./calculation";
import { parseFormula, tokenize, translateFormula } from "./formula";
import { parseInput, Scalar } from "./number-format";
import { Rect } from "./references";
import { StyleChange } from "./styles";
import { Journal, Sheet, Workbook } from "./workbook-model";

interface Saved {
  sheet: Sheet;
  row: number;
  col: number;
  before: Element | null;
  after?: Element | null;
}

class Transaction implements Journal {
  cells = new Map<string, Saved>();
  parts = new Map<string, { before: string | null; after?: string | null }>();

  constructor(private wb: Workbook) {}

  cell(sheet: Sheet, row: number, col: number): void {
    const id = `${sheet.ref.path}|${row}|${col}`;
    // A part saved whole is restored whole; its cells needn't be saved one by one.
    if (this.cells.has(id) || this.parts.has(sheet.ref.path)) return;
    const el = sheet.at(row, col)?.el;
    this.cells.set(id, { sheet, row, col, before: el ? (el.cloneNode(true) as Element) : null });
  }

  part(path: string): void {
    if (this.parts.has(path)) return;
    const doc = this.wb.pkg.xml(path);
    this.parts.set(path, { before: doc ? new XMLSerializer().serializeToString(doc) : null });
  }

  get empty(): boolean {
    return !this.cells.size && !this.parts.size;
  }
}

/** Largest area a formatting change creates cells for; bigger selections format their used part. */
const MAX_FORMAT_CELLS = 100000;

export class Edits {
  calc: Calc;
  private undoStack: Transaction[] = [];
  private tx: Transaction | null = null;
  private redoStack: Transaction[] = [];
  /** Called after an edit, undo or redo with the sheets whose structure was reloaded. */
  onReload: (() => void) | null = null;

  constructor(public wb: Workbook) {
    this.calc = new Calc(wb);
    this.calc.computeMissing();
  }

  /** Run `fn` as one undoable step; it returns the cells it changed, which are then recalculated. */
  run(fn: () => Change[]): void {
    if (this.tx) {
      // Nested: part of the running edit.
      const changes = fn();
      if (changes.length) this.calc.recalc(changes);
      return;
    }
    const tx = new Transaction(this.wb);
    this.tx = tx;
    this.wb.journal = tx;
    try {
      const changes = fn();
      if (changes.length) this.calc.recalc(changes);
    } catch (err) {
      this.wb.journal = null;
      this.tx = null;
      if (!tx.empty) this.apply(tx, "before");
      throw err;
    } finally {
      this.wb.journal = null;
      this.tx = null;
    }
    if (tx.empty) return;
    this.undoStack.push(tx);
    if (this.undoStack.length > 100) this.undoStack.shift();
    this.redoStack = [];
  }

  /**
   * Run a structural edit on the package XML: every XML part except the append-only string and style tables is saved
   * whole, the workbook model is rebuilt, `after` runs on the rebuilt model, and all formulas are recalculated.
   */
  structural(fn: () => void, after?: () => void): void {
    const pkg = this.wb.pkg;
    const before = new Set(pkg.xmlPaths());
    this.run(() => {
      for (const p of before) if (p !== pkg.stringsPath && p !== pkg.stylesPath) this.tx!.part(p);
      try { fn(); }
      finally {
        for (const p of pkg.xmlPaths()) if (!before.has(p)) this.tx!.parts.set(p, { before: null });
      }
      this.reload();
      this.wb.journal = this.tx;
      after?.();
      return this.calc.formulaCells();
    });
    this.onReload?.();
  }

  private reload(): void {
    this.wb = Workbook.open(this.wb.pkg);
    this.calc = new Calc(this.wb);
    this.calc.computeMissing();
  }

  get canUndo(): boolean {
    return this.undoStack.length > 0;
  }
  get canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  undo(): boolean {
    const tx = this.undoStack.pop();
    if (!tx) return false;
    this.apply(tx, "before");
    this.redoStack.push(tx);
    return true;
  }

  redo(): boolean {
    const tx = this.redoStack.pop();
    if (!tx) return false;
    this.apply(tx, "after");
    this.undoStack.push(tx);
    return true;
  }

  private apply(tx: Transaction, which: "before" | "after"): void {
    const pkg: XlsxWorkbook = this.wb.pkg;
    if (tx.parts.size) {
      for (const [path, p] of tx.parts) {
        if (which === "before" && p.after === undefined) {
          const doc = pkg.xml(path);
          p.after = doc ? new XMLSerializer().serializeToString(doc) : null;
        }
        const text = p[which];
        if (text) pkg.putXml(path, parseXml(text)!);
        else pkg.removeXml(path);
      }
    }
    // Sheets are looked up again by path: a structural undo above may have rebuilt them.
    const sheetOf = (s: Sheet) => this.wb.sheets.find((x) => x.ref.path === s.ref.path) ?? s;
    const touched = new Set<Sheet>();
    const changes: Change[] = [];
    for (const saved of tx.cells.values()) {
      const sheet = sheetOf(saved.sheet);
      if (which === "before" && saved.after === undefined) {
        const el = sheet.at(saved.row, saved.col)?.el;
        saved.after = el ? (el.cloneNode(true) as Element) : null;
      }
    }
    if (tx.parts.size) this.wb = Workbook.open(pkg);
    for (const saved of tx.cells.values()) {
      const sheet = this.wb.sheets.find((x) => x.ref.path === saved.sheet.ref.path);
      if (!sheet) continue;
      sheet.restoreCell(saved.row, saved.col, saved[which] ?? null);
      touched.add(sheet);
      changes.push({ sheet, row: saved.row, col: saved.col });
    }
    for (const s of touched) s.resolveShared();
    if (tx.parts.size) {
      this.calc = new Calc(this.wb);
      this.calc.computeMissing();
      this.onReload?.();
    } else if (changes.length) {
      // Restored formula caches are the old results; recalculating just re-confirms them (no new journal).
      this.calc.recalc(changes);
    }
  }

  // ---------------------------------------------------------------------------
  // Cell edits
  // ---------------------------------------------------------------------------

  /** A formula as typed ("=sum(a1" -> "sum(a1)"): checked, with missing closing parentheses added. */
  static normalizeFormula(text: string): string {
    let f = text.replace(/^=/, "");
    let depth = 0;
    try {
      for (const t of tokenize(f)) if (t.kind === "punct") depth += t.text === "(" ? 1 : t.text === ")" ? -1 : 0;
    } catch {
      // parseFormula below reports it
    }
    if (depth > 0) f += ")".repeat(depth);
    parseFormula(f);
    const tokens = tokenize(f);
    return tokens.map((t, i) => {
      if (t.text === "#" && tokens[i - 1]?.kind === "ref") return "";
      if (t.kind === "ref" && tokens[i + 1]?.text === "#") return `_xlfn.ANCHORARRAY(${t.text})`;
      if (t.kind === "func" && !t.text.startsWith("_")) {
        const name=t.text.toUpperCase();
        if (["FILTER","SORT"].includes(name)) return `_xlfn._xlws.${name}`;
        if (["SEQUENCE","UNIQUE"].includes(name)) return `_xlfn.${name}`;
      }
      return t.text;
    }).join("");
  }

  /** What typing `raw` into a cell writes (without applying it). */
  private inputOf(sheet: Sheet, row: number, col: number, raw: string): { value: Scalar; formula?: string; numFmt?: string } {
    if (raw.startsWith("=") && raw.length > 1) return { value: null, formula: Edits.normalizeFormula(raw) };
    const cell = sheet.at(row, col);
    const fmt = this.wb.styles.numFmt(cell?.style ?? 0);
    if (fmt === "@") return { value: raw === "" ? null : raw };
    const p = parseInput(raw, this.wb.date1904);
    return { value: p.value, numFmt: p.format && fmt === "General" ? p.format : undefined };
  }

  private write(sheet: Sheet, row: number, col: number, v: { value: Scalar; formula?: string; numFmt?: string; style?: number }): void {
    let style = v.style;
    if (v.numFmt) {
      style = this.wb.styles.derive(style ?? sheet.at(row, col)?.style ?? 0, { numFmt: v.numFmt });
      this.wb.touchStyles();
    }
    sheet.setCell(row, col, v.formula !== undefined ? { formula: v.formula, value: null, style } : { value: v.value, formula: null, style });
  }

  /** Enter typed text into a cell. Throws FormulaSyntaxError for malformed formulas (nothing changes). */
  input(sheet: Sheet, row: number, col: number, raw: string): void {
    const v = this.inputOf(sheet, row, col, raw);
    this.run(() => {
      this.write(sheet, row, col, v);
      return [{ sheet, row, col }];
    });
  }

  /** Write typed text inside a running edit (see `run`). */
  writeInput(sheet: Sheet, row: number, col: number, raw: string): void {
    this.write(sheet, row, col, this.inputOf(sheet, row, col, raw));
  }

  /** Delete contents (values and formulas) in a range, keeping formatting. */
  clear(sheet: Sheet, r: Rect): void {
    this.run(() => {
      const changes: Change[] = [];
      for (const c of sheet.cellsIn(r)) {
        if (c.value === null && c.formula === undefined) continue;
        if (c.array && (c.array.r1 !== c.row || c.array.c1 !== c.col) && !(c.array.r1 >= r.r1 && c.array.c1 >= r.c1)) continue;
        changes.push({ sheet, row: c.row, col: c.col });
        sheet.setCell(c.row, c.col, { value: null, formula: null });
      }
      return changes;
    });
  }

  /** Clip huge selections (whole columns / rows) to the used area so formatting doesn't create a million cells. */
  private formatArea(sheet: Sheet, r: Rect): Rect | null {
    if ((r.r2 - r.r1 + 1) * (r.c2 - r.c1 + 1) <= MAX_FORMAT_CELLS) return r;
    const used = sheet.used();
    if (!used) return null;
    // ponytail: whole-row/column formats aren't written as <row>/<col> styles; only the used cells get them.
    const clipped = { r1: r.r1, c1: r.c1, r2: Math.min(r.r2, used.r2), c2: Math.min(r.c2, used.c2) };
    return clipped.r1 <= clipped.r2 && clipped.c1 <= clipped.c2 ? clipped : null;
  }

  /** Apply a style change to every cell of a range. */
  format(sheet: Sheet, r: Rect, change: StyleChange): void {
    assertAction(sheet, "formatCells");
    if (change.locked !== undefined && isProtected(sheet)) throw new Error("Unprotect the sheet before changing cell locks.");
    const area = this.formatArea(sheet, r);
    if (!area) return;
    this.run(() => {
      const derived = new Map<number, number>();
      for (let row = area.r1; row <= area.r2; row++) {
        for (let col = area.c1; col <= area.c2; col++) {
          const from = sheet.at(row, col)?.style ?? 0;
          if (!derived.has(from)) derived.set(from, this.wb.styles.derive(from, change));
          const to = derived.get(from)!;
          if (to !== from) sheet.setCell(row, col, { style: to });
        }
      }
      this.wb.touchStyles();
      return [];
    });
  }

  /** Set each cell's style: paint format tiles the copied styles over the range, clear formatting sets the default (0). */
  restyle(sheet: Sheet, r: Rect, styleAt: (row: number, col: number) => number): void {
    assertAction(sheet, "formatCells");
    const area = this.formatArea(sheet, r);
    if (!area) return;
    this.run(() => {
      for (let row = area.r1; row <= area.r2; row++) {
        for (let col = area.c1; col <= area.c2; col++) {
          const to = styleAt(row, col);
          if ((sheet.at(row, col)?.style ?? 0) !== to) sheet.setCell(row, col, { style: to });
        }
      }
      return [];
    });
  }

  /** The value a toggle (bold, italic, ...) should get: the opposite of the range's first cell. */
  toggled(sheet: Sheet, r: Rect, key: "bold" | "italic" | "underline" | "strike" | "wrap"): boolean {
    return !this.wb.styles.format(sheet.at(r.r1, r.c1)?.style ?? 0)[key];
  }

  /** Put cell contents (from a copy) at a position: formulas move by the offset, styles come along. */
  place(sheet: Sheet, row: number, col: number, src: { value: Scalar; formula?: string; style: number; row: number; col: number }, styles = true): Change {
    const formula = src.formula !== undefined ? translateFormula(src.formula, row - src.row, col - src.col) : undefined;
    sheet.setCell(row, col, formula !== undefined ? { formula, value: null, style: styles ? src.style : undefined } : { value: src.value, formula: null, style: styles ? src.style : undefined });
    return { sheet, row, col };
  }

  /**
   * Fill `target` from `src` (a corner of it), like dragging the fill handle: numbers in a run of 2+ continue their
   * linear series, everything else (and single numbers) repeats; formulas move their references.
   */
  fill(sheet: Sheet, src: Rect, target: Rect): void {
    const down = target.r2 > src.r2 || target.r1 < src.r1;
    const len = down ? src.r2 - src.r1 + 1 : src.c2 - src.c1 + 1;
    const lines = down ? src.c2 - src.c1 + 1 : src.r2 - src.r1 + 1;
    this.run(() => {
      const changes: Change[] = [];
      for (let k = 0; k < lines; k++) {
        const at = (i: number) => (down ? { row: src.r1 + i, col: src.c1 + k } : { row: src.r1 + k, col: src.c1 + i });
        const cells = Array.from({ length: len }, (_, i) => sheet.at(at(i).row, at(i).col));
        const nums = cells.every((c) => c && c.formula === undefined && typeof c.value === "number") && len > 1;
        const step = nums ? linearStep(cells.map((c) => c!.value as number)) : null;
        const [from, to] = down ? [target.r1, target.r2] : [target.c1, target.c2];
        const start = down ? src.r1 : src.c1;
        for (let p = from; p <= to; p++) {
          const i = p - start;
          if (i >= 0 && i < len) continue;
          const row = down ? p : src.r1 + k, col = down ? src.c1 + k : p;
          const j = ((i % len) + len) % len;
          const c = cells[j];
          if (step && c) {
            const base = cells[0]!.value as number;
            sheet.setCell(row, col, { value: base + step * i, formula: null, style: c.style });
          } else if (c) this.place(sheet, row, col, c);
          else {
            const cur = sheet.at(row, col);
            if (cur) sheet.setCell(row, col, { value: null, formula: null, style: 0 });
          }
          changes.push({ sheet, row, col });
        }
      }
      return changes;
    });
  }
}

/** Common difference of an arithmetic run (least squares, like spreadsheet fill). */
function linearStep(ns: number[]): number {
  const n = ns.length;
  const mx = (n - 1) / 2;
  const my = ns.reduce((a, b) => a + b, 0) / n;
  let num = 0, den = 0;
  ns.forEach((y, x) => {
    num += (x - mx) * (y - my);
    den += (x - mx) ** 2;
  });
  return num / den;
}
