import { assertAction } from "./protection";
/*
 * Structural edits: rows / columns, sheets, sizes, merges, panes, sorting and filters. Inserting or deleting rows
 * and columns rewrites every reference that can point at them (formulas on all sheets, names, merges, conditional
 * formats, validations, hyperlinks, filters, tables, drawing anchors and chart ranges) in one undo step.
 */
import { NS_S } from "../../infrastructure/xlsx/xlsx-workbook";
import { parseXml } from "../../infrastructure/ooxml/ooxml-package";
import { Change } from "./calculation";
import { Edits } from "./commands";
import { ShiftOptions, renameSheetInFormula, shiftFormula } from "./formula";
import { compareValues } from "./formula-evaluator";
import { Rect, cellName, colName, quoteSheet, rangeName } from "./references";
import { Scalar } from "./number-format";
import { Sheet } from "./workbook-model";

const XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const C = "http://schemas.openxmlformats.org/drawingml/2006/chart";

/** CT_Worksheet child order, for inserting new elements where Excel expects them. */
export const WORKSHEET_ORDER = [
  "sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetCalcPr", "sheetProtection", "protectedRanges",
  "scenarios", "autoFilter", "sortState", "dataConsolidate", "customSheetViews", "mergeCells", "phoneticPr", "conditionalFormatting",
  "dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks",
  "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing", "legacyDrawing", "legacyDrawingHF", "drawingHF",
  "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst",
];

export function child(sheet: Sheet, tag: string, create = false): Element | null {
  const root = sheet.doc.documentElement;
  let el = sheet.first(root, tag);
  if (el || !create) return el;
  el = sheet.doc.createElementNS(NS_S, tag);
  const rank = WORKSHEET_ORDER.indexOf(tag);
  const next = Array.from(root.children).find((c) => WORKSHEET_ORDER.indexOf(c.localName) > rank) ?? null;
  root.insertBefore(el, next);
  return el;
}

const XM = "http://schemas.microsoft.com/office/excel/2006/main";
const sheetEls = (e: Edits) => Array.from(e.wb.pkg.workbook.getElementsByTagNameNS(NS_S, "sheet"));
const definedNames = (e: Edits) => Array.from(e.wb.pkg.workbook.getElementsByTagNameNS(NS_S, "definedName"));

/** Apply `fn` to every formula text in the workbook: cells, names, validations, conditional formats and charts. */
export function rewriteFormulas(e: Edits, fn: (text: string, sheet: Sheet | null) => string): void {
  for (const sheet of e.wb.sheets) {
    for (const f of Array.from(sheet.doc.getElementsByTagNameNS(NS_S, "f"))) if (f.textContent) f.textContent = fn(f.textContent, sheet);
    for (const tag of ["formula", "formula1", "formula2"]) {
      for (const f of Array.from(sheet.doc.getElementsByTagNameNS(NS_S, tag))) if (f.textContent) f.textContent = fn(f.textContent, sheet);
    }
    // Excel 2010 extensions: sparkline sources and conditional format points.
    for (const f of Array.from(sheet.doc.getElementsByTagNameNS(XM, "f"))) if (f.textContent) f.textContent = fn(f.textContent, sheet);
    sheet.touch();
  }
  for (const n of definedNames(e)) n.textContent = fn(n.textContent ?? "", null);
  e.wb.pkg.touch(e.wb.pkg.workbookPath);
  for (const path of e.wb.pkg.xmlPaths()) {
    const doc = e.wb.pkg.xml(path)!;
    if (doc.documentElement.namespaceURI !== C) continue;
    for (const f of Array.from(doc.getElementsByTagNameNS(C, "f"))) f.textContent = fn(f.textContent ?? "", null);
    e.wb.pkg.touch(path);
  }
}

/** Shift a space-separated list of ranges; ranges that were deleted entirely drop out. */
function shiftSqref(sqref: string, o: ShiftOptions): string {
  return sqref
    .split(/\s+/)
    .filter(Boolean)
    .map((r) => shiftFormula(r, o))
    .filter((r) => !r.includes("#REF!"))
    .join(" ");
}

function shiftAttr(el: Element, attr: string, o: ShiftOptions): boolean {
  const v = el.getAttribute(attr);
  if (v === null) return true;
  const out = shiftSqref(v, o);
  if (!out) return false;
  el.setAttribute(attr, out);
  return true;
}

/** Sheet-level XML for rows / columns inserted (count > 0) or deleted (count < 0) at `at`. */
function shiftSheet(e: Edits, sheet: Sheet, axis: "row" | "col", at: number, count: number): void {
  const o: ShiftOptions = { sheet: sheet.name, axis, at, count, formulaSheet: sheet.name };
  const n = -count, end = at + n - 1;
  const moved = (i: number) => (count > 0 ? (i >= at ? i + count : i) : i < at ? i : i > end ? i - n : null);
  const data = sheet.sheetData();

  if (axis === "row") {
    for (const row of sheet.all(data, "row")) {
      const r = Number(row.getAttribute("r")) - 1;
      const to = moved(r);
      if (to === null) {
        row.remove();
        continue;
      }
      if (to === r) continue;
      row.setAttribute("r", String(to + 1));
      for (const c of sheet.all(row, "c")) c.setAttribute("r", (c.getAttribute("r") ?? "").replace(/\d+$/, String(to + 1)));
    }
  } else {
    for (const row of sheet.all(data, "row")) {
      row.removeAttribute("spans");
      for (const c of sheet.all(row, "c")) {
        const m = /^([A-Z]+)(\d+)$/.exec(c.getAttribute("r") ?? "");
        if (!m) continue;
        let col = 0;
        for (const ch of m[1]) col = col * 26 + ch.charCodeAt(0) - 64;
        const to = moved(col - 1);
        if (to === null) c.remove();
        else c.setAttribute("r", colName(to) + m[2]);
      }
    }
    const cols = child(sheet, "cols");
    for (const col of sheet.all(cols, "col")) {
      const min = Number(col.getAttribute("min")) - 1, max = Number(col.getAttribute("max")) - 1;
      let a: number | null, b: number | null;
      if (count > 0) [a, b] = [min >= at ? min + count : min, max >= at ? Math.min(max + count, 16383) : max];
      else {
        a = moved(min) ?? at;
        b = moved(max) ?? at - 1;
      }
      if (b < a || a > 16383) col.remove();
      else {
        col.setAttribute("min", String(a + 1));
        col.setAttribute("max", String(b + 1));
      }
    }
    if (cols && !cols.children.length) cols.remove();
  }

  const root = sheet.doc.documentElement;
  for (const f of Array.from(sheet.doc.getElementsByTagNameNS(NS_S, "f"))) if (f.hasAttribute("ref")) shiftAttr(f, "ref", o);
  const dim = child(sheet, "dimension");
  if (dim && !shiftAttr(dim, "ref", o)) dim.setAttribute("ref", "A1");
  for (const [parent, tag, attr] of [["mergeCells", "mergeCell", "ref"], ["hyperlinks", "hyperlink", "ref"], ["dataValidations", "dataValidation", "sqref"]] as const) {
    const p = child(sheet, parent);
    for (const el of sheet.all(p, tag)) if (!shiftAttr(el, attr, o)) el.remove();
    if (p && !sheet.all(p, tag).length) p.remove();
    else p?.setAttribute("count", String(sheet.all(p, tag).length));
  }
  for (const cf of sheet.all(root, "conditionalFormatting")) if (!shiftAttr(cf, "sqref", o)) cf.remove();
  // Excel 2010 ranges (sparkline cells, conditional formats): what loses its range goes, with emptied wrappers.
  for (const ref of Array.from(sheet.doc.getElementsByTagNameNS(XM, "sqref"))) {
    const out = shiftSqref(ref.textContent ?? "", o);
    if (out) ref.textContent = out;
    else for (let gone: Element | null = ref.parentElement; gone && gone !== root; ) {
      const parent: Element | null = gone.parentElement;
      gone.remove();
      // A sparkline group without sparklines goes too (its colours alone are not a group).
      gone = parent && parent !== root && (!parent.children.length || gone.localName === "sparklines") ? parent : null;
    }
  }
  const filter = child(sheet, "autoFilter");
  if (filter && !shiftAttr(filter, "ref", o)) filter.remove();
  for (const sel of Array.from(sheet.doc.getElementsByTagNameNS(NS_S, "selection"))) {
    sel.removeAttribute("sqref");
    sel.removeAttribute("activeCell");
  }

  // Tables and drawings that belong to this sheet.
  for (const rel of e.wb.pkg.rels(sheet.ref.path)) {
    const doc = rel.external ? null : e.wb.pkg.xml(rel.target);
    if (!doc) continue;
    if (rel.type === "table") {
      for (const el of [doc.documentElement, ...Array.from(doc.getElementsByTagNameNS(NS_S, "autoFilter"))]) shiftAttr(el, "ref", o);
      e.wb.pkg.touch(rel.target);
    } else if (rel.type === "threadedComment") {
      for (const c of Array.from(doc.getElementsByTagNameNS("http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments", "threadedComment"))) if (!shiftAttr(c, "ref", o)) c.remove();
      e.wb.pkg.touch(rel.target);
    } else if (rel.type === "comments") {
      // Notes move with their cells; ones on deleted cells go (their VML drawing is rebuilt on save).
      for (const c of Array.from(doc.getElementsByTagNameNS(NS_S, "comment"))) if (!shiftAttr(c, "ref", o)) c.remove();
      e.wb.pkg.touch(rel.target);
      if (!doc.getElementsByTagNameNS(NS_S, "comment").length) e.wb.pkg.removeComments(sheet.ref.path);
    } else if (rel.type === "drawing") {
      const tag = axis === "row" ? "row" : "col";
      for (const el of Array.from(doc.getElementsByTagNameNS(XDR, tag))) {
        const i = Number(el.textContent);
        el.textContent = String(moved(i) ?? at);
      }
      e.wb.pkg.touch(rel.target);
    }
  }
  sheet.touch();
}

function insertOrDelete(e: Edits, sheet: Sheet, axis: "row" | "col", at: number, count: number): void {
  assertAction(sheet, axis === "row" ? (count > 0 ? "insertRows" : "deleteRows") : (count > 0 ? "insertColumns" : "deleteColumns"));
  const name = sheet.name, path = sheet.ref.path;
  e.structural(
    () => {
      for (const s of e.wb.sheets) s.unshareAll();
      rewriteFormulas(e, (text, owner) => shiftFormula(text, { sheet: name, axis, at, count, formulaSheet: owner?.name ?? "\u0000" }));
      shiftSheet(e, e.wb.sheets.find((s) => s.ref.path === path)!, axis, at, count);
      e.wb.pkg.dropCalcChain();
    },
    () => {
      // New rows take the formats of the row above, new columns those of the column to the left.
      if (count <= 0 || at === 0) return;
      const s = e.wb.sheets.find((x) => x.ref.path === path)!;
      const src = axis === "row" ? { r1: at - 1, c1: 0, r2: at - 1, c2: 16383 } : { r1: 0, c1: at - 1, r2: 1048575, c2: at - 1 };
      for (const c of s.cellsIn(src)) {
        if (!c.style) continue;
        for (let k = 0; k < count; k++) s.setCell(axis === "row" ? at + k : c.row, axis === "row" ? c.col : at + k, { style: c.style });
      }
    }
  );
}

export const insertRows = (e: Edits, sheet: Sheet, at: number, count: number) => insertOrDelete(e, sheet, "row", at, count);
export const deleteRows = (e: Edits, sheet: Sheet, at: number, count: number) => insertOrDelete(e, sheet, "row", at, -count);
export const insertCols = (e: Edits, sheet: Sheet, at: number, count: number) => insertOrDelete(e, sheet, "col", at, count);
export const deleteCols = (e: Edits, sheet: Sheet, at: number, count: number) => insertOrDelete(e, sheet, "col", at, -count);

// ---------------------------------------------------------------------------
// Sheets
// ---------------------------------------------------------------------------

function uniqueName(e: Edits, base: (n: number) => string, start: number): string {
  for (let n = start; ; n++) if (!e.wb.sheet(base(n))) return base(n);
}

export function validateSheetName(e: Edits, name: string, except?: Sheet): void {
  if (!name.trim()) throw new Error("A sheet name can't be empty.");
  if (name.length > 31) throw new Error("A sheet name can be at most 31 characters.");
  if (/[\\/?*[\]:]/.test(name)) throw new Error("A sheet name can't contain the character \\ / ? * [ ] or :.");
  if (/^'|'$/.test(name)) throw new Error("A sheet name can't start or end with an apostrophe.");
  const other = e.wb.sheet(name);
  if (other && other !== except) throw new Error(`A sheet named "${name}" already exists.`);
}

/** Add an empty sheet after tab `after`, named "SheetN" or `base` ("base (2)" when taken); returns its name. */
export function addSheet(e: Edits, after: number, base?: string): string {
  const name = base ? uniqueName(e, (n) => (n === 1 ? base : `${base.slice(0, 28 - String(n).length)} (${n})`), 1) : uniqueName(e, (n) => `Sheet${n}`, e.wb.sheets.length + 1);
  e.structural(() => {
    e.wb.pkg.addSheet(name, `<worksheet xmlns="${NS_S}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetFormatPr defaultRowHeight="15"/><sheetData/></worksheet>`, after);
    renumberLocalNames(e, (i) => (i > after ? i + 1 : i));
  });
  return name;
}

/** definedName localSheetId values after tabs moved (`map` old index -> new, null = removed). */
function renumberLocalNames(e: Edits, map: (i: number) => number | null): void {
  for (const n of definedNames(e)) {
    const id = n.getAttribute("localSheetId");
    if (id === null) continue;
    const to = map(Number(id));
    if (to === null) n.remove();
    else n.setAttribute("localSheetId", String(to));
  }
  const names = e.wb.pkg.workbook.getElementsByTagNameNS(NS_S, "definedNames")[0];
  if (names && !names.children.length) names.remove();
}

function setActiveTab(e: Edits, index: number): void {
  const view = e.wb.pkg.workbook.getElementsByTagNameNS(NS_S, "workbookView")[0];
  if (!view) return;
  view.setAttribute("activeTab", String(index));
  view.removeAttribute("firstSheet");
}

export function renameSheet(e: Edits, sheet: Sheet, name: string): void {
  name = name.trim();
  validateSheetName(e, name, sheet);
  if (name === sheet.name) return;
  const old = sheet.name;
  e.structural(() => {
    rewriteFormulas(e, (text) => renameSheetInFormula(text, old, name));
    sheetEls(e).find((el) => el.getAttribute("name") === old)!.setAttribute("name", name);
  });
}

export function deleteSheet(e: Edits, sheet: Sheet): void {
  const sheets = e.wb.sheets;
  if (sheets.filter((s) => s.ref.state === "visible" && s !== sheet).length === 0) throw new Error("A workbook needs at least one visible sheet; you can't delete the last one.");
  const index = sheets.indexOf(sheet);
  e.structural(() => {
    rewriteFormulas(e, (text) => renameSheetInFormula(text, sheet.name, null));
    renumberLocalNames(e, (i) => (i === index ? null : i > index ? i - 1 : i));
    e.wb.pkg.deleteSheet(e.wb.pkg.sheets()[index]);
    setActiveTab(e, Math.max(0, index - 1));
  });
}

export function moveSheet(e: Edits, from: number, to: number): void {
  if (from === to) return;
  e.structural(() => {
    const els = sheetEls(e);
    const el = els[from];
    const parent = el.parentNode!;
    el.remove();
    const rest = sheetEls(e);
    parent.insertBefore(el, rest[to] ?? null);
    renumberLocalNames(e, (i) => (i === from ? to : from < to && i > from && i <= to ? i - 1 : from > to && i >= to && i < from ? i + 1 : i));
    setActiveTab(e, to);
  });
}

export function duplicateSheet(e: Edits, sheet: Sheet): string {
  const name = uniqueName(e, (n) => (n === 1 ? `Copy of ${sheet.name}` : `Copy ${n} of ${sheet.name}`), 1);
  const index = e.wb.sheets.indexOf(sheet);
  e.structural(() => {
    const doc = parseXml(new XMLSerializer().serializeToString(sheet.doc))!;
    // Parts the copy would share by relationship (drawings, comments, tables) aren't copied.
    for (const tag of ["drawing", "legacyDrawing", "legacyDrawingHF", "picture", "oleObjects", "controls", "tableParts"]) {
      for (const el of Array.from(doc.getElementsByTagNameNS(NS_S, tag))) el.remove();
    }
    for (const v of Array.from(doc.getElementsByTagNameNS(NS_S, "sheetView"))) v.removeAttribute("tabSelected");
    e.wb.pkg.addSheet(name, new XMLSerializer().serializeToString(doc), index);
    renumberLocalNames(e, (i) => (i > index ? i + 1 : i));
  });
  return name;
}

export function setSheetHidden(e: Edits, sheet: Sheet, hidden: boolean): void {
  if (hidden && e.wb.sheets.filter((s) => s.ref.state === "visible" && s !== sheet).length === 0) throw new Error("At least one sheet must stay visible.");
  e.structural(() => {
    const el = sheetEls(e).find((x) => x.getAttribute("name") === sheet.name)!;
    if (hidden) el.setAttribute("state", "hidden");
    else el.removeAttribute("state");
    const firstVisible = e.wb.pkg.sheets().findIndex((s) => s.state === "visible");
    const view = e.wb.pkg.workbook.getElementsByTagNameNS(NS_S, "workbookView")[0];
    if (view && e.wb.pkg.sheets()[Number(view.getAttribute("activeTab") ?? 0)]?.state !== "visible") setActiveTab(e, firstVisible);
  });
}

// ---------------------------------------------------------------------------
// Sizes, visibility, panes, merges
// ---------------------------------------------------------------------------

/** Rewrite the <col> ranges covering columns c1..c2 (0-based) with `fn` applied to each column's element. */
export function editCols(sheet: Sheet, c1: number, c2: number, fn: (col: Element) => void): void {
  const cols = child(sheet, "cols", true)!;
  const pieces: Element[] = [];
  const covered = new Set<number>();
  for (const col of sheet.all(cols, "col")) {
    const min = Number(col.getAttribute("min")) - 1, max = Number(col.getAttribute("max")) - 1;
    col.remove();
    const parts: [number, number][] = [[min, Math.min(max, c1 - 1)], [Math.max(min, c1), Math.min(max, c2)], [Math.max(min, c2 + 1), max]];
    parts.forEach(([a, b], k) => {
      if (a > b) return;
      const el = col.cloneNode(true) as Element;
      el.setAttribute("min", String(a + 1));
      el.setAttribute("max", String(b + 1));
      if (k === 1) {
        for (let i = a; i <= b; i++) covered.add(i);
        fn(el);
      }
      pieces.push(el);
    });
  }
  for (let i = c1; i <= c2; i++) {
    if (covered.has(i)) continue;
    let j = i;
    while (j + 1 <= c2 && !covered.has(j + 1)) j++;
    const el = sheet.doc.createElementNS(NS_S, "col");
    el.setAttribute("min", String(i + 1));
    el.setAttribute("max", String(j + 1));
    el.setAttribute("width", String(+(sheet.cols.size / 7).toFixed(4)));
    fn(el);
    pieces.push(el);
    i = j;
  }
  pieces.sort((a, b) => Number(a.getAttribute("min")) - Number(b.getAttribute("min")));
  for (const p of pieces) cols.appendChild(p);
  sheet.touch();
}

export function setColWidth(e: Edits, sheet: Sheet, c1: number, c2: number, px: number): void {
  assertAction(sheet, "formatColumns");
  e.structural(() =>
    editCols(sheet, c1, c2, (col) => {
      col.setAttribute("width", String(+(px / 7).toFixed(4)));
      col.setAttribute("customWidth", "1");
      col.removeAttribute("hidden");
    })
  );
}

export function setRowHeight(e: Edits, sheet: Sheet, r1: number, r2: number, px: number): void {
  assertAction(sheet, "formatRows");
  e.structural(() => {
    for (let r = r1; r <= r2; r++) {
      const row = sheet.rowEl(r, true)!;
      row.setAttribute("ht", String(+(px * 0.75).toFixed(2)));
      row.setAttribute("customHeight", "1");
      row.removeAttribute("hidden");
    }
    sheet.touch();
  });
}

export function setHidden(e: Edits, sheet: Sheet, axis: "row" | "col", a: number, b: number, hidden: boolean): void {
  assertAction(sheet, axis === "row" ? "formatRows" : "formatColumns");
  e.structural(() => {
    if (axis === "col") {
      editCols(sheet, a, b, (col) => {
        if (hidden) col.setAttribute("hidden", "1");
        else col.removeAttribute("hidden");
      });
      return;
    }
    for (let r = a; r <= b; r++) {
      const row = sheet.rowEl(r, hidden);
      if (!row) continue;
      if (hidden) row.setAttribute("hidden", "1");
      else row.removeAttribute("hidden");
    }
    sheet.touch();
  });
}

export function freeze(e: Edits, sheet: Sheet, rows: number, cols: number): void {
  assertAction(sheet, "sheet");
  e.structural(() => {
    const views = child(sheet, "sheetViews", true)!;
    let view = sheet.first(views, "sheetView");
    if (!view) {
      view = views.appendChild(sheet.doc.createElementNS(NS_S, "sheetView"));
      view.setAttribute("workbookViewId", "0");
    }
    for (const el of [...sheet.all(view, "pane"), ...sheet.all(view, "selection")]) el.remove();
    if (rows || cols) {
      const pane = sheet.doc.createElementNS(NS_S, "pane");
      if (cols) pane.setAttribute("xSplit", String(cols));
      if (rows) pane.setAttribute("ySplit", String(rows));
      pane.setAttribute("topLeftCell", cellName(rows, cols));
      pane.setAttribute("activePane", rows && cols ? "bottomRight" : rows ? "bottomLeft" : "topRight");
      pane.setAttribute("state", "frozen");
      view.insertBefore(pane, view.firstChild);
    }
    sheet.touch();
  });
}

export function merge(e: Edits, sheet: Sheet, r: Rect): void {
  assertAction(sheet, "formatCells");
  if (r.r1 === r.r2 && r.c1 === r.c2) return;
  e.structural(() => {
    for (const c of sheet.cellsIn(r)) if (c.row !== r.r1 || c.col !== r.c1) sheet.setCell(c.row, c.col, { value: null, formula: null });
    const merges = child(sheet, "mergeCells", true)!;
    for (const m of sheet.all(merges, "mergeCell")) {
      const x = sheet.merges.find((g) => rangeName(g) === m.getAttribute("ref"));
      if (x && x.r1 <= r.r2 && r.r1 <= x.r2 && x.c1 <= r.c2 && r.c1 <= x.c2) m.remove();
    }
    const el = merges.appendChild(sheet.doc.createElementNS(NS_S, "mergeCell"));
    el.setAttribute("ref", rangeName(r));
    merges.setAttribute("count", String(sheet.all(merges, "mergeCell").length));
    sheet.touch();
  });
}

/** Unmerge every merged range overlapping `r`. */
export function unmerge(e: Edits, sheet: Sheet, r: Rect): void {
  assertAction(sheet, "formatCells");
  const hits = sheet.merges.filter((x) => x.r1 <= r.r2 && r.r1 <= x.r2 && x.c1 <= r.c2 && r.c1 <= x.c2).map(rangeName);
  if (!hits.length) return;
  e.structural(() => {
    const merges = child(sheet, "mergeCells")!;
    for (const m of sheet.all(merges, "mergeCell")) if (hits.includes(m.getAttribute("ref") ?? "")) m.remove();
    if (!sheet.all(merges, "mergeCell").length) merges.remove();
    else merges.setAttribute("count", String(sheet.all(merges, "mergeCell").length));
    sheet.touch();
  });
}

// ---------------------------------------------------------------------------
// Sorting and filters
// ---------------------------------------------------------------------------

/** Sort order for values: numbers, text (case-insensitive), booleans, errors; blanks always last. */
function sortCompare(a: Scalar, b: Scalar, desc: boolean): number {
  const blank = (v: Scalar) => v === null || v === "";
  if (blank(a) || blank(b)) return Number(blank(a)) - Number(blank(b));
  const err = (v: Scalar) => typeof v === "object" && v !== null;
  if (err(a) || err(b)) return Number(err(a)) - Number(err(b));
  const d = compareValues(a, b);
  return desc ? -d : d;
}

/** Sort the rows of `r` by columns (absolute indexes); rows move with their formats and formulas. */
export function sortRange(e: Edits, sheet: Sheet, r: Rect, keys: { col: number; desc?: boolean }[]): void {
  assertAction(sheet, "sort");
  type Moved = { value: Scalar; formula?: string; style: number; row: number; col: number };
  const rows: { row: number; cells: (Moved | null)[]; keys: Scalar[] }[] = [];
  for (let row = r.r1; row <= r.r2; row++) {
    const cells: (Moved | null)[] = [];
    for (let col = r.c1; col <= r.c2; col++) {
      const c = sheet.at(row, col);
      cells.push(c ? { value: c.value, formula: c.formula, style: c.style, row, col } : null);
    }
    rows.push({ row, cells, keys: keys.map((k) => sheet.at(row, k.col)?.value ?? null) });
  }
  const sorted = [...rows].sort((a, b) => {
    for (let k = 0; k < keys.length; k++) {
      const d = sortCompare(a.keys[k], b.keys[k], !!keys[k].desc);
      if (d) return d;
    }
    return a.row - b.row;
  });
  if (sorted.every((x, i) => x === rows[i])) return;
  e.run(() => {
    const changes: Change[] = [];
    sorted.forEach((src, i) => {
      const row = r.r1 + i;
      src.cells.forEach((c, j) => {
        const col = r.c1 + j;
        if (c) changes.push(e.place(sheet, row, col, c));
        else if (sheet.at(row, col)) {
          sheet.setCell(row, col, { value: null, formula: null, style: 0 });
          changes.push({ sheet, row, col });
        }
      });
    });
    return changes;
  });
}

/** Turn the filter on for `r` (header row first), or off with null (showing every row again). */
export function setFilter(e: Edits, sheet: Sheet, r: Rect | null): void {
  assertAction(sheet, "sheet");
  e.structural(() => {
    const old = child(sheet, "autoFilter");
    if (old && sheet.filter) for (let row = sheet.filter.r1 + 1; row <= sheet.filter.r2; row++) sheet.rowEl(row, false)?.removeAttribute("hidden");
    old?.remove();
    const index = e.wb.sheets.indexOf(sheet);
    for (const n of definedNames(e)) if (n.getAttribute("name") === "_xlnm._FilterDatabase" && n.getAttribute("localSheetId") === String(index)) n.remove();
    if (r) {
      child(sheet, "autoFilter", true)!.setAttribute("ref", rangeName(r));
      let names = e.wb.pkg.workbook.getElementsByTagNameNS(NS_S, "definedNames")[0];
      if (!names) {
        names = e.wb.pkg.workbook.createElementNS(NS_S, "definedNames");
        const sheetsEl = e.wb.pkg.workbook.getElementsByTagNameNS(NS_S, "sheets")[0];
        sheetsEl.parentNode!.insertBefore(names, sheetsEl.nextSibling);
      }
      const n = names.appendChild(e.wb.pkg.workbook.createElementNS(NS_S, "definedName"));
      n.setAttribute("name", "_xlnm._FilterDatabase");
      n.setAttribute("localSheetId", String(index));
      n.setAttribute("hidden", "1");
      const abs = rangeName(r).replace(/([A-Z]+)(\d+)/g, "$$$1$$$2");
      n.textContent = `${quoteSheet(sheet.name)}!${abs}`;
    } else {
      const names = e.wb.pkg.workbook.getElementsByTagNameNS(NS_S, "definedNames")[0];
      if (names && !names.children.length) names.remove();
    }
    sheet.touch();
  });
}

/** Values shown (by displayed text) for a filter column, or null when the column isn't filtered. */
export function filterValues(sheet: Sheet, col: number): Set<string> | null {
  const f = sheet.filter;
  const af = child(sheet, "autoFilter");
  if (!f || !af) return null;
  const fc = sheet.all(af, "filterColumn").find((x) => Number(x.getAttribute("colId")) === col - f.c1);
  const filters = sheet.first(fc, "filters");
  if (!filters) return null;
  const out = new Set(sheet.all(filters, "filter").map((x) => x.getAttribute("val") ?? ""));
  if (filters.getAttribute("blank") === "1") out.add("");
  return out;
}

/** Show only rows whose displayed value in `col` is in `allowed` (null clears the column's filter). */
export function filterColumn(e: Edits, sheet: Sheet, col: number, allowed: Set<string> | null): void {
  assertAction(sheet, "autoFilter");
  const f = sheet.filter;
  if (!f) return;
  e.structural(() => {
    const af = child(sheet, "autoFilter")!;
    const colId = col - f.c1;
    for (const fc of sheet.all(af, "filterColumn")) if (Number(fc.getAttribute("colId")) === colId) fc.remove();
    if (allowed) {
      const fc = sheet.doc.createElementNS(NS_S, "filterColumn");
      fc.setAttribute("colId", String(colId));
      const filters = fc.appendChild(sheet.doc.createElementNS(NS_S, "filters"));
      if (allowed.has("")) filters.setAttribute("blank", "1");
      for (const v of [...allowed].filter(Boolean).sort()) filters.appendChild(sheet.doc.createElementNS(NS_S, "filter")).setAttribute("val", v);
      const after = sheet.all(af, "filterColumn").find((x) => Number(x.getAttribute("colId")) > colId) ?? sheet.first(af, "sortState");
      af.insertBefore(fc, after);
    }
    // Every filtered column must allow a row for it to show.
    const columns = sheet.all(af, "filterColumn").map((fc) => {
      const c = f.c1 + Number(fc.getAttribute("colId"));
      const filters = sheet.first(fc, "filters");
      const vals = new Set(sheet.all(filters, "filter").map((x) => x.getAttribute("val") ?? ""));
      if (filters?.getAttribute("blank") === "1") vals.add("");
      return { c, vals, ok: !!filters };
    });
    for (let row = f.r1 + 1; row <= f.r2; row++) {
      const show = columns.every((x) => !x.ok || x.vals.has(sheet.display(row, x.c).text));
      const el = sheet.rowEl(row, !show);
      if (!el) continue;
      if (show) el.removeAttribute("hidden");
      else el.setAttribute("hidden", "1");
    }
    sheet.touch();
  });
}
