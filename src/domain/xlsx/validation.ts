import { assertAction } from "./protection";
/*
 * Data validation: dropdown lists (items typed in, or a range's values), number limits and custom formulas, read
 * from a sheet's dataValidation elements, checked on input, and written as Google's data validation sidebar does.
 * As in Excel, a rule rejects bad input unless its error style is a warning (or it shows no error at all).
 */
import { NS_S } from "../../infrastructure/xlsx/xlsx-workbook";
import { Edits } from "./commands";
import { compareValues } from "./formula-evaluator";
import { translateFormula } from "./formula";
import { Scalar } from "./number-format";
import { parseRange, Rect, rangeName } from "./references";
import { Sheet } from "./workbook-model";

export interface Validation {
  el: Element;
  ranges: Rect[];
  /** list, whole, decimal, date, time, textLength, custom. */
  type: string;
  operator: string;
  formulas: string[];
  /** A dropdown's items when typed into the rule ("a,b,c"); null when they come from a range or formula. */
  items: string[] | null;
  /** Bad input is refused (else only marked). */
  reject: boolean;
  allowBlank: boolean;
}

export interface NewValidation {
  type: "list" | "decimal" | "custom";
  items?: string[];
  /** A dropdown's items from a range, e.g. "Sheet1!A1:A10". */
  source?: string;
  operator?: string;
  formulas?: string[];
  reject: boolean;
}

const scalarText = (v: Scalar) => (typeof v === "object" && v !== null ? v.error : String(v));
const kids = (el: Element | null | undefined, tag: string) => (el ? Array.from(el.children).filter((c) => c.localName === tag) : []);
const on = (el: Element, attr: string) => ["1", "true"].includes(el.getAttribute(attr) ?? "");

/** Split Excel's quoted item list: "a,b" with "" for a quote inside. */
function literalItems(formula: string): string[] | null {
  const m = /^"((?:[^"]|"")*)"$/.exec(formula.trim());
  return m ? m[1].replace(/""/g, '"').split(",") : null;
}

const cache = new WeakMap<Document, Validation[]>();

/** Forget the rules read for `sheet` (after an edit or undo changed its XML). */
export function invalidateValidations(sheet: Sheet): void {
  cache.delete(sheet.doc);
}

/** The sheet's rules, read once until `invalidateValidations` (the grid asks for every drawn cell). */
export function readValidations(sheet: Sheet): Validation[] {
  const hit = cache.get(sheet.doc);
  if (hit) return hit;
  const container = kids(sheet.doc.documentElement, "dataValidations")[0];
  const rules = kids(container, "dataValidation").map((el): Validation => {
    const formulas = ["formula1", "formula2"].map((t) => kids(el, t)[0]?.textContent ?? null).filter((f): f is string => f !== null);
    const type = el.getAttribute("type") ?? "none";
    const style = el.getAttribute("errorStyle") ?? "stop";
    return {
      el,
      ranges: (el.getAttribute("sqref") ?? "").split(/\s+/).map(parseRange).filter((r): r is Rect => !!r),
      type,
      operator: el.getAttribute("operator") ?? "between",
      formulas,
      items: type === "list" ? literalItems(formulas[0] ?? "") : null,
      reject: on(el, "showErrorMessage") && style === "stop",
      allowBlank: on(el, "allowBlank"),
    };
  });
  cache.set(sheet.doc, rules);
  return rules;
}

export function validationAt(sheet: Sheet, row: number, col: number): Validation | null {
  return readValidations(sheet).find((v) => v.ranges.some((r) => row >= r.r1 && row <= r.r2 && col >= r.c1 && col <= r.c2)) ?? null;
}

/** A dropdown's choices: its typed items, or the distinct non-blank values of its range. */
export function listItems(v: Validation, sheet: Sheet): string[] {
  if (v.items) return v.items;
  const ref = (v.formulas[0] ?? "").replace(/\$/g, "");
  const bang = ref.lastIndexOf("!");
  const name = bang >= 0 ? ref.slice(0, bang).replace(/^'(.*)'$/, "$1").replace(/''/g, "'") : null;
  const source = name ? sheet.wb.sheet(name) : sheet;
  const rect = parseRange(bang >= 0 ? ref.slice(bang + 1) : ref);
  if (!source || !rect) return [];
  const out: string[] = [];
  for (let r = rect.r1; r <= Math.min(rect.r2, rect.r1 + 5000); r++)
    for (let c = rect.c1; c <= rect.c2; c++) {
      const t = source.display(r, c).text;
      if (t && !out.includes(t)) out.push(t);
    }
  return out;
}

/** Whether `value` (now in the cell, for custom formulas) passes the rule. */
export function checkValue(v: Validation, value: Scalar, sheet: Sheet, row: number, col: number, evaluate: (row: number, col: number, formula: string) => Scalar): boolean {
  // Clearing a cell is always allowed, as in Sheets and Excel.
  if (value === null || value === "") return true;
  const operand = (i: number): Scalar => {
    const f = v.formulas[i] ?? "";
    const base = v.ranges[0];
    const n = Number(f);
    return f.trim() !== "" && Number.isFinite(n) ? n : evaluate(row, col, translateFormula(f, row - base.r1, col - base.c1));
  };
  const compare = (x: Scalar) => {
    const a = operand(0);
    switch (v.operator) {
      case "between":
      case "notBetween": {
        const b = operand(1);
        const inside = compareValues(x, a) >= 0 && compareValues(x, b) <= 0;
        return v.operator === "between" ? inside : !inside;
      }
      case "equal":
        return compareValues(x, a) === 0;
      case "notEqual":
        return compareValues(x, a) !== 0;
      case "greaterThan":
        return compareValues(x, a) > 0;
      case "greaterThanOrEqual":
        return compareValues(x, a) >= 0;
      case "lessThan":
        return compareValues(x, a) < 0;
      case "lessThanOrEqual":
        return compareValues(x, a) <= 0;
      default:
        return true;
    }
  };
  switch (v.type) {
    case "list": {
      const text = typeof value === "string" ? value : sheet.display(row, col).text || scalarText(value);
      return listItems(v, sheet).some((i) => i.trim().toLowerCase() === text.trim().toLowerCase());
    }
    case "whole":
      return typeof value === "number" && Number.isInteger(value) && compare(value);
    case "decimal":
    case "date":
    case "time":
      return typeof value === "number" && compare(value);
    case "textLength":
      return compare(scalarText(value).length);
    case "custom": {
      const base = v.ranges[0];
      const r = evaluate(row, col, translateFormula(v.formulas[0] ?? "TRUE", row - base.r1, col - base.c1));
      return r === true || (typeof r === "number" && r !== 0) || (typeof r === "string" && r.toUpperCase() === "TRUE");
    }
    default:
      return true;
  }
}

// -----------------------------------------------------------------------------
// Editing
// -----------------------------------------------------------------------------

/** `a` without the cells of `b`, as up to four rectangles. */
export function subtract(a: Rect, b: Rect): Rect[] {
  if (b.r1 > a.r2 || b.r2 < a.r1 || b.c1 > a.c2 || b.c2 < a.c1) return [a];
  const out: Rect[] = [];
  if (b.r1 > a.r1) out.push({ ...a, r2: b.r1 - 1 });
  if (b.r2 < a.r2) out.push({ ...a, r1: b.r2 + 1 });
  const r1 = Math.max(a.r1, b.r1), r2 = Math.min(a.r2, b.r2);
  if (b.c1 > a.c1) out.push({ r1, r2, c1: a.c1, c2: b.c1 - 1 });
  if (b.c2 < a.c2) out.push({ r1, r2, c1: b.c2 + 1, c2: a.c2 });
  return out;
}

/** Take `ranges` out of every rule's sqref, dropping rules (and the container) left empty. */
function clear(sheet: Sheet, ranges: Rect[]): void {
  const container = kids(sheet.doc.documentElement, "dataValidations")[0];
  if (!container) return;
  for (const el of kids(container, "dataValidation")) {
    let rects = (el.getAttribute("sqref") ?? "").split(/\s+/).map(parseRange).filter((r): r is Rect => !!r);
    for (const cut of ranges) rects = rects.flatMap((r) => subtract(r, cut));
    if (rects.length) el.setAttribute("sqref", rects.map(rangeName).join(" "));
    else el.remove();
  }
  const left = kids(container, "dataValidation").length;
  if (left) container.setAttribute("count", String(left));
  else container.remove();
}

const escape = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;");

/** Set `rule` on `ranges` (replacing any rule there). */
export function setValidation(e: Edits, sheet: Sheet, ranges: Rect[], rule: NewValidation): void {
  assertAction(sheet, "sheet");
  const formulas =
    rule.type === "list" ? [rule.items ? `"${rule.items.map((i) => i.replace(/"/g, '""')).join(",")}"` : rule.source ?? ""] : rule.formulas ?? [];
  const attrs = [`type="${rule.type}"`, `allowBlank="1"`, `showErrorMessage="1"`, rule.reject ? "" : `errorStyle="warning"`, rule.type === "decimal" ? `operator="${rule.operator ?? "between"}"` : "", `sqref="${ranges.map(rangeName).join(" ")}"`].filter(Boolean).join(" ");
  const xml = `<dataValidation xmlns="${NS_S}" ${attrs}>${formulas.map((f, i) => `<formula${i + 1}>${escape(f)}</formula${i + 1}>`).join("")}</dataValidation>`;
  e.structural(() => {
    clear(sheet, ranges);
    invalidateValidations(sheet);
    const root = sheet.doc.documentElement;
    let container = kids(root, "dataValidations")[0];
    if (!container) {
      container = sheet.doc.createElementNS(NS_S, "dataValidations");
      const AFTER = ["hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks", "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing", "legacyDrawing", "legacyDrawingHF", "drawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst"];
      root.insertBefore(container, Array.from(root.children).find((c) => AFTER.includes(c.localName)) ?? null);
    }
    container.appendChild(sheet.doc.importNode(new DOMParser().parseFromString(xml, "application/xml").documentElement, true));
    container.setAttribute("count", String(kids(container, "dataValidation").length));
    sheet.wb.pkg.touch(sheet.ref.path);
  });
}

export function removeValidation(e: Edits, sheet: Sheet, ranges: Rect[]): void {
  assertAction(sheet, "sheet");
  e.structural(() => {
    clear(sheet, ranges);
    invalidateValidations(sheet);
    sheet.wb.pkg.touch(sheet.ref.path);
  });
}
