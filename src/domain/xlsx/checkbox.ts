/**
 * Cell checkboxes, as Google's Insert › Checkbox: TRUE / FALSE cells whose format carries Excel's native checkbox
 * (Excel 365 draws them; other apps show TRUE / FALSE). The format points through an xfComplement index at a chain
 * of feature property bags: XFComplements -> XFComplement -> XFControls -> Checkbox.
 */
import { assertAction } from "./protection";
import type { Change } from "./calculation";
import type { Edits } from "./commands";
import type { Rect } from "./references";
import { NS_FPB } from "./styles";
import type { Sheet, Workbook } from "./workbook-model";

const bagsOf = (doc: Document) => Array.from(doc.documentElement.children).filter((b) => b.localName === "bag");
const kids = (el: Element, name: string) => Array.from(el.children).filter((c) => c.localName === name);
const bagRef = (bag: Element | undefined, key: string) => {
  const id = bag && kids(bag, "bagId").find((b) => b.getAttribute("k") === key)?.textContent;
  return id === undefined || id === null ? null : Number(id);
};
const mapping = (doc: Document) => {
  const list = bagsOf(doc).find((b) => b.getAttribute("type") === "XFComplements");
  return list ? kids(list, "a").find((a) => a.getAttribute("k") === "MappedFeaturePropertyBags") ?? null : null;
};

/** The xfComplement index that means "checkbox" in this workbook, or null. */
function checkboxIndex(doc: Document | null): number | null {
  if (!doc) return null;
  const bags = bagsOf(doc);
  const list = mapping(doc);
  const ids = list ? kids(list, "bagId").map((b) => Number(b.textContent)) : [];
  for (const [i, id] of ids.entries()) {
    const controls = bags[bagRef(bags[id], "XFControls") ?? -1];
    if (bags[bagRef(controls, "CellControl") ?? -1]?.getAttribute("type") === "Checkbox") return i;
  }
  return null;
}

const cache = new WeakMap<Workbook, number | null>();
function checkboxComplement(wb: Workbook): number | null {
  if (!cache.has(wb)) {
    const path = wb.pkg.featureBags();
    cache.set(wb, checkboxIndex(path ? wb.pkg.xml(path) : null));
  }
  return cache.get(wb)!;
}

/** Whether cell format `style` is a checkbox. */
export function isCheckbox(wb: Workbook, style: number): boolean {
  const i = wb.styles.complement(style);
  return i !== null && i === checkboxComplement(wb);
}

/** Add the checkbox bags (Checkbox, XFControls, XFComplement and its mapping) unless there; returns its index. */
function ensureCheckbox(wb: Workbook): number {
  const doc = wb.pkg.xml(wb.pkg.featureBags(true)!)!;
  const found = checkboxIndex(doc);
  if (found !== null) return found;
  const root = doc.documentElement;
  const bag = (type: string, ref?: [string, number]) => {
    const el = doc.createElementNS(NS_FPB, "bag");
    el.setAttribute("type", type);
    if (ref) {
      const id = el.appendChild(doc.createElementNS(NS_FPB, "bagId"));
      id.setAttribute("k", ref[0]);
      id.textContent = String(ref[1]);
    }
    // Before any XFComplements bag and the extension elements.
    root.insertBefore(el, bagsOf(doc).find((b) => b.getAttribute("type") === "XFComplements") ?? kids(root, "extLst")[0] ?? null);
    return bagsOf(doc).indexOf(el);
  };
  const checkbox = bag("Checkbox");
  const controls = bag("XFControls", ["CellControl", checkbox]);
  const complement = bag("XFComplement", ["XFControls", controls]);
  let list = mapping(doc);
  if (!list) {
    const complements = doc.createElementNS(NS_FPB, "bag");
    complements.setAttribute("type", "XFComplements");
    complements.setAttribute("extRef", "XFComplementsMapperExtRef");
    list = complements.appendChild(doc.createElementNS(NS_FPB, "a"));
    list.setAttribute("k", "MappedFeaturePropertyBags");
    root.insertBefore(complements, kids(root, "extLst")[0] ?? null);
  }
  list.appendChild(doc.createElementNS(NS_FPB, "bagId")).textContent = String(complement);
  return kids(list, "bagId").length - 1;
}

/** Insert › Checkbox: the range becomes checkboxes; empty cells start unchecked (FALSE), others keep their value. */
export function insertCheckboxes(e: Edits, sheet: Sheet, r: Rect): void {
  assertAction(sheet, "formatCells");
  let i = 0;
  e.structural(() => {
    i = ensureCheckbox(e.wb);
  }, () => {
    const s = e.wb.sheets.find((x) => x.ref.path === sheet.ref.path)!;
    const derived = new Map<number, number>();
    for (let row = r.r1; row <= r.r2; row++) {
      for (let col = r.c1; col <= r.c2; col++) {
        const cell = s.at(row, col);
        const from = cell?.style ?? 0;
        if (!derived.has(from)) derived.set(from, e.wb.styles.derive(from, { complement: i }));
        const empty = !cell || (cell.value === null && cell.formula === undefined);
        s.setCell(row, col, empty ? { style: derived.get(from)!, value: false, formula: null } : { style: derived.get(from)! });
      }
    }
    e.wb.touchStyles();
  });
}

/** Click / Space on checkboxes: all checked if any is unchecked, else all unchecked (formulas are left alone). */
export function toggleCheckboxes(e: Edits, sheet: Sheet, r: Rect): void {
  const boxes = sheet.cellsIn(r).filter((c) => c.formula === undefined && isCheckbox(e.wb, c.style));
  if (!boxes.length) return;
  const to = boxes.some((c) => c.value !== true);
  e.run(() => {
    const changes: Change[] = [];
    for (const c of boxes) {
      sheet.setCell(c.row, c.col, { value: to, formula: null });
      changes.push({ sheet, row: c.row, col: c.col });
    }
    return changes;
  });
}
