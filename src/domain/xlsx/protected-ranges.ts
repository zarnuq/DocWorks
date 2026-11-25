/**
 * Google Sheets' protected ranges in its "show a warning when editing this range" mode: editing a cell in one asks
 * first. Saved as Excel's x:protectedRanges, which Excel only acts on in a protected sheet (where it means the
 * opposite: ranges left editable), so a protected sheet keeps Excel's locking and these ranges don't warn there.
 */
import type { Edits } from "./commands";
import { child } from "./sheet-edit";
import { contains, parseRange, Rect } from "./references";
import { isProtected } from "./protection";
import type { Sheet } from "./workbook-model";

export interface ProtectedRange {
  /** Google's description; Excel's range title. */
  name: string;
  /** Space-separated A1 ranges. */
  sqref: string;
}

const rects = (sqref: string): Rect[] => sqref.split(/\s+/).map(parseRange).filter((r): r is Rect => !!r);

export function protectedRanges(sheet: Sheet): ProtectedRange[] {
  return sheet.all(child(sheet, "protectedRanges"), "protectedRange").map((el) => ({ name: el.getAttribute("name") ?? "", sqref: el.getAttribute("sqref") ?? "" }));
}

/** The protected range a cell edit should warn about, if any (none in a protected sheet: Excel's locking rules there). */
export function protectedRangeAt(sheet: Sheet, row: number, col: number): ProtectedRange | null {
  if (isProtected(sheet)) return null;
  return protectedRanges(sheet).find((p) => rects(p.sqref).some((r) => contains(r, row, col))) ?? null;
}

/** A name not used by the sheet's other ranges ("Range 1", "Range 2"…), as Excel needs one per range. */
export function newRangeName(sheet: Sheet): string {
  const used = new Set(protectedRanges(sheet).map((p) => p.name));
  let n = 1;
  while (used.has(`Range ${n}`)) n++;
  return `Range ${n}`;
}

/** Check a typed range ("A1:B5" or several separated by spaces / commas); null when it isn't one. */
export function normalizeSqref(text: string): string | null {
  const parts = text.split(/[\s,]+/).filter(Boolean).map((t) => t.replace(/\$/g, "").toUpperCase());
  return parts.length && parts.every((t) => parseRange(t)) ? parts.join(" ") : null;
}

/**
 * Add (`index` null) or change a range; `null` as the range removes it. Unmodelled attributes of an existing range
 * (Excel's password hash, security descriptor) are kept.
 */
export function setProtectedRange(e: Edits, sheet: Sheet, index: number | null, range: ProtectedRange | null): void {
  const path = sheet.ref.path;
  e.structural(() => {
    const s = e.wb.sheets.find((x) => x.ref.path === path)!;
    const wrap = child(s, "protectedRanges", true)!;
    let el = index === null ? null : s.all(wrap, "protectedRange")[index] ?? null;
    if (!range) el?.remove();
    else {
      if (!el) el = wrap.appendChild(s.doc.createElementNS(wrap.namespaceURI, "protectedRange"));
      el.setAttribute("sqref", range.sqref);
      el.setAttribute("name", range.name);
    }
    if (!s.all(wrap, "protectedRange").length) wrap.remove();
    s.touch();
  });
}
