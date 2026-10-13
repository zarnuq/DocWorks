/*
 * Row and column groups (outlines), as Google Sheets' "Group rows": each row / column has an outline level; a group is
 * a run at some level or deeper. Collapsing hides the run and marks the summary row / column after it (before it when
 * sheetPr outlinePr says the summary is above) with collapsed="1", which is how Excel stores it.
 */
import { assertAction } from "./protection";
import { Edits } from "./commands";
import { child, editCols } from "./sheet-edit";
import { Sheet } from "./workbook-model";

export type OutlineAxis = "row" | "col";

export interface OutlineGroup {
  level: number;
  start: number;
  end: number;
  collapsed: boolean;
}

const MAX_LEVEL = 7;
const levelsOf = (sheet: Sheet, axis: OutlineAxis) => (axis === "row" ? sheet.rowLevels : sheet.colLevels);
const hiddenOf = (sheet: Sheet, axis: OutlineAxis) => (axis === "row" ? sheet.hiddenRows : sheet.hiddenCols);

/** The index holding a group's +/- (and its collapsed flag). */
export function summaryIndex(sheet: Sheet, axis: OutlineAxis, g: { start: number; end: number }): number {
  const after = axis === "row" ? sheet.summaryBelow : sheet.summaryRight;
  return after ? g.end + 1 : g.start - 1;
}

/** Every group, by start then level. */
export function outlineGroups(sheet: Sheet, axis: OutlineAxis): OutlineGroup[] {
  const levels = levelsOf(sheet, axis);
  const marked = axis === "row" ? sheet.collapsedRows : sheet.collapsedCols;
  const hidden = hiddenOf(sheet, axis);
  const keys = [...levels.keys()].sort((a, b) => a - b);
  const out: OutlineGroup[] = [];
  const max = Math.max(0, ...levels.values());
  for (let level = 1; level <= max; level++) {
    let start = -1, prev = -2;
    const close = () => {
      if (start < 0) return;
      const g = { level, start, end: prev, collapsed: false };
      let all = true;
      for (let i = start; i <= prev && all; i++) all = hidden.has(i);
      g.collapsed = marked.has(summaryIndex(sheet, axis, g)) || all;
      out.push(g);
    };
    for (const i of keys) {
      if (levels.get(i)! < level) continue;
      if (i !== prev + 1) {
        close();
        start = i;
      }
      prev = i;
    }
    close();
  }
  return out.sort((a, b) => a.start - b.start || a.level - b.level);
}

/** Set an attribute on row / column `i` (null removes it). */
function setAttr(sheet: Sheet, axis: OutlineAxis, i: number, name: string, value: string | null): void {
  if (i < 0) return;
  if (axis === "col") {
    editCols(sheet, i, i, (col) => (value === null ? col.removeAttribute(name) : col.setAttribute(name, value)));
    return;
  }
  const row = sheet.rowEl(i, value !== null);
  if (!row) return;
  if (value === null) row.removeAttribute(name);
  else row.setAttribute(name, value);
}

/** sheetFormatPr outlineLevelRow / outlineLevelCol: the deepest level, which Excel uses to size its gutter. */
function writeMaxLevel(sheet: Sheet, axis: OutlineAxis, levels: Map<number, number>): void {
  const max = Math.max(0, ...levels.values());
  const attr = axis === "row" ? "outlineLevelRow" : "outlineLevelCol";
  let fmt = child(sheet, "sheetFormatPr");
  if (!fmt && !max) return;
  if (!fmt) {
    fmt = child(sheet, "sheetFormatPr", true)!;
    fmt.setAttribute("defaultRowHeight", "15");
  }
  if (max) fmt.setAttribute(attr, String(max));
  else fmt.removeAttribute(attr);
}

function changeLevels(e: Edits, sheet: Sheet, axis: OutlineAxis, a: number, b: number, delta: number, before?: () => void): void {
  assertAction(sheet, axis === "row" ? "formatRows" : "formatColumns");
  const levels = new Map(levelsOf(sheet, axis));
  e.structural(() => {
    before?.();
    for (let i = a; i <= b; i++) {
      const level = Math.max(0, Math.min(MAX_LEVEL, (levels.get(i) ?? 0) + delta));
      if (level === (levels.get(i) ?? 0)) continue;
      if (level) levels.set(i, level);
      else levels.delete(i);
      setAttr(sheet, axis, i, "outlineLevel", level ? String(level) : null);
    }
    writeMaxLevel(sheet, axis, levels);
    sheet.touch();
  });
}

/** Group rows / columns a..b one level deeper. */
export function groupRange(e: Edits, sheet: Sheet, axis: OutlineAxis, a: number, b: number): void {
  changeLevels(e, sheet, axis, a, b, 1);
}

/** Take rows / columns a..b out of one group level; whatever a collapsed group there hid is shown again. */
export function ungroupRange(e: Edits, sheet: Sheet, axis: OutlineAxis, a: number, b: number): void {
  const collapsed = outlineGroups(sheet, axis).filter((g) => g.collapsed && g.start <= b && a <= g.end);
  changeLevels(e, sheet, axis, a, b, -1, () => {
    for (const g of collapsed) {
      for (let i = g.start; i <= g.end; i++) setAttr(sheet, axis, i, "hidden", null);
      setAttr(sheet, axis, summaryIndex(sheet, axis, g), "collapsed", null);
    }
  });
}

/** Collapse (hide) or expand a group; expanding keeps groups nested inside it that are collapsed themselves. */
export function setGroupCollapsed(e: Edits, sheet: Sheet, axis: OutlineAxis, g: OutlineGroup, collapsed: boolean): void {
  const nested = outlineGroups(sheet, axis).filter((n) => n.collapsed && n.level > g.level && n.start >= g.start && n.end <= g.end);
  e.structural(() => {
    for (let i = g.start; i <= g.end; i++) {
      const keep = !collapsed && nested.some((n) => i >= n.start && i <= n.end);
      if (!keep) setAttr(sheet, axis, i, "hidden", collapsed ? "1" : null);
    }
    setAttr(sheet, axis, summaryIndex(sheet, axis, g), "collapsed", collapsed ? "1" : null);
    sheet.touch();
  });
}
