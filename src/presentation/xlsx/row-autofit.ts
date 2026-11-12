import { allowsAction } from "../../domain/xlsx/protection";
import type { Rect } from "../../domain/xlsx/references";
import type { Cell, Sheet } from "../../domain/xlsx/workbook-model";
import { cellTextFont, cellTextHeight } from "./paint";

// Excel caps row height at 409.5 points (546 CSS pixels).
const MAX_ROW_HEIGHT = 546;

/** Merges in different columns can share row heights. Fit overlapping groups together. */
function affectedMerges(sheet: Sheet, rows: Set<number>): Rect[] {
  const merges = sheet.merges.filter((merge) => {
    if (merge.r1 === merge.r2) return false;
    const anchor = sheet.at(merge.r1, merge.c1);
    return anchor && sheet.wb.styles.format(anchor.style).wrap;
  }).sort((a, b) => a.r1 - b.r1);
  if (!merges.length) return [];
  const selected = [...rows].sort((a, b) => a - b);
  const affected: Rect[] = [];
  let selectedIndex = 0;
  for (let i = 0; i < merges.length;) {
    const first = merges[i].r1;
    let last = merges[i].r2;
    let end = i + 1;
    while (end < merges.length && merges[end].r1 <= last) {
      last = Math.max(last, merges[end++].r2);
    }
    while (selectedIndex < selected.length && selected[selectedIndex] < first) selectedIndex++;
    if (selectedIndex >= selected.length) break;
    if (selected[selectedIndex] <= last) {
      for (let row = first; row <= last; row++) rows.add(row);
      for (let j = i; j < end; j++) affected.push(merges[j]);
    }
    i = end;
  }
  return affected;
}

function growMergedRows(sheet: Sheet, merge: Rect, heights: Map<number, number>, required: number): void {
  const growable: number[] = [];
  let missing = required;
  for (let row = merge.r1; row <= merge.r2; row++) {
    const fitted = heights.get(row);
    missing -= fitted ?? sheet.rowHeight(row);
    if (fitted !== undefined && fitted < MAX_ROW_HEIGHT) growable.push(row);
  }
  if (missing <= 0) return;
  // Share the extra height evenly where possible. Apply the smallest remaining capacities first.
  growable.sort((a, b) => heights.get(b)! - heights.get(a)!);
  for (const [i, row] of growable.entries()) {
    if (missing <= 0) break;
    const current = heights.get(row)!;
    const extra = Math.min(MAX_ROW_HEIGHT - current, Math.ceil(missing / (growable.length - i)));
    heights.set(row, current + extra);
    missing -= extra;
  }
}

/** Fit rows and wrapped merges using the grid's own line breaks and font metrics. */
export function fitRowHeights(sheet: Sheet, rows: Iterable<number>, explicit = false): void {
  if (!allowsAction(sheet, "formatRows")) return;
  const targets = new Set(rows);
  if (!targets.size) return;
  const merges = affectedMerges(sheet, targets);
  let ctx: CanvasRenderingContext2D | null | undefined;
  const measure = (cell: Cell, width: number): number => {
    const text = sheet.display(cell.row, cell.col).text;
    if (!text || width <= 0) return 0;
    const format = sheet.wb.styles.format(cell.style);
    const size = format.fontSize * 4 / 3;
    const css = sheet.wb.styles.css(cell.style);
    const indent = parseFloat(css["padding-left"] ?? css["padding-right"] ?? "0") || 0;
    if (format.wrap) {
      if (ctx === undefined) ctx = createEl("canvas").getContext("2d");
      if (ctx) ctx.font = cellTextFont(format, size);
    }
    return cellTextHeight(ctx ?? null, { text, w: width, size, wrap: format.wrap, pad: indent ? indent - 3 : 0 });
  };
  const heights = new Map<number, number>();
  for (const row of targets) {
    const [, custom] = sheet.rowHeightAttributes(row);
    if (sheet.hiddenRows.has(row) || (!explicit && (custom === "1" || custom === "true"))) continue;
    let height = sheet.rows.size;
    for (const cell of sheet.cellsInRow(row)) {
      const { col } = cell;
      if (!sheet.colWidth(col)) continue;
      const merge = sheet.mergeAt(row, col);
      if (merge && (merge.r1 !== merge.r2 || col !== merge.c1)) continue;
      const width = sheet.colX((merge?.c2 ?? col) + 1) - sheet.colX(col);
      height = Math.max(height, measure(cell, width));
      if (height >= MAX_ROW_HEIGHT) break;
    }
    heights.set(row, Math.min(height, MAX_ROW_HEIGHT));
  }
  for (const merge of merges) {
    const width = sheet.colX(merge.c2 + 1) - sheet.colX(merge.c1);
    const required = measure(sheet.at(merge.r1, merge.c1)!, width);
    if (required) growMergedRows(sheet, merge, heights, required);
  }
  for (const [row, height] of heights) sheet.setFittedRowHeight(row, height);
}
