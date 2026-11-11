import { allowsAction } from "../../domain/xlsx/protection";
import type { Sheet } from "../../domain/xlsx/workbook-model";
import { cellTextFont, cellTextHeight } from "./paint";

// Excel caps row height at 409.5 points (546 CSS pixels).
const MAX_ROW_HEIGHT = 546;

/** Fit rows to the tallest visible cell, using the grid's own line breaks and font metrics. */
export function fitRowHeights(sheet: Sheet, rows: Iterable<number>, explicit = false): void {
  if (!allowsAction(sheet, "formatRows")) return;
  let ctx: CanvasRenderingContext2D | null | undefined;
  for (const row of rows) {
    const [, custom] = sheet.rowHeightAttributes(row);
    if (sheet.hiddenRows.has(row) || (!explicit && (custom === "1" || custom === "true"))) continue;
    let height = sheet.rows.size;
    for (const cell of sheet.cellsInRow(row)) {
      const { col, style } = cell;
      if (!sheet.colWidth(col)) continue;
      const text = sheet.display(row, col).text;
      if (!text) continue;
      const merge = sheet.mergeAt(row, col);
      // Leave vertical merges alone: their content spans multiple row heights.
      if (merge && (merge.r1 !== merge.r2 || col !== merge.c1)) continue;
      const format = sheet.wb.styles.format(style);
      const size = format.fontSize * 4 / 3;
      const width = sheet.colX((merge?.c2 ?? col) + 1) - sheet.colX(col);
      const css = sheet.wb.styles.css(style);
      const indent = parseFloat(css["padding-left"] ?? css["padding-right"] ?? "0") || 0;
      if (format.wrap) {
        if (ctx === undefined) ctx = createEl("canvas").getContext("2d");
        if (ctx) ctx.font = cellTextFont(format, size);
      }
      const needed = cellTextHeight(ctx ?? null, { text, w: width, size, wrap: format.wrap, pad: indent ? indent - 3 : 0 });
      height = Math.max(height, needed);
      if (height >= MAX_ROW_HEIGHT) break;
    }
    sheet.setFittedRowHeight(row, Math.min(height, MAX_ROW_HEIGHT));
  }
}
