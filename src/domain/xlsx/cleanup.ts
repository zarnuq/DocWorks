/*
 * Data cleanup, as Google Sheets' Data menu: remove duplicate rows and split text to columns. Both are one undo step
 * of ordinary cell writes.
 */
import { Change } from "./calculation";
import { Edits } from "./commands";
import { Rect } from "./references";
import { Sheet } from "./workbook-model";

/**
 * Remove rows of `r` that repeat an earlier row in the compared columns (absolute indexes). The first of each stays;
 * the unique rows move up inside the range and the rows left over at its bottom are cleared.
 */
export function removeDuplicates(e: Edits, sheet: Sheet, r: Rect, o: { header: boolean; cols: number[] }): { removed: number; kept: number } {
  const first = r.r1 + (o.header ? 1 : 0);
  const seen = new Set<string>();
  const keep: number[] = [];
  for (let row = first; row <= r.r2; row++) {
    const key = JSON.stringify(o.cols.map((col) => sheet.at(row, col)?.value ?? null));
    if (seen.has(key)) continue;
    seen.add(key);
    keep.push(row);
  }
  const removed = r.r2 - first + 1 - keep.length;
  if (removed) {
    e.run(() => {
      const changes: Change[] = [];
      // Rows only move up, so reading a source row after writing the rows above it is safe.
      keep.forEach((src, i) => {
        const row = first + i;
        if (src === row) return;
        for (let col = r.c1; col <= r.c2; col++) {
          const c = sheet.at(src, col);
          if (c) changes.push(e.place(sheet, row, col, { value: c.value, formula: c.formula, style: c.style, row: src, col }));
          else if (sheet.at(row, col)) changes.push(clearCell(sheet, row, col));
        }
      });
      for (let row = first + keep.length; row <= r.r2; row++) {
        for (let col = r.c1; col <= r.c2; col++) if (sheet.at(row, col)) changes.push(clearCell(sheet, row, col));
      }
      return changes;
    });
  }
  return { removed, kept: keep.length };
}

function clearCell(sheet: Sheet, row: number, col: number): Change {
  sheet.setCell(row, col, { value: null, formula: null, style: 0 });
  return { sheet, row, col };
}

const SEPARATORS = [",", ";", ".", " "];

/** The separator ("Detect automatically") found in the most cells, or null when none has any. */
export function detectSeparator(texts: string[]): string | null {
  let best: string | null = null, most = 0;
  for (const sep of SEPARATORS) {
    const n = texts.filter((t) => t.includes(sep)).length;
    if (n > most) [best, most] = [sep, n];
  }
  return best;
}

/** The text cells of a one-column range (what splitting reads). */
export function splitSources(sheet: Sheet, r: Rect): string[] {
  return sheet.cellsIn(r).flatMap((c) => (typeof c.value === "string" && c.formula === undefined ? [c.value] : []));
}

/**
 * Split each text cell of a one-column range at `sep` into it and the cells to its right, each part entered like
 * typed input. Numbers, formulas and cells without the separator stay as they are.
 */
export function splitText(e: Edits, sheet: Sheet, r: Rect, sep: string): void {
  if (r.c1 !== r.c2) throw new Error("Select one column to split.");
  if (!sep) return;
  e.run(() => {
    const changes: Change[] = [];
    for (const c of sheet.cellsIn(r)) {
      if (typeof c.value !== "string" || c.formula !== undefined || !c.value.includes(sep)) continue;
      c.value.split(sep).forEach((part, i) => {
        e.writeInput(sheet, c.row, r.c1 + i, part);
        changes.push({ sheet, row: c.row, col: r.c1 + i });
      });
    }
    return changes;
  });
}
