/** Find and replace over cell contents: shown text for values, formula text for formulas. */
import { Change } from "./calculation";
import { Edits } from "./commands";
import { Sheet } from "./workbook-model";
import { notesOf } from "./notes";
import { commentsOf } from "./comments";
import { parseRange } from "./references";
import { textMatches, TextMatch } from "../../core/utils/text-matches";

export interface FindOptions {
  matchCase?: boolean;
  /** The whole cell must match. */
  entire?: boolean;
  /** Search formula text instead of results. */
  formulas?: boolean;
  notes?: boolean;
  comments?: boolean;
}

export type SheetSearchMatch = ({ kind: "cell" } & CellMatch) | (CellMatch & TextMatch & {
  kind: "note" | "comment";
  text: string;
  author: string;
  messageId?: string;
  resolved?: boolean;
});

/** Annotation-only cells are included even when they have no stored value. */
export function findSheetMatches(sheets: Sheet[], query: string, o: FindOptions = {}): SheetSearchMatch[] {
  if (!query) return [];
  const out: SheetSearchMatch[] = [];
  for (const sheet of sheets) {
    const hits: SheetSearchMatch[] = findCells([sheet], query, o).map(m => ({ ...m, kind: "cell" }));
    const add = (ref: string, text: string, author: string, kind: "note" | "comment", messageId?: string, resolved?: boolean) => {
      const range = parseRange(ref);
      if (!range) return;
      hits.push(...textMatches(text, query, o.matchCase, o.entire).map(m => ({ ...m, sheet, row: range.r1, col: range.c1, kind, text, author, messageId, resolved })));
    };
    if (o.notes) for (const [ref, note] of notesOf(sheet.wb.pkg, sheet.ref.path)) add(ref, note.text, note.author, "note");
    if (o.comments) for (const [ref, threads] of commentsOf(sheet.wb.pkg, sheet.ref.path)) for (const thread of threads) for (const message of thread.messages) add(ref, message.text, message.author, "comment", message.id, thread.resolved);
    hits.sort((a, b) => a.row - b.row || a.col - b.col);
    out.push(...hits);
  }
  return out;
}

export interface CellMatch {
  sheet: Sheet;
  row: number;
  col: number;
}

const escape = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

function pattern(query: string, o: FindOptions, global = false): RegExp {
  const body = escape(query);
  return new RegExp(o.entire ? `^${body}$` : body, (o.matchCase ? "" : "i") + (global ? "g" : ""));
}

/** Text searched for a cell. */
function content(sheet: Sheet, row: number, col: number, o: FindOptions): string {
  const c = sheet.at(row, col);
  if (!c) return "";
  if (o.formulas && c.formula !== undefined) return "=" + c.formula;
  return sheet.display(row, col).text;
}

/** Matches in reading order (row by row), on `sheets` in tab order. */
export function findCells(sheets: Sheet[], query: string, o: FindOptions = {}): CellMatch[] {
  if (!query) return [];
  const re = pattern(query, o);
  const out: CellMatch[] = [];
  for (const sheet of sheets) {
    const cells = [...sheet.cells.values()].sort((a, b) => a.row - b.row || a.col - b.col);
    for (const c of cells) if (re.test(content(sheet, c.row, c.col, o))) out.push({ sheet, row: c.row, col: c.col });
  }
  return out;
}

/** Replace in the given matches as one undo step; values are re-read like typed input. Returns how many changed. */
export function replaceCells(edits: Edits, matches: CellMatch[], query: string, replacement: string, o: FindOptions = {}): number {
  let n = 0;
  edits.run(() => {
    const changes: Change[] = [];
    for (const m of matches) {
      const c = m.sheet.at(m.row, m.col);
      if (!c) continue;
      // Formulas are only rewritten when searching formulas; replacing in a formula's result would lose it.
      if (c.formula !== undefined && !o.formulas) continue;
      const before = c.formula !== undefined ? "=" + c.formula : typeof c.value === "string" ? c.value : m.sheet.display(m.row, m.col).text;
      const after = before.replace(pattern(query, o, true), () => replacement);
      if (after === before) continue;
      try {
        edits.writeInput(m.sheet, m.row, m.col, after);
      } catch {
        continue; // the replacement would break the formula
      }
      changes.push(m);
      n++;
    }
    return changes;
  });
  return n;
}
