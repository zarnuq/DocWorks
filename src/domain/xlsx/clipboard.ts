/*
 * Copy and paste of ranges. Inside Doc Works a copy keeps formulas and styles; to other apps it's tab-separated
 * displayed values plus an HTML table (what spreadsheets and documents paste best).
 */
import { escapeXml } from "../../core/utils/xml";
import { Change } from "./calculation";
import { Edits } from "./commands";
import { Scalar } from "./number-format";
import { Rect } from "./references";
import { Sheet } from "./workbook-model";

export interface ClipCell {
  value: Scalar;
  formula?: string;
  style: number;
  row: number;
  col: number;
}

export interface Clip {
  rows: (ClipCell | null)[][];
  text: string;
  html: string;
  /** Workbook and sheet path it came from: styles only carry over within the same workbook. */
  source: { pkg: object; path: string; rect: Rect };
  cut: boolean;
}

const quote = (s: string) => (/[\t\n"]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s);

export function copyRange(sheet: Sheet, r: Rect, cut = false): Clip {
  const rows: (ClipCell | null)[][] = [];
  const text: string[] = [];
  const html: string[] = [];
  const styles = sheet.wb.styles;
  for (let row = r.r1; row <= r.r2; row++) {
    const cells: (ClipCell | null)[] = [], line: string[] = [], tds: string[] = [];
    for (let col = r.c1; col <= r.c2; col++) {
      const c = sheet.at(row, col);
      cells.push(c ? { value: c.value, formula: c.formula, style: c.style, row, col } : null);
      const shown = sheet.display(row, col);
      line.push(quote(shown.text));
      const css = c ? { ...styles.css(c.style), ...(shown.color ? { color: shown.color } : {}) } : {};
      delete css["align-items"];
      const style = Object.entries(css).map(([k, v]) => `${k}:${v}`).join(";");
      tds.push(`<td${style ? ` style="${escapeXml(style)}"` : ""}>${escapeXml(shown.text).replace(/\n/g, "<br>")}</td>`);
    }
    rows.push(cells);
    text.push(line.join("\t"));
    html.push(`<tr>${tds.join("")}</tr>`);
  }
  return {
    rows,
    text: text.join("\n"),
    html: `<table style="border-collapse:collapse">${html.join("")}</table>`,
    source: { pkg: sheet.wb.pkg, path: sheet.ref.path, rect: r },
    cut,
  };
}

/** Rows of fields from tab-separated text (quoted fields may hold tabs, newlines and doubled quotes). */
export function parseTsv(text: string): string[][] {
  const rows: string[][] = [[]];
  let field = "", i = 0, quoted = false;
  text = text.replace(/\r\n?/g, "\n").replace(/\n$/, "");
  while (i < text.length) {
    const ch = text[i];
    if (quoted) {
      if (ch === '"' && text[i + 1] === '"') {
        field += '"';
        i += 2;
        continue;
      }
      if (ch === '"') quoted = false;
      else field += ch;
    } else if (ch === '"' && field === "") quoted = true;
    else if (ch === "\t") {
      rows[rows.length - 1].push(field);
      field = "";
    } else if (ch === "\n") {
      rows[rows.length - 1].push(field);
      rows.push([]);
      field = "";
    } else field += ch;
    i++;
  }
  rows[rows.length - 1].push(field);
  return rows;
}

/** Target area: the selection when it's a multiple of the copied size, else the copy's size at its corner. */
function target(sel: Rect, h: number, w: number): Rect {
  const sh = sel.r2 - sel.r1 + 1, sw = sel.c2 - sel.c1 + 1;
  if (sh % h === 0 && sw % w === 0) return sel;
  return { r1: sel.r1, c1: sel.c1, r2: sel.r1 + h - 1, c2: sel.c1 + w - 1 };
}

export function pasteClip(edits: Edits, sheet: Sheet, sel: Rect, clip: Clip): Rect {
  const h = clip.rows.length, w = clip.rows[0]?.length ?? 0;
  const area = clip.cut ? { r1: sel.r1, c1: sel.c1, r2: sel.r1 + h - 1, c2: sel.c1 + w - 1 } : target(sel, h, w);
  const sameBook = clip.source.pkg === edits.wb.pkg;
  edits.run(() => {
    const changes: Change[] = [];
    const src = clip.cut ? edits.wb.sheets.find((s) => s.ref.path === clip.source.path) : undefined;
    // A cut moves: clear the source first (cells the paste lands on are written again below).
    if (src) {
      for (const c of src.cellsIn(clip.source.rect)) {
        src.setCell(c.row, c.col, { value: null, formula: null, style: 0 });
        changes.push({ sheet: src, row: c.row, col: c.col });
      }
    }
    for (let row = area.r1; row <= area.r2; row++) {
      for (let col = area.c1; col <= area.c2; col++) {
        const c = clip.rows[(row - area.r1) % h][(col - area.c1) % w];
        if (c) {
          // Moved formulas keep pointing at the same cells; copied ones move their relative references.
          const from = clip.cut ? { ...c, row, col } : c;
          changes.push(edits.place(sheet, row, col, from, sameBook));
        } else if (sheet.at(row, col)) {
          sheet.setCell(row, col, { value: null, formula: null, style: sameBook ? 0 : undefined });
          changes.push({ sheet, row, col });
        }
      }
    }
    return changes;
  });
  return area;
}

/** Paste text from another app: each field is read like typed input (numbers, dates, formulas). */
export function pasteText(edits: Edits, sheet: Sheet, sel: Rect, text: string): Rect {
  const rows = parseTsv(text);
  const h = rows.length, w = Math.max(...rows.map((r) => r.length));
  const area = target(sel, h, w);
  edits.run(() => {
    const changes: Change[] = [];
    for (let row = area.r1; row <= area.r2; row++) {
      for (let col = area.c1; col <= area.c2; col++) {
        const raw = rows[(row - area.r1) % h][(col - area.c1) % w] ?? "";
        try {
          edits.writeInput(sheet, row, col, raw);
        } catch {
          edits.writeInput(sheet, row, col, "'" + raw); // a malformed formula pastes as text
        }
        changes.push({ sheet, row, col });
      }
    }
    return changes;
  });
  return area;
}
