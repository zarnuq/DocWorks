/*
 * CSV (RFC 4180) in and out, as Google Sheets: import adds a tab with text converted like typed input; download
 * writes the sheet's displayed values.
 */
import { Change } from "./calculation";
import { Edits } from "./commands";
import { addSheet } from "./sheet-edit";
import { Sheet } from "./workbook-model";
import type { Rect } from "./references";

export function parseCsv(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [], field = "", quoted = false;
  const s = text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (quoted) {
      if (ch !== '"') field += ch;
      else if (s[i + 1] === '"') field += s[i++];
      else quoted = false;
    } else if (ch === '"') quoted = true;
    else if (ch === ",") {
      row.push(field);
      field = "";
    } else if (ch === "\n" || ch === "\r") {
      if (ch === "\r" && s[i + 1] === "\n") i++;
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
    } else field += ch;
  }
  if (field || row.length) rows.push([...row, field]);
  return rows;
}

const quote = (v: string) => (/[",\r\n]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v);

/** The sheet's used range from A1, as displayed, one CRLF-terminated line per row. */
export function sheetToCsv(sheet: Sheet): string {
  const used = sheet.used();
  if (!used) return "";
  let out = "";
  for (let row = 0; row <= used.r2; row++) {
    const fields: string[] = [];
    for (let col = 0; col <= used.c2; col++) fields.push(quote(sheet.display(row, col).text));
    out += fields.join(",") + "\r\n";
  }
  return out;
}

/**
 * Copy as Markdown table: the shown text of `r` (cut to the used area, hidden rows and columns left out), the first
 * row as the header, columns of numbers right-aligned. "" when nothing is in the range.
 */
export function rangeToMarkdown(sheet: Sheet, r: Rect): string {
  const used = sheet.used();
  if (!used) return "";
  const r2 = Math.min(r.r2, used.r2), c2 = Math.min(r.c2, used.c2);
  const rows: number[] = [], cols: number[] = [];
  for (let i = r.r1; i <= r2; i++) if (sheet.rowHeight(i) > 0) rows.push(i);
  for (let i = r.c1; i <= c2; i++) if (sheet.colWidth(i) > 0) cols.push(i);
  if (!rows.some((row) => cols.some((col) => sheet.at(row, col)))) return "";
  const cell = (row: number, col: number) => sheet.display(row, col).text.replace(/\|/g, "\\|").replace(/\r?\n/g, "<br>");
  const line = (cells: string[]) => `| ${cells.join(" | ")} |`;
  const numeric = (col: number) => {
    const body = rows.slice(1).map((row) => sheet.at(row, col)?.value).filter((v) => v !== undefined && v !== null && v !== "");
    return body.length > 0 && body.every((v) => typeof v === "number");
  };
  return [line(cols.map((col) => cell(rows[0], col))), line(cols.map((col) => (numeric(col) ? "---:" : "---"))), ...rows.slice(1).map((row) => line(cols.map((col) => cell(row, col)))), ""].join("\n");
}

/** Add a tab after tab `after`, named from `name` (made valid and unique), holding `rows`; returns the tab's name. */
export function importCsv(e: Edits, rows: string[][], name: string, after: number): string {
  const base = name.replace(/[\\/?*[\]:]/g, " ").replace(/^'+|'+$/g, "").trim().slice(0, 31) || "Imported";
  let added = "";
  e.run(() => {
    added = addSheet(e, after, base);
    const sheet = e.wb.sheet(added)!;
    const changes: Change[] = [];
    rows.forEach((fields, row) =>
      fields.forEach((text, col) => {
        if (text === "") return;
        e.writeInput(sheet, row, col, text);
        changes.push({ sheet, row, col });
      })
    );
    return changes;
  });
  return added;
}
