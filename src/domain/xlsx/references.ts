/** A1-style cell references. Rows and columns are 0-based internally. */

export const MAX_ROW = 1048575;
export const MAX_COL = 16383;

export interface CellRef {
  row: number;
  col: number;
  absRow: boolean;
  absCol: boolean;
}

/** Inclusive rectangle, r1 <= r2 and c1 <= c2. */
export interface Rect {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

export function colName(col: number): string {
  let s = "";
  for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

export function colIndex(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64;
  return n - 1;
}

export const cellName = (row: number, col: number) => `${colName(col)}${row + 1}`;

export function parseCellRef(text: string): CellRef | null {
  const m = /^(\$?)([A-Za-z]{1,3})(\$?)([0-9]{1,7})$/.exec(text);
  if (!m) return null;
  const col = colIndex(m[2]);
  const row = Number(m[4]) - 1;
  if (col > MAX_COL || row < 0 || row > MAX_ROW) return null;
  return { row, col, absRow: m[3] === "$", absCol: m[1] === "$" };
}

/** "A1:B2", "C:C", "2:3" or a single cell; null if it isn't one. */
export function parseRange(text: string): Rect | null {
  const [a, b = a, extra] = text.split(":");
  if (extra !== undefined) return null;
  const col = /^\$?([A-Za-z]{1,3})$/;
  const row = /^\$?([0-9]{1,7})$/;
  const ca = col.exec(a), cb = col.exec(b);
  if (ca && cb && text.includes(":")) return order({ r1: 0, c1: colIndex(ca[1]), r2: MAX_ROW, c2: colIndex(cb[1]) });
  const ra = row.exec(a), rb = row.exec(b);
  if (ra && rb && text.includes(":")) return order({ r1: Number(ra[1]) - 1, c1: 0, r2: Number(rb[1]) - 1, c2: MAX_COL });
  const p = parseCellRef(a), q = parseCellRef(b);
  return p && q ? order({ r1: p.row, c1: p.col, r2: q.row, c2: q.col }) : null;
}

const order = (r: Rect): Rect => ({ r1: Math.min(r.r1, r.r2), c1: Math.min(r.c1, r.c2), r2: Math.max(r.r1, r.r2), c2: Math.max(r.c1, r.c2) });

export function rangeName(r: Rect): string {
  if (r.r1 === 0 && r.r2 === MAX_ROW) return `${colName(r.c1)}:${colName(r.c2)}`;
  if (r.c1 === 0 && r.c2 === MAX_COL) return `${r.r1 + 1}:${r.r2 + 1}`;
  const a = cellName(r.r1, r.c1);
  return r.r1 === r.r2 && r.c1 === r.c2 ? a : `${a}:${cellName(r.r2, r.c2)}`;
}

export const contains = (r: Rect, row: number, col: number) => row >= r.r1 && row <= r.r2 && col >= r.c1 && col <= r.c2;
export const intersects = (a: Rect, b: Rect) => a.r1 <= b.r2 && b.r1 <= a.r2 && a.c1 <= b.c2 && b.c1 <= a.c2;

/** Sheet name as written in a reference: quoted when it isn't a plain identifier. */
export function quoteSheet(name: string): string {
  const plain = /^[A-Za-z_À-￿][\w.À-￿]*$/.test(name) && !parseCellRef(name) && !/^(R|C|RC|TRUE|FALSE)$/i.test(name);
  return plain ? name : `'${name.replace(/'/g, "''")}'`;
}
