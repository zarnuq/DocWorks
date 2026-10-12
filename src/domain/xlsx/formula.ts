/*
 * Spreadsheet formulas (as stored in the file: no leading "=", comma separators). Tokens keep their exact text, so
 * reference rewriting (copy, insert/delete rows, sheet renames) changes only the references it has to.
 */
import { MAX_COL, MAX_ROW, Rect, colIndex, colName, quoteSheet } from "./references";

export class FormulaSyntaxError extends Error {}

/** One end of a reference; row/col undefined for whole-column / whole-row references. */
interface RefEnd {
  row?: number;
  col?: number;
  absRow: boolean;
  absCol: boolean;
}

export interface RefToken {
  kind: "ref";
  text: string;
  /** Unquoted sheet name, if qualified. */
  sheet?: string;
  a: RefEnd;
  b?: RefEnd;
}

export type Token =
  | RefToken
  | { kind: "ws" | "num" | "str" | "bool" | "err" | "func" | "name" | "op" | "punct" | "struct"; text: string };

const ERRORS = ["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA", "#SPILL!", "#CALC!"];
const SHEET = String.raw`(?:'((?:[^']|'')+)'|([A-Za-z_À-￿][\w.À-￿]*))!`;
const CELL = String.raw`\$?[A-Za-z]{1,3}\$?\d{1,7}`;
const AREA = String.raw`(?:${CELL}(?::${CELL})?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d{1,7}:\$?\d{1,7})`;
const REF_RE = new RegExp(String.raw`^(?:${SHEET})?(${AREA}|#REF!)(?![\w(.!\[])`);
const RULES: [Token["kind"], RegExp][] = [
  ["ws", /^\s+/],
  ["str", /^"(?:[^"]|"")*"/],
  ["num", /^(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/],
  ["op", /^(?:<>|<=|>=|[-+*/^&=<>%])/],
  ["punct", /^[(),;{}:]/],
];

function refEnd(text: string): RefEnd {
  const m = /^(\$?)([A-Za-z]*)(\$?)(\d*)$/.exec(text)!;
  return { col: m[2] ? colIndex(m[2]) : undefined, row: m[4] ? Number(m[4]) - 1 : undefined, absCol: !!m[2] && m[1] === "$", absRow: m[3] === "$" || (!m[2] && m[1] === "$") };
}

export function tokenize(src: string): Token[] {
  const out: Token[] = [];
  let i = 0;
  while (i < src.length) {
    const rest = src.slice(i);
    let tok: Token | null = null;
    const ref = REF_RE.exec(rest);
    const structured = /^(?:[A-Za-z_\\À-￿][\w.\\?À-￿]*)?\[/.exec(rest);
    if (structured) {
      let depth = 0, j = rest.indexOf("[");
      for (; j < rest.length; j++) {
        if (rest[j] === "'" && j + 1 < rest.length) { j++; continue; }
        if (rest[j] === "[") depth++;
        else if (rest[j] === "]" && --depth === 0) break;
      }
      if (depth) throw new FormulaSyntaxError("Missing ]");
      tok = { kind: "struct", text: rest.slice(0, j + 1) };
    } else if (ref) {
      const sheet = ref[1] !== undefined ? ref[1].replace(/''/g, "'") : ref[2];
      if (ref[3] === "#REF!") tok = { kind: "err", text: ref[0] };
      else {
        const [a, b] = ref[3].split(":");
        tok = { kind: "ref", text: ref[0], sheet, a: refEnd(a), b: b === undefined ? undefined : refEnd(b) };
      }
    } else if (rest[0] === "#") {
      const err = ERRORS.find((e) => rest.toUpperCase().startsWith(e));
      if (!err && out.at(-1)?.kind !== "ref") throw new FormulaSyntaxError(`Unknown error value at ${i + 1}`);
      tok = err ? { kind: "err", text: rest.slice(0, err.length) } : { kind: "op", text: "#" };
    } else {
      for (const [kind, re] of RULES) {
        const m = re.exec(rest);
        if (m) {
          tok = { kind, text: m[0] } as Token;
          break;
        }
      }
      if (!tok) {
        const id = /^[A-Za-z_\\À-￿][\w.\\?À-￿]*/.exec(rest);
        if (!id) throw new FormulaSyntaxError(rest[0] === '"' ? "Missing closing quote" : `Unexpected "${rest[0]}"`);
        let text = id[0];
        if (rest[text.length] === "[") {
          // Structured table reference: Table1[Column], Table1[[#This Row],[Col]].
          let depth = 0, j = text.length;
          for (; j < rest.length; j++) {
            if (rest[j] === "[") depth++;
            else if (rest[j] === "]" && --depth === 0) break;
          }
          if (depth) throw new FormulaSyntaxError("Missing ]");
          tok = { kind: "struct", text: rest.slice(0, j + 1) };
        } else if (/^\s*\(/.test(rest.slice(text.length))) tok = { kind: "func", text };
        else if (/^(TRUE|FALSE)$/i.test(text)) tok = { kind: "bool", text };
        else tok = { kind: "name", text };
        text = tok.text;
      }
    }
    out.push(tok);
    i += tok.text.length;
  }
  return out;
}

export type Node =
  | { type: "num"; value: number }
  | { type: "str"; value: string }
  | { type: "bool"; value: boolean }
  | { type: "err"; value: string }
  | { type: "ref"; sheet?: string; range: Rect; area: boolean }
  | { type: "name"; name: string }
  | { type: "struct"; text: string }
  | { type: "spill"; arg: Node }
  | { type: "call"; name: string; args: Node[] }
  | { type: "un"; op: string; arg: Node }
  | { type: "bin"; op: string; left: Node; right: Node }
  | { type: "array"; rows: (number | string | boolean)[][] }
  | { type: "missing" }
  | { type: "unsupported"; text: string };

export const funcName = (text: string) => text.replace(/^(_xlfn\.|_xlws\.)+/i, "").toUpperCase();

/** Rectangle of a reference token (whole columns / rows span the sheet). */
export function refRect(t: RefToken): Rect {
  const b = t.b ?? t.a;
  const r1 = t.a.row ?? 0, r2 = b.row ?? (t.a.row === undefined ? MAX_ROW : r1);
  const c1 = t.a.col ?? 0, c2 = b.col ?? (t.a.col === undefined ? MAX_COL : c1);
  return { r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) };
}

export function parseFormula(src: string): Node {
  const toks = tokenize(src).filter((t) => t.kind !== "ws");
  let i = 0;
  const peek = () => toks[i];
  const is = (text: string) => toks[i] && (toks[i].kind === "op" || toks[i].kind === "punct") && toks[i].text === text;
  const expect = (text: string) => {
    if (!is(text)) throw new FormulaSyntaxError(`Expected "${text}"`);
    i++;
  };

  const levels = [["=", "<>", "<", ">", "<=", ">="], ["&"], ["+", "-"], ["*", "/"], ["^"]];
  const binary = (level: number): Node => {
    if (level === levels.length) return unary();
    let left = binary(level + 1);
    while (peek()?.kind === "op" && levels[level].includes(peek().text)) {
      const op = toks[i++].text;
      left = { type: "bin", op, left, right: binary(level + 1) };
    }
    return left;
  };
  const unary = (): Node => {
    if (is("-") || is("+")) {
      const op = toks[i++].text;
      return { type: "un", op, arg: unary() };
    }
    let node = primary();
    while (is("%") || is("#")) {
      const op = toks[i++].text;
      node = op === "#" ? { type: "spill", arg: node } : { type: "un", op: "%", arg: node };
    }
    return node;
  };
  const literal = (): number | string | boolean => {
    const neg = is("-") ? (i++, -1) : 1;
    const t = toks[i++];
    if (t?.kind === "num") return neg * Number(t.text);
    if (t?.kind === "str") return t.text.slice(1, -1).replace(/""/g, '"');
    if (t?.kind === "bool") return t.text.toUpperCase() === "TRUE";
    throw new FormulaSyntaxError("Expected a constant in the array");
  };
  const primary = (): Node => {
    const t = toks[i++];
    if (!t) throw new FormulaSyntaxError("Formula ends too early");
    switch (t.kind) {
      case "num":
        return { type: "num", value: Number(t.text) };
      case "str":
        return { type: "str", value: t.text.slice(1, -1).replace(/""/g, '"') };
      case "bool":
        return { type: "bool", value: t.text.toUpperCase() === "TRUE" };
      case "err":
        return { type: "err", value: t.text.slice(t.text.lastIndexOf("#")) };
      case "ref":
        return { type: "ref", sheet: t.sheet, range: refRect(t), area: !!t.b };
      case "name":
        return { type: "name", name: t.text };
      case "struct":
        return { type: "struct", text: t.text };
      case "func": {
        expect("(");
        const args: Node[] = [];
        if (!is(")")) {
          for (;;) {
            args.push(is(",") || is(")") ? { type: "missing" } : binary(0));
            if (is(",")) i++;
            else break;
          }
        }
        expect(")");
        return { type: "call", name: funcName(t.text), args };
      }
    }
    if (t.text === "(") {
      const node = binary(0);
      expect(")");
      return node;
    }
    if (t.text === "{") {
      const rows: (number | string | boolean)[][] = [[]];
      for (;;) {
        rows[rows.length - 1].push(literal());
        if (is(",")) i++;
        else if (is(";")) {
          i++;
          rows.push([]);
        } else break;
      }
      expect("}");
      return { type: "array", rows };
    }
    throw new FormulaSyntaxError(`Unexpected "${t.text}"`);
  };

  const node = binary(0);
  if (i < toks.length) throw new FormulaSyntaxError(`Unexpected "${toks[i].text}"`);
  return node;
}

// ---------------------------------------------------------------------------
// Rewriting references
// ---------------------------------------------------------------------------

function endText(e: RefEnd): string {
  return (e.col !== undefined ? (e.absCol ? "$" : "") + colName(e.col) : "") + (e.row !== undefined ? (e.absRow ? "$" : "") + (e.row + 1) : "");
}

function refText(t: RefToken, a: RefEnd, b?: RefEnd): string {
  const prefix = t.sheet !== undefined ? t.text.slice(0, t.text.lastIndexOf("!") + 1) : "";
  return prefix + endText(a) + (b ? ":" + endText(b) : "");
}

/** Replace each reference token by `fn`'s text (or keep it when `fn` returns null). */
function mapRefs(src: string, fn: (t: RefToken) => string | null): string {
  let toks: Token[];
  try {
    toks = tokenize(src);
  } catch {
    return src;
  }
  return toks.map((t) => (t.kind === "ref" ? fn(t) ?? t.text : t.text)).join("");
}

const inBounds = (e: RefEnd) => (e.row === undefined || (e.row >= 0 && e.row <= MAX_ROW)) && (e.col === undefined || (e.col >= 0 && e.col <= MAX_COL));

/** A formula copied `dr` rows and `dc` columns away: relative references move with it. */
export function translateFormula(src: string, dr: number, dc: number): string {
  if (!dr && !dc) return src;
  return mapRefs(src, (t) => {
    const move = (e: RefEnd): RefEnd => ({ ...e, row: e.row === undefined || e.absRow ? e.row : e.row + dr, col: e.col === undefined || e.absCol ? e.col : e.col + dc });
    const a = move(t.a), b = t.b && move(t.b);
    return inBounds(a) && (!b || inBounds(b)) ? refText(t, a, b) : "#REF!";
  });
}

export interface ShiftOptions {
  /** Sheet whose rows / columns were inserted or deleted. */
  sheet: string;
  axis: "row" | "col";
  /** First inserted / deleted index. */
  at: number;
  /** Positive to insert, negative to delete. */
  count: number;
  /** Sheet the formula lives on (for unqualified references). */
  formulaSheet: string;
}

/** References after rows / columns were inserted or deleted (absolute ones too; deleted cells become #REF!). */
export function shiftFormula(src: string, o: ShiftOptions): string {
  const key = o.axis;
  const n = -o.count, end = o.at + n - 1;
  return mapRefs(src, (t) => {
    if ((t.sheet ?? o.formulaSheet).toLowerCase() !== o.sheet.toLowerCase()) return null;
    const a = { ...t.a }, b = t.b && { ...t.b };
    if (a[key] === undefined) return null; // whole rows when shifting columns, or the reverse
    if (o.count > 0) {
      for (const e of [a, b]) if (e && e[key]! >= o.at) e[key]! += o.count;
      return inBounds(a) && (!b || inBounds(b)) ? refText(t, a, b) : "#REF!";
    }
    const lo = Math.min(a[key], (b ?? a)[key]!), hi = Math.max(a[key], (b ?? a)[key]!);
    if (lo >= o.at && hi <= end) return "#REF!";
    const first = lo < o.at ? lo : lo > end ? lo - n : o.at;
    const last = hi < o.at ? hi : hi > end ? hi - n : o.at - 1;
    const [ea, eb] = a[key] <= (b ?? a)[key]! ? [a, b] : [b!, a];
    ea[key] = first;
    if (eb) eb[key] = last;
    return refText(t, a, b);
  });
}

/** References to sheet `from` renamed to `to`; `to` null turns them into #REF! (deleted sheet). */
export function renameSheetInFormula(src: string, from: string, to: string | null): string {
  return mapRefs(src, (t) => {
    if (t.sheet?.toLowerCase() !== from.toLowerCase()) return null;
    return to === null ? "#REF!" : quoteSheet(to) + "!" + t.text.slice(t.text.lastIndexOf("!") + 1);
  });
}

/** Every reference token (for dependency tracking and highlighting references while editing). */
export function formulaRefs(src: string): RefToken[] {
  try {
    return tokenize(src).filter((t): t is RefToken => t.kind === "ref");
  } catch {
    return [];
  }
}
