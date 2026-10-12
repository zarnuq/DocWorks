import { tableReference } from "./tables";
/*
 * Formula values and functions. Ranges stay references (sheet + rectangle) until a function needs their cells, so
 * SUM(A:A) reads only the cells that exist. Anything not implemented throws Unsupported, so the caller keeps the
 * file's cached result instead of showing a guess.
 */
import { Node, parseFormula } from "./formula";
import { runQuery } from "./query";
import { Rect } from "./references";
import { ErrorValue, Scalar, dateToSerial, formatValue, parseInput, serialToDate } from "./number-format";
import type { Sheet, Workbook } from "./workbook-model";

export interface RangeValue {
  kind: "range";
  sheet: Sheet;
  rect: Rect;
}
export interface Matrix {
  kind: "matrix";
  rows: Scalar[][];
}
export type Value = Scalar | RangeValue | Matrix;

export class Unsupported extends Error {}

export interface EvalContext {
  arrays?: boolean;
  wb: Workbook;
  sheet: Sheet;
  row: number;
  col: number;
  /** A cell's current value (calculating it first if it's waiting to be). */
  value(sheet: Sheet, row: number, col: number): Scalar;
  /** A defined name's formula, as seen from `sheet`. */
  name(name: string, sheet: Sheet): Node | null;
}

const err = (code: string): ErrorValue => ({ error: code });
export const isErr = (v: unknown): v is ErrorValue => typeof v === "object" && v !== null && "error" in v;
const isRange = (v: Value): v is RangeValue => typeof v === "object" && v !== null && "kind" in v && v.kind === "range";
const isMatrix = (v: Value): v is Matrix => typeof v === "object" && v !== null && "kind" in v && v.kind === "matrix";
class ErrorSignal extends Error {
  constructor(readonly value: ErrorValue) {
    super(value.error);
  }
}
/** Stop evaluating and return this error (errors propagate through almost everything). */
const fail = (code: string | ErrorValue): never => {
  throw new ErrorSignal(typeof code === "string" ? err(code) : code);
};

// ---------------------------------------------------------------------------
// Coercion
// ---------------------------------------------------------------------------

/** Numbers as text: up to 15 significant digits, like a spreadsheet. */
export function numText(n: number): string {
  const s = String(Number(n.toPrecision(15)));
  return s.replace("e+", "E+").replace("e-", "E-");
}

function toNum(v: Scalar): number {
  if (typeof v === "number") return v;
  if (v === null) return 0;
  if (typeof v === "boolean") return v ? 1 : 0;
  if (isErr(v)) return fail(v);
  const p = parseInput(v.trim());
  if (typeof p.value === "number" && !p.quoted) return p.value;
  return fail("#VALUE!");
}

function toText(v: Scalar): string {
  if (v === null) return "";
  if (typeof v === "number") return numText(v);
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  if (isErr(v)) return fail(v);
  return v;
}

function toBool(v: Scalar): boolean {
  if (v === null) return false;
  if (typeof v === "boolean") return v;
  if (typeof v === "number") return v !== 0;
  if (isErr(v)) return fail(v);
  if (/^true$/i.test(v)) return true;
  if (/^false$/i.test(v)) return false;
  return fail("#VALUE!");
}

function checked(n: number): number {
  return Number.isFinite(n) ? n : fail("#NUM!");
}

// ---------------------------------------------------------------------------
// Evaluation
// ---------------------------------------------------------------------------

/** Cells of a range, clipped to the sheet's used area (blank beyond it), as rows. */
function matrixOf(v: Value, ctx: EvalContext): Scalar[][] {
  if (isMatrix(v)) return v.rows;
  if (!isRange(v)) return [[v]];
  const used = v.sheet.used();
  const bounded=(v.rect.r2-v.rect.r1+1)*(v.rect.c2-v.rect.c1+1)<=100000;
  const r2 = bounded ? v.rect.r2 : Math.min(v.rect.r2, used ? Math.max(used.r2, v.rect.r1) : v.rect.r1);
  const c2 = bounded ? v.rect.c2 : Math.min(v.rect.c2, used ? Math.max(used.c2, v.rect.c1) : v.rect.c1);
  if((r2-v.rect.r1+1)*(c2-v.rect.c1+1)>100000) throw new Unsupported("Array exceeds 100000 cells");
  const rows: Scalar[][] = [];
  for (let r = v.rect.r1; r <= r2; r++) {
    const row: Scalar[] = [];
    for (let c = v.rect.c1; c <= c2; c++) row.push(ctx.value(v.sheet, r, c));
    rows.push(row);
  }
  return rows;
}

/** Values of existing cells in a range (sparse), or of a matrix / scalar. */
function* valuesOf(v: Value, ctx: EvalContext): Generator<Scalar> {
  if (isRange(v)) {
    for (const cell of v.sheet.cellsIn(v.rect)) yield ctx.value(v.sheet, cell.row, cell.col);
  } else if (isMatrix(v)) for (const row of v.rows) yield* row;
  else yield v;
}

/** One value from a range in a single-value position: the cell in the formula's row or column. */
function scalar(v: Value, ctx: EvalContext): Scalar {
  if (isMatrix(v)) return v.rows[0]?.[0] ?? null;
  if (!isRange(v)) return v;
  const { r1, c1, r2, c2 } = v.rect;
  if (r1 === r2 && c1 === c2) return ctx.value(v.sheet, r1, c1);
  if (c1 === c2 && ctx.row >= r1 && ctx.row <= r2 && v.sheet === ctx.sheet) return ctx.value(v.sheet, ctx.row, c1);
  if (r1 === r2 && ctx.col >= c1 && ctx.col <= c2 && v.sheet === ctx.sheet) return ctx.value(v.sheet, r1, ctx.col);
  return fail("#VALUE!");
}

/** Spreadsheet ordering: numbers < text < booleans; text case-insensitive; blank matches the other side's type. */
export function compareValues(a: Scalar, b: Scalar): number {
  if (isErr(a)) return fail(a);
  if (isErr(b)) return fail(b);
  if (a === null) a = typeof b === "string" ? "" : typeof b === "boolean" ? false : 0;
  if (b === null) b = typeof a === "string" ? "" : typeof a === "boolean" ? false : 0;
  const rank = (v: Scalar) => (typeof v === "number" ? 0 : typeof v === "string" ? 1 : 2);
  if (rank(a) !== rank(b)) return rank(a) - rank(b);
  if (typeof a === "string" && typeof b === "string") {
    const x = a.toLowerCase(), y = b.toLowerCase();
    return x < y ? -1 : x > y ? 1 : 0;
  }
  return Number(a) - Number(b);
}

function binary(op: string, a: Scalar, b: Scalar): Scalar {
  if (isErr(a)) return a;
  if (isErr(b)) return b;
  switch (op) {
    case "&":
      return toText(a) + toText(b);
    case "=": return compareValues(a, b) === 0;
    case "<>": return compareValues(a, b) !== 0;
    case "<": return compareValues(a, b) < 0;
    case ">": return compareValues(a, b) > 0;
    case "<=": return compareValues(a, b) <= 0;
    case ">=": return compareValues(a, b) >= 0;
  }
  const x = toNum(a), y = toNum(b);
  switch (op) {
    case "+": return fix(x + y);
    case "-": return fix(x - y);
    case "*": return checked(x * y);
    case "/": return y === 0 ? err("#DIV/0!") : checked(x / y);
    default: return x === 0 && y < 0 ? err("#DIV/0!") : checked(x ** y);
  }
}

/** Trim binary noise from sums (0.1 + 0.2 -> 0.3) at 15 significant digits. */
const fix = (n: number) => checked(Number(n.toPrecision(15)));

export function evalNode(node: Node, ctx: EvalContext): Value {
  try {
    return evalInner(node, ctx);
  } catch (e) {
    if (e instanceof ErrorSignal) return e.value;
    throw e;
  }
}

function evalInner(node: Node, ctx: EvalContext): Value {
  switch (node.type) {
    case "num": case "str": case "bool":
      return node.value;
    case "err":
      return err(node.value);
    case "missing":
      return null;
    case "array":
      return { kind: "matrix", rows: node.rows };
    case "spill": {
      const v = evalInner(node.arg, ctx);
      if (!isRange(v)) return err("#REF!");
      const cell = v.sheet.at(v.rect.r1, v.rect.c1);
      if (cell && isErr(cell.value)) return cell.value;
      return cell?.array ? { ...v, rect: cell.array } : err("#REF!");
    }
    case "struct":
      return evalInner(tableReference(ctx.wb, ctx.sheet, ctx.row, ctx.col, node.text), ctx);
    case "unsupported":
      throw new Unsupported(node.text);
    case "ref": {
      const sheet = node.sheet === undefined ? ctx.sheet : ctx.wb.sheet(node.sheet);
      if (!sheet) return err("#REF!");
      return { kind: "range", sheet, rect: node.range };
    }
    case "name": {
      const def = ctx.name(node.name, ctx.sheet);
      return evalInner(def ?? tableReference(ctx.wb, ctx.sheet, ctx.row, ctx.col, node.name), ctx);
    }
    case "un": {
      const v = scalar(evalInner(node.arg, ctx), ctx);
      if (isErr(v)) return v;
      if (node.op === "%") return toNum(v) / 100;
      return node.op === "-" ? -toNum(v) : toNum(v);
    }
    case "bin": {
      const a = evalInner(node.left, ctx), b = evalInner(node.right, ctx);
      if (isMatrix(a) || isMatrix(b) || ctx.arrays && (isRange(a) || isRange(b))) {
        const left = matrixOf(a, ctx), right = matrixOf(b, ctx);
        const h = Math.max(left.length, right.length), w = Math.max(left[0]?.length ?? 1, right[0]?.length ?? 1);
        if (h * w > 100000) return err("#NUM!");
        const get = (m: Scalar[][], r: number, c: number) => {
          const row=m[m.length === 1 ? 0 : r], col=m[0]?.length === 1 ? 0 : c;
          return row && col < row.length ? row[col] : err("#N/A");
        };
        return { kind: "matrix", rows: Array.from({length: h}, (_, r) => Array.from({length: w}, (_, c) => binary(node.op, get(left,r,c), get(right,r,c)))) };
      }
      return binary(node.op, scalar(a, ctx), scalar(b, ctx));
    }
    case "call": {
      const fn = FUNCTIONS[node.name];
      if (!fn) throw new Unsupported(node.name);
      const args = node.args.map((a) => () => evalInner(a, ctx));
      return fn(args, ctx, node.args);
    }
  }
}

// ---------------------------------------------------------------------------
// Functions
// ---------------------------------------------------------------------------

type Thunk = () => Value;
type Fn = (args: Thunk[], ctx: EvalContext, nodes: Node[]) => Value;

const arity = (args: Thunk[], min: number, max = min) => {
  if (args.length < min || args.length > max) fail("#N/A");
};
const num = (t: Thunk | undefined, ctx: EvalContext, dflt?: number) => (t ? toNum(scalar(t(), ctx)) : dflt === undefined ? fail("#N/A") : dflt);
const text = (t: Thunk | undefined, ctx: EvalContext, dflt = "") => (t ? toText(scalar(t(), ctx)) : dflt);
const bool = (t: Thunk | undefined, ctx: EvalContext, dflt = false) => (t ? toBool(scalar(t(), ctx)) : dflt);

/** Numbers for SUM-like functions: from ranges only real numbers count; direct arguments are coerced. */
function numbers(args: Thunk[], ctx: EvalContext): number[] {
  const out: number[] = [];
  for (const a of args) {
    const v = a();
    if (isRange(v) || isMatrix(v)) {
      for (const x of valuesOf(v, ctx)) {
        if (isErr(x)) fail(x);
        if (typeof x === "number") out.push(x);
      }
    } else if (v !== null) out.push(toNum(v));
  }
  return out;
}

function roundTo(x: number, d: number, mode: "half" | "up" | "down"): number {
  const f = (n: number) => (mode === "half" ? Math.round(n) : mode === "up" ? Math.ceil(n) : Math.floor(n));
  const shift = (n: number, by: number) => {
    const [m, e = "0"] = String(n).split("e");
    return Number(`${m}e${Number(e) + by}`);
  };
  const a = Math.abs(x);
  // Rounding the shifted value at 15 digits first stops binary noise (1.21*10 = 12.100000000000001) rounding up.
  const shifted = Number(shift(a, d).toPrecision(15));
  return Math.sign(x) * shift(f(shifted), -d);
}

function aggregate(fn: (ns: number[]) => number): Fn {
  return (args, ctx) => checked(fn(numbers(args, ctx)));
}

const wildcard = (pattern: string) =>
  new RegExp("^" + pattern.replace(/~([*?~])|([*?])|([.+^${}()|[\]\\])/g, (_, lit, wild, special) => (lit ? "\\" + lit : wild ? (wild === "*" ? ".*" : ".") : "\\" + special)) + "$", "is");

/** Match test for criteria like ">5", "<>x", "a*", 3. */
function criteria(c: Scalar): (v: Scalar) => boolean {
  if (typeof c !== "string") return (v) => v !== null && !isErr(v) && compareValues(v, c) === 0;
  const m = /^(<=|>=|<>|<|>|=)?(.*)$/s.exec(c)!;
  const op = m[1] ?? "=";
  const rhs = m[2];
  const parsed = parseInput(rhs);
  const n = typeof parsed.value === "number" ? parsed.value : null;
  if (op === "=" || op === "<>") {
    const neg = op === "<>";
    if (rhs === "") return (v) => (v === null || v === "") !== neg;
    const re = wildcard(rhs);
    return (v) => {
      let hit: boolean;
      if (n !== null && typeof v === "number") hit = v === n;
      else if (typeof v === "boolean") hit = /^(true|false)$/i.test(rhs) && v === /^true$/i.test(rhs);
      else hit = typeof v === "string" && re.test(v);
      return hit !== neg;
    };
  }
  return (v) => {
    if (v === null || isErr(v)) return false;
    if (n !== null) {
      if (typeof v !== "number") return false;
    } else if (typeof v !== "string") return false;
    const d = compareValues(v, n ?? rhs);
    return op === "<" ? d < 0 : op === ">" ? d > 0 : op === "<=" ? d <= 0 : d >= 0;
  };
}

/** Rows of `range` matching every (criteria range, criteria) pair; positions are relative to each range. */
function matching(pairs: [Value, Scalar][], ctx: EvalContext): [number, number][] {
  const grids = pairs.map(([r]) => matrixOf(r, ctx));
  const tests = pairs.map(([, c]) => criteria(c));
  const out: [number, number][] = [];
  grids[0].forEach((row, i) =>
    row.forEach((_, j) => {
      if (grids.every((g, k) => tests[k](g[i]?.[j] ?? null))) out.push([i, j]);
    })
  );
  return out;
}

function valueAt(range: Value, i: number, j: number, ctx: EvalContext): Scalar {
  if (isRange(range)) {
    const r = range.rect.r1 + i, c = range.rect.c1 + j;
    return r <= range.rect.r2 && c <= range.rect.c2 ? ctx.value(range.sheet, r, c) : null;
  }
  return isMatrix(range) ? range.rows[i]?.[j] ?? null : i === 0 && j === 0 ? range : null;
}

function conditional(kind: "sum" | "count" | "avg", multi: boolean): Fn {
  return (args, ctx) => {
    let target: Value | null = null;
    const pairs: [Value, Scalar][] = [];
    if (multi) {
      let k = 0;
      if (kind !== "count") target = args[k++]();
      for (; k + 1 < args.length + 1 && args[k]; k += 2) pairs.push([args[k](), scalar(args[k + 1]?.() ?? null, ctx)]);
    } else {
      pairs.push([args[0](), scalar(args[1](), ctx)]);
      target = kind === "count" ? null : args[2] ? args[2]() : args[0]();
    }
    const hits = matching(pairs, ctx);
    if (kind === "count") return hits.length;
    const ns: number[] = [];
    for (const [i, j] of hits) {
      const v = valueAt(target, i, j, ctx);
      if (isErr(v)) fail(v);
      if (typeof v === "number") ns.push(v);
    }
    if (kind === "sum") return fix(ns.reduce((a, b) => a + b, 0));
    return ns.length ? ns.reduce((a, b) => a + b, 0) / ns.length : err("#DIV/0!");
  };
}

/** Index of `x` in `list`: exact (with wildcards for text), or the last value <= x in ascending data. */
function lookupIndex(x: Scalar, list: Scalar[], exact: boolean, wild = true): number {
  if (exact) {
    const re = typeof x === "string" && wild && /[*?~]/.test(x) ? wildcard(x) : null;
    return list.findIndex((v) => (re ? typeof v === "string" && re.test(v) : v !== null && !isErr(v) && typeof v === typeof x && compareValues(v, x) === 0));
  }
  let found = -1;
  for (let i = 0; i < list.length; i++) {
    const v = list[i];
    if (v === null || isErr(v) || typeof v !== typeof x) continue;
    if (compareValues(v, x) <= 0) found = i;
    else break;
  }
  return found;
}

function date(serial: number, ctx: EvalContext) {
  if (serial < 0) fail("#NUM!");
  return serialToDate(serial, ctx.wb.date1904);
}

function nowSerial(ctx: EvalContext): number {
  const d = new Date();
  const day = dateToSerial(d.getFullYear(), d.getMonth() + 1, d.getDate(), ctx.wb.date1904);
  return day + (d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds()) / 86400;
}

function addMonths(serial: number, months: number, ctx: EvalContext, endOfMonth: boolean): number {
  const p = date(serial, ctx);
  const m = p.m - 1 + Math.trunc(months);
  const y = p.y + Math.floor(m / 12), mo = ((m % 12) + 12) % 12 + 1;
  const last = new Date(Date.UTC(y, mo, 0)).getUTCDate();
  return dateToSerial(y, mo, endOfMonth ? last : Math.min(p.d, last), ctx.wb.date1904);
}

function datedif(a: number, b: number, unit: string, ctx: EvalContext): number {
  if (a > b) fail("#NUM!");
  const p = date(a, ctx), q = date(b, ctx);
  let months = (q.y - p.y) * 12 + q.m - p.m - (q.d < p.d ? 1 : 0);
  switch (unit.toUpperCase()) {
    case "Y": return Math.floor(months / 12);
    case "M": return months;
    case "D": return Math.floor(b) - Math.floor(a);
    case "MD": return q.d >= p.d ? q.d - p.d : Math.floor(b) - addMonths(Math.floor(a), months, ctx, false);
    case "YM": return months % 12;
    case "YD": {
      months = (q.y - p.y) * 12;
      let start = addMonths(Math.floor(a), months, ctx, false);
      if (start > b) start = addMonths(Math.floor(a), months - 12, ctx, false);
      return Math.floor(b) - start;
    }
    default: return fail("#NUM!");
  }
}

const VOLATILE = new Set(["TODAY", "NOW", "RAND", "RANDBETWEEN", "OFFSET", "INDIRECT"]);
export const isVolatile = (name: string) => VOLATILE.has(name);

/** Apply `f` once, or element by element when an argument is an array (in ARRAYFORMULA), errors staying per cell. */
function lifted(args: Thunk[], ctx: EvalContext, f: (vals: Scalar[]) => Scalar): Value {
  const vs = args.map((a) => a());
  const arrays = vs.map((v) => (isMatrix(v) || (ctx.arrays && isRange(v)) ? matrixOf(v, ctx) : null));
  const one = (vals: Scalar[]) => {
    try {
      return f(vals);
    } catch (e) {
      if (e instanceof ErrorSignal) return e.value;
      throw e;
    }
  };
  if (!arrays.some(Boolean)) return f(vs.map((v) => scalar(v, ctx)));
  const h = Math.max(...arrays.map((m) => m?.length ?? 1)), w = Math.max(...arrays.map((m) => m?.[0]?.length ?? 1));
  const at = (i: number, r: number, c: number) => {
    const m = arrays[i];
    if (!m) return scalar(vs[i], ctx);
    const row = m[m.length === 1 ? 0 : r];
    return row?.[row.length === 1 ? 0 : c] ?? null;
  };
  return { kind: "matrix", rows: Array.from({ length: h }, (_, r) => Array.from({ length: w }, (_, c) => one(vs.map((_, i) => at(i, r, c))))) };
}

/** A regular expression from a formula (RE2 syntax is close enough to JavaScript's for common patterns). */
function regex(pattern: string, flags = ""): RegExp {
  try {
    return new RegExp(pattern, flags);
  } catch {
    return fail("#REF!");
  }
}

const FUNCTIONS: Record<string, Fn> = {
  // Google Sheets: work the argument out over whole ranges (arithmetic and lifted functions go element by element).
  ARRAYFORMULA: (a, c, nodes) => (arity(a, 1), evalInner(nodes[0], { ...c, arrays: true })),
  // Google's export of its own functions: IFERROR(__xludf.DUMMYFUNCTION("REGEXMATCH(...)"), cached result).
  "__XLUDF.DUMMYFUNCTION": (a, c) => evalInner(parseFormula(text(a[0], c)), c),
  REGEXMATCH: (a, c) => (arity(a, 2), lifted(a, c, ([t, re]) => regex(toText(re)).test(toText(t)))),
  REGEXTEST: (a, c) => (arity(a, 2, 3), lifted(a, c, ([t, re, ci]) => regex(toText(re), ci !== null && ci !== undefined && toBool(ci) ? "i" : "").test(toText(t)))),
  REGEXEXTRACT: (a, c) => (arity(a, 2, 4), lifted(a.slice(0, 2), c, ([t, re]) => {
    const m = regex(toText(re)).exec(toText(t));
    return m ? m[1] ?? m[0] : fail("#N/A");
  })),
  REGEXREPLACE: (a, c) => (arity(a, 3, 5), lifted(a.slice(0, 3), c, ([t, re, to]) => toText(t).replace(regex(toText(re), "g"), toText(to)))),
  QUERY: (a, c, nodes) => {
    arity(a, 1, 3);
    const data = matrixOf(a[0](), c);
    let headers = a[2] ? Math.trunc(num(a[2], c)) : -1;
    // -1: Google guesses one header row when the first row is text and a later one isn't.
    if (headers < 0) headers = data.length > 1 && data[0].every((v) => typeof v === "string") && data.slice(1).some((r) => r.some((v) => typeof v !== "string" && v !== null)) ? 1 : 0;
    try {
      const rows = runQuery(data, a[1] ? text(a[1], c) : "", headers, nodes[0]?.type === "array");
      return rows.length && rows[0].length ? { kind: "matrix", rows } : err("#N/A");
    } catch (e) {
      if (e instanceof ErrorSignal) throw e;
      return err("#VALUE!");
    }
  },
  // Drawn by the grid (see sparklines.ts); the cell itself shows nothing.
  SPARKLINE: () => "",

  ANCHORARRAY: (a,c) => {
    const v=a[0]();
    if(!isRange(v)) return err("#REF!");
    const value=c.value(v.sheet,v.rect.r1,v.rect.c1);
    if(isErr(value)) return value;
    const cell=v.sheet.at(v.rect.r1,v.rect.c1);
    return cell?.array ? {...v,rect:cell.array} : err("#REF!");
  },
  SEQUENCE: (a,c) => {
    const h = Math.trunc(num(a[0],c,1)), w = Math.trunc(num(a[1],c,1)), start = num(a[2],c,1), step = num(a[3],c,1);
    if (h < 1 || w < 1 || h * w > 100000) return err("#NUM!");
    return {kind: "matrix", rows: Array.from({length:h}, (_,r) => Array.from({length:w}, (_,col) => start+(r*w+col)*step))};
  },
  TRANSPOSE: (a,c) => {
    const rows = matrixOf(a[0](),c);
    return {kind: "matrix", rows: (rows[0] ?? []).map((_,i) => rows.map(r => r[i] ?? null))};
  },
  FILTER: (a,c) => {
    const rows = matrixOf(a[0](),c), mask = matrixOf(a[1](),{...c, arrays:true});
    let result: Scalar[][];
    if (mask.length === rows.length && mask.every(r => r.length === 1)) result = rows.filter((_,i) => toBool(mask[i][0]));
    else if (mask.length === 1 && mask[0].length === rows[0]?.length) result = rows.map(r => r.filter((_,i) => toBool(mask[0][i])));
    else return err("#VALUE!");
    return result.length && result[0].length ? {kind:"matrix",rows:result} : a[2] ? a[2]() : err("#CALC!");
  },
  SORT: (a,c) => {
    const rows = matrixOf(a[0](),c), index = Math.trunc(num(a[1],c,1))-1, order = num(a[2],c,1), byCol = a[3] ? toBool(scalar(a[3](),c)) : false;
    if (![1,-1].includes(order) || index < 0 || index >= (byCol ? rows.length : rows[0]?.length ?? 0)) return err("#VALUE!");
    if (byCol) {
      const cols = (rows[0] ?? []).map((_,i) => i).sort((i,j) => order * compareValues(rows[index][i],rows[index][j]));
      return {kind:"matrix",rows:rows.map(r => cols.map(i => r[i]))};
    }
    return {kind:"matrix",rows:[...rows].sort((a,b) => order * compareValues(a[index],b[index]))};
  },
  UNIQUE: (a,c) => {
    const input = matrixOf(a[0](),c), byCol = a[1] ? toBool(scalar(a[1](),c)) : false, once = a[2] ? toBool(scalar(a[2](),c)) : false;
    const rows = byCol ? (input[0] ?? []).map((_,i) => input.map(r => r[i])) : input;
    const key = (r: Scalar[]) => JSON.stringify(r.map(v => typeof v === "string" ? v.toLowerCase() : v));
    const counts = new Map<string,number>();
    for (const r of rows) counts.set(key(r),(counts.get(key(r)) ?? 0)+1);
    const seen = new Set<string>();
    const out = rows.filter(r => {const k=key(r); if (seen.has(k) || once && counts.get(k)! > 1) return false; seen.add(k);return true;});
    if (!out.length) return err("#CALC!");
    return {kind:"matrix",rows:byCol ? out[0].map((_,i) => out.map(r => r[i])) : out};
  },
  SUM: aggregate((ns) => fix(ns.reduce((a, b) => a + b, 0))),
  PRODUCT: aggregate((ns) => ns.reduce((a, b) => a * b, 1)),
  AVERAGE: (args, ctx) => {
    const ns = numbers(args, ctx);
    return ns.length ? ns.reduce((a, b) => a + b, 0) / ns.length : err("#DIV/0!");
  },
  MIN: aggregate((ns) => (ns.length ? Math.min(...ns) : 0)),
  MAX: aggregate((ns) => (ns.length ? Math.max(...ns) : 0)),
  MEDIAN: (args, ctx) => {
    const ns = numbers(args, ctx).sort((a, b) => a - b);
    if (!ns.length) return err("#NUM!");
    const m = ns.length >> 1;
    return ns.length % 2 ? ns[m] : (ns[m - 1] + ns[m]) / 2;
  },
  COUNT: (args, ctx) => {
    let n = 0;
    for (const a of args) {
      const v = a();
      if (isRange(v) || isMatrix(v)) for (const x of valuesOf(v, ctx)) n += typeof x === "number" ? 1 : 0;
      else if (typeof v === "number" || (typeof v === "string" && typeof parseInput(v).value === "number") || typeof v === "boolean") n++;
    }
    return n;
  },
  COUNTA: (args, ctx) => {
    let n = 0;
    for (const a of args) for (const x of valuesOf(a(), ctx)) n += x !== null ? 1 : 0;
    return n;
  },
  COUNTBLANK: (args, ctx) => {
    const v = args[0]();
    if (!isRange(v)) return matrixOf(v, ctx).flat().filter((x) => x === null || x === "").length;
    const { r1, c1, r2, c2 } = v.rect;
    let filled = 0;
    for (const cell of v.sheet.cellsIn(v.rect)) {
      const x = ctx.value(v.sheet, cell.row, cell.col);
      if (x !== null && x !== "") filled++;
    }
    return (r2 - r1 + 1) * (c2 - c1 + 1) - filled;
  },
  SUMIF: conditional("sum", false),
  SUMIFS: conditional("sum", true),
  COUNTIF: conditional("count", false),
  COUNTIFS: conditional("count", true),
  AVERAGEIF: conditional("avg", false),
  AVERAGEIFS: conditional("avg", true),
  SUMPRODUCT: (args, ctx) => {
    const grids = args.map((a) => matrixOf(a(), ctx));
    let total = 0;
    grids[0].forEach((row, i) =>
      row.forEach((_, j) => {
        total += grids.reduce((p, g) => {
          const v = g[i]?.[j] ?? null;
          if (isErr(v)) fail(v);
          return p * (typeof v === "number" ? v : 0);
        }, 1);
      })
    );
    return fix(total);
  },
  ROUND: (a, c) => roundTo(num(a[0], c), num(a[1], c, 0), "half"),
  ROUNDUP: (a, c) => roundTo(num(a[0], c), num(a[1], c, 0), "up"),
  ROUNDDOWN: (a, c) => roundTo(num(a[0], c), num(a[1], c, 0), "down"),
  INT: (a, c) => Math.floor(num(a[0], c)),
  TRUNC: (a, c) => roundTo(num(a[0], c), num(a[1], c, 0), "down"),
  ABS: (a, c) => Math.abs(num(a[0], c)),
  SIGN: (a, c) => Math.sign(num(a[0], c)),
  MOD: (a, c) => {
    const x = num(a[0], c), y = num(a[1], c);
    return y === 0 ? err("#DIV/0!") : fix(x - y * Math.floor(x / y));
  },
  POWER: (a, c) => binary("^", num(a[0], c), num(a[1], c)),
  SQRT: (a, c) => {
    const x = num(a[0], c);
    return x < 0 ? err("#NUM!") : Math.sqrt(x);
  },
  EXP: (a, c) => checked(Math.exp(num(a[0], c))),
  LN: (a, c) => {
    const x = num(a[0], c);
    return x <= 0 ? err("#NUM!") : Math.log(x);
  },
  LOG10: (a, c) => {
    const x = num(a[0], c);
    return x <= 0 ? err("#NUM!") : Math.log10(x);
  },
  LOG: (a, c) => {
    const x = num(a[0], c), b = num(a[1], c, 10);
    return x <= 0 || b <= 0 || b === 1 ? err("#NUM!") : Math.log(x) / Math.log(b);
  },
  PI: () => Math.PI,
  CEILING: (a, c) => {
    const x = num(a[0], c), s = num(a[1], c, 1);
    return s === 0 ? 0 : fix(Math.ceil(x / s) * s);
  },
  FLOOR: (a, c) => {
    const x = num(a[0], c), s = num(a[1], c, 1);
    return s === 0 ? err("#DIV/0!") : fix(Math.floor(x / s) * s);
  },
  RAND: () => Math.random(),
  RANDBETWEEN: (a, c) => {
    const lo = Math.ceil(num(a[0], c)), hi = Math.floor(num(a[1], c));
    return lo > hi ? err("#NUM!") : lo + Math.floor(Math.random() * (hi - lo + 1));
  },

  IF: (a, c) => {
    arity(a, 1, 3);
    const cond = scalar(a[0](), c);
    if (isErr(cond)) return cond;
    return toBool(cond) ? (a[1] ? a[1]() : true) : a[2] ? a[2]() : false;
  },
  IFS: (a, c) => {
    for (let i = 0; i + 1 < a.length; i += 2) if (toBool(scalar(a[i](), c))) return a[i + 1]();
    return err("#N/A");
  },
  SWITCH: (a, c) => {
    const x = scalar(a[0](), c);
    let i = 1;
    for (; i + 1 < a.length; i += 2) if (compareValues(x, scalar(a[i](), c)) === 0) return a[i + 1]();
    return i < a.length ? a[i]() : err("#N/A");
  },
  IFERROR: (a, c) => {
    const v = evalSafe(a[0], c);
    return isErr(v) ? a[1]() : v;
  },
  IFNA: (a, c) => {
    const v = evalSafe(a[0], c);
    return isErr(v) && v.error === "#N/A" ? a[1]() : v;
  },
  AND: (a, c) => logical(a, c).every(Boolean),
  OR: (a, c) => logical(a, c).some(Boolean),
  XOR: (a, c) => logical(a, c).filter(Boolean).length % 2 === 1,
  NOT: (a, c) => !bool(a[0], c),
  TRUE: () => true,
  FALSE: () => false,

  CONCATENATE: (a, c) => a.map((t) => text(t, c)).join(""),
  CONCAT: (a, c) => a.map((t) => Array.from(valuesOf(t(), c)).map(toText).join("")).join(""),
  TEXTJOIN: (a, c) => {
    const sep = text(a[0], c), skip = bool(a[1], c);
    const parts: string[] = [];
    for (const t of a.slice(2)) {
      const v = t();
      for (const x of isRange(v) || isMatrix(v) ? matrixOf(v, c).flat() : [scalar(v, c)]) {
        const s = toText(x);
        if (s || !skip) parts.push(s);
      }
    }
    return parts.join(sep);
  },
  LEN: (a, c) => text(a[0], c).length,
  LEFT: (a, c) => text(a[0], c).slice(0, Math.max(0, num(a[1], c, 1))),
  RIGHT: (a, c) => {
    const s = text(a[0], c), n = num(a[1], c, 1);
    return n <= 0 ? "" : s.slice(-n);
  },
  MID: (a, c) => {
    const start = num(a[1], c), n = num(a[2], c);
    return start < 1 || n < 0 ? err("#VALUE!") : text(a[0], c).slice(start - 1, start - 1 + n);
  },
  UPPER: (a, c) => text(a[0], c).toUpperCase(),
  LOWER: (a, c) => text(a[0], c).toLowerCase(),
  PROPER: (a, c) => text(a[0], c).toLowerCase().replace(/(^|[^a-z])([a-z])/g, (_, p: string, ch: string) => p + ch.toUpperCase()),
  TRIM: (a, c) => text(a[0], c).trim().replace(/ +/g, " "),
  SUBSTITUTE: (a, c) => {
    const s = text(a[0], c), from = text(a[1], c), to = text(a[2], c);
    if (!from) return s;
    if (!a[3]) return s.split(from).join(to);
    const n = num(a[3], c);
    let i = -1;
    for (let k = 0; k < n; k++) if ((i = s.indexOf(from, i + 1)) < 0) return s;
    return s.slice(0, i) + to + s.slice(i + from.length);
  },
  REPLACE: (a, c) => {
    const s = text(a[0], c), start = num(a[1], c), n = num(a[2], c);
    return s.slice(0, start - 1) + text(a[3], c) + s.slice(start - 1 + n);
  },
  FIND: (a, c) => {
    const i = text(a[1], c).indexOf(text(a[0], c), num(a[2], c, 1) - 1);
    return i < 0 ? err("#VALUE!") : i + 1;
  },
  SEARCH: (a, c) => {
    const hay = text(a[1], c), start = num(a[2], c, 1) - 1;
    const re = wildcard(text(a[0], c));
    for (let i = start; i < hay.length; i++) for (let j = hay.length; j > i; j--) if (re.test(hay.slice(i, j))) return i + 1;
    return err("#VALUE!");
  },
  TEXT: (a, c) => {
    const v = scalar(a[0](), c);
    return formatValue(typeof v === "string" && typeof parseInput(v).value === "number" ? (parseInput(v).value) : v, text(a[1], c), c.wb.date1904).text;
  },
  VALUE: (a, c) => {
    const v = scalar(a[0](), c);
    if (typeof v === "number") return v;
    const p = parseInput(toText(v).trim());
    return typeof p.value === "number" ? p.value : err("#VALUE!");
  },
  REPT: (a, c) => text(a[0], c).repeat(Math.max(0, num(a[1], c))),
  EXACT: (a, c) => text(a[0], c) === text(a[1], c),
  T: (a, c) => {
    const v = scalar(a[0](), c);
    return typeof v === "string" ? v : "";
  },
  N: (a, c) => {
    const v = scalar(a[0](), c);
    return typeof v === "number" ? v : typeof v === "boolean" ? Number(v) : 0;
  },
  CHAR: (a, c) => String.fromCharCode(num(a[0], c)),
  CODE: (a, c) => text(a[0], c).charCodeAt(0) || err("#VALUE!"),

  VLOOKUP: (a, c) => lookup(a, c, false),
  HLOOKUP: (a, c) => lookup(a, c, true),
  INDEX: (a, c) => {
    const v = a[0]();
    const r = num(a[1], c, 0), col = num(a[2], c, 0);
    if (isRange(v)) {
      const { r1, c1, r2, c2 } = v.rect;
      const single = r1 === r2 ? "row" : c1 === c2 ? "col" : null;
      // INDEX(A1:A5, 3) picks the 3rd cell even in a one-column range.
      const [ri, ci] = a.length === 2 && single === "row" ? [1, r] : [r, col || (single ? 1 : 0)];
      if (ri > r2 - r1 + 1 || ci > c2 - c1 + 1 || ri < 0 || ci < 0) return err("#REF!");
      return { kind: "range", sheet: v.sheet, rect: { r1: ri ? r1 + ri - 1 : r1, r2: ri ? r1 + ri - 1 : r2, c1: ci ? c1 + ci - 1 : c1, c2: ci ? c1 + ci - 1 : c2 } };
    }
    const rows = matrixOf(v, c);
    const row = rows[Math.max(0, r - 1)];
    return row?.[Math.max(0, (col || 1) - 1)] ?? err("#REF!");
  },
  MATCH: (a, c) => {
    const x = scalar(a[0](), c);
    const list = matrixOf(a[1](), c).flat();
    const type = num(a[2], c, 1);
    if (type === -1) {
      let found = -1;
      for (let i = 0; i < list.length; i++) if (list[i] !== null && compareValues(list[i], x) >= 0) found = i;
      return found < 0 ? err("#N/A") : found + 1;
    }
    const i = lookupIndex(x, list, type === 0);
    return i < 0 ? err("#N/A") : i + 1;
  },
  XLOOKUP: (a, c) => {
    const x = scalar(a[0](), c);
    const keys = matrixOf(a[1](), c);
    const vertical = keys.length > 1 || keys[0]?.length === 1;
    const list = vertical ? keys.map((r) => r[0]) : keys[0];
    const mode = num(a[4], c, 0);
    let i = lookupIndex(x, list, true, mode === 2);
    if (i < 0 && (mode === -1 || mode === 1)) {
      let best = -1;
      list.forEach((v, k) => {
        if (v === null || isErr(v) || typeof v !== typeof x) return;
        const d = compareValues(v, x);
        if ((mode === -1 && d < 0 && (best < 0 || compareValues(v, list[best]) > 0)) || (mode === 1 && d > 0 && (best < 0 || compareValues(v, list[best]) < 0))) best = k;
      });
      i = best;
    }
    if (i < 0) return a[3] ? a[3]() : err("#N/A");
    const ret = a[2]();
    return vertical ? valueAt(ret, i, 0, c) : valueAt(ret, 0, i, c);
  },
  CHOOSE: (a, c) => {
    const i = Math.trunc(num(a[0], c));
    return i < 1 || i >= a.length ? err("#VALUE!") : a[i]();
  },
  ROW: (a, c) => (a[0] ? rangeOf(a[0]()).rect.r1 + 1 : c.row + 1),
  COLUMN: (a, c) => (a[0] ? rangeOf(a[0]()).rect.c1 + 1 : c.col + 1),
  ROWS: (a, c) => {
    const v = a[0]();
    return isRange(v) ? v.rect.r2 - v.rect.r1 + 1 : matrixOf(v, c).length;
  },
  COLUMNS: (a, c) => {
    const v = a[0]();
    return isRange(v) ? v.rect.c2 - v.rect.c1 + 1 : matrixOf(v, c)[0]?.length ?? 1;
  },
  OFFSET: (a, c) => {
    const base = rangeOf(a[0]());
    const dr = num(a[1], c), dc = num(a[2], c);
    const h = num(a[3], c, base.rect.r2 - base.rect.r1 + 1), w = num(a[4], c, base.rect.c2 - base.rect.c1 + 1);
    const r1 = base.rect.r1 + dr, c1 = base.rect.c1 + dc;
    if (r1 < 0 || c1 < 0 || h < 1 || w < 1) return err("#REF!");
    return { kind: "range", sheet: base.sheet, rect: { r1, c1, r2: r1 + h - 1, c2: c1 + w - 1 } };
  },
  INDIRECT: (a, c) => {
    const ref = text(a[0], c);
    try {
      const node = parseFormula(ref);
      if (node.type === "ref") return evalInner(node, c);
    } catch {
      // not a reference
    }
    return err("#REF!");
  },

  DATE: (a, c) => {
    let y = Math.trunc(num(a[0], c));
    if (y < 1900) y += 1900;
    const serial = dateToSerial(y, Math.trunc(num(a[1], c)), Math.trunc(num(a[2], c)), c.wb.date1904);
    return serial < 0 ? err("#NUM!") : serial;
  },
  TIME: (a, c) => {
    const s = (num(a[0], c) * 3600 + num(a[1], c) * 60 + num(a[2], c)) / 86400;
    return s < 0 ? err("#NUM!") : s - Math.floor(s);
  },
  YEAR: (a, c) => date(num(a[0], c), c).y,
  MONTH: (a, c) => date(num(a[0], c), c).m,
  DAY: (a, c) => date(num(a[0], c), c).d,
  HOUR: (a, c) => date(num(a[0], c), c).h,
  MINUTE: (a, c) => date(num(a[0], c), c).mi,
  SECOND: (a, c) => date(num(a[0], c), c).s,
  WEEKDAY: (a, c) => {
    const dow = date(num(a[0], c), c).dow;
    const type = num(a[1], c, 1);
    return type === 2 ? ((dow + 6) % 7) + 1 : type === 3 ? (dow + 6) % 7 : dow + 1;
  },
  TODAY: (_, c) => Math.floor(nowSerial(c)),
  NOW: (_, c) => nowSerial(c),
  EDATE: (a, c) => addMonths(Math.floor(num(a[0], c)), num(a[1], c), c, false),
  EOMONTH: (a, c) => addMonths(Math.floor(num(a[0], c)), num(a[1], c), c, true),
  DATEDIF: (a, c) => datedif(num(a[0], c), num(a[1], c), text(a[2], c), c),
  DAYS: (a, c) => Math.floor(num(a[0], c)) - Math.floor(num(a[1], c)),

  ISBLANK: (a, c) => scalarSafe(a[0], c) === null,
  ISNUMBER: (a, c) => typeof scalarSafe(a[0], c) === "number",
  ISTEXT: (a, c) => typeof scalarSafe(a[0], c) === "string",
  ISLOGICAL: (a, c) => typeof scalarSafe(a[0], c) === "boolean",
  ISERROR: (a, c) => isErr(scalarSafe(a[0], c)),
  ISERR: (a, c) => {
    const v = scalarSafe(a[0], c);
    return isErr(v) && v.error !== "#N/A";
  },
  ISNA: (a, c) => {
    const v = scalarSafe(a[0], c);
    return isErr(v) && v.error === "#N/A";
  },
  NA: () => err("#N/A"),
};

function rangeOf(v: Value): RangeValue {
  return isRange(v) ? v : fail(isErr(v) ? v : "#VALUE!");
}

function evalSafe(t: Thunk, ctx: EvalContext): Value {
  try {
    const v = t();
    return isRange(v) || isMatrix(v) ? scalar(v, ctx) : v;
  } catch (e) {
    if (e instanceof ErrorSignal) return e.value;
    throw e;
  }
}

function scalarSafe(t: Thunk, ctx: EvalContext): Scalar {
  return evalSafe(t, ctx) as Scalar;
}

function logical(args: Thunk[], ctx: EvalContext): boolean[] {
  const out: boolean[] = [];
  for (const a of args) {
    const v = a();
    if (isRange(v) || isMatrix(v)) {
      for (const x of valuesOf(v, ctx)) {
        if (isErr(x)) fail(x);
        if (typeof x === "number" || typeof x === "boolean") out.push(Boolean(x));
      }
    } else out.push(toBool(v));
  }
  return out.length ? out : fail("#VALUE!");
}

function lookup(a: Thunk[], c: EvalContext, horizontal: boolean): Value {
  const x = scalar(a[0](), c);
  if (isErr(x)) return x;
  const table = matrixOf(a[1](), c);
  const n = num(a[2], c);
  const exact = a[3] ? !bool(a[3], c) : false;
  const keys = horizontal ? table[0] ?? [] : table.map((r) => r[0]);
  if (n < 1 || n > (horizontal ? table.length : table[0]?.length ?? 0)) return err("#REF!");
  const i = lookupIndex(x, keys, exact);
  if (i < 0) return err("#N/A");
  return (horizontal ? table[n - 1][i] : table[i][n - 1]) ?? null;
}

/** Evaluate a parsed formula to the value a cell shows (single values; ranges show their top-left cell). */
export function evaluate(node: Node, ctx: EvalContext): Scalar {
  let v = evalNode(node, ctx);
  if (isRange(v)) {
    const range = v;
    try {
      v = scalar(range, ctx);
    } catch (e) {
      if (!(e instanceof ErrorSignal)) throw e;
      v = ctx.value(range.sheet, range.rect.r1, range.rect.c1);
    }
  }
  if (isMatrix(v)) v = v.rows[0]?.[0] ?? null;
  // A formula pointing at a blank cell shows 0, like a spreadsheet.
  return v ?? 0;
}

/** Full result for a dynamic array anchor; bounded ranges become rows without losing their shape. */
export function evaluateArray(node: Node, ctx: EvalContext): Scalar[][] {
  const v = evalNode(node, {...ctx, arrays:true});
  return matrixOf(v,ctx);
}
