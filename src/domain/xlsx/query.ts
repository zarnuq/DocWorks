/*
 * Google Sheets' QUERY: a subset of the Visualization API query language over a block of rows.
 * select (columns, *, sum / count / avg / min / max), where (comparisons, contains, starts with, ends with,
 * matches, like, is [not] null, and / or / not, parentheses), group by, order by, limit, offset, label.
 * Columns are letters (A, B... relative to the data) or Col1, Col2... for arrays. pivot and format are refused.
 */
import { colIndex } from "./references";
import { Scalar } from "./number-format";

type Agg = "sum" | "count" | "avg" | "min" | "max";
interface Item {
  col: number;
  agg?: Agg;
}
type Cond =
  | { op: "and" | "or"; a: Cond; b: Cond }
  | { op: "not"; a: Cond }
  | { op: "cmp"; cmp: string; left: Operand; right: Operand }
  | { op: "null"; left: Operand; not: boolean };
type Operand = { col: number } | { value: Scalar };

const AGGS = new Set(["sum", "count", "avg", "min", "max"]);
/** A value as text (errors by their code). */
const str = (v: Scalar | undefined): string => (v === null || v === undefined ? "" : typeof v === "object" ? v.error : String(v));

function tokenize(q: string): string[] {
  const out: string[] = [];
  const re = /\s*(?:(\d+(?:\.\d+)?)|'([^']*)'|"([^"]*)"|`([^`]*)`|(<=|>=|!=|<>|[=<>(),*])|([A-Za-z_][A-Za-z0-9_]*))/y;
  let m: RegExpExecArray | null;
  let at = 0;
  while (at < q.length) {
    re.lastIndex = at;
    if (!(m = re.exec(q))) {
      if (!q.slice(at).trim()) break;
      throw new Error(`Can't read the query near "${q.slice(at, at + 10)}"`);
    }
    at = re.lastIndex;
    // Strings keep a marker so they're never taken for keywords or columns.
    out.push(m[1] ?? (m[2] ?? m[3]) !== undefined ? m[1] ?? `'${m[2] ?? m[3]}` : m[4] !== undefined ? m[4] : m[5] ?? m[6]);
  }
  return out;
}

export function runQuery(data: Scalar[][], query: string, headers: number, colNames = false): Scalar[][] {
  const width = Math.max(0, ...data.map((r) => r.length));
  const heads = Array.from({ length: width }, (_, i) => data.slice(0, headers).map((r) => str(r[i])).filter(Boolean).join(" "));
  const rows = data.slice(headers);
  const t = tokenize(query);
  let p = 0;
  const peek = (k = 0) => t[p + k]?.toLowerCase();
  const next = () => t[p++];
  const eat = (word: string) => (peek() === word ? (p++, true) : false);
  const expect = (word: string) => {
    if (!eat(word)) throw new Error(`Expected "${word}" in the query`);
  };
  const column = (): number => {
    const id = next() ?? "";
    const n = colNames ? /^col(\d+)$/i.exec(id) : null;
    const c = n ? Number(n[1]) - 1 : /^[A-Za-z]{1,3}$/.test(id) ? colIndex(id.toUpperCase()) : -1;
    if (c < 0 || c >= width) throw new Error(`Unknown column ${id}`);
    return c;
  };
  const item = (): Item => {
    if (AGGS.has(peek() ?? "") && t[p + 1] === "(") {
      const agg = next().toLowerCase() as Agg;
      next();
      const col = column();
      expect(")");
      return { col, agg };
    }
    return { col: column() };
  };
  const list = <T>(one: () => T): T[] => {
    const out = [one()];
    while (eat(",")) out.push(one());
    return out;
  };
  const operand = (): Operand => {
    const tok = t[p];
    if (tok === undefined) throw new Error("The query ends too early");
    if (tok.startsWith("'")) return (p++, { value: tok.slice(1) });
    if (/^\d/.test(tok)) return (p++, { value: Number(tok) });
    if (/^(true|false)$/i.test(tok)) return (p++, { value: tok.toLowerCase() === "true" });
    return { col: column() };
  };
  const cond = (): Cond => {
    let a = andCond();
    while (eat("or")) a = { op: "or", a, b: andCond() };
    return a;
  };
  const andCond = (): Cond => {
    let a = notCond();
    while (eat("and")) a = { op: "and", a, b: notCond() };
    return a;
  };
  const notCond = (): Cond => {
    if (eat("not")) return { op: "not", a: notCond() };
    if (eat("(")) {
      const c = cond();
      expect(")");
      return c;
    }
    const left = operand();
    if (eat("is")) {
      const not = eat("not");
      expect("null");
      return { op: "null", left, not };
    }
    let cmp = next()?.toLowerCase() ?? "";
    if (cmp === "starts" || cmp === "ends") {
      expect("with");
      cmp += " with";
    }
    if (!["=", "!=", "<>", "<", ">", "<=", ">=", "contains", "starts with", "ends with", "matches", "like"].includes(cmp)) throw new Error(`Unknown comparison "${cmp}"`);
    return { op: "cmp", cmp, left, right: operand() };
  };

  let select: Item[] | null = null;
  let where: Cond | null = null;
  let group: number[] = [];
  let order: (Item & { desc: boolean })[] = [];
  let limit = Infinity, offset = 0;
  const labels = new Map<string, string>();
  const key = (i: Item) => `${i.agg ?? ""}:${i.col}`;
  while (p < t.length) {
    const word = next().toLowerCase();
    if (word === "select") select = eat("*") ? null : list(item);
    else if (word === "where") where = cond();
    else if (word === "group") {
      expect("by");
      group = list(column);
    } else if (word === "order") {
      expect("by");
      order = list(() => {
        const i = item();
        const desc = eat("desc");
        if (!desc) eat("asc");
        return { ...i, desc };
      });
    }
    else if (word === "limit") limit = Number(next());
    else if (word === "offset") offset = Number(next());
    else if (word === "label")
      for (const [i, text] of list(() => [item(), next() ?? ""] as const)) labels.set(key(i), text.replace(/^'/, ""));
    else if (word === "options") p = t.length;
    else throw new Error(`QUERY's "${word}" clause isn't supported here`);
  }

  // Filter.
  const val = (o: Operand, r: Scalar[]): Scalar => ("col" in o ? r[o.col] ?? null : o.value);
  const compare = (a: Scalar, b: Scalar) => (typeof a === "number" && typeof b === "number" ? a - b : str(a).localeCompare(str(b)));
  const test = (c: Cond, r: Scalar[]): boolean => {
    switch (c.op) {
      case "and": return test(c.a, r) && test(c.b, r);
      case "or": return test(c.a, r) || test(c.b, r);
      case "not": return !test(c.a, r);
      case "null": {
        const v = val(c.left, r);
        return (v === null || v === "") !== c.not;
      }
      case "cmp": {
        const a = val(c.left, r), b = val(c.right, r);
        if (a === null || a === "") return c.cmp === "!=" || c.cmp === "<>";
        const s = str(a), q = str(b);
        switch (c.cmp) {
          case "=": return compare(a, b) === 0;
          case "!=": case "<>": return compare(a, b) !== 0;
          case "<": return compare(a, b) < 0;
          case ">": return compare(a, b) > 0;
          case "<=": return compare(a, b) <= 0;
          case ">=": return compare(a, b) >= 0;
          case "contains": return s.includes(q);
          case "starts with": return s.startsWith(q);
          case "ends with": return s.endsWith(q);
          case "matches": return new RegExp(`^(?:${q})$`).test(s);
          default: return new RegExp(`^${q.replace(/[.+?^${}()|[\]\\*]/g, "\\$&").replace(/%/g, ".*").replace(/_/g, ".")}$`).test(s);
        }
      }
    }
  };
  let kept = where ? rows.filter((r) => test(where, r)) : rows;

  // Select and group.
  const items: Item[] = select ?? Array.from({ length: width }, (_, col) => ({ col }));
  const aggregated = items.some((i) => i.agg) || group.length > 0;
  const aggregate = (agg: Agg, vals: Scalar[]): Scalar => {
    const present = vals.filter((v) => v !== null && v !== "");
    if (agg === "count") return present.length;
    const nums = present.filter((v): v is number => typeof v === "number");
    if (agg === "sum") return nums.reduce((a, b) => a + b, 0);
    if (!nums.length) return null;
    if (agg === "avg") return nums.reduce((a, b) => a + b, 0) / nums.length;
    return agg === "min" ? Math.min(...nums) : Math.max(...nums);
  };
  // Each result row also carries every item it might be ordered by.
  const all: Item[] = [...items, ...order.filter((o) => !items.some((i) => key(i) === key(o)))];
  let out: Scalar[][];
  if (aggregated) {
    const groups = new Map<string, Scalar[][]>();
    for (const r of kept) {
      const k = JSON.stringify(group.map((g) => r[g] ?? null));
      groups.set(k, [...(groups.get(k) ?? []), r]);
    }
    if (!group.length && !groups.size) groups.set("[]", []);
    out = [...groups.values()].map((rs) => all.map((i) => (i.agg ? aggregate(i.agg, rs.map((r) => r[i.col] ?? null)) : rs[0]?.[i.col] ?? null)));
    // Groups come out in key order, as Google sorts them.
    if (group.length) out.sort((a, b) => {
      for (const g of group) {
        const at = all.findIndex((i) => !i.agg && i.col === g);
        const d = at < 0 ? 0 : compare(a[at], b[at]);
        if (d) return d;
      }
      return 0;
    });
  } else out = kept.map((r) => all.map((i) => r[i.col] ?? null));
  kept = [];
  if (order.length)
    out.sort((a, b) => {
      for (const o of order) {
        const at = all.findIndex((i) => key(i) === key(o));
        const d = compare(a[at], b[at]);
        if (d) return o.desc ? -d : d;
      }
      return 0;
    });
  out = out.slice(offset, offset + limit).map((r) => r.slice(0, items.length));
  const head = items.map((i) => labels.get(key(i)) ?? (i.agg ? `${i.agg} ${heads[i.col] || `Col${i.col + 1}`}`.trim() : heads[i.col] ?? ""));
  return headers > 0 || aggregated && items.some((i) => i.agg) || labels.size ? [head, ...out] : out;
}
