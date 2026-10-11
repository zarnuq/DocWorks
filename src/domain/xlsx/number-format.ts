/*
 * Number format codes ("#,##0.00", "yyyy-mm-dd", "[Red]0;-0") applied to cell values, and typed input read the
 * way Google Sheets reads it. Dates are serial day numbers (1900 or 1904 system), computed in UTC so no timezone
 * ever shifts them.
 */

export type ErrorValue = { error: string };
export type Scalar = number | string | boolean | ErrorValue | null;

const BUILTIN: Record<number, string> = {
  0: "General", 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00", 9: "0%", 10: "0.00%", 11: "0.00E+00", 12: "# ?/?", 13: "# ??/??",
  14: "m/d/yyyy", 15: "d-mmm-yy", 16: "d-mmm", 17: "mmm-yy", 18: "h:mm AM/PM", 19: "h:mm:ss AM/PM", 20: "h:mm", 21: "h:mm:ss",
  22: "m/d/yyyy h:mm", 37: "#,##0 ;(#,##0)", 38: "#,##0 ;[Red](#,##0)", 39: "#,##0.00;(#,##0.00)", 40: "#,##0.00;[Red](#,##0.00)",
  45: "mm:ss", 46: "[h]:mm:ss", 47: "mmss.0", 48: "##0.0E+0", 49: "@",
};

export const builtinFormat = (id: number): string => BUILTIN[id] ?? "General";

const COLORS: Record<string, string> = {
  black: "#000000", blue: "#0000FF", cyan: "#00FFFF", green: "#00FF00", magenta: "#FF00FF", red: "#FF0000", white: "#FFFFFF", yellow: "#FFFF00",
};
const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const DAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

// ---------------------------------------------------------------------------
// Dates
// ---------------------------------------------------------------------------

export interface DateParts { y: number; m: number; d: number; h: number; mi: number; s: number; ms: number; dow: number }

const DAY_MS = 86400000;
const EPOCH_1900 = Date.UTC(1899, 11, 30);
const EPOCH_1904 = Date.UTC(1904, 0, 1);

export function serialToDate(serial: number, date1904 = false): DateParts {
  const totalMs = Math.round(serial * DAY_MS);
  let days = Math.floor(totalMs / DAY_MS);
  const msOfDay = totalMs - days * DAY_MS;
  const time = { h: Math.floor(msOfDay / 3600000), mi: Math.floor(msOfDay / 60000) % 60, s: Math.floor(msOfDay / 1000) % 60, ms: msOfDay % 1000 };
  // 1900 system: serial 60 is Lotus's 1900-02-29, which never existed; serials before it are one day later than the epoch suggests.
  if (!date1904 && days === 60) return { y: 1900, m: 2, d: 29, ...time, dow: 3 };
  if (!date1904 && days < 60) days += 1;
  const dt = new Date((date1904 ? EPOCH_1904 : EPOCH_1900) + days * DAY_MS);
  return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate(), ...time, dow: dt.getUTCDay() };
}

export function dateToSerial(y: number, m: number, d: number, date1904 = false): number {
  const days = Math.round((Date.UTC(y, m - 1, d) - (date1904 ? EPOCH_1904 : EPOCH_1900)) / DAY_MS);
  return !date1904 && days < 61 ? days - 1 : days;
}

// ---------------------------------------------------------------------------
// Format codes
// ---------------------------------------------------------------------------

interface Section {
  tokens: string[];
  color?: string;
  cond?: { op: string; value: number };
}

const cache = new Map<string, Section[]>();

/** Split a code into sections; each section into tokens: quoted / escaped literals, placeholders, date parts. */
function parseCode(code: string): Section[] {
  const hit = cache.get(code);
  if (hit) return hit;
  const sections: Section[] = [{ tokens: [] }];
  const tokenRe = /"[^"]*"|\\.|_.|\*.|\[[^\]]*\]|E[+-]|AM\/PM|A\/P|y+|m+|d+|h+|s+(?:\.0+)?|[0#?]|[.,%@;/]|./gi;
  for (const [tok] of code.matchAll(tokenRe)) {
    const sec = sections[sections.length - 1];
    if (tok === ";") sections.push({ tokens: [] });
    else if (tok[0] === "[") {
      const inner = tok.slice(1, -1);
      const cond = /^(<=|>=|<>|<|>|=)(-?[\d.]+)$/.exec(inner);
      if (cond) sec.cond = { op: cond[1], value: Number(cond[2]) };
      else if (COLORS[inner.toLowerCase()]) sec.color = COLORS[inner.toLowerCase()];
      else if (/^\$/.test(inner)) sec.tokens.push(`"${inner.slice(1).split("-")[0]}"`);
      else if (/^(h+|m+|s+)$/i.test(inner)) sec.tokens.push(tok.toLowerCase());
    } else sec.tokens.push(tok);
  }
  cache.set(code, sections);
  return sections;
}

const isDateToken = (t: string) => /^(y+|m+|d+|h+|s+(\.0+)?|AM\/PM|A\/P|\[[hms]+\])$/i.test(t);
export const isDateFormat = (code: string) => code !== "General" && parseCode(code).some((s) => s.tokens.some(isDateToken));

const literal = (t: string): string => {
  if (t[0] === '"') return t.slice(1, -1);
  if (t[0] === "\\") return t[1];
  if (t[0] === "_") return " ";
  if (t[0] === "*") return "";
  return t;
};

function compare(v: number, c: { op: string; value: number }): boolean {
  switch (c.op) {
    case "<": return v < c.value;
    case "<=": return v <= c.value;
    case ">": return v > c.value;
    case ">=": return v >= c.value;
    case "=": return v === c.value;
    default: return v !== c.value;
  }
}

/** Round half away from zero at `d` decimals without binary artefacts (1.005 -> 1.01). */
function round(x: number, d: number): number {
  const shift = (n: number, by: number) => {
    const [m, e = "0"] = String(n).split("e");
    return Number(`${m}e${Number(e) + by}`);
  };
  return Math.sign(x) * shift(Math.round(shift(Math.abs(x), d)), -d);
}

export function formatGeneral(v: number): string {
  if (v === 0) return "0";
  const abs = Math.abs(v);
  const sci = () => {
    const [m, e] = abs.toExponential(5).split("e");
    const exp = Number(e);
    return `${v < 0 ? "-" : ""}${m.replace(/\.?0+$/, "")}E${exp < 0 ? "-" : "+"}${String(Math.abs(exp)).padStart(2, "0")}`;
  };
  if (abs >= 1e11) return sci();
  const intDigits = abs < 1 ? 1 : Math.floor(Math.log10(abs)) + 1;
  const r = round(v, Math.max(0, 10 - intDigits));
  if (r === 0) return sci();
  return String(r).includes("e") ? sci() : String(r);
}

function formatNumber(value: number, tokens: string[]): string {
  const ph = (t: string) => t === "0" || t === "#" || t === "?";
  const first = tokens.findIndex(ph);
  if (first < 0) return tokens.map(literal).join("");
  let last = tokens.length - 1;
  while (!ph(tokens[last])) last--;
  // Thousands scaling: commas right after the last digit placeholder.
  let scale = 1, j = last + 1;
  for (; tokens[j] === ","; j++) scale *= 1000;
  const body = tokens.slice(first, last + 1);
  const pct = tokens.filter((t) => t === "%").length;
  let v = (value * 100 ** pct) / scale;

  const e = body.findIndex((t) => /^E[+-]$/i.test(t));
  const mant = e < 0 ? body : body.slice(0, e);
  const dot = mant.indexOf(".");
  const intPh = (dot < 0 ? mant : mant.slice(0, dot)).filter(ph);
  const decPh = dot < 0 ? [] : mant.slice(dot + 1).filter(ph);
  const grouping = (dot < 0 ? mant : mant.slice(0, dot)).includes(",");

  let exp = "";
  if (e >= 0) {
    const expDigits = body.slice(e + 1).filter(ph).length;
    let x = v === 0 ? 0 : Math.floor(Math.log10(Math.abs(v)));
    if (round(v / 10 ** x, decPh.length) >= 10) x++;
    v = v === 0 ? 0 : v / 10 ** x;
    exp = (x < 0 ? "E-" : body[e][1] === "+" ? "E+" : "E") + String(Math.abs(x)).padStart(expDigits, "0");
  }
  const r = round(Math.abs(v), decPh.length);
  let [intStr, decStr = ""] = r.toFixed(decPh.length).split(".");
  const minInt = intPh.filter((t) => t === "0").length;
  if (intStr === "0" && minInt === 0) intStr = "";
  intStr = intStr.padStart(minInt, "0");
  if (grouping) intStr = intStr.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  const minDec = decPh.filter((t) => t === "0").length;
  while (decStr.length > minDec && decStr.endsWith("0")) decStr = decStr.slice(0, -1);
  const num = intStr + (dot >= 0 ? "." + decStr : "") + exp;
  const lit = (ts: string[]) => ts.filter((t) => t !== ",").map(literal).join("");
  return lit(tokens.slice(0, first)) + num + tokens.slice(j).map(literal).join("");
}

function formatDate(serial: number, tokens: string[], date1904: boolean): string {
  // Without fractional seconds, times round to the nearest second (23:59:59.9 shows as the next day's 00:00:00).
  if (!tokens.some((t) => /^s+\.0/i.test(t))) serial = Math.round(serial * 86400) / 86400;
  const p = serialToDate(serial, date1904);
  const ampm = tokens.some((t) => /^(AM\/PM|A\/P)$/i.test(t));
  const pad = (n: number, w: number) => String(n).padStart(w, "0");
  const out: string[] = [];
  tokens.forEach((t, i) => {
    const lower = t.toLowerCase();
    if (/^y+$/.test(lower)) out.push(lower.length <= 2 ? pad(p.y % 100, 2) : String(p.y));
    else if (/^m+$/.test(lower)) {
      // m is minutes right after an hour or right before seconds.
      const prev = tokens.slice(0, i).reverse().find(isDateToken);
      const next = tokens.slice(i + 1).find(isDateToken);
      const minutes = lower.length <= 2 && ((prev && /^\[?h/i.test(prev)) || (next && /^\[?s/i.test(next)));
      if (minutes) out.push(pad(p.mi, lower.length));
      else if (lower.length <= 2) out.push(pad(p.m, lower.length));
      else out.push(lower.length === 3 ? MONTHS[p.m - 1].slice(0, 3) : lower.length === 5 ? MONTHS[p.m - 1][0] : MONTHS[p.m - 1]);
    } else if (/^d+$/.test(lower)) out.push(lower.length <= 2 ? pad(p.d, lower.length) : lower.length === 3 ? DAYS[p.dow].slice(0, 3) : DAYS[p.dow]);
    else if (/^h+$/.test(lower)) out.push(pad(ampm ? p.h % 12 || 12 : p.h, Math.min(2, lower.length)));
    else if (/^s+/.test(lower)) {
      const frac = lower.split(".")[1];
      out.push(pad(p.s, Math.min(2, lower.split(".")[0].length)) + (frac ? "." + pad(Math.floor(p.ms / 10 ** (3 - frac.length)), frac.length) : ""));
    } else if (lower === "am/pm") out.push(p.h < 12 ? "AM" : "PM");
    else if (lower === "a/p") out.push(p.h < 12 ? "A" : "P");
    else if (/^\[h+\]$/.test(lower)) out.push(pad(Math.floor(serial * 24 + 1e-9), lower.length - 2));
    else if (/^\[m+\]$/.test(lower)) out.push(pad(Math.floor(serial * 1440 + 1e-9), lower.length - 2));
    else if (/^\[s+\]$/.test(lower)) out.push(pad(Math.round(serial * 86400), lower.length - 2));
    else out.push(literal(t));
  });
  return out.join("");
}

export interface Formatted {
  text: string;
  color?: string;
}

/** A value as the cell shows it under format `code`. */
export function formatValue(v: Scalar, code: string, date1904 = false): Formatted {
  if (v === null) return { text: "" };
  if (typeof v === "boolean") return { text: v ? "TRUE" : "FALSE" };
  if (typeof v === "object") return { text: v.error };
  const sections = parseCode(code || "General");
  if (typeof v === "string") {
    const sec = sections[3] ?? (sections.length === 1 && sections[0].tokens.includes("@") ? sections[0] : null);
    if (!sec) return { text: v };
    return { text: sec.tokens.map((t) => (t === "@" ? v : literal(t))).join(""), color: sec.color };
  }
  if (!Number.isFinite(v)) return { text: "#NUM!" };
  let sec: Section, value = v;
  if (sections.some((s) => s.cond)) {
    sec = sections.find((s, i) => i < 2 && s.cond && compare(v, s.cond)) ?? sections.find((s, i) => i < 3 && !s.cond) ?? sections[0];
    value = Math.abs(v);
    if (v < 0 && sec === sections[0] && !sec.cond) value = v;
  } else if (v < 0 && sections.length > 1) {
    sec = sections[1];
    value = -v;
  } else sec = v === 0 && sections.length > 2 ? sections[2] : sections[0];
  const color = sec.color;
  const general = sec.tokens.join("").toLowerCase() === "general" || !sec.tokens.length;
  if (general) return { text: formatGeneral(value), color };
  if (sec.tokens.some(isDateToken)) return { text: value < 0 ? "#".repeat(8) : formatDate(value, sec.tokens, date1904), color };
  const text = formatNumber(value, sec.tokens.map((t) => (t.toLowerCase() === "general" ? "0" : t)));
  // A single section shows negatives with a minus sign (unless the number rounded to zero).
  const sign = value < 0 && /[1-9]/.test(text) ? "-" : "";
  return { text: sign + text, color };
}

// ---------------------------------------------------------------------------
// Typed input
// ---------------------------------------------------------------------------

export interface ParsedInput {
  value: Scalar;
  /** Number format the input implies ("12%" -> 0%), when the cell has none. */
  format?: string;
  /** Leading apostrophe: keep as text. */
  quoted?: boolean;
}

const ERROR_VALUES = ["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"];
const NUM = /^[-+]?(\d{1,3}(,\d{3})+|\d+)?(\.\d+)?([eE][-+]?\d+)?$/;

function toNumber(s: string): number | null {
  if (!NUM.test(s) || !/\d/.test(s)) return null;
  if (/[eE]/.test(s) && s.includes(",")) return null;
  return Number(s.replace(/,/g, ""));
}

function parseTime(s: string): { value: number; format: string } | null {
  const m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?$/i.exec(s);
  if (!m) return null;
  let h = Number(m[1]);
  const mi = Number(m[2]), sec = Number(m[3] ?? 0);
  if (m[4]) {
    if (h < 1 || h > 12) return null;
    h = (h % 12) + (m[4].toUpperCase() === "PM" ? 12 : 0);
  }
  if (h > 23 || mi > 59 || sec > 59) return null;
  return { value: (h * 3600 + mi * 60 + sec) / 86400, format: (m[3] ? "h:mm:ss" : "h:mm") + (m[4] ? " AM/PM" : "") };
}

function parseDate(s: string, date1904: boolean): { value: number; format: string } | null {
  const [datePart, ...rest] = s.split(/\s+/);
  let y: number, mo: number, d: number, format: string;
  let m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(datePart);
  if (m) [y, mo, d, format] = [Number(m[1]), Number(m[2]), Number(m[3]), "yyyy-mm-dd"];
  else if ((m = /^(\d{1,2})\/(\d{1,2})\/(\d{4}|\d{2})$/.exec(datePart))) {
    [mo, d, y, format] = [Number(m[1]), Number(m[2]), Number(m[3]), "m/d/yyyy"];
    if (y < 100) y += y < 30 ? 2000 : 1900;
  } else return null;
  if (mo < 1 || mo > 12 || d < 1 || d > new Date(Date.UTC(y, mo, 0)).getUTCDate()) return null;
  const time = rest.length ? parseTime(rest.join(" ")) : null;
  if (rest.length && !time) return null;
  return { value: dateToSerial(y, mo, d, date1904) + (time?.value ?? 0), format: time ? `${format} ${time.format}` : format };
}

/** What a typed entry means: number, boolean, error, date, or text. Formulas are handled by the caller. */
export function parseInput(raw: string, date1904 = false): ParsedInput {
  if (raw.startsWith("'")) return { value: raw.slice(1), quoted: true };
  const s = raw.trim();
  if (!s) return { value: null };
  if (/^(true|false)$/i.test(s)) return { value: s.toLowerCase() === "true" };
  const err = ERROR_VALUES.find((e) => e === s.toUpperCase());
  if (err) return { value: { error: err } };
  const n = toNumber(s);
  if (n !== null) return { value: n };
  const pct = /^(.*)%$/.exec(s);
  if (pct) {
    const p = toNumber(pct[1].trim());
    if (p !== null) return { value: round(p / 100, 12), format: pct[1].includes(".") ? "0.00%" : "0%" };
  }
  const cur = /^([-+]?)\$\s*(.+)$/.exec(s);
  if (cur) {
    const c = toNumber(cur[2]);
    if (c !== null && c >= 0) return { value: cur[1] === "-" ? -c : c, format: cur[2].includes(".") ? '"$"#,##0.00' : '"$"#,##0' };
  }
  return parseDate(s, date1904) ?? parseTime(s) ?? { value: raw };
}

/** A format with one more (delta 1) or one fewer (-1) decimal place; General starts from what `value` shows. */
export function adjustDecimals(code: string, delta: number, value = 0): string {
  if (!code || code === "General") {
    const shown = formatGeneral(value).split(/E/)[0];
    const places = Math.max(0, (shown.split(".")[1]?.length ?? 0) + delta);
    return places ? "0." + "0".repeat(places) : "0";
  }
  if (isDateFormat(code)) return code;
  // Literal and bracketed parts stay as they are; only the first number in each section changes.
  const parts: string[] = code.match(/"[^"]*"|\[[^\]]*\]|\\.|;|[^"[\\;]+/g) ?? [];
  let done = false;
  return parts
    .map((p: string) => {
      if (p === ";") {
        done = false;
        return p;
      }
      if (done || /^["[\\]/.test(p)) return p;
      return p.replace(/([0#?,]*[0#?])(\.([0#?]*))?/, (m, int: string, _dot: string, dec: string | undefined) => {
        done = true;
        const places = Math.max(0, (dec?.length ?? 0) + delta);
        return int + (places ? "." + (dec ?? "").padEnd(places, "0").slice(0, places) : "");
      });
    })
    .join("");
}
