import { assertAction } from "./protection";
/*
 * Conditional formatting: the sheet's cfRule elements read, evaluated per cell (lazily, for the cells being
 * drawn) and written as Google's sidebar offers them. Formulas in rules are relative to the top-left cell of the
 * rule's first range, as in Excel; colour scales and data bars use the values across all of the rule's ranges.
 */
import { NS_S } from "../../infrastructure/xlsx/xlsx-workbook";
import { Edits } from "./commands";
import { compareValues } from "./formula-evaluator";
import { translateFormula } from "./formula";
import { dateToSerial, Scalar, serialToDate } from "./number-format";
import { parseRange, Rect, rangeName } from "./references";
import { DxfStyle } from "./styles";
import { Sheet } from "./workbook-model";

export interface CfStyle extends DxfStyle {
  /** A data bar: the part of the cell's width it fills. */
  bar?: { fraction: number; color: string };
  /** An icon set's icon, drawn at the cell's left. */
  icon?: { glyph: string; color: string };
  /** An icon set with showValue="0": the icon only. */
  hideValue?: boolean;
}

interface Point {
  type: string;
  val: string | null;
  /** Icon set thresholds: reached at >= (default) or only above. */
  gte: boolean;
}

export interface CfRule {
  el: Element;
  block: Element;
  ranges: Rect[];
  type: string;
  operator: string | null;
  formulas: string[];
  text: string | null;
  dxf: number | null;
  priority: number;
  stopIfTrue: boolean;
  /** Colour scale, data bar or icon set points, with their colours. */
  points: Point[];
  colors: string[];
}

const kids = (el: Element | null | undefined, tag: string) => (el ? Array.from(el.children).filter((c) => c.localName === tag) : []);
const on = (el: Element, attr: string, dflt: boolean) => (el.hasAttribute(attr) ? ["1", "true"].includes(el.getAttribute(attr)!) : dflt);

const X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";

export function readRules(sheet: Sheet): CfRule[] {
  const out: CfRule[] = [];
  const read = (block: Element, ranges: Rect[]) => {
    for (const el of kids(block, "cfRule")) {
      const scale = kids(el, "colorScale")[0] ?? kids(el, "dataBar")[0] ?? kids(el, "iconSet")[0];
      out.push({
        el,
        block,
        ranges,
        type: el.getAttribute("type") ?? "",
        operator: el.getAttribute("operator"),
        formulas: kids(el, "formula").map((f) => f.textContent ?? ""),
        text: el.getAttribute("text"),
        dxf: el.hasAttribute("dxfId") ? Number(el.getAttribute("dxfId")) : null,
        priority: Number(el.getAttribute("priority") ?? 0),
        stopIfTrue: on(el, "stopIfTrue", false),
        // Excel 2010 rules keep a point's value in an xm:f child.
        points: kids(scale, "cfvo").map((p) => ({ type: p.getAttribute("type") ?? "min", val: p.getAttribute("val") ?? kids(p, "f")[0]?.textContent ?? null, gte: on(p, "gte", true) })),
        colors: kids(scale, "color").map((c) => sheet.wb.styles.color(c) ?? "#000000"),
      });
    }
  };
  const root = sheet.doc.documentElement;
  for (const block of kids(root, "conditionalFormatting")) read(block, (block.getAttribute("sqref") ?? "").split(/\s+/).map(parseRange).filter((r): r is Rect => !!r));
  // Excel 2010 rules in the extension list (custom icon sets); data bar extensions belong to a rule above.
  const paired = new Set(out.flatMap((r) => Array.from(r.el.getElementsByTagNameNS(X14, "id")).map((i) => i.textContent ?? "")));
  for (const block of Array.from(root.getElementsByTagNameNS(X14, "conditionalFormatting"))) {
    const ranges = (kids(block, "sqref")[0]?.textContent ?? "").split(/\s+/).map(parseRange).filter((r): r is Rect => !!r);
    const before = out.length;
    read(block, ranges);
    out.splice(before, out.length - before, ...out.slice(before).filter((r) => !paired.has(r.el.getAttribute("id") ?? "")));
  }
  return out.sort((a, b) => a.priority - b.priority);
}

const R = "#D93025", Y = "#F4B400", G = "#0F9D58", K = "#5F6368", B = "#000000", U = "#4285F4";
const icons = (glyphs: string, colors: string[]) => Array.from(glyphs).map((glyph, i) => ({ glyph, color: colors[i] }));
const bars = (n: number) => Array.from({ length: n }, (_, i) => ({ glyph: "▂▄▆█".slice(0, Math.max(1, i + 5 - n)), color: i + 5 - n > 0 ? U : "#C0C0C0" }));
// ponytail: Excel's icons drawn as Unicode glyphs (lowest first); custom x14 icon sets are not read.
const ICON_SETS: Record<string, { glyph: string; color: string }[]> = {
  "3Arrows": icons("↓→↑", [R, Y, G]),
  "3ArrowsGray": icons("↓→↑", [K, K, K]),
  "3Flags": icons("⚑⚑⚑", [R, Y, G]),
  "3TrafficLights1": icons("●●●", [R, Y, G]),
  "3TrafficLights2": icons("●●●", [R, Y, G]),
  "3Signs": icons("◆▲●", [R, Y, G]),
  "3Symbols": icons("✖!✔", [R, Y, G]),
  "3Symbols2": icons("✖!✔", [R, Y, G]),
  "3Stars": icons("☆★★", [Y, "#F9DC8C", Y]),
  "3Triangles": icons("▼▬▲", [R, Y, G]),
  "4Arrows": icons("↓↘↗↑", [R, Y, Y, G]),
  "4ArrowsGray": icons("↓↘↗↑", [K, K, K, K]),
  "4RedToBlack": icons("●●●●", [B, K, "#F4A7A0", R]),
  "4Rating": bars(4),
  "4TrafficLights": icons("●●●●", [B, R, Y, G]),
  "5Arrows": icons("↓↘→↗↑", [R, Y, Y, Y, G]),
  "5ArrowsGray": icons("↓↘→↗↑", [K, K, K, K, K]),
  "5Rating": bars(5),
  "5Quarters": icons("○◔◑◕●", [K, K, K, K, K]),
  "5Boxes": icons("□◔◑◕■", [U, U, U, U, U]),
};

const inRanges = (ranges: Rect[], row: number, col: number) => ranges.some((r) => row >= r.r1 && row <= r.r2 && col >= r.c1 && col <= r.c2);
const isErr = (v: Scalar) => typeof v === "object" && v !== null;
const truthy = (v: Scalar) => (typeof v === "number" ? v !== 0 : typeof v === "boolean" ? v : typeof v === "string" ? v.toUpperCase() === "TRUE" : false);

function mix(a: string, b: string, t: number): string {
  const ch = (h: string, i: number) => parseInt(h.slice(i, i + 2), 16);
  return "#" + [1, 3, 5].map((i) => Math.round(ch(a, i) + (ch(b, i) - ch(a, i)) * t).toString(16).padStart(2, "0")).join("").toUpperCase();
}

function percentile(sorted: number[], p: number): number {
  if (!sorted.length) return 0;
  const at = (sorted.length - 1) * Math.min(1, Math.max(0, p));
  const lo = Math.floor(at);
  return sorted[lo] + (sorted[Math.min(lo + 1, sorted.length - 1)] - sorted[lo]) * (at - lo);
}

/** Rule results for the cells of one sheet, worked out as cells are drawn and kept until `invalidate`. */
export class ConditionalFormats {
  private rules: CfRule[];
  private memo = new Map<number, CfStyle | null>();
  private stats = new Map<CfRule, { numbers: number[]; counts: Map<string, number> }>();

  constructor(
    private sheet: Sheet,
    /** Evaluate a formula as if it were in (row, col). */
    private evaluate: (row: number, col: number, formula: string) => Scalar
  ) {
    this.rules = readRules(sheet);
  }

  get empty(): boolean {
    return !this.rules.length;
  }

  invalidate(): void {
    this.memo.clear();
    this.stats.clear();
  }

  at(row: number, col: number): CfStyle | null {
    if (!this.rules.length) return null;
    const key = row * 16384 + col;
    if (this.memo.has(key)) return this.memo.get(key)!;
    let out: CfStyle | null = null;
    for (const rule of this.rules) {
      if (!inRanges(rule.ranges, row, col)) continue;
      let hit: CfStyle | null = null;
      try {
        hit = this.apply(rule, row, col);
      } catch {
        hit = null; // a formula the evaluator can't handle: no format
      }
      if (!hit) continue;
      // The higher priority (earlier) rule wins each property.
      out = { ...hit, ...(out ?? {}) };
      if (rule.stopIfTrue) break;
    }
    this.memo.set(key, out);
    return out;
  }

  private value(row: number, col: number): Scalar {
    return this.sheet.at(row, col)?.value ?? null;
  }

  private text(row: number, col: number): string {
    const v = this.value(row, col);
    return typeof v === "string" ? v : this.sheet.display(row, col).text;
  }

  private formula(rule: CfRule, i: number, row: number, col: number): Scalar {
    const base = rule.ranges[0];
    return this.evaluate(row, col, translateFormula(rule.formulas[i] ?? "", row - base.r1, col - base.c1));
  }

  private statsOf(rule: CfRule) {
    let s = this.stats.get(rule);
    if (!s) {
      const numbers: number[] = [];
      const counts = new Map<string, number>();
      for (const r of rule.ranges) {
        for (const cell of this.sheet.cellsIn(r)) {
          if (typeof cell.value === "number") numbers.push(cell.value);
          if (cell.value !== null && cell.value !== "") {
            const k = String(typeof cell.value === "object" ? cell.value.error : cell.value).toLowerCase();
            counts.set(k, (counts.get(k) ?? 0) + 1);
          }
        }
      }
      numbers.sort((a, b) => a - b);
      this.stats.set(rule, (s = { numbers, counts }));
    }
    return s;
  }

  /** A colour scale / data bar point's value. */
  private point(rule: CfRule, p: Point, row: number, col: number): number {
    const { numbers } = this.statsOf(rule);
    const min = numbers[0] ?? 0, max = numbers[numbers.length - 1] ?? 0;
    const n = Number(p.val);
    switch (p.type) {
      case "min":
        return min;
      case "max":
        return max;
      case "percent":
        return min + ((max - min) * n) / 100;
      case "percentile":
        return percentile(numbers, n / 100);
      case "formula": {
        if (Number.isFinite(n) && p.val?.trim() !== "") return n;
        const v = this.evaluate(row, col, p.val ?? "0");
        return typeof v === "number" ? v : 0;
      }
      default:
        return n;
    }
  }

  private apply(rule: CfRule, row: number, col: number): CfStyle | null {
    const style = (): CfStyle => (rule.dxf === null ? {} : { ...this.sheet.wb.styles.dxf(rule.dxf) });
    const v = this.value(row, col);
    const yes = (b: boolean) => (b ? style() : null);
    switch (rule.type) {
      case "cellIs": {
        if (v === null) return null;
        const a = this.formula(rule, 0, row, col);
        const cmp = (x: Scalar) => compareValues(v, x);
        switch (rule.operator) {
          case "between":
          case "notBetween": {
            const b = this.formula(rule, 1, row, col);
            const [lo, hi] = compareValues(a, b) <= 0 ? [a, b] : [b, a];
            const inside = cmp(lo) >= 0 && cmp(hi) <= 0;
            return yes(rule.operator === "between" ? inside : !inside);
          }
          case "equal":
            return yes(cmp(a) === 0);
          case "notEqual":
            return yes(cmp(a) !== 0);
          case "greaterThan":
            return yes(cmp(a) > 0);
          case "greaterThanOrEqual":
            return yes(cmp(a) >= 0);
          case "lessThan":
            return yes(cmp(a) < 0);
          case "lessThanOrEqual":
            return yes(cmp(a) <= 0);
          default:
            return null;
        }
      }
      case "expression":
        return yes(truthy(this.formula(rule, 0, row, col)));
      case "containsText":
      case "notContainsText":
      case "beginsWith":
      case "endsWith": {
        const t = this.text(row, col).toLowerCase(), q = (rule.text ?? "").toLowerCase();
        const hit = rule.type === "containsText" ? t.includes(q) : rule.type === "notContainsText" ? !t.includes(q) : rule.type === "beginsWith" ? t.startsWith(q) : t.endsWith(q);
        return yes(hit);
      }
      case "containsBlanks":
        return yes(this.text(row, col).trim() === "");
      case "notContainsBlanks":
        return yes(this.text(row, col).trim() !== "");
      case "containsErrors":
        return yes(isErr(v));
      case "notContainsErrors":
        return yes(!isErr(v));
      case "duplicateValues":
      case "uniqueValues": {
        if (v === null || v === "") return null;
        const n = this.statsOf(rule).counts.get((typeof v === "object" ? v.error : String(v)).toLowerCase()) ?? 0;
        return yes(rule.type === "duplicateValues" ? n > 1 : n === 1);
      }
      case "top10": {
        if (typeof v !== "number") return null;
        const { numbers } = this.statsOf(rule);
        const rank = Number(rule.el.getAttribute("rank") ?? 10);
        const count = on(rule.el, "percent", false) ? Math.max(1, Math.floor((numbers.length * rank) / 100)) : rank;
        if (on(rule.el, "bottom", false)) return yes(v <= numbers[Math.min(count, numbers.length) - 1]);
        return yes(v >= numbers[Math.max(0, numbers.length - count)]);
      }
      case "aboveAverage": {
        if (typeof v !== "number") return null;
        const { numbers } = this.statsOf(rule);
        const avg = numbers.reduce((a, b) => a + b, 0) / (numbers.length || 1);
        const above = on(rule.el, "aboveAverage", true), equal = on(rule.el, "equalAverage", false);
        return yes(above ? v > avg || (equal && v === avg) : v < avg || (equal && v === avg));
      }
      case "colorScale": {
        if (typeof v !== "number" || rule.points.length < 2) return null;
        const stops = rule.points.map((p) => this.point(rule, p, row, col));
        if (v <= stops[0]) return { fill: rule.colors[0] };
        for (let i = 1; i < stops.length; i++) {
          if (v <= stops[i]) {
            const t = stops[i] === stops[i - 1] ? 1 : (v - stops[i - 1]) / (stops[i] - stops[i - 1]);
            return { fill: mix(rule.colors[i - 1], rule.colors[i], t) };
          }
        }
        return { fill: rule.colors[rule.colors.length - 1] };
      }
      case "dataBar": {
        if (typeof v !== "number" || rule.points.length < 2) return null;
        const lo = this.point(rule, rule.points[0], row, col), hi = this.point(rule, rule.points[1], row, col);
        const fraction = hi === lo ? 1 : Math.min(1, Math.max(0, (v - lo) / (hi - lo)));
        return { bar: { fraction, color: rule.colors[0] ?? "#638EC6" } };
      }
      case "iconSet": {
        if (typeof v !== "number" || !rule.points.length) return null;
        const set = rule.el.getElementsByTagNameNS(rule.el.namespaceURI, "iconSet")[0] ?? null;
        const glyphs = ICON_SETS[set?.getAttribute("iconSet") ?? "3TrafficLights1"] ?? ICON_SETS["3TrafficLights1"];
        // Custom icons (Excel 2010): one per threshold, from any set, or none.
        const custom = kids(set, "cfIcon").map((c) => (c.getAttribute("iconSet") === "NoIcons" ? null : ICON_SETS[c.getAttribute("iconSet") ?? ""]?.[Number(c.getAttribute("iconId") ?? 0)] ?? null));
        let i = 0;
        rule.points.forEach((p, k) => {
          const t = this.point(rule, p, row, col);
          if (k > 0 && (v > t || (p.gte && v === t))) i = k;
        });
        if (set && on(set, "reverse", false)) i = rule.points.length - 1 - i;
        const icon = custom.length ? custom[i] : glyphs[Math.min(i, glyphs.length - 1)];
        if (!icon) return null;
        const out: CfStyle = { icon };
        if (set && !on(set, "showValue", true)) out.hideValue = true;
        return out;
      }
      case "timePeriod": {
        if (typeof v !== "number") return null;
        const now = new Date();
        const date1904 = this.sheet.wb.date1904;
        const today = dateToSerial(now.getFullYear(), now.getMonth() + 1, now.getDate(), date1904);
        const day = Math.floor(v), dow = now.getDay();
        const month = (dm: number) => {
          const d = serialToDate(day, date1904);
          const t = new Date(now.getFullYear(), now.getMonth() + dm, 1);
          return d.y === t.getFullYear() && d.m === t.getMonth() + 1;
        };
        const week = (dw: number) => day >= today - dow + 7 * dw && day <= today - dow + 7 * dw + 6;
        const hits: Record<string, () => boolean> = {
          today: () => day === today,
          yesterday: () => day === today - 1,
          tomorrow: () => day === today + 1,
          last7Days: () => day <= today && day > today - 7,
          thisWeek: () => week(0),
          lastWeek: () => week(-1),
          nextWeek: () => week(1),
          thisMonth: () => month(0),
          lastMonth: () => month(-1),
          nextMonth: () => month(1),
        };
        return yes(hits[rule.el.getAttribute("timePeriod") ?? ""]?.() ?? false);
      }
      default:
        return null;
    }
  }
}

// -----------------------------------------------------------------------------
// Editing (Google's conditional format rules sidebar)
// -----------------------------------------------------------------------------

export type NewRuleType =
  | "cellIs"
  | "expression"
  | "containsText"
  | "notContainsText"
  | "beginsWith"
  | "endsWith"
  | "containsBlanks"
  | "notContainsBlanks"
  | "colorScale";

export interface NewRule {
  type: NewRuleType;
  /** cellIs: between, notBetween, equal, notEqual, greaterThan, greaterThanOrEqual, lessThan, lessThanOrEqual. */
  operator?: string;
  /** Values or text typed in the sidebar; a leading "=" makes a value a formula. */
  values?: string[];
  style?: DxfStyle;
  /** Colour scale: two or three colours from min (through the midpoint) to max. */
  colors?: string[];
}

const escape = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;");
/** A typed value as a rule formula: numbers as they are, "=..." as a formula, anything else as quoted text. */
function operand(raw: string): string {
  const t = raw.trim();
  if (t.startsWith("=")) return t.slice(1);
  if (t !== "" && Number.isFinite(Number(t))) return t;
  return `"${t.replace(/"/g, '""')}"`;
}

/** Add a rule for `ranges`, below the sheet's existing rules. */
export function addRule(e: Edits, sheet: Sheet, ranges: Rect[], rule: NewRule): void {
  assertAction(sheet, "formatCells");
  const priority = Math.max(0, ...readRules(sheet).map((r) => r.priority)) + 1;
  const first = rangeName({ ...ranges[0], r2: ranges[0].r1, c2: ranges[0].c1 });
  const values = rule.values ?? [];
  const q = (values[0] ?? "").replace(/"/g, '""');
  let body: string;
  if (rule.type === "colorScale") {
    const colors = rule.colors ?? ["#FFFFFF", "#57BB8A"];
    const points = colors.length === 3 ? `<cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/>` : `<cfvo type="min"/><cfvo type="max"/>`;
    body = `<cfRule type="colorScale" priority="${priority}"><colorScale>${points}${colors.map((c) => `<color rgb="FF${c.slice(1).toUpperCase()}"/>`).join("")}</colorScale></cfRule>`;
  } else {
    const dxf = sheet.wb.styles.addDxf(rule.style ?? { fill: "#B7E1CD" });
    sheet.wb.touchStyles();
    const head = (attrs: string) => `<cfRule type="${rule.type}" dxfId="${dxf}" priority="${priority}"${attrs}>`;
    const f = (x: string) => `<formula>${escape(x)}</formula>`;
    switch (rule.type) {
      case "cellIs":
        body = head(` operator="${rule.operator ?? "equal"}"`) + values.slice(0, rule.operator === "between" || rule.operator === "notBetween" ? 2 : 1).map((v) => f(operand(v))).join("") + "</cfRule>";
        break;
      case "expression":
        body = head("") + f(operand(values[0] ?? "=TRUE").replace(/^"(.*)"$/, "$1")) + "</cfRule>";
        break;
      case "containsText":
        body = head(` operator="containsText" text="${escape(values[0] ?? "")}"`) + f(`NOT(ISERROR(SEARCH("${q}",${first})))`) + "</cfRule>";
        break;
      case "notContainsText":
        body = head(` operator="notContains" text="${escape(values[0] ?? "")}"`) + f(`ISERROR(SEARCH("${q}",${first}))`) + "</cfRule>";
        break;
      case "beginsWith":
        body = head(` operator="beginsWith" text="${escape(values[0] ?? "")}"`) + f(`LEFT(${first},LEN("${q}"))="${q}"`) + "</cfRule>";
        break;
      case "endsWith":
        body = head(` operator="endsWith" text="${escape(values[0] ?? "")}"`) + f(`RIGHT(${first},LEN("${q}"))="${q}"`) + "</cfRule>";
        break;
      case "containsBlanks":
        body = head("") + f(`LEN(TRIM(${first}))=0`) + "</cfRule>";
        break;
      case "notContainsBlanks":
        body = head("") + f(`LEN(TRIM(${first}))>0`) + "</cfRule>";
        break;
    }
  }
  const xml = `<conditionalFormatting xmlns="${NS_S}" sqref="${ranges.map(rangeName).join(" ")}">${body}</conditionalFormatting>`;
  e.structural(() => {
    const el = sheet.doc.importNode(new DOMParser().parseFromString(xml, "application/xml").documentElement, true);
    const root = sheet.doc.documentElement;
    // After the last existing block, or where CT_Worksheet puts it (before data validations and what follows).
    const last = kids(root, "conditionalFormatting").pop();
    const AFTER = ["dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks", "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing", "legacyDrawing", "legacyDrawingHF", "drawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst"];
    root.insertBefore(el, last ? last.nextSibling : Array.from(root.children).find((c) => AFTER.includes(c.localName)) ?? null);
    sheet.wb.pkg.touch(sheet.ref.path);
  });
}

/** Remove one rule (and its block when it was the block's last). */
export function deleteRule(e: Edits, sheet: Sheet, rule: CfRule): void {
  assertAction(sheet, "formatCells");
  e.structural(() => {
    const live = readRules(sheet).find((r) => r.priority === rule.priority && r.type === rule.type);
    if (!live) return;
    live.el.remove();
    if (!kids(live.block, "cfRule").length) {
      // An emptied Excel 2010 block takes its now empty wrappers with it (conditionalFormattings, ext, extLst).
      let gone: Element | null = live.block;
      while (gone && gone !== sheet.doc.documentElement) {
        const parent: Element | null = gone.parentElement;
        gone.remove();
        gone = parent && parent !== sheet.doc.documentElement && !parent.children.length ? parent : null;
      }
    }
    sheet.wb.pkg.touch(sheet.ref.path);
  });
}
