/*
 * Sparklines (Excel 2010's x14:sparklineGroups in the sheet's extension list): small line, column or win/loss
 * charts drawn inside a cell from a row or column of values. Read only; the XML stays as it is.
 */
import { parseRange, Rect } from "./references";
import { Sheet } from "./workbook-model";

const X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";

export interface Sparkline {
  row: number;
  col: number;
  type: "line" | "column" | "stacked";
  /** In the source range's order; null for blanks and text. */
  values: (number | null)[];
  color: string;
  /** Point colours, set only when the group switches them on. */
  negative?: string;
  markers?: string;
  high?: string;
  low?: string;
  first?: string;
  last?: string;
  /** The axis colour when the group shows the horizontal axis. */
  axis?: string;
  /** Fixed axis range (custom, or the same for the group); otherwise each sparkline's own. */
  min?: number;
  max?: number;
  /** Line width in px. */
  weight: number;
  rtl: boolean;
}

const kid = (el: Element | null | undefined, name: string) => (el ? Array.from(el.children).find((c) => c.localName === name) ?? null : null);
const on = (el: Element, attr: string) => ["1", "true"].includes(el.getAttribute(attr) ?? "");

/** The numbers in a sheet-qualified range ("Data!A1:A5"); null where a cell is blank or not a number. */
function valuesOf(sheet: Sheet, ref: string): (number | null)[] {
  const bang = ref.lastIndexOf("!");
  const name = bang < 0 ? sheet.name : ref.slice(0, bang).replace(/^'(.*)'$/, "$1").replace(/''/g, "'");
  const src = sheet.wb.sheets.find((s) => s.name === name);
  const r: Rect | null = parseRange(ref.slice(bang + 1).replace(/\$/g, ""));
  if (!src || !r) return [];
  const out: (number | null)[] = [];
  for (let row = r.r1; row <= r.r2; row++)
    for (let col = r.c1; col <= r.c2; col++) {
      const v = src.at(row, col)?.value;
      out.push(typeof v === "number" ? v : null);
    }
  return out;
}

export function readSparklines(sheet: Sheet): Sparkline[] {
  const out: Sparkline[] = [];
  for (const group of Array.from(sheet.doc.getElementsByTagNameNS(X14, "sparklineGroup"))) {
    const color = (name: string) => sheet.wb.styles.color(kid(group, name)) ?? undefined;
    const flag = (attr: string, name: string) => (on(group, attr) ? color(name) : undefined);
    const type = group.getAttribute("type") === "column" ? "column" : group.getAttribute("type") === "stacked" ? "stacked" : "line";
    const lines = Array.from(kid(group, "sparklines")?.children ?? []).map((sp) => ({
      at: parseRange(kid(sp, "sqref")?.textContent ?? ""),
      values: valuesOf(sheet, kid(sp, "f")?.textContent ?? ""),
    }));
    const all = lines.flatMap((l) => l.values).filter((v): v is number => v !== null);
    const bound = (which: "min" | "max") => {
      const kind = group.getAttribute(`${which}AxisType`);
      if (kind === "custom") return Number(group.getAttribute(which === "min" ? "manualMin" : "manualMax") ?? 0);
      if (kind === "group" && all.length) return which === "min" ? Math.min(...all) : Math.max(...all);
      return undefined;
    };
    for (const l of lines) {
      if (!l.at) continue;
      out.push({
        row: l.at.r1,
        col: l.at.c1,
        type,
        values: l.values,
        color: color("colorSeries") ?? "#376092",
        negative: flag("negative", "colorNegative"),
        markers: type === "line" ? flag("markers", "colorMarkers") : undefined,
        high: flag("high", "colorHigh"),
        low: flag("low", "colorLow"),
        first: flag("first", "colorFirst"),
        last: flag("last", "colorLast"),
        axis: flag("displayXAxis", "colorAxis"),
        min: bound("min"),
        max: bound("max"),
        weight: (Number(group.getAttribute("lineWeight") ?? 0.75) * 4) / 3,
        rtl: on(group, "rightToLeft"),
      });
    }
  }
  return out;
}
