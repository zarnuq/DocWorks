/*
 * Pivot tables: the definition (rows, columns, values) is read from the pivot table part and its cache; refreshing
 * works the table out again from the source range's current values, writes it into the sheet (Excel's compact
 * layout: row fields nested in one column with subtotals at the top, one column field, any number of values) and
 * rewrites the definition and the cache to match. The cache is marked to refresh on load, so Excel redoes its own
 * layout when it opens the file.
 */
import { Edits } from "./commands";
import { compareValues } from "./formula-evaluator";
import { Scalar } from "./number-format";
import { parseRange, Rect, rangeName } from "./references";
import { Sheet } from "./workbook-model";

export type PivotFn = "sum" | "count" | "average" | "max" | "min";
export interface PivotValue {
  field: number;
  fn: PivotFn;
  name?: string;
}
export interface Pivot {
  name: string;
  /** The pivot table and cache parts. */
  path: string;
  cachePath: string;
  /** Cache field names (the source's header row). */
  fields: string[];
  rows: number[];
  cols: number[];
  values: PivotValue[];
  source: { sheet: string; ref: Rect } | null;
  location: Rect;
  /** Report filters, which are kept but not applied here. */
  filtered: boolean;
}
export interface PivotLayout {
  rows: number[];
  cols: number[];
  values: PivotValue[];
}

const kids = (el: Element | null | undefined, name: string) => (el ? Array.from(el.children).filter((c) => c.localName === name) : []);
const kid = (el: Element | null | undefined, name: string) => kids(el, name)[0] ?? null;
const FN_LABEL: Record<PivotFn, string> = { sum: "Sum", count: "Count", average: "Average", max: "Max", min: "Min" };
export const valueName = (fields: string[], v: PivotValue) => v.name ?? `${FN_LABEL[v.fn]} of ${fields[v.field]}`;

/** The pivot tables on `sheet`. */
export function readPivots(sheet: Sheet): Pivot[] {
  const pkg = sheet.wb.pkg;
  const out: Pivot[] = [];
  for (const rel of pkg.rels(sheet.ref.path).filter((r) => r.type === "pivotTable")) {
    const def = pkg.xml(rel.target)?.documentElement;
    const cachePath = pkg.rels(rel.target).find((r) => r.type === "pivotCacheDefinition")?.target;
    const cache = cachePath ? pkg.xml(cachePath)?.documentElement : null;
    const location = parseRange(kid(def, "location")?.getAttribute("ref") ?? "");
    if (!def || !cachePath || !cache || !location) continue;
    const ws = kid(kid(cache, "cacheSource"), "worksheetSource");
    const ref = parseRange(ws?.getAttribute("ref") ?? "");
    const axis = (name: string) => kids(kid(def, name), "field").map((f) => Number(f.getAttribute("x"))).filter((x) => x >= 0);
    out.push({
      name: def.getAttribute("name") ?? "",
      path: rel.target,
      cachePath,
      fields: kids(kid(cache, "cacheFields"), "cacheField").map((f) => f.getAttribute("name") ?? ""),
      rows: axis("rowFields"),
      cols: axis("colFields"),
      values: kids(kid(def, "dataFields"), "dataField").map((d) => ({ field: Number(d.getAttribute("fld")), fn: (d.getAttribute("subtotal") ?? "sum") as PivotFn, name: d.getAttribute("name") ?? undefined })),
      source: ref && ws ? { sheet: ws.getAttribute("sheet") ?? sheet.name, ref } : null,
      location,
      filtered: kids(kid(def, "pageFields"), "pageField").length > 0,
    });
  }
  return out;
}

export function pivotAt(sheet: Sheet, row: number, col: number): Pivot | null {
  return readPivots(sheet).find((p) => row >= p.location.r1 && row <= p.location.r2 && col >= p.location.c1 && col <= p.location.c2) ?? null;
}

const blank = (v: Scalar) => v === null || v === "";
const label = (v: Scalar): string | number | boolean => (v === null || v === "" ? "(blank)" : typeof v === "object" ? v.error : v);
const keyOf = (v: Scalar) => `${typeof v}:${String(label(v)).toLowerCase()}`;
const sorted = (vals: Scalar[]) => {
  const seen = new Map<string, Scalar>();
  for (const v of vals) if (!seen.has(keyOf(v))) seen.set(keyOf(v), v);
  // Blanks last, as Excel lists "(blank)".
  return [...seen.values()].sort((a, b) => (blank(a) ? 1 : 0) - (blank(b) ? 1 : 0) || compareValues(label(a), label(b)));
};

function aggregate(fn: PivotFn, vals: Scalar[]): number | null {
  if (fn === "count") {
    const n = vals.filter((v) => !blank(v)).length;
    return n || null;
  }
  const nums = vals.filter((v): v is number => typeof v === "number");
  if (!nums.length) return null;
  if (fn === "sum") return nums.reduce((a, b) => a + b, 0);
  if (fn === "average") return nums.reduce((a, b) => a + b, 0) / nums.length;
  return fn === "max" ? Math.max(...nums) : Math.min(...nums);
}

/** The source as a header and records, read from the sheet's current values. */
function sourceOf(sheet: Sheet, p: Pivot): { header: string[]; records: Scalar[][] } {
  const src = p.source && sheet.wb.sheets.find((s) => s.name === p.source!.sheet);
  if (!p.source || !src) throw new Error("This pivot table's source range couldn't be found.");
  const { r1, c1, r2, c2 } = p.source.ref;
  const at = (r: number, c: number) => src.at(r, c)?.value ?? null;
  const header = Array.from({ length: c2 - c1 + 1 }, (_, i) => (blank(at(r1, c1 + i)) ? `Column${i + 1}` : String(label(at(r1, c1 + i)))));
  const records: Scalar[][] = [];
  for (let r = r1 + 1; r <= r2; r++) {
    const rec = header.map((_, i) => at(r, c1 + i));
    if (rec.some((v) => !blank(v))) records.push(rec);
  }
  return { header, records };
}

/** The pivot table's cells (row-major, null for empty). */
export function pivotGrid(header: string[], records: Scalar[][], layout: PivotLayout): Scalar[][] {
  const { rows, cols, values } = layout;
  const names = values.map((v) => valueName(header, v));
  const colKeys = cols.length ? sorted(records.map((r) => r[cols[0]])) : [];
  const cells = (recs: Scalar[][]): Scalar[] => {
    if (!cols.length) return values.map((v) => aggregate(v.fn, recs.map((r) => r[v.field])));
    const out: Scalar[] = [];
    for (const key of colKeys) {
      const inCol = recs.filter((r) => keyOf(r[cols[0]]) === keyOf(key));
      for (const v of values) out.push(aggregate(v.fn, inCol.map((r) => r[v.field])));
    }
    for (const v of values) out.push(aggregate(v.fn, recs.map((r) => r[v.field])));
    return out;
  };
  const grid: Scalar[][] = [];
  const width = 1 + (cols.length ? (colKeys.length + 1) * values.length : values.length);
  const pad = (row: Scalar[]) => [...row, ...Array<Scalar>(width - row.length).fill(null)];
  if (!cols.length) grid.push(["Row Labels", ...names]);
  else if (values.length === 1) {
    grid.push(pad([names[0], "Column Labels"]));
    grid.push(["Row Labels", ...colKeys.map(label), "Grand Total"]);
  } else {
    grid.push(pad([null, "Column Labels"]));
    grid.push([null, ...colKeys.flatMap((k) => [label(k), ...Array<Scalar>(values.length - 1).fill(null)]), ...names.map((n) => `Total ${n}`)]);
    grid.push(pad(["Row Labels", ...colKeys.flatMap(() => names)]));
  }
  // Row fields nest in one column; a parent row carries its subtotal.
  const walk = (recs: Scalar[][], level: number) => {
    if (level >= rows.length) return;
    for (const key of sorted(recs.map((r) => r[rows[level]]))) {
      const inRow = recs.filter((r) => keyOf(r[rows[level]]) === keyOf(key));
      grid.push([label(key), ...cells(inRow)]);
      walk(inRow, level + 1);
    }
  };
  walk(records, 0);
  grid.push(["Grand Total", ...cells(records)]);
  return grid;
}

const S_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const TABLE_ORDER = ["location", "pivotFields", "rowFields", "rowItems", "colFields", "colItems", "pageFields", "dataFields", "formats", "conditionalFormats", "chartFormats", "pivotHierarchies", "pivotTableStyleInfo", "filters", "rowHierarchiesUsage", "colHierarchiesUsage", "extLst"];
const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;");

function element(doc: Document, xml: string): Element {
  return doc.importNode(new DOMParser().parseFromString(xml.replace(/^<(\w+)/, `<$1 xmlns="${S_NS}"`), "application/xml").documentElement, true);
}
function place(parent: Element, el: Element): void {
  const old = kid(parent, el.localName);
  if (old) return void old.replaceWith(el);
  const rank = TABLE_ORDER.indexOf(el.localName);
  parent.insertBefore(el, Array.from(parent.children).find((c) => TABLE_ORDER.indexOf(c.localName) > rank) ?? null);
}

const itemXml = (v: Scalar) => (blank(v) ? "<m/>" : typeof v === "number" ? `<n v="${v}"/>` : typeof v === "boolean" ? `<b v="${v ? 1 : 0}"/>` : `<s v="${esc(String(label(v)))}"/>`);

/**
 * Work the pivot table out again (with a new layout, when given) and write it: cells, definition and cache, as one
 * undo step.
 */
export function refreshPivot(e: Edits, sheet: Sheet, p: Pivot, next?: PivotLayout): void {
  if (p.filtered) throw new Error("Pivot tables with report filters are refreshed in Excel or Google Sheets.");
  const layout: PivotLayout = next ?? { rows: p.rows, cols: p.cols, values: p.values };
  if (layout.cols.length > 1) throw new Error("A pivot table here can have one column field.");
  if (!layout.values.length) throw new Error("Add at least one value.");
  const { header, records } = sourceOf(sheet, p);
  // Sums, averages, maxima and minima show in their source column's number format (as Google's pivot tables do).
  const src = sheet.wb.sheets.find((x) => x.name === p.source?.sheet);
  const formats = layout.values.map((v) => {
    if (v.fn === "count" || !src || !p.source) return null;
    const { r1: top, r2: bottom, c1: left } = p.source.ref;
    for (let r = top + 1; r <= bottom; r++) {
      const cell = src.at(r, left + v.field);
      if (typeof cell?.value === "number") return sheet.wb.styles.numFmt(cell.style);
    }
    return null;
  });
  const headRows = !layout.cols.length ? 1 : layout.values.length > 1 ? 3 : 2;
  const grid = pivotGrid(header, records, layout);
  const { r1, c1 } = p.location;
  const location: Rect = { r1, c1, r2: r1 + grid.length - 1, c2: c1 + grid[0].length - 1 };
  const old = p.location, path = sheet.ref.path, pkg = sheet.wb.pkg;
  e.structural(
    () => {
      // Cache: fields from the header; shared items for fields on an axis or holding text; records to match.
      const cacheDoc = pkg.xml(p.cachePath)!;
      const cache = cacheDoc.documentElement;
      const axes = new Set([...layout.rows, ...layout.cols]);
      const shared = header.map((_, i) => (axes.has(i) || records.some((r) => !blank(r[i]) && typeof r[i] !== "number") ? sorted(records.map((r) => r[i])) : null));
      const fieldsXml = header.map((name, i) => {
        const items = shared[i];
        if (items) return `<cacheField name="${esc(name)}" numFmtId="0"><sharedItems count="${items.length}">${items.map(itemXml).join("")}</sharedItems></cacheField>`;
        const nums = records.map((r) => r[i]).filter((v): v is number => typeof v === "number");
        const range = nums.length ? ` minValue="${Math.min(...nums)}" maxValue="${Math.max(...nums)}"` : "";
        return `<cacheField name="${esc(name)}" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1"${nums.every(Number.isInteger) ? ' containsInteger="1"' : ""}${records.some((r) => blank(r[i])) ? ' containsBlank="1"' : ""}${range}/></cacheField>`;
      });
      kid(cache, "cacheFields").replaceWith(element(cacheDoc, `<cacheFields count="${header.length}">${fieldsXml.join("")}</cacheFields>`));
      cache.setAttribute("refreshOnLoad", "1");
      cache.setAttribute("recordCount", String(records.length));
      pkg.touch(p.cachePath);
      const recPath = pkg.rels(p.cachePath).find((r) => r.type === "pivotCacheRecords")?.target;
      const recDoc = recPath ? pkg.xml(recPath) : null;
      if (recPath && recDoc) {
        const index = shared.map((items) => items && new Map(items.map((v, k) => [keyOf(v), k])));
        const rows = records.map((r) => `<r>${r.map((v, i) => (index[i] ? `<x v="${index[i].get(keyOf(v))}"/>` : itemXml(v))).join("")}</r>`).join("");
        recDoc.replaceChild(element(recDoc, `<pivotCacheRecords count="${records.length}">${rows}</pivotCacheRecords>`), recDoc.documentElement);
        pkg.touch(recPath);
      }
      // Definition: fields on their axes with their items in display order, the value fields, the new place.
      const defDoc = pkg.xml(p.path)!;
      const def = defDoc.documentElement;
      const pf = header.map((_, i) => {
        const on = layout.rows.includes(i) ? "axisRow" : layout.cols.includes(i) ? "axisCol" : null;
        const data = layout.values.some((v) => v.field === i) ? ' dataField="1"' : "";
        if (!on) return `<pivotField${data} showAll="0"/>`;
        const items = shared[i]!.map((_, k) => `<item x="${k}"/>`).join("");
        return `<pivotField axis="${on}"${data} showAll="0"><items count="${shared[i]!.length + 1}">${items}<item t="default"/></items></pivotField>`;
      });
      place(def, element(defDoc, `<pivotFields count="${header.length}">${pf.join("")}</pivotFields>`));
      const fieldList = (name: string, xs: number[]) => (xs.length ? place(def, element(defDoc, `<${name} count="${xs.length}">${xs.map((x) => `<field x="${x}"/>`).join("")}</${name}>`)) : kid(def, name)?.remove());
      fieldList("rowFields", layout.rows);
      // Several values make a "Values" field across the columns.
      fieldList("colFields", layout.values.length > 1 ? [...layout.cols, -2] : layout.cols);
      // Excel lays out the items again when it refreshes on load.
      kid(def, "rowItems")?.remove();
      kid(def, "colItems")?.remove();
      const dataXml = layout.values.map((v) => `<dataField name="${esc(valueName(header, { ...v, name: undefined }))}" fld="${v.field}"${v.fn === "sum" ? "" : ` subtotal="${v.fn}"`} baseField="0" baseItem="0"/>`);
      place(def, element(defDoc, `<dataFields count="${layout.values.length}">${dataXml.join("")}</dataFields>`));
      const loc = kid(def, "location");
      loc.setAttribute("ref", rangeName(location));
      loc.setAttribute("firstHeaderRow", "1");
      loc.setAttribute("firstDataRow", String(!layout.cols.length ? 1 : layout.values.length > 1 ? 3 : 2));
      loc.setAttribute("firstDataCol", "1");
      pkg.touch(p.path);
    },
    () => {
      // Cells: clear the old area, then write the table.
      const s = e.wb.sheets.find((x) => x.ref.path === path)!;
      for (const c of s.cellsIn({ r1: old.r1, c1: old.c1, r2: Math.max(old.r2, location.r2), c2: Math.max(old.c2, location.c2) })) if (c.value !== null || c.formula !== undefined) s.setCell(c.row, c.col, { value: null, formula: null });
      const styles = new Map<string, number>();
      const styleOf = (fmt: string | null) => {
        if (!fmt || fmt === "General") return 0;
        if (!styles.has(fmt)) styles.set(fmt, s.wb.styles.derive(0, { numFmt: fmt }));
        return styles.get(fmt)!;
      };
      if (formats.some((f) => f && f !== "General")) s.wb.touchStyles();
      grid.forEach((row, i) =>
        row.forEach((v, j) => {
          if (v === null) return;
          const fmt = typeof v === "number" && i >= headRows && j >= 1 ? formats[(j - 1) % layout.values.length] : null;
          s.setCell(r1 + i, c1 + j, { value: v, style: styleOf(fmt) });
        })
      );
    }
  );
}

/**
 * Google's Insert › Pivot table: a new sheet ("Pivot Table 1"...) with a pivot table over `r` on `source`,
 * starting from the first text column as rows and the sum of the first number column (or a count). One undo step.
 */
export function insertPivot(e: Edits, source: Sheet, r: Rect): string {
  if (r.r2 <= r.r1) throw new Error("Select a range with a header row and at least one row of data.");
  const header = Array.from({ length: r.c2 - r.c1 + 1 }, (_, i) => source.at(r.r1, r.c1 + i)?.value ?? null);
  const column = (i: number) => Array.from({ length: r.r2 - r.r1 }, (_, k) => source.at(r.r1 + 1 + k, r.c1 + i)?.value ?? null);
  const text = header.findIndex((_, i) => column(i).some((v) => typeof v === "string"));
  const number = header.findIndex((_, i) => i !== text && column(i).some((v) => typeof v === "number"));
  const layout: PivotLayout = {
    rows: text >= 0 ? [text] : [0],
    cols: [],
    values: [number >= 0 ? { field: number, fn: "sum" } : { field: text >= 0 ? text : 0, fn: "count" }],
  };
  let n = 1;
  while (e.wb.sheet(`Pivot Table ${n}`)) n++;
  const name = `Pivot Table ${n}`;
  const at = e.wb.sheets.indexOf(source), sourceName = source.name;
  e.run(() => {
    e.structural(() => {
      const ref = e.wb.pkg.addSheet(name, `<worksheet xmlns="${S_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetFormatPr defaultRowHeight="15"/><sheetData/></worksheet>`, at);
      e.wb.pkg.addPivotTable(ref.path, sourceName, rangeName(r));
    });
    const sheet = e.wb.sheet(name)!;
    refreshPivot(e, sheet, readPivots(sheet)[0], layout);
    return [];
  });
  return name;
}
