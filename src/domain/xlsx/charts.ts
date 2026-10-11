/*
 * Charts on sheets: graphic frames in the sheet's drawing part pointing at chart parts. They're drawn with the
 * slide chart renderer from the sheet's current cell values (Google's exports carry no cached values), formatted
 * with chart-format.ts, and inserted from a selection as Google's Insert › Chart does.
 */
import { parseXml } from "../../infrastructure/ooxml/ooxml-package";
import { chartData, ChartData, ChartModel } from "../pptx/chart";
import { axesXml, chartPalette, ChartFormat, readChartFormat, setChartFormat } from "../pptx/chart-format";
import { ColorContext } from "../pptx/color";
import { Edits } from "./commands";
import { anchorBox } from "./drawing";
import { cellName, colName, parseRange, quoteSheet, Rect } from "./references";
import { child as sheetChild } from "./sheet-edit";
import { Sheet, Workbook } from "./workbook-model";

const XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const EMU_PER_PX = 9525;

export interface SheetChart {
  /** The chart part. */
  path: string;
  /** Its anchor in the drawing part, and the drawing's relationship to the chart. */
  anchor: Element;
  drawing: string;
  rId: string;
  box: { x: number; y: number; w: number; h: number };
  name: string;
}

export function sheetCharts(wb: Workbook, sheet: Sheet): SheetChart[] {
  const rel = wb.pkg.rels(sheet.ref.path).find((r) => r.type === "drawing" && !r.external);
  const doc = rel && wb.pkg.xml(rel.target);
  if (!rel || !doc) return [];
  const out: SheetChart[] = [];
  for (const anchor of Array.from(doc.documentElement.children)) {
    const frame = Array.from(anchor.children).find((c) => c.namespaceURI === XDR && c.localName === "graphicFrame");
    const rId = frame?.getElementsByTagNameNS(C, "chart")[0]?.getAttributeNS(R, "id");
    const path = rId ? wb.pkg.relTarget(rel.target, rId) : null;
    const box = frame && anchorBox(sheet, anchor, frame);
    if (!frame || !rId || !path || !box || !wb.pkg.xml(path)) continue;
    out.push({ path, anchor, drawing: rel.target, rId, box, name: frame.getElementsByTagNameNS(XDR, "cNvPr")[0]?.getAttribute("name") ?? "Chart" });
  }
  return out;
}

/** Theme colours for charts (scheme slots plus the usual text / background mapping). */
export function chartColors(wb: Workbook): ColorContext {
  return { scheme: wb.styles.theme, map: { tx1: "dk1", bg1: "lt1", tx2: "dk2", bg2: "lt2" } };
}

/** The cells a chart reference (Sheet!$A$1:$A$5) points at, as a sheet and rectangle. */
function resolve(wb: Workbook, f: string): { sheet: Sheet; rect: Rect } | null {
  const ref = f.trim().replace(/^\(|\)$/g, "").replace(/\$/g, "");
  const bang = ref.lastIndexOf("!");
  if (bang < 0) return null;
  const sheet = wb.sheet(ref.slice(0, bang).replace(/^'(.*)'$/, "$1").replace(/''/g, "'"));
  const rect = parseRange(ref.slice(bang + 1));
  return sheet && rect ? { sheet, rect } : null;
}

function cellsOf(wb: Workbook, f: string | null | undefined): { text: string[]; numbers: number[] } | null {
  const at = f ? resolve(wb, f) : null;
  if (!at) return null;
  const text: string[] = [], numbers: number[] = [];
  for (let r = at.rect.r1; r <= at.rect.r2; r++)
    for (let c = at.rect.c1; c <= at.rect.c2; c++) {
      const v = at.sheet.at(r, c)?.value ?? null;
      text.push(at.sheet.display(r, c).text);
      numbers.push(typeof v === "number" ? v : typeof v === "boolean" ? Number(v) : NaN);
    }
  return { text, numbers };
}

const firstF = (el: Element | undefined) => el?.getElementsByTagNameNS(C, "f")[0]?.textContent ?? null;
const serPart = (ser: Element, name: string) => Array.from(ser.children).find((c) => c.namespaceURI === C && c.localName === name);

/** The chart's data from the cells it refers to, falling back to its cached values. */
export function chartModel(wb: Workbook, path: string): ChartModel {
  const xml = wb.pkg.xml(path)!;
  const cached = chartData(xml);
  const series = Array.from(xml.getElementsByTagNameNS(C, "ser"));
  const data: ChartData = { categories: cached.categories, series: [] };
  const cat = series[0] && cellsOf(wb, firstF(serPart(series[0], "cat") ?? serPart(series[0], "xVal")));
  if (cat) data.categories = cat.text;
  series.forEach((ser, i) => {
    const name = cellsOf(wb, firstF(serPart(ser, "tx")))?.text.join(" ");
    const values = cellsOf(wb, firstF(serPart(ser, "val") ?? serPart(ser, "yVal")))?.numbers;
    data.series.push({ name: name || cached.series[i]?.name || `Series ${i + 1}`, values: values ?? cached.series[i]?.values ?? [] });
  });
  if (!data.categories.length && data.series[0]) data.categories = data.series[0].values.map((_, i) => String(i + 1));
  const colors = chartColors(wb);
  const type = Array.from(xml.getElementsByTagNameNS(C, "plotArea")[0]?.children ?? []).find((e) => e.localName.endsWith("Chart"))?.localName ?? "unknown";
  return { path, type, data, format: readChartFormat(xml, colors), palette: chartPalette(colors) };
}

/** The range a chart's data spans (for the chart editor), e.g. "Data!A1:C4". */
export function chartSource(wb: Workbook, path: string): string {
  const refs = Array.from(wb.pkg.xml(path)?.getElementsByTagNameNS(C, "f") ?? []).map((f) => resolve(wb, f.textContent ?? "")).filter((x): x is NonNullable<typeof x> => !!x);
  if (!refs.length) return "";
  const r1 = Math.min(...refs.map((x) => x.rect.r1)), c1 = Math.min(...refs.map((x) => x.rect.c1));
  const r2 = Math.max(...refs.map((x) => x.rect.r2)), c2 = Math.max(...refs.map((x) => x.rect.c2));
  return `${quoteSheet(refs[0].sheet.name)}!${cellName(r1, c1)}:${cellName(r2, c2)}`;
}

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;");
const absolute = (sheet: Sheet, r1: number, c1: number, r2: number, c2: number) =>
  `${quoteSheet(sheet.name)}!$${colName(c1)}$${r1 + 1}${r1 === r2 && c1 === c2 ? "" : `:$${colName(c2)}$${r2 + 1}`}`;

/**
 * Chart XML for the selection. Series run down columns, or along rows when the data is wider than it is tall;
 * a text first row / first column gives series names and categories.
 */
function chartXml(sheet: Sheet, r: Rect): string {
  const isText = (row: number, col: number) => {
    const v = sheet.at(row, col)?.value;
    return typeof v === "string" && v !== "";
  };
  let header = false;
  for (let c = r.c1; c <= r.c2; c++) if (isText(r.r1, c)) header = true;
  const first = header ? r.r1 + 1 : r.r1;
  let labels = r.c2 > r.c1;
  if (labels) for (let row = first; row <= r.r2; row++) if (!isText(row, r.c1) && sheet.at(row, r.c1)?.value != null) labels = false;
  const left = labels ? r.c1 + 1 : r.c1;
  const byRow = r.c2 - left + 1 > r.r2 - first + 1;
  type Ref = { r1: number; c1: number; r2: number; c2: number };
  // Each series: its name cell (if any) and value cells; categories along the other side.
  const series: { name: Ref | null; values: Ref }[] = [];
  let cats: Ref | null;
  if (byRow) {
    for (let row = first; row <= r.r2; row++) series.push({ name: labels ? { r1: row, c1: r.c1, r2: row, c2: r.c1 } : null, values: { r1: row, c1: left, r2: row, c2: r.c2 } });
    cats = header ? { r1: r.r1, c1: left, r2: r.r1, c2: r.c2 } : null;
  } else {
    for (let c = left; c <= r.c2; c++) series.push({ name: header ? { r1: r.r1, c1: c, r2: r.r1, c2: c } : null, values: { r1: first, c1: c, r2: r.r2, c2: c } });
    cats = labels ? { r1: first, c1: r.c1, r2: r.r2, c2: r.c1 } : null;
  }
  const cells = (x: Ref) => {
    const out: [number, number][] = [];
    for (let row = x.r1; row <= x.r2; row++) for (let c = x.c1; c <= x.c2; c++) out.push([row, c]);
    return out;
  };
  const f = (x: Ref) => `<c:f>${esc(absolute(sheet, x.r1, x.c1, x.r2, x.c2))}</c:f>`;
  const strCache = (vals: string[]) => `<c:strCache><c:ptCount val="${vals.length}"/>${vals.map((v, i) => `<c:pt idx="${i}"><c:v>${esc(v)}</c:v></c:pt>`).join("")}</c:strCache>`;
  const numCache = (vals: number[]) => `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${vals.length}"/>${vals.map((v, i) => (Number.isFinite(v) ? `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>` : "")).join("")}</c:numCache>`;
  const text = (x: Ref) => cells(x).map(([row, c]) => sheet.display(row, c).text);
  const ser = series
    .map((s, i) => {
      const name = s.name ? `<c:tx><c:strRef>${f(s.name)}${strCache(text(s.name))}</c:strRef></c:tx>` : "";
      const cat = cats ? `<c:cat><c:strRef>${f(cats)}${strCache(text(cats))}</c:strRef></c:cat>` : "";
      const vals = cells(s.values).map(([row, c]) => {
        const v = sheet.at(row, c)?.value;
        return typeof v === "number" ? v : NaN;
      });
      return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${name}${cat}<c:val><c:numRef>${f(s.values)}${numCache(vals)}</c:numRef></c:val></c:ser>`;
    })
    .join("");
  return (
    `<c:chartSpace xmlns:c="${C}" xmlns:a="${A}" xmlns:r="${R}"><c:roundedCorners val="0"/><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/>` +
    `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${ser}<c:gapWidth val="150"/><c:axId val="1"/><c:axId val="2"/></c:barChart>${axesXml(1, 2)}</c:plotArea>` +
    `<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>`
  );
}

/** A drawing marker at sheet pixel (x, y). */
function marker(sheet: Sheet, tag: string, x: number, y: number): string {
  const col = sheet.colAt(x), row = sheet.rowAt(y);
  const off = (v: number) => Math.max(0, Math.round(v * EMU_PER_PX));
  return `<xdr:${tag}><xdr:col>${col}</xdr:col><xdr:colOff>${off(x - sheet.colX(col))}</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>${off(y - sheet.rowY(row))}</xdr:rowOff></xdr:${tag}>`;
}

/** Insert a chart of the selection beside it (600 × 371 px, Google's default size); returns the chart part's path. */
export function insertChart(e: Edits, sheet: Sheet, r: Rect, format?: (f: ChartFormat) => void): string {
  const xml = parseXml(chartXml(sheet, r))!;
  if (format) {
    const f = readChartFormat(xml, chartColors(e.wb));
    format(f);
    setChartFormat(xml, f, chartColors(e.wb));
  }
  let path = "";
  e.structural(() => {
    const pkg = e.wb.pkg;
    const drawing = pkg.ensureDrawing(sheet.ref.path);
    if (!sheetChild(sheet, "drawing")) {
      sheetChild(sheet, "drawing", true)!.setAttributeNS(R, "r:id", drawing.rId);
      pkg.touch(sheet.ref.path);
    }
    const part = pkg.addChartPart(drawing.path, xml);
    path = part.path;
    const doc = pkg.xml(drawing.path)!;
    const ids = Array.from(doc.getElementsByTagNameNS(XDR, "cNvPr")).map((c) => Number(c.getAttribute("id")) || 0);
    const id = Math.max(0, ...ids) + 1;
    const x = sheet.colX(r.c2 + 1) + 24, y = sheet.rowY(r.r1);
    const anchor =
      `<xdr:twoCellAnchor xmlns:xdr="${XDR}" xmlns:a="${A}" editAs="oneCell">${marker(sheet, "from", x, y)}${marker(sheet, "to", x + 600, y + 371)}` +
      `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="Chart ${id}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
      `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${C}"><c:chart xmlns:c="${C}" xmlns:r="${R}" r:id="${part.rId}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`;
    doc.documentElement.appendChild(doc.importNode(parseXml(anchor)!.documentElement, true));
    pkg.touch(drawing.path);
  });
  return path;
}

/** Apply chart editor settings (type, title, legend...) to a chart part. */
export function formatChart(e: Edits, path: string, format: ChartFormat): void {
  e.structural(() => {
    const xml = e.wb.pkg.xml(path);
    if (!xml) return;
    setChartFormat(xml, format, chartColors(e.wb));
    e.wb.pkg.touch(path);
  });
}

/** Point a chart at another range: series per column, as when it was inserted; keeps its formatting. */
export function setChartSource(e: Edits, path: string, source: string): void {
  const at = resolve(e.wb, source.includes("!") ? source : `${quoteSheet(e.wb.sheets[0].name)}!${source}`);
  if (!at) throw new Error("Enter a range like Sheet1!A1:C10.");
  e.structural(() => {
    const xml = e.wb.pkg.xml(path);
    if (!xml) return;
    const fresh = parseXml(chartXml(at.sheet, at.rect))!;
    const target = Array.from(xml.getElementsByTagNameNS(C, "plotArea")[0]?.children ?? []).find((c) => c.localName.endsWith("Chart"));
    if (!target) return;
    const old = Array.from(target.children).filter((c) => c.localName === "ser");
    const next = Array.from(fresh.getElementsByTagNameNS(C, "ser"));
    // Keep each kept series' formatting (spPr, labels); drop or add series to match the new columns.
    next.forEach((ser, i) => {
      const keep = old[i];
      const imported = xml.importNode<Element>(ser, true);
      if (keep) {
        for (const tag of ["spPr", "dPt", "dLbls", "marker", "invertIfNegative", "smooth"]) {
          for (const el of Array.from(keep.children).filter((c) => c.localName === tag)) {
            const before = Array.from(imported.children).find((c) => ["cat", "val"].includes(c.localName)) ?? null;
            imported.insertBefore(el.cloneNode(true), tag === "smooth" ? null : before);
          }
        }
        keep.replaceWith(imported);
      } else {
        // After the last series (or where series go: before labels, gaps and axes).
        const last = Array.from(target.children).filter((c) => c.localName === "ser").pop();
        const AFTER_SER = ["dLbls", "dropLines", "hiLowLines", "upDownBars", "gapWidth", "overlap", "serLines", "marker", "smooth", "firstSliceAng", "holeSize", "shape", "axId", "extLst"];
        target.insertBefore(imported, last ? last.nextSibling : Array.from(target.children).find((c) => AFTER_SER.includes(c.localName)) ?? null);
      }
    });
    for (const ser of old.slice(next.length)) ser.remove();
    e.wb.pkg.touch(path);
  });
}

/** Move / resize a chart to `box` (sheet pixels): its anchor's cells and offsets follow, as when dragged in Sheets. */
export function moveChart(e: Edits, path: string, box: { x: number; y: number; w: number; h: number }): void {
  const sheet = e.wb.sheets.find((s) => sheetCharts(e.wb, s).some((c) => c.path === path));
  const chart = sheet && sheetCharts(e.wb, sheet).find((c) => c.path === path);
  if (!sheet || !chart) return;
  const x = Math.max(0, box.x), y = Math.max(0, box.y), w = Math.max(20, box.w), h = Math.max(20, box.h);
  e.structural(() => {
    const live = e.wb.pkg.xml(chart.drawing);
    const s = e.wb.sheets.find((x) => x.ref.path === sheet.ref.path)!;
    const anchor = Array.from(live?.documentElement.children ?? []).find((a) => Array.from(a.getElementsByTagNameNS(C, "chart")).some((c) => c.getAttributeNS(R, "id") === chart.rId));
    if (!anchor) return;
    const swap = (tag: string, px: number, py: number) => {
      const old = Array.from(anchor.children).find((c) => c.namespaceURI === XDR && c.localName === tag);
      const next = live!.importNode(parseXml(marker(s, tag, px, py).replace(`<xdr:${tag}>`, `<xdr:${tag} xmlns:xdr="${XDR}">`))!.documentElement, true);
      if (old) old.replaceWith(next);
      return !!old;
    };
    const ext = (el: Element | undefined) => {
      el?.setAttribute("cx", String(Math.round(w * EMU_PER_PX)));
      el?.setAttribute("cy", String(Math.round(h * EMU_PER_PX)));
    };
    const kid = (name: string) => Array.from(anchor.children).find((c) => c.namespaceURI === XDR && c.localName === name);
    if (anchor.localName === "absoluteAnchor") {
      kid("pos")?.setAttribute("x", String(Math.round(x * EMU_PER_PX)));
      kid("pos")?.setAttribute("y", String(Math.round(y * EMU_PER_PX)));
      ext(kid("ext"));
    } else {
      swap("from", x, y);
      if (!swap("to", x + w, y + h)) ext(kid("ext"));
    }
    e.wb.pkg.touch(chart.drawing);
  });
}

export function deleteChart(e: Edits, _sheet: Sheet, chart: SheetChart): void {
  e.structural(() => {
    const live = e.wb.pkg.xml(chart.drawing);
    const anchor = Array.from(live?.documentElement.children ?? []).find((a) => Array.from(a.getElementsByTagNameNS(C, "chart")).some((c) => c.getAttributeNS(R, "id") === chart.rId));
    anchor?.remove();
    e.wb.pkg.removeChartPart(chart.drawing, chart.rId);
    e.wb.pkg.touch(chart.drawing);
  });
}
