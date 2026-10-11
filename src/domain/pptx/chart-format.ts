/**
 * Chart formatting, as Google's chart editor offers it: chart type and stacking, title, legend, data labels, axis
 * titles, gridlines, ranges and series colours. Read from a c:chartSpace, and written back changing only what
 * differs, so formatting the editor doesn't model is kept.
 */
import { NS } from "../../core/utils/xml";
import { ColorContext, resolveColor } from "./color";

const C = "http://schemas.openxmlformats.org/drawingml/2006/chart";

export type ChartKind = "column" | "bar" | "line" | "area" | "pie" | "doughnut" | "scatter" | "other";
export type Stacking = "none" | "stacked" | "percent";
export type LegendPos = "b" | "t" | "l" | "r" | "none";

export interface ChartAxis {
  title: string;
  visible: boolean;
  gridlines: boolean;
  min: number | null;
  max: number | null;
  /** Drawn from the far end (orientation maxMin): a horizontal bar chart's first category at the top. */
  reversed: boolean;
  /** Number format of the labels, null for the data's own. */
  format: string | null;
}

export interface ChartFormat {
  kind: ChartKind;
  stacking: Stacking;
  /** "" for no title. */
  title: string;
  legend: LegendPos;
  /** Data labels: values, or percentages on a pie. */
  labels: boolean;
  /** Per series (per slice on a pie / doughnut); null for the default colour. */
  colors: (string | null)[];
  category: ChartAxis;
  value: ChartAxis;
}

/** Default series colours: the theme's accents, else Google's chart palette. */
export function chartPalette(ctx?: ColorContext): string[] {
  const accents = [1, 2, 3, 4, 5, 6].map((i) => ctx?.scheme[`accent${i}`]).filter((c): c is string => !!c);
  return accents.length === 6 ? accents : ["#4285F4", "#EA4335", "#FBBC04", "#34A853", "#FF6D01", "#46BDC6"];
}

// -----------------------------------------------------------------------------
// XML helpers
// -----------------------------------------------------------------------------

const kids = (el: Element | null | undefined, name: string, ns = C) => (el ? Array.from(el.children).filter((c) => c.namespaceURI === ns && c.localName === name) : []);
const kid = (el: Element | null | undefined, name: string, ns = C) => kids(el, name, ns)[0] ?? null;
const val = (el: Element | null) => el?.getAttribute("val") ?? null;
/** CT_Boolean: present means true unless val says otherwise. */
const on = (el: Element | null) => !!el && !["0", "false"].includes(el.getAttribute("val") ?? "1");
const num = (el: Element | null) => (el && Number.isFinite(Number(val(el))) ? Number(val(el)) : null);

function make(doc: Document, name: string, attrs: Record<string, string | number> = {}): Element {
  const e = doc.createElementNS(C, `c:${name}`);
  for (const [k, v] of Object.entries(attrs)) e.setAttribute(k, String(v));
  return e;
}
function parse(doc: Document, xml: string): Element {
  const root = new DOMParser().parseFromString(`<r xmlns:c="${C}" xmlns:a="${NS.a}">${xml}</r>`, "application/xml").documentElement;
  return doc.importNode<Element>(root.firstElementChild!, true);
}

const CHART_ORDER = ["title", "autoTitleDeleted", "pivotFmts", "view3D", "floor", "sideWall", "backWall", "plotArea", "legend", "plotVisOnly", "dispBlanksAs", "showDLblsOverMax", "extLst"];
const TYPE_ORDER = ["barDir", "radarStyle", "scatterStyle", "grouping", "varyColors", "ser", "dLbls", "dropLines", "hiLowLines", "upDownBars", "gapWidth", "overlap", "serLines", "marker", "smooth", "firstSliceAng", "holeSize", "shape", "axId", "extLst"];
const SER_ORDER = ["idx", "order", "tx", "spPr", "invertIfNegative", "pictureOptions", "marker", "explosion", "dPt", "dLbls", "trendline", "errBars", "xVal", "yVal", "cat", "val", "smooth", "shape", "bubbleSize", "bubble3D", "extLst"];
const AX_ORDER = ["axId", "scaling", "delete", "axPos", "majorGridlines", "minorGridlines", "title", "numFmt", "majorTickMark", "minorTickMark", "tickLblPos", "spPr", "txPr", "crossAx", "crosses", "crossesAt", "crossBetween", "auto", "lblAlgn", "lblOffset", "baseTimeUnit", "majorUnit", "majorTimeUnit", "minorUnit", "minorTimeUnit", "tickLblSkip", "tickMarkSkip", "dispUnits", "noMultiLvlLbl", "extLst"];
const SCALING_ORDER = ["logBase", "orientation", "max", "min", "extLst"];
const LEGEND_ORDER = ["legendPos", "legendEntry", "layout", "overlay", "spPr", "txPr", "extLst"];

/** Insert `child` in `parent` in schema `order`, after any elements of its own name. */
function insert(parent: Element, child: Element, order: string[]): Element {
  const at = order.indexOf(child.localName);
  const before = Array.from(parent.children).find((c) => order.indexOf(c.localName) > at) ?? null;
  return parent.insertBefore(child, before);
}
/** Put `child` in `parent` in schema `order`, replacing an element of the same name. */
function place(parent: Element, child: Element, order: string[]): Element {
  const old = kid(parent, child.localName, child.namespaceURI ?? C);
  if (!old) return insert(parent, child, order);
  old.replaceWith(child);
  return child;
}
const setVal = (parent: Element, name: string, value: string | number, order: string[]) => place(parent, make(parent.ownerDocument, name, { val: value }), order);

const richText = (el: Element | null) => {
  const tx = kid(el, "tx");
  if (!tx) return null;
  const rich = kid(tx, "rich");
  if (rich) return Array.from(rich.getElementsByTagNameNS(NS.a, "p")).map((p) => Array.from(p.getElementsByTagNameNS(NS.a, "t")).map((t) => t.textContent).join("")).join(" ");
  return Array.from(tx.getElementsByTagNameNS(C, "v")).map((v) => v.textContent).join(" ");
};

/** A c:title holding `text`, keeping the existing title's formatting (its first run) when there is one. */
function titleEl(doc: Document, old: Element | null, text: string, vertical: boolean): Element {
  const rich = kid(kid(old, "tx"), "rich");
  const run = rich?.getElementsByTagNameNS(NS.a, "r")[0];
  if (old && rich && run) {
    const p = run.parentElement!;
    for (const other of Array.from(rich.getElementsByTagNameNS(NS.a, "p"))) if (other !== p) other.remove();
    for (const r of Array.from(p.children)) if (r !== run && (r.localName === "r" || r.localName === "br" || r.localName === "fld")) r.remove();
    run.getElementsByTagNameNS(NS.a, "t")[0].textContent = text;
    return old;
  }
  const esc = text.replace(/&/g, "&amp;").replace(/</g, "&lt;");
  const title = parse(doc, `<c:title><c:tx><c:rich><a:bodyPr${vertical ? ' rot="-5400000" vert="horz"' : ""}/><a:lstStyle/><a:p><a:pPr><a:defRPr b="0"/></a:pPr><a:r><a:t>${esc}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`);
  // An existing title without text (an automatic one) keeps its layout and formatting.
  if (old) for (const c of Array.from(old.children)) if (c.localName !== "tx") title.appendChild(c.cloneNode(true));
  kids(title, "overlay").slice(1).forEach((e) => e.remove());
  return title;
}

// -----------------------------------------------------------------------------
// Reading
// -----------------------------------------------------------------------------

interface Parts {
  chart: Element;
  plot: Element;
  type: Element | null;
  series: Element[];
  category: Element | null;
  value: Element | null;
}

function parts(xml: Document): Parts | null {
  const chart = xml.getElementsByTagNameNS(C, "chart")[0];
  const plot = kid(chart, "plotArea");
  if (!chart || !plot) return null;
  const type = Array.from(plot.children).find((c) => c.namespaceURI === C && c.localName.endsWith("Chart")) ?? null;
  const ids = kids(type, "axId").map(val);
  const axes = Array.from(plot.children).filter((c) => c.namespaceURI === C && c.localName.endsWith("Ax") && ids.includes(val(kid(c, "axId"))));
  let category = axes.find((a) => a.localName === "catAx" || a.localName === "dateAx") ?? null;
  let value = axes.find((a) => a.localName === "valAx") ?? null;
  if (!category && type?.localName === "scatterChart") {
    // A scatter chart's x axis is a value axis along the bottom.
    category = axes.find((a) => a.localName === "valAx" && ["b", "t"].includes(val(kid(a, "axPos")) ?? "")) ?? null;
    value = axes.find((a) => a.localName === "valAx" && a !== category) ?? null;
  }
  return { chart, plot, type, series: kids(type, "ser"), category, value };
}

function kindOf(type: Element | null): ChartKind {
  switch (type?.localName) {
    case "barChart":
    case "bar3DChart":
      return val(kid(type, "barDir")) === "bar" ? "bar" : "column";
    case "lineChart":
    case "line3DChart":
      return "line";
    case "areaChart":
    case "area3DChart":
      return "area";
    case "pieChart":
    case "pie3DChart":
    case "ofPieChart":
      return "pie";
    case "doughnutChart":
      return "doughnut";
    case "scatterChart":
      return "scatter";
    default:
      return "other";
  }
}

const isPie = (kind: ChartKind) => kind === "pie" || kind === "doughnut";

function readAxis(ax: Element | null): ChartAxis {
  const scaling = kid(ax, "scaling");
  const fmt = kid(ax, "numFmt");
  return {
    title: richText(kid(ax, "title")) ?? (kid(ax, "title") ? "Axis Title" : ""),
    visible: !!ax && !on(kid(ax, "delete")),
    gridlines: !!kid(ax, "majorGridlines"),
    min: num(kid(scaling, "min")),
    max: num(kid(scaling, "max")),
    reversed: val(kid(scaling, "orientation")) === "maxMin",
    format: fmt && fmt.getAttribute("sourceLinked") !== "1" && fmt.getAttribute("formatCode") !== "General" ? fmt.getAttribute("formatCode") : null,
  };
}

const labelled = (dLbls: Element | null) => !!dLbls && !on(kid(dLbls, "delete")) && (on(kid(dLbls, "showVal")) || on(kid(dLbls, "showPercent")));

function fillOf(spPr: Element | null, ctx: ColorContext, line: boolean): string | null {
  const fill = line ? kid(kid(spPr, "ln", NS.a), "solidFill", NS.a) : kid(spPr, "solidFill", NS.a);
  return fill ? resolveColor(fill, ctx)?.slice(0, 7) ?? null : null;
}

export function readChartFormat(xml: Document, ctx: ColorContext = { scheme: {}, map: {} }): ChartFormat {
  const p = parts(xml);
  const kind = kindOf(p?.type ?? null);
  const grouping = val(kid(p?.type, "grouping"));
  const titleNode = kid(p?.chart, "title");
  const autoTitle = p && p.series.length === 1 ? richText(kid(p.series[0], "tx")) ?? "" : "Chart Title";
  const legend = kid(p?.chart, "legend");
  const typeLabels = kid(p?.type, "dLbls");
  let colors: (string | null)[];
  if (isPie(kind)) {
    const ser = p?.series[0];
    const count = Number(val(kid(ser, "val")?.getElementsByTagNameNS(C, "ptCount")[0] ?? null) ?? 0);
    const points = new Map(kids(ser, "dPt").map((d) => [Number(val(kid(d, "idx"))), fillOf(kid(d, "spPr"), ctx, false)]));
    colors = Array.from({ length: count }, (_, i) => points.get(i) ?? null);
  } else colors = (p?.series ?? []).map((s) => fillOf(kid(s, "spPr"), ctx, kind === "line" || kind === "scatter"));
  return {
    kind,
    stacking: grouping === "stacked" ? "stacked" : grouping === "percentStacked" ? "percent" : "none",
    title: titleNode && !on(kid(p!.chart, "autoTitleDeleted")) ? richText(titleNode) ?? autoTitle : "",
    legend: legend ? ((val(kid(legend, "legendPos")) ?? "r") as LegendPos) : "none",
    labels: (p?.series ?? []).some((s) => (kid(s, "dLbls") ? labelled(kid(s, "dLbls")) : labelled(typeLabels))),
    colors,
    category: readAxis(p?.category ?? null),
    value: readAxis(p?.value ?? null),
  };
}

// -----------------------------------------------------------------------------
// Writing
// -----------------------------------------------------------------------------

const TYPE_TAG: Partial<Record<ChartKind, string>> = { column: "barChart", bar: "barChart", line: "lineChart", area: "areaChart", pie: "pieChart", doughnut: "doughnutChart" };
/** Series children every converted chart type accepts. */
const SER_KEEP = new Set(["idx", "order", "tx", "spPr", "dLbls", "cat", "val", "extLst"]);

/** A category and a value axis, as a new chart gets them (gridlines on the value axis, as in Google's charts). */
export function axesXml(cat: number, value: number): string {
  return (
    `<c:catAx><c:axId val="${cat}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${value}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>` +
    `<c:valAx><c:axId val="${value}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${cat}"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`
  );
}

/** Rebuild the plot's chart-type element as another type, keeping its series, data labels and axes. */
function convert(p: Parts, kind: ChartKind, stacking: Stacking): void {
  const doc = p.plot.ownerDocument;
  const old = p.type!;
  const tag = TYPE_TAG[kind]!;
  const next = make(doc, tag);
  const add = (name: string, value: string | number) => next.appendChild(make(doc, name, { val: value }));
  const pie = isPie(kind);
  if (tag === "barChart") add("barDir", kind === "bar" ? "bar" : "col");
  if (!pie) add("grouping", stacking === "stacked" ? "stacked" : stacking === "percent" ? "percentStacked" : tag === "barChart" ? "clustered" : "standard");
  add("varyColors", pie ? 1 : 0);
  for (const ser of p.series) {
    const s = make(doc, "ser");
    for (const c of Array.from(ser.children)) if (SER_KEEP.has(c.localName)) s.appendChild(c);
    next.appendChild(s);
  }
  const labels = kid(old, "dLbls");
  if (labels) next.appendChild(labels);
  if (tag === "barChart") {
    add("gapWidth", 150);
    if (stacking !== "none") add("overlap", 100);
  }
  if (tag === "lineChart") add("marker", 1);
  if (pie) add("firstSliceAng", 0);
  if (kind === "doughnut") add("holeSize", 50);
  const ids = kids(old, "axId").map(val);
  if (pie) {
    for (const ax of Array.from(p.plot.children)) if (ax.localName.endsWith("Ax") && ids.includes(val(kid(ax, "axId")))) ax.remove();
  } else if (ids.length >= 2 && p.category && p.value) {
    for (const id of ids) add("axId", id!);
  } else {
    const used = Array.from(p.plot.getElementsByTagNameNS(C, "axId")).map((a) => Number(val(a)) || 0);
    const cat = Math.max(1000, ...used) + 1;
    add("axId", cat);
    add("axId", cat + 1);
    const axes = parse(doc, `<x xmlns:c="${C}">${axesXml(cat, cat + 1)}</x>`);
    const before = Array.from(p.plot.children).find((c) => ["dTable", "spPr", "extLst"].includes(c.localName)) ?? null;
    for (const ax of Array.from(axes.children)) p.plot.insertBefore(ax, before);
  }
  old.replaceWith(next);
}

function setFill(spPr: Element, color: string | null, line: boolean): void {
  const doc = spPr.ownerDocument;
  const FILLS = ["noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"];
  let target = spPr;
  if (line) {
    target = kid(spPr, "ln", NS.a) ?? spPr.insertBefore(doc.createElementNS(NS.a, "a:ln"), Array.from(spPr.children).find((c) => ["effectLst", "effectDag", "scene3d", "sp3d", "extLst"].includes(c.localName)) ?? null);
    if (!target.hasAttribute("w")) target.setAttribute("w", "28575");
  }
  for (const c of Array.from(target.children)) if (FILLS.includes(c.localName)) c.remove();
  if (!color) return;
  const fill = doc.createElementNS(NS.a, "a:solidFill");
  fill.appendChild(doc.createElementNS(NS.a, "a:srgbClr")).setAttribute("val", color.replace("#", "").slice(0, 6).toUpperCase());
  const before = line ? target.firstElementChild : Array.from(target.children).find((c) => !["xfrm", "custGeom", "prstGeom"].includes(c.localName)) ?? null;
  target.insertBefore(fill, before);
}

function setAxis(ax: Element | null, was: ChartAxis, next: ChartAxis, vertical: boolean): void {
  if (!ax) return;
  const doc = ax.ownerDocument;
  if (next.title !== was.title) {
    if (next.title) place(ax, titleEl(doc, kid(ax, "title"), next.title, vertical), AX_ORDER);
    else kid(ax, "title")?.remove();
  }
  if (next.visible !== was.visible) setVal(ax, "delete", next.visible ? 0 : 1, AX_ORDER);
  if (next.gridlines !== was.gridlines) {
    if (next.gridlines) place(ax, make(doc, "majorGridlines"), AX_ORDER);
    else kid(ax, "majorGridlines")?.remove();
  }
  const scaling = kid(ax, "scaling") ?? place(ax, make(doc, "scaling"), AX_ORDER);
  for (const end of ["min", "max"] as const) {
    if (next[end] === was[end]) continue;
    if (next[end] === null) kid(scaling, end)?.remove();
    else setVal(scaling, end, next[end], SCALING_ORDER);
  }
  if (next.reversed !== was.reversed) setVal(scaling, "orientation", next.reversed ? "maxMin" : "minMax", SCALING_ORDER);
  if (next.format !== was.format) {
    const fmt = make(doc, "numFmt", { formatCode: next.format ?? "General", sourceLinked: next.format ? 0 : 1 });
    place(ax, fmt, AX_ORDER);
  }
}

/** Write `next` into the chart, changing only what differs from what it shows now. */
export function setChartFormat(xml: Document, next: ChartFormat, ctx: ColorContext = { scheme: {}, map: {} }): void {
  let p = parts(xml);
  if (!p?.type) return;
  const doc = xml;
  let was = readChartFormat(xml, ctx);
  const convertible = (k: ChartKind) => !!TYPE_TAG[k];
  const single = Array.from(p.plot.children).filter((c) => c.localName.endsWith("Chart")).length === 1;
  if (next.kind !== was.kind && convertible(next.kind) && convertible(was.kind) && single && TYPE_TAG[next.kind] !== p.type.localName) {
    convert(p, next.kind, next.stacking);
    p = parts(xml)!;
    was = readChartFormat(xml, ctx);
  }
  const type = p.type!;
  const kind = was.kind;
  if ((kind === "column" || kind === "bar") && (next.kind === "column" || next.kind === "bar") && next.kind !== kind) {
    setVal(type, "barDir", next.kind === "bar" ? "bar" : "col", TYPE_ORDER);
    if (p.category) setVal(p.category, "axPos", next.kind === "bar" ? "l" : "b", AX_ORDER);
    if (p.value) setVal(p.value, "axPos", next.kind === "bar" ? "b" : "l", AX_ORDER);
  }
  const kindNow = kindOf(type);
  if (next.stacking !== was.stacking && !isPie(kindNow) && kindNow !== "scatter") {
    const bar = type.localName === "barChart";
    setVal(type, "grouping", next.stacking === "stacked" ? "stacked" : next.stacking === "percent" ? "percentStacked" : bar ? "clustered" : "standard", TYPE_ORDER);
    if (bar) {
      if (next.stacking === "none") kid(type, "overlap")?.remove();
      else setVal(type, "overlap", 100, TYPE_ORDER);
    }
  }
  if (next.title !== was.title) {
    if (next.title) {
      place(p.chart, titleEl(doc, kid(p.chart, "title"), next.title, false), CHART_ORDER);
      setVal(p.chart, "autoTitleDeleted", 0, CHART_ORDER);
    } else {
      kid(p.chart, "title")?.remove();
      setVal(p.chart, "autoTitleDeleted", 1, CHART_ORDER);
    }
  }
  if (next.legend !== was.legend) {
    if (next.legend === "none") kid(p.chart, "legend")?.remove();
    else {
      const legend = kid(p.chart, "legend") ?? place(p.chart, parse(doc, `<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>`), CHART_ORDER);
      setVal(legend, "legendPos", next.legend, LEGEND_ORDER);
    }
  }
  if (next.labels !== was.labels) {
    for (const s of p.series) kid(s, "dLbls")?.remove();
    kid(type, "dLbls")?.remove();
    if (next.labels) {
      const pie = isPie(kindNow);
      const flags = `<c:showLegendKey val="0"/><c:showVal val="${pie ? 0 : 1}"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="${pie ? 1 : 0}"/><c:showBubbleSize val="0"/>`;
      place(type, parse(doc, `<c:dLbls>${flags}</c:dLbls>`), TYPE_ORDER);
    }
  }
  const barVertical = kindNow === "bar";
  setAxis(p.category, was.category, next.category, barVertical);
  setAxis(p.value, was.value, next.value, !barVertical);
  // A reversed category axis on a bar chart keeps its value axis at the bottom (crossing at the last category).
  if (p.value && p.category && next.category.reversed !== was.category.reversed && kindNow === "bar")
    setVal(p.value, "crosses", next.category.reversed ? "max" : "autoZero", AX_ORDER);

  const colorsNow = readChartFormat(xml, ctx).colors;
  next.colors.forEach((color, i) => {
    if ((color ?? null) === (colorsNow[i] ?? null) || i >= colorsNow.length) return;
    if (isPie(kindNow)) {
      const ser = p.series[0];
      let point = kids(ser, "dPt").find((d) => Number(val(kid(d, "idx"))) === i);
      if (!point) {
        point = parse(doc, `<c:dPt><c:idx val="${i}"/><c:bubble3D val="0"/><c:spPr/></c:dPt>`);
        const after = kids(ser, "dPt").find((d) => Number(val(kid(d, "idx"))) > i);
        if (after) ser.insertBefore(point, after);
        else insert(ser, point, SER_ORDER);
      }
      const spPr = kid(point, "spPr") ?? point.appendChild(make(doc, "spPr"));
      setFill(spPr, color, false);
    } else {
      const ser = p.series[i];
      const spPr = kid(ser, "spPr") ?? place(ser, make(doc, "spPr"), SER_ORDER);
      setFill(spPr, color, kindNow === "line" || kindNow === "scatter");
    }
  });
}
