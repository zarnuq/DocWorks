import type { ChartModel } from "../../domain/pptx/chart";
import { chartPalette, ChartAxis, ChartFormat } from "../../domain/pptx/chart-format";

const NS = "http://www.w3.org/2000/svg";
const INK = "#595959";
const LINE = "#BFBFBF";
const GRID = "#D9D9D9";
/** Label and title sizes, in the chart's units (points). */
const FONT = 9;
const TITLE = 14;
const PAD = 6;

function node(parent: Element, tag: string, attrs: Record<string, string | number> = {}, text?: string): SVGElement {
  const el = document.createElementNS(NS, tag);
  for (const [key, value] of Object.entries(attrs)) el.setAttribute(key, String(value));
  if (text !== undefined) el.textContent = text;
  parent.append(el);
  return el;
}

/** Rough text width: charts have short labels and no layout pass to measure them in. */
const textWidth = (text: string, size = FONT) => text.length * size * 0.55;

/** Round axis steps: about `count` steps of 1, 2, 2.5 or 5 × 10^n covering lo..hi. */
export function niceTicks(lo: number, hi: number, count = 5): number[] {
  if (hi - lo < 1e-12) hi = lo + 1;
  const raw = (hi - lo) / count;
  const mag = 10 ** Math.floor(Math.log10(raw));
  const step = [1, 2, 2.5, 5, 10].map((m) => m * mag).find((s) => s >= raw - 1e-12)!;
  const out: number[] = [];
  for (let v = Math.floor(lo / step + 1e-9) * step; v <= Math.ceil(hi / step - 1e-9) * step + step / 2; v += step) out.push(Number(v.toPrecision(12)));
  return out;
}

/** A value in an axis / label number format: percentages, fixed decimals and thousands separators. */
export function formatNumber(v: number, format: string | null): string {
  const decimals = (format?.match(/\.(0+)/)?.[1].length ?? null);
  if (format?.includes("%")) return `${(v * 100).toFixed(decimals ?? 0)}%`;
  if (format?.includes(",")) return v.toLocaleString("en-US", { minimumFractionDigits: decimals ?? 0, maximumFractionDigits: decimals ?? 0 });
  if (decimals !== null) return v.toFixed(decimals);
  return String(Number(v.toPrecision(10)));
}

const DEFAULT_AXIS: ChartAxis = { title: "", visible: true, gridlines: false, min: null, max: null, reversed: false, format: null };
const DEFAULT_FORMAT: ChartFormat = { kind: "column", stacking: "none", title: "", legend: "b", labels: false, colors: [], category: DEFAULT_AXIS, value: { ...DEFAULT_AXIS, gridlines: true } };

interface Box { left: number; top: number; right: number; bottom: number; }

/** The legend in its place, shrinking `box` (the space left for the plot). */
function legend(svg: Element, box: Box, pos: ChartFormat["legend"], labels: string[], color: (i: number) => string): void {
  if (pos === "none" || !labels.length) return;
  const g = node(svg, "g", { class: "chart-legend" });
  const lh = FONT * 1.6;
  const item = (label: string, x: number, y: number, i: number) => {
    node(g, "rect", { x, y: y - FONT * 0.8, width: FONT * 0.9, height: FONT * 0.9, fill: color(i) });
    node(g, "text", { x: x + FONT * 1.3, y, "font-size": FONT, fill: INK }, label);
  };
  const widths = labels.map((l) => FONT * 2 + textWidth(l));
  if (pos === "l" || pos === "r") {
    const w = Math.min((box.right - box.left) / 3, Math.max(...widths));
    const top = (box.top + box.bottom) / 2 - (labels.length * lh) / 2 + FONT;
    const x = pos === "r" ? box.right - w : box.left;
    labels.forEach((l, i) => item(l, x, top + i * lh, i));
    if (pos === "r") box.right -= w + PAD;
    else box.left += w + PAD;
    return;
  }
  // Rows of entries, centred, wrapping to the chart's width.
  const rows: number[][] = [[]];
  let used = 0;
  widths.forEach((w, i) => {
    if (used + w > box.right - box.left && rows[rows.length - 1].length) {
      rows.push([]);
      used = 0;
    }
    rows[rows.length - 1].push(i);
    used += w;
  });
  const height = rows.length * lh;
  const top = pos === "t" ? box.top + FONT : box.bottom - height + FONT;
  rows.forEach((row, r) => {
    let x = (box.left + box.right) / 2 - row.reduce((s, i) => s + widths[i], 0) / 2;
    for (const i of row) {
      item(labels[i], x, top + r * lh, i);
      x += widths[i];
    }
  });
  if (pos === "t") box.top += height + PAD;
  else box.bottom -= height + PAD;
}

/** Common chart types drawn from their saved caches and formatting, `size` in points (the frame's size). */
export function renderChart(host: HTMLElement, chart: ChartModel, size = { w: 480, h: 300 }): void {
  const { w: W, h: H } = size;
  const svg = node(host, "svg", { viewBox: `0 0 ${W} ${H}`, width: "100%", height: "100%", role: "img", "aria-label": "Chart", "font-family": "Arial, sans-serif" });
  const { categories, series } = chart.data;
  const f = chart.format ?? { ...DEFAULT_FORMAT, kind: chart.type === "barChart" ? "column" : chart.type === "lineChart" ? "line" : chart.type === "areaChart" ? "area" : chart.type === "pieChart" ? "pie" : chart.type === "doughnutChart" ? "doughnut" : "column" };
  if (!series.length || !categories.length) {
    node(svg, "text", { x: PAD, y: TITLE + PAD, "font-size": FONT, fill: INK }, "Chart has no cached data");
    return;
  }
  const palette = chart.palette ?? chartPalette();
  const color = (i: number) => f.colors[i] ?? palette[i % palette.length];
  const box: Box = { left: PAD, top: PAD, right: W - PAD, bottom: H - PAD };
  if (f.title) {
    node(svg, "text", { x: W / 2, y: box.top + TITLE, "text-anchor": "middle", "font-size": TITLE, fill: INK, class: "chart-title" }, f.title);
    box.top += TITLE * 1.6;
  }
  const pie = f.kind === "pie" || f.kind === "doughnut";
  legend(svg, box, f.legend, pie ? categories : series.map((s) => s.name), color);

  if (pie) {
    const values = series[0].values.map((v) => Math.max(0, v));
    const total = values.reduce((a, b) => a + b, 0);
    const cx = (box.left + box.right) / 2, cy = (box.top + box.bottom) / 2;
    const r = Math.max(4, Math.min(box.right - box.left, box.bottom - box.top) / 2 - 2);
    let angle = -Math.PI / 2;
    values.forEach((value, i) => {
      if (!total || !value) return;
      const next = angle + (value / total) * Math.PI * 2;
      const at = (a: number, k = 1) => `${cx + r * k * Math.cos(a)} ${cy + r * k * Math.sin(a)}`;
      if (value === total) node(svg, "circle", { cx, cy, r, fill: color(i), class: "chart-mark" });
      else node(svg, "path", { d: `M${cx} ${cy} L${at(angle)} A${r} ${r} 0 ${next - angle > Math.PI ? 1 : 0} 1 ${at(next)} Z`, fill: color(i), stroke: "#FFFFFF", "stroke-width": 1, class: "chart-mark" });
      if (f.labels) {
        const mid = (angle + next) / 2, k = f.kind === "doughnut" ? 0.75 : 0.62;
        node(svg, "text", { x: cx + r * k * Math.cos(mid), y: cy + r * k * Math.sin(mid) + FONT / 3, "text-anchor": "middle", "font-size": FONT, fill: "#FFFFFF", class: "chart-label" }, `${Math.round((value / total) * 100)}%`);
      }
      angle = next;
    });
    if (f.kind === "doughnut") node(svg, "circle", { cx, cy, r: r / 2, fill: "#FFFFFF" });
    return;
  }

  // Axes charts. Values are stacked (and made fractions of the total for 100%) before scaling.
  const horizontal = f.kind === "bar";
  const combo = f.kind === "combo";
  /** Series drawn as lines (a line chart, or a combo's line series); the rest of a combo are columns. */
  const isLine = (j: number) => f.kind === "line" || (combo && f.seriesKinds?.[j] === "line");
  const stacked = f.stacking !== "none" && f.kind !== "scatter" && !combo;
  const percent = f.stacking === "percent";
  const n = categories.length;
  // A scatter chart's categories are x values: placed by value along a numeric axis.
  const xs = categories.map(Number);
  const numericX = f.kind === "scatter" && n > 0 && categories.every((c, i) => c.trim() !== "" && Number.isFinite(xs[i]));
  const xTicks = numericX ? niceTicks(Math.min(...xs), Math.max(...xs)) : [];
  const totals = categories.map((_, i) => series.reduce((s, ser) => s + Math.abs(ser.values[i] || 0), 0) || 1);
  const shown = series.map((s) => s.values.map((v, i) => (percent ? (v || 0) / totals[i] : v || 0)));
  /** Start and end of each series' segment per category. */
  const spans = shown.map((vals) => vals.map((v) => [0, v]));
  if (stacked) {
    for (let i = 0; i < n; i++) {
      let pos = 0, neg = 0;
      shown.forEach((vals, j) => {
        const v = vals[i];
        if (f.kind !== "column" && f.kind !== "bar") {
          spans[j][i] = [pos, pos + v];
          pos += v;
        } else if (v >= 0) {
          spans[j][i] = [pos, pos + v];
          pos += v;
        } else {
          spans[j][i] = [neg, neg + v];
          neg += v;
        }
      });
    }
  }
  const ends = spans.flatMap((s) => s.flatMap((x) => x));
  const ticks = niceTicks(f.value.min ?? Math.min(0, ...ends), f.value.max ?? Math.max(0, ...ends));
  const lo = f.value.min ?? ticks[0], hi = f.value.max ?? ticks[ticks.length - 1];
  const tickValues = ticks.filter((t) => t >= lo - 1e-9 && t <= hi + 1e-9);
  const valueFormat = f.value.format ?? (percent ? "0%" : null);
  const tickLabels = tickValues.map((t) => formatNumber(t, valueFormat));

  // Room for axis labels and titles.
  const titleRoom = FONT * 1.8;
  const valueRoom = f.value.visible ? (horizontal ? FONT * 1.6 : Math.max(...tickLabels.map((t) => textWidth(t))) + 4) : 0;
  const catRoom = f.category.visible ? (horizontal ? Math.min((W - 2 * PAD) / 3, Math.max(...categories.map((c) => textWidth(c))) + 4) : FONT * 1.6) : 0;
  const leftTitle = horizontal ? f.category.title : f.value.title;
  const bottomTitle = horizontal ? f.value.title : f.category.title;
  if (leftTitle) {
    node(svg, "text", { x: 0, y: 0, transform: `translate(${box.left + FONT} ${(box.top + box.bottom) / 2}) rotate(-90)`, "text-anchor": "middle", "font-size": FONT, fill: INK, class: "chart-axis-title" }, leftTitle);
    box.left += titleRoom;
  }
  if (bottomTitle) {
    node(svg, "text", { x: (box.left + box.right) / 2, y: box.bottom - 2, "text-anchor": "middle", "font-size": FONT, fill: INK, class: "chart-axis-title" }, bottomTitle);
    box.bottom -= titleRoom;
  }
  box.left += horizontal ? catRoom : valueRoom;
  box.bottom -= horizontal ? valueRoom : catRoom;
  box.right -= PAD;
  const plotW = Math.max(1, box.right - box.left), plotH = Math.max(1, box.bottom - box.top);

  /** Where a value falls along the value axis (y for columns, x for bars). */
  const scale = (v: number) => {
    let t = (Math.min(hi, Math.max(lo, v)) - lo) / (hi - lo || 1);
    if (f.value.reversed) t = 1 - t;
    return horizontal ? box.left + t * plotW : box.bottom - t * plotH;
  };
  const band = (horizontal ? plotH : plotW) / n;
  const xLo = xTicks[0] ?? 0, xHi = xTicks[xTicks.length - 1] ?? 1;
  const xScale = (v: number) => box.left + ((v - xLo) / (xHi - xLo || 1)) * plotW;
  /** Centre of category `i`: left to right for columns, bottom to top for bars (Office's order), unless reversed. */
  const centre = (i: number) => {
    if (numericX) return xScale(xs[i]);
    const k = f.category.reversed ? n - 1 - i : i;
    return horizontal ? box.bottom - band * (k + 0.5) : box.left + band * (k + 0.5);
  };

  // Gridlines and value labels.
  for (const [k, t] of tickValues.entries()) {
    const p = scale(t);
    if (f.value.gridlines) node(svg, "line", horizontal ? { x1: p, x2: p, y1: box.top, y2: box.bottom, stroke: GRID } : { x1: box.left, x2: box.right, y1: p, y2: p, stroke: GRID });
    if (f.value.visible)
      node(svg, "text", horizontal ? { x: p, y: box.bottom + FONT * 1.3, "text-anchor": "middle", "font-size": FONT, fill: INK } : { x: box.left - 4, y: p + FONT / 3, "text-anchor": "end", "font-size": FONT, fill: INK }, tickLabels[k]);
  }
  if (f.category.gridlines && numericX) for (const t of xTicks) node(svg, "line", { x1: xScale(t), x2: xScale(t), y1: box.top, y2: box.bottom, stroke: GRID });
  else if (f.category.gridlines)
    for (let i = 0; i <= n; i++) {
      const p = horizontal ? box.bottom - band * i : box.left + band * i;
      node(svg, "line", horizontal ? { x1: box.left, x2: box.right, y1: p, y2: p, stroke: GRID } : { x1: p, x2: p, y1: box.top, y2: box.bottom, stroke: GRID });
    }
  const zero = scale(0);
  node(svg, "line", horizontal ? { x1: zero, x2: zero, y1: box.top, y2: box.bottom, stroke: LINE } : { x1: box.left, x2: box.right, y1: zero, y2: zero, stroke: LINE });
  if (f.category.visible && numericX) {
    for (const t of xTicks) node(svg, "text", { x: xScale(t), y: box.bottom + FONT * 1.3, "text-anchor": "middle", "font-size": FONT, fill: INK }, formatNumber(t, f.category.format));
  } else if (f.category.visible) {
    // Labels fit their slot: shortened with "…", and only every k-th one when even short ones would collide.
    const charW = FONT * 0.55;
    // Room for the longest label, or 8 characters, whichever is less, decides how many slots a label may use.
    const need = Math.min(Math.max(...categories.map((c) => c.length)), 8) * charW;
    const step = horizontal ? Math.max(1, Math.ceil((FONT * 1.2) / band)) : Math.max(1, Math.ceil(need / band));
    const chars = Math.max(3, Math.floor((horizontal ? box.left - PAD - 4 : band * step) / charW));
    const fit = (t: string) => (t.length > chars ? `${t.slice(0, chars - 1)}…` : t);
    categories.forEach((name, i) => {
      if (i % step) return;
      node(svg, "text", horizontal ? { x: box.left - 4, y: centre(i) + FONT / 3, "text-anchor": "end", "font-size": FONT, fill: INK } : { x: centre(i), y: box.bottom + FONT * 1.3, "text-anchor": "middle", "font-size": FONT, fill: INK }, fit(name));
    });
  }

  const label = (text: string, along: number, across: number, anchorEnd: boolean, inside: boolean) => {
    const attrs = horizontal
      ? { x: along + (inside ? 0 : anchorEnd ? -3 : 3), y: across + FONT / 3, "text-anchor": inside ? "middle" : anchorEnd ? "end" : "start" }
      : { x: across, y: along + (inside ? FONT / 3 : anchorEnd ? FONT + 2 : -3), "text-anchor": "middle" };
    node(svg, "text", { ...attrs, "font-size": FONT, fill: inside ? "#FFFFFF" : INK, class: "chart-label" }, text);
  };
  const labelText = (j: number, i: number) => formatNumber(percent ? shown[j][i] : series[j].values[i], percent ? "0%" : f.value.format);

  if (f.kind === "column" || f.kind === "bar" || f.kind === "other" || combo) {
    const group = band * 0.6;
    const bars = spans.map((_, j) => j).filter((j) => !isLine(j));
    const thick = stacked ? group : group / Math.max(1, bars.length);
    spans.forEach((vals, j) => {
      if (isLine(j)) return;
      const slot = bars.indexOf(j);
      vals.forEach(([a, b], i) => {
        const p1 = scale(a), p2 = scale(b);
        const offset = centre(i) - group / 2 + (stacked ? 0 : thick * slot);
        const rect = horizontal
          ? { x: Math.min(p1, p2), y: offset, width: Math.abs(p2 - p1), height: Math.max(0.5, thick - 1) }
          : { x: offset, y: Math.min(p1, p2), width: Math.max(0.5, thick - 1), height: Math.abs(p2 - p1) };
        node(svg, "rect", { ...rect, fill: color(j), class: "chart-mark" });
        if (f.labels) label(labelText(j, i), stacked ? (p1 + p2) / 2 : p2, offset + thick / 2, b < a, stacked);
      });
    });
    if (!combo) return;
  }
  // Lines, areas and scatter points: one point per category at its centre.
  const point = (i: number, v: number) => (horizontal ? `${scale(v)},${centre(i)}` : `${centre(i)},${scale(v)}`);
  if (f.kind === "area") {
    spans.forEach((vals, j) => {
      const top = vals.map(([, b], i) => point(i, b));
      const base = vals.map(([a], i) => point(i, stacked ? a : Math.max(lo, Math.min(0, hi)))).reverse();
      node(svg, "polygon", { points: [...top, ...base].join(" "), fill: color(j), "fill-opacity": stacked ? 1 : 0.7, class: "chart-mark" });
    });
  }
  spans.forEach((vals, j) => {
    if (combo && !isLine(j)) return;
    if (isLine(j)) node(svg, "polyline", { points: vals.map(([, b], i) => point(i, b)).join(" "), fill: "none", stroke: color(j), "stroke-width": 2, class: "chart-mark" });
    if (f.kind !== "area") vals.forEach(([, b], i) => node(svg, "circle", { cx: horizontal ? scale(b) : centre(i), cy: horizontal ? centre(i) : scale(b), r: 2.5, fill: color(j), class: "chart-mark" }));
    if (f.labels) vals.forEach(([, b], i) => label(labelText(j, i), scale(b), centre(i), false, false));
  });
}
