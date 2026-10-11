import type { ChartModel } from "../../domain/pptx/chart";
const NS = "http://www.w3.org/2000/svg";
const COLORS = ["#4285f4", "#ea4335", "#fbbc04", "#34a853", "#a142f4", "#24c1e0"];
function node(parent: Element, tag: string, attrs: Record<string, string | number> = {}, text?: string): SVGElement {
  const el = document.createElementNS(NS, tag);
  for (const [key, value] of Object.entries(attrs))
    el.setAttribute(key, String(value));
  if (text !== undefined)
    el.textContent = text;
  parent.append(el);
  return el;
}
/** A lightweight preview of common chart types, driven by their saved caches. */
export function renderChart(host: HTMLElement, chart: ChartModel): void {
  const svg = node(host, "svg", { viewBox: "0 0 640 400", width: "100%", height: "100%", role: "img", "aria-label": "Chart" });
  const { categories, series } = chart.data;
  if (!series.length || !categories.length) {
    node(svg, "text", { x: 20, y: 40 }, "Chart has no cached data");
    return;
  }
  const pie = chart.type === "pieChart" || chart.type === "doughnutChart";
  if (pie) {
    const values = series[0].values.map(v => Math.max(0, v));
    const total = values.reduce((a, b) => a + b, 0);
    let angle = -Math.PI / 2;
    values.forEach((value, i) => {
      if (!total || !value)
        return;
      const next = angle + value / total * Math.PI * 2;
      const x = 320 + 145 * Math.cos(angle), y = 175 + 145 * Math.sin(angle);
      const xx = 320 + 145 * Math.cos(next), yy = 175 + 145 * Math.sin(next);
      if (value === total)
        node(svg, "circle", { cx: 320, cy: 175, r: 145, fill: COLORS[i % COLORS.length] });
      else
        node(svg, "path", { d: `M320 175 L${x} ${y} A145 145 0 ${next - angle > Math.PI ? 1 : 0} 1 ${xx} ${yy} Z`, fill: COLORS[i % COLORS.length] });
      angle = next;
    });
    if (chart.type === "doughnutChart")
      node(svg, "circle", { cx: 320, cy: 175, r: 72, fill: "white" });
  }
  else {
    const values = series.flatMap(s => s.values).filter(Number.isFinite);
    const lo = Math.min(0, ...values), hi = Math.max(1, ...values);
    const y = (v: number) => 315 - (v - lo) / (hi - lo) * 275;
    const step = 540 / categories.length;
    for (let i = 0; i <= 4; i++) {
      const value = lo + (hi - lo) * i / 4;
      node(svg, "line", { x1: 60, x2: 600, y1: y(value), y2: y(value), stroke: "#ddd" });
      node(svg, "text", { x: 53, y: y(value) + 4, "text-anchor": "end", "font-size": 12, fill: "#555" }, String(Math.round(value * 100) / 100));
    }
    categories.forEach((name, i) => node(svg, "text", { x: 60 + step * (i + .5), y: 335, "text-anchor": "middle", "font-size": 12, fill: "#555" }, name));
    series.forEach((s, j) => {
      const color = COLORS[j % COLORS.length];
      if (chart.type === "lineChart" || chart.type === "areaChart" || chart.type === "scatterChart") {
        const pts = s.values.map((v, i) => `${60 + step * (i + .5)},${y(v)}`).join(" ");
        if (chart.type === "areaChart")
          node(svg, "polygon", { points: `${60 + step * .5},${y(0)} ${pts} ${60 + step * (s.values.length - .5)},${y(0)}`, fill: color, "fill-opacity": .25 });
        if (chart.type !== "scatterChart")
          node(svg, "polyline", { points: pts, fill: "none", stroke: color, "stroke-width": 3 });
        s.values.forEach((v, i) => node(svg, "circle", { cx: 60 + step * (i + .5), cy: y(v), r: 4, fill: color }));
      }
      else
        s.values.forEach((v, i) => { const w = step * .8 / series.length; node(svg, "rect", { x: 60 + step * i + step * .1 + w * j, y: Math.min(y(0), y(v)), width: w - 2, height: Math.abs(y(v) - y(0)), fill: color }); });
    });
  }
  const labels = pie ? categories : series.map(s => s.name);
  labels.forEach((label, i) => { const x = 20 + (i % 4) * 155, y = 365 + Math.floor(i / 4) * 18; node(svg, "rect", { x, y: y - 10, width: 10, height: 10, fill: COLORS[i % COLORS.length] }); node(svg, "text", { x: x + 16, y, "font-size": 12, fill: "#555" }, label); });
}
