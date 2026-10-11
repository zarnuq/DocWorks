import { describe, expect, it } from "vitest";
import { blankPptx } from "../src/infrastructure/blank";
import { PptxDocument } from "../src/infrastructure/pptx/pptx-document";
import { addChart, chartData, chartWorkbook, ChartData } from "../src/domain/pptx/chart";
import { ChartFormat, readChartFormat, setChartFormat } from "../src/domain/pptx/chart-format";
import { parseSlide } from "../src/domain/pptx/slide-model";
import { niceTicks, renderChart } from "../src/presentation/pptx/chart-renderer";

const C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const data: ChartData = { categories: ["North", "South", "East"], series: [{ name: "Sales", values: [3, 7, 5] }, { name: "Costs", values: [2, 4, 1] }] };

async function chart(type = "barChart", d = data) {
  const doc = await PptxDocument.load(await blankPptx());
  const slide = doc.slides()[0].path;
  const frame = addChart(doc, slide, d, type, await chartWorkbook(d));
  const path = parseSlide(doc, slide).shapes.find((s) => s.el === frame)!.chart!.path;
  return { doc, slide, path, xml: doc.xml(path)! };
}
const serialize = (x: Document) => new XMLSerializer().serializeToString(x);
const edit = (xml: Document, change: (f: ChartFormat) => void) => {
  const f = readChartFormat(xml);
  change(f);
  setChartFormat(xml, f);
  return readChartFormat(xml);
};
const local = (xml: Document, name: string) => Array.from(xml.getElementsByTagNameNS(C, name));

describe("chart format", () => {
  it("reads a new chart as a clustered column chart with a bottom legend and value gridlines", async () => {
    const { xml } = await chart();
    expect(readChartFormat(xml)).toEqual({
      kind: "column", stacking: "none", title: "", legend: "b", labels: false, colors: [null, null],
      category: { title: "", visible: true, gridlines: false, min: null, max: null, reversed: false, format: null },
      value: { title: "", visible: true, gridlines: true, min: null, max: null, reversed: false, format: null },
    });
  });

  it("writing the format it reads changes nothing", async () => {
    const { xml } = await chart();
    const before = serialize(xml);
    setChartFormat(xml, readChartFormat(xml));
    expect(serialize(xml)).toBe(before);
  });

  it("turns columns into stacked horizontal bars, moving the axes", async () => {
    const { xml } = await chart();
    const f = edit(xml, (f) => { f.kind = "bar"; f.stacking = "stacked"; });
    expect([f.kind, f.stacking]).toEqual(["bar", "stacked"]);
    expect(local(xml, "overlap")[0].getAttribute("val")).toBe("100");
    expect(local(xml, "catAx")[0].getElementsByTagNameNS(C, "axPos")[0].getAttribute("val")).toBe("l");
    expect(local(xml, "valAx")[0].getElementsByTagNameNS(C, "axPos")[0].getAttribute("val")).toBe("b");
    expect(edit(xml, (f) => { f.stacking = "percent"; }).stacking).toBe("percent");
    expect(edit(xml, (f) => { f.stacking = "none"; }).stacking).toBe("none");
    expect(local(xml, "overlap")).toHaveLength(0);
  });

  it("converts between chart types keeping the data, and drops or adds axes for pies", async () => {
    const { xml } = await chart();
    edit(xml, (f) => { f.kind = "pie"; });
    expect(readChartFormat(xml).kind).toBe("pie");
    expect(local(xml, "catAx")).toHaveLength(0);
    expect(local(xml, "axId")).toHaveLength(0);
    expect(chartData(xml)).toEqual(data);
    const f = edit(xml, (f) => { f.kind = "area"; f.stacking = "stacked"; });
    expect([f.kind, f.stacking]).toEqual(["area", "stacked"]);
    expect(local(xml, "catAx")).toHaveLength(1);
    expect(local(xml, "valAx")).toHaveLength(1);
    expect(local(xml, "areaChart")[0].getElementsByTagNameNS(C, "axId")).toHaveLength(2);
    expect(chartData(xml)).toEqual(data);
    expect(edit(xml, (f) => { f.kind = "line"; }).kind).toBe("line");
    expect(chartData(xml)).toEqual(data);
  });

  it("sets and removes the title, legend and data labels", async () => {
    const { xml } = await chart();
    expect(edit(xml, (f) => { f.title = "Revenue & costs"; }).title).toBe("Revenue & costs");
    expect(edit(xml, (f) => { f.title = "Revenue"; }).title).toBe("Revenue");
    expect(local(xml, "title")).toHaveLength(1);
    expect(edit(xml, (f) => { f.title = ""; }).title).toBe("");
    expect(local(xml, "autoTitleDeleted")[0].getAttribute("val")).toBe("1");
    expect(edit(xml, (f) => { f.legend = "r"; }).legend).toBe("r");
    expect(edit(xml, (f) => { f.legend = "none"; }).legend).toBe("none");
    expect(edit(xml, (f) => { f.legend = "t"; }).legend).toBe("t");
    expect(edit(xml, (f) => { f.labels = true; }).labels).toBe(true);
    expect(edit(xml, (f) => { f.labels = false; }).labels).toBe(false);
  });

  it("puts percentages on pie labels", async () => {
    const { xml } = await chart("pieChart");
    edit(xml, (f) => { f.labels = true; });
    expect(local(xml, "showPercent")[0].getAttribute("val")).toBe("1");
  });

  it("sets axis titles, gridlines, range, direction and number format", async () => {
    const { xml } = await chart();
    const f = edit(xml, (f) => {
      f.category.title = "Region";
      f.value = { ...f.value, title: "Dollars", gridlines: false, min: 0, max: 10, format: "0%" };
      f.category.reversed = true;
    });
    expect(f.category).toMatchObject({ title: "Region", reversed: true });
    expect(f.value).toMatchObject({ title: "Dollars", gridlines: false, min: 0, max: 10, format: "0%" });
    // Vertical axis titles are turned like Office's.
    expect(local(xml, "valAx")[0].getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/main", "bodyPr")[0].getAttribute("rot")).toBe("-5400000");
    const g = edit(xml, (f) => { f.value.min = null; f.value.visible = false; });
    expect(g.value).toMatchObject({ min: null, max: 10, visible: false });
  });

  it("colours series, lines and pie slices", async () => {
    const bar = await chart();
    expect(edit(bar.xml, (f) => { f.colors = ["#FF0000", null]; }).colors).toEqual(["#FF0000", null]);
    const line = await chart("lineChart");
    expect(edit(line.xml, (f) => { f.colors[1] = "#00AA00"; }).colors).toEqual([null, "#00AA00"]);
    expect(local(line.xml, "ser")[1].getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/main", "ln")).toHaveLength(1);
    const pie = await chart("pieChart");
    expect(readChartFormat(pie.xml).colors).toEqual([null, null, null]);
    expect(edit(pie.xml, (f) => { f.colors = [null, "#123456", "#ABCDEF"]; }).colors).toEqual([null, "#123456", "#ABCDEF"]);
    expect(edit(pie.xml, (f) => { f.colors[0] = "#000000"; }).colors).toEqual(["#000000", "#123456", "#ABCDEF"]);
    const points = local(pie.xml, "dPt").map((d) => d.getElementsByTagNameNS(C, "idx")[0].getAttribute("val"));
    expect(points).toEqual(["0", "1", "2"]);
  });

  it("reads Office-style formatting: rich title, right legend, series labels and theme colours", () => {
    const xml = new DOMParser().parseFromString(
      `<c:chartSpace xmlns:c="${C}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Q1 </a:t></a:r><a:r><a:t>results</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:barChart><c:barDir val="bar"/><c:grouping val="percentStacked"/><c:ser><c:idx val="0"/><c:order val="0"/><c:spPr><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></c:spPr><c:dLbls><c:showVal val="1"/></c:dLbls><c:val><c:numRef><c:numCache><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser><c:axId val="5"/><c:axId val="6"/></c:barChart><c:catAx><c:axId val="5"/><c:scaling><c:orientation val="maxMin"/></c:scaling><c:delete/></c:catAx><c:valAx><c:axId val="6"/><c:scaling/></c:valAx></c:plotArea><c:legend><c:legendPos val="r"/></c:legend></c:chart></c:chartSpace>`,
      "application/xml");
    const f = readChartFormat(xml, { scheme: { accent2: "#ED7D31" }, map: {} });
    expect(f).toMatchObject({ kind: "bar", stacking: "percent", title: "Q1 results", legend: "r", labels: true, colors: ["#ED7D31"] });
    expect(f.category).toMatchObject({ visible: false, reversed: true });
  });
});

describe("chart rendering", () => {
  const draw = (format: Partial<ChartFormat>, d = data, kind = "barChart") => {
    const host = document.createElement("div");
    const base = readChartFormat(new DOMParser().parseFromString(`<c:chartSpace xmlns:c="${C}"/>`, "application/xml"));
    renderChart(host, { path: "x", type: kind, data: d, format: { ...base, kind: "column", legend: "b", ...format, value: { ...base.value, gridlines: true, visible: true, ...format.value }, category: { ...base.category, visible: true, ...format.category } } }, { w: 400, h: 300 });
    return host;
  };
  const rects = (host: HTMLElement) => Array.from(host.querySelectorAll("rect.chart-mark")).map((r) => ["x", "y", "width", "height"].map((a) => Number(r.getAttribute(a))));
  const texts = (host: HTMLElement, cls = "") => Array.from(host.querySelectorAll(`text${cls}`)).map((t) => t.textContent);

  it("draws horizontal bars with lengths along x, first category at the bottom", () => {
    const r = rects(draw({ kind: "bar" }));
    expect(r).toHaveLength(6);
    // Sales: 3, 7, 5 -> widths in proportion, same heights.
    expect(r[1][2] / r[0][2]).toBeCloseTo(7 / 3, 1);
    expect(r[0][3]).toBeCloseTo(r[1][3]);
    expect(r[0][1]).toBeGreaterThan(r[1][1]);
    const reversed = rects(draw({ kind: "bar", category: { reversed: true } as ChartFormat["category"] }));
    expect(reversed[0][1]).toBeLessThan(reversed[1][1]);
  });

  it("stacks columns, and scales 100% stacks to the same height", () => {
    const r = rects(draw({ stacking: "stacked" }));
    // Costs sit on top of Sales in each column.
    expect(r[3][1] + r[3][3]).toBeCloseTo(r[0][1]);
    expect(r[3][0]).toBeCloseTo(r[0][0]);
    const p = rects(draw({ stacking: "percent" }));
    const height = (i: number) => p[i][3] + p[i + 3][3];
    expect(height(0)).toBeCloseTo(height(1));
    expect(texts(draw({ stacking: "percent" }))).toContain("100%");
  });

  it("draws the title, axis titles, data labels and the legend where asked", () => {
    const host = draw({ title: "Revenue", labels: true, legend: "r", value: { title: "Dollars" } as ChartFormat["value"], category: { title: "Region" } as ChartFormat["category"] });
    expect(texts(host, ".chart-title")).toEqual(["Revenue"]);
    expect(texts(host, ".chart-axis-title").sort()).toEqual(["Dollars", "Region"]);
    expect(texts(host, ".chart-label")).toContain("7");
    const legend = Array.from(host.querySelectorAll(".chart-legend text"));
    expect(legend.map((t) => t.textContent)).toEqual(["Sales", "Costs"]);
    const rightmostBar = Math.max(...rects(host).map((r) => r[0] + r[2]));
    expect(Number(legend[0].getAttribute("x"))).toBeGreaterThan(rightmostBar);
    expect(draw({ legend: "none" }).querySelector(".chart-legend")).toBeNull();
  });

  it("uses series colours, and percentages on pie labels", () => {
    const host = draw({ colors: ["#FF0000", null] });
    expect(host.querySelector("rect.chart-mark")!.getAttribute("fill")).toBe("#FF0000");
    const pie = draw({ kind: "pie", labels: true }, { categories: ["A", "B"], series: [{ name: "S", values: [1, 3] }] }, "pieChart");
    expect(texts(pie, ".chart-label")).toEqual(["25%", "75%"]);
  });

  it("picks round axis steps", () => {
    expect(niceTicks(0, 7)).toEqual([0, 2, 4, 6, 8]);
    expect(niceTicks(0, 1)).toEqual([0, 0.2, 0.4, 0.6, 0.8, 1]);
    expect(niceTicks(-3, 9)).toEqual([-5, -2.5, 0, 2.5, 5, 7.5, 10]);
  });
});
