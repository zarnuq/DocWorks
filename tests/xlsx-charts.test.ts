import { describe, expect, it } from "vitest";
import { XlsxWorkbook } from "../src/infrastructure/xlsx/xlsx-workbook";
import { Workbook } from "../src/domain/xlsx/workbook-model";
import { Edits } from "../src/domain/xlsx/commands";
import { chartModel, chartSource, deleteChart, formatChart, insertChart, moveChart, setChartSource, sheetCharts } from "../src/domain/xlsx/charts";
import { readChartFormat } from "../src/domain/pptx/chart-format";
import { sheetXml, S, xlsx } from "./xlsx-fixtures";

const C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const str = (ref: string, v: string) => `<c r="${ref}" t="inlineStr"><is><t>${v}</t></is></c>`;
const num = (ref: string, v: number) => `<c r="${ref}"><v>${v}</v></c>`;
const DATA =
  `<row r="1">${str("A1", "Region")}${str("B1", "Sales")}${str("C1", "Costs")}</row>` +
  `<row r="2">${str("A2", "North")}${num("B2", 10)}${num("C2", 4)}</row>` +
  `<row r="3">${str("A3", "South")}${num("B3", 20)}${num("C3", 6)}</row>` +
  `<row r="4">${str("A4", "East")}${num("B4", 15)}${num("C4", 9)}</row>`;

// Google's export: a pie with references but empty caches.
const PIE = `<c:chartSpace xmlns:c="${C}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:layout/><c:pieChart><c:varyColors val="1"/><c:ser><c:idx val="0"/><c:order val="0"/><c:dPt><c:idx val="0"/><c:spPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></c:spPr></c:dPt><c:cat><c:strRef><c:f>Data!$A$2:$A$4</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>Data!$B$2:$B$4</c:f><c:numCache/></c:numRef></c:val></c:ser><c:firstSliceAng val="0"/></c:pieChart></c:plotArea><c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/></c:chart></c:chartSpace>`;
const ANCHOR = `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:oneCellAnchor><xdr:from><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="5715000" cy="3533775"/><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="1" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${C}"><c:chart xmlns:c="${C}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>`;

async function setup(withChart = true) {
  const bytes = await xlsx({
    sheets: [sheetXml(DATA, { before: "", after: `<drawing r:id="rId1"/>` })],
    names: ["Data"],
    ...(withChart ? { drawing: ANCHOR, drawingRels: [["rId1", "../charts/chart1.xml"]] as [string, string][], parts: { "xl/charts/chart1.xml": PIE } } : {}),
  });
  return new Edits(Workbook.open(await XlsxWorkbook.load(bytes)));
}

describe("charts on sheets", () => {
  it("finds a chart's place and draws it from the sheet's cells when the file has no cached values", async () => {
    const e = await setup();
    const [chart] = sheetCharts(e.wb, e.wb.sheets[0]);
    expect(chart.box).toMatchObject({ x: e.wb.sheets[0].colX(4), y: e.wb.sheets[0].rowY(1), w: 600, h: 371 });
    const model = chartModel(e.wb, chart.path);
    expect(model.data).toEqual({ categories: ["North", "South", "East"], series: [{ name: "Series 1", values: [10, 20, 15] }] });
    expect(model.format).toMatchObject({ kind: "pie", legend: "r" });
    // Theme colours resolve against the workbook's theme (here none: null, so the palette applies).
    e.input(e.wb.sheets[0], 2, 1, "50");
    expect(chartModel(e.wb, chart.path).data.series[0].values).toEqual([10, 50, 15]);
    expect(chartSource(e.wb, chart.path)).toBe("Data!A2:B4");
  });

  it("inserts a chart from the selection: header row as series names, first column as categories", async () => {
    const e = await setup(false);
    const sheet = e.wb.sheets[0];
    const path = insertChart(e, sheet, { r1: 0, c1: 0, r2: 3, c2: 2 });
    const charts = sheetCharts(e.wb, e.wb.sheets[0]);
    expect(charts.map((c) => c.path)).toEqual([path]);
    expect(chartModel(e.wb, path).data).toEqual({ categories: ["North", "South", "East"], series: [{ name: "Sales", values: [10, 20, 15] }, { name: "Costs", values: [4, 6, 9] }] });
    expect(charts[0].box.x).toBeGreaterThan(sheet.colX(3));
    const saved = await XlsxWorkbook.load(await e.wb.pkg.save());
    const back = Workbook.open(saved);
    expect(sheetCharts(back, back.sheets[0])).toHaveLength(1);
    e.undo();
    expect(sheetCharts(e.wb, e.wb.sheets[0])).toHaveLength(0);
  });

  it("runs series along rows when the selection is wider than it is tall", async () => {
    const e = await setup(false);
    // A1:C2: one header row and one data row (Region / North, Sales / 10, Costs / 4).
    const path = insertChart(e, e.wb.sheets[0], { r1: 0, c1: 0, r2: 1, c2: 2 });
    expect(chartModel(e.wb, path).data).toEqual({ categories: ["Sales", "Costs"], series: [{ name: "North", values: [10, 4] }] });
  });

  it("inserts into a sheet without a drawing, adding the drawing part and its link", async () => {
    const bytes = await xlsx({ sheets: [sheetXml(DATA)], names: ["Data"] });
    const e = new Edits(Workbook.open(await XlsxWorkbook.load(bytes)));
    // The fixture links a drawing to sheet 1 regardless; drop it to test a bare sheet.
    const path = insertChart(e, e.wb.sheets[0], { r1: 0, c1: 0, r2: 3, c2: 1 }, (f) => { f.kind = "bar"; f.title = "Sales"; });
    const root = e.wb.sheets[0].doc.documentElement;
    expect(root.getElementsByTagNameNS(S, "drawing")).toHaveLength(1);
    expect(readChartFormat(e.wb.pkg.xml(path)!)).toMatchObject({ kind: "bar", title: "Sales" });
  });

  it("points a chart at another range, keeping series colours", async () => {
    const e = await setup(false);
    const path = insertChart(e, e.wb.sheets[0], { r1: 0, c1: 0, r2: 3, c2: 1 }, (f) => { f.colors = ["#FF0000"]; });
    setChartSource(e, path, "Data!A1:C4");
    expect(chartModel(e.wb, path).data.series.map((s) => s.name)).toEqual(["Sales", "Costs"]);
    expect(readChartFormat(e.wb.pkg.xml(path)!).colors[0]).toBe("#FF0000");
    setChartSource(e, path, "A1:B3");
    expect(chartModel(e.wb, path).data).toEqual({ categories: ["North", "South"], series: [{ name: "Sales", values: [10, 20] }] });
    expect(() => setChartSource(e, path, "nonsense")).toThrow();
  });

  it("formats and deletes a chart, each one undo step", async () => {
    const e = await setup();
    const [chart] = sheetCharts(e.wb, e.wb.sheets[0]);
    const f = chartModel(e.wb, chart.path).format!;
    formatChart(e, chart.path, { ...f, kind: "doughnut", title: "Share" });
    expect(readChartFormat(e.wb.pkg.xml(chart.path)!)).toMatchObject({ kind: "doughnut", title: "Share" });
    e.undo();
    expect(readChartFormat(e.wb.pkg.xml(chart.path)!).kind).toBe("pie");
    deleteChart(e, e.wb.sheets[0], sheetCharts(e.wb, e.wb.sheets[0])[0]);
    expect(sheetCharts(e.wb, e.wb.sheets[0])).toHaveLength(0);
    expect(e.wb.pkg.xml(chart.path)).toBeNull();
    e.undo();
    expect(sheetCharts(e.wb, e.wb.sheets[0])).toHaveLength(1);
    expect(e.wb.pkg.xml(chart.path)?.getElementsByTagNameNS(C, "pieChart")).toHaveLength(1);
  });
});

describe("moving and resizing charts", () => {
  it("moves and resizes a chart to a new box (its cell anchors follow), with undo", async () => {
    const e = await setup();
    const before = sheetCharts(e.wb, e.wb.sheets[0])[0];
    const box = { x: 40, y: 300, w: 400, h: 250 };
    moveChart(e, before.path, box);
    const after = sheetCharts(e.wb, e.wb.sheets[0])[0];
    expect(after.box.x).toBeCloseTo(box.x, 0);
    expect(after.box.y).toBeCloseTo(box.y, 0);
    expect(after.box.w).toBeCloseTo(box.w, 0);
    expect(after.box.h).toBeCloseTo(box.h, 0);
    e.undo();
    expect(sheetCharts(e.wb, e.wb.sheets[0])[0].box).toEqual(before.box);
  });
});
