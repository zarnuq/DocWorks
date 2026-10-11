import { expect, it } from "vitest";
import JSZip from "jszip";
import { blankPptx } from "../src/infrastructure/blank";
import { PptxDocument } from "../src/infrastructure/pptx/pptx-document";
import { addChart, chartData, prepareChartEdit, chartWorkbook } from "../src/domain/pptx/chart";
import { parseSlide } from "../src/domain/pptx/slide-model";
import { copyItems, pasteItems } from "../src/domain/pptx/arrange";
const data = { categories: ["A", "B"], series: [{ name: "Sales", values: [3, 7] }] };
it("creates, renders, edits and independently copies a chart and workbook", async () => {
  const doc = await PptxDocument.load(await blankPptx());
  const slide = doc.slides()[0].path;
  const frame = addChart(doc, slide, data, "barChart", await chartWorkbook(data));
  const model = parseSlide(doc, slide).shapes.find(s => s.el === frame)!;
  expect(model.chart?.data).toEqual(data);
  const path = model.chart!.path;
  const edited = { categories: ["North", "South", "East"], series: [{ name: "Income", values: [4, 5, 6] }] };
  const update = await prepareChartEdit(doc, path, edited);
  doc.putXml(path, update.xml);
  doc.putBinary(update.workbookPath!, update.workbook!);
  expect(chartData(doc.xml(path)!)).toEqual(edited);
  const zip = await JSZip.loadAsync(await doc.media(update.workbookPath!));
  expect(await zip.file("xl/worksheets/sheet1.xml")!.async("text")).toContain("Income");
  const clip = await copyItems(doc, slide, [frame]);
  const [copy] = pasteItems(doc, slide, clip, { dx: 0, dy: 0 }).pasted;
  const cp = parseSlide(doc, slide).shapes.find(s => s.el === copy)!.chart!;
  expect(cp.path).not.toBe(path);
  expect(cp.data).toEqual(edited);
  const back = await PptxDocument.load(await doc.save());
  expect(chartData(back.xml(path)!)).toEqual(edited);
});
it("rejects invalid chart values without touching the package", async () => {
  const doc = await PptxDocument.load(await blankPptx());
  const before = doc.snapshot();
  await expect(chartWorkbook({ categories: ["A"], series: [{ name: "Bad", values: [NaN] }] })).rejects.toThrow();
  expect(doc.snapshot()).toEqual(before);
});

it("retains unrelated embedded workbook parts and chart number formatting",async()=>{
 const doc=await PptxDocument.load(await blankPptx());const slide=doc.slides()[0].path;
 const workbook=await JSZip.loadAsync(await chartWorkbook(data));workbook.file("custom/untouched.bin",new Uint8Array([9,8,7]));
 const frame=addChart(doc,slide,data,"lineChart",await workbook.generateAsync({type:"uint8array"}));const path=parseSlide(doc,slide).shapes.find(s=>s.el===frame)!.chart!.path;
 const cache=doc.xml(path)!.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/chart","formatCode")[0];cache.textContent="0.00%";
 const prepared=await prepareChartEdit(doc,path,{...data,series:[{name:"New",values:[5,10]}]});
 expect(prepared.xml.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/chart","formatCode")[0].textContent).toBe("0.00%");
 const edited=await JSZip.loadAsync(prepared.workbook!);expect(await edited.file("custom/untouched.bin")!.async("uint8array")).toEqual(new Uint8Array([9,8,7]));
});
