/** Opt-in measurements: npm run benchmark -- [workbook.xlsx]. No timing assertions on CI. */
import { it, expect } from "vitest";
import { readFileSync, writeFileSync } from "node:fs";
import { cpus, totalmem } from "node:os";
import { XlsxWorkbook } from "../src/infrastructure/xlsx/xlsx-workbook";
import { Workbook } from "../src/domain/xlsx/workbook-model";
import { Edits } from "../src/domain/xlsx/commands";
import { Grid, GridHost } from "../src/presentation/xlsx/grid";
import { xlsx, sheetXml } from "./xlsx-fixtures";

it.skipIf(!process.env.DOC_WORKS_BENCH)("measures large workbook load, edit, render and save", async () => {
  const path = process.env.DOC_WORKS_BENCH_FILE;
  const synthetic = async () => {
    const rows = Array.from({length: 20000}, (_, i) => {
      const n = i + 1;
      return `<row r="${n}"><c r="A${n}"><v>${n}</v></c><c r="B${n}"><f>A${n}*2</f><v>${n*2}</v></c><c r="C${n}"><v>3</v></c><c r="D${n}"><v>4</v></c><c r="E${n}"><v>5</v></c></row>`;
    }).join("");
    return xlsx({sheets: [sheetXml(rows)]});
  };
  const bytes = path ? new Uint8Array(readFileSync(path)) : await synthetic();
  if (process.env.DOC_WORKS_BENCH_FIXTURE) writeFileSync(process.env.DOC_WORKS_BENCH_FIXTURE, bytes);
  const times: Record<string, number> = {};
  const timed = async <T>(name: string, fn: () => T | Promise<T>): Promise<T> => {
    const start = performance.now(); const result = await fn(); times[name] = performance.now()-start; return result;
  };
  const memoryBefore = process.memoryUsage().rss;
  const pkg = await timed("packageLoadMs", () => XlsxWorkbook.load(bytes));
  const wb = await timed("modelOpenMs", () => Workbook.open(pkg));
  const edits = await timed("calculationIndexMs", () => new Edits(wb));
  const sheet = wb.sheets[0];
  const host: GridHost = {
    sheet: () => sheet, isStale: c => edits.calc.isStale(c), isCircular: c => edits.calc.isCircular(c),
    commit: (r,c,text) => { edits.input(sheet,r,c,text); return true; },
    selectionChanged: () => {}, editing: () => {}, clear: () => {}, fill: () => {}, resize: () => {},
    autofit: () => {}, contextMenu: () => {}, filterMenu: () => {},
  };
  const grid = new Grid(document.body.createDiv(), host);
  await timed("firstGridRenderMs", () => grid.render());
  const samples = (fn: (i: number) => void) => Array.from({length: 30}, (_, i) => {const t = performance.now(); fn(i); return performance.now()-t;}).sort((a,b) => a-b);
  const render = samples(i => {grid.el.querySelector<HTMLElement>(".xlsx-scroll")!.scrollTop = i*500; grid.render();});
  const edit = samples(i => edits.input(sheet,0,0,String(i+1)));
  times.renderMedianMs = render[15]; times.renderP95Ms = render[28];
  times.editMedianMs = edit[15]; times.editP95Ms = edit[28];
  const saved = await timed("saveMs", () => pkg.save());
  expect(saved.length).toBeGreaterThan(0);
  if (!path) expect(sheet.cell("B1")?.value).toBe(60);
  const result = {date: new Date().toISOString(), runtime: process.version, cpu: cpus()[0]?.model,
    systemMemoryGiB: totalmem()/2**30, environment: "Node/jsdom; canvas paint mocked; no browser frame-time claim",
    source: path ? "provided workbook" : "synthetic 100000 cells / 20000 formulas", bytes: bytes.length,
    cells: wb.sheets.reduce((n,s) => n+s.cells.size,0), sheets: wb.sheets.length,
    rssBeforeMiB: memoryBefore/2**20, rssAfterMiB: process.memoryUsage().rss/2**20,
    domElements: grid.el.querySelectorAll("*").length, ...times};
  console.log(JSON.stringify(result,null,2));
  if (process.env.DOC_WORKS_BENCH_OUT) writeFileSync(process.env.DOC_WORKS_BENCH_OUT, JSON.stringify(result,null,2)+"\n");
  grid.destroy(); document.body.empty();
}, 180000);
