import { describe, it, expect } from "vitest";
import JSZip from "jszip";
import { XlsxWorkbook } from "../src/infrastructure/xlsx/xlsx-workbook";
import { Workbook } from "../src/domain/xlsx/workbook-model";
import { Edits } from "../src/domain/xlsx/commands";
import { DEFAULT_PRINT, pagePlan, printSettings, setPrintSettings } from "../src/domain/xlsx/print";
import { xlsx } from "./xlsx-fixtures";

const setup = async () => new Edits(Workbook.open(await XlsxWorkbook.load(await xlsx())));

describe("print settings", () => {
  it("reads Excel's defaults when the sheet has none", async () => {
    const e = await setup();
    expect(printSettings(e.wb.sheets[0])).toEqual(DEFAULT_PRINT);
  });

  it("saves paper, orientation, scaling, margins and gridlines into the sheet, as Excel stores them", async () => {
    const e = await setup();
    const want = { paper: "a4" as const, landscape: true, fit: "width" as const, scale: 100, margins: { top: 0.5, bottom: 0.5, left: 0.25, right: 0.25 }, gridlines: true };
    setPrintSettings(e, e.wb.sheets[0], want);
    expect(printSettings(e.wb.sheets[0])).toEqual(want);
    const zip = await JSZip.loadAsync(await e.wb.pkg.save());
    const xml = await zip.file(e.wb.sheets[0].ref.path)!.async("text");
    expect(xml).toMatch(/<pageSetup[^>]*paperSize="9"[^>]*orientation="landscape"/);
    expect(xml).toMatch(/fitToHeight="0"/);
    expect(xml).toMatch(/<pageSetUpPr fitToPage="1"\/>/);
    expect(xml).toMatch(/<printOptions gridLines="1"\/>/);
    // Schema order: printOptions, pageMargins, pageSetup.
    expect(xml.indexOf("<printOptions")).toBeLessThan(xml.indexOf("<pageMargins"));
    expect(xml.indexOf("<pageMargins")).toBeLessThan(xml.indexOf("<pageSetup"));
  });

  it("splits a sheet into pages down, then across, at the chosen scale", () => {
    // 10 columns of 100 px and 60 rows of 20 px on Letter portrait with half-inch margins (720 x 960 px of room).
    const cols = Array(10).fill(100), rows = Array(60).fill(20);
    const normal = pagePlan(cols, rows, { ...DEFAULT_PRINT, margins: { top: 0.5, bottom: 0.5, left: 0.5, right: 0.5 } });
    expect(normal.scale).toBe(1);
    expect(normal.pages.map((p) => [p.c1, p.c2, p.r1, p.r2])).toEqual([[0, 6, 0, 47], [0, 6, 48, 59], [7, 9, 0, 47], [7, 9, 48, 59]]);
    const width = pagePlan(cols, rows, { ...DEFAULT_PRINT, fit: "width", margins: { top: 0.5, bottom: 0.5, left: 0.5, right: 0.5 } });
    expect(width.scale).toBeCloseTo(0.72);
    expect(width.pages.every((p) => p.c1 === 0 && p.c2 === 9)).toBe(true);
    const page = pagePlan(cols, rows, { ...DEFAULT_PRINT, fit: "page", margins: { top: 0.5, bottom: 0.5, left: 0.5, right: 0.5 } });
    expect(page.pages).toHaveLength(1);
    expect(page.scale).toBeCloseTo(0.72);
  });
});
