/* eslint-disable @typescript-eslint/no-explicit-any */
vi.mock("../src/presentation/common/pdf", () => ({ canExportPdf: () => true, renderPdf: vi.fn(async () => null) }));
import { afterEach, describe, expect, it, vi } from "vitest";
import type { TFile as RealTFile, WorkspaceLeaf as RealLeaf } from "obsidian";
import { App, Menu, TFile, WorkspaceLeaf } from "./obsidian-stub";
import { XlsxView } from "../src/presentation/xlsx/xlsx-view";
import { XlsxWorkbook } from "../src/infrastructure/xlsx/xlsx-workbook";
import { Workbook } from "../src/domain/xlsx/workbook-model";
import { Edits } from "../src/domain/xlsx/commands";
import { setFilter } from "../src/domain/xlsx/sheet-edit";
import { addSlicer, deleteSlicer, makeTable, setSlicerSelection, sheetSlicers, sheetTables, slicerItems, tableFilter } from "../src/domain/xlsx/slicers";
import { sheetXml, xlsx } from "./xlsx-fixtures";

afterEach(() => {
  document.body.empty();
  vi.restoreAllMocks();
});

const str = (r: string, t: string) => `<c r="${r}" t="inlineStr"><is><t>${t}</t></is></c>`;
const DATA = [["Region", "Sales"], ["East", 1], ["West", 2], ["East", 3], ["North", 4]]
  .map((row, i) => `<row r="${i + 1}">${row.map((v, j) => (typeof v === "string" ? str(`${"AB"[j]}${i + 1}`, v) : `<c r="${"AB"[j]}${i + 1}"><v>${v}</v></c>`)).join("")}</row>`).join("");
const RANGE = { r1: 0, c1: 0, r2: 4, c2: 1 };
const load = async (bytes: Uint8Array) => new Edits(Workbook.open(await XlsxWorkbook.load(bytes)));
const setup = async (data = DATA) => load(await xlsx({ sheets: [sheetXml(data)], names: ["Data"] }));
const hidden = (e: Edits) => [1, 2, 3, 4].filter((r) => e.wb.sheets[0].rowEl(r, false)?.getAttribute("hidden") === "1");
const xml = (e: Edits, path: string) => new XMLSerializer().serializeToString(e.wb.pkg.xml(path)!);
const withSlicer = async () => {
  const e = await setup();
  const table = makeTable(e, e.wb.sheets[0], RANGE);
  addSlicer(e, e.wb.sheets[0], sheetTables(e.wb, e.wb.sheets[0])[0], 0, { x: 300, y: 0 });
  return { e, table, slicer: () => sheetSlicers(e.wb, e.wb.sheets[0])[0] };
};

describe("makeTable", () => {
  it("writes a table part with the headings as its columns", async () => {
    const e = await setup();
    const t = makeTable(e, e.wb.sheets[0], RANGE);
    expect(t).toMatchObject({ id: 1, name: "Table1", columns: [{ id: 1, name: "Region" }, { id: 2, name: "Sales" }] });
    expect(sheetTables(e.wb, e.wb.sheets[0])).toMatchObject([{ name: "Table1", ref: RANGE }]);
    expect(xml(e, t.path)).toContain('<autoFilter ref="A1:B5"/>');
    expect(xml(e, "[Content_Types].xml")).toContain(`PartName="/${t.path}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"`);
    expect(new XMLSerializer().serializeToString(e.wb.sheets[0].doc)).toMatch(/<tableParts count="1"><tablePart r:id="rId\d+"\/><\/tableParts><\/worksheet>/);
  });

  it("needs a heading on every column, all different, and doesn't overlap another table", async () => {
    const e = await setup();
    expect(() => makeTable(e, e.wb.sheets[0], { ...RANGE, c2: 2 })).toThrow(/heading/);
    const dup = await setup(DATA.replace(">Sales<", ">Region<"));
    expect(() => makeTable(dup, dup.wb.sheets[0], RANGE)).toThrow(/Two columns/);
    makeTable(e, e.wb.sheets[0], RANGE);
    expect(() => makeTable(e, e.wb.sheets[0], RANGE)).toThrow(/overlaps Table1/);
  });

  it("a sheet filter on the range goes (a table has its own)", async () => {
    const e = await setup();
    setFilter(e, e.wb.sheets[0], RANGE);
    makeTable(e, e.wb.sheets[0], RANGE);
    expect(e.wb.sheets[0].filter).toBeNull();
  });
});

describe("table slicers", () => {
  it("adding one writes Excel's five places", async () => {
    const { e, slicer } = await withSlicer();
    const book = xml(e, e.wb.pkg.workbookPath);
    expect(book).toContain('<definedName name="Slicer_Region">#N/A</definedName>');
    expect(book).toMatch(/<ext uri="\{46BE6895-7355-4a93-B00E-2C351335B9C9\}" xmlns:x15="[^"]+"><x15:slicerCaches xmlns:x14="[^"]+"><x14:slicerCache r:id="rId\d+"\/><\/x15:slicerCaches><\/ext><\/extLst><\/workbook>/);
    const cacheRel = e.wb.pkg.rels(e.wb.pkg.workbookPath).find((r) => r.type === "slicerCache")!;
    expect(cacheRel.uri).toBe("http://schemas.microsoft.com/office/2007/relationships/slicerCache");
    expect(xml(e, cacheRel.target)).toMatch(/name="Slicer_Region" sourceName="Region".*<x15:tableSlicerCache tableId="1" column="1"\/>/);
    const sheetXmlText = new XMLSerializer().serializeToString(e.wb.sheets[0].doc);
    expect(sheetXmlText).toMatch(/<ext uri="\{3A4CF648-6AED-40f4-86FF-DC5316D8AED3\}"[^>]*><x14:slicerList[^>]*><x14:slicer r:id="rId\d+"\/><\/x14:slicerList><\/ext>/);
    const slicerRel = e.wb.pkg.rels(e.wb.sheets[0].ref.path).find((r) => r.type === "slicer")!;
    expect(xml(e, slicerRel.target)).toContain('<slicer name="Region" cache="Slicer_Region" caption="Region" rowHeight="241300"/>');
    const types = xml(e, "[Content_Types].xml");
    expect(types).toContain("application/vnd.ms-excel.slicerCache+xml");
    expect(types).toContain("application/vnd.ms-excel.slicer+xml");
    expect(slicer()).toMatchObject({ name: "Region", caption: "Region", column: 0, box: { x: 300, y: 0, w: 180, h: 240 } });
  });

  it("the drawing holds the slicer frame in an sle15 choice with a text fallback", async () => {
    const { e, slicer } = await withSlicer();
    const drawing = xml(e, slicer().drawing!);
    expect(drawing).toMatch(/<mc:Choice xmlns:sle15="[^"]+" Requires="sle15"><xdr:graphicFrame macro="">.*<a:graphicData uri="http:\/\/schemas.microsoft.com\/office\/drawing\/2010\/slicer"><sle:slicer xmlns:sle="[^"]+" name="Region"\/>/);
    expect(drawing).toContain("<mc:Fallback><xdr:sp");
  });

  it("survives a save and reopen", async () => {
    const { e } = await withSlicer();
    const again = await load(await e.wb.pkg.save());
    expect(sheetSlicers(again.wb, again.wb.sheets[0]).map((s) => s.name)).toEqual(["Region"]);
  });

  it("lists the column's values sorted, all chosen while it's unfiltered", async () => {
    const { e, slicer } = await withSlicer();
    expect(slicerItems(e.wb, e.wb.sheets[0], slicer())).toEqual([{ value: "East", selected: true }, { value: "North", selected: true }, { value: "West", selected: true }]);
  });

  it("choosing values filters the table (the choice is its autoFilter) and hides the other rows", async () => {
    const { e, slicer, table } = await withSlicer();
    setSlicerSelection(e, e.wb.sheets[0], slicer(), new Set(["East"]));
    expect(hidden(e)).toEqual([2, 4]);
    expect(tableFilter(e.wb, table, 0)).toEqual(new Set(["East"]));
    expect(xml(e, table.path)).toContain('<autoFilter ref="A1:B5"><filterColumn colId="0"><filters><filter val="East"/></filters></filterColumn></autoFilter>');
    expect(slicerItems(e.wb, e.wb.sheets[0], slicer()).filter((i) => i.selected).map((i) => i.value)).toEqual(["East"]);
    setSlicerSelection(e, e.wb.sheets[0], slicer(), null);
    expect(hidden(e)).toEqual([]);
    e.undo();
    expect(hidden(e)).toEqual([2, 4]);
  });

  it("two slicers share the sheet's slicer part and filter together", async () => {
    const { e, slicer } = await withSlicer();
    addSlicer(e, e.wb.sheets[0], sheetTables(e.wb, e.wb.sheets[0])[0], 1, { x: 500, y: 0 });
    const all = sheetSlicers(e.wb, e.wb.sheets[0]);
    expect(all.map((s) => s.name)).toEqual(["Region", "Sales"]);
    expect(e.wb.pkg.rels(e.wb.sheets[0].ref.path).filter((r) => r.type === "slicer")).toHaveLength(1);
    setSlicerSelection(e, e.wb.sheets[0], slicer(), new Set(["East"]));
    setSlicerSelection(e, e.wb.sheets[0], sheetSlicers(e.wb, e.wb.sheets[0])[1], new Set(["3"]));
    expect(hidden(e)).toEqual([1, 2, 4]);
  });

  it("names don't repeat", async () => {
    const { e } = await withSlicer();
    addSlicer(e, e.wb.sheets[0], sheetTables(e.wb, e.wb.sheets[0])[0], 0, { x: 500, y: 0 });
    expect(sheetSlicers(e.wb, e.wb.sheets[0]).map((s) => s.name)).toEqual(["Region", "Region 1"]);
    expect(xml(e, e.wb.pkg.workbookPath)).toContain('<definedName name="Slicer_Region1">#N/A</definedName>');
  });

  it("deleting the last one removes every part and entry, and its filter", async () => {
    const { e, slicer } = await withSlicer();
    const before = { workbook: xml(e, e.wb.pkg.workbookPath) };
    setSlicerSelection(e, e.wb.sheets[0], slicer(), new Set(["West"]));
    deleteSlicer(e, e.wb.sheets[0], slicer());
    expect(sheetSlicers(e.wb, e.wb.sheets[0])).toEqual([]);
    expect(hidden(e)).toEqual([]);
    const book = xml(e, e.wb.pkg.workbookPath);
    expect(book).not.toContain("Slicer_Region");
    expect(book).not.toContain("extLst");
    expect(new XMLSerializer().serializeToString(e.wb.sheets[0].doc)).not.toContain("slicerList");
    expect(e.wb.pkg.xmlPaths().some((p) => p.includes("slicer"))).toBe(false);
    expect(xml(e, slicer.length >= 0 ? sheetTables(e.wb, e.wb.sheets[0])[0].path : "")).not.toContain("filterColumn");
    expect(before.workbook).toContain("Slicer_Region");
  });
});

describe("Data › Add a slicer", () => {
  async function mount() {
    const app = new App();
    const bytes = await xlsx({ sheets: [sheetXml(DATA)], names: ["Data"] });
    app.vault.readBinary = async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
    app.vault.modifyBinary = async () => undefined;
    const view = new XlsxView(new WorkspaceLeaf(app) as unknown as RealLeaf);
    document.body.appendChild(view.containerEl);
    const file = new TFile("slicers.xlsx") as unknown as RealTFile;
    view.file = file;
    await view.onLoadFile(file);
    const v = view as any;
    const flush = () => new Promise((r) => setTimeout(r, 0));
    const modal = () => Array.from(document.querySelectorAll<HTMLElement>(".test-modal")).at(-1);
    const addSlicerMenu = () => {
      Array.from(view.contentEl.querySelectorAll<HTMLButtonElement>(".dw-menu-btn")).find((x) => x.textContent === "Data")!.click();
      Menu.last!.items.find((i) => i.title === "Add a slicer")!.run!();
    };
    const box = () => view.contentEl.querySelector<HTMLElement>(".xlsx-slicer");
    const click = (el: Element, init: MouseEventInit = {}) => el.dispatchEvent(new MouseEvent("click", { bubbles: true, ...init }));
    return { view, v, flush, modal, addSlicerMenu, box, click };
  }

  it("outside a table it offers to make one, then draws the slicer for the selected column", async () => {
    const { v, flush, modal, addSlicerMenu, box } = await mount();
    v.grid.select({ row: 2, col: 0 });
    addSlicerMenu();
    expect(modal()!.textContent).toContain("Make A1:B5 a table");
    Array.from(modal()!.querySelectorAll("button")).find((b) => b.textContent === "Make table")!.click();
    await flush();
    await flush();
    v.grid.render?.();
    expect(sheetSlicers(v.wb, v.sheet).map((s: any) => s.caption)).toEqual(["Region"]);
    expect(Array.from(box()!.querySelectorAll(".xlsx-slicer-item")).map((b) => b.textContent)).toEqual(["East", "North", "West"]);
  });

  it("a click shows one value; Ctrl+click adds another; ✕ clears", async () => {
    const { v, flush, modal, addSlicerMenu, box, click } = await mount();
    v.grid.select({ row: 1, col: 0 });
    addSlicerMenu();
    Array.from(modal()!.querySelectorAll("button")).find((b) => b.textContent === "Make table")!.click();
    await flush();
    await flush();
    const item = (text: string) => Array.from(box()!.querySelectorAll<HTMLElement>(".xlsx-slicer-item")).find((b) => b.textContent === text)!;
    click(item("West"));
    expect(Array.from(box()!.querySelectorAll(".xlsx-slicer-item.is-selected")).map((b) => b.textContent)).toEqual(["West"]);
    expect(v.sheet.rowEl(1, false)?.getAttribute("hidden")).toBe("1");
    click(item("North"), { ctrlKey: true });
    expect(Array.from(box()!.querySelectorAll(".xlsx-slicer-item.is-selected")).map((b) => b.textContent)).toEqual(["North", "West"]);
    click(box()!.querySelector(".xlsx-slicer-clear")!);
    expect(box()!.querySelectorAll(".xlsx-slicer-item.is-selected")).toHaveLength(3);
    expect(v.sheet.rowEl(1, false)?.getAttribute("hidden") ?? null).toBeNull();
  });

  it("Make table can be declined: nothing changes", async () => {
    const { v, flush, modal, addSlicerMenu } = await mount();
    addSlicerMenu();
    Array.from(modal()!.querySelectorAll("button")).find((b) => b.textContent === "Cancel")!.click();
    await flush();
    expect(sheetTables(v.wb, v.sheet)).toEqual([]);
  });
});

describe("regionAround", () => {
  it("is the block of filled cells around a cell, as Excel's current region", async () => {
    const { regionAround } = await import("../src/domain/xlsx/slicers");
    const e = await setup(DATA + `<row r="8"><c r="D8"><v>9</v></c></row>`);
    expect(regionAround(e.wb.sheets[0], 2, 0)).toEqual(RANGE);
    expect(regionAround(e.wb.sheets[0], 7, 3)).toEqual({ r1: 7, c1: 3, r2: 7, c2: 3 });
  });
});
