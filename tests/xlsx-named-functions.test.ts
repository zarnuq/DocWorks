import { describe, it, expect } from "vitest";
import { XlsxWorkbook } from "../src/infrastructure/xlsx/xlsx-workbook";
import { Workbook } from "../src/domain/xlsx/workbook-model";
import { Edits } from "../src/domain/xlsx/commands";
import { deleteName, lambdaText, listFunctions, listNames, setFunction } from "../src/domain/xlsx/names";
import { xlsx, sheetXml } from "./xlsx-fixtures";

// Names as Excel writes them: a LAMBDA with _xlpm. parameters, one calling another, one recursive.
const NAMES =
  `<definedNames>` +
  `<definedName name="ADDTAX" comment="Adds sales tax">_xlfn.LAMBDA(_xlpm.price,_xlpm.rate,_xlpm.price*(1+_xlpm.rate))</definedName>` +
  `<definedName name="DOUBLETAX">_xlfn.LAMBDA(_xlpm.x,ADDTAX(_xlpm.x,Data!$B$1)*2)</definedName>` +
  `<definedName name="FACT2">_xlfn.LAMBDA(_xlpm.n,IF(_xlpm.n&lt;=1,1,_xlpm.n*FACT2(_xlpm.n-1)))</definedName>` +
  `<definedName name="Rate">Data!$B$1</definedName>` +
  `</definedNames>`;
const ROWS =
  `<row r="1"><c r="A1"><v>100</v></c><c r="B1"><v>0.5</v></c></row>` +
  `<row r="2"><c r="A2"><f>ADDTAX(A1,B1)</f></c><c r="B2"><f>DOUBLETAX(A1)</f></c><c r="C2"><f>FACT2(5)</f></c></row>`;
const setup = async (extra = ROWS) => new Edits(Workbook.open(await XlsxWorkbook.load(await xlsx({ sheets: [sheetXml(extra)], names: ["Data"], workbookExtra: NAMES }))));
const val = (e: Edits, row: number, col: number) => e.wb.sheets[0].at(row, col)?.value;

describe("named functions: calculation", () => {
  it("a LAMBDA name is called with its arguments bound", async () => {
    const e = await setup();
    expect(val(e, 1, 0)).toBe(150);
  });

  it("named functions call each other, see cells, and recurse", async () => {
    const e = await setup();
    expect(val(e, 1, 1)).toBe(300);
    expect(val(e, 1, 2)).toBe(120);
  });

  it("recalculates when a cell the body reads changes (not only the arguments)", async () => {
    const e = await setup();
    e.input(e.wb.sheets[0], 0, 1, "1");
    expect(val(e, 1, 1)).toBe(400);
    expect(val(e, 1, 0)).toBe(200);
  });

  it("wrong argument count is #VALUE!; a LAMBDA left uncalled is #CALC!; runaway recursion is #NUM!", async () => {
    const e = await setup(`<row r="1"><c r="A1"><f>ADDTAX(1)</f></c><c r="B1"><f>ADDTAX</f></c><c r="C1"><f>_xlfn.LAMBDA(_xlpm.x,_xlpm.x)</f></c><c r="D1"><f>FACT2(1000)</f></c></row>`);
    expect(val(e, 0, 0)).toEqual({ error: "#VALUE!" });
    expect(val(e, 0, 1)).toEqual({ error: "#CALC!" });
    expect(val(e, 0, 2)).toEqual({ error: "#CALC!" });
    expect(val(e, 0, 3)).toEqual({ error: "#NUM!" });
  });

  it("a parameter hides a defined name of the same name only inside the function", async () => {
    const e = await setup(`<row r="1"><c r="B1"><v>0.25</v></c><c r="C1"><f>ADDTAX(10,Rate)</f></c></row>`);
    expect(val(e, 0, 2)).toBe(12.5);
  });

  it("ranges pass through as ranges", async () => {
    const names = NAMES.replace("</definedNames>", `<definedName name="TOTAL2">_xlfn.LAMBDA(_xlpm.r,SUM(_xlpm.r)*2)</definedName></definedNames>`);
    const e = new Edits(Workbook.open(await XlsxWorkbook.load(await xlsx({ sheets: [sheetXml(`<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>TOTAL2(A1:A2)</f></c></row><row r="2"><c r="A2"><v>2</v></c></row>`)], names: ["Data"], workbookExtra: names }))));
    expect(val(e, 0, 1)).toBe(6);
  });
});

describe("named functions: Data › Named functions", () => {
  it("lists functions apart from named ranges, without the file's prefixes", async () => {
    const e = await setup();
    expect(listNames(e.wb).map((n) => n.name)).toEqual(["Rate"]);
    expect(listFunctions(e.wb)).toEqual([
      { name: "ADDTAX", params: ["price", "rate"], body: "price*(1+rate)", description: "Adds sales tax" },
      { name: "DOUBLETAX", params: ["x"], body: "ADDTAX(x,Data!$B$1)*2", description: "" },
      { name: "FACT2", params: ["n"], body: "IF(n<=1,1,n*FACT2(n-1))", description: "" },
    ]);
  });

  it("stores a typed definition as Excel does", () => {
    expect(lambdaText(["a", "b"], "=a + B * 2")).toBe("_xlfn.LAMBDA(_xlpm.a,_xlpm.b,_xlpm.a + _xlpm.B * 2)");
    expect(lambdaText(["r"], "unique(r)")).toBe("_xlfn.LAMBDA(_xlpm.r,_xlfn.UNIQUE(_xlpm.r))");
  });

  it("adds a function that cells can use at once", async () => {
    const e = await setup();
    setFunction(e, null, { name: "HALF", params: ["v"], body: "v/2", description: "Half of it" });
    const s = e.wb.sheets[0];
    e.input(s, 4, 0, "=HALF(A1)");
    expect(val(e, 4, 0)).toBe(50);
    const xml = new XMLSerializer().serializeToString(e.wb.pkg.workbook);
    expect(xml).toContain(`<definedName name="HALF" comment="Half of it">_xlfn.LAMBDA(_xlpm.v,_xlpm.v/2)</definedName>`);
  });

  it("editing changes results; renaming rewrites the calls", async () => {
    const e = await setup();
    setFunction(e, "ADDTAX", { name: "WITHTAX", params: ["price", "rate"], body: "price*(1+rate)+1", description: "" });
    expect(e.wb.sheets[0].at(1, 0)?.formula).toBe("WITHTAX(A1,B1)");
    expect(val(e, 1, 0)).toBe(151);
    expect(listFunctions(e.wb).find((f) => f.name === "DOUBLETAX")!.body).toBe("WITHTAX(x,Data!$B$1)*2");
  });

  it("refuses built-in names, bad placeholders, repeats and empty definitions", async () => {
    const e = await setup();
    const f = (o: Partial<Parameters<typeof setFunction>[2]>) => () => setFunction(e, null, { name: "MYFN", params: [], body: "1", description: "", ...o });
    expect(f({ name: "SUM" })).toThrow(/already a function/);
    expect(f({ name: "ADDTAX" })).toThrow(/already exists/);
    expect(f({ params: ["A1"] })).toThrow(/placeholder/);
    expect(f({ params: ["two words"] })).toThrow(/placeholder/);
    expect(f({ params: ["x", "X"] })).toThrow(/twice/);
    expect(f({ body: " " })).toThrow(/definition/);
    expect(f({ body: "1+(" })).toThrow();
  });

  it("deleting a function leaves its calls unsupported (they keep their saved results)", async () => {
    const e = await setup();
    deleteName(e, "ADDTAX");
    expect(listFunctions(e.wb).map((f) => f.name)).toEqual(["DOUBLETAX", "FACT2"]);
  });
});

import { afterEach, vi } from "vitest";
import type { TFile as RealTFile, WorkspaceLeaf as RealLeaf } from "obsidian";
import { App, Menu, TFile, WorkspaceLeaf } from "./obsidian-stub";
import { XlsxView } from "../src/presentation/xlsx/xlsx-view";

vi.mock("../src/presentation/common/pdf", () => ({ canExportPdf: () => true, renderPdf: vi.fn(async () => null) }));
afterEach(() => document.body.empty());

describe("Data › Named functions dialog", () => {
  async function mount() {
    const app = new App();
    const bytes = await xlsx({ sheets: [sheetXml(ROWS)], names: ["Data"], workbookExtra: NAMES });
    app.vault.readBinary = async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
    app.vault.modifyBinary = async () => undefined;
    const view = new XlsxView(new WorkspaceLeaf(app) as unknown as RealLeaf);
    document.body.appendChild(view.containerEl);
    const file = new TFile("fns.xlsx") as unknown as RealTFile;
    view.file = file;
    await view.onLoadFile(file);
    Array.from(view.contentEl.querySelectorAll<HTMLButtonElement>(".dw-menu-btn")).find((x) => x.textContent === "Data")!.click();
    Menu.last!.items.find((i) => i.title === "Named functions…")!.run!();
    const modals = () => Array.from(document.querySelectorAll<HTMLElement>(".test-modal"));
    const button = (m: HTMLElement, text: string) => Array.from(m.querySelectorAll("button")).find((b) => b.textContent === text)!;
    const set = (m: HTMLElement, name: string, value: string) => {
      const input = Array.from(m.querySelectorAll(".setting-item")).find((s) => s.querySelector(".setting-item-name")?.textContent === name)!.querySelector<HTMLInputElement>("input, textarea")!;
      input.value = value;
      input.dispatchEvent(new Event("change"));
    };
    return { view, v: view as any, modals, button, set };
  }

  it("lists the functions with their placeholders", async () => {
    const { modals } = await mount();
    expect(Array.from(modals()[0].querySelectorAll(".setting-item-name")).map((n) => n.textContent)).toEqual(["ADDTAX(price, rate)", "DOUBLETAX(x)", "FACT2(n)"]);
  });

  it("Add new function: the form creates it and the list shows it", async () => {
    const { v, modals, button, set } = await mount();
    button(modals()[0], "Add new function").click();
    const form = modals()[1];
    set(form, "Function name", "SQUARE");
    set(form, "Argument placeholders", "x");
    set(form, "Formula definition", "=x*x");
    button(form, "Create").click();
    expect(listFunctions(v.wb).map((f: any) => f.name)).toContain("SQUARE");
    expect(Array.from(modals()[0].querySelectorAll(".setting-item-name")).map((n) => n.textContent)).toContain("SQUARE(x)");
    v.edits.input(v.sheet, 5, 0, "=SQUARE(4)");
    expect(v.sheet.at(5, 0).value).toBe(16);
  });

  it("a bad definition keeps the form open with nothing added", async () => {
    const { v, modals, button, set } = await mount();
    button(modals()[0], "Add new function").click();
    const form = modals()[1];
    set(form, "Function name", "SUM");
    set(form, "Formula definition", "=1");
    button(form, "Create").click();
    expect(form.isConnected).toBe(true);
    expect(listFunctions(v.wb)).toHaveLength(3);
  });
});
