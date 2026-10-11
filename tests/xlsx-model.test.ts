import { describe, it, expect } from "vitest";
import { XlsxWorkbook } from "../src/infrastructure/xlsx/xlsx-workbook";
import { Workbook } from "../src/domain/xlsx/workbook-model";
import { xlsx, sheetXml, S } from "./xlsx-fixtures";

const open = async (bytes?: Uint8Array) => Workbook.open(await XlsxWorkbook.load(bytes ?? (await xlsx())));
const xmlOf = (el: Element) => new XMLSerializer().serializeToString(el).replace(/ xmlns(:\w+)?="[^"]*"/g, "");

describe("reading a workbook", () => {
  it("reads values of every cell type", async () => {
    const wb = await open();
    const s = wb.sheets[0];
    const v = (a: string) => s.cell(a)?.value;
    expect([v("A1"), v("B2"), v("B3"), v("B4"), v("A2"), v("C5"), v("B5")]).toEqual(["Name", "rich text", "inline", true, 10, { error: "#DIV/0!" }, 45292]);
    expect(s.cell("Z99")).toBeUndefined();
  });

  it("reads formulas, expanding shared formulas for every member", async () => {
    const s = (await open()).sheets[0];
    expect(s.cell("A5")).toMatchObject({ formula: "SUM(A2:A4)", value: 60 });
    expect(s.cell("C2")?.formula).toBe("A2*2");
    expect(s.cell("C3")).toMatchObject({ formula: "A3*2", value: 40 });
    expect(s.cell("C4")?.formula).toBe("A4*2");
    expect(s.cell("A6")).toMatchObject({ formula: 'Totals!A1&"!"', value: "x!" });
  });

  it("shows values through their number formats", async () => {
    const s = (await open()).sheets[0];
    expect(s.text("B5")).toBe("2024-01-01");
    expect(s.text("A2")).toBe("10");
    expect(s.text("B4")).toBe("TRUE");
  });

  it("reads sheet layout: widths, heights, hidden rows and columns, merges and frozen panes", async () => {
    const s = (await open()).sheets[0];
    expect(s.colWidth(0)).toBe(64);
    expect(s.colWidth(1)).toBe(140);
    expect(s.colWidth(3)).toBe(0);
    expect(s.rowHeight(0)).toBe(20);
    expect(s.rowHeight(3)).toBe(40);
    expect(s.rowHeight(5)).toBe(0);
    expect(s.merges).toEqual([{ r1: 5, c1: 1, r2: 5, c2: 2 }]);
    expect(s.frozen).toEqual({ rows: 1, cols: 0 });
  });

  it("finds positions and the row / column at a point", async () => {
    const s = (await open()).sheets[0];
    expect(s.colX(2)).toBe(64 + 140);
    expect(s.colAt(64 + 139)).toBe(1);
    expect(s.colAt(64 + 140)).toBe(2);
    expect(s.colAt(64 + 140 + 64)).toBe(4); // D is hidden
    expect(s.rowY(4)).toBe(20 * 3 + 40);
    expect(s.rowAt(20 * 3 + 39)).toBe(3);
    expect(s.colX(1000)).toBe(64 * 998 + 140); // B is wider, D hidden
  });

  it("knows the used range", async () => {
    const s = (await open()).sheets[0];
    expect(s.used()).toEqual({ r1: 0, c1: 0, r2: 5, c2: 2 });
  });

  it("uses the 1904 date system when the workbook asks for it", async () => {
    const bytes = await xlsx({ sheets: [sheetXml(`<row r="1"><c r="A1" s="2"><v>0</v></c></row>`)], workbookExtra: "" });
    const wb = await open(bytes);
    expect(wb.date1904).toBe(false);
    const pkg = await XlsxWorkbook.load(bytes);
    pkg.workbook.getElementsByTagNameNS(S, "workbookPr")[0].setAttribute("date1904", "1");
    expect(Workbook.open(pkg).sheets[0].text("A1")).toBe("1904-01-01");
  });
});

describe("writing cells", () => {
  it("adds cells in column order inside their row, and new rows in row order", async () => {
    const wb = await open();
    const s = wb.sheets[0];
    s.setCell(1, 3, { value: 5 });
    s.setCell(1, 0, { value: 7 });
    s.setCell(9, 0, { value: 1 });
    s.setCell(7, 1, { value: 2 });
    const rows = Array.from(s.doc.getElementsByTagNameNS(S, "row"));
    expect(rows.map((r) => r.getAttribute("r"))).toEqual(["1", "2", "3", "4", "5", "6", "8", "10"]);
    expect(Array.from(rows[1].children).map((c) => c.getAttribute("r"))).toEqual(["A2", "B2", "C2", "D2"]);
    expect(xmlOf(s.cell("D2")!.el)).toBe('<c r="D2"><v>5</v></c>');
    expect(xmlOf(s.cell("A2")!.el)).toBe('<c r="A2"><v>7</v></c>');
    expect(s.doc.getElementsByTagNameNS(S, "dimension")[0].getAttribute("ref")).toBe("A1:D10");
  });

  it("writes text to the shared string table, reusing an existing entry", async () => {
    const wb = await open();
    const s = wb.sheets[0];
    s.setCell(0, 5, { value: "Name" });
    s.setCell(1, 5, { value: "brand new" });
    expect(xmlOf(s.cell("F1")!.el)).toBe('<c r="F1" t="s"><v>0</v></c>');
    expect(xmlOf(s.cell("F2")!.el)).toBe('<c r="F2" t="s"><v>5</v></c>');
    const sst = wb.pkg.strings().documentElement;
    expect(sst.getAttribute("uniqueCount")).toBe("6");
    expect(sst.lastElementChild!.textContent).toBe("brand new");
  });

  it("keeps a cell's style, and writes booleans, errors and text with spaces", async () => {
    const s = (await open()).sheets[0];
    s.setCell(0, 0, { value: false });
    expect(xmlOf(s.cell("A1")!.el)).toBe('<c r="A1" s="1" t="b"><v>0</v></c>');
    s.setCell(0, 0, { value: { error: "#N/A" } });
    expect(xmlOf(s.cell("A1")!.el)).toBe('<c r="A1" s="1" t="e"><v>#N/A</v></c>');
    s.setCell(0, 0, { value: null });
    expect(xmlOf(s.cell("A1")!.el)).toBe('<c r="A1" s="1"/>');
  });

  it("removes an unstyled cleared cell", async () => {
    const s = (await open()).sheets[0];
    s.setCell(1, 0, { value: null });
    expect(s.cell("A2")).toBeUndefined();
    expect(s.doc.querySelector('c[r="A2"]')).toBeNull();
  });

  it("writes formulas with their result type", async () => {
    const s = (await open()).sheets[0];
    s.setCell(7, 0, { formula: 'A1&"x"', value: "Namex" });
    expect(xmlOf(s.cell("A8")!.el)).toBe('<c r="A8" t="str"><f>A1&amp;"x"</f><v>Namex</v></c>');
    s.setCell(7, 0, { formula: "1+1", value: 2 });
    expect(xmlOf(s.cell("A8")!.el)).toBe('<c r="A8"><f>1+1</f><v>2</v></c>');
  });

  it("splits a shared formula group into plain formulas before changing a member", async () => {
    const s = (await open()).sheets[0];
    s.setCell(2, 2, { value: 1 });
    expect(xmlOf(s.cell("C2")!.el)).toBe('<c r="C2"><f>A2*2</f><v>20</v></c>');
    expect(xmlOf(s.cell("C3")!.el)).toBe('<c r="C3"><v>1</v></c>');
    expect(xmlOf(s.cell("C4")!.el)).toBe('<c r="C4"><f>A4*2</f><v>60</v></c>');
  });

  it("updates a formula's cached value without touching the formula", async () => {
    const s = (await open()).sheets[0];
    s.setCached(s.cell("C3")!, 99);
    expect(xmlOf(s.cell("C3")!.el)).toBe('<c r="C3"><f t="shared" si="0"/><v>99</v></c>');
    s.setCached(s.cell("A6")!, 3);
    expect(xmlOf(s.cell("A6")!.el)).toBe('<c r="A6"><f>Totals!A1&amp;"!"</f><v>3</v></c>');
  });
});
