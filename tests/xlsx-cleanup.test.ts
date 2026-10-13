import { describe, it, expect } from "vitest";
import { XlsxWorkbook } from "../src/infrastructure/xlsx/xlsx-workbook";
import { Workbook } from "../src/domain/xlsx/workbook-model";
import { Edits } from "../src/domain/xlsx/commands";
import { detectSeparator, removeDuplicates, splitText } from "../src/domain/xlsx/cleanup";
import { xlsx, sheetXml } from "./xlsx-fixtures";

const str = (ref: string, text: string) => `<c r="${ref}" t="inlineStr"><is><t>${text}</t></is></c>`;
const num = (ref: string, n: number) => `<c r="${ref}"><v>${n}</v></c>`;
const setup = async (rows: string) => new Edits(Workbook.open(await XlsxWorkbook.load(await xlsx({ sheets: [sheetXml(rows)] }))));
const rect = (r1: number, c1: number, r2 = r1, c2 = c1) => ({ r1, c1, r2, c2 });
const column = (e: Edits, col: string, n: number) => Array.from({ length: n }, (_, i) => e.wb.sheets[0].cell(`${col}${i + 1}`)?.value ?? null);

describe("remove duplicates", () => {
  const rows =
    `<row r="1">${str("A1", "Name")}${str("B1", "Qty")}</row>` +
    `<row r="2">${str("A2", "pen")}${num("B2", 1)}</row>` +
    `<row r="3">${str("A3", "ink")}${num("B3", 2)}</row>` +
    `<row r="4">${str("A4", "pen")}${num("B4", 1)}</row>` +
    `<row r="5">${str("A5", "pen")}${num("B5", 3)}</row>` +
    `<row r="6">${str("A6", "Pen")}${num("B6", 1)}</row>` +
    `<row r="7">${str("A7", "after")}</row>`;

  it("keeps the first of each duplicate row and moves the rest up inside the range", async () => {
    const e = await setup(rows);
    const result = removeDuplicates(e, e.wb.sheets[0], rect(0, 0, 5, 1), { header: true, cols: [0, 1] });
    expect(result).toEqual({ removed: 1, kept: 4 });
    expect(column(e, "A", 7)).toEqual(["Name", "pen", "ink", "pen", "Pen", null, "after"]);
    expect(column(e, "B", 6)).toEqual(["Qty", 1, 2, 3, 1, null]);
  });

  it("compares only the chosen columns", async () => {
    const e = await setup(rows);
    expect(removeDuplicates(e, e.wb.sheets[0], rect(0, 0, 5, 1), { header: true, cols: [0] })).toEqual({ removed: 2, kept: 3 });
    expect(column(e, "A", 6)).toEqual(["Name", "pen", "ink", "Pen", null, null]);
    expect(column(e, "B", 4)).toEqual(["Qty", 1, 2, 1]);
  });

  it("treats the first row as data without a header, and undoes in one step", async () => {
    const e = await setup(`<row r="1">${num("A1", 5)}</row><row r="2">${num("A2", 5)}</row>`);
    expect(removeDuplicates(e, e.wb.sheets[0], rect(0, 0, 1, 0), { header: false, cols: [0] })).toEqual({ removed: 1, kept: 1 });
    expect(column(e, "A", 2)).toEqual([5, null]);
    e.undo();
    expect(column(e, "A", 2)).toEqual([5, 5]);
  });

  it("changes nothing when there are no duplicates", async () => {
    const e = await setup(`<row r="1">${num("A1", 1)}</row><row r="2">${num("A2", 2)}</row>`);
    expect(removeDuplicates(e, e.wb.sheets[0], rect(0, 0, 1, 0), { header: false, cols: [0] })).toEqual({ removed: 0, kept: 2 });
    expect(e.canUndo).toBe(false);
  });
});

describe("split text to columns", () => {
  it("detects the separator used in most cells", () => {
    expect(detectSeparator(["a,b", "c,d", "e f"])).toBe(",");
    expect(detectSeparator(["a;b", "x"])).toBe(";");
    expect(detectSeparator(["one two"])).toBe(" ");
    expect(detectSeparator(["plain"])).toBeNull();
  });

  it("splits each cell into the columns to its right, parsing parts like typed input", async () => {
    const e = await setup(`<row r="1">${str("A1", "Ann,42,2026-01-05")}</row><row r="2">${str("A2", "Bo,7")}${str("C2", "old")}</row><row r="3">${num("A3", 9)}</row>`);
    splitText(e, e.wb.sheets[0], rect(0, 0, 2, 0), ",");
    const s = e.wb.sheets[0];
    expect([s.cell("A1")!.value, s.cell("B1")!.value, typeof s.cell("C1")!.value]).toEqual(["Ann", 42, "number"]);
    expect([s.cell("A2")!.value, s.cell("B2")!.value, s.cell("C2")!.value]).toEqual(["Bo", 7, "old"]);
    expect(s.cell("A3")!.value).toBe(9);
    e.undo();
    expect(e.wb.sheets[0].cell("A1")!.value).toBe("Ann,42,2026-01-05");
  });

  it("refuses more than one column", async () => {
    const e = await setup(`<row r="1">${str("A1", "a,b")}</row>`);
    expect(() => splitText(e, e.wb.sheets[0], rect(0, 0, 0, 1), ",")).toThrow(/one column/);
  });
});
