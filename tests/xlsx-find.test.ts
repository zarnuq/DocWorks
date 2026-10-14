import { describe, it, expect } from "vitest";
import { XlsxWorkbook } from "../src/infrastructure/xlsx/xlsx-workbook";
import { Workbook } from "../src/domain/xlsx/workbook-model";
import { Edits } from "../src/domain/xlsx/commands";
import { findCells, replaceCells } from "../src/domain/xlsx/find";
import { xlsx } from "./xlsx-fixtures";

const setup = async () => new Edits(Workbook.open(await XlsxWorkbook.load(await xlsx())));
const refs = (ms: { row: number; col: number; sheet: { name: string } }[]) => ms.map((m) => `${m.sheet.name}!${m.row},${m.col}`);

describe("find", () => {
  it("finds shown text in reading order across sheets, ignoring case by default", async () => {
    const e = await setup();
    expect(refs(findCells(e.wb.sheets, "x"))).toEqual(["Data!1,1", "Data!5,0", "Totals!0,0"]);
    expect(refs(findCells(e.wb.sheets, "NAME"))).toEqual(["Data!0,0"]);
    expect(findCells(e.wb.sheets, "NAME", { matchCase: true })).toHaveLength(0);
    expect(refs(findCells(e.wb.sheets, "2024"))).toEqual(["Data!4,1"]);
  });

  it("matches whole cells and formula text when asked", async () => {
    const e = await setup();
    expect(refs(findCells(e.wb.sheets, "10", { entire: true }))).toEqual(["Data!1,0"]);
    expect(refs(findCells(e.wb.sheets, "sum(", { formulas: true }))).toEqual(["Data!4,0", "Totals!0,1"]);
    expect(findCells(e.wb.sheets, "sum(")).toHaveLength(0);
  });
});

describe("replace", () => {
  it("replaces text values and re-reads numbers, in one undo step", async () => {
    const e = await setup();
    const s = e.wb.sheets[0];
    const n = replaceCells(e, findCells([s], "inline"), "inline", "42");
    expect(n).toBe(1);
    expect(s.cell("B3")!.value).toBe(42);
    e.undo();
    expect(s.cell("B3")!.value).toBe("inline");
  });

  it("uses Google's options: regular expressions with $1, whole words, ignoring accents", async () => {
    const e = await setup();
    const s = e.wb.sheets[0];
    e.input(s, 20, 0, "Smith, Jo");
    e.input(s, 21, 0, "café");
    expect(refs(findCells([s], "^s\\w+,", { regex: true }))).toEqual(["Data!20,0"]);
    expect(refs(findCells([s], "cafe", { ignoreAccents: true }))).toEqual(["Data!21,0"]);
    expect(findCells([s], "Smit", { wholeWord: true })).toHaveLength(0);
    expect(replaceCells(e, findCells([s], "(\\w+), (\\w+)", { regex: true }), "(\\w+), (\\w+)", "$2 $1", { regex: true })).toBe(1);
    expect(s.at(20, 0)!.value).toBe("Jo Smith");
    expect(replaceCells(e, findCells([s], "cafe", { ignoreAccents: true }), "cafe", "tea", { ignoreAccents: true })).toBe(1);
    expect(s.at(21, 0)!.value).toBe("tea");
  });

  it("rewrites formulas only when searching formulas, and recalculates", async () => {
    const e = await setup();
    const s = e.wb.sheets[0];
    expect(replaceCells(e, findCells([s], "60"), "60", "1")).toBe(0);
    expect(replaceCells(e, findCells([s], "SUM", { formulas: true }), "SUM", "MAX", { formulas: true })).toBe(1);
    expect(s.cell("A5")).toMatchObject({ formula: "MAX(A2:A4)", value: 30 });
  });
});
