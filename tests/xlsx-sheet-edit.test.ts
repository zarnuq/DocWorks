import { describe, it, expect } from "vitest";
import { XlsxWorkbook } from "../src/infrastructure/xlsx/xlsx-workbook";
import { Workbook } from "../src/domain/xlsx/workbook-model";
import { Edits } from "../src/domain/xlsx/commands";
import * as ops from "../src/domain/xlsx/sheet-edit";
import { xlsx, S } from "./xlsx-fixtures";

const setup = async () => new Edits(Workbook.open(await XlsxWorkbook.load(await xlsx())));
const rect = (r1: number, c1: number, r2 = r1, c2 = c1) => ({ r1, c1, r2, c2 });
const attr = (doc: Document, tag: string, a: string) => Array.from(doc.getElementsByTagNameNS(S, tag)).map((e) => e.getAttribute(a));

describe("inserting and deleting rows", () => {
  it("moves cells down and updates formulas on every sheet", async () => {
    const e = await setup();
    ops.insertRows(e, e.wb.sheets[0], 2, 2);
    const s = e.wb.sheets[0];
    expect(s.cell("A3")).toBeUndefined();
    expect(s.cell("A5")!.value).toBe(20);
    expect(s.cell("A7")).toMatchObject({ formula: "SUM(A2:A6)", value: 60 });
    expect(s.cell("C5")!.formula).toBe("A5*2");
    expect(e.wb.sheets[1].cell("B1")!.formula).toBe("SUM(Data!A2:A6)");
    expect(s.rowHeight(5)).toBe(40); // the tall row moved with its cells
    expect(s.merges).toEqual([rect(7, 1, 7, 2)]);
    expect(s.doc.getElementsByTagNameNS(S, "dimension")[0].getAttribute("ref")).toBe("A1:D8");
  });

  it("deletes rows, shrinking ranges and turning lost references into #REF!", async () => {
    const e = await setup();
    ops.deleteRows(e, e.wb.sheets[0], 1, 1);
    const s = e.wb.sheets[0];
    expect(s.cell("A2")!.value).toBe(20);
    expect(s.cell("A4")).toMatchObject({ formula: "SUM(A2:A3)", value: 50 });
    ops.deleteRows(e, s, 0, 1);
    expect(e.wb.sheets[0].cell("A4")!.formula).toBe('Totals!A1&"!"');
    expect(e.wb.sheets[0].merges).toEqual([rect(3, 1, 3, 2)]);
  });

  it("is one undo step", async () => {
    const e = await setup();
    const before = new XMLSerializer().serializeToString(e.wb.sheets[0].doc);
    ops.insertRows(e, e.wb.sheets[0], 0, 3);
    e.undo();
    expect(new XMLSerializer().serializeToString(e.wb.sheets[0].doc)).toBe(before);
    expect(e.wb.sheets[0].cell("A5")!.formula).toBe("SUM(A2:A4)");
    e.redo();
    expect(e.wb.sheets[0].cell("A8")!.formula).toBe("SUM(A5:A7)");
  });
});

describe("inserting and deleting columns", () => {
  it("moves cells right with their column widths", async () => {
    const e = await setup();
    ops.insertCols(e, e.wb.sheets[0], 1, 1);
    const s = e.wb.sheets[0];
    expect(s.cell("C1")!.value).toBe("Notes");
    expect(s.colWidth(2)).toBe(140);
    expect(s.colWidth(1)).toBe(64);
    expect(s.cell("D3")!.formula).toBe("A3*2");
    expect(attr(s.doc, "col", "min")).toEqual(["3", "5"]);
  });

  it("deletes columns", async () => {
    const e = await setup();
    ops.deleteCols(e, e.wb.sheets[0], 0, 1);
    const s = e.wb.sheets[0];
    expect(s.cell("A1")!.value).toBe("Notes");
    expect(s.cell("B2")!.formula).toBe("#REF!*2");
    expect(s.cell("B2")!.value).toEqual({ error: "#REF!" });
  });
});

describe("sheets", () => {
  it("adds a sheet with a unique name after the current one", async () => {
    const e = await setup();
    const name = ops.addSheet(e, 0);
    expect(name).toBe("Sheet3");
    expect(e.wb.sheets.map((s) => s.name)).toEqual(["Data", "Sheet3", "Totals"]);
    e.undo();
    expect(e.wb.sheets.map((s) => s.name)).toEqual(["Data", "Totals"]);
  });

  it("renames a sheet and every reference to it", async () => {
    const e = await setup();
    ops.renameSheet(e, e.wb.sheets[0], "Raw data");
    expect(e.wb.sheets[0].name).toBe("Raw data");
    expect(e.wb.sheets[1].cell("B1")!.formula).toBe("SUM('Raw data'!A2:A4)");
    expect(() => ops.renameSheet(e, e.wb.sheets[0], "totals")).toThrow(/already/);
    expect(() => ops.renameSheet(e, e.wb.sheets[0], "a/b")).toThrow(/character/);
  });

  it("deletes a sheet, turning references to it into #REF!", async () => {
    const e = await setup();
    ops.deleteSheet(e, e.wb.sheets[1]);
    expect(e.wb.sheets.map((s) => s.name)).toEqual(["Data"]);
    expect(e.wb.sheets[0].cell("A6")).toMatchObject({ formula: '#REF!&"!"', value: { error: "#REF!" } });
    expect(() => ops.deleteSheet(e, e.wb.sheets[0])).toThrow(/last/);
    e.undo();
    expect(e.wb.sheets.map((s) => s.name)).toEqual(["Data", "Totals"]);
  });

  it("moves and duplicates sheets", async () => {
    const e = await setup();
    ops.moveSheet(e, 1, 0);
    expect(e.wb.sheets.map((s) => s.name)).toEqual(["Totals", "Data"]);
    ops.duplicateSheet(e, e.wb.sheets[1]);
    expect(e.wb.sheets.map((s) => s.name)).toEqual(["Totals", "Data", "Copy of Data"]);
    expect(e.wb.sheets[2].cell("A5")!.formula).toBe("SUM(A2:A4)");
    expect(e.wb.sheets[2].doc.getElementsByTagNameNS(S, "drawing")).toHaveLength(0);
  });

  it("hides and shows sheets, but never the last visible one", async () => {
    const e = await setup();
    ops.setSheetHidden(e, e.wb.sheets[1], true);
    expect(e.wb.sheets[1].ref.state).toBe("hidden");
    expect(() => ops.setSheetHidden(e, e.wb.sheets[0], true)).toThrow(/visible/);
  });
});

describe("layout", () => {
  it("sets column widths and row heights", async () => {
    const e = await setup();
    const s = e.wb.sheets[0];
    ops.setColWidth(e, s, 0, 2, 100);
    expect([0, 1, 2].map((c) => e.wb.sheets[0].colWidth(c))).toEqual([100, 100, 100]);
    ops.setRowHeight(e, e.wb.sheets[0], 0, 0, 30);
    expect(e.wb.sheets[0].rowHeight(0)).toBe(30);
  });

  it("hides and unhides rows and columns", async () => {
    const e = await setup();
    ops.setHidden(e, e.wb.sheets[0], "col", 0, 0, true);
    expect(e.wb.sheets[0].colWidth(0)).toBe(0);
    ops.setHidden(e, e.wb.sheets[0], "col", 3, 3, false);
    expect(e.wb.sheets[0].colWidth(3)).toBe(63);
    ops.setHidden(e, e.wb.sheets[0], "row", 5, 5, false);
    expect(e.wb.sheets[0].rowHeight(5)).toBe(20);
  });

  it("freezes panes", async () => {
    const e = await setup();
    ops.freeze(e, e.wb.sheets[0], 2, 1);
    expect(e.wb.sheets[0].frozen).toEqual({ rows: 2, cols: 1 });
    ops.freeze(e, e.wb.sheets[0], 0, 0);
    expect(e.wb.sheets[0].frozen).toEqual({ rows: 0, cols: 0 });
    expect(e.wb.sheets[0].doc.getElementsByTagNameNS(S, "pane")).toHaveLength(0);
  });

  it("merges, keeping only the top-left value, and unmerges", async () => {
    const e = await setup();
    ops.merge(e, e.wb.sheets[0], rect(1, 0, 2, 1));
    const s = e.wb.sheets[0];
    expect(s.merges).toContainEqual(rect(1, 0, 2, 1));
    expect(s.cell("A2")!.value).toBe(10);
    expect(s.cell("B2")?.value ?? null).toBeNull();
    ops.unmerge(e, s, rect(1, 0));
    expect(e.wb.sheets[0].merges).not.toContainEqual(rect(1, 0, 2, 1));
  });
});

describe("sorting", () => {
  it("sorts rows by a column, moving whole rows and their formulas", async () => {
    const e = await setup();
    const s = e.wb.sheets[0];
    ops.sortRange(e, s, rect(1, 0, 3, 2), [{ col: 0, desc: true }]);
    expect([1, 2, 3].map((r) => s.at(r, 0)!.value)).toEqual([30, 20, 10]);
    expect([1, 2, 3].map((r) => s.at(r, 2)!.formula)).toEqual(["A2*2", "A3*2", "A4*2"]);
    expect([1, 2, 3].map((r) => s.at(r, 1)?.value)).toEqual([true, "inline", "rich text"]);
  });

  it("puts blanks last either way and sorts text case-insensitively", async () => {
    const e = await setup();
    const s = e.wb.sheets[0];
    for (const [r, v] of [[10, "b"], [11, ""], [12, "A"], [13, "c"]] as const) e.input(s, r, 5, v);
    ops.sortRange(e, s, rect(10, 5, 13, 5), [{ col: 5 }]);
    expect([10, 11, 12, 13].map((r) => s.at(r, 5)?.value ?? null)).toEqual(["A", "b", "c", null]);
  });
});

describe("filters", () => {
  it("hides rows that don't match and shows them again", async () => {
    const e = await setup();
    ops.setFilter(e, e.wb.sheets[0], rect(0, 0, 4, 2));
    expect(e.wb.sheets[0].filter).toEqual(rect(0, 0, 4, 2));
    ops.filterColumn(e, e.wb.sheets[0], 0, new Set(["10", "30"]));
    const s = e.wb.sheets[0];
    expect([1, 2, 3, 4].map((r) => s.hiddenRows.has(r))).toEqual([false, true, false, true]);
    expect(s.doc.getElementsByTagNameNS(S, "filterColumn")).toHaveLength(1);
    ops.filterColumn(e, s, 0, null);
    expect([1, 2, 3, 4].map((r) => e.wb.sheets[0].hiddenRows.has(r))).toEqual([false, false, false, false]);
    ops.setFilter(e, e.wb.sheets[0], null);
    expect(e.wb.sheets[0].filter).toBeNull();
  });
});

describe("formatting of inserted rows and columns", () => {
  it("gives inserted rows the formats of the row above, and columns those of the column to the left", async () => {
    const e = await setup();
    ops.insertRows(e, e.wb.sheets[0], 1, 2); // below the bold header row
    const s = e.wb.sheets[0];
    expect([s.cell("A2")?.style, s.cell("B3")?.style]).toEqual([1, 1]);
    expect(s.cell("A2")!.value).toBeNull();
    ops.insertCols(e, e.wb.sheets[0], 1, 1);
    expect(e.wb.sheets[0].cell("B1")?.style).toBe(1);
  });
});
