import { describe, it, expect } from "vitest";
import { XlsxWorkbook } from "../src/infrastructure/xlsx/xlsx-workbook";
import { Workbook } from "../src/domain/xlsx/workbook-model";
import { Edits } from "../src/domain/xlsx/commands";
import { groupRange, outlineGroups, setGroupCollapsed, ungroupRange } from "../src/domain/xlsx/outline";
import { xlsx, sheetXml, S } from "./xlsx-fixtures";

const rows = Array.from({ length: 10 }, (_, i) => `<row r="${i + 1}"><c r="A${i + 1}"><v>${i}</v></c></row>`).join("");
const setup = async (data = rows, extra?: { before: string; after: string }) =>
  new Edits(Workbook.open(await XlsxWorkbook.load(await xlsx({ sheets: [sheetXml(data, extra)] }))));
const rowAttr = (e: Edits, r: number, a: string) => e.wb.sheets[0].rowEl(r, false)?.getAttribute(a) ?? null;
const fmt = (e: Edits) => e.wb.sheets[0].doc.getElementsByTagNameNS(S, "sheetFormatPr")[0];

describe("row and column groups", () => {
  it("reads groups from outline levels, nested ones inside", async () => {
    const data = rows
      .replace('<row r="3">', '<row r="3" outlineLevel="1">')
      .replace('<row r="4">', '<row r="4" outlineLevel="2" hidden="1">')
      .replace('<row r="5">', '<row r="5" outlineLevel="2" hidden="1">')
      .replace('<row r="6">', '<row r="6" outlineLevel="1" collapsed="1">');
    const e = await setup(data);
    expect(outlineGroups(e.wb.sheets[0], "row")).toEqual([
      { level: 1, start: 2, end: 5, collapsed: false },
      { level: 2, start: 3, end: 4, collapsed: true },
    ]);
  });

  it("groups rows: outline levels, sheetFormatPr level, one undo step", async () => {
    const e = await setup();
    groupRange(e, e.wb.sheets[0], "row", 2, 5);
    expect([2, 3, 4, 5].map((r) => rowAttr(e, r, "outlineLevel"))).toEqual(["1", "1", "1", "1"]);
    expect(fmt(e).getAttribute("outlineLevelRow")).toBe("1");
    groupRange(e, e.wb.sheets[0], "row", 3, 4);
    expect(outlineGroups(e.wb.sheets[0], "row").map((g) => [g.level, g.start, g.end])).toEqual([[1, 2, 5], [2, 3, 4]]);
    expect(fmt(e).getAttribute("outlineLevelRow")).toBe("2");
    e.undo();
    e.undo();
    expect(rowAttr(e, 2, "outlineLevel")).toBeNull();
    expect(fmt(e)?.hasAttribute("outlineLevelRow") ?? false).toBe(false);
  });

  it("collapses a group (hiding its rows, marking the row below) and expands it again", async () => {
    const e = await setup();
    groupRange(e, e.wb.sheets[0], "row", 2, 5);
    groupRange(e, e.wb.sheets[0], "row", 3, 4);
    const inner = () => outlineGroups(e.wb.sheets[0], "row")[1];
    const outer = () => outlineGroups(e.wb.sheets[0], "row")[0];
    setGroupCollapsed(e, e.wb.sheets[0], "row", inner(), true);
    setGroupCollapsed(e, e.wb.sheets[0], "row", outer(), true);
    expect([2, 3, 4, 5].every((r) => e.wb.sheets[0].hiddenRows.has(r))).toBe(true);
    expect(rowAttr(e, 6, "collapsed")).toBe("1");
    expect(outer().collapsed).toBe(true);
    setGroupCollapsed(e, e.wb.sheets[0], "row", outer(), false);
    // The nested group stays collapsed.
    expect([2, 3, 4, 5].map((r) => e.wb.sheets[0].hiddenRows.has(r))).toEqual([false, true, true, false]);
    expect(rowAttr(e, 6, "collapsed")).toBeNull();
    expect(inner().collapsed).toBe(true);
  });

  it("ungroups one level and shows rows a collapsed group hid", async () => {
    const e = await setup();
    groupRange(e, e.wb.sheets[0], "row", 2, 5);
    setGroupCollapsed(e, e.wb.sheets[0], "row", outlineGroups(e.wb.sheets[0], "row")[0], true);
    ungroupRange(e, e.wb.sheets[0], "row", 2, 3);
    expect(outlineGroups(e.wb.sheets[0], "row")).toEqual([{ level: 1, start: 4, end: 5, collapsed: false }]);
    expect([2, 3, 4, 5].some((r) => e.wb.sheets[0].hiddenRows.has(r))).toBe(false);
    ungroupRange(e, e.wb.sheets[0], "row", 0, 9);
    expect(outlineGroups(e.wb.sheets[0], "row")).toEqual([]);
    expect(fmt(e)?.hasAttribute("outlineLevelRow") ?? false).toBe(false);
  });

  it("groups and collapses columns", async () => {
    const e = await setup(rows, { before: `<cols><col min="2" max="4" width="20" customWidth="1"/></cols>`, after: "" });
    groupRange(e, e.wb.sheets[0], "col", 1, 2);
    expect(outlineGroups(e.wb.sheets[0], "col")).toEqual([{ level: 1, start: 1, end: 2, collapsed: false }]);
    expect(fmt(e).getAttribute("outlineLevelCol")).toBe("1");
    setGroupCollapsed(e, e.wb.sheets[0], "col", outlineGroups(e.wb.sheets[0], "col")[0], true);
    const s = e.wb.sheets[0];
    expect([s.hiddenCols.has(1), s.hiddenCols.has(2), s.hiddenCols.has(3)]).toEqual([true, true, false]);
    expect(s.colWidth(3)).toBe(140);
    setGroupCollapsed(e, s, "col", outlineGroups(s, "col")[0], false);
    expect(e.wb.sheets[0].colWidth(1)).toBe(140);
  });

  it("puts the summary above when sheetPr says so", async () => {
    const data = rows.replace('<row r="2">', '<row r="2" collapsed="1">').replace('<row r="3">', '<row r="3" outlineLevel="1" hidden="1">');
    const e = await setup(data, { before: `<sheetPr><outlinePr summaryBelow="0"/></sheetPr>`, after: "" });
    expect(outlineGroups(e.wb.sheets[0], "row")).toEqual([{ level: 1, start: 2, end: 2, collapsed: true }]);
  });
});
