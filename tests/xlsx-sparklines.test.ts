import { describe, expect, it } from "vitest";
import { XlsxWorkbook } from "../src/infrastructure/xlsx/xlsx-workbook";
import { Workbook } from "../src/domain/xlsx/workbook-model";
import { Edits } from "../src/domain/xlsx/commands";
import { readSparklines } from "../src/domain/xlsx/sparklines";
import * as ops from "../src/domain/xlsx/sheet-edit";
import { sheetXml, xlsx } from "./xlsx-fixtures";

const X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main", XM = "http://schemas.microsoft.com/office/excel/2006/main";
const num = (ref: string, v: number) => `<c r="${ref}"><v>${v}</v></c>`;
const DATA = [3, -1, 4, 1, 5].map((v, i) => `<row r="${i + 1}">${num(`A${i + 1}`, v)}</row>`).join("");
const group = (attrs: string, colors: string, f: string, sqref: string) =>
  `<x14:sparklineGroup ${attrs}>${colors}<x14:sparklines><x14:sparkline><xm:f>${f}</xm:f><xm:sqref>${sqref}</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup>`;
const ext = (groups: string) =>
  `<extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="${X14}"><x14:sparklineGroups xmlns:xm="${XM}">${groups}</x14:sparklineGroups></ext></extLst>`;

async function setup(groups: string) {
  const bytes = await xlsx({ sheets: [sheetXml(DATA, { before: "", after: ext(groups) })], names: ["Data"] });
  return new Edits(Workbook.open(await XlsxWorkbook.load(bytes)));
}

describe("sparklines", () => {
  it("reads a line sparkline with its values and colours", async () => {
    const e = await setup(group(`displayEmptyCellsAs="gap" markers="1"`, `<x14:colorSeries rgb="FF376092"/><x14:colorMarkers rgb="FFD00000"/>`, "Data!A1:A5", "B1"));
    const [s] = readSparklines(e.wb.sheets[0]);
    expect(s).toMatchObject({ row: 0, col: 1, type: "line", values: [3, -1, 4, 1, 5], color: "#376092", markers: "#D00000" });
  });

  it("reads column and win/loss sparklines, with negative and high point colours only when switched on", async () => {
    const colors = `<x14:colorSeries rgb="FF376092"/><x14:colorNegative rgb="FFFF0000"/><x14:colorHigh rgb="FF00FF00"/>`;
    const e = await setup(group(`type="column" negative="1" high="1"`, colors, "Data!A1:A5", "B1") + group(`type="stacked"`, colors, "Data!A1:A5", "C1"));
    const [col, wl] = readSparklines(e.wb.sheets[0]);
    expect(col).toMatchObject({ type: "column", negative: "#FF0000", high: "#00FF00" });
    expect(wl.type).toBe("stacked");
    expect(wl.negative).toBeUndefined();
  });

  it("uses a custom axis minimum and maximum", async () => {
    const e = await setup(group(`minAxisType="custom" maxAxisType="custom" manualMin="-10" manualMax="10"`, `<x14:colorSeries rgb="FF376092"/>`, "Data!A1:A5", "B1"));
    expect(readSparklines(e.wb.sheets[0])[0]).toMatchObject({ min: -10, max: 10 });
  });

  it("moves with inserted rows and goes when its cell's row is deleted", async () => {
    const e = await setup(group("", `<x14:colorSeries rgb="FF376092"/>`, "Data!A1:A5", "B5"));
    ops.insertRows(e, e.wb.sheets[0], 0, 2);
    expect(readSparklines(e.wb.sheets[0])[0]).toMatchObject({ row: 6, col: 1, values: [3, -1, 4, 1, 5] });
    expect(e.wb.sheets[0].doc.getElementsByTagNameNS(XM, "f")[0].textContent).toBe("Data!A3:A7");
    ops.deleteRows(e, e.wb.sheets[0], 6, 1);
    expect(readSparklines(e.wb.sheets[0])).toHaveLength(0);
    expect(e.wb.sheets[0].doc.getElementsByTagNameNS(X14, "sparklineGroups")).toHaveLength(0);
  });
});
