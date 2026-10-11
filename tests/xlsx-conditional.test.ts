import { describe, expect, it } from "vitest";
import { XlsxWorkbook } from "../src/infrastructure/xlsx/xlsx-workbook";
import { Workbook } from "../src/domain/xlsx/workbook-model";
import { Edits } from "../src/domain/xlsx/commands";
import { addRule, ConditionalFormats, deleteRule, readRules } from "../src/domain/xlsx/conditional";
import { sheetXml, STYLES, S, xlsx } from "./xlsx-fixtures";

const DXFS = `<dxfs count="2"><dxf><font><b/></font><fill><patternFill><bgColor rgb="FFFF0000"/></patternFill></fill></dxf><dxf><font><color rgb="FF0000FF"/></font></dxf></dxfs>`;
const num = (ref: string, v: number) => `<c r="${ref}"><v>${v}</v></c>`;
const str = (ref: string, v: string) => `<c r="${ref}" t="inlineStr"><is><t>${v}</t></is></c>`;
const DATA =
  [1, 2, 3, 4, 5].map((v, i) => `<row r="${i + 1}">${num(`A${i + 1}`, v)}${i < 4 ? str(`B${i + 1}`, ["banana", "apple", "", "banana"][i]) : ""}</row>`).join("");

async function setup(cf: string) {
  const bytes = await xlsx({ sheets: [sheetXml(DATA, { before: "", after: cf })], names: ["Data"], styles: STYLES.replace("</styleSheet>", `${DXFS}</styleSheet>`) });
  const e = new Edits(Workbook.open(await XlsxWorkbook.load(bytes)));
  const sheet = e.wb.sheets[0];
  const cf_ = () => new ConditionalFormats(e.wb.sheets[0], (row, col, f) => e.calc.evaluate(e.wb.sheets[0], row, col, f));
  return { e, sheet, cf: cf_ };
}
const rule = (sqref: string, body: string) => `<conditionalFormatting sqref="${sqref}">${body}</conditionalFormatting>`;

describe("conditional formatting", () => {
  it("styles cells by value with the rule's format (cell is greater than)", async () => {
    const { cf } = await setup(rule("A1:A5", `<cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>3</formula></cfRule>`));
    const f = cf();
    expect(f.at(2, 0)).toBeNull();
    expect(f.at(3, 0)).toEqual({ fill: "#FF0000", bold: true });
    expect(f.at(4, 0)).toEqual({ fill: "#FF0000", bold: true });
    expect(f.at(3, 1)).toBeNull();
  });

  it("evaluates custom formulas relative to each cell, and lets the higher priority rule win", async () => {
    const { cf } = await setup(
      rule("A1:A5", `<cfRule type="expression" dxfId="1" priority="1"><formula>A1=MAX($A$1:$A$5)</formula></cfRule>`) +
        rule("A1:A5", `<cfRule type="cellIs" dxfId="0" priority="2" operator="between"><formula>4</formula><formula>5</formula></cfRule>`)
    );
    const f = cf();
    expect(f.at(4, 0)).toEqual({ color: "#0000FF", fill: "#FF0000", bold: true });
    expect(f.at(3, 0)).toEqual({ fill: "#FF0000", bold: true });
    expect(f.at(0, 0)).toBeNull();
  });

  it("matches text rules, blanks, duplicates and top / bottom ranks", async () => {
    const { cf } = await setup(
      rule("B1:B4", `<cfRule type="containsText" dxfId="1" priority="1" operator="containsText" text="AN"><formula>NOT(ISERROR(SEARCH("AN",B1)))</formula></cfRule>`) +
        rule("B1:B4", `<cfRule type="containsBlanks" dxfId="0" priority="2"><formula>LEN(TRIM(B1))=0</formula></cfRule>`) +
        rule("A1:A5", `<cfRule type="top10" dxfId="1" priority="3" rank="2" bottom="1"/>`) +
        rule("B1:B4", `<cfRule type="duplicateValues" dxfId="0" priority="4"/>`)
    );
    const f = cf();
    expect(f.at(0, 1)?.color).toBe("#0000FF");
    expect(f.at(1, 1)).toBeNull();
    expect(f.at(2, 1)).toEqual({ fill: "#FF0000", bold: true });
    expect(f.at(3, 1)).toMatchObject({ color: "#0000FF", fill: "#FF0000" });
    expect([0, 1, 2].map((r) => f.at(r, 0)?.color ?? null)).toEqual(["#0000FF", "#0000FF", null]);
  });

  it("shades a colour scale between its end colours (Google's percentile and number points)", async () => {
    const { cf } = await setup(rule("A1:A5", `<cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFFFFFFF"/><color rgb="FF57BB8A"/></colorScale></cfRule>`));
    const f = cf();
    expect(f.at(0, 0)!.fill).toBe("#FFFFFF");
    expect(f.at(4, 0)!.fill).toBe("#57BB8A");
    expect(f.at(2, 0)!.fill).toBe("#ABDDC5");
    const three = await setup(rule("A1:A5", `<cfRule type="colorScale" priority="1"><colorScale><cfvo type="num" val="1"/><cfvo type="percentile" val="50"/><cfvo type="formula" val="5"/><color rgb="FFE67C73"/><color rgb="FFFFFFFF"/><color rgb="FF57BB8A"/></colorScale></cfRule>`));
    expect(three.cf().at(2, 0)!.fill).toBe("#FFFFFF");
    expect(three.cf().at(0, 0)!.fill).toBe("#E67C73");
  });

  it("draws data bars as a fraction of the range", async () => {
    const { cf } = await setup(rule("A1:A5", `<cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar></cfRule>`));
    expect(cf().at(4, 0)!.bar).toEqual({ fraction: 1, color: "#638EC6" });
    expect(cf().at(2, 0)!.bar!.fraction).toBeCloseTo(0.5);
  });

  it("adds and deletes rules as one undo step each, with their own format record", async () => {
    const { e, cf } = await setup("");
    addRule(e, e.wb.sheets[0], [{ r1: 0, c1: 0, r2: 4, c2: 0 }], { type: "cellIs", operator: "lessThan", values: ["3"], style: { fill: "#00FF00", italic: true } });
    expect(cf().at(0, 0)).toEqual({ fill: "#00FF00", italic: true });
    addRule(e, e.wb.sheets[0], [{ r1: 0, c1: 1, r2: 3, c2: 1 }], { type: "containsText", values: ["apple"], style: { color: "#123456" } });
    expect(cf().at(1, 1)).toEqual({ color: "#123456" });
    addRule(e, e.wb.sheets[0], [{ r1: 0, c1: 0, r2: 4, c2: 0 }], { type: "colorScale", colors: ["#FFFFFF", "#000000"] });
    const rules = readRules(e.wb.sheets[0]);
    expect(rules.map((r) => r.type)).toEqual(["cellIs", "containsText", "colorScale"]);
    // Worksheet child order: conditional formatting before page margins.
    const names = Array.from(e.wb.sheets[0].doc.documentElement.children).map((c) => c.localName);
    expect(names.indexOf("conditionalFormatting")).toBeLessThan(names.indexOf("pageMargins") < 0 ? Infinity : names.indexOf("pageMargins"));
    deleteRule(e, e.wb.sheets[0], readRules(e.wb.sheets[0])[0]);
    expect(readRules(e.wb.sheets[0]).map((r) => r.type)).toEqual(["containsText", "colorScale"]);
    e.undo();
    expect(readRules(e.wb.sheets[0])).toHaveLength(3);
    expect(e.wb.sheets[0].doc.getElementsByTagNameNS(S, "cfRule")).toHaveLength(3);
  });

  it("reads Google's colour scale with formula points from a real export", async () => {
    const { cf } = await setup(rule("A1:A5", `<cfRule type="colorScale" priority="1"><colorScale><cfvo type="formula" val="0"/><cfvo type="formula" val="0.5"/><cfvo type="formula" val="1"/><color rgb="FFE67C73"/><color rgb="FFFFFFFF"/><color rgb="FF57BB8A"/></colorScale></cfRule>`));
    // Every value is above the top point: the top colour.
    expect(cf().at(1, 0)!.fill).toBe("#57BB8A");
  });
});
