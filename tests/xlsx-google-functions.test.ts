import { describe, expect, it } from "vitest";
import { Edits } from "../src/domain/xlsx/commands";
import { Workbook } from "../src/domain/xlsx/workbook-model";
import { XlsxWorkbook } from "../src/infrastructure/xlsx/xlsx-workbook";
import { formulaSparklines } from "../src/domain/xlsx/sparklines";
import { xlsx, sheetXml } from "./xlsx-fixtures";

const str = (ref: string, v: string) => `<c r="${ref}" t="inlineStr"><is><t>${v}</t></is></c>`;
const num = (ref: string, v: number) => `<c r="${ref}"><v>${v}</v></c>`;
const ROWS: [string, string, number][] = [["East", "A", 10], ["East", "B", 20], ["West", "A", 5], ["West", "B", 15], ["North", "B", 3]];
const DATA = `<row r="1">${str("A1", "Region")}${str("B1", "Product")}${str("C1", "Amount")}</row>` + ROWS.map(([r, p, a], i) => `<row r="${i + 2}">${str(`A${i + 2}`, r)}${str(`B${i + 2}`, p)}${num(`C${i + 2}`, a)}</row>`).join("");
const setup = async (extra = "") => new Edits(Workbook.open(await XlsxWorkbook.load(await xlsx({ sheets: [sheetXml(DATA + extra)] }))));
const col = (e: Edits, c: number, r1: number, r2: number) => Array.from({ length: r2 - r1 + 1 }, (_, i) => e.wb.sheets[0].at(r1 + i, c)?.value ?? null);

describe("Google Sheets functions", () => {
  it("ARRAYFORMULA works out arithmetic over ranges and spills it", async () => {
    const e = await setup();
    e.input(e.wb.sheets[0], 1, 4, "=ARRAYFORMULA(C2:C4*2)");
    expect(col(e, 4, 1, 3)).toEqual([20, 40, 10]);
  });

  it("REGEXMATCH, REGEXEXTRACT and REGEXREPLACE, also element by element in ARRAYFORMULA", async () => {
    const e = await setup(), s = e.wb.sheets[0];
    e.input(s, 10, 0, '=REGEXMATCH(A2,"^Ea")');
    e.input(s, 10, 1, '=REGEXEXTRACT("order-1234-x","\\d+")');
    e.input(s, 10, 2, '=REGEXEXTRACT("key=value","(\\w+)=")');
    e.input(s, 10, 3, '=REGEXREPLACE("a-b-c","-","+")');
    e.input(s, 10, 4, '=REGEXEXTRACT("abc","\\d")');
    expect([s.at(10, 0)?.value, s.at(10, 1)?.value, s.at(10, 2)?.value, s.at(10, 3)?.value, s.at(10, 4)?.value]).toEqual([true, "1234", "key", "a+b+c", { error: "#N/A" }]);
    e.input(s, 1, 5, '=ARRAYFORMULA(REGEXMATCH(A2:A4,"st$"))');
    expect(col(e, 5, 1, 3)).toEqual([true, true, true]);
    e.input(s, 12, 0, '=REGEXTEST("Hello","^h")');
    expect(s.at(12, 0)?.value).toBe(false);
  });

  it("QUERY returns a table that spills, with its header", async () => {
    const e = await setup(), s = e.wb.sheets[0];
    e.input(s, 0, 6, '=QUERY(A1:C6,"select A, sum(C) group by A order by sum(C) desc",1)');
    expect(col(e, 6, 0, 3)).toEqual(["Region", "East", "West", "North"]);
    expect(col(e, 7, 0, 3)).toEqual(["sum Amount", 30, 20, 3]);
  });

  it("evaluates Google-only functions exported as __xludf.DUMMYFUNCTION, instead of keeping the cache", async () => {
    const e = await setup(`<row r="20"><c r="A20"><f>IFERROR(__xludf.DUMMYFUNCTION("REGEXMATCH(A2,""Ea"")"),FALSE)</f><v>0</v></c></row>`);
    // The file's cached result stands until something it reads changes; then it's worked out here.
    e.input(e.wb.sheets[0], 1, 0, "Eagle");
    expect(e.wb.sheets[0].at(19, 0)?.value).toBe(true);
    e.input(e.wb.sheets[0], 1, 0, "West");
    expect(e.wb.sheets[0].at(19, 0)?.value).toBe(false);
  });

  it("SPARKLINE shows nothing as a value and draws a sparkline in its cell", async () => {
    const e = await setup(), s = e.wb.sheets[0];
    e.input(s, 1, 8, '=SPARKLINE(C2:C6,{"charttype","column";"color1","red"})');
    expect(s.at(1, 8)?.value).toBe("");
    const [sp] = formulaSparklines(s, (sheet, row, column, f) => e.calc.evaluateArray(sheet, row, column, f));
    expect(sp).toMatchObject({ row: 1, col: 8, type: "column", values: [10, 20, 5, 15, 3], color: "red" });
  });
});
