import { describe, it, expect } from "vitest";
import { XlsxWorkbook } from "../src/infrastructure/xlsx/xlsx-workbook";
import { Workbook } from "../src/domain/xlsx/workbook-model";
import { Edits } from "../src/domain/xlsx/commands";
import { importCsv, parseCsv, sheetToCsv } from "../src/domain/xlsx/csv";
import { xlsx, sheetXml } from "./xlsx-fixtures";

const setup = async (sheets?: string[]) => new Edits(Workbook.open(await XlsxWorkbook.load(await xlsx(sheets ? { sheets } : {}))));

describe("parseCsv", () => {
  it("reads quoted fields, doubled quotes, embedded newlines and CRLF", () => {
    expect(parseCsv('a,"b,c","say ""hi"""\r\n"multi\nline",,x\n')).toEqual([
      ["a", "b,c", 'say "hi"'],
      ["multi\nline", "", "x"],
    ]);
  });

  it("drops a byte order mark and keeps a last line without newline", () => {
    expect(parseCsv("﻿a,b\n1,2")).toEqual([["a", "b"], ["1", "2"]]);
  });
});

describe("sheetToCsv", () => {
  it("writes displayed values of the used range, quoting where needed", async () => {
    const e = await setup([
      sheetXml(
        `<row r="1"><c r="A1" t="inlineStr"><is><t>x,y</t></is></c><c r="C1" t="inlineStr"><is><t>say "hi"</t></is></c></row>` +
          `<row r="3"><c r="B3"><v>1.5</v></c><c r="C3"><f>B3*2</f><v>3</v></c></row>`
      ),
    ]);
    expect(sheetToCsv(e.wb.sheets[0])).toBe('"x,y",,"say ""hi"""\r\n,,\r\n,1.5,3\r\n');
  });

  it("writes formatted dates as shown", async () => {
    const e = await setup();
    expect(sheetToCsv(e.wb.sheets[0]).split("\r\n")[4]).toBe("60,2024-01-01,#DIV/0!");
  });
});

describe("importCsv", () => {
  it("adds a tab named after the file with values converted like typed input, in one undo step", async () => {
    const e = await setup();
    const name = importCsv(e, [["Item", "Price"], ["pen", "1.50"], ["total", "=SUM(B2)"]], "Prices", 1);
    expect(name).toBe("Prices");
    const s = e.wb.sheet("Prices")!;
    expect(e.wb.sheets.indexOf(s)).toBe(2);
    expect(s.cell("A1")!.value).toBe("Item");
    expect(s.cell("B2")!.value).toBe(1.5);
    expect(s.cell("B3")).toMatchObject({ formula: "SUM(B2)", value: 1.5 });
    e.undo();
    expect(e.wb.sheet("Prices")).toBeUndefined();
    expect(e.wb.sheets.length).toBe(2);
  });

  it("makes the tab name valid and unique", async () => {
    const e = await setup();
    expect(importCsv(e, [["a"]], "Data", 0)).toBe("Data (2)");
    expect(importCsv(e, [["a"]], "a/b:c?" + "x".repeat(40), 0)).toBe("a b c " + "x".repeat(25));
  });
});
