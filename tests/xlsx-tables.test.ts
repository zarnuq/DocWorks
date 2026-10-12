import { describe, expect, it } from "vitest";
import { XlsxWorkbook } from "../src/infrastructure/xlsx/xlsx-workbook";
import { Workbook } from "../src/domain/xlsx/workbook-model";
import { Edits } from "../src/domain/xlsx/commands";
import { insertRows } from "../src/domain/xlsx/sheet-edit";
import { translateFormula } from "../src/domain/xlsx/formula";
import { xlsx, sheetXml, S } from "./xlsx-fixtures";
const setup = async () => {
 const pkg = await XlsxWorkbook.load(await xlsx({sheets: [sheetXml('<row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Amount</t></is></c></row><row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>10</v></c></row><row r="3"><c r="A3"><v>2</v></c><c r="B3"><v>20</v></c></row><row r="4"><c r="B4"><v>30</v></c></row>')], parts: {'xl/worksheets/_rels/sheet1.xml.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>', 'xl/tables/table1.xml': `<table xmlns="${S}" id="1" name="Sales" displayName="Sales" ref="A1:B4" totalsRowCount="1"><tableColumns count="2"><tableColumn id="1" name="Name"/><tableColumn id="2" name="Amount"/></tableColumns></table>`}}));
 return new Edits(Workbook.open(pkg));
};
describe('structured table references', () => {
 it('evaluates data, headers, totals and column spans', async () => {
  const e = await setup(), s = e.wb.sheets[0];
  expect(e.calc.evaluate(s, 8, 4, 'SUM(Sales[Amount])')).toBe(30);
  expect(e.calc.evaluate(s, 8, 4, 'SUM(Sales[[#All],[Amount]])')).toBe(60);
  expect(e.calc.evaluate(s, 8, 4, 'Sales[[#Headers],[Amount]]')).toBe('Amount');
  expect(e.calc.evaluate(s, 8, 4, 'Sales[[#Totals],[Amount]]')).toBe(30);
  expect(e.calc.evaluate(s, 8, 4, 'SUM(Sales[[Name]:[Amount]])')).toBe(33);
 });
 it('resolves this-row references and short references inside a table', async () => {
  const e = await setup(), s = e.wb.sheets[0];
  expect(e.calc.evaluate(s, 2, 0, 'Sales[@Amount]*2')).toBe(40);
  expect(e.calc.evaluate(s, 1, 0, '[@Amount]*2')).toBe(20);
  expect(e.calc.evaluate(s, 1, 0, 'SUM([Amount])')).toBe(30);
  expect(e.calc.evaluate(s, 8, 4, 'Sales[@Amount]')).toEqual({error: '#VALUE!'});
  expect(e.calc.evaluate(s, 8, 4, 'Sales[Missing]')).toEqual({error: '#REF!'});
 });
 it('indexes dependencies and follows table growth after inserting rows', async () => {
  const e = await setup();
  e.input(e.wb.sheets[0], 0, 3, '=SUM(Sales[Amount])');
  e.input(e.wb.sheets[0], 1, 1, '15');
  expect(e.wb.sheets[0].at(0, 3)?.value).toBe(35);
  insertRows(e, e.wb.sheets[0], 2, 1);
  e.input(e.wb.sheets[0], 2, 1, '7');
  expect(e.wb.sheets[0].at(0, 3)?.value).toBe(42);
  expect(translateFormula('SUM(Sales[Amount])+A1', 1, 1)).toBe('SUM(Sales[Amount])+B2');
 });
});
