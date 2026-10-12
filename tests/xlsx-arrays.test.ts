import { describe, expect, it } from "vitest";
import { Edits } from "../src/domain/xlsx/commands";
import { Workbook } from "../src/domain/xlsx/workbook-model";
import { XlsxWorkbook } from "../src/infrastructure/xlsx/xlsx-workbook";
import { xlsx, sheetXml } from "./xlsx-fixtures";
const setup = async () => new Edits(Workbook.open(await XlsxWorkbook.load(await xlsx({sheets:[sheetXml('')]}))));
const rect = (r1: number, c1: number, r2 = r1, c2 = c1) => ({r1,c1,r2,c2});

describe('dynamic arrays', () => {
 it('spills sequence results and preserves the formula across save and reopen', async () => {
  const e = await setup(), s = e.wb.sheets[0];
  e.input(s,0,0,'=SEQUENCE(3,2,10,2)');
  expect([s.at(0,0)?.value,s.at(0,1)?.value,s.at(2,0)?.value,s.at(2,1)?.value]).toEqual([10,12,18,20]);
  expect(s.at(0,0)?.array).toEqual(rect(0,0,2,1));
  expect(() => e.input(s,1,1,'99')).toThrow(/array/i);
  const again = Workbook.open(await XlsxWorkbook.load(await e.wb.pkg.save()));
  expect(again.sheets[0].at(2,1)?.value).toBe(20);
  expect(again.sheets[0].at(0,0)?.formula).toContain('SEQUENCE');
  expect(again.sheets[0].at(0,0)?.el.hasAttribute('cm')).toBe(true);
 });
 it('grows and shrinks spills, recalculates readers, and undoes the entire edit', async () => {
  const e = await setup(), s = e.wb.sheets[0];
  e.input(s,0,3,'3'); e.input(s,0,0,'=SEQUENCE(D1)'); e.input(s,0,4,'=SUM(A1#)');
  expect(s.at(0,4)?.value).toBe(6);
  e.input(s,0,3,'2');
  expect(s.at(2,0)?.value ?? null).toBeNull(); expect(s.at(0,4)?.value).toBe(3);
  e.undo();
  expect(e.wb.sheets[0].at(2,0)?.value).toBe(3); expect(e.wb.sheets[0].at(0,4)?.value).toBe(6);
  e.redo(); expect(e.wb.sheets[0].at(0,4)?.value).toBe(3);
  e.clear(e.wb.sheets[0],rect(0,0));
  expect(e.wb.sheets[0].at(1,0)?.value ?? null).toBeNull();
 });
 it('never overwrites occupied or merged cells, and retries when a blocker is cleared', async () => {
  const e = await setup(), s = e.wb.sheets[0];
  e.input(s,1,0,'blocked'); e.input(s,0,0,'=SEQUENCE(3)');
  expect(s.at(0,0)?.value).toEqual({error:'#SPILL!'}); expect(s.at(1,0)?.value).toBe('blocked');
  e.clear(s,rect(1,0)); expect(s.at(2,0)?.value).toBe(3);
  s.merges.push(rect(0,2,1,2)); e.input(s,0,2,'=SEQUENCE(2)');
  expect(s.at(0,2)?.value).toEqual({error:'#SPILL!'});
 });
 it('filters, sorts, deduplicates and transposes values with array conditions', async () => {
  const e = await setup(), s = e.wb.sheets[0];
  [3,1,3,2].forEach((n,r) => e.input(s,r,0,String(n)));
  e.input(s,0,2,'=SORT(UNIQUE(FILTER(A1:A4,A1:A4>1)))');
  expect([s.at(0,2)?.value,s.at(1,2)?.value]).toEqual([2,3]);
  e.input(s,0,4,'=TRANSPOSE(A1:A4)');
  expect([s.at(0,4)?.value,s.at(0,7)?.value]).toEqual([3,2]);
 });
 it('recalculates a formula reading a spill member after a resize', async () => {
  const e = await setup(), s = e.wb.sheets[0];
  e.input(s,0,3,'2');e.input(s,0,0,'=SEQUENCE(D1,1,10)');e.input(s,0,4,'=A2*2');
  expect(s.at(0,4)?.value).toBe(22);
  e.input(s,0,3,'1');expect(s.at(0,4)?.value).toBe(0);
  e.input(s,0,3,'3');expect(s.at(0,4)?.value).toBe(22);
 });
 it('supports arithmetic arrays and clears all outputs when replacing the anchor', async () => {
  const e=await setup(),s=e.wb.sheets[0];
  e.input(s,0,0,'=SEQUENCE(2,2)*2');expect(s.at(1,1)?.value).toBe(8);
  e.input(s,0,0,'hello');expect(s.at(1,1)?.value ?? null).toBeNull();
 });
 it('keeps existing cached spills byte-identical when merely opened', async () => {
  const e = await setup(); e.input(e.wb.sheets[0],0,0,'=SEQUENCE(2)');
  const bytes = await e.wb.pkg.save();
  const again = new Edits(Workbook.open(await XlsxWorkbook.load(bytes)));
  expect(await again.wb.pkg.save()).toEqual(bytes);
 });
});
