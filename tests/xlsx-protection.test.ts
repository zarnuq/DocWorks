import { describe, expect, it } from "vitest";
import { XlsxWorkbook } from "../src/infrastructure/xlsx/xlsx-workbook";
import { Workbook } from "../src/domain/xlsx/workbook-model";
import { Edits } from "../src/domain/xlsx/commands";
import { pasteText } from "../src/domain/xlsx/clipboard";
import { setProtected, unprotectWithPassword } from "../src/domain/xlsx/protection";
import { setColWidth } from "../src/domain/xlsx/sheet-edit";
import { sheetXml, STYLES, xlsx } from "./xlsx-fixtures";

const setup = async (flags = '') => new Edits(Workbook.open(await XlsxWorkbook.load(await xlsx({
  styles: STYLES.replace('<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>', '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1"><protection locked="0"/></xf>'),
  sheets: [sheetXml('<row r="1"><c r="A1" s="2"><v>1</v></c><c r="B1"><v>2</v></c></row>', { before: '', after: `<sheetProtection sheet="1" ${flags}/>` })],
}))));

describe('sheet protection', () => {
  it('allows unlocked cells and refuses locked and empty cells', async () => {
    const e = await setup(), s = e.wb.sheets[0];
    e.input(s, 0, 0, '3');
    expect(s.at(0, 0)?.value).toBe(3);
    expect(() => e.input(s, 0, 1, '4')).toThrow(/protected/i);
    expect(() => e.input(s, 1, 0, '4')).toThrow(/protected/i);
    expect(s.at(0, 1)?.value).toBe(2);
  });
  it('protects, unprotects and preserves the setting across undo', async () => {
    const e = await setup();
    setProtected(e, e.wb.sheets[0], false);
    e.format(e.wb.sheets[0], {r1: 1, c1: 0, r2: 1, c2: 0}, {locked: false});
    setProtected(e, e.wb.sheets[0], true);
    e.input(e.wb.sheets[0], 1, 0, '7');
    expect(e.wb.sheets[0].at(1, 0)?.value).toBe(7);
    e.undo(); e.undo();
    expect(e.wb.sheets[0].first(e.wb.sheets[0].doc.documentElement, 'sheetProtection')).toBeNull();
  });
  it('does not remove password protection without authentication', async () => {
    const e = await setup('password="ABCD"');
    expect(() => setProtected(e, e.wb.sheets[0], false)).toThrow(/password/i);
  });
  it('removes legacy password protection only with the right password', async () => {
    const e = await setup('password="83AF"');
    expect(await unprotectWithPassword(e, e.wb.sheets[0], 'wrong')).toBe(false);
    expect(e.wb.sheets[0].first(e.wb.sheets[0].doc.documentElement, 'sheetProtection')).not.toBeNull();
    expect(await unprotectWithPassword(e, e.wb.sheets[0], 'password')).toBe(true);
    expect(e.wb.sheets[0].first(e.wb.sheets[0].doc.documentElement, 'sheetProtection')).toBeNull();
  });
  it('checks SHA-512 password hashes with salt and spin count', async () => {
    const e = await setup('algorithmName="SHA-512" hashValue="beRsOEW5hgnfGpC0r+gdvIpkWH0Z1MMQdzvbTFscKwUs/P8cgikVLIjRJliVxPNFIY8IKWPw9fTenrNFQcPUYQ==" saltValue="AAECAwQFBgcICQoLDA0ODw==" spinCount="1000"');
    expect(await unprotectWithPassword(e, e.wb.sheets[0], 'Secret')).toBe(false);
    expect(await unprotectWithPassword(e, e.wb.sheets[0], 'secret')).toBe(true);
    e.undo();
    expect(e.wb.sheets[0].first(e.wb.sheets[0].doc.documentElement, 'sheetProtection')?.getAttribute('spinCount')).toBe('1000');
  });
  it('rolls back an entire paste when a later cell is locked', async () => {
    const e = await setup();
    expect(() => pasteText(e, e.wb.sheets[0], {r1: 0, c1: 0, r2: 0, c2: 0}, '5\t6')).toThrow(/protected/i);
    expect(e.wb.sheets[0].at(0, 0)?.value).toBe(1);
    expect(e.canUndo).toBe(false);
  });
  it('enforces formatting permissions independently of cell locks', async () => {
    const e = await setup();
    expect(() => e.format(e.wb.sheets[0], {r1: 0, c1: 0, r2: 0, c2: 0}, {bold: true})).toThrow(/protected/i);
    const allowed = await setup('formatCells="0" formatColumns="0"');
    allowed.format(allowed.wb.sheets[0], {r1: 0, c1: 1, r2: 0, c2: 1}, {bold: true});
    setColWidth(allowed, allowed.wb.sheets[0], 0, 0, 140);
    expect(allowed.wb.sheets[0].colWidth(0)).toBeCloseTo(140, 0);
    expect(() => setColWidth(e, e.wb.sheets[0], 0, 0, 140)).toThrow(/protected/i);
  });
});
