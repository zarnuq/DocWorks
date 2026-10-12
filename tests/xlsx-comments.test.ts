import { describe, expect, it } from 'vitest';
import { Edits } from '../src/domain/xlsx/commands';
import { Workbook } from '../src/domain/xlsx/workbook-model';
import { XlsxWorkbook } from '../src/infrastructure/xlsx/xlsx-workbook';
import { addComment, commentsOf, editComment, removeThread, resolveThread } from '../src/domain/xlsx/comments';
import { insertRows, deleteRows } from '../src/domain/xlsx/sheet-edit';
import { xlsx } from './xlsx-fixtures';
const setup = async () => new Edits(Workbook.open(await XlsxWorkbook.load(await xlsx())));
describe('threaded spreadsheet comments', () => {
 it('adds a thread and replies, preserving authors and resolution across save and reopen', async () => {
  const e=await setup(),path=e.wb.sheets[0].ref.path;
  const id=addComment(e,e.wb.sheets[0],'B3','Check this','Miles');
  addComment(e,e.wb.sheets[0],'B3','Agreed','Sam',id);
  resolveThread(e,e.wb.sheets[0],id,true);
  const again=Workbook.open(await XlsxWorkbook.load(await e.wb.pkg.save()));
  const thread=commentsOf(again.pkg,path).get('B3')![0];
  expect(thread.resolved).toBe(true);
  expect(thread.messages.map(m=>[m.text,m.author])).toEqual([['Check this','Miles'],['Agreed','Sam']]);
 });
 it('edits text without losing replies, deletes a whole thread, and undoes it', async () => {
  const e=await setup(),path=e.wb.sheets[0].ref.path;
  const id=addComment(e,e.wb.sheets[0],'A2','First','Miles');
  addComment(e,e.wb.sheets[0],'A2','Reply','Sam',id);
  editComment(e,e.wb.sheets[0],id,'Corrected');
  expect(commentsOf(e.wb.pkg,path).get('A2')![0].messages[0].text).toBe('Corrected');
  removeThread(e,e.wb.sheets[0],id);
  expect(commentsOf(e.wb.pkg,path).size).toBe(0);
  e.undo();expect(commentsOf(e.wb.pkg,path).get('A2')![0].messages).toHaveLength(2);
 });
 it('moves complete threads with structural edits and removes threads on deleted cells', async () => {
  const e=await setup(),path=e.wb.sheets[0].ref.path;
  const id=addComment(e,e.wb.sheets[0],'A3','First','Miles');
  addComment(e,e.wb.sheets[0],'A3','Reply','Sam',id);
  insertRows(e,e.wb.sheets[0],0,1);
  expect(commentsOf(e.wb.pkg,path).get('A4')![0].messages).toHaveLength(2);
  deleteRows(e,e.wb.sheets[0],3,1);
  expect(commentsOf(e.wb.pkg,path).size).toBe(0);
 });
});
