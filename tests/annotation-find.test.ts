import { describe, expect, it } from "vitest";
import { PptxDocument } from "../src/infrastructure/pptx/pptx-document";
import { DocxDocument } from "../src/infrastructure/docx/docx-document";
import { XlsxWorkbook } from "../src/infrastructure/xlsx/xlsx-workbook";
import { blankDocx, blankPptx } from "../src/infrastructure/blank";
import { findSlideMatches, replaceAll } from "../src/domain/pptx/find";
import { addSlideComment, resolveSlideThread } from "../src/domain/pptx/comments";
import { findCommentMatches } from "../src/presentation/docx/docx-find";
import { findSheetMatches, replaceCells } from "../src/domain/xlsx/find";
import { Workbook } from "../src/domain/xlsx/workbook-model";
import { Edits } from "../src/domain/xlsx/commands";
import { addComment, resolveThread } from "../src/domain/xlsx/comments";
import { setNote } from "../src/domain/xlsx/notes";
import { textMatches } from "../src/core/utils/text-matches";
import { xlsx } from "./xlsx-fixtures";

describe("annotation search", () => {
  it.each(['.', '[x]', '$', '\\', '(a)', '*'])("treats %s as literal text", query => {
    expect(textMatches(`before ${query} after ${query}`, query)).toHaveLength(2);
  });
  it("returns original Unicode offsets and respects case", () => {
    expect(textMatches('İx NEEDLE needle', 'needle')).toEqual([{ start: 3, end: 9 }, { start: 10, end: 16 }]);
    expect(textMatches('İx NEEDLE needle', 'needle', true)).toEqual([{ start: 10, end: 16 }]);
    expect(textMatches('nothing', '')).toEqual([]);
  });
  it("finds speaker notes, comments, replies and resolved threads without changing the deck", async () => {
    const doc = await PptxDocument.load(await blankPptx());
    doc.setNotes(0, 'Needle\nneedle');
    const path = doc.slides()[0].path;
    const id = addSlideComment(doc, path, 'needle root', 'Author');
    addSlideComment(doc, path, 'reply needle', 'Reply author', id);
    resolveSlideThread(doc, path, id, true);
    const before = doc.snapshot();
    const hits = findSlideMatches(doc, 'needle', { matchCase: false, notes: true, comments: true });
    expect(hits.map(h => h.kind)).toEqual(['notes', 'notes', 'comment', 'comment']);
    expect(hits.every(h => h.slide === 0)).toBe(true);
    expect(findSlideMatches(doc, 'needle', { matchCase: true, notes: true, comments: false })).toHaveLength(1);
    expect(findSlideMatches(doc, 'needle', { matchCase: false, notes: false, comments: false })).toEqual([]);
    expect(doc.snapshot()).toEqual(before);
    expect(replaceAll(doc, 'needle', 'changed', { matchCase: false })).toBe(0);
    expect(doc.notesText(0)).toBe('Needle\nneedle');
  });
  it("searches saved Word comment replies and resolved comments", async () => {
    const doc = await DocxDocument.load(await blankDocx());
    const id = doc.addComment('needle root', 'Author');
    doc.replyToComment(id, 'Reply NEEDLE', 'Reply author');
    doc.setCommentResolved(id, true);
    const reloaded = await DocxDocument.load(await doc.save());
    expect(findCommentMatches(reloaded, 'needle', false)).toHaveLength(2);
    expect(findCommentMatches(reloaded, 'needle', true)).toHaveLength(1);
    expect(findCommentMatches(reloaded, '', false)).toEqual([]);
  });
  it("keeps repeated annotation offsets across line breaks and treats punctuation literally", async () => {
    const doc = await PptxDocument.load(await blankPptx());
    doc.setNotes(0, 'a.b\na.b');
    const hits = findSlideMatches(doc, 'a.b', { matchCase: true, notes: true });
    expect(hits.map(h => [h.start, h.end])).toEqual([[0, 3], [4, 7]]);
    expect(findSlideMatches(doc, 'ab', { matchCase: false, notes: true })).toEqual([]);
    expect(findSlideMatches(doc, '', { matchCase: false, notes: true, comments: true })).toEqual([]);
  });
  it("does not join separate Word comments into one match", async () => {
    const doc = await DocxDocument.load(await blankDocx());
    doc.addComment('hello', 'Author');
    doc.addComment('world', 'Author');
    expect(findCommentMatches(doc, 'hello world', false)).toEqual([]);
    expect(findCommentMatches(doc, 'hello', false)[0]).toMatchObject({ start: 0, end: 5 });
  });
  it("retains sheet annotation matches after saving and does not include disabled sources", async () => {
    const e = new Edits(Workbook.open(await XlsxWorkbook.load(await xlsx())));
    const sheet = e.wb.sheets[0];
    setNote(e.wb.pkg, sheet.ref.path, 'Z99', 'needle needle', 'Author');
    addComment(e, sheet, 'Z99', 'needle thread', 'Author');
    const wb = Workbook.open(await XlsxWorkbook.load(await e.wb.pkg.save()));
    expect(findSheetMatches(wb.sheets, 'needle', { notes: true }).map(h => h.kind)).toEqual(['note', 'note']);
    expect(findSheetMatches(wb.sheets, 'needle', { comments: true }).map(h => h.kind)).toEqual(['comment']);
    expect(findSheetMatches(wb.sheets, 'needle')).toEqual([]);
    expect(findSheetMatches(wb.sheets, '', { notes: true, comments: true })).toEqual([]);
  });
  it("finds sheet notes and comment replies on otherwise empty cells, across sheets", async () => {
    const e = new Edits(Workbook.open(await XlsxWorkbook.load(await xlsx())));
    const [a, b] = e.wb.sheets;
    setNote(e.wb.pkg, a.ref.path, 'Z99', 'Needle note', 'Author');
    const id = addComment(e, b, 'Y98', 'needle root', 'Author');
    addComment(e, b, 'Y98', 'reply NEEDLE', 'Reply author', id);
    resolveThread(e, b, id, true);
    const hits = findSheetMatches(e.wb.sheets, 'needle', { notes: true, comments: true });
    expect(hits.map(h => [h.kind, h.sheet.name, h.row, h.col])).toEqual([['note', a.name, 98, 25], ['comment', b.name, 97, 24], ['comment', b.name, 97, 24]]);
    expect(findSheetMatches([a], 'needle', { notes: true })).toHaveLength(1);
    expect(findSheetMatches([a], 'needle', { notes: true, entire: true })).toHaveLength(0);
    expect(findSheetMatches(e.wb.sheets, 'needle', { notes: true, comments: true, matchCase: true })).toHaveLength(1);
    expect(replaceCells(e, hits.filter(h => h.kind === 'cell'), 'needle', 'changed')).toBe(0);
    expect(findSheetMatches(e.wb.sheets, 'needle', { notes: true, comments: true })).toHaveLength(3);
  });
});
