import { describe, expect, it } from 'vitest';
import { newDrawing, addDrawingShape, readDrawing, writeDrawing } from '../src/domain/docx/drawings';
import { readDocx } from '../src/domain/docx/docx-reader';
import { writeBody } from '../src/domain/docx/docx-writer';
import { NS } from '../src/core/utils/xml';

describe('Word drawings', () => {
 it('saves a drawing with multiple shapes and editable text as native Word content', () => {
  const drawing=newDrawing(400,200);
  addDrawingShape(drawing,'rect');addDrawingShape(drawing,'ellipse');
  drawing.shapes[0].text='Hello & <world>';drawing.shapes[1].x=220;
  const xml=writeDrawing(drawing,7),again=readDrawing(xml)!;
  expect(again.shapes.map(s=>s.kind)).toEqual(['rect','ellipse']);
  expect(again.shapes[0].text).toBe('Hello & <world>');
  expect(again.shapes[1].x).toBe(220);
  const root=new DOMParser().parseFromString(xml,'application/xml');
  expect(root.getElementsByTagNameNS(NS.wp,'docPr')[0].getAttribute('id')).toBe('7');
  expect(root.getElementsByTagNameNS(NS.w,'txbxContent')).toHaveLength(2);
 });
 it('preserves unedited drawings verbatim through the document reader and writer', () => {
  const drawing=newDrawing();addDrawingShape(drawing,'rightArrow');
  const xml=writeDrawing(drawing,1);
  const body=new DOMParser().parseFromString(`<w:body xmlns:w="${NS.w}"><w:p>${xml}</w:p></w:body>`,'application/xml').documentElement;
  const before=new XMLSerializer().serializeToString(body);
  const {doc,origins}=readDocx(body,{imageSrc:()=>null,hyperlinkHref:()=>null});
  writeBody(body,doc,origins,{nextDocPrId:()=>2,ensureParagraphStyles:()=>{}});
  expect(new XMLSerializer().serializeToString(body)).toBe(before);
 });
 it('patches changed shape properties while retaining unmodeled XML', () => {
  const drawing=newDrawing();addDrawingShape(drawing,'rect');
  const xml=writeDrawing(drawing,1).replace('</wps:spPr>','<a:effectLst><a:outerShdw blurRad="30000"/></a:effectLst></wps:spPr>');
  const model=readDrawing(xml)!;model.shapes[0].fill='#FF0000';model.shapes[0].w=220;
  const out=writeDrawing(model,1);
  expect(out).toContain('outerShdw');expect(out).toContain('FF0000');
  expect(readDrawing(out)!.shapes[0].w).toBe(220);
 });
 it('keeps shape identifiers unique after reordering and adding shapes', () => {
  const drawing=newDrawing();addDrawingShape(drawing,'rect');addDrawingShape(drawing,'ellipse');
  const edited=readDrawing(writeDrawing(drawing,1))!;
  edited.shapes.reverse();addDrawingShape(edited,'triangle');
  edited.shapes.push(edited.shapes.shift()!);
  const shapes=readDrawing(writeDrawing(edited,1))!.shapes;
  expect(new Set(shapes.map(s=>s.id)).size).toBe(3);
 });
 it('keeps existing text styles when editing the words', () => {
  const drawing=newDrawing();addDrawingShape(drawing,'rect').text='Before';
  const xml=writeDrawing(drawing,1).replace('<w:rPr>','<w:rPr><w:b/>').replace('w:val="21"','w:val="28"');
  const edited=readDrawing(xml)!;edited.shapes[0].text='After';
  const saved=writeDrawing(edited,1);
  expect(saved).toContain('<w:b');expect(saved).toContain('w:val="28"');
 });
});
