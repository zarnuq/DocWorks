import { describe, it, expect, vi } from 'vitest';
import { newDrawing, addDrawingShape } from '../src/domain/docx/drawings';
import { drawingEditor } from '../src/presentation/docx/drawing-editor';

describe('drawing canvas',()=>{
 it('moves and resizes in drawing coordinates on a scaled canvas',()=>{
  const drawing=newDrawing(),shape=addDrawingShape(drawing,'rect'),host=createDiv();
  drawingEditor(host,drawing);
  const canvas=host.querySelector<HTMLElement>('.docx-drawing-canvas')!;
  vi.spyOn(canvas,'getBoundingClientRect').mockReturnValue(new DOMRect(0,0,240,120));
  const event=(type:string,x:number,y:number)=>new MouseEvent(type,{clientX:x,clientY:y,bubbles:true,cancelable:true});
  canvas.querySelector('path')!.dispatchEvent(event('pointerdown',20,20));
  canvas.dispatchEvent(event('pointermove',40,35));canvas.dispatchEvent(event('pointerup',40,35));
  expect([shape.x,shape.y]).toEqual([60,50]);
  canvas.querySelector('[data-resize]')!.dispatchEvent(event('pointerdown',100,60));
  canvas.dispatchEvent(event('pointermove',120,70));canvas.dispatchEvent(event('pointerup',120,70));
  expect([shape.w,shape.h]).toEqual([180,100]);
  expect(canvas.style.height).toBe('auto');
 });
});
