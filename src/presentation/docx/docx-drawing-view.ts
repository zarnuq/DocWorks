import { DOMSerializer, Node as PMNode } from 'prosemirror-model';
import type { EditorView, NodeView } from 'prosemirror-view';
import { drawingGeometry, readDrawing, WordDrawing } from '../../domain/docx/drawings';
import { placementCss } from '../../domain/docx/drawing';

const SVG = 'http://www.w3.org/2000/svg';
export function renderDrawing(host: HTMLElement, drawing: WordDrawing, selected?: string): SVGSVGElement {
  host.empty();
  const svg = document.createElementNS(SVG,'svg');
  svg.setAttribute('viewBox',`0 0 ${drawing.width} ${drawing.height}`);
  svg.setAttribute('width','100%');svg.setAttribute('height','100%');
  // Named by <title>: an aria-label would get Obsidian's hover tooltip, which fails on SVG elements.
  svg.setAttribute('role','img');svg.appendChild(document.createElementNS(SVG,'title')).textContent='Drawing';host.appendChild(svg);
  for (const shape of drawing.shapes) {
    const g=document.createElementNS(SVG,'g');g.dataset.shape=shape.id;g.setAttribute('transform',`translate(${shape.x} ${shape.y})`);svg.appendChild(g);
    const geom=drawingGeometry(shape);
    for(const p of geom?.paths ?? []) {
      const path=document.createElementNS(SVG,'path');path.setAttribute('d',p.d);
      path.setAttribute('fill',p.fill==='none' ? 'none' : shape.fill ?? 'none');
      path.setAttribute('stroke',p.stroke ? shape.line ?? 'none' : 'none');path.setAttribute('stroke-width','1.33');g.appendChild(path);
    }
    const lines=shape.text.split('\n');
    lines.forEach((line,i)=>{
      const text=document.createElementNS(SVG,'text');text.textContent=line;
      text.setAttribute('x',String(shape.w/2));text.setAttribute('y',String(shape.h/2+(i-(lines.length-1)/2)*18));
      text.setAttribute('text-anchor','middle');text.setAttribute('dominant-baseline','middle');text.setAttribute('fill','#222222');text.setAttribute('font-size','14');g.appendChild(text);
    });
    if(shape.id===selected) {
      const border=document.createElementNS(SVG,'rect');
      for(const [k,v] of Object.entries({x:0,y:0,width:shape.w,height:shape.h,fill:'none',stroke:'#1a73e8','stroke-dasharray':'4 2'})) border.setAttribute(k,String(v));
      border.setAttribute('pointer-events','none');g.appendChild(border);
      const handle=document.createElementNS(SVG,'rect');
      for(const [k,v] of Object.entries({x:shape.w-4,y:shape.h-4,width:8,height:8,fill:'#1a73e8'})) handle.setAttribute(k,String(v));
      handle.dataset.resize=shape.id;g.appendChild(handle);
    }
  }
  return svg;
}

/** Draw a native Word drawing as one selectable object; double-click opens its drawing editor. */
export class DrawingView implements NodeView {
  dom: HTMLElement;
  constructor(node: PMNode, _view: EditorView, getPos: () => number | undefined, edit: (pos: number) => void) {
    const drawing=readDrawing(node.attrs.xml as string);
    if(!drawing) {this.dom=DOMSerializer.renderSpec(document,node.type.spec.toDOM!(node)).dom;return;}
    const p=placementCss(drawing.placement);
    this.dom=createSpan({cls:`docx-drawing ${p.cls}`,attr:{contenteditable:'false'}});
    this.dom.setCssStyles({display:'inline-block',width:`${drawing.width}px`,height:`${drawing.height}px`});
    if(p.style) this.dom.style.cssText+=`;${p.style}`;
    renderDrawing(this.dom,drawing);
    this.dom.addEventListener('dblclick',e=>{e.preventDefault();const pos=getPos();if(pos!==undefined && drawing.editable)edit(pos);});
  }
  ignoreMutation(): boolean {return true;}
}
