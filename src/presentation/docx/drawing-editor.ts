/** The drawing canvas and controls, shared by insert and edit dialogs. */
import { addDrawingShape, WordDrawing } from '../../domain/docx/drawings';
import { renderDrawing } from './docx-drawing-view';

const SHAPES: [string,string][]=[['Rectangle','rect'],['Rounded rectangle','roundRect'],['Ellipse','ellipse'],['Triangle','triangle'],['Diamond','diamond'],['Right arrow','rightArrow'],['Line','line'],['Text box','rect']];
export function drawingEditor(host: HTMLElement, drawing: WordDrawing): void {
  let selected: string | undefined=drawing.shapes[0]?.id;
  const tools=host.createDiv({cls:'docx-drawing-tools'});
  const choice=tools.createEl('select',{attr:{'aria-label':'Shape'}});
  for(const [label,value] of SHAPES) choice.createEl('option',{text:label,value});
  const button=(parent:HTMLElement,label:string,fn:()=>void)=>{const b=parent.createEl('button',{text:label});b.addEventListener('click',fn);};
  button(tools,'Add shape',()=>{const s=addDrawingShape(drawing,choice.value);selected=s.id;draw();});
  const canvas=host.createDiv({cls:'docx-drawing-canvas'});
  canvas.setCssStyles({width:`${drawing.width}px`,height:'auto',maxWidth:'100%',aspectRatio:`${drawing.width} / ${drawing.height}`});
  const props=host.createDiv({cls:'docx-drawing-properties'});
  const active=()=>drawing.shapes.find(s=>s.id===selected);
  const draw=()=>{
    renderDrawing(canvas,drawing,selected);props.empty();const s=active();if(!s)return;
    const field=(label:string,value:string,type:string,change:(v:string)=>void)=>{
      const row=props.createEl('label',{cls:'docx-drawing-field'});row.createSpan({text:label});
      const input=row.createEl('input',{type,value,attr:{'aria-label':label}});input.addEventListener('change',()=>{change(input.value);draw();});return input;
    };
    for(const [label,key] of [['X','x'],['Y','y'],['Width','w'],['Height','h']] as const) field(label,String(s[key]),'number',v=>{const n=Number(v);if(Number.isFinite(n))s[key]=Math.max(key==='w'||key==='h'?10:0,n);});
    field('Fill',s.fill??'#ffffff','color',v=>s.fill=v);field('Border',s.line??'#000000','color',v=>s.line=v);
    button(props,'No fill',()=>{s.fill=null;draw();});button(props,'No border',()=>{s.line=null;draw();});
    const text=props.createEl('textarea',{value:s.text,attr:{'aria-label':'Shape text',rows:'3'}});
    text.addEventListener('input',()=>{s.text=text.value;renderDrawing(canvas,drawing,selected);});
    button(props,'Bring to front',()=>{drawing.shapes=drawing.shapes.filter(x=>x!==s);drawing.shapes.push(s);draw();});
    button(props,'Send to back',()=>{drawing.shapes=drawing.shapes.filter(x=>x!==s);drawing.shapes.unshift(s);draw();});
    button(props,'Delete shape',()=>{drawing.shapes=drawing.shapes.filter(x=>x!==s);selected=drawing.shapes.at(-1)?.id;draw();});
  };
  canvas.addEventListener('pointerdown',e=>{
    const target=e.target as Element,s=drawing.shapes.find(s=>s.id===target.closest<SVGElement>('[data-shape]')?.dataset.shape);if(!s)return;
    e.preventDefault();selected=s.id;
    const resize=target.hasAttribute('data-resize'),box=canvas.getBoundingClientRect(),scale=drawing.width/box.width;
    const start={x:e.clientX,y:e.clientY,sx:s.x,sy:s.y,w:s.w,h:s.h};
    canvas.setPointerCapture?.(e.pointerId);
    const move=(ev:PointerEvent)=>{
      const dx=(ev.clientX-start.x)*scale,dy=(ev.clientY-start.y)*scale;
      if(resize){s.w=Math.max(10,start.w+dx);s.h=Math.max(10,start.h+dy);}else{s.x=Math.max(0,start.sx+dx);s.y=Math.max(0,start.sy+dy);}
      renderDrawing(canvas,drawing,selected);
    };
    const end=()=>{canvas.removeEventListener('pointermove',move);canvas.removeEventListener('pointerup',end);canvas.removeEventListener('pointercancel',end);draw();};
    canvas.addEventListener('pointermove',move);canvas.addEventListener('pointerup',end);canvas.addEventListener('pointercancel',end);draw();
  });
  draw();
}
