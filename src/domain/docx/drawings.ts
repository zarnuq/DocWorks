/** Native Word shapes and groups. Edited properties are patched into each shape's original XML. */
import { NS } from '../../core/utils/xml';
import { parseXml } from '../../infrastructure/ooxml/ooxml-package';
import { EMU_PER_PX } from "./docx-reader";
import { Placement, placementOf, setPlacement } from './drawing';
import { presetGeometry, customGeometry, Geometry } from '../pptx/geometry';

export const WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
export const WPG = 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup';
export interface DrawingShape {
  id: string; kind: string; x: number; y: number; w: number; h: number;
  fill: string | null; line: string | null; text: string;
  /** Original shape XML, retained for properties this editor does not model. */
  xml: string;
}
export interface WordDrawing {
  width: number; height: number; placement: Placement; shapes: DrawingShape[];
  xml: string; editable: boolean;
}
const first = (el: Element | undefined | null, ns: string, tag: string) => Array.from(el?.children ?? []).find(c => c.namespaceURI === ns && c.localName === tag);
const serialize = (el: Element) => new XMLSerializer().serializeToString(el);
const px = (el: Element | undefined, attr: string, fallback = 0) => el?.hasAttribute(attr) ? Number(el.getAttribute(attr)) / EMU_PER_PX : fallback;
const color = (el: Element | undefined) => {
  const value = first(el, NS.a, 'srgbClr')?.getAttribute('val');
  return value && /^[0-9a-f]{6}$/i.test(value) ? `#${value}` : null;
};
export function newDrawing(width = 480, height = 240): WordDrawing {
  return { width, height, placement: { wrap: 'inline', x: 0, y: 0, align: null }, shapes: [], xml: '', editable: true };
}
export function addDrawingShape(drawing: WordDrawing, kind: string): DrawingShape {
  if (!presetGeometry(kind, 100, 100)) throw new Error('Unknown shape.');
  const shape: DrawingShape = { id: crypto.randomUUID(), kind, x: 20 + drawing.shapes.length * 20, y: 20 + drawing.shapes.length * 20, w: 140, h: 80, fill: kind === 'line' ? null : '#D9EAF7', line: '#34699A', text: '', xml: '' };
  drawing.shapes.push(shape); return shape;
}
export function readDrawing(xml: string): WordDrawing | null {
  if (!xml.includes('wordprocessingShape')) return null;
  const doc = parseXml(xml); if (!doc) return null;
  const frame = doc.getElementsByTagNameNS(NS.wp, 'inline')[0] ?? doc.getElementsByTagNameNS(NS.wp, 'anchor')[0];
  const graphic = doc.getElementsByTagNameNS(NS.a, 'graphicData')[0];
  if (!frame || !graphic || ![WPS, WPG].includes(graphic.getAttribute('uri') ?? '')) return null;
  const extent = first(frame, NS.wp, 'extent');
  const group = first(graphic, WPG, 'wgp');
  const elements = group ? Array.from(group.children).filter(c => c.namespaceURI === WPS && c.localName === 'wsp') : [first(graphic, WPS, 'wsp')].filter((s): s is Element => !!s);
  if (!elements.length) return null;
  const shapes = elements.map((el, i): DrawingShape => {
    const props = first(el, WPS, 'spPr'), transform = first(props, NS.a, 'xfrm');
    const off = first(transform, NS.a, 'off'), ext = first(transform, NS.a, 'ext');
    const text = Array.from(el.getElementsByTagNameNS(NS.w, 'p')).map(p => Array.from(p.getElementsByTagNameNS(NS.w, 't')).map(t => t.textContent).join('')).join('\n');
    return { id: first(el, WPS, 'cNvPr')?.getAttribute('id') ?? String(i + 1), kind: first(props, NS.a, 'prstGeom')?.getAttribute('prst') ?? 'custom', x: group ? px(off, 'x') : 0, y: group ? px(off, 'y') : 0, w: px(ext, 'cx', px(extent, 'cx')), h: px(ext, 'cy', px(extent, 'cy')), fill: color(first(props, NS.a, 'solidFill')), line: color(first(first(props, NS.a, 'ln'), NS.a, 'solidFill')), text, xml: serialize(el) };
  });
  // Nested groups / pictures and transformed coordinate systems stay preserved and cannot be flattened by this editor.
  const groupTransform = first(first(group, WPG, 'grpSpPr'), NS.a, 'xfrm');
  const groupExt = first(groupTransform, NS.a, 'ext'), childExt = first(groupTransform, NS.a, 'chExt');
  const nested = group && Array.from(group.children).some(c => !['cNvPr', 'cNvGrpSpPr', 'grpSpPr', 'wsp', 'extLst'].includes(c.localName));
  const childOff = first(groupTransform, NS.a, 'chOff');
  const transformed = groupTransform && (['rot', 'flipH', 'flipV'].some(a => groupTransform.hasAttribute(a)) || ['cx', 'cy'].some(a => groupExt?.getAttribute(a) !== childExt?.getAttribute(a)) || ['x','y'].some(a=>Number(childOff?.getAttribute(a)??0)!==0));
  const editable = !nested && !transformed && shapes.every(s => s.kind !== 'custom' && !['rot','flipH','flipV'].some(a=>first(first(parseXml(s.xml)?.documentElement,WPS,'spPr'),NS.a,'xfrm')?.hasAttribute(a)));
  return { width: px(extent, 'cx', 480), height: px(extent, 'cy', 240), placement: placementOf(doc.documentElement), shapes, xml, editable };
}
export function drawingGeometry(shape: DrawingShape): Geometry | null {
  if (shape.kind !== 'custom') return presetGeometry(shape.kind, shape.w, shape.h);
  const doc = parseXml(shape.xml), custom = doc?.getElementsByTagNameNS(NS.a, 'custGeom')[0];
  return custom ? customGeometry(custom, shape.w, shape.h) : null;
}

function shapeXml(shape: DrawingShape, index: number): string {
  const doc = parseXml(shape.xml || `<wps:wsp xmlns:wps="${WPS}" xmlns:a="${NS.a}" xmlns:w="${NS.w}"><wps:cNvPr id="${index}" name="Shape ${index}"/><wps:cNvSpPr/><wps:spPr/><wps:bodyPr anchor="ctr"/></wps:wsp>`)!;
  const root = doc.documentElement;
  const previous = shape.xml ? readDrawing(`<w:r xmlns:w="${NS.w}" xmlns:wp="${NS.wp}" xmlns:a="${NS.a}"><w:drawing><wp:inline><wp:extent cx="${Math.round(shape.w * EMU_PER_PX)}" cy="${Math.round(shape.h * EMU_PER_PX)}"/><a:graphic><a:graphicData uri="${WPS}">${shape.xml}</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`)?.shapes[0] : null;
  const make = (ns: string, tag: string, attrs: Record<string, string> = {}) => {
    const el = doc.createElementNS(ns, `${ns === WPS ? 'wps' : ns === NS.w ? 'w' : 'a'}:${tag}`);
    for (const [k,v] of Object.entries(attrs)) el.setAttribute(k,v); return el;
  };
  const props = first(root, WPS, 'spPr')!;
  const order = ['xfrm','prstGeom','custGeom','noFill','solidFill','gradFill','blipFill','pattFill','grpFill','ln','effectLst','effectDag','scene3d','sp3d','extLst'];
  const put = (el: Element) => props.insertBefore(el, Array.from(props.children).find(c => order.indexOf(c.localName) > order.indexOf(el.localName)) ?? null);
  let transform = first(props, NS.a, 'xfrm');
  if (!transform) transform = put(make(NS.a, 'xfrm'));
  for (const [tag,attrs] of [['off',{x:shape.x,y:shape.y}],['ext',{cx:shape.w,cy:shape.h}]] as const) {
    let el = first(transform, NS.a, tag); if (!el) el = transform.appendChild(make(NS.a, tag));
    for (const [k,v] of Object.entries(attrs)) el.setAttribute(k,String(Math.round(v*EMU_PER_PX)));
  }
  if (!previous || previous.kind !== shape.kind) {
    first(props, NS.a, 'prstGeom')?.remove(); first(props, NS.a, 'custGeom')?.remove();
    const geom = put(make(NS.a, 'prstGeom', {prst:shape.kind})); geom.appendChild(make(NS.a, 'avLst'));
  }
  const fill = (parent: Element, value: string | null) => {
    const el = make(NS.a, value ? 'solidFill' : 'noFill');
    if (value) el.appendChild(make(NS.a, 'srgbClr', {val:value.replace('#','').toUpperCase()}));
    for (const c of Array.from(parent.children)) if (['noFill','solidFill','gradFill','blipFill','pattFill','grpFill'].includes(c.localName)) c.remove();
    if (parent === props) put(el); else parent.insertBefore(el,parent.firstChild);
  };
  if (!previous || previous.fill !== shape.fill) fill(props,shape.fill);
  if (!previous || previous.line !== shape.line) {
    let line = first(props,NS.a,'ln'); if (!line) line = put(make(NS.a,'ln',{w:'12700'})); fill(line,shape.line);
  }
  if (!previous || previous.text !== shape.text) {
    let box = first(root,WPS,'txbx');
    if (!box) box = root.insertBefore(make(WPS,'txbx'),first(root,WPS,'bodyPr') ?? null);
    const oldParagraphs = Array.from(box.getElementsByTagNameNS(NS.w,'p'));
    const content = make(NS.w,'txbxContent');
    for (const [i,text] of shape.text.split('\n').entries()) {
      const p = make(NS.w,'p'), pPr=make(NS.w,'pPr'), jc=make(NS.w,'jc',{'w:val':'center'});
      pPr.appendChild(jc);p.appendChild(pPr);
      const run=make(NS.w,'r'),rPr=make(NS.w,'rPr'),t=make(NS.w,'t');
      rPr.appendChild(make(NS.w,'rFonts',{'w:ascii':'Arial','w:hAnsi':'Arial'}));
      rPr.appendChild(make(NS.w,'color',{'w:val':'222222'}));
      rPr.appendChild(make(NS.w,'sz',{'w:val':'21'}));
      const original = oldParagraphs[i] ?? oldParagraphs[0];
      const oldPPr = first(original,NS.w,'pPr');
      if (oldPPr) p.replaceChild(oldPPr.cloneNode(true),pPr);
      const oldRPr = original?.getElementsByTagNameNS(NS.w,'rPr')[0];
      run.appendChild(oldRPr ? oldRPr.cloneNode(true) : rPr);
      t.setAttribute('xml:space','preserve');t.textContent=text;run.appendChild(t);p.appendChild(run);content.appendChild(p);
    }
    box.replaceChildren(content);
  }
  return serialize(root);
}

export function writeDrawing(drawing: WordDrawing, docPrId: number): string {
  if (!drawing.editable) throw new Error('This drawing contains objects that cannot be edited here.');
  const cx=Math.round(drawing.width*EMU_PER_PX),cy=Math.round(drawing.height*EMU_PER_PX);
  let doc = drawing.xml ? parseXml(drawing.xml) : null;
  if (!doc) doc = parseXml(`<w:r xmlns:w="${NS.w}" xmlns:wp="${NS.wp}" xmlns:a="${NS.a}" xmlns:wps="${WPS}" xmlns:wpg="${WPG}"><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${docPrId}" name="Drawing"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="${WPG}"><wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/><a:chOff x="0" y="0"/><a:chExt cx="${cx}" cy="${cy}"/></a:xfrm></wpg:grpSpPr></wpg:wgp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`)!;
  const graphic = doc.getElementsByTagNameNS(NS.a,'graphicData')[0];
  const group = first(graphic,WPG,'wgp');
  if (!group && drawing.shapes.length !== 1) throw new Error('Add shapes in a new drawing.');
  const parent=group ?? graphic;
  for (const el of Array.from(parent.children)) if (el.namespaceURI === WPS && el.localName === 'wsp') el.remove();
  let nextId = Math.max(0, ...drawing.shapes.filter(s=>s.xml).map(s=>Number(s.id)||0)) + 1;
  drawing.shapes.forEach(shape => parent.insertBefore(doc.importNode(parseXml(shapeXml(shape,shape.xml ? Number(shape.id) : nextId++))!.documentElement,true),Array.from(parent.children).find(c=>c.localName==='extLst') ?? null));
  const extent=doc.getElementsByTagNameNS(NS.wp,'extent')[0];extent.setAttribute('cx',String(cx));extent.setAttribute('cy',String(cy));
  if(group) for(const tag of ['ext','chExt']) {const ext=first(first(first(group,WPG,'grpSpPr'),NS.a,'xfrm'),NS.a,tag);ext?.setAttribute('cx',String(cx));ext?.setAttribute('cy',String(cy));}
  setPlacement(doc.documentElement,drawing.placement);
  return serialize(doc.documentElement);
}
