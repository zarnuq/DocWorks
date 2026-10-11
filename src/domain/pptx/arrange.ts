/*
 * Arranging slide items (top-level shapes, pictures, groups...): bounds,
 * grouping, align / distribute, snapping, and copy / paste between slides and
 * presentations.
 */
import { NS, childNS } from "../../core/utils/xml";
import type { ResourceBundle } from "../../infrastructure/ooxml/ooxml-package";
import type { PptxDocument } from "../../infrastructure/pptx/pptx-document";
import type { Rect } from "./geometry";
import { nextShapeId, setRect } from "./slide-edit";

const SHAPE_TAGS = new Set(["sp", "cxnSp", "pic", "graphicFrame", "grpSp", "contentPart", "AlternateContent"]);
const kid = (el: Element | null | undefined, ns: string, local: string) => (el ? childNS(el, ns, local) : null);
const num = (el: Element | null, attr: string) => Number(el?.getAttribute(attr) ?? 0);

/** The slide item (direct child of p:spTree) that `el` belongs to. */
export function topLevel(el: Element): Element {
  let cur = el;
  while (cur.parentElement && cur.parentElement.localName !== "spTree") cur = cur.parentElement;
  return cur;
}

export function xfrmOf(item: Element): Element | null {
  if (item.localName === "grpSp") return kid(kid(item, NS.p, "grpSpPr"), NS.a, "xfrm");
  if (item.localName === "graphicFrame") return kid(item, NS.p, "xfrm");
  return kid(kid(item, NS.p, "spPr"), NS.a, "xfrm");
}

/** Position and size from the item's own xfrm (zeros when it inherits them). */
export function itemRect(item: Element): Rect {
  const xfrm = xfrmOf(item);
  const off = kid(xfrm, NS.a, "off");
  const ext = kid(xfrm, NS.a, "ext");
  return { x: num(off, "x"), y: num(off, "y"), w: num(ext, "cx"), h: num(ext, "cy") };
}

/** Move / resize an item; a group keeps its child coordinates, so its children follow. */
export function setItemRect(item: Element, r: Rect): void {
  if (item.localName !== "grpSp") return setRect(item, r);
  const xfrm = xfrmOf(item)!;
  const round = (v: number) => String(Math.round(v));
  const off = kid(xfrm, NS.a, "off")!;
  const ext = kid(xfrm, NS.a, "ext")!;
  off.setAttribute("x", round(r.x));
  off.setAttribute("y", round(r.y));
  ext.setAttribute("cx", round(r.w));
  ext.setAttribute("cy", round(r.h));
}

const union = (rects: Rect[]): Rect => {
  const x = Math.min(...rects.map((r) => r.x));
  const y = Math.min(...rects.map((r) => r.y));
  return { x, y, w: Math.max(...rects.map((r) => r.x + r.w)) - x, h: Math.max(...rects.map((r) => r.y + r.h)) - y };
};

/** Put sibling items into a new p:grpSp at the position of the topmost one. */
export function groupItems(slide: Document, items: Element[], rects = items.map(itemRect)): Element {
  const tree = items[0].parentElement!;
  const order = Array.from(tree.children);
  const sorted = items.map((el, i) => ({ el, r: rects[i] })).sort((a, b) => order.indexOf(a.el) - order.indexOf(b.el));
  const b = union(rects);
  const id = nextShapeId(slide);
  const xml =
    `<p:grpSp xmlns:p="${NS.p}" xmlns:a="${NS.a}"><p:nvGrpSpPr><p:cNvPr id="${id}" name="Group ${id}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
    `<p:grpSpPr><a:xfrm><a:off x="${b.x}" y="${b.y}"/><a:ext cx="${b.w}" cy="${b.h}"/>` +
    `<a:chOff x="${b.x}" y="${b.y}"/><a:chExt cx="${b.w}" cy="${b.h}"/></a:xfrm></p:grpSpPr></p:grpSp>`;
  const group = slide.importNode(new DOMParser().parseFromString(xml, "application/xml").documentElement, true);
  tree.insertBefore(group, sorted[sorted.length - 1].el.nextSibling);
  for (const { el, r } of sorted) {
    group.appendChild(el);
    // Placeholders may inherit their position; give every child an explicit one.
    setItemRect(el, r);
  }
  return group;
}

/** Dissolve a group, keeping each child where it appears. Returns the children. */
export function ungroupItem(group: Element): Element[] {
  const xfrm = xfrmOf(group);
  const off = kid(xfrm, NS.a, "off");
  const ext = kid(xfrm, NS.a, "ext");
  const chOff = kid(xfrm, NS.a, "chOff");
  const chExt = kid(xfrm, NS.a, "chExt");
  const sx = num(chExt, "cx") ? num(ext, "cx") / num(chExt, "cx") : 1;
  const sy = num(chExt, "cy") ? num(ext, "cy") / num(chExt, "cy") : 1;
  const children = Array.from(group.children).filter((c) => SHAPE_TAGS.has(c.localName));
  for (const child of children) {
    const r = itemRect(child);
    setItemRect(child, {
      x: num(off, "x") + (r.x - num(chOff, "x")) * sx,
      y: num(off, "y") + (r.y - num(chOff, "y")) * sy,
      w: r.w * sx,
      h: r.h * sy,
    });
    group.parentElement!.insertBefore(child, group);
  }
  group.remove();
  return children;
}

export type Alignment = "left" | "center" | "right" | "top" | "middle" | "bottom";

/** Align rects to their common bounds (or to `bounds`, e.g. the slide for a single shape). */
export function alignRects(rects: Rect[], mode: Alignment, bounds = union(rects)): Rect[] {
  return rects.map((r) => {
    switch (mode) {
      case "left":
        return { ...r, x: bounds.x };
      case "center":
        return { ...r, x: bounds.x + (bounds.w - r.w) / 2 };
      case "right":
        return { ...r, x: bounds.x + bounds.w - r.w };
      case "top":
        return { ...r, y: bounds.y };
      case "middle":
        return { ...r, y: bounds.y + (bounds.h - r.h) / 2 };
      case "bottom":
        return { ...r, y: bounds.y + bounds.h - r.h };
    }
  });
}

/** Equal gaps between rects along an axis; the outermost two stay put. */
export function distributeRects(rects: Rect[], axis: "horizontal" | "vertical"): Rect[] {
  if (rects.length < 3) return rects;
  const [pos, size] = axis === "horizontal" ? (["x", "w"] as const) : (["y", "h"] as const);
  const order = rects.map((r, i) => i).sort((a, b) => rects[a][pos] - rects[b][pos]);
  const first = rects[order[0]];
  const last = rects[order[order.length - 1]];
  const span = last[pos] + last[size] - first[pos];
  const gap = (span - rects.reduce((s, r) => s + r[size], 0)) / (rects.length - 1);
  const out = rects.map((r) => ({ ...r }));
  let at = first[pos];
  for (const i of order) {
    out[i][pos] = at;
    at += rects[i][size] + gap;
  }
  return out;
}

export interface Snap {
  dx: number;
  dy: number;
  /** Guide lines to show: a vertical line at x, or a horizontal one at y. */
  guides: ({ x: number } | { y: number })[];
}

/** Nudge `moving` so an edge or center lines up with another shape or the slide, within `tolerance`. */
export function snapRect(moving: Rect, others: Rect[], slide: { cx: number; cy: number }, tolerance: number): Snap {
  const axis = (lines: number[], targets: number[]) => {
    let best: { delta: number; at: number } | null = null;
    for (const line of lines) {
      for (const t of targets) {
        const delta = t - line;
        if (Math.abs(delta) <= tolerance && (!best || Math.abs(delta) < Math.abs(best.delta))) best = { delta, at: t };
      }
    }
    return best;
  };
  const xs = [0, slide.cx / 2, slide.cx, ...others.flatMap((o) => [o.x, o.x + o.w / 2, o.x + o.w])];
  const ys = [0, slide.cy / 2, slide.cy, ...others.flatMap((o) => [o.y, o.y + o.h / 2, o.y + o.h])];
  const sx = axis([moving.x, moving.x + moving.w / 2, moving.x + moving.w], xs);
  const sy = axis([moving.y, moving.y + moving.h / 2, moving.y + moving.h], ys);
  const guides: Snap["guides"] = [];
  if (sx) guides.push({ x: sx.at });
  if (sy) guides.push({ y: sy.at });
  return { dx: sx?.delta ?? 0, dy: sy?.delta ?? 0, guides };
}

// -----------------------------------------------------------------------------
// Copy / paste
// -----------------------------------------------------------------------------

interface ClipRel {
  type: string;
  target: string;
  external: boolean;
  /** Image bytes, for pasting into another presentation. */
  bytes?: Uint8Array;
  uri?: string;
  resources?: ResourceBundle;
  drawingRel?: string;
}

export interface Clip {
  doc: PptxDocument;
  items: string[];
  rels: Map<string, ClipRel>;
}

/** Relationship types a pasted item may carry; anything else (charts, diagrams, media) is skipped. */
const PASTABLE = new Set(["image", "hyperlink", "chart", "diagramData", "diagramLayout", "diagramQuickStyle", "diagramColors", "diagramDrawing", "audio", "video", "media"]);

const relAttrs = (el: Element) =>
  [el, ...Array.from(el.getElementsByTagName("*"))].flatMap((e) => Array.from(e.attributes).filter((a) => a.namespaceURI === NS.r && a.value !== "").map((a) => ({ e, a })));

export async function copyItems(doc: PptxDocument, slidePath: string, items: Element[]): Promise<Clip> {
  const rels = new Map<string, ClipRel>();
  for (const item of items) {
    const dataId = item.getElementsByTagNameNS(NS.dgm,"relIds")[0]?.getAttributeNS(NS.r,"dm");
    const dataPath = dataId && doc.relTarget(slidePath,dataId);
    const drawingRel = dataPath && doc.xml(dataPath)?.getElementsByTagNameNS(NS.dsp,"dataModelExt")[0]?.getAttribute("relId");
    if (drawingRel && !rels.has(drawingRel)) {
      const rel = doc.rels(slidePath).find(r=>r.id===drawingRel);
      if(rel) rels.set(rel.id,{type:rel.type,uri:rel.uri,target:rel.target,external:rel.external,resources:await doc.captureResources([rel.target])});
    }
    for (const { a } of relAttrs(item)) {
      const rel = doc.rels(slidePath).find((r) => r.id === a.value);
      if (!rel || rels.has(rel.id)) continue;
      const bytes = rel.type === "image" && !rel.external ? await doc.media(rel.target).catch(() => undefined) : undefined;
      const resources = PASTABLE.has(rel.type) && rel.type !== "image" && !rel.external ? await doc.captureResources([rel.target]).catch(() => undefined) : undefined;
      rels.set(rel.id, { type: rel.type, uri: rel.uri, target: rel.target, external: rel.external, bytes, resources, drawingRel: rel.id === dataId ? drawingRel || undefined : undefined });
    }
  }
  return { doc, items: items.map((i) => new XMLSerializer().serializeToString(i)), rels };
}

/** Paste copied items onto a slide with new ids and relationships, shifted by `offset`. */
export function pasteItems(doc: PptxDocument, slidePath: string, clip: Clip, offset: { dx: number; dy: number }): { pasted: Element[]; skipped: number } {
  const slide = doc.xml(slidePath)!;
  const tree = slide.getElementsByTagNameNS(NS.p, "spTree")[0];
  const relMap = new Map<string, string>();
  const idMap = new Map<string, string>();
  const pasted: Element[] = [];
  let skipped = 0;

  const relFor = (oldId: string): string | null => {
    if (relMap.has(oldId)) return relMap.get(oldId)!;
    const rel = clip.rels.get(oldId);
    if (!rel || !PASTABLE.has(rel.type)) return null;
    let id: string;
    if (rel.resources) {
      const paths = doc.importResources(rel.resources);
      if(rel.drawingRel) {
        const drawing = relFor(rel.drawingRel);
        if (!drawing) return null;
        const data = doc.xml(paths.get(rel.target)!);
        data?.getElementsByTagNameNS(NS.dsp,"dataModelExt")[0]?.setAttribute("relId",drawing);
      }
      id = doc.addSlideRel(slidePath, rel.uri ?? rel.type, paths.get(rel.target)!);
    } else if (rel.external || (clip.doc === doc && rel.type === "image")) id = doc.addSlideRel(slidePath, rel.uri ?? rel.type, rel.target, rel.external);
    else if (rel.bytes) id = doc.addImage(slidePath, rel.bytes, rel.target.slice(rel.target.lastIndexOf(".") + 1));
    else return null;
    relMap.set(oldId, id);
    return id;
  };

  for (const xml of clip.items) {
    const item = slide.importNode(new DOMParser().parseFromString(xml, "application/xml").documentElement, true);
    const frames = item.localName === "graphicFrame" ? [item] : Array.from(item.getElementsByTagNameNS(NS.p, "graphicFrame"));
    const unsupported = frames.some((f) => !f.getElementsByTagNameNS(NS.a, "tbl").length && !relAttrs(f).length);
    const refs = relAttrs(item);
    const missing=refs.some(({a})=>{const rel=clip.rels.get(a.value);return !rel||!PASTABLE.has(rel.type)||(!rel.external&&!rel.resources&&rel.type!=="image")||(!rel.external&&rel.type==="image"&&clip.doc!==doc&&!rel.bytes);});
    if (unsupported || missing || refs.some(({ a }) => relFor(a.value) === null)) {
      skipped++;
      continue;
    }
    for (const { e, a } of refs) e.setAttributeNS(NS.r, a.name, relFor(a.value)!);
    for (const c of Array.from(item.getElementsByTagNameNS(NS.p, "cNvPr"))) {
      const id = String(nextShapeId(slide) + idMap.size);
      idMap.set(c.getAttribute("id") ?? "", id);
      c.setAttribute("id", id);
    }
    const r = itemRect(item);
    setItemRect(item, { ...r, x: r.x + offset.dx, y: r.y + offset.dy });
    tree.insertBefore(item, kid(tree, NS.p, "extLst"));
    pasted.push(item);
  }
  // Connectors pasted together with their shapes stay glued to the copies.
  for (const item of pasted) {
    for (const g of [...Array.from(item.getElementsByTagNameNS(NS.a, "stCxn")), ...Array.from(item.getElementsByTagNameNS(NS.a, "endCxn"))]) {
      const to = idMap.get(g.getAttribute("id") ?? "");
      if (to) g.setAttribute("id", to);
    }
  }
  return { pasted, skipped };
}
