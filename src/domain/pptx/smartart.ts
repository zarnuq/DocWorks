/**
 * SmartArt text: the items live in the diagram's data model (dgm:pt); PowerPoint also keeps a drawing of them
 * (the shapes we show), linked through the presentation points that stand for each item (presAssocID). Editing
 * an item's text changes both, so the slide shows it at once and PowerPoint agrees. Top-level items are added
 * and removed in the data model, and the drawing of a diagram laid out in a row or column is spread again over the
 * same space (PowerPoint redoes its own layout from the data).
 */
import { NS } from "../../core/utils/xml";

export interface SmartArtItem {
  id: string;
  text: string;
  /** 0 for top-level items, 1 for their children... */
  depth: number;
}

const kids = (el: Element | null, ns: string, name: string) => Array.from(el?.children ?? []).filter((c) => c.namespaceURI === ns && c.localName === name);
const kid = (el: Element | null, ns: string, name: string) => kids(el, ns, name)[0] ?? null;

const textOf = (body: Element | null) => kids(body, NS.a, "p").map((p) => Array.from(p.getElementsByTagNameNS(NS.a, "t")).map((t) => t.textContent).join("")).join("\n");

/** The diagram's items (its data points, not its connectors or presentation shapes), in order and by level. */
export function smartArtItems(data: Document): SmartArtItem[] {
  const points = new Map(Array.from(data.getElementsByTagNameNS(NS.dgm, "pt")).map((p) => [p.getAttribute("modelId") ?? "", p]));
  // Parent-of connections (no type, or "parOf"), by parent, in the parent's order.
  const children = new Map<string, { id: string; ord: number }[]>();
  for (const c of Array.from(data.getElementsByTagNameNS(NS.dgm, "cxn"))) {
    if ((c.getAttribute("type") ?? "parOf") !== "parOf") continue;
    const list = children.get(c.getAttribute("srcId") ?? "") ?? [];
    list.push({ id: c.getAttribute("destId") ?? "", ord: Number(c.getAttribute("srcOrd")) || 0 });
    children.set(c.getAttribute("srcId") ?? "", list);
  }
  const root = [...points.values()].find((p) => p.getAttribute("type") === "doc");
  const out: SmartArtItem[] = [];
  const walk = (id: string, depth: number) => {
    for (const { id: child } of (children.get(id) ?? []).sort((a, b) => a.ord - b.ord)) {
      const p = points.get(child);
      if (!p || !["node", null].includes(p.getAttribute("type"))) continue;
      out.push({ id: child, text: textOf(kid(p, NS.dgm, "t")), depth });
      walk(child, depth + 1);
    }
  };
  if (root) walk(root.getAttribute("modelId") ?? "", 0);
  return out;
}

/** Replace a text body's paragraphs with `text` (one per line), keeping the first paragraph's and run's formatting. */
function setBody(body: Element, text: string): void {
  const doc = body.ownerDocument;
  const paras = kids(body, NS.a, "p");
  const pPr = kid(paras[0] ?? null, NS.a, "pPr");
  const rPr = paras.flatMap((p) => kids(p, NS.a, "r")).map((r) => kid(r, NS.a, "rPr")).find(Boolean) ?? kid(paras[0] ?? null, NS.a, "endParaRPr");
  paras.forEach((p) => p.remove());
  for (const line of text.split("\n")) {
    const p = body.appendChild(doc.createElementNS(NS.a, "a:p"));
    if (pPr) p.appendChild(pPr.cloneNode(true));
    const r = p.appendChild(doc.createElementNS(NS.a, "a:r"));
    if (rPr) {
      const props = r.appendChild(doc.createElementNS(NS.a, "a:rPr"));
      for (const attr of Array.from(rPr.attributes)) props.setAttribute(attr.name, attr.value);
      props.append(...Array.from(rPr.childNodes).map((n) => n.cloneNode(true)));
    }
    r.appendChild(doc.createElementNS(NS.a, "a:t")).textContent = line;
  }
}

/** Set item `id`'s text in the data model and in the drawing shapes that show it. */
export function setSmartArtText(data: Document, drawing: Document | null, id: string, text: string): void {
  const point = Array.from(data.getElementsByTagNameNS(NS.dgm, "pt")).find((p) => p.getAttribute("modelId") === id);
  const body = kid(point ?? null, NS.dgm, "t");
  if (body) setBody(body, text);
  if (!drawing) return;
  // Presentation points standing for this item; the drawing's shapes carry their ids.
  const shownBy = new Set(Array.from(data.getElementsByTagNameNS(NS.dgm, "pt")).filter((p) => p.getAttribute("type") === "pres" && kid(p, NS.dgm, "prSet")?.getAttribute("presAssocID") === id).map((p) => p.getAttribute("modelId")));
  for (const sp of Array.from(drawing.getElementsByTagNameNS(NS.dsp, "sp"))) {
    const tx = kid(sp, NS.dsp, "txBody");
    if (tx && shownBy.has(sp.getAttribute("modelId"))) setBody(tx, text);
  }
}

/** The data model and drawing parts of the SmartArt in graphic frame `frame` on part `part`. */
export function smartArtParts(doc: { relTarget(part: string, rId: string): string | null; xml(path: string): Document | null }, frame: Element, part: string): { data: string; drawing: string | null } | null {
  const dm = frame.getElementsByTagNameNS(NS.dgm, "relIds")[0]?.getAttributeNS(NS.r, "dm");
  const data = dm ? doc.relTarget(part, dm) : null;
  if (!data) return null;
  const relId = doc.xml(data)?.getElementsByTagNameNS(NS.dsp, "dataModelExt")[0]?.getAttribute("relId");
  return { data, drawing: (relId && doc.relTarget(part, relId)) || null };
}


// -----------------------------------------------------------------------------
// Adding and removing items
// -----------------------------------------------------------------------------

const guid = () => `{${crypto.randomUUID().toUpperCase()}}`;
const pointsOf = (data: Document) => Array.from(data.getElementsByTagNameNS(NS.dgm, "pt"));
const cxnsOf = (data: Document) => Array.from(data.getElementsByTagNameNS(NS.dgm, "cxn"));
const isParOf = (c: Element) => (c.getAttribute("type") ?? "parOf") === "parOf";
const parOf = (data: Document, id: string) => cxnsOf(data).find((c) => isParOf(c) && c.getAttribute("destId") === id) ?? null;
const ptById = (data: Document, id: string | null) => pointsOf(data).find((p) => p.getAttribute("modelId") === id) ?? null;

interface Box { x: number; y: number; w: number; h: number }
/** The shape's own transform and its text box, both moved together. */
const xfrmsOf = (sp: Element) => [kid(kid(sp, NS.dsp, "spPr"), NS.a, "xfrm"), kid(sp, NS.dsp, "txXfrm")].filter((x): x is Element => !!x);
function boxOf(x: Element): Box {
  const off = kid(x, NS.a, "off"), ext = kid(x, NS.a, "ext");
  return { x: Number(off?.getAttribute("x") ?? 0), y: Number(off?.getAttribute("y") ?? 0), w: Number(ext?.getAttribute("cx") ?? 0), h: Number(ext?.getAttribute("cy") ?? 0) };
}
function setBox(x: Element, b: Box): void {
  kid(x, NS.a, "off")?.setAttribute("x", String(Math.round(b.x)));
  kid(x, NS.a, "off")?.setAttribute("y", String(Math.round(b.y)));
  kid(x, NS.a, "ext")?.setAttribute("cx", String(Math.round(b.w)));
  kid(x, NS.a, "ext")?.setAttribute("cy", String(Math.round(b.h)));
}
const union = (boxes: Box[]): Box => {
  const x = Math.min(...boxes.map((b) => b.x)), y = Math.min(...boxes.map((b) => b.y));
  return { x, y, w: Math.max(...boxes.map((b) => b.x + b.w)) - x, h: Math.max(...boxes.map((b) => b.y + b.h)) - y };
};

/** Drawing shapes standing for data point `id` (through the presentation points associated with it). */
function shapesFor(data: Document, drawing: Document, id: string | null): Element[] {
  if (!id) return [];
  const pres = new Set(pointsOf(data).filter((p) => p.getAttribute("type") === "pres" && kid(p, NS.dgm, "prSet")?.getAttribute("presAssocID") === id).map((p) => p.getAttribute("modelId")));
  return Array.from(drawing.getElementsByTagNameNS(NS.dsp, "sp")).filter((sp) => pres.has(sp.getAttribute("modelId")));
}

/** One top-level item's shapes, and the shapes of the connector after it (an arrow in a process). */
interface Group { shapes: Element[]; gap: Element[]; box: Box }

/** The diagram's top-level items with their drawing, refusing layouts that aren't one row or one column. */
function rowOf(data: Document, drawing: Document | null): { root: string; items: string[]; groups: Group[]; across: boolean } {
  const doc = pointsOf(data).find((p) => p.getAttribute("type") === "doc")?.getAttribute("modelId") ?? "";
  const top = cxnsOf(data).filter((c) => isParOf(c) && c.getAttribute("srcId") === doc).sort((a, b) => Number(a.getAttribute("srcOrd")) - Number(b.getAttribute("srcOrd")));
  const items = top.map((c) => c.getAttribute("destId") ?? "");
  if (!drawing) return { root: doc, items, groups: [], across: true };
  const groups = top.map((c): Group => {
    const shapes = shapesFor(data, drawing, c.getAttribute("destId"));
    return { shapes, gap: shapesFor(data, drawing, c.getAttribute("sibTransId")), box: union(shapes.flatMap(xfrmsOf).map(boxOf)) };
  });
  if (groups.some((g) => !g.shapes.length)) throw new Error("This diagram's items can only be added or removed in a row or column layout.");
  const cx = groups.map((g) => g.box.x + g.box.w / 2), cy = groups.map((g) => g.box.y + g.box.h / 2);
  const spread = (v: number[]) => Math.max(...v) - Math.min(...v);
  const across = spread(cx) >= spread(cy);
  // One row (or column): items the same size, in order along it, lined up across it.
  const size = (g: Group) => (across ? g.box.w : g.box.h), cross = across ? cy : cx;
  const crossSize = across ? groups[0].box.h : groups[0].box.w;
  const ok = groups.every((g, i) => Math.abs(size(g) - size(groups[0])) <= size(groups[0]) * 0.05 && Math.abs(cross[i] - cross[0]) <= crossSize * 0.05 && (i === 0 || (across ? cx : cy)[i] > (across ? cx : cy)[i - 1]));
  if (!ok) throw new Error("This diagram's items can only be added or removed in a row or column layout.");
  return { root: doc, items, groups, across };
}

/** Spread `groups` (in their new order) over the span the old row covered, scaled to fit. */
function relayout(old: Group[], groups: Group[], across: boolean): void {
  if (!old.length) return;
  const pos = (b: Box) => (across ? b.x : b.y), len = (b: Box) => (across ? b.w : b.h);
  const start = pos(old[0].box), end = pos(old[old.length - 1].box) + len(old[old.length - 1].box);
  const w = len(old[0].box);
  const gap = old.length > 1 ? pos(old[1].box) - pos(old[0].box) - w : w * 0.15;
  const n = groups.length;
  // Items shrink to fit more; with fewer they keep their size and the row is centred (growing would overflow).
  const s = Math.min(1, (end - start) / (n + ((n - 1) * gap) / w) / w), w2 = w * s, gap2 = gap * s;
  const first = start + (end - start - (n * w2 + (n - 1) * gap2)) / 2;
  const move = (b: Box, from: Box, at: number): Box => {
    // Along the row: from the item's start; across it: scaled about the item's middle line.
    const mid = across ? from.y + from.h / 2 : from.x + from.w / 2;
    const along = at + (pos(b) - pos(from)) * s, side = mid + ((across ? b.y : b.x) - mid) * s;
    return across ? { x: along, y: side, w: b.w * s, h: b.h * s } : { x: side, y: along, w: b.w * s, h: b.h * s };
  };
  groups.forEach((g, j) => {
    const at = first + j * (w2 + gap2);
    const from = g.box;
    for (const sp of g.shapes) for (const x of xfrmsOf(sp)) setBox(x, move(boxOf(x), from, at));
    // The connector sits in the gap after the item, where it sat relative to the old gap.
    for (const sp of g.gap)
      for (const x of xfrmsOf(sp)) setBox(x, move(boxOf(x), { ...from, ...(across ? { x: pos(from) + w + gap / 2 } : { y: pos(from) + w + gap / 2 }) }, at + w2 + gap2 / 2));
    if (s < 1)
      for (const sp of [...g.shapes, ...g.gap])
        for (const r of Array.from(sp.getElementsByTagNameNS(NS.a, "*")).filter((e) => e.hasAttribute("sz") && ["rPr", "endParaRPr", "defRPr"].includes(e.localName)))
          r.setAttribute("sz", String(Math.max(100, Math.round((Number(r.getAttribute("sz")) * s) / 100) * 100)));
  });
}

/** Clone a drawing shape with a fresh model id, right after `after`. */
function cloneShape(sp: Element, after: Element): Element {
  const copy = sp.cloneNode(true) as Element;
  copy.setAttribute("modelId", guid());
  after.after(copy);
  return copy;
}

/** Add a top-level item after `afterId`; returns its id. */
export function addSmartArtItem(data: Document, drawing: Document | null, afterId: string, text: string): string {
  const { root, items, groups, across } = rowOf(data, drawing);
  const k = items.indexOf(afterId);
  const cxn = parOf(data, afterId);
  const node = ptById(data, afterId);
  if (k < 0 || !cxn || !node) throw new Error("Only top-level items can be added to.");
  const id = guid(), cxnId = guid(), parT = guid(), sibT = guid();
  // Data: the item, its two transition points, and its connection, ordered after `afterId`.
  const pt = node.cloneNode(true) as Element;
  pt.setAttribute("modelId", id);
  const body = kid(pt, NS.dgm, "t");
  if (body) setBody(body, text);
  const trans = (from: string | null, newId: string) => {
    const t = ptById(data, from);
    const copy = (t?.cloneNode(true) as Element | undefined) ?? data.createElementNS(NS.dgm, "dgm:pt");
    copy.setAttribute("modelId", newId);
    copy.setAttribute("type", from === cxn.getAttribute("parTransId") ? "parTrans" : "sibTrans");
    copy.setAttribute("cxnId", cxnId);
    return copy;
  };
  const anchor = ptById(data, cxn.getAttribute("sibTransId")) ?? node;
  anchor.after(pt, trans(cxn.getAttribute("parTransId"), parT), trans(cxn.getAttribute("sibTransId"), sibT));
  const ord = Number(cxn.getAttribute("srcOrd")) || 0;
  for (const c of cxnsOf(data)) if (isParOf(c) && c.getAttribute("srcId") === root && Number(c.getAttribute("srcOrd")) > ord) c.setAttribute("srcOrd", String(Number(c.getAttribute("srcOrd")) + 1));
  const link = cxn.cloneNode(true) as Element;
  for (const [a, v] of [["modelId", cxnId], ["destId", id], ["srcOrd", String(ord + 1)], ["parTransId", parT], ["sibTransId", sibT]]) link.setAttribute(a, v);
  cxn.after(link);
  // Presentation points (and their links) for the new item and its connector, copied from the item's, so its
  // drawing shapes stand for points in the data like every other shape.
  const map = new Map<string, string>([[afterId, id], [cxn.getAttribute("parTransId") ?? "", parT], [cxn.getAttribute("sibTransId") ?? "", sibT]]);
  const clonePres = (assoc: string, to: string) => {
    for (const p of pointsOf(data).filter((p) => p.getAttribute("type") === "pres" && kid(p, NS.dgm, "prSet")?.getAttribute("presAssocID") === assoc)) {
      const copy = p.cloneNode(true) as Element;
      const fresh = guid();
      map.set(p.getAttribute("modelId") ?? "", fresh);
      copy.setAttribute("modelId", fresh);
      kid(copy, NS.dgm, "prSet")?.setAttribute("presAssocID", to);
      p.parentElement!.appendChild(copy);
    }
  };
  clonePres(afterId, id);
  clonePres(cxn.getAttribute("sibTransId") ?? "", sibT);
  const linkPres = (only: Set<string>) => {
    for (const c of cxnsOf(data).filter((c) => !isParOf(c) && (only.has(c.getAttribute("srcId") ?? "") || only.has(c.getAttribute("destId") ?? "")))) {
      const copy = c.cloneNode(true) as Element;
      copy.setAttribute("modelId", guid());
      for (const a of ["srcId", "destId"]) copy.setAttribute(a, map.get(c.getAttribute(a) ?? "") ?? c.getAttribute(a) ?? "");
      c.parentElement!.appendChild(copy);
    }
  };
  linkPres(new Set(map.keys()));
  if (!drawing) return id;
  const shapeCopy = (sp: Element, after: Element) => {
    const copy = cloneShape(sp, after);
    const to = map.get(sp.getAttribute("modelId") ?? "");
    if (to) copy.setAttribute("modelId", to);
    return copy;
  };
  // Drawing: copies of the item's shapes; whoever is now followed by an item has a connector after it.
  const src = groups[k];
  const shapes: Element[] = [];
  let last = src.gap[src.gap.length - 1] ?? src.shapes[src.shapes.length - 1];
  for (const sp of src.shapes) shapes.push((last = shapeCopy(sp, last)));
  for (const sp of shapes) {
    const tx = kid(sp, NS.dsp, "txBody");
    if (tx) setBody(tx, text);
  }
  const added: Group = { shapes, gap: [], box: src.box };
  if (src.gap.length) for (const sp of src.gap) added.gap.push((last = shapeCopy(sp, last)));
  else if (k > 0) {
    // The last item gains the connector the one before it has, standing for its own (so far undrawn) transition.
    const prev = groups[k - 1];
    const prevSib = parOf(data, items[k - 1])?.getAttribute("sibTransId") ?? "";
    const sib = cxn.getAttribute("sibTransId") ?? "";
    const presOf = (assoc: string) => pointsOf(data).filter((p) => p.getAttribute("type") === "pres" && kid(p, NS.dgm, "prSet")?.getAttribute("presAssocID") === assoc);
    map.clear();
    const own = presOf(sib);
    if (own.length) {
      // Its transition already has presentation points (PowerPoint keeps them): the shapes stand for those.
      const name = (p: Element) => kid(p, NS.dgm, "prSet")?.getAttribute("presName");
      for (const p of presOf(prevSib)) {
        const match = own.find((o) => name(o) === name(p));
        if (match) map.set(p.getAttribute("modelId") ?? "", match.getAttribute("modelId") ?? "");
      }
    } else {
      clonePres(prevSib, sib);
      linkPres(new Set(map.keys()));
    }
    src.gap = prev.gap.map((sp) => shapeCopy(sp, src.shapes[src.shapes.length - 1]));
    for (const sp of src.gap) for (const x of xfrmsOf(sp)) setBox(x, { ...boxOf(x), ...(across ? { x: boxOf(x).x + (src.box.x - prev.box.x) } : { y: boxOf(x).y + (src.box.y - prev.box.y) }) });
  }
  const next = [...groups];
  next.splice(k + 1, 0, added);
  relayout(groups, next, across);
  return id;
}

/** Remove a top-level item, with its children, connections and shapes. */
export function removeSmartArtItem(data: Document, drawing: Document | null, id: string): void {
  const { root, items, groups, across } = rowOf(data, drawing);
  const k = items.indexOf(id);
  if (k < 0) throw new Error("Only top-level items can be removed.");
  if (items.length < 2) throw new Error("A diagram needs at least one item.");
  // The item and everything under it, with their transition points.
  const gone = new Set<string>();
  const collect = (at: string) => {
    gone.add(at);
    for (const c of cxnsOf(data)) {
      if (!isParOf(c)) continue;
      if (c.getAttribute("destId") === at) for (const a of ["modelId", "parTransId", "sibTransId"]) gone.add(c.getAttribute(a) ?? "");
      if (c.getAttribute("srcId") === at) collect(c.getAttribute("destId") ?? "");
    }
  };
  collect(id);
  const ord = Number(parOf(data, id)?.getAttribute("srcOrd")) || 0;
  // Presentation points standing for them go too, and every connection touching any of these.
  for (const p of pointsOf(data)) if (p.getAttribute("type") === "pres" && gone.has(kid(p, NS.dgm, "prSet")?.getAttribute("presAssocID") ?? "")) gone.add(p.getAttribute("modelId") ?? "");
  if (drawing) {
    const g = groups[k];
    for (const sp of [...g.shapes, ...g.gap]) sp.remove();
    // A new last item drops its connector.
    if (k === groups.length - 1 && k > 0) {
      for (const sp of groups[k - 1].gap) sp.remove();
      groups[k - 1].gap = [];
    }
  }
  for (const p of pointsOf(data)) if (gone.has(p.getAttribute("modelId") ?? "")) p.remove();
  for (const c of cxnsOf(data)) if (gone.has(c.getAttribute("modelId") ?? "") || gone.has(c.getAttribute("srcId") ?? "") || gone.has(c.getAttribute("destId") ?? "")) c.remove();
  for (const c of cxnsOf(data)) if (isParOf(c) && c.getAttribute("srcId") === root && Number(c.getAttribute("srcOrd")) > ord) c.setAttribute("srcOrd", String(Number(c.getAttribute("srcOrd")) - 1));
  if (drawing) relayout(groups, groups.filter((_, i) => i !== k), across);
}
