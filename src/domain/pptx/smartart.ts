/**
 * SmartArt text: the items live in the diagram's data model (dgm:pt); PowerPoint also keeps a drawing of them
 * (the shapes we show), linked through the presentation points that stand for each item (presAssocID). Editing
 * an item's text changes both, so the slide shows it at once and PowerPoint agrees. Items are added and removed
 * in the data model, and the drawing is laid out again in the space it had: a row or column is spread again, a
 * cycle spaced round its circle, a top-down tree (a hierarchy) redrawn level by level with its connecting lines,
 * the items round a radial diagram's middle spread over their arc.
 * Which of these a drawing is comes from where its shapes are, not from the layout's name. PowerPoint redoes its
 * own layout from the data.
 */
import { NS, kid, kids, guid } from "../../core/utils/xml";

interface SmartArtItem {
  id: string;
  text: string;
  /** 0 for top-level items, 1 for their children... */
  depth: number;
}


const textOf = (body: Element | null) => kids(body, "p", NS.a).map((p) => Array.from(p.getElementsByTagNameNS(NS.a, "t")).map((t) => t.textContent).join("")).join("\n");

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
      out.push({ id: child, text: textOf(kid(p, "t", NS.dgm)), depth });
      walk(child, depth + 1);
    }
  };
  if (root) walk(root.getAttribute("modelId") ?? "", 0);
  return out;
}

/** Replace a text body's paragraphs with `text` (one per line), keeping the first paragraph's and run's formatting. */
function setBody(body: Element, text: string): void {
  const doc = body.ownerDocument;
  const paras = kids(body, "p", NS.a);
  const pPr = kid(paras[0] ?? null, "pPr", NS.a);
  const rPr = paras.flatMap((p) => kids(p, "r", NS.a)).map((r) => kid(r, "rPr", NS.a)).find(Boolean) ?? kid(paras[0] ?? null, "endParaRPr", NS.a);
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
  const body = kid(point ?? null, "t", NS.dgm);
  if (body) setBody(body, text);
  if (!drawing) return;
  // Presentation points standing for this item; the drawing's shapes carry their ids.
  const shownBy = new Set(Array.from(data.getElementsByTagNameNS(NS.dgm, "pt")).filter((p) => p.getAttribute("type") === "pres" && kid(p, "prSet", NS.dgm)?.getAttribute("presAssocID") === id).map((p) => p.getAttribute("modelId")));
  for (const sp of Array.from(drawing.getElementsByTagNameNS(NS.dsp, "sp"))) {
    const tx = kid(sp, "txBody", NS.dsp);
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

const pointsOf = (data: Document) => Array.from(data.getElementsByTagNameNS(NS.dgm, "pt"));
const cxnsOf = (data: Document) => Array.from(data.getElementsByTagNameNS(NS.dgm, "cxn"));
const isParOf = (c: Element) => (c.getAttribute("type") ?? "parOf") === "parOf";
const ptById = (data: Document, id: string | null) => pointsOf(data).find((p) => p.getAttribute("modelId") === id) ?? null;

interface Box { x: number; y: number; w: number; h: number }
/** The shape's own transform and its text box, both moved together. */
const xfrmsOf = (sp: Element) => [kid(kid(sp, "spPr", NS.dsp), "xfrm", NS.a), kid(sp, "txXfrm", NS.dsp)].filter((x): x is Element => !!x);
function boxOf(x: Element): Box {
  const off = kid(x, "off", NS.a), ext = kid(x, "ext", NS.a);
  return { x: Number(off?.getAttribute("x") ?? 0), y: Number(off?.getAttribute("y") ?? 0), w: Number(ext?.getAttribute("cx") ?? 0), h: Number(ext?.getAttribute("cy") ?? 0) };
}
function setBox(x: Element, b: Box): void {
  kid(x, "off", NS.a)?.setAttribute("x", String(Math.round(b.x)));
  kid(x, "off", NS.a)?.setAttribute("y", String(Math.round(b.y)));
  kid(x, "ext", NS.a)?.setAttribute("cx", String(Math.round(b.w)));
  kid(x, "ext", NS.a)?.setAttribute("cy", String(Math.round(b.h)));
}
const union = (boxes: Box[]): Box => {
  const x = Math.min(...boxes.map((b) => b.x)), y = Math.min(...boxes.map((b) => b.y));
  return { x, y, w: Math.max(...boxes.map((b) => b.x + b.w)) - x, h: Math.max(...boxes.map((b) => b.y + b.h)) - y };
};

/** Drawing shapes standing for data point `id` (through the presentation points associated with it). */
function shapesFor(data: Document, drawing: Document, id: string | null): Element[] {
  if (!id) return [];
  const pres = new Set(pointsOf(data).filter((p) => p.getAttribute("type") === "pres" && kid(p, "prSet", NS.dgm)?.getAttribute("presAssocID") === id).map((p) => p.getAttribute("modelId")));
  return Array.from(drawing.getElementsByTagNameNS(NS.dsp, "sp")).filter((sp) => pres.has(sp.getAttribute("modelId")));
}

/** How the drawing is laid out, which decides how it is laid out again after a change. */
export type SmartArtLayout = "row" | "cycle" | "tree" | "radial";
const REFUSED = "This diagram's items can only be added or removed in a row or column, a cycle, a top-down tree or round a middle item.";
const RING_ONLY = "Only the items around the middle can be added or removed.";

/** An item of the data model with the drawing shapes that stand for it. */
interface Item {
  id: string;
  /** Its connection from its parent (the document point for top-level items). */
  cxn: Element;
  parent: Item | null;
  children: Item[];
  depth: number;
  shapes: Element[];
  /** The connector after it (its sibling transition: an arrow in a process or cycle). */
  gap: Element[];
  /** The line from its parent (its parent transition: a hierarchy's connector). */
  conn: Element[];
  /** What its own shapes cover. */
  box: Box;
}
const NO_BOX: Box = { x: 0, y: 0, w: 0, h: 0 };
const ordOf = (c: Element) => Number(c.getAttribute("srcOrd")) || 0;
const flat = (items: Item[]): Item[] => items.flatMap((i) => [i, ...flat(i.children)]);
const mid = (b: Box) => ({ x: b.x + b.w / 2, y: b.y + b.h / 2 });

/** The data model's items as a tree, each with its shapes in the drawing. */
function readItems(data: Document, drawing: Document | null): Item[] {
  const root = pointsOf(data).find((p) => p.getAttribute("type") === "doc")?.getAttribute("modelId") ?? "";
  const byParent = new Map<string, Element[]>();
  for (const c of cxnsOf(data)) if (isParOf(c)) byParent.set(c.getAttribute("srcId") ?? "", [...(byParent.get(c.getAttribute("srcId") ?? "") ?? []), c]);
  const build = (cxn: Element, parent: Item | null, depth: number): Item | null => {
    const id = cxn.getAttribute("destId") ?? "";
    const point = ptById(data, id);
    // Assistants (an organisation chart's boxes between the levels) count: they move with their parent in a row, and
    // a tree holding one doesn't pass for a plain tree, so it is never laid out again around them.
    if (!point || !["node", "asst", null].includes(point.getAttribute("type"))) return null;
    const of = (assoc: string | null) => (drawing ? shapesFor(data, drawing, assoc) : []);
    const shapes = of(id);
    const item: Item = { id, cxn, parent, children: [], depth, shapes, gap: of(cxn.getAttribute("sibTransId")), conn: of(cxn.getAttribute("parTransId")), box: shapes.length ? union(shapes.flatMap(xfrmsOf).map(boxOf)) : NO_BOX };
    item.children = childrenOf(id, item, depth + 1);
    return item;
  };
  const childrenOf = (id: string, parent: Item | null, depth: number) =>
    (byParent.get(id) ?? []).sort((a, b) => ordOf(a) - ordOf(b)).map((c) => build(c, parent, depth)).filter((i): i is Item => !!i);
  return childrenOf(root, null, 0);
}

/** A tree drawn top-down: every item drawn, each level in one row, children wholly below their parents. */
function isTree(top: Item[]): boolean {
  const all = flat(top);
  if (!all.some((i) => i.depth > 0 && i.shapes.length) || all.some((i) => !i.shapes.length)) return false;
  const first = new Map<number, Item>();
  return all.every((i) => {
    const level = first.get(i.depth) ?? first.set(i.depth, i).get(i.depth)!;
    const below = !i.parent || i.box.y >= i.parent.box.y + i.parent.box.h - i.parent.box.h * 0.05;
    return below && Math.abs(mid(i.box).y - mid(level.box).y) <= level.box.h * 0.1;
  });
}

/** Top-level items in one row or column: in order along it, lined up across it (their lengths may differ). */
function rowAcross(top: Item[]): boolean | null {
  const c = top.map((g) => mid(g.box));
  const spread = (v: number[]) => Math.max(...v) - Math.min(...v);
  const across = spread(c.map((p) => p.x)) >= spread(c.map((p) => p.y));
  const along = c.map((p) => (across ? p.x : p.y)), cross = c.map((p) => (across ? p.y : p.x));
  const crossSize = across ? top[0].box.h : top[0].box.w;
  const ok = top.every((_, i) => Math.abs(cross[i] - cross[0]) <= crossSize * 0.05 && (i === 0 || along[i] > along[i - 1]));
  return ok ? across : null;
}

/** Top-level items round a circle: the last has a connector back to the first, and all are as far from the middle. */
function isCycle(top: Item[]): boolean {
  if (top.length < 2 || !top[top.length - 1].gap.length) return false;
  const c = top.map((g) => mid(g.box));
  const cx = c.reduce((a, p) => a + p.x, 0) / c.length, cy = c.reduce((a, p) => a + p.y, 0) / c.length;
  const r = c.map((p) => Math.hypot(p.x - cx, p.y - cy));
  return Math.max(...r) - Math.min(...r) <= Math.max(...r) * 0.1;
}

/**
 * The first item in the middle with its children round it, all as far from it (a radial diagram); PowerPoint draws
 * no other top-level item.
 */
function isRadial(top: Item[]): boolean {
  const [centre, ...others] = top;
  if (!centre?.shapes.length || !centre.children.length || others.some((i) => i.shapes.length) || centre.children.some((c) => !c.shapes.length || c.children.some((g) => g.shapes.length))) return false;
  const c = mid(centre.box), r = centre.children.map((i) => Math.hypot(mid(i.box).x - c.x, mid(i.box).y - c.y));
  return Math.min(...r) > 0 && Math.max(...r) - Math.min(...r) <= Math.max(...r) * 0.15;
}

function layoutOf(top: Item[], drawing: Document | null): SmartArtLayout {
  if (!drawing) return "row";
  if (isTree(top)) return "tree";
  if (isRadial(top)) return "radial";
  if (!top.length || top.some((g) => !g.shapes.length)) throw new Error(REFUSED);
  if (isCycle(top)) return "cycle";
  if (rowAcross(top) === null) throw new Error(REFUSED);
  return "row";
}

/** How the diagram is drawn, or null when its items can't be added or removed here. */
export function smartArtLayout(data: Document, drawing: Document | null): SmartArtLayout | null {
  try {
    return layoutOf(readItems(data, drawing), drawing);
  } catch {
    return null;
  }
}

/** One top-level item of a row or cycle: its shapes, the connector after it, and what moves along (its children's shapes). */
interface Group { shapes: Element[]; gap: Element[]; extra: Element[]; box: Box }
const groupOf = (i: Item): Group => ({ shapes: i.shapes, gap: i.gap, extra: flat(i.children).flatMap((c) => [...c.shapes, ...c.gap, ...c.conn]), box: i.box });

/** Text sizes in `shapes` scaled by `s` (to whole points). */
function scaleText(shapes: Element[], s: number): void {
  if (s >= 1) return;
  for (const sp of shapes)
    for (const r of Array.from(sp.getElementsByTagNameNS(NS.a, "*")).filter((e) => e.hasAttribute("sz") && ["rPr", "endParaRPr", "defRPr"].includes(e.localName)))
      r.setAttribute("sz", String(Math.max(100, Math.round((Number(r.getAttribute("sz")) * s) / 100) * 100)));
}

/** Spread `groups` (in their new order) over the span the old row covered, scaled to fit. */
function relayout(old: Group[], groups: Group[], across: boolean): void {
  if (!old.length) return;
  const pos = (b: Box) => (across ? b.x : b.y), len = (b: Box) => (across ? b.w : b.h);
  const start = pos(old[0].box), end = pos(old[old.length - 1].box) + len(old[old.length - 1].box);
  // The gap between neighbours (negative when they overlap, as a process's arrows do).
  const gap = old.length > 1 ? pos(old[1].box) - pos(old[0].box) - len(old[0].box) : len(old[0].box) * 0.15;
  const n = groups.length, total = groups.reduce((a, g) => a + len(g.box), 0) + (n - 1) * gap;
  // Items shrink to fit more; with fewer they keep their size and the row is centred (growing would overflow).
  const s = Math.min(1, (end - start) / total), gap2 = gap * s;
  let at = start + (end - start - total * s) / 2;
  const move = (b: Box, from: Box, at: number): Box => {
    // Along the row: from the item's start; across it: scaled about the item's middle line.
    const middle = across ? from.y + from.h / 2 : from.x + from.w / 2;
    const along = at + (pos(b) - pos(from)) * s, side = middle + ((across ? b.y : b.x) - middle) * s;
    return across ? { x: along, y: side, w: b.w * s, h: b.h * s } : { x: side, y: along, w: b.w * s, h: b.h * s };
  };
  for (const g of groups) {
    const from = g.box, w = len(from);
    for (const sp of [...g.shapes, ...g.extra]) for (const x of xfrmsOf(sp)) setBox(x, move(boxOf(x), from, at));
    // The connector sits in the gap after the item, where it sat relative to the old gap.
    for (const sp of g.gap)
      for (const x of xfrmsOf(sp)) setBox(x, move(boxOf(x), { ...from, ...(across ? { x: pos(from) + w + gap / 2 } : { y: pos(from) + w + gap / 2 }) }, at + w * s + gap2 / 2));
    scaleText([...g.shapes, ...g.gap, ...g.extra], s);
    at += w * s + gap2;
  }
}

/**
 * Space `groups` (in their new order) evenly round the circle the old ones were on, from where the first was.
 * Items shrink when their neighbours come closer; each connector goes half-way to the next item, turned as far.
 */
function relayoutCycle(old: Group[], groups: Group[]): void {
  const c = old.map((g) => mid(g.box));
  const C = { x: c.reduce((a, p) => a + p.x, 0) / c.length, y: c.reduce((a, p) => a + p.y, 0) / c.length };
  const R = c.reduce((a, p) => a + Math.hypot(p.x - C.x, p.y - C.y), 0) / c.length;
  const angle = (p: { x: number; y: number }) => Math.atan2(p.y - C.y, p.x - C.x);
  const a0 = angle(c[0]);
  // The way round the old items go (clockwise on screen unless the second one says otherwise).
  const turn = old.length > 2 && (((angle(c[1]) - a0) % (2 * Math.PI)) + 2 * Math.PI) % (2 * Math.PI) > Math.PI ? -1 : 1;
  const step0 = (turn * 2 * Math.PI) / old.length, step = (turn * 2 * Math.PI) / groups.length;
  // One item alone has no neighbour to make room for.
  const s = groups.length < 2 ? 1 : Math.min(1, Math.sin(Math.PI / groups.length) / Math.sin(Math.PI / old.length));
  groups.forEach((g, j) => {
    const from = mid(g.box), a = a0 + j * step;
    const to = { x: C.x + R * Math.cos(a), y: C.y + R * Math.sin(a) };
    for (const sp of [...g.shapes, ...g.extra])
      for (const x of xfrmsOf(sp)) {
        const b = boxOf(x), m = mid(b);
        setBox(x, { x: to.x + (m.x - from.x) * s - (b.w * s) / 2, y: to.y + (m.y - from.y) * s - (b.h * s) / 2, w: b.w * s, h: b.h * s });
      }
    // From the middle of the old arc after the item to the middle of the new one.
    const d = a + step / 2 - (angle(from) + step0 / 2);
    for (const sp of g.gap)
      for (const x of xfrmsOf(sp)) {
        const b = boxOf(x), m = mid(b), v = { x: m.x - C.x, y: m.y - C.y };
        // An arc drawn round the circle's own middle turns in place and keeps its size.
        const k = Math.hypot(v.x, v.y) < R * 0.2 ? 1 : s;
        const at = { x: C.x + v.x * Math.cos(d) - v.y * Math.sin(d), y: C.y + v.x * Math.sin(d) + v.y * Math.cos(d) };
        setBox(x, { x: at.x - (b.w * k) / 2, y: at.y - (b.h * k) / 2, w: b.w * k, h: b.h * k });
        const rot = Math.round(Number(x.getAttribute("rot") ?? 0) + (d * 180 * 60000) / Math.PI);
        x.setAttribute("rot", String(((rot % 21600000) + 21600000) % 21600000));
      }
    scaleText([...g.shapes, ...g.gap, ...g.extra], s);
  });
}

/**
 * Space `now` (the middle item's children, in order) over the arc `old` covered round `centre`: its two ends stay
 * (a full ring stays full), each child's line from the middle turns with it, and children shrink only when
 * neighbours would touch.
 */
function relayoutArc(centre: Item, old: Item[], now: Item[]): void {
  const C = mid(centre.box);
  const angle = (i: Item) => Math.atan2(mid(i.box).y - C.y, mid(i.box).x - C.x);
  const R = old.reduce((a, i) => a + Math.hypot(mid(i.box).x - C.x, mid(i.box).y - C.y), 0) / old.length;
  // The old angles in order, each within half a turn of the one before.
  const a: number[] = [];
  for (const i of old) {
    let v = angle(i);
    while (a.length && v - a[a.length - 1] > Math.PI) v -= 2 * Math.PI;
    while (a.length && v - a[a.length - 1] < -Math.PI) v += 2 * Math.PI;
    a.push(v);
  }
  const span = a[a.length - 1] - a[0], n = now.length;
  const full = old.length === 1 || Math.abs(Math.abs(span) - (2 * Math.PI * (old.length - 1)) / old.length) < 0.2;
  const step = full ? ((span < 0 ? -1 : 1) * 2 * Math.PI) / n : n > 1 ? span / (n - 1) : 0;
  const first = full || n > 1 ? a[0] : a[0] + span / 2;
  const size = Math.max(...now.map((i) => Math.max(i.box.w, i.box.h)));
  const s = n > 1 ? Math.min(1, (2 * R * Math.abs(Math.sin(step / 2))) / (size * 1.05)) : 1;
  now.forEach((i, j) => {
    const from = mid(i.box), to = { x: C.x + R * Math.cos(first + j * step), y: C.y + R * Math.sin(first + j * step) };
    for (const sp of [...i.shapes, ...i.gap])
      for (const x of xfrmsOf(sp)) {
        const b = boxOf(x), m = mid(b);
        setBox(x, { x: to.x + (m.x - from.x) * s - (b.w * s) / 2, y: to.y + (m.y - from.y) * s - (b.h * s) / 2, w: b.w * s, h: b.h * s });
      }
    const d = first + j * step - angle(i);
    for (const sp of i.conn)
      for (const x of xfrmsOf(sp)) {
        const b = boxOf(x), v = { x: mid(b).x - C.x, y: mid(b).y - C.y };
        setBox(x, { x: C.x + v.x * Math.cos(d) - v.y * Math.sin(d) - b.w / 2, y: C.y + v.x * Math.sin(d) + v.y * Math.cos(d) - b.h / 2, w: b.w, h: b.h });
        const rot = Math.round(Number(x.getAttribute("rot") ?? 0) + (d * 180 * 60000) / Math.PI);
        x.setAttribute("rot", String(((rot % 21600000) + 21600000) % 21600000));
      }
    scaleText([...i.shapes, ...i.gap], s);
  });
}

/** What a tree's layout keeps: the space it has, each level's top and height, and the gap between neighbours. */
interface TreePlan { bounds: Box; levels: { y: number; h: number }[]; gap: number }
function treePlan(top: Item[]): TreePlan {
  const all = flat(top);
  const levels: TreePlan["levels"] = [];
  for (const i of all) if (!levels[i.depth] || i.box.y < levels[i.depth].y) levels[i.depth] = { y: i.box.y, h: i.box.h };
  // Between neighbours with nothing under them where there are any (others are pushed apart by their children).
  const gaps = (leaves: boolean) =>
    [top, ...all.map((i) => i.children)].flatMap((list) => list.slice(1).flatMap((n, k) => (!leaves || (!n.children.length && !list[k].children.length) ? [n.box.x - list[k].box.x - list[k].box.w] : []))).filter((g) => g > 0);
  const gap = Math.min(...(gaps(true).length ? gaps(true) : gaps(false).length ? gaps(false) : [all[0].box.w * 0.25]));
  return { bounds: union(all.map((i) => i.box)), levels, gap };
}

/** Draw a hierarchy's connector from (px, py), its parent's bottom middle, to (cx, cy), the item's top middle. */
function setConnector(sp: Element, px: number, py: number, cx: number, cy: number): void {
  const h = Math.max(1, Math.round(cy - py)), straight = Math.abs(cx - px) < 1;
  const x0 = Math.min(px, cx);
  // A straight line still needs a box with some width (PowerPoint's is 91440 wide, the line down its middle).
  const box: Box = straight ? { x: px - 45720, y: py, w: 91440, h } : { x: x0, y: py, w: Math.abs(cx - px), h };
  const geom = kid(kid(sp, "spPr", NS.dsp), "custGeom", NS.a);
  for (const x of xfrmsOf(sp)) {
    setBox(x, box);
    for (const a of ["rot", "flipV"]) x.removeAttribute(a);
    // A preset line runs corner to corner: flipped when the item is left of its parent.
    if (!geom && cx < px) x.setAttribute("flipH", "1");
    else x.removeAttribute("flipH");
  }
  if (!geom) return;
  const sx = straight ? 45720 : Math.round(px - x0), ex = straight ? 45720 : Math.round(cx - x0), half = Math.round(h / 2);
  const points = straight ? [[sx, 0], [ex, h]] : [[sx, 0], [sx, half], [ex, half], [ex, h]];
  const doc = sp.ownerDocument;
  const el = (parent: Element, name: string) => parent.appendChild(doc.createElementNS(NS.a, `a:${name}`));
  const list = doc.createElementNS(NS.a, "a:pathLst");
  const path = el(list, "path");
  points.forEach(([x, y], i) => {
    const p = el(el(path, i ? "lnTo" : "moveTo"), "pt");
    p.setAttribute("x", String(x));
    p.setAttribute("y", String(y));
  });
  const old = kid(geom, "pathLst", NS.a);
  if (old) old.replaceWith(list);
  else geom.appendChild(list);
}

/**
 * Lay the tree out again in the space it had: each item over the middle of its children, neighbours `plan.gap`
 * apart, levels where they were (a new level as far below as the last two are apart); shrunk to fit, never grown.
 */
function relayoutTree(top: Item[], plan: TreePlan): void {
  const all = flat(top);
  const levels = [...plan.levels];
  for (let d = levels.length; d <= Math.max(...all.map((i) => i.depth)); d++) levels[d] = { y: levels[d - 1].y + (levels[d - 2] ? levels[d - 1].y - levels[d - 2].y : levels[d - 1].h * 1.4), h: levels[d - 1].h };
  const span = (list: Item[]): number => list.reduce((a, i) => a + width(i), 0) + plan.gap * Math.max(0, list.length - 1);
  const width = (i: Item): number => Math.max(i.box.w, span(i.children));
  const left = new Map<Item, number>();
  const assign = (list: Item[], x: number) => {
    for (const i of list) {
      left.set(i, x + (width(i) - i.box.w) / 2);
      assign(i.children, x + (width(i) - span(i.children)) / 2);
      x += width(i) + plan.gap;
    }
  };
  assign(top, 0);
  const deepest = levels[Math.max(...all.map((i) => i.depth))];
  const s = Math.min(1, plan.bounds.w / span(top), plan.bounds.h / (deepest.y + deepest.h - levels[0].y));
  const x0 = plan.bounds.x + (plan.bounds.w - span(top) * s) / 2, y0 = levels[0].y;
  const placed = new Map<Item, Box>();
  for (const i of all) {
    const to: Box = { x: x0 + left.get(i)! * s, y: y0 + (levels[i.depth].y - y0) * s, w: i.box.w * s, h: i.box.h * s };
    for (const sp of [...i.shapes, ...i.gap])
      for (const x of xfrmsOf(sp)) {
        const b = boxOf(x);
        setBox(x, { x: to.x + (b.x - i.box.x) * s, y: to.y + (b.y - i.box.y) * s, w: b.w * s, h: b.h * s });
      }
    scaleText([...i.shapes, ...i.gap, ...i.conn], s);
    placed.set(i, to);
  }
  for (const i of all) {
    const p = i.parent && placed.get(i.parent), c = placed.get(i)!;
    if (p) for (const sp of i.conn) setConnector(sp, p.x + p.w / 2, p.y + p.h, c.x + c.w / 2, c.y);
  }
}

/** Clone a drawing shape with a fresh model id, right after `after`. */
function cloneShape(sp: Element, after: Element): Element {
  const copy = sp.cloneNode(true) as Element;
  copy.setAttribute("modelId", guid());
  after.after(copy);
  return copy;
}

/**
 * Add an item after `afterId`, at its level, or (`below`) under it as its last child; returns its id. Rows and
 * cycles take top-level items; a tree takes them anywhere; a radial diagram round its middle item.
 */
export function addSmartArtItem(data: Document, drawing: Document | null, afterId: string, text: string, below = false): string {
  const top = readItems(data, drawing);
  const kind = layoutOf(top, drawing);
  let like = flat(top).find((i) => i.id === afterId);
  if (kind === "radial" && (!like || (below ? like !== top[0] : like.parent !== top[0]))) throw new Error(RING_ONLY);
  if (!like || (kind === "row" || kind === "cycle" ? like.depth > 0 || below : false)) throw new Error("Only top-level items can be added to.");
  // Under an item that has children: after the last of them.
  if (below && like.children.length) {
    like = like.children[like.children.length - 1];
    below = false;
  }
  const cxn = like.cxn, node = ptById(data, like.id);
  if (!node) throw new Error("Only top-level items can be added to.");
  const parentId = below ? like.id : cxn.getAttribute("srcId") ?? "";
  const plan = kind === "tree" ? treePlan(top) : null;
  const ring = kind === "radial" ? [...top[0].children] : null;
  const groups = plan || ring ? [] : top.map(groupOf);
  const k = top.indexOf(like);
  const id = guid(), cxnId = guid(), parT = guid(), sibT = guid();
  // Data: the item, its two transition points, and its connection, ordered after `afterId` (or first under it).
  const pt = node.cloneNode(true) as Element;
  pt.setAttribute("modelId", id);
  const body = kid(pt, "t", NS.dgm);
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
  const ord = below ? -1 : ordOf(cxn);
  for (const c of cxnsOf(data)) if (isParOf(c) && c.getAttribute("srcId") === parentId && ordOf(c) > ord) c.setAttribute("srcOrd", String(ordOf(c) + 1));
  const link = cxn.cloneNode(true) as Element;
  for (const [a, v] of [["modelId", cxnId], ["srcId", parentId], ["destId", id], ["srcOrd", String(ord + 1)], ["parTransId", parT], ["sibTransId", sibT]]) link.setAttribute(a, v);
  cxn.after(link);
  // Presentation points (and their links) for the new item and its connector, copied from the item's, so its
  // drawing shapes stand for points in the data like every other shape.
  const map = new Map<string, string>([[like.id, id], [cxn.getAttribute("parTransId") ?? "", parT], [cxn.getAttribute("sibTransId") ?? "", sibT]]);
  const clonePres = (assoc: string, to: string) => {
    for (const p of pointsOf(data).filter((p) => p.getAttribute("type") === "pres" && kid(p, "prSet", NS.dgm)?.getAttribute("presAssocID") === assoc)) {
      const copy = p.cloneNode(true) as Element;
      const fresh = guid();
      map.set(p.getAttribute("modelId") ?? "", fresh);
      copy.setAttribute("modelId", fresh);
      kid(copy, "prSet", NS.dgm)?.setAttribute("presAssocID", to);
      p.parentElement!.appendChild(copy);
    }
  };
  clonePres(like.id, id);
  clonePres(cxn.getAttribute("sibTransId") ?? "", sibT);
  // A tree draws the line from the parent: the parent transition's points.
  if (plan || ring) clonePres(cxn.getAttribute("parTransId") ?? "", parT);
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
  // Drawing: copies of the item's shapes, with the new text.
  const shapes: Element[] = [];
  let last = like.gap[like.gap.length - 1] ?? like.shapes[like.shapes.length - 1];
  for (const sp of like.shapes) shapes.push((last = shapeCopy(sp, last)));
  for (const sp of shapes) {
    const tx = kid(sp, "txBody", NS.dsp);
    if (tx) setBody(tx, text);
  }
  if (plan || ring) {
    // Tree or radial: the copy takes its place among its parent's children, with a copy of the line the item has.
    const parent = below ? like : like.parent;
    const added: Item = { id, cxn: link, parent, children: [], depth: below ? like.depth + 1 : like.depth, shapes, gap: [], conn: like.conn.map((sp) => (last = shapeCopy(sp, last))), box: like.box };
    const list = parent ? parent.children : top;
    list.splice(below ? 0 : list.indexOf(like) + 1, 0, added);
    if (plan) relayoutTree(top, plan);
    else relayoutArc(top[0], ring!, top[0].children);
    return id;
  }
  // Row or cycle: whoever is now followed by an item has a connector after it.
  const src = groups[k];
  const added: Group = { shapes, gap: [], extra: [], box: src.box };
  if (src.gap.length) for (const sp of src.gap) added.gap.push((last = shapeCopy(sp, last)));
  else if (k > 0) {
    // The last item gains the connector the one before it has, standing for its own (so far undrawn) transition.
    const prev = groups[k - 1];
    const prevSib = top[k - 1].cxn.getAttribute("sibTransId") ?? "";
    const sib = cxn.getAttribute("sibTransId") ?? "";
    const presOf = (assoc: string) => pointsOf(data).filter((p) => p.getAttribute("type") === "pres" && kid(p, "prSet", NS.dgm)?.getAttribute("presAssocID") === assoc);
    map.clear();
    const own = presOf(sib);
    if (own.length) {
      // Its transition already has presentation points (PowerPoint keeps them): the shapes stand for those.
      const name = (p: Element) => kid(p, "prSet", NS.dgm)?.getAttribute("presName");
      for (const p of presOf(prevSib)) {
        const match = own.find((o) => name(o) === name(p));
        if (match) map.set(p.getAttribute("modelId") ?? "", match.getAttribute("modelId") ?? "");
      }
    } else {
      clonePres(prevSib, sib);
      linkPres(new Set(map.keys()));
    }
    const across = rowAcross(top) ?? true;
    src.gap = prev.gap.map((sp) => shapeCopy(sp, src.shapes[src.shapes.length - 1]));
    for (const sp of src.gap) for (const x of xfrmsOf(sp)) setBox(x, { ...boxOf(x), ...(across ? { x: boxOf(x).x + (src.box.x - prev.box.x) } : { y: boxOf(x).y + (src.box.y - prev.box.y) }) });
  }
  const next = [...groups];
  next.splice(k + 1, 0, added);
  if (kind === "cycle") relayoutCycle(groups, next);
  else relayout(groups, next, rowAcross(top) ?? true);
  return id;
}

/** Remove an item, with its children, connections and shapes. Rows and cycles give up top-level items; a tree any. */
export function removeSmartArtItem(data: Document, drawing: Document | null, id: string): void {
  const top = readItems(data, drawing);
  const kind = layoutOf(top, drawing);
  const item = flat(top).find((i) => i.id === id);
  if (kind === "radial" && item?.parent !== top[0]) throw new Error(RING_ONLY);
  if (!item || ((kind === "row" || kind === "cycle") && item.depth > 0)) throw new Error("Only top-level items can be removed.");
  if (item.parent ? kind === "radial" && item.parent.children.length < 2 : top.length < 2) throw new Error("A diagram needs at least one item.");
  const plan = kind === "tree" ? treePlan(top) : null;
  const ring = kind === "radial" ? [...top[0].children] : null;
  const groups = plan || ring ? [] : top.map(groupOf);
  const across = kind === "row" ? rowAcross(top) ?? true : true;
  const k = top.indexOf(item);
  // The item and everything under it, with their transition points.
  const gone = new Set<string>();
  const collect = (at: string) => {
    gone.add(at);
    for (const c of cxnsOf(data)) {
      if (!isParOf(c)) continue;
      if (c.getAttribute("destId") === at) for (const a of ["modelId", "parTransId", "sibTransId"]) { const v = c.getAttribute(a); if (v) gone.add(v); }
      const dest = c.getAttribute("destId");
      if (c.getAttribute("srcId") === at && dest) collect(dest);
    }
  };
  collect(id);
  const ord = ordOf(item.cxn), parentId = item.cxn.getAttribute("srcId");
  // Presentation points standing for them go too, with the shapes drawn for those, and every connection touching any.
  for (const p of pointsOf(data)) { const assoc = kid(p, "prSet", NS.dgm)?.getAttribute("presAssocID"), mid = p.getAttribute("modelId"); if (p.getAttribute("type") === "pres" && assoc && gone.has(assoc) && mid) gone.add(mid); }
  if (drawing) {
    for (const sp of Array.from(drawing.getElementsByTagNameNS(NS.dsp, "sp"))) { const mid = sp.getAttribute("modelId"); if (mid && gone.has(mid)) sp.remove(); }
    // A row's new last item drops its connector (a cycle's goes back to the first item).
    if (kind === "row" && k === groups.length - 1 && k > 0) {
      for (const sp of groups[k - 1].gap) sp.remove();
      groups[k - 1].gap = [];
    }
  }
  for (const p of pointsOf(data)) if (gone.has(p.getAttribute("modelId") ?? "")) p.remove();
  for (const c of cxnsOf(data)) if (gone.has(c.getAttribute("modelId") ?? "") || gone.has(c.getAttribute("srcId") ?? "") || gone.has(c.getAttribute("destId") ?? "")) c.remove();
  for (const c of cxnsOf(data)) if (isParOf(c) && c.getAttribute("srcId") === parentId && ordOf(c) > ord) c.setAttribute("srcOrd", String(ordOf(c) - 1));
  if (!drawing) return;
  if (plan) {
    const list = item.parent ? item.parent.children : top;
    list.splice(list.indexOf(item), 1);
    relayoutTree(top, plan);
  } else if (ring) relayoutArc(top[0], ring, ring.filter((i) => i !== item));
  else if (kind === "cycle") relayoutCycle(groups, groups.filter((_, i) => i !== k));
  else relayout(groups, groups.filter((_, i) => i !== k), across);
}
