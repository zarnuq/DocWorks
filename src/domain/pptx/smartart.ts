/**
 * SmartArt text: the items live in the diagram's data model (dgm:pt); PowerPoint also keeps a drawing of them
 * (the shapes we show), linked through the presentation points that stand for each item (presAssocID). Editing
 * an item's text changes both, so the slide shows it at once and PowerPoint agrees. Items aren't added or
 * removed here: that needs the diagram's layout run again, which only Office does.
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

