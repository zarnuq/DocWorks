/*
 * Links on text runs and whole shapes (a:hlinkClick), as Google Slides' link dialog: a web address, a slide in this
 * deck (Office's ppaction://hlinksldjump with a slide relationship) or the next / previous / first / last slide
 * (ppaction://hlinkshowjump, no relationship). In the model a link is its address, "#slide:<slide part path>" or
 * "#jump:<nextslide|previousslide|firstslide|lastslide>", so moving slides around never breaks one.
 */
import { NS, childNS } from "../../core/utils/xml";
import type { PptxDocument } from "../../infrastructure/pptx/pptx-document";

const SLIDE_JUMP = "ppaction://hlinksldjump";
const SHOW_JUMP = "ppaction://hlinkshowjump?jump=";
export const JUMPS = ["nextslide", "previousslide", "firstslide", "lastslide"] as const;
const JUMP_LABELS: Record<string, string> = { nextslide: "Next slide", previousslide: "Previous slide", firstslide: "First slide", lastslide: "Last slide" };

/** The link an a:hlinkClick holds, or null for none (or an action we don't model, like playing media). */
export function readLink(h: Element | null | undefined, target: (rId: string) => string | null): string | null {
  if (!h) return null;
  const action = h.getAttribute("action") ?? "";
  const rId = h.getAttributeNS(NS.r, "id") ?? "";
  if (action === SLIDE_JUMP) {
    const path = rId && target(rId);
    return path ? `#slide:${path}` : null;
  }
  if (action.startsWith(SHOW_JUMP)) return `#jump:${action.slice(SHOW_JUMP.length)}`;
  if (action) return null;
  return (rId && target(rId)) || null;
}

/** r:id and action attributes for writing `link` from part `part`. */
export function linkAttrs(doc: PptxDocument, part: string, link: string): { id: string; action?: string } {
  if (link.startsWith("#slide:")) return { id: doc.slideRel(part, link.slice(7)), action: SLIDE_JUMP };
  if (link.startsWith("#jump:")) return { id: "", action: SHOW_JUMP + link.slice(6) };
  return { id: doc.linkRel(part, link) };
}

/** Fill an a:hlinkClick element's attributes for `link`. */
export function setLinkAttrs(h: Element, attrs: { id: string; action?: string }): void {
  h.setAttributeNS(NS.r, "r:id", attrs.id);
  if (attrs.action) h.setAttribute("action", attrs.action);
  else h.removeAttribute("action");
}

/** The shape's p:cNvPr (sp, pic, graphicFrame, cxnSp, grpSp). */
export function cNvPrOf(shape: Element): Element | null {
  const nv = Array.from(shape.children).find((c) => c.localName.startsWith("nv") && c.localName.endsWith("Pr"));
  return nv ? childNS(nv, NS.p, "cNvPr") : null;
}

/** Link a whole shape (null removes the link). */
export function setShapeLink(doc: PptxDocument, part: string, shape: Element, link: string | null): void {
  const cNvPr = cNvPrOf(shape);
  if (!cNvPr) return;
  childNS(cNvPr, NS.a, "hlinkClick")?.remove();
  if (link) {
    const h = cNvPr.ownerDocument.createElementNS(NS.a, "a:hlinkClick");
    setLinkAttrs(h, linkAttrs(doc, part, link));
    cNvPr.insertBefore(h, cNvPr.firstChild);
  }
  doc.touch(part);
}

/** The slide index a link shows from slide `current`, or null for a web address or a slide that is gone. */
export function linkTarget(link: string, current: number, slides: string[]): number | null {
  if (link.startsWith("#slide:")) {
    const i = slides.indexOf(link.slice(7));
    return i < 0 ? null : i;
  }
  const last = slides.length - 1;
  switch (link.slice(6)) {
    case "nextslide": return link.startsWith("#jump:") ? Math.min(last, current + 1) : null;
    case "previousslide": return link.startsWith("#jump:") ? Math.max(0, current - 1) : null;
    case "firstslide": return link.startsWith("#jump:") ? 0 : null;
    case "lastslide": return link.startsWith("#jump:") ? last : null;
  }
  return null;
}

/** How the link dialog and tooltips name a link. */
export function linkLabel(link: string, slides: string[]): string {
  if (link.startsWith("#slide:")) {
    const i = slides.indexOf(link.slice(7));
    return i < 0 ? "Missing slide" : `Slide ${i + 1}`;
  }
  if (link.startsWith("#jump:")) return JUMP_LABELS[link.slice(6)] ?? link;
  return link;
}

/** Remove hlinkClick elements using relationship ids `ids` from an XML part. */
export function dropLinks(xml: Document, ids: Set<string>): void {
  for (const h of Array.from(xml.getElementsByTagNameNS(NS.a, "hlinkClick"))) if (ids.has(h.getAttributeNS(NS.r, "id") ?? "")) h.remove();
}
