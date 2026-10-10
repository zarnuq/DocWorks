/**
 * Where a picture or text box sits: in line with the text, or floating (wp:anchor) with text wrapping around it,
 * above and below it, or not at all (behind / in front of the text) - Word's and Google Docs' wrap options.
 */
import { NS } from "../../core/utils/xml";

export type Wrap = "inline" | "square" | "topBottom" | "behind" | "front";

export interface Placement {
  wrap: Wrap;
  /** Offset (px) from the column's left edge and from the paragraph's top. */
  x: number;
  y: number;
  /** Horizontal alignment instead of an offset. */
  align: "left" | "right" | "center" | null;
}

export const INLINE: Placement = { wrap: "inline", x: 0, y: 0, align: null };

const EMU_PER_PX = 9525;
const kid = (el: Element | null | undefined, ns: string, local: string) =>
  (el ? Array.from(el.children).find((c) => c.namespaceURI === ns && c.localName === local) : undefined) ?? null;

/** The wp:inline / wp:anchor of a w:drawing (or of the first one inside `el`). */
function frameOf(el: Element): Element | null {
  return el.getElementsByTagNameNS(NS.wp, "anchor")[0] ?? el.getElementsByTagNameNS(NS.wp, "inline")[0] ?? null;
}

export function placementOf(drawing: Element | null): Placement {
  const frame = drawing && frameOf(drawing);
  if (!frame || frame.localName !== "anchor") return INLINE;
  const has = (local: string) => !!kid(frame, NS.wp, local);
  const behind = frame.getAttribute("behindDoc") === "1";
  const wrap: Wrap = has("wrapSquare") || has("wrapTight") || has("wrapThrough") ? "square" : has("wrapTopAndBottom") ? "topBottom" : behind ? "behind" : "front";
  const offset = (local: string) => Math.round(Number(kid(kid(frame, NS.wp, local), NS.wp, "posOffset")?.textContent ?? 0) / EMU_PER_PX);
  const align = kid(kid(frame, NS.wp, "positionH"), NS.wp, "align")?.textContent?.trim();
  // ponytail: offsets from the page or margin are taken from the column / paragraph; fine for the usual Word and Docs output.
  return { wrap, x: offset("positionH"), y: offset("positionV"), align: align === "left" || align === "right" || align === "center" ? align : null };
}

/** Width of the usual text column (6.5 in), for deciding which side a wrapped picture floats to. */
const COLUMN = 624;

/** CSS for a placement: floats for wrapped text, a block for break text, absolute for behind / in front. */
export function placementCss(p: Placement, width = 0): { cls: string; style: string } {
  switch (p.wrap) {
    case "inline":
      return { cls: "", style: "" };
    case "square":
      if (p.align === "center") return { cls: "docx-wrap", style: "display: block; margin-left: auto; margin-right: auto" };
      if (p.align === "right") return { cls: "docx-wrap", style: `float: right; margin-top: ${p.y}px` };
      // ponytail: a 6.5 in column is assumed; text wraps on one side only (CSS floats), Word wraps both.
      if (!p.align && width && p.x + width / 2 > COLUMN / 2) return { cls: "docx-wrap", style: `float: right; margin-right: ${Math.max(0, COLUMN - p.x - width)}px; margin-top: ${p.y}px` };
      return { cls: "docx-wrap", style: `float: left; margin-left: ${p.align ? 0 : p.x}px; margin-top: ${p.y}px` };
    case "topBottom":
      if (p.align === "center") return { cls: "docx-wrap", style: "display: block; margin-left: auto; margin-right: auto" };
      if (p.align === "right") return { cls: "docx-wrap", style: "display: block; margin-left: auto" };
      return { cls: "docx-wrap", style: `display: block; margin-left: ${p.align ? 0 : p.x}px; margin-top: ${p.y}px` };
    default: {
      const left = p.align === "right" ? "right: 0" : p.align === "center" ? "left: 50%; transform: translateX(-50%)" : `left: ${p.align ? 0 : p.x}px`;
      return { cls: `docx-float docx-${p.wrap}`, style: `${left}; top: ${p.y}px` };
    }
  }
}

/** Rewrite the w:drawing in `run` to `p`: wp:inline for in line, else wp:anchor with its wrap. Keeps the picture itself. */
export function setPlacement(run: Element, p: Placement): void {
  const drawing = run.getElementsByTagNameNS(NS.w, "drawing")[0];
  const frame = drawing && frameOf(drawing);
  if (!frame) return;
  const doc = run.ownerDocument;
  const el = (local: string, attrs: Record<string, string> = {}, text?: string) => {
    const e = doc.createElementNS(NS.wp, `wp:${local}`);
    for (const [k, v] of Object.entries(attrs)) e.setAttribute(k, v);
    if (text !== undefined) e.textContent = text;
    return e;
  };
  const keep = ["extent", "effectExtent", "docPr", "cNvGraphicFramePr"].map((l) => kid(frame, NS.wp, l)).filter(Boolean) as Element[];
  const graphic = Array.from(frame.children).find((c) => c.localName === "graphic") ?? null;
  const [extent, ...rest] = keep;
  const effect = rest.find((e) => e.localName === "effectExtent") ?? null;
  const tail = rest.filter((e) => e !== effect);

  let next: Element;
  if (p.wrap === "inline") {
    next = el("inline", { distT: "0", distB: "0", distL: "0", distR: "0" });
    for (const e of [extent, effect, ...tail, graphic]) if (e) next.appendChild(e);
  } else {
    const wasAnchor = frame.localName === "anchor";
    next = el("anchor", {
      distT: "0",
      distB: "0",
      distL: "114300",
      distR: "114300",
      simplePos: "0",
      relativeHeight: frame.getAttribute("relativeHeight") ?? "251659264",
      behindDoc: p.wrap === "behind" ? "1" : "0",
      locked: "0",
      layoutInCell: "1",
      allowOverlap: "1",
    });
    if (wasAnchor) for (const a of ["distT", "distB", "distL", "distR"]) next.setAttribute(a, frame.getAttribute(a) ?? next.getAttribute(a)!);
    next.appendChild(el("simplePos", { x: "0", y: "0" }));
    const h = next.appendChild(el("positionH", { relativeFrom: "column" }));
    h.appendChild(p.align ? el("align", {}, p.align) : el("posOffset", {}, String(Math.round(p.x * EMU_PER_PX))));
    next.appendChild(el("positionV", { relativeFrom: "paragraph" })).appendChild(el("posOffset", {}, String(Math.round(p.y * EMU_PER_PX))));
    for (const e of [extent, effect]) if (e) next.appendChild(e);
    next.appendChild(p.wrap === "square" ? el("wrapSquare", { wrapText: "bothSides" }) : p.wrap === "topBottom" ? el("wrapTopAndBottom") : el("wrapNone"));
    for (const e of [...tail, graphic]) if (e) next.appendChild(e);
  }
  frame.replaceWith(next);
}

/** A text box (wps:txbx) in a w:drawing: its paragraphs' text, size, placement and fill / border colours. */
export function textBoxOf(xml: string): { paragraphs: string[]; width: number; height: number; placement: Placement; fill: string | null; line: string | null } | null {
  if (!xml.includes("txbxContent")) return null;
  const root = new DOMParser().parseFromString(xml, "application/xml").documentElement;
  const content = root.getElementsByTagNameNS(NS.w, "txbxContent")[0];
  const frame = frameOf(root);
  if (!content || !frame) return null;
  const extent = kid(frame, NS.wp, "extent");
  const px = (attr: string) => Math.round(Number(extent?.getAttribute(attr) ?? 0) / EMU_PER_PX);
  const spPr = root.getElementsByTagNameNS("http://schemas.microsoft.com/office/word/2010/wordprocessingShape", "spPr")[0];
  const color = (el: Element | null | undefined) => {
    const v = el?.getElementsByTagNameNS(NS.a, "srgbClr")[0]?.getAttribute("val");
    return v && /^[0-9a-f]{6}$/i.test(v) ? `#${v}` : null;
  };
  const ln = spPr && kid(spPr, NS.a, "ln");
  return {
    paragraphs: Array.from(content.children)
      .filter((c) => c.localName === "p")
      .map((p) => Array.from(p.getElementsByTagNameNS(NS.w, "t")).map((t) => t.textContent).join("")),
    width: px("cx"),
    height: px("cy"),
    placement: placementOf(frame.parentElement ?? frame),
    fill: color(spPr && kid(spPr, NS.a, "solidFill")),
    line: ln && !kid(ln, NS.a, "noFill") ? color(kid(ln, NS.a, "solidFill")) : null,
  };
}
