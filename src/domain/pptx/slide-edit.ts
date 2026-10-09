import { NS, childNS, escapeXml } from "../../core/utils/xml";
import { Para, Rect, Run, Spacing, TextBody, TextStyle, resolveLevel, runProps } from "./slide-model";
import { fitTable } from "./table";

/**
 * What the editor hands back per paragraph; extra fields from Para are ignored.
 * Spacing left undefined is not touched.
 */
export type ParaInput = Pick<Para, "runs" | "align" | "lvl" | "bullet"> & Partial<Pick<Para, "lnSpc" | "spcBef" | "spcAft">>;

const RUN_TAGS = new Set(["r", "br", "fld"]);
const RUN_KEYS: (keyof Run)[] = ["text", "b", "i", "u", "strike", "sz", "color", "font", "spc", "cap", "link", "field"];
const sameRun = (a: Run, b: Run) => RUN_KEYS.every((k) => String(a[k] ?? "").toUpperCase() === String(b[k] ?? "").toUpperCase());
const BULLET_TAGS = new Set(["buNone", "buAutoNum", "buChar", "buBlip"]);
/** a:pPr children in schema order (bullet color/size/font and the bullet itself share slots). */
const PPR_ORDER = ["lnSpc", "spcBef", "spcAft", "buClrTx", "buClr", "buSzTx", "buSzPct", "buSzPts", "buFontTx", "buFont", "bullet", "tabLst", "defRPr", "extLst"];
const pprSlot = (local: string) => (BULLET_TAGS.has(local) ? "bullet" : local);
const sameSpacing = (a: Spacing | null | undefined, b: Spacing | null | undefined) => JSON.stringify(a ?? null) === JSON.stringify(b ?? null);
const FILL_TAGS = new Set(["noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"]);
/** Children of a:rPr that must come after a:latin. */
const AFTER_LATIN = new Set(["ea", "cs", "sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst"]);

const el = (doc: Document, name: string) => doc.createElementNS(NS.a, `a:${name}`);
const kids = (parent: Element | null, local: string) =>
  parent ? Array.from(parent.children).filter((c) => c.localName === local && c.namespaceURI === NS.a) : [];

function insertBeforeFirst(parent: Element, child: Element, after: Set<string>): void {
  const anchor = Array.from(parent.children).find((c) => after.has(c.localName)) ?? null;
  parent.insertBefore(child, anchor);
}

/** Put `child` into a:pPr at its schema position, replacing any element in the same slot. */
function placeInPPr(pPr: Element, child: Element | null, slot: string): void {
  Array.from(pPr.children).filter((c) => pprSlot(c.localName) === slot).forEach((c) => c.remove());
  if (!child) return;
  const at = PPR_ORDER.indexOf(slot);
  const next = Array.from(pPr.children).find((c) => PPR_ORDER.indexOf(pprSlot(c.localName)) > at) ?? null;
  pPr.insertBefore(child, next);
}

function spacingEl(doc: Document, local: string, s: Spacing): Element {
  const outer = el(doc, local);
  const inner = outer.appendChild(el(doc, "pct" in s ? "spcPct" : "spcPts"));
  inner.setAttribute("val", String(Math.round("pct" in s ? s.pct * 100000 : s.pts * 100)));
  return outer;
}

function writeRunProps(rPr: Element, run: Run, base: ReturnType<typeof runProps>, style: TextStyle): void {
  const doc = rPr.ownerDocument;
  if (Math.abs(run.spc - base.spc) > 0.001) {
    if (run.spc) rPr.setAttribute("spc", String(Math.round(run.spc * 100)));
    else rPr.removeAttribute("spc");
  }
  if (run.cap !== base.cap) rPr.setAttribute("cap", run.cap);
  if (run.link !== base.link) {
    childNS(rPr, NS.a, "hlinkClick")?.remove();
    if (run.link && style.links) {
      const h = el(doc, "hlinkClick");
      h.setAttributeNS(NS.r, "r:id", style.links.add(run.link));
      insertBeforeFirst(rPr, h, new Set(["hlinkMouseOver", "rtl", "extLst"]));
    }
  }
  if (run.b !== base.b) rPr.setAttribute("b", run.b ? "1" : "0");
  if (run.i !== base.i) rPr.setAttribute("i", run.i ? "1" : "0");
  if (run.u !== base.u) rPr.setAttribute("u", run.u ? "sng" : "none");
  if (run.strike !== base.strike) rPr.setAttribute("strike", run.strike ? "sngStrike" : "noStrike");
  if (Math.abs(run.sz - base.sz) > 0.001) rPr.setAttribute("sz", String(Math.round(run.sz * 100)));

  if (run.color.toUpperCase() !== base.color.toUpperCase()) {
    Array.from(rPr.children).filter((c) => FILL_TAGS.has(c.localName)).forEach((c) => c.remove());
    const fill = el(doc, "solidFill");
    const clr = el(doc, "srgbClr");
    clr.setAttribute("val", run.color.replace("#", "").toUpperCase());
    fill.appendChild(clr);
    // Fill follows an optional a:ln and precedes everything else.
    const ln = childNS(rPr, NS.a, "ln");
    rPr.insertBefore(fill, ln ? ln.nextSibling : rPr.firstChild);
  }

  if (run.font !== base.font) {
    let latin = childNS(rPr, NS.a, "latin");
    if (!latin) {
      latin = el(doc, "latin");
      insertBeforeFirst(rPr, latin, AFTER_LATIN);
    }
    latin.setAttribute("typeface", run.font);
  }
}

/**
 * Replace the paragraphs of a text body. Only properties that differ from what
 * the XML already says (directly or by inheritance) are written, so untouched
 * text keeps its original markup, theme colors and fonts.
 */
export function writeParagraphs(body: TextBody, input: ParaInput[]): void {
  const txBody = body.el;
  const doc = txBody.ownerDocument;
  const style = body.style;
  const original = kids(txBody, "p");
  const paras: ParaInput[] = input.length ? input : [{ ...(body.paras[0] ?? { align: "l" as const, lvl: 0, bullet: null }), runs: [] }];

  const written = paras.map((para, k) => {
    const same = original[k] ?? null;
    const template = same ?? original[original.length - 1] ?? null;
    const p = el(doc, "p");

    const pPr = (kids(template, "pPr")[0]?.cloneNode(true) as Element | undefined) ?? el(doc, "pPr");
    // An explicit lvl="0" stays when the level didn't change.
    if (para.lvl !== Number(pPr.getAttribute("lvl") ?? 0)) {
      if (para.lvl) pPr.setAttribute("lvl", String(para.lvl));
      else pPr.removeAttribute("lvl");
    }
    const level = resolveLevel(style, para.lvl, pPr);
    if (para.align !== level.align) pPr.setAttribute("algn", para.align);
    if (para.bullet !== level.bullet) {
      const auto = para.bullet?.startsWith("#") ? /^#(\w+)(?:@(\d+))?$/.exec(para.bullet) : null;
      const bu = el(doc, para.bullet === null ? "buNone" : auto ? "buAutoNum" : "buChar");
      if (auto) {
        bu.setAttribute("type", auto[1]);
        if (auto[2] && auto[2] !== "1") bu.setAttribute("startAt", auto[2]);
      } else if (para.bullet !== null) bu.setAttribute("char", para.bullet);
      placeInPPr(pPr, bu, "bullet");
    }
    for (const key of ["lnSpc", "spcBef", "spcAft"] as const) {
      const v = para[key];
      if (v === undefined || sameSpacing(v, level[key])) continue;
      placeInPPr(pPr, v ? spacingEl(doc, key, v) : null, key);
    }
    if (pPr.attributes.length || pPr.children.length) p.appendChild(pPr);

    // Run-like children line up 1:1 with the runs readParagraphs produced.
    const origItems = template ? Array.from(template.children).filter((c) => RUN_TAGS.has(c.localName)) : [];
    const origRuns = same ? body.paras[k]?.runs ?? [] : [];
    para.runs.forEach((run, j) => {
      const twin = same && origItems[j];
      if (twin && origRuns[j] && sameRun(origRuns[j], run)) {
        p.appendChild(twin.cloneNode(true));
        return;
      }
      const baseRPr = kids(origItems[j] ?? origItems[origItems.length - 1] ?? null, "rPr")[0];
      const rPr = (baseRPr?.cloneNode(true) as Element | undefined) ?? el(doc, "rPr");
      writeRunProps(rPr, run, runProps(rPr, level.defaults, style), style);
      const hasProps = rPr.attributes.length > 0 || rPr.children.length > 0;
      if (run.field !== undefined) {
        const fld = el(doc, "fld");
        fld.setAttribute("id", `{${crypto.randomUUID().toUpperCase()}}`);
        fld.setAttribute("type", run.field);
        fld.appendChild(hasProps ? rPr : el(doc, "rPr"));
        fld.appendChild(el(doc, "t")).textContent = run.text;
        p.appendChild(fld);
        return;
      }

      run.text.split("\n").forEach((segment, s) => {
        if (s > 0) {
          const br = el(doc, "br");
          if (hasProps) br.appendChild(rPr.cloneNode(true));
          p.appendChild(br);
        }
        if (!segment) return;
        const r = el(doc, "r");
        if (hasProps) r.appendChild(rPr.cloneNode(true));
        const t = el(doc, "t");
        t.textContent = segment;
        r.appendChild(t);
        p.appendChild(r);
      });
    });

    const endPr = kids(template, "endParaRPr")[0];
    if (endPr) p.appendChild(endPr.cloneNode(true));
    return p;
  });

  original.forEach((p) => p.remove());
  written.forEach((p) => txBody.appendChild(p));
}

export interface BodyPrPatch {
  anchor?: "t" | "ctr" | "b";
  insets?: { l: number; t: number; r: number; b: number };
  wrap?: boolean;
  autofit?: { type: "none" } | { type: "resize" } | { type: "shrink"; fontScale?: number; lnSpcReduction?: number };
}

const AUTOFIT_TAGS = new Set(["noAutofit", "normAutofit", "spAutoFit"]);

/** Change text box settings on a:bodyPr (anchor, insets, wrapping, autofit). */
export function writeBodyPr(txBody: Element, patch: BodyPrPatch): void {
  const doc = txBody.ownerDocument;
  const bodyPr = kids(txBody, "bodyPr")[0] ?? txBody.insertBefore(el(doc, "bodyPr"), txBody.firstChild);
  if (patch.anchor) bodyPr.setAttribute("anchor", patch.anchor);
  if (patch.insets) {
    for (const [k, attr] of [["l", "lIns"], ["t", "tIns"], ["r", "rIns"], ["b", "bIns"]] as const) {
      bodyPr.setAttribute(attr, String(Math.round(patch.insets[k])));
    }
  }
  if (patch.wrap !== undefined) bodyPr.setAttribute("wrap", patch.wrap ? "square" : "none");
  if (patch.autofit) {
    Array.from(bodyPr.children).filter((c) => AUTOFIT_TAGS.has(c.localName)).forEach((c) => c.remove());
    const fit = patch.autofit;
    const tag = fit.type === "shrink" ? "normAutofit" : fit.type === "resize" ? "spAutoFit" : "noAutofit";
    const child = el(doc, tag);
    if (fit.type === "shrink") {
      if (fit.fontScale !== undefined && fit.fontScale < 1) child.setAttribute("fontScale", String(Math.round(fit.fontScale * 100000)));
      if (fit.lnSpcReduction) child.setAttribute("lnSpcReduction", String(Math.round(fit.lnSpcReduction * 100000)));
    }
    // a:prstTxWarp comes first; autofit precedes scene3d, sp3d, flatTx, extLst.
    const warp = kids(bodyPr, "prstTxWarp")[0];
    bodyPr.insertBefore(child, warp ? warp.nextSibling : bodyPr.firstChild);
  }
}

// -----------------------------------------------------------------------------
// Shapes
// -----------------------------------------------------------------------------

function ensureChild(parent: Element, ns: string, qname: string, afterLocal: string | null): Element {
  const local = qname.split(":")[1];
  let child = childNS(parent, ns, local);
  if (!child) {
    child = parent.ownerDocument.createElementNS(ns, qname);
    const after = afterLocal ? Array.from(parent.children).find((c) => c.localName.startsWith(afterLocal)) : null;
    parent.insertBefore(child, after ? after.nextSibling : parent.firstChild);
  }
  return child;
}

export function setRect(shape: Element, rect: Rect): void {
  let xfrm: Element;
  if (shape.localName === "graphicFrame") {
    xfrm = ensureChild(shape, NS.p, "p:xfrm", "nvGraphicFramePr");
    // A table's grid stretches with its frame.
    const tbl = shape.getElementsByTagNameNS(NS.a, "tbl")[0];
    if (tbl) fitTable(tbl, rect.w, rect.h);
  } else {
    const spPr = ensureChild(shape, NS.p, "p:spPr", "nv");
    xfrm = ensureChild(spPr, NS.a, "a:xfrm", null);
  }
  const off = ensureChild(xfrm, NS.a, "a:off", null);
  const ext = ensureChild(xfrm, NS.a, "a:ext", "off");
  off.setAttribute("x", String(Math.round(rect.x)));
  off.setAttribute("y", String(Math.round(rect.y)));
  ext.setAttribute("cx", String(Math.round(rect.w)));
  ext.setAttribute("cy", String(Math.round(rect.h)));
}

export function nextShapeId(doc: Document): number {
  return Array.from(doc.getElementsByTagNameNS(NS.p, "cNvPr")).reduce((m, c) => Math.max(m, Number(c.getAttribute("id")) || 0), 0) + 1;
}

const xfrmXml = (r: Rect) =>
  `<a:xfrm><a:off x="${Math.round(r.x)}" y="${Math.round(r.y)}"/><a:ext cx="${Math.round(r.w)}" cy="${Math.round(r.h)}"/></a:xfrm>`;

export function appendShape(doc: Document, xml: string): Element {
  const frag = new DOMParser().parseFromString(
    `<root xmlns:a="${NS.a}" xmlns:p="${NS.p}" xmlns:r="${NS.r}">${xml}</root>`,
    "application/xml"
  ).documentElement.firstElementChild!;
  const node = doc.importNode(frag, true);
  const tree = doc.getElementsByTagNameNS(NS.p, "spTree")[0];
  tree.insertBefore(node, childNS(tree, NS.p, "extLst"));
  return node;
}

export function addTextBox(doc: Document, rect: Rect, text: string): Element {
  const id = nextShapeId(doc);
  return appendShape(
    doc,
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="TextBox ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr>${xfrmXml(rect)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>` +
      `<p:txBody><a:bodyPr wrap="square" rtlCol="0"/><a:lstStyle/>` +
      `<a:p><a:r><a:rPr lang="en-US"/><a:t>${escapeXml(text)}</a:t></a:r></a:p></p:txBody></p:sp>`
  );
}

/** Presets drawn as connector shapes (p:cxnSp): no text, styled by the theme's line. */
export const CONNECTORS = new Set([
  "line", "straightConnector1",
  "bentConnector2", "bentConnector3", "bentConnector4", "bentConnector5",
  "curvedConnector2", "curvedConnector3", "curvedConnector4", "curvedConnector5",
]);

const flipAttrs = (f?: { flipH?: boolean; flipV?: boolean }) => `${f?.flipH ? ' flipH="1"' : ""}${f?.flipV ? ' flipV="1"' : ""}`;

/** Insert a preset shape styled like PowerPoint's defaults (accent fill, darker outline, light text). */
export function addShape(
  doc: Document,
  prst: string,
  rect: Rect,
  opts: { flipH?: boolean; flipV?: boolean; tailArrow?: boolean } = {}
): Element {
  const id = nextShapeId(doc);
  const xfrm = xfrmXml(rect).replace("<a:xfrm>", `<a:xfrm${flipAttrs(opts)}>`);
  const geom = `<a:prstGeom prst="${escapeXml(prst)}"><a:avLst/></a:prstGeom>`;
  if (CONNECTORS.has(prst)) {
    const ln = opts.tailArrow ? `<a:ln><a:tailEnd type="triangle"/></a:ln>` : "";
    return appendShape(
      doc,
      `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${id}" name="Connector ${id}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>` +
        `<p:spPr>${xfrm}${geom}${ln}</p:spPr>` +
        `<p:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef>` +
        `<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></p:style></p:cxnSp>`
    );
  }
  return appendShape(
    doc,
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Shape ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr>${xfrm}${geom}</p:spPr>` +
      `<p:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef>` +
      `<a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>` +
      `<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>` +
      `<p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`
  );
}

export const shapeId = (shape: Element) => shape.getElementsByTagNameNS(NS.p, "cNvPr")[0]?.getAttribute("id") ?? "";

function xfrmOf(shape: Element): Element {
  const spPr = ensureChild(shape, NS.p, "p:spPr", "nv");
  return ensureChild(spPr, NS.a, "a:xfrm", null);
}

/** Set rotation (degrees clockwise) and flips; 0 / false remove the attribute. */
export function setTransform(shape: Element, t: { rot?: number; flipH?: boolean; flipV?: boolean }): void {
  const xfrm = xfrmOf(shape);
  if (t.rot !== undefined) {
    const rot = Math.round((((t.rot % 360) + 360) % 360) * 60000);
    if (rot) xfrm.setAttribute("rot", String(rot));
    else xfrm.removeAttribute("rot");
  }
  for (const k of ["flipH", "flipV"] as const) {
    if (t[k] === undefined) continue;
    if (t[k]) xfrm.setAttribute(k, "1");
    else xfrm.removeAttribute(k);
  }
}

/** Set one preset adjust value (a:avLst/a:gd). */
export function setAdjust(shape: Element, name: string, value: number): void {
  const prstGeom = shape.getElementsByTagNameNS(NS.a, "prstGeom")[0];
  if (!prstGeom) return;
  const avLst = childNS(prstGeom, NS.a, "avLst") ?? prstGeom.appendChild(el(shape.ownerDocument, "avLst"));
  const gd = Array.from(avLst.children).find((g) => g.getAttribute("name") === name) ?? avLst.appendChild(el(shape.ownerDocument, "gd"));
  gd.setAttribute("name", name);
  gd.setAttribute("fmla", `val ${Math.round(value)}`);
}

export type ZOrder = "front" | "back" | "forward" | "backward";

/** Move a shape in the drawing order among its sibling shapes. */
export function reorder(shape: Element, to: ZOrder): void {
  const tree = shape.parentElement!;
  const shapes = Array.from(tree.children).filter((c) => !["nvGrpSpPr", "grpSpPr", "extLst"].includes(c.localName));
  const i = shapes.indexOf(shape);
  const target =
    to === "front" ? null : to === "back" ? shapes[0] : to === "forward" ? (shapes[i + 1]?.nextSibling ?? null) : shapes[i - 1];
  if (to === "forward" && i === shapes.length - 1) return;
  if (to === "backward" && i === 0) return;
  tree.insertBefore(shape, target === null ? childNS(tree, NS.p, "extLst") : target);
}

export interface Glue {
  /** cNvPr id of the shape */
  id: string;
  /** connection site index */
  idx: number;
}

/** Glue (or unglue, with null) one end of a connector to a shape's connection site. */
export function connect(cxn: Element, end: "st" | "end", target: Glue | null): void {
  const nv = cxn.getElementsByTagNameNS(NS.p, "cNvCxnSpPr")[0];
  if (!nv) return;
  childNS(nv, NS.a, `${end}Cxn`)?.remove();
  if (!target) return;
  const g = el(cxn.ownerDocument, `${end}Cxn`);
  g.setAttribute("id", target.id);
  g.setAttribute("idx", String(target.idx));
  // a:cxnSpLocks, a:stCxn, a:endCxn, a:extLst
  const before = end === "st" ? (childNS(nv, NS.a, "endCxn") ?? childNS(nv, NS.a, "extLst")) : childNS(nv, NS.a, "extLst");
  nv.insertBefore(g, before);
}

export function connections(cxn: Element): { st: Glue | null; end: Glue | null } {
  const nv = cxn.getElementsByTagNameNS(NS.p, "cNvCxnSpPr")[0];
  const read = (end: string): Glue | null => {
    const g = nv && childNS(nv, NS.a, `${end}Cxn`);
    return g ? { id: g.getAttribute("id") ?? "", idx: Number(g.getAttribute("idx")) } : null;
  };
  return { st: read("st"), end: read("end") };
}

/** Bounding box and flips of a line from `a` to `b`. */
export function lineBox(a: { x: number; y: number }, b: { x: number; y: number }): { rect: Rect; flipH: boolean; flipV: boolean } {
  return {
    rect: { x: Math.min(a.x, b.x), y: Math.min(a.y, b.y), w: Math.abs(b.x - a.x), h: Math.abs(b.y - a.y) },
    flipH: b.x < a.x,
    flipV: b.y < a.y,
  };
}

export function addPicture(doc: Document, rect: Rect, rId: string, name: string): Element {
  const id = nextShapeId(doc);
  return appendShape(
    doc,
    `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="Picture ${id}" descr="${escapeXml(name)}"/>` +
      `<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>` +
      `<p:blipFill><a:blip r:embed="${escapeXml(rId)}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
      `<p:spPr>${xfrmXml(rect)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`
  );
}


const paragraphText = (p: Element) =>
  Array.from(p.getElementsByTagNameNS(NS.a, "t"))
    .map((t) => t.textContent)
    .join("");

/**
 * Replace a text body's paragraphs with plain-text lines (used for speaker
 * notes). Paragraphs whose text is unchanged are kept as they are; new or
 * changed ones copy the paragraph and run properties of their neighbour.
 */
export function writePlainParagraphs(txBody: Element, text: string): void {
  const doc = txBody.ownerDocument;
  const original = kids(txBody, "p");
  const written = text.split("\n").map((line, k) => {
    const same = original[k];
    if (same && paragraphText(same) === line && !same.getElementsByTagNameNS(NS.a, "br").length) return same;
    const template = same ?? original[original.length - 1] ?? null;
    const p = el(doc, "p");
    const pPr = kids(template, "pPr")[0];
    if (pPr) p.appendChild(pPr.cloneNode(true));
    if (line) {
      const r = p.appendChild(el(doc, "r"));
      const rPr = template?.getElementsByTagNameNS(NS.a, "rPr")[0];
      if (rPr) r.appendChild(rPr.cloneNode(true));
      r.appendChild(el(doc, "t")).textContent = line;
    }
    const end = kids(template, "endParaRPr")[0];
    if (end) p.appendChild(end.cloneNode(true));
    return p;
  });
  original.forEach((p) => p.remove());
  written.forEach((p) => txBody.appendChild(p));
}

interface Placed {
  el: Element;
  rect: Rect;
  rot: number;
  flipH: boolean;
  flipV: boolean;
  sites: { x: number; y: number }[];
}

/** Slide position of a shape's connection site, after flips and rotation. */
export function sitePoint(s: Placed, idx: number): { x: number; y: number } | null {
  const site = s.sites[idx];
  if (!site) return null;
  let x = (s.flipH ? s.rect.w - site.x : site.x) - s.rect.w / 2;
  let y = (s.flipV ? s.rect.h - site.y : site.y) - s.rect.h / 2;
  const a = (s.rot * Math.PI) / 180;
  [x, y] = [x * Math.cos(a) - y * Math.sin(a), x * Math.sin(a) + y * Math.cos(a)];
  return { x: s.rect.x + s.rect.w / 2 + x, y: s.rect.y + s.rect.h / 2 + y };
}

/** Move the ends of glued connectors onto their shapes' connection sites. Returns whether anything changed. */
export function rerouteConnectors(slide: Document, shapes: Placed[]): boolean {
  const byId = new Map(shapes.map((s) => [shapeId(s.el), s]));
  let changed = false;
  for (const cxn of Array.from(slide.getElementsByTagNameNS(NS.p, "cxnSp"))) {
    const glue = connections(cxn);
    const self = shapes.find((s) => s.el === cxn);
    if (!self || (!glue.st && !glue.end)) continue;
    const { rect: r, flipH, flipV } = self;
    let a = { x: flipH ? r.x + r.w : r.x, y: flipV ? r.y + r.h : r.y };
    let b = { x: flipH ? r.x : r.x + r.w, y: flipV ? r.y : r.y + r.h };
    const at = (g: Glue | null) => (g && byId.has(g.id) ? sitePoint(byId.get(g.id)!, g.idx) : null);
    a = at(glue.st) ?? a;
    b = at(glue.end) ?? b;
    const box = lineBox(a, b);
    const round = (v: Rect) => ({ x: Math.round(v.x), y: Math.round(v.y), w: Math.round(v.w), h: Math.round(v.h) });
    const next = round(box.rect);
    if (JSON.stringify(next) === JSON.stringify(round(r)) && box.flipH === flipH && box.flipV === flipV) continue;
    setRect(cxn, next);
    setTransform(cxn, { flipH: box.flipH, flipV: box.flipV });
    changed = true;
  }
  return changed;
}

/** Picture crop as fractions cut from each side (a:srcRect). */
export interface Crop {
  l: number;
  t: number;
  r: number;
  b: number;
}

const SIDES = ["l", "t", "r", "b"] as const;

export function readCrop(pic: Element): Crop | null {
  const src = pic.getElementsByTagNameNS(NS.a, "srcRect")[0];
  if (!src) return null;
  const crop = Object.fromEntries(SIDES.map((k) => [k, (Number(src.getAttribute(k)) || 0) / 100000])) as unknown as Crop;
  return SIDES.some((k) => crop[k]) ? crop : null;
}

/** Crop a picture like PowerPoint: the image keeps its scale and the frame shrinks or grows around it. */
// ponytail: ignores rotation; a rotated picture's frame moves along unrotated axes.
export function setCrop(pic: Element, next: Crop): void {
  const blipFill = childNS(pic, NS.p, "blipFill");
  const xfrm = childNS(childNS(pic, NS.p, "spPr") ?? pic, NS.a, "xfrm");
  if (!blipFill || !xfrm) return;
  const old = readCrop(pic) ?? { l: 0, t: 0, r: 0, b: 0 };
  const clamp = (v: number) => Math.min(0.95, Math.max(0, v || 0));
  const c = { l: clamp(next.l), t: clamp(next.t), r: clamp(next.r), b: clamp(next.b) };
  if (c.l + c.r > 0.95) c.r = 0.95 - c.l;
  if (c.t + c.b > 0.95) c.b = 0.95 - c.t;
  const off = childNS(xfrm, NS.a, "off")!;
  const ext = childNS(xfrm, NS.a, "ext")!;
  const r = { x: Number(off.getAttribute("x")), y: Number(off.getAttribute("y")), w: Number(ext.getAttribute("cx")), h: Number(ext.getAttribute("cy")) };
  // A file can crop away the whole image; treat that as 5% left visible.
  const fullW = r.w / Math.max(0.05, 1 - old.l - old.r);
  const fullH = r.h / Math.max(0.05, 1 - old.t - old.b);
  setRect(pic, { x: r.x + fullW * (c.l - old.l), y: r.y + fullH * (c.t - old.t), w: fullW * (1 - c.l - c.r), h: fullH * (1 - c.t - c.b) });

  blipFill.getElementsByTagNameNS(NS.a, "srcRect")[0]?.remove();
  if (!SIDES.some((k) => c[k])) return;
  const src = pic.ownerDocument.createElementNS(NS.a, "a:srcRect");
  for (const k of SIDES) if (c[k]) src.setAttribute(k, String(Math.round(c[k] * 100000)));
  const blip = childNS(blipFill, NS.a, "blip");
  blipFill.insertBefore(src, blip ? blip.nextSibling : blipFill.firstChild);
}
