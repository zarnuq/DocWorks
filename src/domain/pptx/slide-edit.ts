import { NS, childNS, escapeXml } from "../../core/utils/xml";
import { Para, Rect, Run, TextBody, resolveLevel, runProps } from "./slide-model";

/** What the editor hands back per paragraph; extra fields from Para are ignored. */
export type ParaInput = Pick<Para, "runs" | "align" | "lvl" | "bullet">;

const RUN_TAGS = new Set(["r", "br", "fld"]);
const RUN_KEYS: (keyof Run)[] = ["text", "b", "i", "u", "strike", "sz", "color", "font"];
const sameRun = (a: Run, b: Run) => RUN_KEYS.every((k) => String(a[k]).toUpperCase() === String(b[k]).toUpperCase());
const BULLET_TAGS = new Set(["buNone", "buAutoNum", "buChar", "buBlip"]);
/** Children of a:pPr that must come after the bullet element. */
const AFTER_BULLET = new Set(["tabLst", "defRPr", "extLst"]);
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

function writeRunProps(rPr: Element, run: Run, base: ReturnType<typeof runProps>): void {
  const doc = rPr.ownerDocument;
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
    if (para.lvl) pPr.setAttribute("lvl", String(para.lvl));
    else pPr.removeAttribute("lvl");
    const level = resolveLevel(style, para.lvl, pPr);
    if (para.align !== level.align) pPr.setAttribute("algn", para.align);
    if (para.bullet !== level.bullet) {
      Array.from(pPr.children).filter((c) => BULLET_TAGS.has(c.localName)).forEach((c) => c.remove());
      const bu = el(doc, para.bullet === null ? "buNone" : "buChar");
      if (para.bullet !== null) bu.setAttribute("char", para.bullet);
      insertBeforeFirst(pPr, bu, AFTER_BULLET);
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
      writeRunProps(rPr, run, runProps(rPr, level.defaults, style));
      const hasProps = rPr.attributes.length > 0 || rPr.children.length > 0;

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

function nextShapeId(doc: Document): number {
  return Array.from(doc.getElementsByTagNameNS(NS.p, "cNvPr")).reduce((m, c) => Math.max(m, Number(c.getAttribute("id")) || 0), 0) + 1;
}

const xfrmXml = (r: Rect) =>
  `<a:xfrm><a:off x="${Math.round(r.x)}" y="${Math.round(r.y)}"/><a:ext cx="${Math.round(r.w)}" cy="${Math.round(r.h)}"/></a:xfrm>`;

function appendShape(doc: Document, xml: string): Element {
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
