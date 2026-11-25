/*
 * Equations on slides, as PowerPoint keeps them: a text box inside mc:AlternateContent whose Choice
 * (Requires="a14") holds the math as a14:m around OMML, and whose Fallback is a picture of it for readers without
 * a14. The Choice is what we draw and edit (LaTeX through domain/docx/math.ts). An edited equation's Fallback becomes
 * the equation as text (we can't draw the picture); a moved one's Fallback is moved with it on save.
 */
import { NS, escapeXml, parseElement } from "../../core/utils/xml";
import { latexToOmml, ommlToLatex } from "../docx/math";
import { appendShape, nextShapeId } from "./slide-edit";
import type { Rect } from "./geometry";

export const A14 = "http://schemas.microsoft.com/office/drawing/2010/main";
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const branches = (ac: Element) => Array.from(ac.children).filter((c) => c.namespaceURI === NS.mc);
const choiceOf = (ac: Element) => branches(ac).find((c) => c.localName === "Choice") ?? null;
const fallbackOf = (ac: Element) => branches(ac).find((c) => c.localName === "Fallback") ?? null;
const hasMath = (el: Element | null) => !!el && el.getElementsByTagNameNS(A14, "m").length > 0;

/** The shape an mc:AlternateContent shows here: its Choice when that is a14 math, else its Fallback's. */
export function shownBranch(ac: Element): Element | null {
  const choice = choiceOf(ac);
  const requires = (choice?.getAttribute("Requires") ?? "").split(/\s+/);
  if (choice && requires.every((r) => r === "a14") && hasMath(choice)) return choice.firstElementChild;
  return fallbackOf(ac)?.firstElementChild ?? null;
}

/** Whether a shape's text holds math (it is then edited as an equation, not as text). */
export const isEquation = (shape: Element) => hasMath(shape);

/** LaTeX of the shape's first equation, or null. */
export function equationLatex(shape: Element): string | null {
  const m = shape.getElementsByTagNameNS(A14, "m")[0];
  const omml = m?.firstElementChild;
  return omml ? ommlToLatex(new XMLSerializer().serializeToString(omml)).trim() : null;
}

/**
 * Word's OMML (from latexToOmml) as PowerPoint writes it: each m:r takes DrawingML run properties in Cambria Math
 * (italic unless the run is upright text) where Word has w:rPr.
 */
export function slideOmml(omml: string): string {
  const root = parseElement(omml);
  const doc = root.ownerDocument;
  for (const r of Array.from(root.getElementsByTagNameNS(NS.m, "r"))) {
    for (const w of Array.from(r.children).filter((c) => c.namespaceURI === W)) w.remove();
    const mrPr = Array.from(r.children).find((c) => c.namespaceURI === NS.m && c.localName === "rPr");
    const upright = !!mrPr && Array.from(mrPr.children).some((c) => (c.localName === "sty" && c.getAttributeNS(NS.m, "val") === "p") || c.localName === "nor");
    const rPr = doc.createElementNS(NS.a, "a:rPr");
    rPr.setAttribute("lang", "en-US");
    rPr.setAttribute("i", upright ? "0" : "1");
    const latin = rPr.appendChild(doc.createElementNS(NS.a, "a:latin"));
    latin.setAttribute("typeface", "Cambria Math");
    r.insertBefore(rPr, mrPr ? mrPr.nextSibling : r.firstChild);
  }
  root.removeAttributeNS("http://www.w3.org/2000/xmlns/", "w");
  return new XMLSerializer().serializeToString(root);
}

/** The Fallback for an edited equation: the Choice's text box with each equation as its LaTeX. */
function textFallback(shape: Element): Element {
  const copy = shape.cloneNode(true) as Element;
  for (const m of Array.from(copy.getElementsByTagNameNS(A14, "m"))) {
    const omml = m.firstElementChild;
    const latex = omml ? ommlToLatex(new XMLSerializer().serializeToString(omml)) : "";
    const run = parseElement(`<a:r xmlns:a="${NS.a}"><a:rPr lang="en-US"/><a:t>${escapeXml(latex)}</a:t></a:r>`);
    m.replaceWith(copy.ownerDocument.importNode(run, true));
  }
  return copy;
}

/** Replace the shape's first equation with `latex` (display math stays display math); its Fallback becomes text. */
export function setEquation(shape: Element, latex: string): void {
  if (!latex.trim()) throw new Error("An equation can't be empty.");
  const m = shape.getElementsByTagNameNS(A14, "m")[0];
  if (!m) throw new Error("This shape has no equation.");
  const display = m.firstElementChild?.localName === "oMathPara";
  m.replaceChildren(shape.ownerDocument.importNode(parseElement(slideOmml(latexToOmml(latex, display))), true));
  const ac = shape.parentElement?.parentElement;
  const fallback = ac?.localName === "AlternateContent" ? fallbackOf(ac) : null;
  if (fallback) fallback.replaceChildren(textFallback(shape));
}

/** Add an equation text box at `rect` (EMU), as PowerPoint makes one; returns its shape (inside the Choice). */
export function insertEquation(slide: Document, latex: string, rect: Rect): Element {
  if (!latex.trim()) throw new Error("An equation can't be empty.");
  const id = nextShapeId(slide);
  const r = (v: number) => Math.round(v);
  const sp =
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Equation ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${r(rect.x)}" y="${r(rect.y)}"/><a:ext cx="${r(rect.w)}" cy="${r(rect.h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>` +
    `<p:txBody><a:bodyPr wrap="none" rtlCol="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a14:m>${slideOmml(latexToOmml(latex, true))}</a14:m><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`;
  const ac = appendShape(slide, `<mc:AlternateContent xmlns:mc="${NS.mc}"><mc:Choice xmlns:a14="${A14}" Requires="a14">${sp}</mc:Choice><mc:Fallback/></mc:AlternateContent>`);
  const shape = choiceOf(ac)!.firstElementChild!;
  fallbackOf(ac)!.replaceChildren(textFallback(shape));
  return shape;
}

const xfrmXml = (shape: Element | null | undefined) => {
  const xfrm = shape?.getElementsByTagNameNS(NS.a, "xfrm")[0];
  return xfrm ? new XMLSerializer().serializeToString(xfrm) : null;
};

/** Move each equation's Fallback to where its Choice now is. True if any changed. */
export function syncEquationFallbacks(slide: Document): boolean {
  let changed = false;
  for (const ac of Array.from(slide.getElementsByTagNameNS(NS.mc, "AlternateContent"))) {
    const choice = choiceOf(ac)?.firstElementChild;
    const fallback = fallbackOf(ac)?.firstElementChild;
    if (!choice || !fallback || !hasMath(choice)) continue;
    const want = choice.getElementsByTagNameNS(NS.a, "xfrm")[0];
    const have = fallback.getElementsByTagNameNS(NS.a, "xfrm")[0];
    if (!want || !have || xfrmXml(choice) === xfrmXml(fallback)) continue;
    have.replaceWith(slide.importNode(want, true));
    changed = true;
  }
  return changed;
}
