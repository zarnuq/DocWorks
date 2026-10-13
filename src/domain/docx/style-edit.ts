/*
 * Google Docs' "Update '<style>' to match": a paragraph's own formatting (alignment, spacing, indents; font, size,
 * colour, bold / italic / underline of its text) becomes its style's, so every paragraph with the style follows.
 * The paragraph is written with the normal writer into a scratch body and its pPr / rPr merged into the w:style.
 */
import { Node as PMNode } from "prosemirror-model";
import { NS } from "../../core/utils/xml";
import type { DocxDocument } from "../../infrastructure/docx/docx-document";
import { ORDER, writeBody } from "./docx-writer";
import { docxSchema as schema, ParagraphAttrs, attrsOf } from "./schema";

/** CT_Style child order. */
const STYLE_ORDER = ["name", "aliases", "basedOn", "next", "link", "autoRedefine", "hidden", "uiPriority", "semiHidden", "unhideWhenUsed", "qFormat", "locked", "personal", "personalCompose", "personalReply", "rsid", "pPr", "rPr", "tblPr", "trPr", "tcPr", "tblStylePr"];
/** Paragraph properties the style takes over; the rest (numbering, section, style) belong to the paragraph. */
const SKIP = new Set(["pStyle", "numPr", "rPr", "sectPr", "pPrChange"]);
/** Elements whose attributes are merged rather than replaced (spacing before, keeping the style's line spacing). */
const MERGE = new Set(["spacing", "ind", "rFonts"]);
export const STYLE_MARKS = ["bold", "italic", "underline", "strike", "font", "size", "color"] as const;

const kids = (el: Element) => Array.from(el.children);

function place(parent: Element, el: Element, order: string[]): void {
  const at = order.indexOf(el.localName);
  const next = kids(parent).find((c) => order.indexOf(c.localName) > at) ?? null;
  parent.insertBefore(el, next);
}

function own(parent: Element, local: string, order: string[]): Element {
  const found = kids(parent).find((c) => c.localName === local && c.namespaceURI === NS.w);
  if (found) return found;
  const el = parent.ownerDocument.createElementNS(NS.w, `w:${local}`);
  place(parent, el, order);
  return el;
}

function merge(target: Element, from: Element, order: string[]): void {
  for (const c of kids(from)) {
    if (SKIP.has(c.localName)) continue;
    const old = kids(target).find((x) => x.localName === c.localName);
    const el = target.ownerDocument.importNode<Element>(c, true);
    if (old && MERGE.has(c.localName)) {
      for (const a of Array.from(el.attributes)) old.setAttributeNS(a.namespaceURI, a.name, a.value);
      continue;
    }
    old?.remove();
    place(target, el, order);
  }
}

/** The paragraph's direct formatting: its paragraph attributes and the marks on its first text. */
export function directFormat(para: PMNode): { attrs: Partial<ParagraphAttrs>; marks: PMNode["marks"] } {
  const a = attrsOf<ParagraphAttrs>(para);
  let marks: PMNode["marks"] = [];
  para.descendants((n) => {
    if (n.isText && !marks.length) marks = n.marks.filter((m) => (STYLE_MARKS as readonly string[]).includes(m.type.name));
    return !marks.length;
  });
  return { attrs: { align: a.align, lineSpacing: a.lineSpacing, spaceBefore: a.spaceBefore, spaceAfter: a.spaceAfter, firstLine: a.firstLine, indent: a.indent, rightIndent: a.rightIndent }, marks };
}

/** Merge `para`'s direct formatting into paragraph style `styleId` (defined first if it's a built-in one). */
export function updateStyleToMatch(docx: DocxDocument, styleId: string, para: PMNode): void {
  const styles = docx.stylesXml();
  if (!styles) return;
  docx.ensureParagraphStyles([styleId]);
  const style = Array.from(styles.getElementsByTagNameNS(NS.w, "style")).find((s) => s.getAttributeNS(NS.w, "type") === "paragraph" && s.getAttributeNS(NS.w, "styleId") === styleId);
  if (!style) return;
  const { attrs, marks } = directFormat(para);
  const scratch = styles.createElementNS(NS.w, "w:body");
  const p = schema.nodes.paragraph.create({ ...attrs, style: null }, schema.text("x", marks));
  writeBody(scratch, schema.nodes.doc.create(null, p), new Map(), { nextDocPrId: () => 1 });
  const wp = scratch.getElementsByTagNameNS(NS.w, "p")[0];
  const pPr = wp && kids(wp).find((c) => c.localName === "pPr");
  const rPr = wp?.getElementsByTagNameNS(NS.w, "r")[0] && kids(wp.getElementsByTagNameNS(NS.w, "r")[0]).find((c) => c.localName === "rPr");
  if (pPr && kids(pPr).some((c) => !SKIP.has(c.localName))) merge(own(style, "pPr", STYLE_ORDER), pPr, ORDER.pPr);
  if (rPr?.children.length) merge(own(style, "rPr", STYLE_ORDER), rPr, ORDER.rPr);
  const path = docx.relOfType(docx.mainPath, "styles");
  if (path) docx.touch(path);
}
