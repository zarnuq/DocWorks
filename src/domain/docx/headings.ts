/**
 * Headings as Obsidian links name them: [[report.docx#Results]] (or #Parent#Child) opens or embeds that heading's
 * section. Only top-level heading paragraphs count (not headings inside tables or content controls).
 */
import type { Node as PMNode } from "prosemirror-model";
import { attrsOf, docxSchema as schema, HEADING_LEVEL, ParagraphAttrs } from "./schema";

export interface Heading {
  /** Position of the heading paragraph and its index among the document's top-level blocks. */
  pos: number;
  index: number;
  level: number;
  text: string;
}

/** A heading's text as link text: the characters Obsidian links can't hold (# | ^ : [ ] \) dropped, spaces squeezed. */
export function headingLinkText(text: string): string {
  return text.replace(/[#|^:[\]\\]/g, " ").replace(/\s+/g, " ").trim();
}

export function headings(doc: PMNode): Heading[] {
  const out: Heading[] = [];
  doc.forEach((node, pos, index) => {
    if (node.type !== schema.nodes.paragraph) return;
    const { style } = attrsOf<ParagraphAttrs>(node);
    const level = style ? HEADING_LEVEL[style] : undefined;
    if (level && node.textContent.trim()) out.push({ pos, index, level, text: node.textContent.trim() });
  });
  return out;
}

/** The heading "#Heading" (or "#Parent#Child": the last part) names, compared as link text without case. */
export function findHeading(doc: PMNode, subpath: string): Heading | null {
  const wanted = headingLinkText(subpath.split("#").filter(Boolean).pop() ?? "").toLowerCase();
  if (!wanted) return null;
  return headings(doc).find((h) => headingLinkText(h.text).toLowerCase() === wanted) ?? null;
}

/** The heading's section, as document positions: from the heading to the next heading at its level or above. */
export function headingSection(doc: PMNode, subpath: string): { from: number; to: number } | null {
  const h = findHeading(doc, subpath);
  if (!h) return null;
  const next = headings(doc).find((o) => o.pos > h.pos && o.level <= h.level);
  return { from: h.pos, to: next?.pos ?? doc.content.size };
}
