/**
 * Sections: a w:sectPr on a paragraph ends a section (the body's own w:sectPr is the last section's). The editor
 * keeps each one as XML (paragraph / doc attr `sectPr`) and draws a section with two or more columns as a `columns`
 * block around its blocks; `regroup` rebuilds those blocks from the section properties.
 */
import { Node as PMNode } from "prosemirror-model";
import { NS } from "../../core/utils/xml";
import { docxSchema as schema } from "./schema";

export interface Columns {
  count: number;
  /** Space between columns, in twips. */
  space: number;
  /** Line between columns. */
  sep: boolean;
}

export const ONE_COLUMN: Columns = { count: 1, space: 720, sep: false };

/** CT_SectPr child order. */
const SECT_ORDER = ["headerReference", "footerReference", "footnotePr", "endnotePr", "type", "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType", "pgNumType", "cols", "formProt", "vAlign", "noEndnote", "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange"];

const parse = (xml: string) => new DOMParser().parseFromString(xml, "application/xml").documentElement;
const child = (el: Element, local: string) => Array.from(el.children).find((c) => c.namespaceURI === NS.w && c.localName === local) ?? null;
const on = (v: string | null) => v === "1" || v === "true" || v === "on";

export function columnsOf(sectPr: string | null): Columns {
  const cols = sectPr ? child(parse(sectPr), "cols") : null;
  if (!cols) return ONE_COLUMN;
  return {
    count: Math.max(1, Number(cols.getAttributeNS(NS.w, "num") ?? 1) || 1),
    space: Number(cols.getAttributeNS(NS.w, "space") ?? 720) || 0,
    sep: on(cols.getAttributeNS(NS.w, "sep")),
  };
}

/** How a section starts: "nextPage" (the default), "continuous", "evenPage", "oddPage" or "nextColumn". */
export function sectionType(sectPr: string | null): string {
  const type = sectPr ? child(parse(sectPr), "type") : null;
  return type?.getAttributeNS(NS.w, "val") ?? "nextPage";
}

/** Replace (or add, in schema order) one child of a w:sectPr; `null` removes it. */
function withChild(sectPr: string, local: string, make: ((doc: Document) => Element) | null): string {
  const el = parse(sectPr);
  child(el, local)?.remove();
  if (make) {
    const at = SECT_ORDER.indexOf(local);
    const before = Array.from(el.children).find((c) => SECT_ORDER.indexOf(c.localName) > at) ?? null;
    el.insertBefore(make(el.ownerDocument), before);
  }
  return new XMLSerializer().serializeToString(el);
}

export function withColumns(sectPr: string, cols: Columns): string {
  if (cols.count <= 1) return child(parse(sectPr), "cols") ? withChild(sectPr, "cols", null) : sectPr;
  return withChild(sectPr, "cols", (doc) => {
    const el = doc.createElementNS(NS.w, "w:cols");
    el.setAttributeNS(NS.w, "w:num", String(cols.count));
    el.setAttributeNS(NS.w, "w:space", String(Math.round(cols.space)));
    if (cols.sep) el.setAttributeNS(NS.w, "w:sep", "1");
    return el;
  });
}

export function withType(sectPr: string, type: string): string {
  return withChild(sectPr, "type", (doc) => {
    const el = doc.createElementNS(NS.w, "w:type");
    el.setAttributeNS(NS.w, "w:val", type);
    return el;
  });
}

/** An empty section's properties, for a document without any. */
export const DEFAULT_SECTPR = `<w:sectPr xmlns:w="${NS.w}"><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>`;

/** Top-level blocks with column sections opened up. */
export function flatten(doc: PMNode): PMNode[] {
  const out: PMNode[] = [];
  doc.forEach((n) => (n.type === schema.nodes.columns ? n.forEach((c) => out.push(c)) : out.push(n)));
  return out;
}

/** Index of the first and last block of each section, with the properties that govern it (null = the body's). */
export function sections(units: PMNode[]): { start: number; end: number; at: number | null }[] {
  const out: { start: number; end: number; at: number | null }[] = [];
  let start = 0;
  units.forEach((u, i) => {
    if (u.type === schema.nodes.paragraph && u.attrs.sectPr) {
      out.push({ start, end: i, at: i });
      start = i + 1;
    }
  });
  if (start < units.length || !out.length) out.push({ start, end: units.length - 1, at: null });
  return out;
}

/** Top-level blocks: each section with two or more columns wrapped in a `columns` block. */
export function regroup(units: PMNode[], finalSectPr: string | null): PMNode[] {
  const out: PMNode[] = [];
  for (const s of sections(units)) {
    const blocks = units.slice(s.start, s.end + 1);
    if (!blocks.length) continue;
    const cols = columnsOf(s.at === null ? finalSectPr : (units[s.at].attrs.sectPr as string));
    if (cols.count > 1) out.push(schema.nodes.columns.create({ count: cols.count, space: cols.space, sep: cols.sep }, blocks));
    else out.push(...blocks);
  }
  return out;
}

/** Blocks that end a section followed by one starting on a new page. */
export function sectionBreaks(doc: PMNode): Set<PMNode> {
  const units = flatten(doc);
  const out = new Set<PMNode>();
  const secs = sections(units);
  secs.forEach((s, k) => {
    if (s.at === null) return;
    const next = secs[k + 1];
    const nextSectPr = !next ? null : next.at === null ? (doc.attrs.sectPr as string | null) : (units[next.at].attrs.sectPr as string);
    if (sectionType(nextSectPr) !== "continuous") out.add(units[s.at]);
  });
  return out;
}
