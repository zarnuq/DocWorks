import { Mark, Node as PMNode } from "prosemirror-model";
import { NS, childNS } from "../../core/utils/xml";
import { docxSchema as schema, Align } from "./schema";

export interface ReadContext {
  /** Preview URL for an image relationship. */
  imageSrc(rId: string): string | null;
  hyperlinkHref(rId: string): string | null;
}

/** The XML a node was read from, plus the node as first read (to detect edits). */
export interface Origin {
  el: Element;
  node: PMNode;
}
export type Origins = Map<string, Origin>;

export const EMU_PER_PX = 9525;

/** Word's named highlight colors. */
export const HIGHLIGHTS: Record<string, string> = {
  yellow: "#FFFF00",
  green: "#00FF00",
  cyan: "#00FFFF",
  magenta: "#FF00FF",
  blue: "#0000FF",
  red: "#FF0000",
  darkBlue: "#000080",
  darkCyan: "#008080",
  darkGreen: "#008000",
  darkMagenta: "#800080",
  darkRed: "#800000",
  darkYellow: "#808000",
  darkGray: "#808080",
  lightGray: "#C0C0C0",
  black: "#000000",
  white: "#FFFFFF",
};

/** Paragraph / cell children we read as content rather than keep opaque. */
const CONTAINERS = new Set(["ins", "smartTag", "customXml", "dir", "bdo"]);
/** Pure rendering hints, safe to drop when a paragraph is rewritten. */
const DROPPED = new Set(["proofErr", "lastRenderedPageBreak", "pPr", "rPr"]);
const ORDERED_ALIGN: Record<string, Align> = { left: "left", start: "left", center: "center", right: "right", end: "right", both: "justify", distribute: "justify" };

export const wVal = (el: Element | null | undefined, attr = "val") => el?.getAttributeNS(NS.w, attr) ?? null;
const w = (el: Element | null | undefined, local: string) => (el ? childNS(el, NS.w, local) : null);
const isOn = (el: Element | null) => !!el && !["0", "false", "off"].includes(wVal(el) ?? "");
const hex = (v: string | null) => (v && /^[0-9a-f]{6}$/i.test(v) ? `#${v.toUpperCase()}` : null);
const xmlOf = (el: Element) => new XMLSerializer().serializeToString(el);

// -----------------------------------------------------------------------------
// Properties shared with the writer (it compares against what we read here)
// -----------------------------------------------------------------------------

export interface ParagraphProps {
  style: string | null;
  align: Align | null;
  numId: string | null;
  ilvl: number;
  lineSpacing: string | null;
  indent: number;
}

export function paragraphProps(pPr: Element | null): ParagraphProps {
  const numPr = w(pPr, "numPr");
  const numId = wVal(w(numPr, "numId"));
  const spacing = w(pPr, "spacing");
  const line = wVal(spacing, "line");
  const rule = wVal(spacing, "lineRule");
  const ind = w(pPr, "ind");
  const left = Number(wVal(ind, "left") ?? wVal(ind, "start") ?? 0);
  return {
    style: wVal(w(pPr, "pStyle")),
    align: ORDERED_ALIGN[wVal(w(pPr, "jc")) ?? ""] ?? null,
    numId: numId && numId !== "0" ? numId : null,
    ilvl: Number(wVal(w(numPr, "ilvl")) ?? 0),
    lineSpacing: line && (!rule || rule === "auto") ? String(Math.round((Number(line) / 240) * 100) / 100) : null,
    indent: left > 0 ? Math.round(left / 720) : 0,
  };
}

/** Marks for the formatting set directly in a w:rPr (styles are display-only, see docx-styles). */
export function runMarks(rPr: Element | null): Mark[] {
  if (!rPr) return [];
  const m = schema.marks;
  const marks: Mark[] = [m.rpr.create({ xml: xmlOf(rPr) })];
  if (isOn(w(rPr, "b"))) marks.push(m.bold.create());
  if (isOn(w(rPr, "i"))) marks.push(m.italic.create());
  const u = w(rPr, "u");
  if (u && wVal(u) !== "none") marks.push(m.underline.create());
  if (isOn(w(rPr, "strike")) || isOn(w(rPr, "dstrike"))) marks.push(m.strike.create());
  const vert = wVal(w(rPr, "vertAlign"));
  if (vert === "subscript") marks.push(m.subscript.create());
  if (vert === "superscript") marks.push(m.superscript.create());
  const fonts = w(rPr, "rFonts");
  const font = wVal(fonts, "ascii") ?? wVal(fonts, "hAnsi");
  if (font) marks.push(m.font.create({ font }));
  const sz = wVal(w(rPr, "sz"));
  if (sz) marks.push(m.size.create({ pt: Number(sz) / 2 }));
  const color = hex(wVal(w(rPr, "color")));
  if (color) marks.push(m.color.create({ color }));
  const highlight = HIGHLIGHTS[wVal(w(rPr, "highlight")) ?? ""] ?? hex(wVal(w(rPr, "shd"), "fill"));
  if (highlight) marks.push(m.highlight.create({ color: highlight }));
  return marks;
}

// -----------------------------------------------------------------------------
// Reader
// -----------------------------------------------------------------------------

class Reader {
  origins: Origins = new Map();
  private nextKey = 1;

  constructor(private ctx: ReadContext) {}

  private keyed(el: Element, make: (key: string) => PMNode): PMNode {
    const key = `k${this.nextKey++}`;
    const node = make(key);
    this.origins.set(key, { el, node });
    return node;
  }

  blocks(container: Element): PMNode[] {
    const out: PMNode[] = [];
    for (const el of Array.from(container.children)) {
      if (el.namespaceURI !== NS.w) {
        out.push(this.raw(el));
      } else if (el.localName === "p") {
        out.push(this.paragraph(el));
      } else if (el.localName === "tbl") {
        out.push(this.table(el));
      } else if (el.localName !== "sectPr" && el.localName !== "tblPr" && el.localName !== "tcPr" && el.localName !== "trPr") {
        out.push(this.raw(el));
      }
    }
    return out;
  }

  private raw(el: Element): PMNode {
    return schema.nodes.raw_block.create({ xml: xmlOf(el), text: el.textContent ?? "" });
  }

  private paragraph(p: Element): PMNode {
    const pPr = w(p, "pPr");
    const bdr = w(pPr, "pBdr");
    // An empty paragraph with only a bottom border is how Word draws a horizontal rule.
    if (w(bdr, "bottom") && !p.getElementsByTagNameNS(NS.w, "t").length) {
      return schema.nodes.raw_block.create({ xml: xmlOf(p), kind: "hr" });
    }
    const inline: PMNode[] = [];
    this.inline(p, [], inline);
    return this.keyed(p, (key) => schema.nodes.paragraph.create({ key, ...paragraphProps(pPr) }, inline));
  }

  private inline(container: Element, marks: readonly Mark[], out: PMNode[]): void {
    for (const el of Array.from(container.children)) {
      const local = el.localName;
      if (el.namespaceURI === NS.w && DROPPED.has(local)) continue;
      if (el.namespaceURI === NS.w && local === "r") {
        this.run(el, marks, out);
      } else if (el.namespaceURI === NS.w && local === "hyperlink") {
        const rId = el.getAttributeNS(NS.r, "id");
        const anchor = wVal(el, "anchor");
        const href = (rId && this.ctx.hyperlinkHref(rId)) ?? null;
        this.inline(el, [...marks, schema.marks.link.create({ rId, anchor, href })], out);
      } else if (el.namespaceURI === NS.w && local === "sdt") {
        const content = w(el, "sdtContent");
        if (content) this.inline(content, marks, out);
      } else if (el.namespaceURI === NS.w && CONTAINERS.has(local)) {
        // ponytail: wrapper (tracked insertion, smart tag...) is dropped if the paragraph is edited.
        this.inline(el, marks, out);
      } else {
        out.push(this.opaque(el, el.localName === "del" || el.localName === "fldSimple" ? (el.textContent ?? "") : "", marks));
      }
    }
  }

  private opaque(el: Element, text: string, marks: readonly Mark[]): PMNode {
    const link = marks.filter((m) => m.type === schema.marks.link);
    return schema.nodes.opaque.create({ xml: xmlOf(el), text }, null, link);
  }

  private run(r: Element, outer: readonly Mark[], out: PMNode[]): void {
    const rPr = w(r, "rPr");
    const marks = [...outer, ...runMarks(rPr)];
    const text = (s: string) => out.push(schema.text(s, marks));

    for (const child of Array.from(r.children)) {
      const local = child.namespaceURI === NS.w ? child.localName : "";
      if (local === "t") {
        if (child.textContent) text(child.textContent);
      } else if (local === "tab") {
        text("\t");
      } else if (local === "noBreakHyphen") {
        text("‑");
      } else if (local === "softHyphen") {
        text("­");
      } else if (local === "br" && wVal(child, "type") === "page") {
        out.push(schema.nodes.page_break.create());
      } else if (local === "cr" || (local === "br" && wVal(child, "type") !== "column")) {
        out.push(schema.nodes.hard_break.create(null, null, marks));
      } else if (local === "drawing" && child.getElementsByTagNameNS(NS.a, "blip").length) {
        out.push(this.image(r, child, outer));
      } else if (!DROPPED.has(local)) {
        // Fields, footnote references, symbols, shapes... keep the whole run around this one child.
        const single = r.cloneNode(false) as Element;
        if (rPr) single.appendChild(rPr.cloneNode(true));
        single.appendChild(child.cloneNode(true));
        const label = local === "sym" ? "□" : local === "footnoteReference" || local === "endnoteReference" ? "*" : "";
        out.push(this.opaque(single, label, outer));
      }
    }
  }

  private image(r: Element, drawing: Element, marks: readonly Mark[]): PMNode {
    const rId = drawing.getElementsByTagNameNS(NS.a, "blip")[0].getAttributeNS(NS.r, "embed");
    const extent = drawing.getElementsByTagNameNS(NS.wp, "extent")[0];
    const px = (attr: string) => (extent ? Math.round(Number(extent.getAttribute(attr)) / EMU_PER_PX) : null);
    const single = r.cloneNode(false) as Element;
    const rPr = w(r, "rPr");
    if (rPr) single.appendChild(rPr.cloneNode(true));
    single.appendChild(drawing.cloneNode(true));
    return schema.nodes.image.create(
      {
        xml: xmlOf(single),
        rId,
        src: (rId && this.ctx.imageSrc(rId)) ?? "",
        width: px("cx"),
        height: px("cy"),
        alt: drawing.getElementsByTagNameNS(NS.wp, "docPr")[0]?.getAttribute("descr") ?? "",
      },
      null,
      marks.filter((m) => m.type === schema.marks.link)
    );
  }

  private table(tbl: Element): PMNode {
    const rows = Array.from(tbl.children)
      .filter((tr) => tr.namespaceURI === NS.w && tr.localName === "tr")
      .map((tr) =>
        this.keyed(tr, (key) => {
          const cells = Array.from(tr.children)
            .filter((tc) => tc.namespaceURI === NS.w && tc.localName === "tc")
            .map((tc) =>
              this.keyed(tc, (cellKey) => {
                const tcPr = w(tc, "tcPr");
                const content = this.blocks(tc);
                return schema.nodes.table_cell.create(
                  {
                    key: cellKey,
                    colspan: Number(wVal(w(tcPr, "gridSpan")) ?? 1),
                    background: hex(wVal(w(tcPr, "shd"), "fill")),
                  },
                  content.length ? content : schema.nodes.paragraph.create()
                );
              })
            );
          return schema.nodes.table_row.create({ key }, cells);
        })
      );
    // A table without rows can't exist in the editor; keep it verbatim instead.
    if (!rows.length || rows.some((r) => !r.childCount)) return this.raw(tbl);
    return this.keyed(tbl, (key) => schema.nodes.table.create({ key }, rows));
  }
}

/** Read <w:body> into a ProseMirror doc plus the origins needed to write it back losslessly. */
export function readDocx(body: Element, ctx: ReadContext): { doc: PMNode; origins: Origins } {
  const reader = new Reader(ctx);
  const blocks = reader.blocks(body);
  const doc = schema.nodes.doc.create(null, blocks.length ? blocks : schema.nodes.paragraph.create());
  return { doc, origins: reader.origins };
}
