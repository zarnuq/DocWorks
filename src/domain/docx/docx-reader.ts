import { Mark, Node as PMNode } from "prosemirror-model";
import { NS, childNS } from "../../core/utils/xml";
import { regroup } from "./sections";
import { placementOf } from "./drawing";
import { docxSchema as schema, Align, TabStop } from "./schema";

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
const CONTAINERS = new Set(["smartTag", "customXml", "dir", "bdo"]);
/** Pure rendering hints, safe to drop when a paragraph is rewritten. */
const DROPPED = new Set(["proofErr", "lastRenderedPageBreak", "pPr", "rPr"]);
const ORDERED_ALIGN: Record<string, Align> = { left: "left", start: "left", center: "center", right: "right", end: "right", both: "justify", distribute: "justify" };

export const wVal = (el: Element | null | undefined, attr = "val") => el?.getAttributeNS(NS.w, attr) ?? null;
export const w = (el: Element | null | undefined, local: string) => (el ? childNS(el, NS.w, local) : null);
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
  rightIndent: number | null;
  tabs: TabStop[] | null;
  spaceBefore: number | null;
  spaceAfter: number | null;
  firstLine: number | null;
  tabLeader: string | null;
  markSize: number | null;
  markFont: string | null;
  keepNext: boolean | null;
  keepLines: boolean | null;
  contextual: boolean | null;
  /** Borders and shading as CSS (display only; the XML keeps the details). */
  frame: string | null;
}

/** w:pBdr / w:shd as CSS: borders (width from eighths of a point, w:space as padding) and the fill. */
function frameCss(pPr: Element | null): string | null {
  const css: string[] = [];
  const bdr = w(pPr, "pBdr");
  for (const side of ["top", "left", "bottom", "right"]) {
    const b = w(bdr, side);
    const val = wVal(b);
    if (!b || !val || val === "none" || val === "nil") continue;
    const px = Math.max(1, Math.round(((Number(wVal(b, "sz") ?? 4) / 8) * 4) / 3 * 100) / 100);
    const style = val === "double" ? "double" : /dash/i.test(val) ? "dashed" : /dot/i.test(val) ? "dotted" : "solid";
    css.push(`border-${side}: ${px}px ${style} ${hex(wVal(b, "color")) ?? "#000000"}`, `padding-${side}: ${Number(wVal(b, "space") ?? 0)}pt`);
  }
  const fill = hex(wVal(w(pPr, "shd"), "fill"));
  if (fill) css.push(`background-color: ${fill}`);
  return css.length ? css.join("; ") : null;
}

const pt = (twips: string | null) => (twips === null ? null : Math.round((Number(twips) / 20) * 100) / 100);

/** Leader of the right-aligned tab stop furthest along, or null. */
function rightTabLeader(tabs: Element | null): string | null {
  const stops = tabs ? Array.from(tabs.children).filter((t) => t.localName === "tab" && wVal(t) !== "clear") : [];
  const last = stops.sort((a, b) => Number(wVal(a, "pos")) - Number(wVal(b, "pos"))).pop();
  return last && ["right", "end"].includes(wVal(last) ?? "") && wVal(last, "leader") ? wVal(last, "leader") : null;
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
    indent: left / 720,
    rightIndent: pt(wVal(ind, "right") ?? wVal(ind, "end")),
    tabs: w(pPr, "tabs") ? Array.from(w(pPr, "tabs")!.children).filter(t => t.localName === "tab").map(t => ({ pt: pt(wVal(t, "pos")) ?? 0, align: wVal(t) ?? "left", leader: wVal(t, "leader") })) : null,
    spaceBefore: pt(wVal(spacing, "before")),
    spaceAfter: pt(wVal(spacing, "after")),
    firstLine: wVal(ind, "hanging") !== null ? -pt(wVal(ind, "hanging"))! : pt(wVal(ind, "firstLine")),
    tabLeader: rightTabLeader(w(pPr, "tabs")),
    markSize: wVal(w(w(pPr, "rPr"), "sz")) ? Number(wVal(w(w(pPr, "rPr"), "sz"))) / 2 : null,
    markFont: wVal(w(w(pPr, "rPr"), "rFonts"), "ascii") ?? wVal(w(w(pPr, "rPr"), "rFonts"), "hAnsi"),
    keepNext: w(pPr, "keepNext") ? isOn(w(pPr, "keepNext")) : null,
    keepLines: w(pPr, "keepLines") ? isOn(w(pPr, "keepLines")) : null,
    contextual: w(pPr, "contextualSpacing") ? isOn(w(pPr, "contextualSpacing")) : null,
    frame: frameCss(pPr),
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

export function vMergeOf(tcPr: Element | null): "restart" | "continue" | null {
  const v = w(tcPr, "vMerge");
  return v ? (wVal(v) === "restart" ? "restart" : "continue") : null;
}

const NO_LINE = new Set(["nil", "none"]);

/** Cell padding as "top right bottom left" in pt ("-" = default), from the cell's w:tcMar, else the table's w:tblCellMar (a table style's comes from DocxStyles.cellPaddingCss). */
function marginsOf(tcPr: Element | null, tblPr: Element | null): string | null {
  const sides = [["top"], ["right", "end"], ["bottom"], ["left", "start"]].map((names) => {
    for (const mar of [w(tcPr, "tcMar"), w(tblPr, "tblCellMar")]) {
      const side = names.map((n) => w(mar, n)).find(Boolean);
      if (side && (wVal(side, "type") ?? "dxa") === "dxa") return `${Number(wVal(side, "w")) / 20}pt`;
    }
    return "-";
  });
  return sides.every((x) => x === "-") ? null : sides.join(" ");
}

/** A cell's own border overrides (w:tcBorders), as "t:none b:line"; null when it has none. Display only. */
export function edgesOf(tcPr: Element | null): string | null {
  const b = w(tcPr, "tcBorders");
  if (!b) return null;
  const side: Record<string, string> = { top: "t", left: "l", start: "l", bottom: "b", right: "r", end: "r" };
  const out = Array.from(b.children)
    .filter((c) => side[c.localName])
    .map((c) => `${side[c.localName]}:${NO_LINE.has(wVal(c) ?? "") ? "none" : "line"}`);
  return out.length ? out.join(" ") : null;
}

/** The border preset a table's own w:tblBorders matches, or null (style borders or something custom). */
export function bordersOf(tblPr: Element | null): "all" | "outer" | "none" | null {
  const b = w(tblPr, "tblBorders");
  if (!b) return null;
  const on = (side: string) => {
    const v = wVal(w(b, side));
    return v === null ? null : !NO_LINE.has(v);
  };
  const outer = ["top", "left", "bottom", "right"].map(on);
  const inner = ["insideH", "insideV"].map(on);
  if (outer.every((x) => x === false) && inner.every((x) => x !== true)) return "none";
  if (outer.every((x) => x === true)) return inner.every((x) => x === true) ? "all" : inner.every((x) => x !== true) ? "outer" : null;
  return null;
}

// -----------------------------------------------------------------------------
// Reader
// -----------------------------------------------------------------------------

class Reader {
  origins: Origins = new Map();
  private nextKey = 1;
  /** Complex fields (w:fldChar begin ... end) the reader is inside, innermost last; fields can span paragraphs. */
  private fields: { instr: string; result: boolean }[] = [];

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
      } else if (el.localName === "sdt" && this.hasBlocks(w(el, "sdtContent"))) {
        out.push(this.sdt(el));
      } else if (el.localName !== "sectPr" && el.localName !== "tblPr" && el.localName !== "tcPr" && el.localName !== "trPr") {
        out.push(this.raw(el));
      }
    }
    return out;
  }

  private hasBlocks(content: Element | null): boolean {
    return !!content && Array.from(content.children).some((c) => c.namespaceURI === NS.w && ["p", "tbl", "sdt"].includes(c.localName));
  }

  private sdt(el: Element): PMNode {
    const gallery = wVal(w(el, "sdtPr")?.getElementsByTagNameNS(NS.w, "docPartGallery")[0]);
    const kind = gallery?.includes("Table of Contents") ? "toc" : "other";
    return this.keyed(el, (key) => schema.nodes.sdt.create({ key, kind }, this.blocks(w(el, "sdtContent")!)));
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
    const sectPr = w(pPr, "sectPr");
    return this.keyed(p, (key) => schema.nodes.paragraph.create({ key, ...paragraphProps(pPr), sectPr: sectPr ? xmlOf(sectPr) : null }, inline));
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
      } else if (el.namespaceURI === NS.w && (local === "ins" || local === "moveTo")) {
        const shell = el.cloneNode(false) as Element;
        const mark = schema.marks.insertion.create({ xml: xmlOf(shell), author: wVal(el, "author") ?? "", date: wVal(el, "date") ?? "" });
        this.inline(el, [...marks, mark], out);
      } else if (el.namespaceURI === NS.w && CONTAINERS.has(local)) {
        // ponytail: wrapper (tracked insertion, smart tag...) is dropped if the paragraph is edited.
        this.inline(el, marks, out);
      } else {
        out.push(this.opaque(el, ["del", "moveFrom", "fldSimple", "oMath", "oMathPara"].includes(el.localName) ? (el.textContent ?? "") : "", marks));
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
    const field = this.fields[this.fields.length - 1];
    const name = field?.result ? /^\s*(PAGE|NUMPAGES)\b/.exec(field.instr)?.[1] : undefined;
    if (name) marks.push(schema.marks.field.create({ name }));
    const text = (s: string) => out.push(schema.text(s, marks));

    for (const child of Array.from(r.children)) {
      const local = child.namespaceURI === NS.w ? child.localName : "";
      if (local === "fldChar") {
        const type = wVal(child, "fldCharType");
        if (type === "begin") this.fields.push({ instr: "", result: false });
        else if (type === "separate" && this.fields.length) this.fields[this.fields.length - 1].result = true;
        else if (type === "end") this.fields.pop();
      } else if (local === "instrText" && this.fields.length) {
        this.fields[this.fields.length - 1].instr += child.textContent ?? "";
      }
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
        ...placementOf(drawing),
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
                    vMerge: vMergeOf(tcPr),
                    edges: edgesOf(tcPr),
                    margins: marginsOf(tcPr, w(tbl, "tblPr")),
                  },
                  content.length ? content : schema.nodes.paragraph.create()
                );
              })
            );
          const trPr = w(tr, "trPr");
          return schema.nodes.table_row.create({ key, header: isOn(w(trPr, "tblHeader")), cantSplit: isOn(w(trPr, "cantSplit")) }, cells);
        })
      );
    // A table without rows can't exist in the editor; keep it verbatim instead.
    if (!rows.length || rows.some((r) => !r.childCount)) return this.raw(tbl);
    const widths = Array.from(w(tbl, "tblGrid")?.children ?? []).map((g) => Number(wVal(g, "w")) || 0);
    return this.keyed(tbl, (key) => schema.nodes.table.create({ key, borders: bordersOf(w(tbl, "tblPr")), widths: widths.length ? widths : null, style: wVal(w(w(tbl, "tblPr"), "tblStyle")) }, rows));
  }
}

/** Read <w:body> into a ProseMirror doc plus the origins needed to write it back losslessly. */
export function readDocx(body: Element, ctx: ReadContext): { doc: PMNode; origins: Origins } {
  const reader = new Reader(ctx);
  const final = Array.from(body.children).filter((c) => c.namespaceURI === NS.w && c.localName === "sectPr").pop();
  const sectPr = final ? xmlOf(final) : null;
  const blocks = regroup(reader.blocks(body), sectPr);
  const doc = schema.nodes.doc.create({ sectPr }, blocks.length ? blocks : schema.nodes.paragraph.create());
  return { doc, origins: reader.origins };
}

/**
 * Whether Word started a page at this paragraph or table when it last laid the document out: its
 * w:lastRenderedPageBreak comes before any of its text.
 */
export function startsWordPage(el: Element): boolean {
  for (const node of Array.from(el.getElementsByTagNameNS(NS.w, "*"))) {
    if (node.localName === "lastRenderedPageBreak") return true;
    if (["t", "tab", "drawing", "pict", "object", "sym", "br"].includes(node.localName)) return false;
  }
  return false;
}
