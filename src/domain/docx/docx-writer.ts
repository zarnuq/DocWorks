import { Mark, Node as PMNode } from "prosemirror-model";
import { NS, escapeXml } from "../../core/utils/xml";
import { attrsOf, CellAttrs, docxSchema as schema, ImageAttrs, LinkAttrs, ParagraphAttrs, TableAttrs, TableBorders, XmlAttrs } from "./schema";
import { bordersOf, EMU_PER_PX, HIGHLIGHTS, Origin, Origins, paragraphProps, runMarks, vMergeOf, w, wVal } from "./docx-reader";

export interface WriteContext {
  /** Unique id for a new wp:docPr. */
  nextDocPrId(): number;
  /** Relationship id of an external link to `href`, added if needed. */
  hyperlinkRel?(href: string): string;
  /** Define paragraph styles that edited paragraphs now use (headings missing from the document...). */
  ensureParagraphStyles?(ids: Iterable<string>): void;
}

/*
 * Child order of the OOXML property elements we touch. Word rejects files whose
 * children are out of schema order, so new elements are inserted in place.
 */
const ORDER: Record<string, string[]> = {
  pPr: ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"],
  rPr: ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath"],
  numPr: ["ilvl", "numId"],
  tcPr: ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark"],
  trPr: ["cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore", "wAfter", "cantSplit", "trHeight", "tblHeader", "tblCellSpacing", "jc", "hidden", "ins", "del", "trPrChange"],
  tblPr: ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook", "tblCaption", "tblDescription", "tblPrChange"],
};

const HIGHLIGHT_NAMES = Object.fromEntries(Object.entries(HIGHLIGHTS).map(([name, hex]) => [hex, name]));
const MODELED_MARKS = ["bold", "italic", "underline", "strike", "subscript", "superscript", "font", "size", "color", "highlight"] as const;

const parse = (xml: string) => new DOMParser().parseFromString(xml, "application/xml").documentElement;

class Writer {
  private used = new Set<string>();
  /** Paragraph styles newly applied by this write. */
  readonly newStyles = new Set<string>();
  private remaining = new Map<string, number>();

  constructor(private owner: Document, private origins: Origins, private ctx: WriteContext, doc: PMNode) {
    doc.descendants((n) => {
      const { key } = attrsOf<{ key?: string | null }>(n);
      if (key) this.remaining.set(key, (this.remaining.get(key) ?? 0) + 1);
    });
  }

  // --- DOM helpers ----------------------------------------------------------

  private el(local: string): Element {
    return this.owner.createElementNS(NS.w, `w:${local}`);
  }

  private set(el: Element, attr: string, value: string): void {
    el.setAttributeNS(NS.w, `w:${attr}`, value);
  }

  private import(xml: string): Element {
    return this.owner.importNode(parse(xml), true);
  }

  /** The child `local` of `parent`, created at its schema position if missing. */
  private upsert(parent: Element, local: string): Element {
    const existing = w(parent, local);
    if (existing) return existing;
    const order = ORDER[parent.localName] ?? [];
    const rank = order.indexOf(local);
    const child = this.el(local);
    const before = Array.from(parent.children).find((c) => order.indexOf(c.localName) > rank) ?? null;
    parent.insertBefore(child, before);
    return child;
  }

  private drop(parent: Element, ...locals: string[]): void {
    for (const local of locals) w(parent, local)?.remove();
  }

  /**
   * Find the original XML of a keyed node. `verbatim` means it can be copied
   * as-is; `first` / `last` tell split or copied nodes apart (ids, section breaks).
   */
  private claim(node: PMNode): { o: Origin | null; verbatim: boolean; first: boolean; last: boolean } {
    const key = node.attrs.key as string | null;
    const o = key ? this.origins.get(key) ?? null : null;
    if (!key || !o) return { o: null, verbatim: false, first: true, last: true };
    const left = (this.remaining.get(key) ?? 1) - 1;
    this.remaining.set(key, left);
    const first = !this.used.has(key);
    this.used.add(key);
    const last = left === 0;
    return { o, verbatim: first && last && node.eq(o.node), first, last };
  }

  private shell(local: string, c: { o: Origin | null; first: boolean }): Element {
    const el = this.el(local);
    // Only the first copy keeps attributes like w14:paraId, which must stay unique.
    if (c.o && c.first) for (const a of Array.from(c.o.el.attributes)) el.setAttributeNS(a.namespaceURI, a.name, a.value);
    return el;
  }

  // --- Blocks ---------------------------------------------------------------

  blocks(container: PMNode): Element[] {
    const out: Element[] = [];
    // A columns block is display only: its section's columns live in the section's w:sectPr.
    container.forEach((node) => (node.type === schema.nodes.columns ? out.push(...this.blocks(node)) : out.push(this.block(node))));
    return out;
  }

  private block(node: PMNode): Element {
    if (node.type === schema.nodes.table) return this.table(node);
    if (node.type === schema.nodes.sdt) return this.sdt(node);
    if (node.type === schema.nodes.raw_block) {
      const { xml } = attrsOf<XmlAttrs>(node);
      if (xml) return this.import(xml);
      return this.import(
        `<w:p xmlns:w="${NS.w}"><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr></w:pPr></w:p>`
      );
    }
    return this.paragraph(node);
  }

  /** A content control: its original properties around the rewritten content. */
  private sdt(node: PMNode): Element {
    const c = this.claim(node);
    if (c.verbatim) return c.o!.el.cloneNode(true) as Element;
    const sdt = c.o ? this.shell("sdt", c) : this.import(attrsOf<XmlAttrs>(node).xml || `<w:sdt xmlns:w="${NS.w}"/>`);
    for (const child of Array.from(c.o?.el.children ?? [])) if (child.localName !== "sdtContent") sdt.appendChild(child.cloneNode(true));
    const content = sdt.appendChild(this.el("sdtContent"));
    for (const el of this.blocks(node)) content.appendChild(el);
    return sdt;
  }

  private paragraph(node: PMNode): Element {
    const c = this.claim(node);
    if (c.verbatim) return c.o!.el.cloneNode(true) as Element;

    const p = this.shell("p", c);
    const base = w(c.o?.el, "pPr");
    const pPr = (base?.cloneNode(true) as Element | undefined) ?? this.el("pPr");
    const was = paragraphProps(base);
    const a = attrsOf<ParagraphAttrs>(node);
    // A section break belongs to the last paragraph of a split; an edited one is written from its XML.
    const sectPr = w(base, "sectPr");
    const wasSect = sectPr ? new XMLSerializer().serializeToString(sectPr) : null;
    if (!c.last || a.sectPr !== wasSect) {
      this.drop(pPr, "sectPr");
      if (a.sectPr && (c.last || a.sectPr !== wasSect)) pPr.insertBefore(this.import(a.sectPr), w(pPr, "pPrChange"));
    }

    if (a.style !== was.style) {
      if (a.style) this.newStyles.add(a.style);
      if (a.style) this.set(this.upsert(pPr, "pStyle"), "val", a.style);
      else this.drop(pPr, "pStyle");
    }
    if (a.numId !== was.numId || a.ilvl !== was.ilvl) {
      this.drop(pPr, "numPr");
      // ilvl without numId nests a list that comes from the paragraph style.
      if (a.numId || a.ilvl > 0) {
        const numPr = this.upsert(pPr, "numPr");
        this.set(this.upsert(numPr, "ilvl"), "val", String(a.ilvl));
        if (a.numId) this.set(this.upsert(numPr, "numId"), "val", a.numId);
      }
    }
    if (a.align !== was.align) {
      if (a.align) this.set(this.upsert(pPr, "jc"), "val", a.align === "justify" ? "both" : a.align);
      else this.drop(pPr, "jc");
    }
    if (a.lineSpacing !== was.lineSpacing) {
      const spacing = this.upsert(pPr, "spacing");
      if (a.lineSpacing) {
        this.set(spacing, "line", String(Math.round(Number(a.lineSpacing) * 240)));
        this.set(spacing, "lineRule", "auto");
      } else {
        spacing.removeAttributeNS(NS.w, "line");
        spacing.removeAttributeNS(NS.w, "lineRule");
        if (!spacing.attributes.length) spacing.remove();
      }
    }
    if (a.indent !== was.indent) {
      const ind = this.upsert(pPr, "ind");
      ind.removeAttributeNS(NS.w, "start");
      if (a.indent > 0) this.set(ind, "left", String(a.indent * 720));
      else ind.removeAttributeNS(NS.w, "left");
      if (!ind.attributes.length) ind.remove();
    }

    const twips = (v: number) => String(Math.round(v * 20));
    for (const side of ["spaceBefore", "spaceAfter"] as const) {
      if (a[side] === was[side]) continue;
      const attr = side === "spaceBefore" ? "before" : "after";
      const spacing = this.upsert(pPr, "spacing");
      // An explicit value replaces Word's automatic spacing.
      spacing.removeAttributeNS(NS.w, `${attr}Autospacing`);
      if (a[side] !== null) this.set(spacing, attr, twips(a[side]));
      else spacing.removeAttributeNS(NS.w, attr);
      if (!spacing.attributes.length) spacing.remove();
    }
    if (a.firstLine !== was.firstLine) {
      const ind = this.upsert(pPr, "ind");
      ind.removeAttributeNS(NS.w, "firstLine");
      ind.removeAttributeNS(NS.w, "hanging");
      if (a.firstLine) this.set(ind, a.firstLine > 0 ? "firstLine" : "hanging", twips(Math.abs(a.firstLine)));
      if (!ind.attributes.length) ind.remove();
    }

    for (const key of ["keepNext", "keepLines"] as const) {
      if (a[key] === was[key]) continue;
      this.drop(pPr, key);
      if (a[key] !== null) {
        const el = this.upsert(pPr, key);
        if (!a[key]) this.set(el, "val", "0");
      }
    }

    if (pPr.attributes.length || pPr.children.length) p.appendChild(pPr);
    this.inline(node, p);
    return p;
  }

  private table(node: PMNode): Element {
    const c = this.claim(node);
    if (c.verbatim) return c.o!.el.cloneNode(true) as Element;

    const tbl = this.shell("tbl", c);
    const tblPr = w(c.o?.el, "tblPr")?.cloneNode(true) as Element | undefined;
    tbl.appendChild(
      tblPr ??
        this.import(
          `<w:tblPr xmlns:w="${NS.w}"><w:tblW w:w="0" w:type="auto"/><w:tblBorders>` +
            ["top", "left", "bottom", "right", "insideH", "insideV"].map((s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join("") +
            `</w:tblBorders><w:tblLook w:val="04A0"/></w:tblPr>`
        )
    );

    const ta = attrsOf<TableAttrs>(node);
    const tblPrEl = tbl.lastElementChild!;
    if (ta.borders && ta.borders !== bordersOf(w(c.o?.el, "tblPr"))) this.setBorders(tblPrEl, ta.borders);

    // ponytail: non-row children of an edited table (bookmarks between rows) are dropped.
    let cols = 0;
    node.forEach((row) => {
      let n = 0;
      row.forEach((cell) => (n += attrsOf<CellAttrs>(cell).colspan));
      cols = Math.max(cols, n);
    });
    const grid = (w(c.o?.el, "tblGrid")?.cloneNode(true) as Element | undefined) ?? this.el("tblGrid");
    const original = Array.from(grid.children).map((g) => Number(wVal(g, "w")) || 0);
    const widthsChanged = !!ta.widths && ta.widths.length === cols && ta.widths.some((x, i) => x !== original[i]);
    const widths = widthsChanged ? ta.widths! : original;
    if (widths.length !== cols || widthsChanged) {
      const avg = widths.length ? Math.round(widths.reduce((s, x) => s + x, 0) / widths.length) : Math.round(9000 / cols);
      grid.replaceChildren(
        ...Array.from({ length: cols }, (_, i) => {
          const col = this.el("gridCol");
          this.set(col, "w", String(widths[i] ?? avg));
          return col;
        })
      );
    }
    tbl.appendChild(grid);
    const colWidths = Array.from(grid.children).map((g) => Number(wVal(g, "w")) || 0);
    node.forEach((row) => tbl.appendChild(this.row(row, colWidths, widthsChanged)));
    return tbl;
  }

  private setBorders(tblPr: Element, kind: TableBorders): void {
    this.drop(tblPr, "tblBorders");
    const borders = this.upsert(tblPr, "tblBorders");
    for (const side of ["top", "left", "bottom", "right", "insideH", "insideV"]) {
      const b = borders.appendChild(this.el(side));
      const on = kind === "all" || (kind === "outer" && !side.startsWith("inside"));
      this.set(b, "val", on ? "single" : "nil");
      if (on) {
        this.set(b, "sz", "4");
        this.set(b, "space", "0");
        this.set(b, "color", "auto");
      }
    }
  }

  private row(node: PMNode, colWidths: number[], widthsChanged: boolean): Element {
    const c = this.claim(node);
    if (c.verbatim && !widthsChanged) return c.o!.el.cloneNode(true) as Element;
    const tr = this.shell("tr", c);
    const base = w(c.o?.el, "trPr");
    const trPr = (base?.cloneNode(true) as Element | undefined) ?? this.el("trPr");
    const header = !!node.attrs.header;
    const wasHeader = !!w(base, "tblHeader") && !["0", "false", "off"].includes(wVal(w(base, "tblHeader")) ?? "");
    if (header !== wasHeader) {
      this.drop(trPr, "tblHeader");
      if (header) this.upsert(trPr, "tblHeader");
    }
    const cantSplit = !!node.attrs.cantSplit;
    const couldNotSplit = !!w(base, "cantSplit") && !["0", "false", "off"].includes(wVal(w(base, "cantSplit")) ?? "");
    if (cantSplit !== couldNotSplit) {
      this.drop(trPr, "cantSplit");
      if (cantSplit) this.upsert(trPr, "cantSplit");
    }
    if (trPr.children.length || trPr.attributes.length) tr.appendChild(trPr);
    let col = 0;
    node.forEach((cell) => {
      const span = attrsOf<CellAttrs>(cell).colspan;
      const width = colWidths.slice(col, col + span).reduce((s, x) => s + x, 0);
      col += span;
      tr.appendChild(this.cell(cell, width, widthsChanged));
    });
    return tr;
  }

  private cell(node: PMNode, width: number, widthsChanged = false): Element {
    const c = this.claim(node);
    if (c.verbatim && !widthsChanged) return c.o!.el.cloneNode(true) as Element;
    const tc = this.shell("tc", c);
    const base = w(c.o?.el, "tcPr");
    const tcPr = (base?.cloneNode(true) as Element | undefined) ?? this.el("tcPr");
    const wasSpan = Number(wVal(w(base, "gridSpan")) ?? 1);
    const wasFill = wVal(w(base, "shd"), "fill");
    const wasBg = wasFill && /^[0-9a-f]{6}$/i.test(wasFill) ? `#${wasFill.toUpperCase()}` : null;

    // New cells get an explicit width, as Word writes them; resized columns update theirs.
    if ((!base || (widthsChanged && w(tcPr, "tcW"))) && width) {
      const tcW = this.upsert(tcPr, "tcW");
      this.set(tcW, "w", String(width));
      this.set(tcW, "type", "dxa");
    }
    if (node.attrs.colspan !== wasSpan) {
      if (node.attrs.colspan > 1) this.set(this.upsert(tcPr, "gridSpan"), "val", String(node.attrs.colspan));
      else this.drop(tcPr, "gridSpan");
    }
    const vMerge = attrsOf<CellAttrs>(node).vMerge;
    if (vMerge !== vMergeOf(base)) {
      this.drop(tcPr, "vMerge");
      if (vMerge) {
        const v = this.upsert(tcPr, "vMerge");
        if (vMerge === "restart") this.set(v, "val", "restart");
      }
    }
    if (node.attrs.background !== wasBg) {
      this.drop(tcPr, "shd");
      const { background } = attrsOf<CellAttrs>(node);
      if (background) this.shade(this.upsert(tcPr, "shd"), background);
    }
    if (tcPr.children.length || tcPr.attributes.length) tc.appendChild(tcPr);
    for (const el of this.blocks(node)) tc.appendChild(el);
    // Word requires every cell to end with a paragraph.
    if (tc.lastElementChild?.localName !== "p") tc.appendChild(this.el("p"));
    return tc;
  }

  private shade(shd: Element, color: string): void {
    this.set(shd, "val", "clear");
    this.set(shd, "color", "auto");
    this.set(shd, "fill", color.replace("#", "").toUpperCase());
  }

  // --- Inline content -------------------------------------------------------

  private inline(node: PMNode, p: Element): void {
    // Runs are grouped into w:ins wrappers (tracked insertions), and links inside them.
    let ins: { mark: Mark; el: Element } | null = null;
    let group: { link: Mark; el: Element; parent: Element } | null = null;
    node.forEach((child) => {
      const insMark = child.marks.find((m) => m.type === schema.marks.insertion) ?? null;
      let target = p;
      if (insMark) {
        if (!ins || !ins.mark.eq(insMark)) {
          ins = { mark: insMark, el: p.appendChild(this.import(attrsOf<{ xml: string }>(insMark).xml)) };
          group = null;
        }
        target = ins.el;
      } else {
        ins = null;
      }
      const link = child.marks.find((m) => m.type === schema.marks.link) ?? null;
      if (link) {
        if (!group || !group.link.eq(link) || group.parent !== target) {
          const h = this.el("hyperlink");
          const { rId, anchor, href } = attrsOf<LinkAttrs>(link);
          const id = rId ?? (href && !anchor ? this.ctx.hyperlinkRel?.(href) : null);
          if (id) h.setAttributeNS(NS.r, "r:id", id);
          if (anchor) this.set(h, "anchor", anchor);
          this.set(h, "history", "1");
          target.appendChild(h);
          group = { link, el: h, parent: target };
        }
        target = group.el;
      } else {
        group = null;
      }
      target.appendChild(this.inlineNode(child));
    });
  }

  private inlineNode(node: PMNode): Element {
    if (node.type === schema.nodes.opaque) return this.import(attrsOf<XmlAttrs>(node).xml);
    if (node.type === schema.nodes.image) return this.image(node);

    const r = this.el("r");
    const rPr = this.runProps(node.marks);
    if (rPr) r.appendChild(rPr);
    if (node.type === schema.nodes.page_break) {
      this.set(r.appendChild(this.el("br")), "type", "page");
    } else if (node.type === schema.nodes.hard_break) {
      r.appendChild(this.el("br"));
    } else {
      node.text!.split("\t").forEach((part, i) => {
        if (i > 0) r.appendChild(this.el("tab"));
        if (!part) return;
        const t = this.el("t");
        if (/^\s|\s$/.test(part)) t.setAttributeNS(NS.xml, "xml:space", "preserve");
        t.textContent = part;
        r.appendChild(t);
      });
    }
    return r;
  }

  /** Original w:rPr with only the modeled properties that changed rewritten. */
  private runProps(marks: readonly Mark[]): Element | null {
    const source = marks.find((m) => m.type === schema.marks.rpr);
    const sourceXml = source ? attrsOf<{ xml: string }>(source).xml : null;
    const rPr = sourceXml ? this.import(sourceXml) : this.el("rPr");
    const was = runMarks(sourceXml ? parse(sourceXml) : null);

    for (const name of MODELED_MARKS) {
      const type = schema.marks[name];
      const now = marks.find((m) => m.type === type);
      const before = was.find((m) => m.type === type);
      if (now === before || (now && before && now.eq(before))) continue;
      switch (name) {
        case "bold":
        case "italic": {
          const local = name === "bold" ? "b" : "i";
          this.drop(rPr, local);
          if (now) this.upsert(rPr, local);
          break;
        }
        case "underline":
          this.drop(rPr, "u");
          if (now) this.set(this.upsert(rPr, "u"), "val", "single");
          break;
        case "strike":
          this.drop(rPr, "strike", "dstrike");
          if (now) this.upsert(rPr, "strike");
          break;
        case "subscript":
        case "superscript": {
          const vert = marks.find((m) => m.type === schema.marks.subscript || m.type === schema.marks.superscript);
          this.drop(rPr, "vertAlign");
          if (vert) this.set(this.upsert(rPr, "vertAlign"), "val", vert.type.name);
          break;
        }
        case "font": {
          const fonts = this.upsert(rPr, "rFonts");
          if (now) {
            const { font } = attrsOf<{ font: string }>(now);
            this.set(fonts, "ascii", font);
            this.set(fonts, "hAnsi", font);
          } else {
            fonts.removeAttributeNS(NS.w, "ascii");
            fonts.removeAttributeNS(NS.w, "hAnsi");
            if (!fonts.attributes.length) fonts.remove();
          }
          break;
        }
        case "size":
          this.drop(rPr, "sz", "szCs");
          if (now) {
            const half = String(Math.round(now.attrs.pt * 2));
            this.set(this.upsert(rPr, "sz"), "val", half);
            this.set(this.upsert(rPr, "szCs"), "val", half);
          }
          break;
        case "color":
          this.drop(rPr, "color");
          if (now) this.set(this.upsert(rPr, "color"), "val", attrsOf<{ color: string }>(now).color.replace("#", "").toUpperCase());
          break;
        case "highlight":
          this.drop(rPr, "highlight", "shd");
          if (now) {
            const color = String(now.attrs.color).toUpperCase();
            if (HIGHLIGHT_NAMES[color]) this.set(this.upsert(rPr, "highlight"), "val", HIGHLIGHT_NAMES[color]);
            else this.shade(this.upsert(rPr, "shd"), color);
          }
          break;
      }
    }
    return rPr.children.length || rPr.attributes.length ? rPr : null;
  }

  private image(node: PMNode): Element {
    const { xml, rId, width, height, alt } = attrsOf<ImageAttrs>(node);
    if (!xml) return this.import(this.newDrawingRun(rId ?? "", width ?? 300, height ?? 200, alt));

    const r = this.import(xml);
    const extent = r.getElementsByTagNameNS(NS.wp, "extent")[0];
    const cx = String(Math.round((width ?? 0) * EMU_PER_PX));
    const cy = String(Math.round((height ?? 0) * EMU_PER_PX));
    const sizeChanged =
      width && height && extent &&
      (Math.round(Number(extent.getAttribute("cx")) / EMU_PER_PX) !== width ||
        Math.round(Number(extent.getAttribute("cy")) / EMU_PER_PX) !== height);
    if (sizeChanged) {
      extent.setAttribute("cx", cx);
      extent.setAttribute("cy", cy);
      for (const ext of Array.from(r.getElementsByTagNameNS(NS.a, "ext")).filter((e) => e.parentElement?.localName === "xfrm")) {
        ext.setAttribute("cx", cx);
        ext.setAttribute("cy", cy);
      }
    }
    const docPr = r.getElementsByTagNameNS(NS.wp, "docPr")[0];
    if (docPr && (docPr.getAttribute("descr") ?? "") !== alt) docPr.setAttribute("descr", alt);
    const blip = r.getElementsByTagNameNS(NS.a, "blip")[0];
    if (blip && rId && blip.getAttributeNS(NS.r, "embed") !== rId) blip.setAttributeNS(NS.r, "r:embed", rId);
    return r;
  }

  private newDrawingRun(rId: string, widthPx: number, heightPx: number, alt: string): string {
    const cx = Math.round(widthPx * EMU_PER_PX);
    const cy = Math.round(heightPx * EMU_PER_PX);
    const id = this.ctx.nextDocPrId();
    const name = `Picture ${id}`;
    return (
      `<w:r xmlns:w="${NS.w}" xmlns:wp="${NS.wp}" xmlns:a="${NS.a}" xmlns:pic="${NS.pic}" xmlns:r="${NS.r}"><w:drawing>` +
      `<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
      `<wp:docPr id="${id}" name="${name}" descr="${escapeXml(alt)}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>` +
      `<pic:nvPicPr><pic:cNvPr id="${id}" name="${name}"/><pic:cNvPicPr/></pic:nvPicPr>` +
      `<pic:blipFill><a:blip r:embed="${escapeXml(rId)}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
      `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
      `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`
    );
  }
}

/**
 * Replace the content of <w:body> with `doc`. Untouched blocks are copied from
 * their original XML; edited ones are rebuilt from it, changing only what the
 * editor models. The final w:sectPr (page setup) is kept last.
 */
export function writeBody(body: Element, doc: PMNode, origins: Origins, ctx: WriteContext): void {
  let finalSectPr = body.lastElementChild?.localName === "sectPr" ? body.lastElementChild : null;
  const sectPr = doc.attrs.sectPr as string | null;
  if (sectPr && (!finalSectPr || new XMLSerializer().serializeToString(finalSectPr) !== sectPr)) {
    finalSectPr = body.ownerDocument.importNode(new DOMParser().parseFromString(sectPr, "application/xml").documentElement, true);
  }
  const writer = new Writer(body.ownerDocument, origins, ctx, doc);
  const blocks = writer.blocks(doc);
  body.replaceChildren(...blocks, ...(finalSectPr ? [finalSectPr] : []));
  if (writer.newStyles.size) ctx.ensureParagraphStyles?.(writer.newStyles);
}
