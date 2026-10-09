import { Schema, Node as PMNode, DOMOutputSpec } from "prosemirror-model";

/**
 * ProseMirror schema mirroring WordprocessingML closely enough to round-trip:
 * lists are paragraph attributes (like w:numPr), headings/quotes are paragraph
 * styles, and anything we can't edit is kept as an opaque atom holding its XML.
 *
 * `key` attrs link paragraphs/tables back to their original XML (see docx-reader).
 */

export const HEADING_LEVEL: Record<string, number> = {
  Title: 1,
  Heading1: 1,
  Heading2: 2,
  Heading3: 3,
  Heading4: 4,
  Heading5: 5,
  Heading6: 6,
};

export type Align = "left" | "center" | "right" | "justify";

// Typed views of node/mark attrs (ProseMirror types them as `any`).
export interface ParagraphAttrs {
  key: string | null;
  style: string | null;
  align: Align | null;
  numId: string | null;
  ilvl: number;
  lineSpacing: string | null;
  indent: number;
  /** Points; null = from the style. */
  spaceBefore: number | null;
  spaceAfter: number | null;
  /** First-line indent in points; negative = hanging. */
  firstLine: number | null;
}
export interface CellAttrs {
  key: string | null;
  colspan: number;
  background: string | null;
  /** Vertical merge: first cell of the merge, or a cell it covers. */
  vMerge: "restart" | "continue" | null;
}
export type TableBorders = "all" | "outer" | "none";
export interface TableAttrs {
  key: string | null;
  /** Border preset, or null for whatever the file defines. */
  borders: TableBorders | null;
  /** Column widths in twips (w:tblGrid), or null for a new table. */
  widths: number[] | null;
}
export interface ImageAttrs {
  xml: string;
  rId: string | null;
  src: string;
  width: number | null;
  height: number | null;
  alt: string;
}
export interface XmlAttrs {
  xml: string;
  text: string;
}
export interface LinkAttrs {
  href: string | null;
  rId: string | null;
  anchor: string | null;
}
export const attrsOf = <T>(n: { attrs: unknown }): T => n.attrs as T;

export const docxSchema = new Schema({
  nodes: {
    doc: { content: "block+" },

    paragraph: {
      group: "block",
      content: "inline*",
      attrs: {
        key: { default: null },
        style: { default: null },
        align: { default: null },
        numId: { default: null },
        ilvl: { default: 0 },
        lineSpacing: { default: null },
        /** Left indent in 0.5" steps. */
        indent: { default: 0 },
        spaceBefore: { default: null },
        spaceAfter: { default: null },
        firstLine: { default: null },
      },
      parseDOM: [
        { tag: "p" },
        ...[1, 2, 3, 4, 5, 6].map((n) => ({ tag: `h${n}`, attrs: { style: `Heading${n}` } })),
        { tag: "blockquote", attrs: { style: "Quote" } },
        { tag: "li" },
      ],
      toDOM(node): DOMOutputSpec {
        const a = attrsOf<ParagraphAttrs>(node);
        const level = a.style ? HEADING_LEVEL[a.style] : undefined;
        const style: string[] = [];
        if (a.align) style.push(`text-align: ${a.align}`);
        if (a.lineSpacing) style.push(`line-height: ${a.lineSpacing}`);
        if (a.indent) style.push(`margin-left: ${a.indent * 36}pt`);
        const attrs: Record<string, string> = {};
        if (style.length) attrs.style = style.join("; ");
        if (a.style) attrs["data-style"] = a.style;
        return [level ? `h${level}` : "p", attrs, 0];
      },
    },

    table: {
      group: "block",
      content: "table_row+",
      isolating: true,
      attrs: { key: { default: null }, borders: { default: null }, widths: { default: null } },
      parseDOM: [{ tag: "table" }],
      toDOM: () => ["table", ["tbody", 0]],
    },
    table_row: {
      content: "table_cell+",
      /** header: repeats at the top of each page (w:tblHeader). */
      attrs: { key: { default: null }, header: { default: false } },
      parseDOM: [{ tag: "tr" }],
      toDOM: () => ["tr", 0],
    },
    table_cell: {
      content: "block+",
      isolating: true,
      attrs: { key: { default: null }, colspan: { default: 1 }, background: { default: null }, vMerge: { default: null } },
      parseDOM: [{ tag: "td" }, { tag: "th" }],
      toDOM(node): DOMOutputSpec {
        const attrs: Record<string, string> = {};
        if (node.attrs.colspan > 1) attrs.colspan = String(node.attrs.colspan);
        if (node.attrs.background) attrs.style = `background-color: ${node.attrs.background}`;
        return ["td", attrs, 0];
      },
    },

    /** Block we can't edit (content control, horizontal rule...), written back verbatim. */
    raw_block: {
      group: "block",
      atom: true,
      selectable: true,
      attrs: { xml: { default: "" }, text: { default: "" }, kind: { default: "other" } },
      parseDOM: [{ tag: "hr", attrs: { kind: "hr" } }],
      toDOM(node): DOMOutputSpec {
        if (node.attrs.kind === "hr") return ["hr"];
        if (node.attrs.kind === "toc") return ["div", { class: "docx-toc", contenteditable: "false", title: "Table of contents (update it from the toolbar)" }, node.attrs.text];
        return ["div", { class: "docx-raw-block", contenteditable: "false", title: "Read-only content" }, node.attrs.text];
      },
    },

    text: { group: "inline" },

    hard_break: {
      group: "inline",
      inline: true,
      selectable: false,
      parseDOM: [{ tag: "br" }],
      toDOM: () => ["br"],
    },

    page_break: {
      group: "inline",
      inline: true,
      atom: true,
      selectable: true,
      parseDOM: [{ tag: "span[data-docx-page-break]" }],
      toDOM: () => [
        "span",
        { "data-docx-page-break": "1", class: "docx-page-break", contenteditable: "false" },
        ["span", { class: "docx-page-break-label" }, "Page Break"],
      ],
    },

    /** Inline picture. `xml` is the original <w:r> (empty for pictures inserted here). */
    image: {
      group: "inline",
      inline: true,
      atom: true,
      draggable: true,
      attrs: {
        xml: { default: "" },
        rId: { default: null },
        src: { default: "" },
        width: { default: null },
        height: { default: null },
        alt: { default: "" },
      },
      toDOM: (node) => {
        const { src, alt } = attrsOf<ImageAttrs>(node);
        return ["img", { src, alt, class: "docx-image-node" }];
      },
    },

    /** Run-level content we don't edit (bookmarks, fields, footnote refs, symbols...). */
    opaque: {
      group: "inline",
      inline: true,
      atom: true,
      attrs: { xml: { default: "" }, text: { default: "" } },
      toDOM: (node) => {
        const xml = node.attrs.xml as string;
        // Tracked deletions show struck through.
        const deleted = /^<w:(del|moveFrom)\b/.test(xml);
        const attrs: Record<string, string> = { class: deleted ? "docx-opaque docx-del" : "docx-opaque", contenteditable: "false" };
        // Page number fields are filled in by the page view.
        const field = /w:instr="\s*(PAGE|NUMPAGES)\b/.exec(node.attrs.xml as string)?.[1];
        if (field) attrs["data-field"] = field;
        return ["span", attrs, node.attrs.text];
      },
    },
  },

  marks: {
    /** Original w:rPr XML: keeps every run property we don't model. */
    rpr: {
      attrs: { xml: {} },
      toDOM: () => ["span", 0],
    },
    /** Tracked insertion (w:ins / w:moveTo); `xml` is the empty wrapper element with all its attributes. */
    insertion: {
      attrs: { xml: {}, author: { default: "" }, date: { default: "" } },
      inclusive: false,
      excludes: "",
      toDOM: (mark) => ["ins", { class: "docx-ins", title: `Inserted${mark.attrs.author ? ` by ${String(mark.attrs.author)}` : ""}` }, 0],
    },
    link: {
      attrs: { href: { default: null }, rId: { default: null }, anchor: { default: null } },
      inclusive: false,
      parseDOM: [{ tag: "a[href]", getAttrs: (dom) => ({ href: dom.getAttribute("href") }) }],
      toDOM: (mark) => {
        const { href, anchor } = attrsOf<LinkAttrs>(mark);
        return ["a", { href: href ?? `#${anchor ?? ""}`, class: "docx-link" }, 0];
      },
    },
    bold: {
      parseDOM: [{ tag: "strong" }, { tag: "b" }, { style: "font-weight=bold" }],
      toDOM: () => ["strong", 0],
    },
    italic: {
      parseDOM: [{ tag: "em" }, { tag: "i" }, { style: "font-style=italic" }],
      toDOM: () => ["em", 0],
    },
    underline: {
      parseDOM: [{ tag: "u" }],
      toDOM: () => ["u", 0],
    },
    strike: {
      parseDOM: [{ tag: "s" }, { tag: "strike" }, { tag: "del" }],
      toDOM: () => ["s", 0],
    },
    subscript: {
      excludes: "superscript",
      parseDOM: [{ tag: "sub" }],
      toDOM: () => ["sub", 0],
    },
    superscript: {
      excludes: "subscript",
      parseDOM: [{ tag: "sup" }],
      toDOM: () => ["sup", 0],
    },
    font: {
      attrs: { font: {} },
      toDOM: (mark) => ["span", { style: `font-family: "${mark.attrs.font}"` }, 0],
    },
    size: {
      attrs: { pt: {} },
      toDOM: (mark) => ["span", { style: `font-size: ${mark.attrs.pt}pt` }, 0],
    },
    color: {
      attrs: { color: {} },
      toDOM: (mark) => ["span", { style: `color: ${mark.attrs.color}` }, 0],
    },
    highlight: {
      attrs: { color: {} },
      toDOM: (mark) => ["mark", { style: `background-color: ${mark.attrs.color}` }, 0],
    },
  },
});

export type DocxNode = PMNode;
