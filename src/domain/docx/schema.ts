import { Schema, Node as PMNode, DOMOutputSpec } from "prosemirror-model";
import { fontStack } from "../../core/utils/fonts";
import { Placement, placementCss, textBoxOf, Wrap } from "./drawing";
import { mathSpec } from "./math";

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
export interface TabStop { pt: number; align: string; leader: string | null; }

// Typed views of node/mark attrs (ProseMirror types them as `any`).
export interface ParagraphAttrs {
  key: string | null;
  style: string | null;
  align: Align | null;
  numId: string | null;
  ilvl: number;
  lineSpacing: string | null;
  indent: number;
  rightIndent: number | null;
  tabs: TabStop[] | null;
  /** Points; null = from the style. */
  spaceBefore: number | null;
  spaceAfter: number | null;
  /** First-line indent in points; negative = hanging. */
  firstLine: number | null;
  /** Leader of the paragraph's right-aligned tab stop ("none", "dot"...), null if it has none (display only). */
  tabLeader: string | null;
  /** Size (pt) and font of the paragraph mark, which Word sizes the lines by (display only). */
  markSize: number | null;
  markFont: string | null;
  /** The w:sectPr XML when the paragraph ends a section. */
  sectPr: string | null;
  /** Keep with next / keep lines together; null = from the style. */
  keepNext: boolean | null;
  keepLines: boolean | null;
  /** Contextual spacing: no space before / after next to a paragraph of the same style; null = from the style. */
  contextual: boolean | null;
  /** Borders and shading as CSS (display only). */
  frame: string | null;
}
export interface CellAttrs {
  key: string | null;
  colspan: number;
  background: string | null;
  /** Vertical merge: first cell of the merge, or a cell it covers. */
  vMerge: "restart" | "continue" | null;
  /** The cell's own border overrides, e.g. "t:none b:line" (display only; the XML keeps the details). */
  edges: string | null;
  /** Padding "top right bottom left" in pt, "-" = default (display only). */
  margins: string | null;
}
export type TableBorders = "all" | "outer" | "none";
export interface TableAttrs {
  key: string | null;
  /** Border preset, or null for whatever the file defines. */
  borders: TableBorders | null;
  /** Column widths in twips (w:tblGrid), or null for a new table. */
  widths: number[] | null;
  /** Table style id (w:tblStyle), for display only: its paragraph settings apply to the text in the table. */
  style: string | null;
}
export interface ImageAttrs extends Placement {
  xml: string;
  rId: string | null;
  src: string;
  width: number | null;
  height: number | null;
  alt: string;
  /** Where the picture sits (see drawing.ts). */
  wrap: Wrap;
  x: number;
  y: number;
  align: Placement["align"];
  /** Share of the picture cut off each side (a:srcRect), or null. */
  crop: { l: number; t: number; r: number; b: number } | null;
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
/** A text box: its paragraphs in a box of its size, placed like a picture. */
function textBoxSpec(box: NonNullable<ReturnType<typeof textBoxOf>>, attrs: Record<string, string>): DOMOutputSpec {
  const { cls, style } = placementCss(box.placement, box.width);
  const css = [`width: ${box.width}px`, `height: ${box.height}px`, box.fill ? `background-color: ${box.fill}` : "", box.line ? `border: 1px solid ${box.line}` : "", style];
  return ["span", { ...attrs, class: `${attrs.class} docx-textbox ${cls}`.trim(), style: css.filter(Boolean).join("; ") }, ...box.paragraphs.map((t) => ["span", { class: "docx-textbox-p" }, t] as DOMOutputSpec)];
}

export const attrsOf = <T>(n: { attrs: unknown }): T => n.attrs as T;

export const docxSchema = new Schema({
  nodes: {
    /** sectPr: the body's own w:sectPr XML (the last section). */
    doc: { content: "block+", attrs: { sectPr: { default: null } } },

    /** A section in two or more columns (display; the columns themselves are saved in the section's w:sectPr). */
    columns: {
      group: "block",
      content: "block+",
      attrs: { count: { default: 2 }, space: { default: 720 }, sep: { default: false } },
      toDOM(node): DOMOutputSpec {
        const { count, space, sep } = node.attrs as { count: number; space: number; sep: boolean };
        // Multi-column properties go inline: the plugin directory's CSS lint flags them in styles.css.
        // The page view measures with --docx-col-count: 1.
        const gap = `${space / 1440}in`;
        const style = `--docx-cols: ${count}; --docx-col-gap: ${gap}; column-count: var(--docx-col-count, ${count}); column-gap: ${gap};${sep ? " column-rule: 1px solid currentColor;" : ""}`;
        return ["div", { class: sep ? "docx-columns docx-columns-sep" : "docx-columns", style }, 0];
      },
    },

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
        rightIndent: { default: null },
        tabs: { default: null },
        spaceBefore: { default: null },
        spaceAfter: { default: null },
        firstLine: { default: null },
        tabLeader: { default: null },
        markSize: { default: null },
        markFont: { default: null },
        sectPr: { default: null },
        keepNext: { default: null },
        keepLines: { default: null },
        contextual: { default: null },
        frame: { default: null },
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
        if (a.rightIndent !== null) style.push(`margin-right: ${a.rightIndent}pt`);
        if (a.frame) style.push(a.frame);
        const attrs: Record<string, string> = {};
        if (style.length) attrs.style = style.join("; ");
        if (a.style) attrs["data-style"] = a.style;
        return [level ? `h${level}` : "p", attrs, 0];
      },
    },

    /** Content control (w:sdt) around blocks. `xml` is the empty w:sdt with its properties, for controls the editor made. */
    sdt: {
      group: "block",
      content: "block+",
      attrs: { key: { default: null }, kind: { default: "other" }, xml: { default: "" } },
      toDOM: (node) => ["div", { class: node.attrs.kind === "toc" ? "docx-sdt docx-toc" : "docx-sdt" }, 0],
    },

    table: {
      group: "block",
      content: "table_row+",
      isolating: true,
      attrs: { key: { default: null }, borders: { default: null }, widths: { default: null }, style: { default: null } },
      parseDOM: [{ tag: "table" }],
      toDOM: (node) => ["table", node.attrs.borders ? { class: `docx-borders-${node.attrs.borders as string}` } : {}, ["tbody", 0]],
    },
    table_row: {
      content: "table_cell+",
      /** header: repeats at the top of each page (w:tblHeader); cantSplit: moves to the next page whole. */
      attrs: { key: { default: null }, header: { default: false }, cantSplit: { default: false } },
      parseDOM: [{ tag: "tr" }],
      toDOM: node => ["tr", node.attrs.cantSplit ? { style: "break-inside: avoid" } : {}, 0],
    },
    table_cell: {
      content: "block+",
      isolating: true,
      attrs: { key: { default: null }, colspan: { default: 1 }, background: { default: null }, vMerge: { default: null }, edges: { default: null }, margins: { default: null } },
      parseDOM: [{ tag: "td" }, { tag: "th" }],
      toDOM(node): DOMOutputSpec {
        const attrs: Record<string, string> = {};
        const edges = node.attrs.edges as string | null;
        if (edges) attrs.class = edges.split(" ").map((e) => `docx-e-${e.replace(":", "-")}`).join(" ");
        if (node.attrs.colspan > 1) attrs.colspan = String(node.attrs.colspan);
        const style: string[] = [];
        if (node.attrs.background) style.push(`background-color: ${node.attrs.background}`);
        const margins = (node.attrs.margins as string | null)?.split(" ") ?? [];
        ["top", "right", "bottom", "left"].forEach((side, i) => margins[i] && margins[i] !== "-" && style.push(`padding-${side}: ${margins[i]}`));
        if (style.length) attrs.style = style.join("; ");
        return ["td", attrs, 0];
      },
    },

    /** Block we can't edit (horizontal rule, unknown elements...), written back verbatim. */
    raw_block: {
      group: "block",
      atom: true,
      selectable: true,
      attrs: { xml: { default: "" }, text: { default: "" }, kind: { default: "other" } },
      parseDOM: [{ tag: "hr", attrs: { kind: "hr" } }],
      toDOM(node): DOMOutputSpec {
        if (node.attrs.kind === "hr") return ["hr"];
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
        wrap: { default: "inline" },
        hRelative: { default: null },
        vRelative: { default: null },
        distances: { default: null },
        x: { default: 0 },
        y: { default: 0 },
        align: { default: null },
        crop: { default: null },
      },
      toDOM: (node) => {
        const { src, alt, width, height, wrap, x, y, align, crop } = attrsOf<ImageAttrs>(node);
        const { cls, style } = placementCss({ ...attrsOf<ImageAttrs>(node), wrap, x, y, align }, width ?? 0);
        const size = `${width ? `width: ${width}px; ` : ""}${height ? `height: ${height}px; ` : ""}`;
        if (!crop || !width || !height) return ["img", { src, alt, class: `docx-image-node ${cls}`.trim(), style: size + style }];
        // Cropped: a clipped box the picture's size, holding the whole image scaled and shifted inside it.
        const c = cropBox(width, height, crop);
        return ["span", { class: `docx-image-node docx-cropped ${cls}`.trim(), style: size + style },
          ["img", { src, alt, style: `width: ${c.width}px; height: ${c.height}px; margin-left: ${c.left}px; margin-top: ${c.top}px` }]];
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
        if (/<w:br\b[^>]*w:type="column"/.test(xml)) {
          attrs.class += " docx-column-break";
          attrs.style = "break-after: column";
        }
        if (/^<m:oMath/.test(xml)) return ["span", attrs, mathSpec(xml)];
        const box = textBoxOf(xml);
        if (box) return textBoxSpec(box, attrs);
        return ["span", attrs, node.attrs.text];
      },
    },
  },

  marks: {
    /** Cached result of a PAGE / NUMPAGES field; the page view shows the real number. */
    field: {
      attrs: { name: {} },
      inclusive: false,
      toDOM: (mark) => ["span", { "data-field": mark.attrs.name as string }, 0],
    },
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
    font: {
      attrs: { font: {} },
      toDOM: (mark) => ["span", { style: `font-family: ${fontStack(mark.attrs.font as string)}` }, 0],
    },
    size: {
      attrs: { pt: {} },
      toDOM: (mark) => ["span", { style: `font-size: ${mark.attrs.pt}pt` }, 0],
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

/** Size and offset of the whole image inside a cropped picture's `width` x `height` box (px). */
export function cropBox(width: number, height: number, crop: NonNullable<ImageAttrs["crop"]>): { width: number; height: number; left: number; top: number } {
  const w = width / Math.max(0.05, 1 - crop.l - crop.r);
  const h = height / Math.max(0.05, 1 - crop.t - crop.b);
  const round = (v: number) => Math.round(v * 100) / 100;
  return { width: round(w), height: round(h), left: round(-crop.l * w), top: round(-crop.t * h) };
}
