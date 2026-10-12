/*
 * A rendered markdown note (HTML, as Obsidian draws it) as a Word document: headings and quotes become Word's
 * paragraph styles, lists numbered lists by depth, code blocks Google's shaded code box, tables bordered tables,
 * pictures inline images, math editable equations, and resolved wikilinks clickable links.
 */
import { Mark, Node as PMNode } from "prosemirror-model";
import { DocxDocument } from "../../infrastructure/docx/docx-document";
import { readDocx } from "./docx-reader";
import { writeBody } from "./docx-writer";
import { docxSchema as schema } from "./schema";
import { latexToOmml } from "./math";
import { escapeXml, NS } from "../../core/utils/xml";

export interface HtmlImageAttrs {
  rId: string;
  src: string;
  width: number;
  height: number;
  alt?: string;
}
export interface FromHtmlContext {
  /** numId of a bullet or numbered list. */
  list(kind: "bullet" | "ordered"): string;
  /** Image node attributes for a picture in the note, or null to leave it out. */
  image(img: HTMLImageElement): HtmlImageAttrs | null;
}

const CODE_FILL = "#F1F3F4";
const CODE_FONT = "Consolas", INLINE_CODE_FONT = "Courier New", CODE_PT = 10;
const BLOCKS = new Set(["P", "H1", "H2", "H3", "H4", "H5", "H6", "UL", "OL", "BLOCKQUOTE", "PRE", "TABLE", "HR", "DIV", "SECTION", "LI", "FIGURE", "DETAILS"]);

function equation(el: HTMLElement): PMNode | null {
  if (!el.classList.contains("math")) return null;
  const source = el.querySelector("mjx-math[data-latex]")?.getAttribute("data-latex") ?? el.querySelector('annotation[encoding="application/x-tex"]')?.textContent;
  if (!source) return null;
  const display = el.classList.contains("math-block") || !!el.querySelector('mjx-container[display="true"]');
  return schema.nodes.opaque.create({ xml: latexToOmml(source, display), text: source });
}

interface Para {
  style?: string | null;
  numId?: string | null;
  ilvl?: number | null;
  align?: "left" | "center" | "right" | null;
}

export function htmlToDoc(root: HTMLElement, ctx: FromHtmlContext): PMNode {
  const out: PMNode[] = [];
  let bookmarkId = 0;
  const control = (tag: string, content: PMNode[]) => schema.nodes.sdt.create({ xml: `<w:sdt xmlns:w="${NS.w}"><w:sdtPr><w:tag w:val="${escapeXml(tag)}"/></w:sdtPr></w:sdt>` }, content);
  // Embeds: a note's text stays (without its title and link chrome); pictures stay; any other file (a deck, sheet,
  // PDF...) is written as its link text instead of the preview drawn in the note.
  for (const embed of Array.from(root.querySelectorAll(".internal-embed"))) {
    if (embed.classList.contains("markdown-embed")) {
      embed.querySelectorAll(".markdown-embed-title, .markdown-embed-link, .embed-title").forEach((e) => e.remove());
      continue;
    }
    if (embed.classList.contains("image-embed")) continue;
    embed.replaceWith(document.createTextNode(embed.getAttribute("alt") || embed.getAttribute("src") || ""));
  }

  /** Inline content of an element (text with marks, pictures, line breaks). */
  const inline = (el: Node, marks: readonly Mark[], into: PMNode[]) => {
    for (const node of Array.from(el.childNodes)) {
      if (node.nodeType === Node.TEXT_NODE) {
        let text = (node.textContent ?? "").replace(/\s+/g, " ");
        if (into[into.length - 1]?.type.name === "hard_break") text = text.replace(/^ /, "");
        if (text) into.push(schema.text(text, marks));
        continue;
      }
      if (!(node.instanceOf(HTMLElement))) continue;
      const tag = node.tagName;
      const math = equation(node);
      if (math) into.push(math);
      else if (tag === "BR") into.push(schema.nodes.hard_break.create());
      else if (tag === "IMG") {
        const attrs = ctx.image(node as HTMLImageElement);
        if (attrs) into.push(schema.nodes.image.create({ ...attrs, alt: attrs.alt ?? node.getAttribute("alt") ?? "" }));
      } else if (tag === "INPUT") {
        // A task list's checkbox.
        if ((node as HTMLInputElement).type === "checkbox") into.push(schema.text((node as HTMLInputElement).checked || node.hasAttribute("checked") ? "☑ " : "☐ ", marks));
      } else if (node.classList.contains("frontmatter") || node.classList.contains("copy-code-button") || node.classList.contains("heading-collapse-indicator")) continue;
      else {
        let m = marks;
        const add = (mark: Mark) => (m = mark.addToSet(m));
        if (tag === "STRONG" || tag === "B") add(schema.marks.bold.create());
        else if (tag === "EM" || tag === "I") add(schema.marks.italic.create());
        else if (tag === "U") add(schema.marks.underline.create());
        else if (tag === "DEL" || tag === "S" || tag === "STRIKE") add(schema.marks.strike.create());
        else if (tag === "MARK") add(schema.marks.highlight.create({ color: "#FFFF00" }));
        else if (tag === "SUB") add(schema.marks.subscript.create());
        else if (tag === "SUP") add(schema.marks.superscript.create());
        else if (tag === "CODE") add(schema.marks.font.create({ font: INLINE_CODE_FONT }));
        else if (tag === "A" && /^(https?|mailto|obsidian):/i.test(node.getAttribute("href") ?? "")) add(schema.marks.link.create({ href: node.getAttribute("href") }));
        inline(node, m, into);
      }
    }
  };

  const paragraph = (el: Node, attrs: Para, prefix: PMNode[] = []) => {
    const content = [...prefix];
    inline(el, [], content);
    // Trim the whitespace HTML leaves at the paragraph's ends.
    while (content[0]?.isText && !content[0].text!.trim()) content.shift();
    while (content[content.length - 1]?.isText && !content[content.length - 1].text!.trim()) content.pop();
    const first = content[0], last = content[content.length - 1];
    if (first?.isText) content[0] = schema.text(first.text!.replace(/^\s+/, ""), first.marks);
    if (last?.isText) content[content.length - 1] = schema.text(content[content.length - 1].text!.replace(/\s+$/, ""), last.marks);
    const cleaned = content;
    if (el.instanceOf(HTMLElement) && el.querySelector('input[type="checkbox"]')) {
      const id = bookmarkId++;
      cleaned.unshift(schema.nodes.opaque.create({ xml: `<w:bookmarkStart xmlns:w="${NS.w}" w:id="${id}" w:name="_DocWorksTask${id}"/>` }), schema.nodes.opaque.create({ xml: `<w:bookmarkEnd xmlns:w="${NS.w}" w:id="${id}"/>` }));
    }
    return schema.nodes.paragraph.create({ style: attrs.style ?? null, numId: attrs.numId ?? null, ilvl: attrs.ilvl ?? null, align: attrs.align ?? null }, cleaned);
  };

  /** Blocks of `el` appended to `into`, with paragraph attributes from the surrounding list or quote. */
  const blocks = (el: HTMLElement, into: PMNode[], attrs: Para, depth: number) => {
    // Loose text and inline elements between blocks gather into paragraphs.
    let loose: Node[] = [];
    const flush = () => {
      if (!loose.some((n) => (n.textContent ?? "").trim() || (n.instanceOf(HTMLElement) && (n.querySelector("img, input") || ["IMG", "INPUT"].includes(n.tagName))))) return void (loose = []);
      const holder = createSpan();
      holder.append(...loose.map((n) => n.cloneNode(true)));
      into.push(paragraph(holder, attrs));
      loose = [];
    };
    for (const node of Array.from(el.childNodes)) {
      if (!(node.instanceOf(HTMLElement)) || !BLOCKS.has(node.tagName)) {
        loose.push(node);
        continue;
      }
      flush();
      block(node, into, attrs, depth);
    }
    flush();
  };

  const block = (el: HTMLElement, into: PMNode[], attrs: Para, depth: number) => {
    const tag = el.tagName;
    const math = equation(el);
    if (math) into.push(schema.nodes.paragraph.create({ align: "center" }, math));
    else if (/^H[1-6]$/.test(tag)) into.push(paragraph(el, { ...attrs, style: `Heading${tag[1]}` }));
    else if (tag === "P") into.push(paragraph(el, attrs));
    else if (tag === "BLOCKQUOTE") blocks(el, into, { ...attrs, style: "Quote" }, depth);
    else if (tag === "UL" || tag === "OL") {
      const numId = ctx.list(tag === "UL" ? "bullet" : "ordered");
      for (const li of Array.from(el.children)) if (li.tagName === "LI") blocks(li as HTMLElement, into, { style: attrs.style, numId, ilvl: depth }, depth + 1);
    } else if (tag === "PRE") {
      const marks = [schema.marks.font.create({ font: CODE_FONT }), schema.marks.size.create({ pt: CODE_PT }), schema.marks.noProof.create()];
      const lines = (el.querySelector("code") ?? el).textContent.replace(/\n$/, "").split("\n");
      const paras = lines.map((line) => schema.nodes.paragraph.create({ spaceBefore: 0, spaceAfter: 0 }, line ? schema.text(line, marks) : []));
      const table = schema.nodes.table.create({ borders: "outer" }, schema.nodes.table_row.create(null, schema.nodes.table_cell.create({ background: CODE_FILL }, paras)));
      const language = /(?:^|\s)language-([^\s]+)/.exec(`${el.querySelector("code")?.className ?? ""} ${el.className}`)?.[1];
      into.push(language ? control(`doc-works:code:${language}`, [table]) : table);
    } else if (tag === "TABLE") {
      const rows = Array.from(el.querySelectorAll("tr")).map((tr) => {
        const cells = Array.from(tr.children).filter((c) => c.tagName === "TD" || c.tagName === "TH").map((c) => {
          const content: PMNode[] = [];
          // A markdown column's alignment (---: / :---:) arrives as the cell's text-align.
          const raw = ((c as HTMLElement).style.textAlign || c.getAttribute("align") || "").toLowerCase();
          const align = raw === "right" || raw === "center" ? raw : null;
          blocks(c as HTMLElement, content, { align }, 0);
          return schema.nodes.table_cell.create(null, content.length ? content : [schema.nodes.paragraph.create()]);
        });
        return schema.nodes.table_row.create({ header: tr.parentElement?.tagName === "THEAD" }, cells);
      });
      // Every row as wide as the widest (markdown tables can be ragged).
      const width = Math.max(...rows.map((r) => r.childCount));
      const even = rows.map((r) => (r.childCount === width ? r : r.type.create(r.attrs, [...Array.from({ length: r.childCount }, (_, i) => r.child(i)), ...Array.from({ length: width - r.childCount }, () => schema.nodes.table_cell.create(null, schema.nodes.paragraph.create()))])));
      if (even.length) into.push(schema.nodes.table.create({ borders: "all" }, even));
    } else if (tag === "HR") into.push(schema.nodes.raw_block.create({ kind: "hr" }));
    else if (el.classList.contains("callout")) {
      // A callout: its title in bold, its body as a quote.
      const title = el.querySelector(".callout-title-inner");
      const contentBlocks: PMNode[] = [];
      if (title?.textContent?.trim()) contentBlocks.push(paragraph(title, { style: "Quote" }));
      else contentBlocks.push(schema.nodes.paragraph.create({ style: "Quote" }));
      const content = el.querySelector(".callout-content");
      if (content) blocks(content as HTMLElement, contentBlocks, { ...attrs, style: "Quote" }, depth);
      const type = el.getAttribute("data-callout") || "note";
      const fold = el.getAttribute("data-callout-fold") || "";
      into.push(control(`doc-works:callout:${type}:${fold}`, contentBlocks));
    } else if (el.classList.contains("frontmatter-container") || el.classList.contains("mod-header") || el.classList.contains("mod-footer")) return;
    else blocks(el, into, attrs, depth);
  };

  blocks(root, out, {}, 0);
  return schema.nodes.doc.create(null, out.length ? out : [schema.nodes.paragraph.create()]);
}

/** Picture bytes and natural size for a note's image, or null to leave it out. */
export type ImageSource = (img: HTMLImageElement) => { bytes: Uint8Array; ext: string; width: number; height: number } | null;

export interface NoteExportOptions {
  pageSize: "letter" | "legal" | "a4" | "a5";
  landscape: boolean;
  margin: "default" | "narrow" | "wide" | "none";
  /** Text and picture size as a percentage of the original (10–100). */
  scale: number;
  title: string;
}

export const DEFAULT_NOTE_EXPORT_OPTIONS: NoteExportOptions = { pageSize: "letter", landscape: false, margin: "default", scale: 100, title: "" };
const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
// Word page dimensions in twips, including ISO sizes rounded to the nearest twip.
const PAPER = { letter: [12240, 15840], legal: [12240, 20160], a4: [11906, 16838], a5: [8391, 11906] } as const;

function exportPage(docx: DocxDocument, options: NoteExportOptions): number {
  const section = Array.from(docx.body.children).find((e) => e.localName === "sectPr")!;
  const size = section.getElementsByTagNameNS(NS_W, "pgSz")[0];
  const margins = section.getElementsByTagNameNS(NS_W, "pgMar")[0];
  const [short, long] = PAPER[options.pageSize];
  const [width, height] = options.landscape ? [long, short] : [short, long];
  size.setAttributeNS(NS_W, "w:w", String(width));
  size.setAttributeNS(NS_W, "w:h", String(height));
  size.setAttributeNS(NS_W, "w:orient", options.landscape ? "landscape" : "portrait");
  const margin = { default: 1440, narrow: 720, wide: 2880, none: 0 }[options.margin];
  for (const side of ["top", "right", "bottom", "left"]) margins.setAttributeNS(NS_W, `w:${side}`, String(margin));
  for (const side of ["header", "footer"]) margins.setAttributeNS(NS_W, `w:${side}`, String(Math.min(720, margin / 2)));
  return (width - margin * 2) / 15;
}

/** Scale font sizes and paragraph spacing, including inherited heading and default styles. */
function downscaleText(root: Element, scale: number): void {
  const shrink = (el: Element, attr: string, minimum = 0) => {
    const value = el.getAttributeNS(NS_W, attr);
    if (value !== null) el.setAttributeNS(NS_W, `w:${attr}`, String(Math.max(minimum, Math.round(Number(value) * scale))));
  };
  for (const name of ["sz", "szCs"]) for (const el of Array.from(root.getElementsByTagNameNS(NS_W, name))) shrink(el, "val", 1);
  for (const el of Array.from(root.getElementsByTagNameNS(NS_W, "spacing"))) {
    shrink(el, "before");
    shrink(el, "after");
    if (["exact", "atLeast"].includes(el.getAttributeNS(NS_W, "lineRule") ?? "")) shrink(el, "line");
  }
}

/** A Word document (from `blank`, an empty .docx) holding the rendered note `root`. */
export async function noteToDocx(blank: Uint8Array, root: HTMLElement, images: ImageSource, settings: Partial<NoteExportOptions> = {}): Promise<Uint8Array> {
  const options = { ...DEFAULT_NOTE_EXPORT_OPTIONS, ...settings };
  const scale = Number.isFinite(options.scale) ? Math.max(10, Math.min(100, options.scale)) / 100 : 1;
  const docx = await DocxDocument.load(blank);
  if (root.dataset.exportFrontmatter) docx.setDocumentVariable("DocWorksFrontmatter", root.dataset.exportFrontmatter);
  const maxWidth = exportPage(docx, options);
  let orderedSeen = 0;
  const read = readDocx(docx.body, { imageSrc: () => null, hyperlinkHref: (id) => docx.hyperlinkHref(id) });
  const content = htmlToDoc(root, {
    // Each numbered list counts from 1 (the shared definition would continue the last list's count).
    list: (kind) => {
      const id = docx.ensureList(kind);
      if (kind !== "ordered") return id;
      return orderedSeen++ ? docx.restartList(id) : id;
    },
    image: (img) => {
      const data = images(img);
      if (!data) return null;
      const width = Math.min(data.width || maxWidth, maxWidth) * scale;
      return { rId: docx.addImage(data.bytes, data.ext), src: "", width, height: Math.round((width / (data.width || width)) * (data.height || width * 0.75)) };
    },
  });
  const title = options.title ? schema.nodes.paragraph.create({ style: "Heading1" }, schema.text(options.title)) : null;
  const body = title ? content.content.addToStart(title) : content.content;
  writeBody(docx.body, schema.nodes.doc.create(read.doc.attrs, body), read.origins, docx);
  if (scale !== 1) {
    downscaleText(docx.body, scale);
    const styles = docx.stylesXml();
    if (styles) downscaleText(styles.documentElement, scale);
  }
  return docx.save();
}
