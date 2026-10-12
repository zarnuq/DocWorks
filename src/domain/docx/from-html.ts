/*
 * A rendered markdown note (HTML, as Obsidian draws it) as a Word document: headings and quotes become Word's
 * paragraph styles, lists numbered lists by depth, code blocks Google's shaded code box, tables bordered tables,
 * pictures inline images. Wikilinks keep their text; external links stay links.
 */
import { Mark, Node as PMNode } from "prosemirror-model";
import { DocxDocument } from "../../infrastructure/docx/docx-document";
import { readDocx } from "./docx-reader";
import { writeBody } from "./docx-writer";
import { docxSchema as schema } from "./schema";

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

interface Para {
  style?: string | null;
  numId?: string | null;
  ilvl?: number | null;
  align?: "left" | "center" | "right" | null;
}

export function htmlToDoc(root: HTMLElement, ctx: FromHtmlContext): PMNode {
  const out: PMNode[] = [];

  /** Inline content of an element (text with marks, pictures, line breaks). */
  const inline = (el: Node, marks: readonly Mark[], into: PMNode[]) => {
    for (const node of Array.from(el.childNodes)) {
      if (node.nodeType === Node.TEXT_NODE) {
        const text = (node.textContent ?? "").replace(/\s+/g, " ");
        if (text) into.push(schema.text(text, marks));
        continue;
      }
      if (!(node.instanceOf(HTMLElement))) continue;
      const tag = node.tagName;
      if (tag === "BR") into.push(schema.nodes.hard_break.create());
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
        else if (tag === "A" && !node.classList.contains("internal-link") && /^(https?|mailto):/i.test(node.getAttribute("href") ?? "")) add(schema.marks.link.create({ href: node.getAttribute("href") }));
        inline(node, m, into);
      }
    }
  };

  const paragraph = (el: Node, attrs: Para, prefix: PMNode[] = []) => {
    const content = [...prefix];
    inline(el, [], content);
    // Trim the whitespace HTML leaves at the paragraph's ends.
    const first = content[0], last = content[content.length - 1];
    if (first?.isText) content[0] = schema.text(first.text!.replace(/^\s+/, "") || " ", first.marks);
    if (last?.isText && content.length > 0) content[content.length - 1] = schema.text(last.text!.replace(/\s+$/, "") || " ", last.marks);
    const cleaned = content.filter((n) => !(n.isText && n.text === " " && content.length === 1));
    return schema.nodes.paragraph.create({ style: attrs.style ?? null, numId: attrs.numId ?? null, ilvl: attrs.ilvl ?? null, align: attrs.align ?? null }, cleaned);
  };

  /** Blocks of `el` appended to `into`, with paragraph attributes from the surrounding list or quote. */
  const blocks = (el: HTMLElement, into: PMNode[], attrs: Para, depth: number) => {
    // Loose text and inline elements between blocks gather into paragraphs.
    let loose: Node[] = [];
    const flush = () => {
      if (!loose.some((n) => (n.textContent ?? "").trim() || (n.instanceOf(HTMLElement) && n.querySelector("img, input")) || (n as HTMLElement).tagName === "IMG")) return void (loose = []);
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
    if (/^H[1-6]$/.test(tag)) into.push(paragraph(el, { ...attrs, style: `Heading${tag[1]}` }));
    else if (tag === "P") into.push(paragraph(el, attrs));
    else if (tag === "BLOCKQUOTE") blocks(el, into, { ...attrs, style: "Quote" }, depth);
    else if (tag === "UL" || tag === "OL") {
      const numId = ctx.list(tag === "UL" ? "bullet" : "ordered");
      for (const li of Array.from(el.children)) if (li.tagName === "LI") blocks(li as HTMLElement, into, { style: attrs.style, numId, ilvl: depth }, depth + 1);
    } else if (tag === "PRE") {
      const marks = [schema.marks.font.create({ font: CODE_FONT }), schema.marks.size.create({ pt: CODE_PT }), schema.marks.noProof.create()];
      const lines = (el.querySelector("code") ?? el).textContent.replace(/\n$/, "").split("\n");
      const paras = lines.map((line) => schema.nodes.paragraph.create({ spaceBefore: 0, spaceAfter: 0 }, line ? schema.text(line, marks) : []));
      into.push(schema.nodes.table.create({ borders: "outer" }, schema.nodes.table_row.create(null, schema.nodes.table_cell.create({ background: CODE_FILL }, paras))));
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
    } else if (tag === "HR") into.push(schema.nodes.paragraph.create());
    else if (el.classList.contains("callout")) {
      // A callout: its title in bold, its body as a quote.
      const title = el.querySelector(".callout-title-inner");
      if (title?.textContent?.trim()) into.push(schema.nodes.paragraph.create({ style: "Quote" }, schema.text(title.textContent.trim(), [schema.marks.bold.create()])));
      const content = el.querySelector(".callout-content");
      if (content) blocks(content as HTMLElement, into, { ...attrs, style: "Quote" }, depth);
    } else if (el.classList.contains("frontmatter-container") || el.classList.contains("mod-header") || el.classList.contains("mod-footer")) return;
    else blocks(el, into, attrs, depth);
  };

  blocks(root, out, {}, 0);
  return schema.nodes.doc.create(null, out.length ? out : [schema.nodes.paragraph.create()]);
}

/** Picture bytes and natural size for a note's image, or null to leave it out. */
export type ImageSource = (img: HTMLImageElement) => { bytes: Uint8Array; ext: string; width: number; height: number } | null;

/** A Word document (from `blank`, an empty .docx) holding the rendered note `root`. */
export async function noteToDocx(blank: Uint8Array, root: HTMLElement, images: ImageSource, maxWidth = 624): Promise<Uint8Array> {
  const docx = await DocxDocument.load(blank);
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
      const width = Math.min(data.width || maxWidth, maxWidth);
      return { rId: docx.addImage(data.bytes, data.ext), src: "", width, height: Math.round((width / (data.width || width)) * (data.height || width * 0.75)) };
    },
  });
  // The blank document's section (page size and margins) is kept.
  writeBody(docx.body, schema.nodes.doc.create(read.doc.attrs, content.content), read.origins, docx);
  return docx.save();
}
