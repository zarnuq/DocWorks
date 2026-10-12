import { Node as PMNode } from "prosemirror-model";
import { NS } from "../../core/utils/xml";
import { DocxDocument } from "../../infrastructure/docx/docx-document";
import { parseXml } from "../../infrastructure/ooxml/ooxml-package";
import { readDocx, w, wVal } from "./docx-reader";
import { numberingLevel } from "./docx-styles";
import { ommlToLatex } from "./math";
import { attrsOf, CellAttrs, HEADING_LEVEL, ImageAttrs, LinkAttrs, ParagraphAttrs, XmlAttrs } from "./schema";

export interface MarkdownExport {
  markdown: string;
  images: { name: string; bytes: Uint8Array }[];
  warnings: string[];
}

const escape = (text: string) => text.replace(/[\\`*_[\]<>!#|~$]/g, "\\$&");
const destination = (url: string) => `<${url.replace(/[\s<>\\]/g, c => encodeURIComponent(c))}>`;
const mono = (font: string) => /^(consolas|courier( new)?|monospace|menlo|monaco|liberation mono)$/i.test(font);
const fence = (text: string, minimum: number) => "`".repeat(Math.max(minimum, ...Array.from(text.matchAll(/`+/g), m => m[0].length + 1)));
const codeSpan = (text: string) => {
  const ticks = fence(text, 1);
  const pad = text.startsWith("`") || text.endsWith("`") || (/^ .* $/.test(text) && text.trim()) ? " " : "";
  return `${ticks}${pad}${text}${pad}${ticks}`;
};
function tablePipes(text: string): string {
  return text.replace(/\|/g, (pipe: string, offset: number) => {
    let slashes = 0;
    for (let i = offset - 1; i >= 0 && text[i] === "\\"; i--) slashes++;
    return slashes % 2 ? pipe : "\\|";
  });
}

interface InlineMark { key: string; wrap: (text: string) => string; }
interface Token { text: string; marks: InlineMark[]; }

/** Keep contiguous marks together even when Word splits a bold/link span into several runs. */
function marked(tokens: Token[]): string {
  let out = "";
  for (let i = 0; i < tokens.length;) {
    const mark = tokens[i].marks[0];
    if (!mark) { out += tokens[i++].text; continue; }
    const group: Token[] = [];
    do { group.push({ ...tokens[i], marks: tokens[i].marks.slice(1) }); i++; } while (i < tokens.length && tokens[i].marks[0]?.key === mark.key);
    out += mark.wrap(marked(group));
  }
  return out;
}

/** Export the body as editable Markdown; page layout, comments and drawing geometry are not Markdown. */
export async function docxToMarkdown(docx: DocxDocument, imageFolder: string, current?: PMNode): Promise<MarkdownExport> {
  const doc = current ?? readDocx(docx.body, { imageSrc: () => null, hyperlinkHref: id => docx.hyperlinkHref(id) }).doc;
  const images: MarkdownExport["images"] = [];
  const warnings = new Set<string>();
  const imageLinks = new Map<string, string>();
  const mediaLinks = new Map<string, string>();
  const relationships = docx.rels(docx.mainPath);
  // Only referenced pictures are extracted. Repeated relationships to the same media share an attachment.
  const ids = new Set<string>();
  doc.descendants(n => { if (n.type.name === "image") { const { rId } = attrsOf<ImageAttrs>(n); if (rId) ids.add(rId); } });
  for (const id of ids) {
    const rel = relationships.find(r => r.id === id && r.type === "image");
    if (!rel) continue;
    if (rel.external) { imageLinks.set(id, rel.target); continue; }
    let link = mediaLinks.get(rel.target);
    if (!link) {
      const ext = /\.([a-z0-9]+)$/i.exec(rel.target)?.[1]?.toLowerCase() ?? "bin";
      const name = `image-${images.length + 1}.${ext}`;
      try {
        images.push({ name, bytes: await docx.media(rel.target) });
        link = `${imageFolder}/${name}`.split("/").map(encodeURIComponent).join("/");
        mediaLinks.set(rel.target, link);
      } catch { warnings.add("Some pictures could not be extracted."); continue; }
    }
    imageLinks.set(id, link);
  }

  const styles = new Map(Array.from(docx.stylesXml()?.getElementsByTagNameNS(NS.w, "style") ?? []).map(st => [wVal(st, "styleId"), st]));
  const chain = (id: string | null): Element[] => {
    const out: Element[] = [];
    while (id && out.length < 10) {
      const st = styles.get(id);
      if (!st || out.includes(st)) break;
      out.push(st); id = wVal(w(st, "basedOn"));
    }
    return out;
  };
  const numbering = docx.numberingXml();
  const counters = new Map<string, number[]>();
  const list = (a: ParagraphAttrs): { prefix: string; left: number } | null => {
    const inherited = chain(a.style).map(st => w(w(st, "pPr"), "numPr")).find(Boolean);
    const id = a.numId ?? wVal(w(inherited, "numId"));
    if (!id || id === "0") return null;
    const level = Math.min(8, Math.max(0, a.numId || a.ilvl ? a.ilvl : Number(wVal(w(inherited, "ilvl")) ?? 0)));
    const num = Array.from(numbering?.getElementsByTagNameNS(NS.w, "num") ?? []).find(n => wVal(n, "numId") === id);
    const override = Array.from(num?.children ?? []).find(n => n.localName === "lvlOverride" && wVal(n, "ilvl") === String(level));
    const lvl = w(override, "lvl") ?? numberingLevel(numbering, id, level);
    const ordered = lvl && !["bullet", "none"].includes(wVal(w(lvl, "numFmt")) ?? "bullet");
    const counts = counters.get(id) ?? [];
    counters.set(id, counts);
    const start = Number(wVal(w(override, "startOverride")) ?? wVal(w(lvl, "start")) ?? 1);
    counts[level] = (counts[level] ?? start - 1) + 1;
    counts.length = level + 1;
    const styleIndent = chain(a.style).map(st => w(w(st, "pPr"), "ind")).find(Boolean);
    const levelIndent = w(w(lvl, "pPr"), "ind");
    const left = a.indent ? a.indent * 720 : Number(wVal(styleIndent, "left") ?? wVal(styleIndent, "start") ?? wVal(levelIndent, "left") ?? wVal(levelIndent, "start") ?? 720 * (level + 1));
    return { prefix: ordered ? `${counts[level]}. ` : "- ", left };
  };

  const usedNotes = new Map<string, string>();
  const opaque = (n: PMNode): string => {
    const { xml, text } = attrsOf<XmlAttrs>(n);
    const el = parseXml(xml)?.documentElement;
    if (!el) return escape(text);
    if (el.namespaceURI === NS.m) {
      const latex = ommlToLatex(xml);
      return el.localName === "oMathPara" ? `\n\n$$\n${latex}\n$$\n\n` : `$${latex}$`;
    }
    if (el.localName === "del" || el.localName === "moveFrom") return "";
    for (const [tag, part, label] of [["footnoteReference", "footnotes", "fn"], ["endnoteReference", "endnotes", "en"]]) {
      const ref = el.getElementsByTagNameNS(NS.w, tag)[0];
      if (!ref) continue;
      const id = wVal(ref, "id") ?? "";
      const path = docx.relOfType(docx.mainPath, part);
      const root = path ? docx.xml(path) : null;
      const note = Array.from(root?.getElementsByTagNameNS(NS.w, part === "footnotes" ? "footnote" : "endnote") ?? []).find(e => wVal(e, "id") === id);
      if (note) {
        const key = `${label}-${id}`;
        const paragraphs = Array.from(note.getElementsByTagNameNS(NS.w, "p")).map(p => Array.from(p.getElementsByTagNameNS(NS.w, "t")).map(t => t.textContent).join(""));
        usedNotes.set(key, paragraphs.map(escape).join("\n").trimStart());
        return `[^${key}]`;
      }
    }
    const paragraphs = Array.from(el.getElementsByTagNameNS(NS.w, "txbxContent")).flatMap(box => Array.from(box.getElementsByTagNameNS(NS.w, "p")).map(p => Array.from(p.getElementsByTagNameNS(NS.w, "t")).map(t => t.textContent).join("")));
    if (paragraphs.length) return paragraphs.map(escape).join("  \n");
    if (el.getElementsByTagNameNS(NS.w, "drawing").length || el.getElementsByTagNameNS(NS.w, "pict").length) warnings.add("Some drawings could not be converted.");
    return escape(text);
  };

  const inline = (p: PMNode): string => {
    const tokens: Token[] = [];
    p.forEach(n => {
      let text = "";
      const font = n.marks.find(m => m.type.name === "font");
      const code = n.isText && font && mono(attrsOf<{ font: string }>(font).font);
      if (n.isText) text = code ? n.text! : escape(n.text!);
      else if (n.type.name === "hard_break") text = "  \n";
      else if (n.type.name === "page_break") text = "\n\n";
      else if (n.type.name === "image") {
        const a = attrsOf<ImageAttrs>(n), link = a.rId && imageLinks.get(a.rId);
        if (link) text = `![${escape(a.alt)}](${destination(link)})`;
        else { text = escape(a.alt); warnings.add("Some pictures could not be extracted."); }
      } else text = opaque(n);
      const marks: InlineMark[] = [];
      const link = n.marks.find(m => m.type.name === "link");
      if (link) {
        const a = attrsOf<LinkAttrs>(link), href = a.href ?? (a.rId ? docx.hyperlinkHref(a.rId) : null) ?? `#${a.anchor ?? ""}`;
        marks.push({ key: `link:${href}`, wrap: s => `[${s}](${destination(href)})` });
      }
      for (const [name, delimiter] of [["bold", "**"], ["italic", "*"], ["strike", "~~"], ["highlight", "=="]]) {
        if (n.marks.some(m => m.type.name === name)) marks.push({ key: name, wrap: s => s.replace(/^(\s*)([\s\S]*?)(\s*)$/, (_, before: string, body: string, after: string) => body ? `${before}${delimiter}${body}${delimiter}${after}` : s) });
      }
      for (const [name, tag] of [["underline", "u"], ["superscript", "sup"], ["subscript", "sub"]]) if (n.marks.some(m => m.type.name === name)) marks.push({ key: name, wrap: s => `<${tag}>${s}</${tag}>` });
      if (code) marks.push({ key: "code", wrap: codeSpan });
      tokens.push({ text, marks });
    });
    return marked(tokens);
  };

  const blocks = (parent: PMNode, unquote = false): string => {
    let out = "", previousList = false;
    let listIndents: number[] = [];
    parent.forEach(n => {
      let text = "", isList = false;
      if (n.type.name === "paragraph") {
        const a = attrsOf<ParagraphAttrs>(n), item = list(a);
        text = inline(n);
        let task = false;
        n.forEach(child => {
          if (child.type.name !== "opaque") return;
          const el = parseXml(attrsOf<XmlAttrs>(child).xml)?.documentElement;
          if (el?.localName === "bookmarkStart" && /^_DocWorksTask\d+$/.test(wVal(el, "name") ?? "")) task = true;
        });
        if (item) {
          isList = true;
          if (!previousList) listIndents = [item.left];
          else {
            while (listIndents.length > 1 && item.left < listIndents[listIndents.length - 1]) listIndents.pop();
            if (item.left > listIndents[listIndents.length - 1]) listIndents.push(item.left);
          }
          const indent = "    ".repeat(listIndents.length - 1);
          const taskPrefix = task && /^[☑☐]/.test(text) ? 4 : 0;
          if (task) text = text.replace(/^([☑☐])\s*/, (_, tick: string) => tick === "☑" ? "[x] " : "[ ] ");
          text = `${indent}${item.prefix}${text.replace(/\n/g, `\n${indent}${" ".repeat(item.prefix.length + taskPrefix)}`)}`;
        } else {
          const outline = chain(a.style).map(st => wVal(w(w(st, "pPr"), "outlineLvl"))).find(v => v !== null);
          const level = outline !== undefined && Number(outline) < 6 ? Number(outline) + 1 : HEADING_LEVEL[a.style ?? ""];
          if (level) text = `${"#".repeat(level)} ${text}`;
          else if (!unquote && (a.style === "Quote" || chain(a.style).some(st => wVal(st, "styleId") === "Quote"))) text = text.split("\n").map(l => `> ${l}`).join("\n");
          else text = text.replace(/^(\d+)([.)])(?=\s)/, "$1\\$2").replace(/^([+-])(?=\s|[-]+$)/, "\\$1");
        }
      } else if (n.type.name === "table") {
        const cell = n.childCount === 1 && n.firstChild?.childCount === 1 ? n.firstChild.firstChild : null;
        let codeBox = !!cell && attrsOf<CellAttrs>(cell).background?.toUpperCase() === "#F1F3F4";
        cell?.descendants(c => { if (c.isText && !c.marks.some(m => m.type.name === "font" && mono(attrsOf<{ font: string }>(m).font))) codeBox = false; });
        if (codeBox && cell) {
          const code = Array.from({ length: cell.childCount }, (_, i) => cell.child(i).textBetween(0, cell.child(i).content.size, "\n", "\n")).join("\n");
          const ticks = fence(code, 3); text = `${ticks}\n${code}\n${ticks}`;
        } else {
          const rows: string[][] = [];
          const alignments: string[] = [];
          n.forEach(row => {
            const cells: string[] = [];
            row.forEach(c => {
              const a = attrsOf<CellAttrs>(c);
              if (!rows.length) {
                const align = c.firstChild?.type.name === "paragraph" ? attrsOf<ParagraphAttrs>(c.firstChild).align : null;
                alignments.push(align === "right" ? "---:" : align === "center" ? ":---:" : "---");
                for (let i = 1; i < a.colspan; i++) alignments.push("---");
              }
              cells.push(a.vMerge === "continue" ? "" : tablePipes(blocks(c).trim().replace(/\n+/g, "<br>")));
              for (let i = 1; i < a.colspan; i++) cells.push("");
            }); rows.push(cells);
          });
          const width = Math.max(...rows.map(row => row.length));
          const line = (cells: string[]) => `| ${Array.from({ length: width }, (_, i) => cells[i] ?? "").join(" | ")} |`;
          text = [line(rows[0]), line(Array.from({ length: width }, (_, i) => alignments[i] ?? "---")), ...rows.slice(1).map(line)].join("\n");
        }
      } else if (n.type.name === "raw_block") {
        text = attrsOf<{ kind: string }>(n).kind === "hr" ? "---" : opaque(n);
      } else if (n.type.name === "sdt") {
        const shell = parseXml(attrsOf<XmlAttrs>(n).xml)?.documentElement;
        const tag = wVal(w(w(shell, "sdtPr"), "tag")) ?? "";
        if (tag.startsWith("doc-works:callout:")) {
          const [, , type, fold] = tag.split(":");
          const title = n.firstChild?.type.name === "paragraph" ? inline(n.firstChild) : "";
          const body = blocks(n.copy(n.content.cut(n.firstChild?.nodeSize ?? 0)), true);
          text = [`[!${type}]${fold === "+" || fold === "-" ? fold : ""}${title ? ` ${title}` : ""}`, ...(body ? ["", ...body.split("\n")] : [])].map(line => line ? `> ${line}` : ">").join("\n");
        } else {
          text = blocks(n, unquote);
          if (tag.startsWith("doc-works:code:")) {
            const language = tag.slice("doc-works:code:".length);
            if (/^[\w.+#-]+$/.test(language)) text = text.replace(/^(`{3,})(?=\n)/, `$1${language}`);
          }
        }
      } else if (n.childCount) text = blocks(n, unquote);
      // Display math supplies its own paragraph boundaries when mixed with inline text.
      text = text.replace(/^\n+|\n+$/g, "");
      if (!text) return;
      out += (out ? previousList && isList ? "\n" : "\n\n" : "") + text;
      previousList = isList;
    });
    return out;
  };
  let markdown = blocks(doc).replace(/^\n+|\n+$/g, "");
  for (const [id, text] of usedNotes) markdown += `\n\n[^${id}]: ${text.replace(/\n/g, "\n    ")}`;
  const properties = docx.documentVariable("DocWorksFrontmatter");
  if (properties && /^---\r?\n[\s\S]*?\r?\n---(?:\r?\n|$)$/.test(properties)) markdown = `${properties.trimEnd()}\n\n${markdown}`;
  return { markdown: markdown ? `${markdown}\n` : "", images, warnings: [...warnings] };
}
