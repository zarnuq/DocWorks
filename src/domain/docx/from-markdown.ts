/*
 * Markdown text (pasted from a note) as document content: the Markdown is drawn as the HTML Obsidian would make
 * and handed to the note export's converter, so pasted headings, lists, tasks, code, tables, equations and vault
 * pictures match an exported note.
 */
import { Fragment, Node as PMNode, Slice } from "prosemirror-model";
import { EditorState, Selection, Transaction } from "prosemirror-state";
import { blockRuns, MdBlock, MdRun, parseMarkdown } from "../../core/utils/markdown";
import { htmlToDoc } from "./from-html";
import { attrsOf } from "./schema";

type Ctx = Parameters<typeof htmlToDoc>[1];

type Pictures = Map<string, ReturnType<Ctx["image"]>>;
/** An equation as Obsidian's rendered math carries its source. */
const mathEl = (parent: HTMLElement, tag: "span" | "div", latex: string) =>
  parent.createEl(tag, { cls: ["math", tag === "div" ? "math-block" : "math-inline"] }).createEl("annotation" as "span", { text: latex, attr: { encoding: "application/x-tex" } });

function appendAll(el: HTMLElement, runs: MdRun[], pictures: Pictures): void {
  for (const run of runs) {
    if (run.math) {
      mathEl(el, "span", run.text);
      continue;
    }
    // A picture that couldn't be loaded keeps its name in the text.
    if (run.image && pictures.get(run.image)) {
      el.createEl("img", { attr: { src: run.image, alt: run.text } });
      continue;
    }
    let target = el;
    const wrap = (tag: keyof HTMLElementTagNameMap, attr?: Record<string, string>) => (target = target.createEl(tag, attr ? { attr } : undefined));
    if (run.href) wrap("a", { href: run.href });
    if (run.b) wrap("strong");
    if (run.i) wrap("em");
    if (run.strike) wrap("del");
    if (run.highlight) wrap("mark");
    if (run.code) wrap("code");
    target.appendText(run.text);
  }
}

/** The blocks as the HTML Obsidian renders for them. */
function markdownDom(blocks: MdBlock[], pictures: Pictures): HTMLElement {
  const appendRuns = (el: HTMLElement, runs: MdRun[]) => appendAll(el, runs, pictures);
  const root = createDiv();
  let lists: { depth: number; el: HTMLElement }[] = [];
  // The body of the callout the quote lines being read belong to.
  let callout: HTMLElement | null = null;
  for (const b of blocks) {
    if (b.kind !== "item" && b.kind !== "blank") lists = [];
    if (b.kind !== "quote") callout = null;
    switch (b.kind) {
      case "blank":
        break;
      case "hr":
        root.createEl("hr");
        break;
      case "para":
        appendRuns(root.createEl("p"), b.runs);
        break;
      case "heading":
        appendRuns(root.createEl(`h${b.level}` as "h1"), b.runs);
        break;
      case "quote":
        if (b.callout) {
          // As Obsidian renders a callout, which the note export's converter turns into its tagged box.
          const box = root.createDiv({ cls: "callout", attr: { "data-callout": b.callout.type, "data-callout-fold": b.callout.fold } });
          appendRuns(box.createDiv({ cls: "callout-title" }).createDiv({ cls: "callout-title-inner" }), b.runs);
          callout = box.createDiv({ cls: "callout-content" });
        } else appendRuns((callout ?? root.createEl("blockquote")).createEl("p"), b.runs);
        break;
      case "math":
        mathEl(root, "div", b.latex);
        break;
      case "code":
        root.createEl("pre").createEl("code", { text: b.text, cls: b.lang ? `language-${b.lang}` : undefined });
        break;
      case "table": {
        const table = root.createEl("table");
        b.rows.forEach((row, r) => {
          const tr = table.createEl(r ? "tbody" : "thead").createEl("tr");
          for (const cell of row) appendRuns(tr.createEl(r ? "td" : "th"), cell);
        });
        break;
      }
      case "item": {
        const tag = b.list === "ordered" ? "ol" : "ul";
        while (lists.length && lists[lists.length - 1].depth > b.depth) lists.pop();
        if (lists.length && lists[lists.length - 1].depth === b.depth && lists[lists.length - 1].el.tagName !== tag.toUpperCase()) lists.pop();
        if (!lists.length || lists[lists.length - 1].depth < b.depth) {
          const parent = lists.length ? (lists[lists.length - 1].el.lastElementChild as HTMLElement | null) ?? root : root;
          lists.push({ depth: b.depth, el: parent.createEl(tag) });
        }
        const li = lists[lists.length - 1].el.createEl("li");
        if (b.list === "task") {
          li.addClass("task-list-item");
          li.createEl("input", { type: "checkbox", attr: b.checked ? { checked: "" } : {} });
        }
        appendRuns(li, b.runs);
        break;
      }
    }
  }
  return root;
}

/**
 * Markdown text as a document (headings, lists, tasks, quotes, code boxes, tables, rules, inline marks, equations).
 * `ctx.image` is asked once for each picture, with an <img> whose src is the link path.
 */
export function markdownToDoc(text: string, ctx: Ctx): PMNode {
  const blocks = parseMarkdown(text);
  const pictures: Pictures = new Map();
  for (const run of blocks.flatMap(blockRuns)) if (run.image && !pictures.has(run.image)) pictures.set(run.image, ctx.image(createEl("img", { attr: { src: run.image, alt: run.text } })));
  return htmlToDoc(markdownDom(blocks, pictures), { ...ctx, image: (img) => pictures.get(img.getAttribute("src") ?? "") ?? null });
}

const plainParagraph = (n: PMNode | null) => n?.type.name === "paragraph" && !attrsOf<{ style: string | null; numId: string | null }>(n).style && !n.attrs.numId;

/** The document as a paste: a plain first / last paragraph joins the paragraph at the caret, anything else stays a block. */
export function pasteSlice(doc: PMNode): Slice {
  return new Slice(Fragment.from(doc.content), plainParagraph(doc.firstChild) ? 1 : 0, plainParagraph(doc.lastChild) ? 1 : 0);
}

/**
 * The transaction that pastes `doc` over the selection, with the caret at the end of what was pasted (as in Word and
 * Google Docs). A slice that ends in a closed block (a list item, a heading) would otherwise leave ProseMirror's caret
 * at the start of the paragraph after it.
 */
export function pasteMarkdown(state: EditorState, doc: PMNode): Transaction {
  const tr = state.tr.replaceSelection(pasteSlice(doc));
  let end = -1;
  tr.mapping.maps[tr.mapping.maps.length - 1]?.forEach((_from, _to, _newFrom, newTo) => (end = newTo));
  return end < 0 ? tr : tr.setSelection(Selection.near(tr.doc.resolve(end), -1));
}
