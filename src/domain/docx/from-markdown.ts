/*
 * Markdown text (pasted from a note) as document content: the Markdown is drawn as the HTML Obsidian would make
 * and handed to the note export's converter, so pasted headings, lists, tasks, code and tables match an exported note.
 */
import { Fragment, Node as PMNode, Slice } from "prosemirror-model";
import { MdBlock, MdRun, parseMarkdown } from "../../core/utils/markdown";
import { htmlToDoc } from "./from-html";
import { attrsOf } from "./schema";

type Ctx = Parameters<typeof htmlToDoc>[1];

function appendRuns(el: HTMLElement, runs: MdRun[]): void {
  for (const run of runs) {
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
function markdownDom(blocks: MdBlock[]): HTMLElement {
  const root = createDiv();
  let lists: { depth: number; el: HTMLElement }[] = [];
  for (const b of blocks) {
    if (b.kind !== "item") lists = [];
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
        appendRuns(root.createEl("blockquote").createEl("p"), b.runs);
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

/** Markdown text as a document (headings, lists, tasks, quotes, code boxes, tables, rules, inline marks). */
export function markdownToDoc(text: string, ctx: Ctx): PMNode {
  return htmlToDoc(markdownDom(parseMarkdown(text)), ctx);
}

const plainParagraph = (n: PMNode | null) => n?.type.name === "paragraph" && !attrsOf<{ style: string | null; numId: string | null }>(n).style && !n.attrs.numId;

/** The document as a paste: a plain first / last paragraph joins the paragraph at the caret, anything else stays a block. */
export function pasteSlice(doc: PMNode): Slice {
  return new Slice(Fragment.from(doc.content), plainParagraph(doc.firstChild) ? 1 : 0, plainParagraph(doc.lastChild) ? 1 : 0);
}
