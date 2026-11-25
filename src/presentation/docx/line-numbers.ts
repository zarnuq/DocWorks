/**
 * Line numbers in the page view (Word's w:lnNumType): a number in the left margin beside each line of body text,
 * counted as `lineNumberLabels` says. Tables, headers, footers and text boxes aren't numbered, as in Word.
 * Drawn after each page layout into one zero-height widget at the top of the text area. No Obsidian imports.
 */
import { Node as PMNode } from "prosemirror-model";
import { Plugin, PluginKey } from "prosemirror-state";
import { Decoration, DecorationSet, EditorView } from "prosemirror-view";
import { lineNumberLabels, lineNumbersOf, LineNumbers } from "../../domain/docx/sections";
import { docxSchema as schema } from "../../domain/docx/schema";

const key = new PluginKey<DecorationSet>("docx-line-numbers");
/** Word's "Auto" distance from the text: 0.25 in. */
const DISTANCE = 24;
let generation = 0;

export function lineNumbersPlugin(): Plugin<DecorationSet> {
  return new Plugin<DecorationSet>({
    key,
    state: {
      init: () => DecorationSet.empty,
      apply: (tr, old) => (tr.getMeta(key) as DecorationSet | undefined) ?? old.map(tr.mapping, tr.doc),
    },
    props: { decorations: (state) => key.getState(state) },
  });
}

/** Body paragraphs in order (content controls and column sections opened up), with their section's numbering. */
function numberedParagraphs(doc: PMNode): { node: PMNode; pos: number; section: number; ln: LineNumbers | null; columns: PMNode | null }[] {
  const found: { node: PMNode; pos: number; section: number; columns: PMNode | null }[] = [];
  const ends: (string | null)[] = [];
  const walk = (parent: PMNode, start: number, columns: PMNode | null) =>
    parent.forEach((node, offset) => {
      const pos = start + offset;
      if (node.type === schema.nodes.sdt) walk(node, pos + 1, columns);
      else if (node.type === schema.nodes.columns) walk(node, pos + 1, node);
      else if (node.type === schema.nodes.paragraph) {
        found.push({ node, pos, section: ends.length, columns });
        if (node.attrs.sectPr) ends.push(node.attrs.sectPr as string);
      }
    });
  walk(doc, 0, null);
  const settings = [...ends, doc.attrs.sectPr as string | null].map(lineNumbersOf);
  return found.map((p) => ({ ...p, ln: settings[p.section] }));
}

/** A paragraph's line boxes (tops and heights, screen px), leaving out page gaps drawn inside it. */
function lineBoxes(el: HTMLElement): { top: number; bottom: number }[] {
  const range = document.createRange();
  const rects: DOMRect[] = [];
  el.childNodes.forEach((n) => {
    if (n.nodeType === Node.ELEMENT_NODE && (n as Element).matches(".docx-page-spacer, .docx-page-gap")) return;
    range.selectNode(n);
    rects.push(...Array.from(range.getClientRects()));
  });
  if (!rects.some((r) => r.height)) rects.push(el.getBoundingClientRect());
  const lines: { top: number; bottom: number }[] = [];
  for (const r of rects) {
    if (!r.height) continue;
    const line = lines.find((l) => r.top < l.bottom - 2 && r.bottom > l.top + 2);
    if (line) {
      line.top = Math.min(line.top, r.top);
      line.bottom = Math.max(line.bottom, r.bottom);
    } else lines.push({ top: r.top, bottom: r.bottom });
  }
  return lines.sort((a, b) => a.top - b.top);
}

/** Draw (or clear, `paged` false) the numbers for the current layout. */
export function drawLineNumbers(view: EditorView, paged: boolean): void {
  if (!key.getState(view.state)) return;
  const doc = view.state.doc;
  const paras = paged ? numberedParagraphs(doc).filter((p) => p.ln) : [];
  if (!paras.length) {
    if (key.getState(view.state) !== DecorationSet.empty) view.dispatch(view.state.tr.setMeta(key, DecorationSet.empty).setMeta("addToHistory", false));
    return;
  }
  const zoom = Number.parseFloat(view.dom.style.zoom) || 1;
  const box = view.dom.getBoundingClientRect();
  const cs = getComputedStyle(view.dom);
  const textLeft = Number.parseFloat(cs.paddingLeft) || 0;
  const textTop = Number.parseFloat(cs.paddingTop) || 0;
  const spaces = Array.from(view.dom.querySelectorAll(".docx-page-space")).map((s) => s.getBoundingClientRect().bottom);
  const lines: { y: number; h: number; x: number; page: number; section: number; ln: LineNumbers }[] = [];
  for (const p of paras) {
    const el = view.nodeDOM(p.pos);
    if (!(el instanceof HTMLElement)) continue;
    // The section's text edge: the paragraph's box less its indent (a section with other margins is moved by `translate`).
    const edge = (el.getBoundingClientRect().left - box.left) / zoom - (Number.parseFloat(getComputedStyle(el).marginLeft) || 0);
    let colLeft = (_: number) => edge;
    if (p.columns) {
      const $pos = doc.resolve(p.pos);
      let d = $pos.depth;
      while (d > 1 && $pos.node(d) !== p.columns) d--;
      const cols = view.nodeDOM($pos.before(d));
      if (cols instanceof HTMLElement) {
        const r = cols.getBoundingClientRect();
        const count = p.columns.attrs.count as number;
        const gap = ((p.columns.attrs.space as number) / 15) * zoom;
        const width = (r.width - gap * (count - 1)) / count;
        colLeft = (x) => (r.left - box.left + Math.max(0, Math.min(count - 1, Math.floor((x - r.left) / (width + gap)))) * (width + gap)) / zoom;
      }
    }
    const left = el.getBoundingClientRect().left;
    for (const l of lineBoxes(el)) {
      lines.push({
        y: (l.top - box.top) / zoom,
        h: (l.bottom - l.top) / zoom,
        x: colLeft(left),
        page: spaces.filter((b) => b <= l.top).length,
        section: p.section,
        ln: p.ln!,
      });
    }
  }
  const labels = lineNumberLabels(lines);
  const gen = ++generation;
  const widget = Decoration.widget(
    0,
    () => {
      const layer = createDiv({ cls: "docx-line-numbers", attr: { contenteditable: "false", "aria-hidden": "true" } });
      lines.forEach((l, i) => {
        if (labels[i] === null) return;
        const n = layer.createDiv({ cls: "docx-line-number", text: String(labels[i]) });
        // The layer sits at the top left of the text area.
        n.setCssProps({ "--docx-ln-top": `${l.y - textTop}px`, "--docx-ln-height": `${l.h}px`, "--docx-ln-end": `${l.x - textLeft - DISTANCE}px` });
      });
      return layer;
    },
    { side: -2, key: `line-numbers-${gen}`, ignoreSelection: true }
  );
  view.dispatch(view.state.tr.setMeta(key, DecorationSet.create(doc, [widget])).setMeta("addToHistory", false));
}
