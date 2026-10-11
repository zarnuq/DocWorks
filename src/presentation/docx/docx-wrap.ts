/**
 * Text on both sides of a floating picture (see domain/docx/wrap.ts): each paragraph the picture covers is copied
 * into a hidden sandbox with the editor's styles, blanks are planned into the copy one line at a time, and the
 * blanks become widgets at the same text positions in the editor.
 */
import { Plugin, PluginKey } from "prosemirror-state";
import { Decoration, DecorationSet, EditorView } from "prosemirror-view";
import { Box, justify, TextLayout, wrapAround } from "../../domain/docx/wrap";
import { repaginate } from "./docx-pages";

const key = new PluginKey<DecorationSet>("docx-wrap-both");
const schedulers = new WeakMap<EditorView, () => void>();

/** Plan the blanks again (after the page width or zoom changed). */
export function rewrap(view: EditorView): void {
  schedulers.get(view)?.();
}

/** A paragraph's characters in DOM order. */
function characters(root: Node): { node: Text; offset: number }[] {
  const out: { node: Text; offset: number }[] = [];
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  for (let n = walker.nextNode() as Text | null; n; n = walker.nextNode() as Text | null) for (let i = 0; i < n.data.length; i++) out.push({ node: n, offset: i });
  return out;
}

interface Planned {
  gaps: { node: Text; offset: number; width: number }[];
  /** Word spacing for character ranges (justified lines beside a picture). */
  spacing: { from: { node: Text; offset: number }; to: { node: Text; offset: number }; px: number }[];
  growth: number;
}

/** Blanks for one paragraph (`el`, laid out in `sandbox`), keeping its text out of `areas` (paragraph coordinates). */
function planParagraph(el: HTMLElement, sandbox: HTMLElement, areas: Box[]): Planned {
  const copy = sandbox.appendChild(el.cloneNode(true) as HTMLElement);
  try {
    copy.querySelectorAll(".docx-wrap-gap").forEach((g) => g.remove());
    copy.querySelectorAll<HTMLElement>(".docx-wrap-spaced").forEach((g) => g.style.removeProperty("word-spacing"));
    const style = getComputedStyle(el);
    const justified = style.textAlign === "justify";
    // Planned unjustified; justified lines get their own spacing per side afterwards.
    if (justified) copy.setCssStyles({ textAlign: "left" });
    const original = characters(el);
    const list = characters(copy);
    const text = list.map((c) => c.node.data[c.offset]).join("");
    const before = copy.getBoundingClientRect().height;
    const range = document.createRange();
    const layout: TextLayout = {
      text,
      box(k) {
        const c = list[k];
        range.setStart(c.node, c.offset);
        range.setEnd(c.node, c.offset + 1);
        const r = range.getBoundingClientRect();
        if (!r.width && !r.height) return null;
        const o = copy.getBoundingClientRect();
        return { left: r.left - o.left, right: r.right - o.left, top: r.top - o.top, bottom: r.bottom - o.top };
      },
      insert(k, width) {
        const c = list[k];
        let node = c.node;
        if (c.offset > 0) {
          const rest = node.splitText(c.offset);
          for (let j = k; j < list.length && list[j].node === node; j++) list[j] = { node: rest, offset: list[j].offset - c.offset };
          node = rest;
        }
        node.parentNode!.insertBefore(gapElement(width), node);
      },
    };
    let gaps = wrapAround(layout, areas);
    let spacing: { from: number; to: number; px: number }[] = [];
    if (justified && gaps.length) {
      const right = copy.getBoundingClientRect().width - (Number.parseFloat(style.paddingRight) || 0) - (Number.parseFloat(style.borderRightWidth) || 0);
      ({ gaps, spacing } = justify(layout, gaps, areas, right));
    }
    return {
      gaps: gaps.map((g) => ({ ...original[g.at], width: g.width })),
      spacing: spacing.map((sp) => ({ from: original[sp.from], to: original[sp.to - 1], px: sp.px })),
      growth: copy.getBoundingClientRect().height - before,
    };
  } finally {
    copy.remove();
  }
}

function gapElement(width: number): HTMLElement {
  const el = createSpan({ cls: "docx-wrap-gap", attr: { contenteditable: "false" } });
  el.setCssProps({ "--docx-wrap-gap": `${width}px` });
  return el;
}

function plan(view: EditorView): Decoration[] {
  const doc = view.state.doc;
  const decos: Decoration[] = [];
  const pictures: { el: HTMLElement; index: number; distances: { top: number; right: number; bottom: number; left: number } }[] = [];
  doc.forEach((node, pos, index) => {
    if (node.type.name !== "paragraph") return;
    node.forEach((child, offset) => {
      if (child.type.name !== "image") return;
      const el = view.nodeDOM(pos + 1 + offset);
      if (el instanceof HTMLElement && el.classList.contains("docx-both"))
        pictures.push({ el, index, distances: (child.attrs.distances as typeof pictures[0]["distances"] | null) ?? { top: 0, right: 12, bottom: 0, left: 12 } });
    });
  });
  if (!pictures.length) return decos;
  const zoom = view.dom.style.zoom;
  view.dom.setCssProps({ zoom: "1" });
  const sandbox = view.dom.cloneNode(false) as HTMLElement;
  sandbox.setAttribute("contenteditable", "false");
  sandbox.setAttribute("aria-hidden", "true");
  sandbox.addClass("docx-wrap-sandbox");
  sandbox.setCssProps({ zoom: "1", "--docx-sandbox-left": `${view.dom.offsetLeft}px`, "--docx-sandbox-width": `${view.dom.getBoundingClientRect().width}px` });
  view.dom.parentElement?.appendChild(sandbox);
  try {
    // Pictures' areas as laid out without blanks; each moves down with its paragraph as blanks above it add lines.
    const areas = pictures.flatMap((pic) => {
      const r = pic.el.getBoundingClientRect();
      const d = pic.distances;
      return r.width && r.height ? [{ index: pic.index, box: { left: r.left - d.left, right: r.right + d.right, top: r.top - d.top, bottom: r.bottom + d.bottom } }] : [];
    });
    if (!areas.length) return decos;
    const starts: number[] = [];
    doc.forEach((_, pos) => starts.push(pos));
    /** How far blanks above it pushed each paragraph down. */
    const pushed: number[] = [];
    let growth = 0;
    for (let i = Math.min(...areas.map((a) => a.index)); i < doc.childCount; i++) {
      pushed[i] = growth;
      const el = view.nodeDOM(starts[i]);
      if (!(el instanceof HTMLElement)) continue;
      const rect = el.getBoundingClientRect();
      const here = areas.filter((a) => a.index <= i).map((a) => {
        const shift = pushed[a.index] - growth;
        return { left: a.box.left - rect.left, right: a.box.right - rect.left, top: a.box.top - rect.top + shift, bottom: a.box.bottom - rect.top + shift };
      });
      const reaching = here.filter((a) => a.bottom > 0 && a.top < rect.height);
      if (!here.some((a) => a.bottom > 0) && areas.every((a) => a.index <= i)) break;
      if (doc.child(i).type.name !== "paragraph" || !reaching.length) continue;
      const result = planParagraph(el, sandbox, reaching);
      growth += result.growth;
      for (const g of result.gaps) {
        const at = view.posAtDOM(g.node, g.offset);
        decos.push(Decoration.widget(at, () => gapElement(g.width), { side: -1, key: `wrap-${at}-${Math.round(g.width * 10)}`, ignoreSelection: true, marks: [] }));
      }
      for (const sp of result.spacing) {
        const from = view.posAtDOM(sp.from.node, sp.from.offset), to = view.posAtDOM(sp.to.node, sp.to.offset) + 1;
        if (to > from) decos.push(Decoration.inline(from, to, { class: "docx-wrap-spaced", style: `word-spacing: ${sp.px.toFixed(2)}px` }));
      }
    }
  } finally {
    sandbox.remove();
    view.dom.style.zoom = zoom;
  }
  return decos;
}

/** Keeps text out of pictures that wrap text on both sides. */
export function wrapBothSidesPlugin(): Plugin<DecorationSet> {
  return new Plugin<DecorationSet>({
    key,
    state: {
      init: () => DecorationSet.empty,
      apply: (tr, old) => (tr.getMeta(key) as DecorationSet | undefined) ?? old.map(tr.mapping, tr.doc),
    },
    props: { decorations: (state) => key.getState(state) },
    view(view) {
      let timer = 0;
      const run = () => {
        timer = 0;
        if (!view.dom.isConnected) return;
        const had = key.getState(view.state)!.find().length > 0;
        // Plan against the text as it flows without blanks.
        if (had) view.dispatch(view.state.tr.setMeta(key, DecorationSet.empty).setMeta("addToHistory", false));
        const decos = plan(view);
        if (decos.length) view.dispatch(view.state.tr.setMeta(key, DecorationSet.create(view.state.doc, decos)).setMeta("addToHistory", false));
        if (had || decos.length) repaginate(view);
      };
      const schedule = () => {
        if (!timer) timer = window.setTimeout(run, 60);
      };
      schedulers.set(view, schedule);
      schedule();
      return {
        update: (v, prev) => {
          if (v.state.doc !== prev.doc) schedule();
        },
        destroy: () => window.clearTimeout(timer),
      };
    },
  });
}
