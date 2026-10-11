import { Node as PMNode } from "prosemirror-model";
import { EditorState, Transaction } from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import { NS } from "../../core/utils/xml";
import { attrsOf, docxSchema as schema, HEADING_LEVEL, ParagraphAttrs } from "../../domain/docx/schema";
import { DEFAULT_SECTPR } from "../../domain/docx/sections";
import { pageSetup, PageSetup } from "../../domain/docx/pages";
import { cmd } from "./docx-editor";

export function outlineEntries(doc: PMNode): { pos: number; level: number; text: string }[] {
  const entries: { pos: number; level: number; text: string }[] = [];
  doc.descendants((node, pos) => {
    if (node.type !== schema.nodes.paragraph) return true;
    const { style } = attrsOf<ParagraphAttrs>(node);
    const level = style ? HEADING_LEVEL[style] : undefined;
    if (level && node.textContent.trim()) entries.push({ pos: pos + 1, level, text: node.textContent.trim() });
    return false;
  });
  return entries;
}

export function rulerSection(state: EditorState): { pos: number | null; xml: string; setup: PageSetup } {
  let section: { pos: number | null; xml: string } = { pos: null, xml: (state.doc.attrs.sectPr as string | null) ?? DEFAULT_SECTPR };
  let found = false;
  state.doc.descendants((node, pos) => {
    if (found) return false;
    if (node.type === schema.nodes.table) return false;
    if (node.type === schema.nodes.paragraph && pos + node.nodeSize >= state.selection.from && attrsOf<ParagraphAttrs>(node).sectPr) {
      section = { pos, xml: attrsOf<ParagraphAttrs>(node).sectPr! };
      found = true;
    }
    return node.type !== schema.nodes.paragraph;
  });
  const body = new DOMParser().parseFromString(`<w:body xmlns:w="${NS.w}">${section.xml}</w:body>`, "application/xml").documentElement;
  return { ...section, setup: pageSetup(body) };
}

export function setRulerMargins(state: EditorState, dispatch: (tr: Transaction) => void, patch: Partial<Pick<PageSetup, "left" | "right">>): void {
  const section = rulerSection(state);
  const root = new DOMParser().parseFromString(section.xml, "application/xml").documentElement;
  let margins = Array.from(root.children).find(el => el.localName === "pgMar");
  if (!margins) {
    margins = root.ownerDocument.createElementNS(NS.w, "w:pgMar");
    for (const [key, value] of Object.entries({ top: 1440, right: 1440, bottom: 1440, left: 1440, header: 720, footer: 720, gutter: 0 })) margins.setAttributeNS(NS.w, `w:${key}`, String(value));
    const after = new Set(["headerReference", "footerReference", "footnotePr", "endnotePr", "type", "pgSz"]);
    root.insertBefore(margins, Array.from(root.children).find(el => !after.has(el.localName)) ?? null);
  }
  for (const [side, px] of Object.entries(patch)) margins.setAttributeNS(NS.w, `w:${side}`, String(Math.round(px * 15)));
  const xml = new XMLSerializer().serializeToString(root);
  const tr = state.tr;
  if (section.pos === null) tr.setDocAttribute("sectPr", xml);
  else tr.setNodeMarkup(section.pos, undefined, { ...state.doc.nodeAt(section.pos)!.attrs, sectPr: xml });
  dispatch(tr);
}

/** A horizontal ruler. Drag changes commit once on mouseup; arrow keys provide precise keyboard control. */
export class DocxRuler {
  readonly dom: HTMLElement;
  private track: HTMLElement;
  private tabType: HTMLSelectElement;
  private cancel: (() => void) | null = null;
  private dragging = false;
  constructor(parent: HTMLElement, private editor: () => EditorView | null, private changed: () => void) {
    this.dom = parent.createDiv({ cls: "docx-ruler", attr: { "aria-label": "Document ruler" } });
    this.tabType = this.dom.createEl("select", { cls: "dropdown docx-ruler-tab-type", attr: { "aria-label": "Tab stop alignment" } });
    for (const type of ["left", "center", "right", "decimal"]) this.tabType.createEl("option", { value: type, text: `${type[0].toUpperCase()}${type.slice(1)} tab` });
    this.track = this.dom.createDiv({ cls: "docx-ruler-track", attr: { title: "Click to add a tab stop. Drag a tab off the ruler to remove it." } });
    this.track.addEventListener("click", e => {
      if (e.target !== this.track) return;
      const view = this.editor();
      if (!view || !this.paragraph(view.state)) return;
      const { setup } = rulerSection(view.state);
      const px = (e.clientX - this.track.getBoundingClientRect().left) / this.scale() - setup.left;
      if (px < 0 || px > setup.width - setup.left - setup.right) return;
      const tabs = attrsOf<ParagraphAttrs>(this.paragraph(view.state)!).tabs ?? [];
      const pt = Math.round(px * 0.75 * 4) / 4;
      cmd.setParagraph({ tabs: [...tabs.filter(t => Math.abs(t.pt - pt) > 2), { pt, align: this.tabType.value, leader: null }].sort((a, b) => a.pt - b.pt), tabLeader: null })(view.state, view.dispatch);
      this.changed();
    });
  }
  private paragraph(state: EditorState): PMNode | null {
    const { $from } = state.selection;
    for (let d = $from.depth; d > 0; d--) if ($from.node(d).type === schema.nodes.paragraph) return $from.node(d);
    return null;
  }
  private scale(): number { return this.track.offsetWidth ? this.track.getBoundingClientRect().width / this.track.offsetWidth : 1; }
  align(): void {
    const editor = this.editor();
    if (!editor || !this.dom.parentElement) return;
    const zoom = Number.parseFloat(this.dom.style.zoom) || 1;
    this.dom.style.marginLeft = `${(editor.dom.getBoundingClientRect().left - this.dom.parentElement.getBoundingClientRect().left) / zoom}px`;
  }
  update(): void {
    if (this.dragging) return;
    const view = this.editor();
    if (!view) return;
    const { setup } = rulerSection(view.state), paragraph = this.paragraph(view.state);
    const width = setup.width - setup.left - setup.right;
    this.dom.style.width = `${setup.width}px`;
    this.track.style.width = `${setup.width}px`;
    this.align();
    this.track.empty();
    const band = this.track.createDiv({ cls: "docx-ruler-text-band" });
    band.style.left = `${setup.left}px`;
    band.style.width = `${width}px`;
    for (let x = 0; x <= setup.width; x += 12) {
      const tick = this.track.createSpan({ cls: x % 96 === 0 ? "docx-ruler-tick is-inch" : "docx-ruler-tick", text: x % 96 === 0 ? String(x / 96) : "" });
      tick.style.left = `${x}px`;
    }
    const handle = (label: string, cls: string, initial: number, commit: (x: number, remove: boolean) => void) => {
      const button = this.track.createEl("button", { cls: `docx-ruler-handle ${cls}`, attr: { "aria-label": label, title: label } });
      button.style.left = `${initial}px`;
      button.addEventListener("click", e => e.stopPropagation());
      button.addEventListener("keydown", e => {
        if (!["ArrowLeft", "ArrowRight", "Delete", "Backspace"].includes(e.key)) return;
        e.preventDefault();
        commit(initial + (e.key === "ArrowLeft" ? -1 : e.key === "ArrowRight" ? 1 : 0), e.key === "Delete" || e.key === "Backspace");
        this.changed();
      });
      button.addEventListener("mousedown", e => {
        if (e.button !== 0) return;
        e.preventDefault();
        this.cancel?.();
        const start = e.clientX;
        let x = initial, remove = false;
        this.dragging = true;
        const cleanup = () => { window.removeEventListener("mousemove", move); window.removeEventListener("mouseup", up); window.removeEventListener("keydown", key); this.dragging = false; this.cancel = null; };
        const move = (ev: MouseEvent) => {
          x = Math.max(0, Math.min(setup.width, initial + (ev.clientX - start) / this.scale()));
          const rect = this.track.getBoundingClientRect();
          remove = ev.clientY < rect.top - 20 || ev.clientY > rect.bottom + 20;
          button.style.left = `${x}px`;
          button.toggleClass("will-remove", cls === "is-tab" && remove);
        };
        const up = () => { cleanup(); if (x !== initial || remove) { commit(x, remove); this.changed(); } this.update(); };
        const key = (ev: KeyboardEvent) => { if (ev.key === "Escape") { ev.preventDefault(); cleanup(); this.update(); } };
        this.cancel = () => { cleanup(); this.update(); };
        window.addEventListener("mousemove", move); window.addEventListener("mouseup", up); window.addEventListener("keydown", key);
      });
    };
    handle("Left margin", "is-margin", setup.left, x => setRulerMargins(view.state, view.dispatch, { left: Math.min(x, setup.width - setup.right - 96) }));
    handle("Right margin", "is-margin", setup.width - setup.right, x => setRulerMargins(view.state, view.dispatch, { right: Math.min(setup.width - x, setup.width - setup.left - 96) }));
    if (!paragraph) return;
    const a = attrsOf<ParagraphAttrs>(paragraph);
    const left = setup.left + a.indent * 48;
    const paragraphPatch = (patch: Partial<ParagraphAttrs>) => cmd.setParagraph(patch)(view.state, view.dispatch);
    handle("Left indent", "is-left-indent", left, x => paragraphPatch({ indent: Math.round((x - setup.left) / 48 * 1000) / 1000 }));
    handle("First line indent", "is-first-indent", left + (a.firstLine ?? 0) * 4 / 3, x => paragraphPatch({ firstLine: Math.round((x - left) * 0.75 * 4) / 4 }));
    handle("Right indent", "is-right-indent", setup.width - setup.right - (a.rightIndent ?? 0) * 4 / 3, x => paragraphPatch({ rightIndent: Math.max(0, Math.round((setup.width - setup.right - x) * 0.75 * 4) / 4) }));
    for (const [i, tab] of (a.tabs ?? []).entries()) if (tab.align !== "clear") handle(`${tab.align} tab at ${(tab.pt / 72).toFixed(2)} inches`, "is-tab", setup.left + tab.pt * 4 / 3, (x, remove) => {
      const tabs = (a.tabs ?? []).filter((_, j) => j !== i);
      if (!remove && x >= setup.left && x <= setup.width - setup.right) tabs.push({ ...tab, pt: Math.round((x - setup.left) * 0.75 * 4) / 4 });
      paragraphPatch({ tabs: tabs.sort((a, b) => a.pt - b.pt), tabLeader: null });
    });
  }
  destroy(): void { this.cancel?.(); this.dom.remove(); }
}
