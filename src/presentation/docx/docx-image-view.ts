import { Node as PMNode } from "prosemirror-model";
import { EditorView, NodeView } from "prosemirror-view";
import { attrsOf, ImageAttrs } from "../../domain/docx/schema";
import { placementCss, Wrap } from "../../domain/docx/drawing";

/** Google Docs' image options, in its order. */
const WRAPS: [Wrap, string][] = [
  ["inline", "In line"],
  ["square", "Wrap text"],
  ["topBottom", "Break text"],
  ["behind", "Behind text"],
  ["front", "In front of text"],
];

export interface ImageActions {
  replace(pos: number): void;
  editAlt(pos: number, current: string): void;
}

/**
 * Inline picture with a corner resize handle and a small action bar when
 * selected. Never dispatches on load, so opening a file can't mark it dirty.
 */
export class DocxImageView implements NodeView {
  dom: HTMLElement;
  private img: HTMLImageElement;
  private wrap: HTMLSelectElement;

  constructor(private node: PMNode, private view: EditorView, private getPos: () => number | undefined, actions: ImageActions) {
    this.dom = createSpan({ cls: "docx-image" });
    this.img = this.dom.createEl("img");
    const handle = this.dom.createSpan({ cls: "docx-image-handle" });
    handle.addEventListener("mousedown", (e) => this.startResize(e));

    const bar = this.dom.createSpan({ cls: "docx-image-bar" });
    const button = (label: string, run: () => void) => {
      const b = bar.createEl("button", { text: label, cls: "dw-btn" });
      b.addEventListener("mousedown", (e) => e.preventDefault());
      b.addEventListener("click", (e) => {
        e.preventDefault();
        if (this.getPos() !== undefined) run();
      });
    };
    this.wrap = bar.createEl("select", { cls: "dropdown docx-image-wrap", attr: { "aria-label": "Text wrapping" } });
    for (const [value, label] of WRAPS) this.wrap.createEl("option", { value, text: label });
    this.wrap.addEventListener("mousedown", (e) => e.stopPropagation());
    this.wrap.addEventListener("change", () => {
      const pos = this.getPos();
      if (pos === undefined) return;
      const wrap = this.wrap.value as Wrap;
      this.view.dispatch(this.view.state.tr.setNodeMarkup(pos, undefined, { ...this.node.attrs, wrap, ...(wrap === "inline" ? { x: 0, y: 0, align: null } : {}) }));
    });
    button("50%", () => this.scale(0.5));
    button("100%", () => this.scale(1));
    button("Alt", () => actions.editAlt(this.getPos()!, attrsOf<ImageAttrs>(this.node).alt));
    button("Replace", () => actions.replace(this.getPos()!));
    button("Delete", () => {
      const pos = this.getPos()!;
      this.view.dispatch(this.view.state.tr.delete(pos, pos + 1));
    });
    this.render();
  }

  private render(): void {
    const { src, alt, width, height } = attrsOf<ImageAttrs>(this.node);
    this.img.src = src;
    this.img.alt = alt;
    this.img.style.width = width ? `${width}px` : "";
    this.img.style.height = height ? `${height}px` : "";
    // Floating pictures sit where the document puts them (see drawing.ts).
    const a = attrsOf<ImageAttrs>(this.node);
    const { cls, style } = placementCss({ wrap: a.wrap, x: a.x, y: a.y, align: a.align }, a.width ?? 0);
    this.dom.className = `docx-image ${cls}`.trim() + (this.dom.classList.contains("is-selected") ? " is-selected" : "");
    this.dom.setAttribute("style", style);
    if (this.wrap) this.wrap.value = a.wrap;
  }

  private setSize(width: number, height: number): void {
    const pos = this.getPos();
    if (pos === undefined) return;
    this.view.dispatch(this.view.state.tr.setNodeMarkup(pos, undefined, { ...this.node.attrs, width: Math.round(width), height: Math.round(height) }));
  }

  /** Scale relative to the picture's natural size, capped to the page width. */
  private scale(factor: number): void {
    const { width: w, height: h } = attrsOf<ImageAttrs>(this.node);
    const nw = this.img.naturalWidth || w || 300;
    const nh = this.img.naturalHeight || h || 200;
    const max = this.view.dom.clientWidth || nw;
    const width = Math.min(nw * factor, max);
    this.setSize(width, (width / nw) * nh);
  }

  private startResize(e: MouseEvent): void {
    e.preventDefault();
    e.stopPropagation();
    const startX = e.clientX;
    const w0 = this.img.offsetWidth;
    const ratio = this.img.offsetHeight / Math.max(1, w0);
    let width = w0;
    const move = (ev: MouseEvent) => {
      width = Math.max(24, w0 + ev.clientX - startX);
      this.img.style.width = `${width}px`;
      this.img.style.height = `${width * ratio}px`;
    };
    const up = () => {
      window.removeEventListener("mousemove", move);
      window.removeEventListener("mouseup", up);
      if (width !== w0) this.setSize(width, width * ratio);
    };
    window.addEventListener("mousemove", move);
    window.addEventListener("mouseup", up);
  }

  update(node: PMNode): boolean {
    if (node.type !== this.node.type) return false;
    this.node = node;
    this.render();
    return true;
  }

  selectNode(): void {
    this.dom.addClass("is-selected");
  }

  deselectNode(): void {
    this.dom.removeClass("is-selected");
  }

  stopEvent(e: Event): boolean {
    return (e.target as HTMLElement).closest(".docx-image-bar, .docx-image-handle") !== null;
  }

  ignoreMutation(): boolean {
    return true;
  }
}
