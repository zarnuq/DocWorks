import { Node as PMNode } from "prosemirror-model";
import { EditorView, NodeView } from "prosemirror-view";
import { attrsOf, cropBox, ImageAttrs } from "../../domain/docx/schema";
import { placementCss, PlacementGeometry, recrop, Wrap } from "../../domain/docx/drawing";

/** Google Docs' image options, in its order. */
export const WRAPS: [Wrap, string][] = [
  ["inline", "In line"],
  ["square", "Wrap text"],
  ["topBottom", "Break text"],
  ["behind", "Behind text"],
  ["front", "In front of text"],
];

export interface ImageActions {
  replace(pos: number): void;
  editAlt(pos: number, current: string): void;
  /** Ask for a new crop (shares cut off each side); null removes it. */
  crop(current: ImageAttrs["crop"]): Promise<ImageAttrs["crop"] | undefined>;
}

/**
 * Inline picture with a corner resize handle and a small action bar when
 * selected. Never dispatches on load, so opening a file can't mark it dirty.
 */
export class DocxImageView implements NodeView {
  dom: HTMLElement;
  private img: HTMLImageElement;
  private wrap: HTMLSelectElement;
  private moveHandle: HTMLButtonElement;
  private cancelGesture: (() => void) | null = null;
  private mountTimer = 0;

  constructor(private node: PMNode, private view: EditorView, private getPos: () => number | undefined, actions: ImageActions) {
    this.dom = createSpan({ cls: "docx-image" });
    this.img = this.dom.createEl("img");
    const handle = this.dom.createSpan({ cls: "docx-image-handle" });
    handle.addEventListener("mousedown", (e) => this.startResize(e));

    const bar = this.dom.createSpan({ cls: "docx-image-bar" });
    this.moveHandle = bar.createEl("button", { text: "Move", cls: "dw-btn", attr: { "aria-label": "Move image", title: "Drag to position the image" } });
    this.moveHandle.addEventListener("mousedown", e => this.startMove(e));
    this.img.draggable = false;
    this.img.addEventListener("mousedown", e => {
      if (this.dom.classList.contains("is-selected") && attrsOf<ImageAttrs>(this.node).wrap !== "inline") this.startMove(e);
    });
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
    const spacing = bar.createEl("input", { type: "number", attr: { "aria-label": "Image text spacing (px)", title: "Space between image and text (px)", min: "0", max: "96", step: "1" } });
    spacing.addEventListener("change", () => {
      const value = Number(spacing.value);
      const pos = this.getPos();
      if (pos === undefined || !Number.isFinite(value) || value < 0 || value > 96) return;
      this.view.dispatch(this.view.state.tr.setNodeMarkup(pos, undefined, { ...this.node.attrs, distances: { top: value, right: value, bottom: value, left: value } }));
    });
    button("Crop", () => void actions.crop(attrsOf<ImageAttrs>(this.node).crop).then((crop) => {
      const pos = this.getPos();
      if (crop === undefined || pos === undefined) return;
      const a = attrsOf<ImageAttrs>(this.node);
      this.view.dispatch(this.view.state.tr.setNodeMarkup(pos, undefined, { ...this.node.attrs, ...recrop(a, crop) }));
    }));
    button("Alt", () => actions.editAlt(this.getPos()!, attrsOf<ImageAttrs>(this.node).alt));
    button("Replace", () => actions.replace(this.getPos()!));
    button("Delete", () => {
      const pos = this.getPos()!;
      this.view.dispatch(this.view.state.tr.delete(pos, pos + 1));
    });
    this.render();
    this.mountTimer = window.setTimeout(() => { this.mountTimer = 0; this.render(); }, 0);
  }

  private render(): void {
    const { src, alt, width, height } = attrsOf<ImageAttrs>(this.node);
    this.img.src = src;
    this.img.alt = alt;
    // Cropped: the box keeps the picture's size and clips the whole image, scaled and shifted inside it.
    const crop = attrsOf<ImageAttrs>(this.node).crop;
    const box = crop && width && height ? cropBox(width, height, crop) : null;
    this.img.style.width = box ? `${box.width}px` : width ? `${width}px` : "";
    this.img.style.height = box ? `${box.height}px` : height ? `${height}px` : "";
    this.img.style.marginLeft = box ? `${box.left}px` : "";
    this.img.style.marginTop = box ? `${box.top}px` : "";
    // Floating pictures sit where the document puts them (see drawing.ts).
    const a = attrsOf<ImageAttrs>(this.node);
    const { cls, style } = placementCss(a, a.width ?? 0, this.geometry());
    this.dom.className = `docx-image ${cls}`.trim() + (box ? " docx-cropped" : "") + (this.dom.classList.contains("is-selected") ? " is-selected" : "");
    this.dom.setAttribute("style", style + (box ? `; --docx-crop-w: ${width}px; --docx-crop-h: ${height}px` : ""));
    if (this.wrap) this.wrap.value = a.wrap;
    this.moveHandle.disabled = a.wrap === "inline";
    const spacing = this.dom.querySelector<HTMLInputElement>('[aria-label="Image text spacing (px)"]');
    if (spacing) { spacing.value = String(a.distances?.left ?? 12); spacing.disabled = a.wrap === "inline"; }
  }

  private geometry(): PlacementGeometry {
    const paragraph = this.dom.closest<HTMLElement>("p, h1, h2, h3, h4, h5, h6");
    const host = this.view.dom.closest<HTMLElement>(".docx-editor-host");
    const p = paragraph?.getBoundingClientRect(), editor = this.view.dom.getBoundingClientRect();
    const zoom = this.view.dom.offsetWidth && editor.width ? editor.width / this.view.dom.offsetWidth : 1;
    const marginLeft = Number.parseFloat(host?.style.getPropertyValue("--docx-m-left") ?? "") || 0;
    const marginTop = Number.parseFloat(host?.style.getPropertyValue("--docx-m-top") ?? "") || 0;
    let pageTop = editor.top;
    for (const gap of Array.from(this.view.dom.querySelectorAll(".docx-page-space"))) {
      const rect = gap.getBoundingClientRect();
      if (p && rect.bottom <= p.top) pageTop = Math.max(pageTop, rect.bottom);
    }
    return { columnWidth: paragraph?.clientWidth || this.view.dom.clientWidth || undefined,
      pageLeft: p ? (p.left - editor.left) / zoom : marginLeft, marginLeft: p ? (p.left - editor.left) / zoom - marginLeft : 0,
      pageTop: marginTop, paragraphTop: p ? (p.top - pageTop) / zoom : marginTop };
  }

  private startMove(e: MouseEvent): void {
    const a = attrsOf<ImageAttrs>(this.node);
    if (e.button !== 0 || a.wrap === "inline") return;
    e.preventDefault();
    e.stopPropagation();
    this.cancelGesture?.();
    const startX = e.clientX, startY = e.clientY;
    const geometry = this.geometry();
    const column = geometry.columnWidth ?? 624;
    const x0 = a.align === "right" ? column - (a.width ?? 0) : a.align === "center" ? (column - (a.width ?? 0)) / 2 : a.x;
    const domRect = this.view.dom.getBoundingClientRect();
    const zoom = this.view.dom.offsetWidth && domRect.width ? domRect.width / this.view.dom.offsetWidth : 1;
    let x = x0, y = a.y;
    const cleanup = () => {
      window.removeEventListener("mousemove", move);
      window.removeEventListener("mouseup", up);
      window.removeEventListener("keydown", key);
      this.cancelGesture = null;
      this.dom.removeClass("is-moving");
    };
    const move = (ev: MouseEvent) => {
      x = Math.round(x0 + (ev.clientX - startX) / zoom);
      y = Math.round(a.y + (ev.clientY - startY) / zoom);
      this.dom.setAttribute("style", placementCss({ ...a, x, y, align: null }, a.width ?? 0, geometry).style);
    };
    const up = () => {
      cleanup();
      const pos = this.getPos();
      if (pos !== undefined && (x !== x0 || y !== a.y)) this.view.dispatch(this.view.state.tr.setNodeMarkup(pos, undefined, { ...this.node.attrs, x, y, align: null }));
      else this.render();
    };
    const key = (ev: KeyboardEvent) => { if (ev.key === "Escape") { ev.preventDefault(); cleanup(); this.render(); } };
    this.cancelGesture = () => { cleanup(); this.render(); };
    this.dom.addClass("is-moving");
    window.addEventListener("mousemove", move);
    window.addEventListener("mouseup", up);
    window.addEventListener("keydown", key);
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
    const rect = this.view.dom.getBoundingClientRect();
    const zoom = this.view.dom.offsetWidth && rect.width ? rect.width / this.view.dom.offsetWidth : 1;
    let width = w0;
    const move = (ev: MouseEvent) => {
      width = Math.max(24, w0 + (ev.clientX - startX) / zoom);
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

  destroy(): void { window.clearTimeout(this.mountTimer); this.cancelGesture?.(); }
}
