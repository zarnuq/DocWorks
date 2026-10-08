import type { ShapeModel, SlideModel } from "../../domain/pptx/slide-model";
import { renderParagraphs } from "./text-html";

export interface RenderOptions {
  /** Slide width in px; height follows the slide aspect ratio. */
  width: number;
  /** Make slide-owned text contentEditable. */
  editable: boolean;
  /** Media path -> URL usable in <img src>, or null if not loaded. */
  imageUrl: (mediaPath: string) => string | null;
}

export interface RenderedShape {
  shape: ShapeModel;
  el: HTMLElement;
  textEl: HTMLElement | null;
  /** px per point for this shape's text, including autofit scale. */
  pxPerPt: number;
}

const EMU_PER_PT = 12700;
const RADIUS: Record<string, string> = { ellipse: "50%", roundRect: "12%", flowChartAlternateProcess: "12%" };
const JUSTIFY = { t: "flex-start", ctr: "center", b: "flex-end" } as const;

const SVG_NS = "http://www.w3.org/2000/svg";

/** Custom geometry and lines: one stretched SVG per path. Gradients fall back to their first color. */
function drawPaths(el: HTMLElement, shape: ShapeModel, scale: number): void {
  const fill = shape.fill?.match(/#[0-9A-F]{6}(?:[0-9A-F]{2})?/i)?.[0] ?? "none";
  for (const path of shape.paths!) {
    const svg = document.createElementNS(SVG_NS, "svg");
    svg.setAttribute("viewBox", `0 0 ${path.w} ${path.h}`);
    svg.setAttribute("preserveAspectRatio", "none");
    svg.classList.add("pptx-geom");
    if (shape.flipH || shape.flipV) svg.style.transform = `scale(${shape.flipH ? -1 : 1}, ${shape.flipV ? -1 : 1})`;
    const p = svg.appendChild(document.createElementNS(SVG_NS, "path"));
    p.setAttribute("d", path.d);
    p.setAttribute("fill", path.fill ? fill : "none");
    p.setAttribute("stroke", path.stroke && shape.line ? shape.line.color : "none");
    p.setAttribute("stroke-width", String(Math.max(1, (shape.line?.width ?? 0) * scale)));
    p.setAttribute("vector-effect", "non-scaling-stroke");
    el.appendChild(svg);
  }
}

/** Draw a slide into `host` (replacing its contents) with plain DOM APIs. */
export function renderSlide(host: HTMLElement, model: SlideModel, size: { cx: number; cy: number }, opts: RenderOptions): RenderedShape[] {
  const scale = opts.width / size.cx;
  const px = (emu: number) => `${emu * scale}px`;

  const slide = createDiv();
  slide.className = "pptx-slide";
  slide.style.width = `${opts.width}px`;
  slide.style.height = px(size.cy);
  if (model.background.color) slide.style.background = model.background.color;
  const bgUrl = model.background.image && opts.imageUrl(model.background.image);
  if (bgUrl) {
    slide.style.background = `url("${bgUrl}") center / cover no-repeat`;
  }
  host.replaceChildren(slide);

  return model.shapes.map((shape, index) => {
    const el = slide.createDiv();
    el.className = `pptx-shape pptx-kind-${shape.kind}`;
    el.classList.toggle("is-inherited", shape.inherited);
    el.dataset.index = String(index);
    Object.assign(el.style, { left: px(shape.rect.x), top: px(shape.rect.y), width: px(shape.rect.w), height: px(shape.rect.h) });
    if (shape.rot) el.style.transform = `rotate(${shape.rot}deg)`;
    if (shape.paths) {
      drawPaths(el, shape, scale);
    } else {
      if (shape.fill) el.style.background = shape.fill;
      if (shape.line) el.style.border = `${Math.max(1, shape.line.width * scale)}px solid ${shape.line.color}`;
      if (RADIUS[shape.geom]) el.style.borderRadius = RADIUS[shape.geom];
    }

    if (shape.image) {
      const img = el.createEl("img");
      img.src = opts.imageUrl(shape.image) ?? "";
      img.draggable = false;
    }

    if (shape.table) {
      const table = el.createEl("table");
      table.style.fontSize = `${14 * EMU_PER_PT * scale}px`;
      for (const row of shape.table) {
        const tr = table.createEl("tr");
        for (const cell of row) tr.createEl("td").textContent = cell;
      }
    } else if (shape.kind === "frame") {
      el.createDiv().className = "pptx-frame-label";
      el.lastElementChild!.textContent = "Chart or diagram (view only)";
    }

    const pxPerPt = EMU_PER_PT * scale * (shape.text?.fontScale ?? 1);
    let textEl: HTMLElement | null = null;
    if (shape.text) {
      const t = shape.text;
      textEl = el.createDiv();
      textEl.className = "pptx-text";
      if (shape.textRect) {
        const r = shape.textRect;
        Object.assign(textEl.style, { left: px(r.x), top: px(r.y), width: px(r.w), height: px(r.h) });
      }
      textEl.style.justifyContent = JUSTIFY[t.anchor];
      textEl.style.padding = [t.insets.t, t.insets.r, t.insets.b, t.insets.l].map(px).join(" ");
      renderParagraphs(textEl, t.paras, pxPerPt, scale, t.lnSpcReduction);
      textEl.classList.toggle("is-empty", shape.placeholder && t.paras.every((p) => p.runs.length === 0));
      if (opts.editable && !shape.inherited) {
        textEl.contentEditable = "true";
        textEl.spellcheck = true;
      }
    }
    return { shape, el, textEl, pxPerPt };
  });
}
