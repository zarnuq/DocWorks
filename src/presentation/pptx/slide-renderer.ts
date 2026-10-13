import { pictureFilter } from "../../domain/pptx/picture-adjust";
import { renderChart } from "./chart-renderer";
import type { ShapeModel, SlideModel } from "../../domain/pptx/slide-model";
import { fillCss } from "../../domain/pptx/shape-style";
import { shapeIds } from "../../domain/pptx/animation";
import { renderParagraphs } from "./text-html";

export interface RenderOptions {
  /** Slide width in px; height follows the slide aspect ratio. */
  width: number;
  /** Make slide-owned text contentEditable. */
  editable: boolean;
  /** Media path -> URL usable in <img src>, or null if not loaded. */
  imageUrl: (mediaPath: string) => string | null;
  /** 1-based number shown by slide-number fields. */
  slideNumber?: number;
  playMedia?: boolean;
}

export interface RenderedShape {
  shape: ShapeModel;
  el: HTMLElement;
  textEl: HTMLElement | null;
  /** px per point for this shape's text, including autofit scale. */
  pxPerPt: number;
}

const EMU_PER_PT = 12700;
/** Placeholders PowerPoint leaves blank when empty, even while editing. */
const NO_PROMPT = new Set(["sldNum", "dt", "ftr", "hdr"]);
const JUSTIFY = { t: "flex-start", ctr: "center", b: "flex-end" } as const;

const SVG_NS = "http://www.w3.org/2000/svg";
const svgEl = <K extends keyof SVGElementTagNameMap>(parent: Element, tag: K, attrs: Record<string, string | number> = {}) => {
  const el = parent.appendChild(document.createElementNS(SVG_NS, tag));
  for (const [k, v] of Object.entries(attrs)) el.setAttribute(k, String(v));
  return el;
};

/** Scale SVG path data (M L C Q A Z with absolute coordinates) by sx, sy. */
export function scalePath(d: string, sx: number, sy: number): string {
  const f = (n: number) => String(Math.round(n * 100) / 100);
  return d.replace(/([MLCQAZ])([^MLCQAZ]*)/g, (_, op: string, args: string) => {
    const n = args.trim() ? args.trim().split(/\s+/).map(Number) : [];
    // Arc: rx ry rotation large-arc sweep x y; others: x y pairs.
    const scaled = n.map((v, i) => (op === "A" ? ([0, 5].includes(i) ? f(v * sx) : [1, 6].includes(i) ? f(v * sy) : String(v)) : f(v * (i % 2 ? sy : sx))));
    return scaled.length ? `${op}${scaled.join(" ")} ` : `${op} `;
  }).trim();
}

/** "#RRGGBB[AA]" -> [color, opacity] for SVG paint attributes. */
const paint = (c: string): [string, number] => [c.slice(0, 7), c.length === 9 ? parseInt(c.slice(7), 16) / 255 : 1];

const DASHES: Record<string, number[]> = {
  dot: [1, 1], dash: [4, 3], lgDash: [8, 3], dashDot: [4, 3, 1, 3], lgDashDot: [8, 3, 1, 3], lgDashDotDot: [8, 3, 1, 3, 1, 3],
  sysDash: [3, 1], sysDot: [1, 1], sysDashDot: [3, 1, 1, 1], sysDashDotDot: [3, 1, 1, 1, 1, 1],
};
const OVERLAY: Record<string, [string, number]> = { lighten: ["#FFFFFF", 0.4], lightenLess: ["#FFFFFF", 0.2], darken: ["#000000", 0.4], darkenLess: ["#000000", 0.2] };
const ARROW_SIZE = { sm: 2, med: 3, lg: 5 };
const ARROWS: Record<string, string> = {
  triangle: "M0 0 L10 5 L0 10 Z",
  stealth: "M0 0 L10 5 L0 10 L3.5 5 Z",
  diamond: "M0 5 L5 0 L10 5 L5 10 Z",
  oval: "M0 5 A5 5 0 1 1 10 5 A5 5 0 1 1 0 5 Z",
  arrow: "M0 0 L10 5 L0 10",
};

let uid = 0;

/** Draw a shape's geometry as one SVG in pixel space: fill (solid/gradient), shading, outline, arrowheads. */
function drawGeometry(el: HTMLElement, shape: ShapeModel, scale: number, withFill: boolean): SVGSVGElement | null {
  const line = shape.line;
  const fill = withFill ? shape.fill : null;
  if (!shape.paths?.length || (!fill && !line)) return null;
  const W = shape.rect.w * scale;
  const H = shape.rect.h * scale;
  const svg = svgEl(el, "svg", { class: "pptx-geom", width: Math.max(W, 1), height: Math.max(H, 1) });
  if (shape.flipH || shape.flipV) svg.style.transform = `scale(${shape.flipH ? -1 : 1}, ${shape.flipV ? -1 : 1})`;
  const defs = svgEl(svg, "defs");

  let fillPaint = "none";
  let fillOpacity = 1;
  if (fill?.type === "solid") [fillPaint, fillOpacity] = paint(fill.color);
  else if (fill?.type === "gradient") {
    const id = `dw-grad-${++uid}`;
    let grad: SVGElement;
    if (fill.radial) grad = svgEl(defs, "radialGradient", { id });
    else {
      const a = (fill.angle * Math.PI) / 180;
      const [dx, dy] = [Math.cos(a), Math.sin(a)];
      const half = (Math.abs(W * dx) + Math.abs(H * dy)) / 2;
      grad = svgEl(defs, "linearGradient", {
        id, gradientUnits: "userSpaceOnUse",
        x1: W / 2 - dx * half, y1: H / 2 - dy * half, x2: W / 2 + dx * half, y2: H / 2 + dy * half,
      });
    }
    for (const stop of fill.stops) {
      const [c, o] = paint(stop.color);
      svgEl(grad, "stop", { offset: stop.pos, "stop-color": c, "stop-opacity": o });
    }
    fillPaint = `url(#${id})`;
  }

  const strokeWidth = line ? Math.max(1, line.width * scale) : 0;
  // Office sizes arrowheads from the line width, but counts lines thinner than 0.7 mm (25200 EMU) as that wide.
  const arrowBase = line ? Math.max(line.width, 25200) * scale : 0;
  const marker = (end: "head" | "tail") => {
    const arrow = line?.[end];
    if (!arrow || !line) return null;
    const id = `dw-arrow-${++uid}`;
    const m = svgEl(defs, "marker", {
      id, viewBox: "0 0 10 10", refX: arrow.type === "arrow" || arrow.type === "triangle" || arrow.type === "stealth" ? 9 : 5, refY: 5,
      markerUnits: "userSpaceOnUse", markerWidth: ARROW_SIZE[arrow.len] * arrowBase, markerHeight: ARROW_SIZE[arrow.w] * arrowBase, orient: "auto-start-reverse",
    });
    const [c, o] = paint(line.color);
    const open = arrow.type === "arrow";
    svgEl(m, "path", { d: ARROWS[arrow.type] ?? ARROWS.triangle, fill: open ? "none" : c, stroke: open ? c : "none", "stroke-width": 1.5, opacity: o });
    return `url(#${id})`;
  };
  const head = marker("head");
  const tail = marker("tail");

  // Paths are in shape EMU (geometry size); the shape may have been resized by a group transform.
  for (const path of shape.paths) {
    const d = scalePath(path.d, (W || 1) / (path.w || 1), (H || 1) / (path.h || 1));
    const p = svgEl(svg, "path", { d, fill: "none", stroke: "none" });
    if (path.fill !== "none" && fill) {
      p.setAttribute("fill", fillPaint);
      if (fillOpacity < 1) p.setAttribute("fill-opacity", String(fillOpacity));
      const shade = OVERLAY[path.fill];
      if (shade) svgEl(svg, "path", { d, fill: shade[0], "fill-opacity": shade[1], stroke: "none" });
    }
    if (path.stroke && line) {
      const [c, o] = paint(line.color);
      Object.entries({
        stroke: c, "stroke-opacity": o, "stroke-width": strokeWidth,
        "stroke-linecap": line.cap === "rnd" ? "round" : line.cap === "sq" ? "square" : "butt", "stroke-linejoin": "round",
      }).forEach(([k, v]) => p.setAttribute(k, String(v)));
      const dash = DASHES[line.dash];
      if (dash) p.setAttribute("stroke-dasharray", dash.map((n) => n * strokeWidth).join(" "));
      if (head) p.setAttribute("marker-start", head);
      if (tail) p.setAttribute("marker-end", tail);
    }
  }
  return svg;
}

/** CSS drop-shadow for a shape's outer shadow. */
function shadowCss(shape: ShapeModel, scale: number): string | null {
  const s = shape.shadow;
  if (!s) return null;
  const a = (s.dir * Math.PI) / 180;
  const [c, o] = paint(s.color);
  const rgb = [1, 3, 5].map((i) => parseInt(c.slice(i, i + 2), 16)).join(", ");
  return `drop-shadow(${s.dist * Math.cos(a) * scale}px ${s.dist * Math.sin(a) * scale}px ${(s.blur * scale) / 2}px rgba(${rgb}, ${o}))`;
}

/** The editable text area of a shape or table cell. */
function textBox(parent: HTMLElement, shape: ShapeModel, pxPerPt: number, scale: number, opts: RenderOptions, cls: string): HTMLElement {
  const px = (emu: number) => `${emu * scale}px`;
  const t = shape.text!;
  const textEl = parent.createDiv();
  textEl.className = cls;
  if (shape.textRect) {
    const r = shape.textRect;
    Object.assign(textEl.style, { left: px(r.x), top: px(r.y), width: px(r.w), height: px(r.h) });
  }
  textEl.style.justifyContent = JUSTIFY[t.anchor];
  textEl.style.padding = [t.insets.t, t.insets.r, t.insets.b, t.insets.l].map(px).join(" ");
  renderParagraphs(textEl, t.paras, pxPerPt, scale, t.lnSpcReduction, { slideNumber: opts.slideNumber });
  const hlink = t.style.colors.scheme.hlink;
  if (hlink) textEl.style.setProperty("--pptx-hlink", hlink);
  textEl.classList.toggle("is-nowrap", !t.wrap);
  // "Click to add text" shows in empty placeholders, as in PowerPoint; never in slide number, date or footer ones.
  const phType = shape.placeholder ? (shape.el.getElementsByTagNameNS(shape.el.namespaceURI, "ph")[0]?.getAttribute("type") ?? "body") : null;
  if (phType && !NO_PROMPT.has(phType)) textEl.dataset.prompt = "";
  textEl.classList.toggle("is-empty", t.paras.every((p) => p.runs.length === 0));
  if (opts.editable && !shape.inherited) {
    textEl.contentEditable = "true";
    textEl.spellcheck = true;
  }
  return textEl;
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

  // Table cells are drawn with their table; their text shapes come right after it in the model.
  const cellEls = new Map<Element, HTMLElement>();

  return model.shapes.map((shape, index) => {
    if (shape.kind === "cell") {
      const td = cellEls.get(shape.el) ?? createDiv();
      const pxPerPt = EMU_PER_PT * scale;
      const textEl = shape.text ? textBox(td, shape, pxPerPt, scale, opts, "pptx-text pptx-cell-text") : null;
      if (shape.text) td.style.verticalAlign = { t: "top", ctr: "middle", b: "bottom" }[shape.text.anchor];
      return { shape, el: td, textEl, pxPerPt };
    }
    const el = slide.createDiv();
    el.className = `pptx-shape pptx-kind-${shape.kind}`;
    el.classList.toggle("is-inherited", shape.inherited);
    el.classList.toggle("pptx-diagram-proxy", !!shape.diagram);
    el.classList.toggle("pptx-chart", !!shape.chart);
    el.dataset.index = String(index);
    if (!shape.inherited) el.dataset.spids = shapeIds(shape.el).join(" ");
    if (shape.link) {
      el.dataset.link = shape.link;
      el.addClass("has-link");
    }
    Object.assign(el.style, { left: px(shape.rect.x), top: px(shape.rect.y), width: px(shape.rect.w), height: px(shape.rect.h) });
    if (shape.rot) el.style.transform = `rotate(${shape.rot}deg)`;
    const shadow = shadowCss(shape, scale);
    if (shape.media) {
      // Linked (not embedded) media plays when the view resolved it to a vault file.
      const mediaUrl = opts.playMedia ? opts.imageUrl(shape.media.path) : null;
      if (opts.playMedia && (!shape.media.external || mediaUrl)) {
        const media = shape.media.kind === "video" ? el.createEl("video") : el.createEl("audio");
        media.src = mediaUrl ?? "";
        media.controls = true; media.autoplay = shape.media.autoPlay; media.loop = shape.media.loop; media.preload = "metadata";
        media.addClass("pptx-media-player");
        for (const event of ["click", "pointerdown", "keydown"]) media.addEventListener(event, e => e.stopPropagation());
      } else {
        if (shape.media.kind === "video" && shape.image) { const poster=el.createEl("img");poster.src=opts.imageUrl(shape.image)??"";poster.addClass("pptx-media-poster"); }
        el.createDiv({cls:"pptx-media-label",text: `${shape.media.kind === "audio" ? "♫ Audio" : "▶ Video"} · ${shape.name}`});
      }
    } else if (shape.image) {
      const img = el.createEl("img");
      img.src = opts.imageUrl(shape.image) ?? "";
      img.draggable = false;
      const c = shape.crop;
      if (c) {
        // Show the srcRect part of the image stretched over the frame.
        const pct = (v: number) => `${Math.round(v * 10000) / 100}%`;
        el.classList.add("is-cropped");
        Object.assign(img.style, {
          position: "absolute",
          maxWidth: "none",
          width: pct(1 / (1 - c.l - c.r)),
          height: pct(1 / (1 - c.t - c.b)),
          left: pct(-c.l / (1 - c.l - c.r)),
          top: pct(-c.t / (1 - c.t - c.b)),
        });
      }
      // Pictures cropped to a shape (ellipse, rounded rectangle...).
      if (shape.geom && shape.geom !== "rect" && shape.paths?.[0]) {
        const p = shape.paths[0];
        img.style.clipPath = `path("${scalePath(p.d, (shape.rect.w * scale) / (p.w || 1), (shape.rect.h * scale) / (p.h || 1))}")`;
      }
      const adjust = shape.adjust;
      if (adjust?.transparency) img.style.opacity = String(+(1 - adjust.transparency).toFixed(3));
      const filter = adjust && pictureFilter(adjust);
      const url = filter ? `url(#${svgFilter(el, filter)})` : "";
      if (url || shadow) img.style.filter = [url, shadow].filter(Boolean).join(" ");
      drawGeometry(el, shape, scale, false);
    } else if (shape.paths) {
      const svg = drawGeometry(el, shape, scale, true);
      if (svg && shadow) svg.style.filter = shadow;
    } else {
      const fill = fillCss(shape.fill, shape.flipH, shape.flipV);
      if (fill) el.style.background = fill;
      if (shape.line) el.style.border = `${Math.max(1, shape.line.width * scale)}px solid ${shape.line.color}`;
      if (shadow) el.style.filter = shadow;
    }

    if (shape.table) {
      const t = shape.table;
      const table = el.createEl("table");
      const cols = table.createEl("colgroup");
      for (const w of t.cols) cols.createEl("col").style.width = px(w);
      t.cells.forEach((row, r) => {
        const tr = table.createEl("tr");
        tr.style.height = px(t.rows[r]);
        for (const cell of row) {
          if (cell.hidden) continue;
          const td = tr.createEl("td");
          if (cell.gridSpan > 1) td.colSpan = cell.gridSpan;
          if (cell.rowSpan > 1) td.rowSpan = cell.rowSpan;
          const fill = fillCss(cell.fill);
          if (fill) td.style.background = fill;
          for (const [side, prop] of [["l", "borderLeft"], ["r", "borderRight"], ["t", "borderTop"], ["b", "borderBottom"]] as const) {
            const b = cell.borders[side];
            td.style[prop] = b ? `${Math.max(1, Math.round(b.width * scale))}px solid ${b.color}` : "none";
          }
          cellEls.set(cell.el, td);
        }
      });
    } else if (shape.chart) {
      renderChart(el, shape.chart, { w: shape.rect.w / EMU_PER_PT, h: shape.rect.h / EMU_PER_PT });
    } else if (shape.kind === "frame" && !shape.diagram) {
      el.createDiv().className = "pptx-frame-label";
      el.lastElementChild!.textContent = "Chart or diagram (view only)";
    }

    const pxPerPt = EMU_PER_PT * scale * (shape.text?.fontScale ?? 1);
    const textEl = shape.text ? textBox(el, shape, pxPerPt, scale, opts, "pptx-text") : null;
    return { shape, el, textEl, pxPerPt };
  });
}

let filterId = 0;

/** Add an SVG filter made of `primitives` (markup we build ourselves) inside `parent`; returns its id. */
function svgFilter(parent: HTMLElement, primitives: string): string {
  const id = `dw-pic-filter-${++filterId}`;
  const doc = new DOMParser().parseFromString(
    `<svg xmlns="http://www.w3.org/2000/svg" width="0" height="0" aria-hidden="true"><filter id="${id}" color-interpolation-filters="sRGB">${primitives}</filter></svg>`,
    "image/svg+xml"
  );
  const svg = parent.ownerDocument.importNode(doc.documentElement, true) as unknown as SVGSVGElement;
  svg.setAttribute("class", "pptx-filter-defs");
  parent.appendChild(svg);
  return id;
}
