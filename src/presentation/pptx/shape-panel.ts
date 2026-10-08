/*
 * Format panel for the selected shape: fill, outline, shadow, position, size,
 * rotation, flips and stacking order. Plain DOM; every change goes through
 * `ctl.edit` so the view can snapshot it for undo and redraw.
 */
import type { ShapeModel } from "../../domain/pptx/slide-model";
import { Fill, Line, writeFill, writeLine, writeShadow } from "../../domain/pptx/shape-style";
import { reorder, setRect, setTransform, writeBodyPr, ZOrder } from "../../domain/pptx/slide-edit";
import type { Alignment } from "../../domain/pptx/arrange";
import { NS } from "../../core/utils/xml";

export interface PanelControl {
  /** Apply a change to the shape's XML as one undoable step, then redraw. */
  edit(fn: (shape: Element) => void): void;
  lockAspect: boolean;
  /** Actions on the whole selection; when present, the panel shows align buttons. */
  arrange?: ArrangeControl;
}

export interface ArrangeControl {
  /** Number of selected items, and how many of them are groups. */
  count: number;
  groups: number;
  align(mode: Alignment): void;
  distribute(axis: "horizontal" | "vertical"): void;
  group(): void;
  ungroup(): void;
  order(to: ZOrder): void;
}

const ALIGN: [string, Alignment][] = [["Left", "left"], ["Center", "center"], ["Right", "right"], ["Top", "top"], ["Middle", "middle"], ["Bottom", "bottom"]];
const ORDER: [string, ZOrder][] = [["Bring to front", "front"], ["Bring forward", "forward"], ["Send backward", "backward"], ["Send to back", "back"]];

function title(host: HTMLElement, text: string): HTMLElement {
  host.createDiv({ cls: "pptx-props-title", text });
  return host.createDiv({ cls: "pptx-props-buttons" });
}

/** Align, distribute and group buttons for the current selection. */
function arrangeButtons(host: HTMLElement, a: ArrangeControl, withOrder: boolean): void {
  const btn = (parent: HTMLElement, text: string, fn: () => void) => parent.createEl("button", { text }).addEventListener("click", fn);
  if (withOrder) {
    const order = title(host, "Arrange");
    for (const [text, to] of ORDER) btn(order, text, () => a.order(to));
  }
  const align = title(host, a.count === 1 ? "Align to slide" : "Align");
  for (const [text, mode] of ALIGN) btn(align, text, () => a.align(mode));
  if (a.count >= 3) {
    const dist = title(host, "Distribute");
    btn(dist, "Horizontally", () => a.distribute("horizontal"));
    btn(dist, "Vertically", () => a.distribute("vertical"));
  }
  if (a.count >= 2 || a.groups) {
    const group = title(host, "Group");
    if (a.count >= 2) btn(group, "Group", () => a.group());
    if (a.groups) btn(group, "Ungroup", () => a.ungroup());
  }
}

/** Panel for several selected items or a group. */
export function buildArrangePanel(host: HTMLElement, a: ArrangeControl): void {
  host.replaceChildren();
  host.createDiv({ cls: "pptx-props-heading", text: a.count === 1 ? "Group" : `${a.count} items selected` });
  arrangeButtons(host, a, true);
}

const EMU_PER_INCH = 914400;
const EMU_PER_PT = 12700;
const DEFAULT_COLOR = "#4472C4";
const DASHES: [string, string][] = [
  ["Solid", "solid"], ["Round dot", "sysDot"], ["Square dot", "sysDash"], ["Dash", "dash"], ["Dash dot", "dashDot"],
  ["Long dash", "lgDash"], ["Long dash dot", "lgDashDot"], ["Long dash dot dot", "lgDashDotDot"],
];
const ARROWS: [string, string][] = [["None", ""], ["Triangle", "triangle"], ["Stealth", "stealth"], ["Open", "arrow"], ["Diamond", "diamond"], ["Oval", "oval"]];

const spPrOf = (el: Element) => Array.from(el.children).find((c) => c.localName === "spPr" && c.namespaceURI === NS.p)!;
const base = (c: string) => c.slice(0, 7).toLowerCase();
const alphaOf = (c: string) => (c.length === 9 ? parseInt(c.slice(7), 16) / 255 : 1);
const withAlpha = (c: string, a: number) =>
  (a >= 1 ? c.slice(0, 7) : c.slice(0, 7) + Math.round(a * 255).toString(16).padStart(2, "0")).toUpperCase();
const round2 = (n: number) => String(Math.round(n * 100) / 100);

export function buildShapePanel(host: HTMLElement, shape: ShapeModel, ctl: PanelControl): void {
  host.replaceChildren();
  const edit = (fn: (el: Element, spPr: Element) => void) => ctl.edit((el) => fn(el, spPrOf(el)));

  const section = (title: string) => {
    host.createDiv({ cls: "pptx-props-title", text: title });
    return host.createDiv({ cls: "pptx-props-section" });
  };
  const row = (parent: HTMLElement, label: string) => {
    const r = parent.createEl("label", { cls: "pptx-props-row" });
    r.createSpan({ text: label });
    return r;
  };
  const select = (parent: HTMLElement, label: string, options: [string, string][], value: string, onChange: (v: string) => void) => {
    const s = row(parent, label).createEl("select", { cls: "dropdown", attr: { "data-field": label } });
    for (const [text, v] of options) s.createEl("option", { text, value: v });
    s.value = value;
    s.addEventListener("change", () => onChange(s.value));
    return s;
  };
  const input = (parent: HTMLElement, label: string, type: string, value: string, onChange: (v: string) => void, attrs: Record<string, string> = {}) => {
    const i = row(parent, label).createEl("input", { type, value, attr: { "data-field": label, ...attrs } });
    i.addEventListener("change", () => onChange(i.value));
    return i;
  };
  const number = (parent: HTMLElement, label: string, value: number, onChange: (v: number) => void, step = "0.01") =>
    input(parent, label, "number", round2(value), (v) => Number.isFinite(Number(v)) && v !== "" && onChange(Number(v)), { step });

  // --- Fill ------------------------------------------------------------------
  const hasFill = shape.kind === "sp" && shape.el.localName === "sp";
  if (hasFill) {
    const fill = shape.fill;
    const s = section("Fill");
    const type = !fill ? "none" : fill.type === "solid" ? "solid" : fill.radial ? "radial" : "linear";
    const first = !fill ? DEFAULT_COLOR : fill.type === "solid" ? fill.color : fill.stops[0].color;
    const setFill = (f: Fill | null) => edit((_, spPr) => writeFill(spPr, f));
    select(s, "Type", [["No fill", "none"], ["Solid", "solid"], ["Gradient", "linear"], ["Radial gradient", "radial"]], type, (v) => {
      if (v === "none") return setFill(null);
      if (v === "solid") return setFill({ type: "solid", color: first });
      const stops = fill?.type === "gradient" ? fill.stops : [{ pos: 0, color: first }, { pos: 1, color: "#FFFFFF" }];
      setFill({ type: "gradient", radial: v === "radial", angle: fill?.type === "gradient" ? fill.angle : 90, stops });
    });
    if (fill?.type === "solid") {
      input(s, "Color", "color", base(fill.color), (v) => setFill({ type: "solid", color: withAlpha(v, alphaOf(fill.color)) }));
      number(s, "Transparency %", (1 - alphaOf(fill.color)) * 100, (t) => setFill({ type: "solid", color: withAlpha(fill.color, 1 - Math.min(100, Math.max(0, t)) / 100) }), "1");
    } else if (fill?.type === "gradient") {
      const stop = (i: number, c: string) => fill.stops.map((st, k) => (k === i ? { ...st, color: c } : st));
      const last = fill.stops.length - 1;
      input(s, "Start color", "color", base(fill.stops[0].color), (v) => setFill({ ...fill, stops: stop(0, withAlpha(v, alphaOf(fill.stops[0].color))) }));
      input(s, "End color", "color", base(fill.stops[last].color), (v) => setFill({ ...fill, stops: stop(last, withAlpha(v, alphaOf(fill.stops[last].color))) }));
      if (!fill.radial) number(s, "Angle", fill.angle, (a) => setFill({ ...fill, angle: ((a % 360) + 360) % 360 }), "1");
    }
  }

  // --- Line ------------------------------------------------------------------
  if (shape.el.localName !== "graphicFrame") {
    const line = shape.line;
    const s = section("Line");
    const setLine = (patch: Partial<Line> | null) => edit((_, spPr) => writeLine(spPr, patch));
    select(s, "Type", [["No line", "none"], ["Solid", "solid"]], line ? "solid" : "none", (v) =>
      setLine(v === "none" ? null : { color: DEFAULT_COLOR, width: line?.width ?? EMU_PER_PT })
    );
    if (line) {
      input(s, "Color", "color", base(line.color), (v) => setLine({ color: v.toUpperCase() }));
      number(s, "Width (pt)", line.width / EMU_PER_PT, (w) => setLine({ width: Math.max(0, w) * EMU_PER_PT }), "0.25");
      select(s, "Dash", DASHES, line.dash, (dash) => setLine({ dash }));
      const arrow = (v: string) => (v ? { type: v, w: "med" as const, len: "med" as const } : null);
      select(s, "Begin arrow", ARROWS, line.head?.type ?? "", (v) => setLine({ head: arrow(v) }));
      select(s, "End arrow", ARROWS, line.tail?.type ?? "", (v) => setLine({ tail: arrow(v) }));
    }
  }

  // --- Shadow ----------------------------------------------------------------
  if (shape.el.localName !== "graphicFrame") {
    const s = section("Effects");
    const box = row(s, "Shadow").createEl("input", { type: "checkbox", attr: { "data-field": "Shadow" } });
    box.checked = !!shape.shadow;
    box.addEventListener("change", () =>
      edit((_, spPr) => writeShadow(spPr, box.checked ? { color: "#00000066", blur: 50800, dist: 38100, dir: 45 } : null))
    );
  }

  // --- Text box --------------------------------------------------------------
  const text = shape.text;
  if (text && !shape.inherited) {
    const s = section("Text box");
    const body = (patch: Parameters<typeof writeBodyPr>[1]) => edit(() => writeBodyPr(text.el, patch));
    select(s, "Vertical alignment", [["Top", "t"], ["Middle", "ctr"], ["Bottom", "b"]], text.anchor, (v) => body({ anchor: v as "t" | "ctr" | "b" }));
    select(s, "Autofit", [["Do not autofit", "none"], ["Shrink text on overflow", "shrink"], ["Resize shape to fit text", "resize"]], text.autofit, (v) =>
      body({ autofit: { type: v as "none" | "shrink" | "resize" } })
    );
    const wrap = row(s, "Wrap text").createEl("input", { type: "checkbox", attr: { "data-field": "Wrap text" } });
    wrap.checked = text.wrap;
    wrap.addEventListener("change", () => body({ wrap: wrap.checked }));
    for (const [label, key] of [["Left margin (in)", "l"], ["Right margin (in)", "r"], ["Top margin (in)", "t"], ["Bottom margin (in)", "b"]] as const) {
      number(s, label, text.insets[key] / EMU_PER_INCH, (v) => body({ insets: { ...text.insets, [key]: Math.max(0, v) * EMU_PER_INCH } }));
    }
  }

  // --- Position and size -----------------------------------------------------
  const s = section("Position and size");
  const r = shape.rect;
  const setR = (patch: Partial<typeof r>) => edit((el) => setRect(el, { ...r, ...patch }));
  number(s, "X (in)", r.x / EMU_PER_INCH, (v) => setR({ x: v * EMU_PER_INCH }));
  number(s, "Y (in)", r.y / EMU_PER_INCH, (v) => setR({ y: v * EMU_PER_INCH }));
  number(s, "Width (in)", r.w / EMU_PER_INCH, (v) => {
    const w = Math.max(0, v * EMU_PER_INCH);
    setR(ctl.lockAspect && r.w ? { w, h: (r.h * w) / r.w } : { w });
  });
  number(s, "Height (in)", r.h / EMU_PER_INCH, (v) => {
    const h = Math.max(0, v * EMU_PER_INCH);
    setR(ctl.lockAspect && r.h ? { h, w: (r.w * h) / r.h } : { h });
  });
  const lock = row(s, "Lock aspect ratio").createEl("input", { type: "checkbox", attr: { "data-field": "Lock aspect ratio" } });
  lock.checked = ctl.lockAspect;
  lock.addEventListener("change", () => (ctl.lockAspect = lock.checked));
  if (shape.el.localName !== "graphicFrame") {
    number(s, "Rotation (°)", shape.rot, (v) => edit((el) => setTransform(el, { rot: v })), "1");
    const flips = s.createDiv({ cls: "pptx-props-buttons" });
    const flipBtn = (text: string, fn: () => void) => flips.createEl("button", { text }).addEventListener("click", fn);
    flipBtn("Flip horizontal", () => edit((el) => setTransform(el, { flipH: !shape.flipH })));
    flipBtn("Flip vertical", () => edit((el) => setTransform(el, { flipV: !shape.flipV })));
  }

  // --- Arrange ---------------------------------------------------------------
  const order = section("Arrange").createDiv({ cls: "pptx-props-buttons" });
  for (const [label, to] of ORDER) {
    order.createEl("button", { text: label }).addEventListener("click", () => edit((el) => reorder(el, to)));
  }
  if (ctl.arrange) arrangeButtons(host, ctl.arrange, false);
}
