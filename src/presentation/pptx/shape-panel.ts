/*
 * Format panel for the selected shape: fill, outline, shadow, position, size,
 * rotation, flips and stacking order. Plain DOM; every change goes through
 * `ctl.edit` so the view can snapshot it for undo and redraw.
 */
import { ACCENTS, NO_ADJUST, PictureAdjust, setPictureAdjust } from "../../domain/pptx/picture-adjust";
import type { ShapeModel } from "../../domain/pptx/slide-model";
import { Fill, Line, writeFill, writeLine, writeShadow } from "../../domain/pptx/shape-style";
import { reorder, setCrop, setRect, setTransform, writeBodyPr, ZOrder } from "../../domain/pptx/slide-edit";
import type { Alignment } from "../../domain/pptx/arrange";
import { NS } from "../../core/utils/xml";
import {
  BorderMode, CellRange, deleteCol, deleteRow, insertCol, insertRow, mergeCells, setCellBorders, setCellFill, setColWidth, setRowHeight, setTableFlags, splitCell, TableFlags,
} from "../../domain/pptx/table";

export interface PanelControl {
  /** Apply a change to the shape's XML as one undoable step, then redraw. */
  edit(fn: (shape: Element) => void): void;
  lockAspect: boolean;
  /** Actions on the whole selection; when present, the panel shows align buttons. */
  arrange?: ArrangeControl;
  /** Tables: the selected cell range (the caret's cell, or a Shift+click range). */
  cells?: CellRange | null;
  tableStyles?: { id: string; name: string }[];
  /** Tables: switch style, writing its definition into the presentation. */
  setTableStyle?(id: string): void;
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

function title(host: HTMLElement, text: string, columns = 2): HTMLElement {
  return sectionIn(host, text).createDiv({ cls: columns === 3 ? "pptx-props-buttons is-grid is-thirds" : "pptx-props-buttons is-grid" });
}

/** Align, distribute and group buttons for the current selection. */
function arrangeButtons(host: HTMLElement, a: ArrangeControl, withOrder: boolean): void {
  const btn = (parent: HTMLElement, text: string, fn: () => void) => parent.createEl("button", { text }).addEventListener("click", fn);
  if (withOrder) {
    const order = title(host, "Arrange");
    for (const [text, to] of ORDER) btn(order, text, () => a.order(to));
  }
  const align = title(host, a.count === 1 ? "Align to slide" : "Align", 3);
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
export const EMU_PER_PT = 12700;
const DEFAULT_COLOR = "#4472C4";
export const DASHES: [string, string][] = [
  ["Solid", "solid"], ["Round dot", "sysDot"], ["Square dot", "sysDash"], ["Dash", "dash"], ["Dash dot", "dashDot"],
  ["Long dash", "lgDash"], ["Long dash dot", "lgDashDot"], ["Long dash dot dot", "lgDashDotDot"],
];
const ARROWS: [string, string][] = [["None", ""], ["Triangle", "triangle"], ["Stealth", "stealth"], ["Open", "arrow"], ["Diamond", "diamond"], ["Oval", "oval"]];

export const spPrOf = (el: Element) => Array.from(el.children).find((c) => c.localName === "spPr" && c.namespaceURI === NS.p)!;
const base = (c: string) => c.slice(0, 7).toLowerCase();
const alphaOf = (c: string) => (c.length === 9 ? parseInt(c.slice(7), 16) / 255 : 1);
const withAlpha = (c: string, a: number) =>
  (a >= 1 ? c.slice(0, 7) : c.slice(0, 7) + Math.round(a * 255).toString(16).padStart(2, "0")).toUpperCase();
const round2 = (n: number) => String(Math.round(n * 100) / 100);

/** Sections open until closed (and closed ones stay closed) while the panel is rebuilt after each edit. */
const openSections = new Set(["Fill", "Line", "Crop", "Adjustments", "Table", "Cell fill", "Table fill", "Borders"]);

/** A collapsible section, like Google Slides' format options. */
function sectionIn(host: HTMLElement, title: string): HTMLElement {
  const details = host.createEl("details", { cls: "pptx-props-group" });
  details.open = openSections.has(title);
  details.createEl("summary", { cls: "pptx-props-title", text: title });
  details.addEventListener("toggle", () => (details.open ? openSections.add(title) : openSections.delete(title)));
  return details.createDiv({ cls: "pptx-props-section" });
}

/** Two short fields side by side (X / Y, width / height, margins). */
const pair = (parent: HTMLElement) => parent.createDiv({ cls: "pptx-props-pair" });

function row(parent: HTMLElement, label: string, text = label): HTMLElement {
  const r = parent.createEl("label", { cls: "pptx-props-row" });
  r.createSpan({ text });
  return r;
}

function select(parent: HTMLElement, label: string, options: [string, string][], value: string, onChange: (v: string) => void): HTMLSelectElement {
  const s = row(parent, label).createEl("select", { cls: "dropdown", attr: { "data-field": label } });
  for (const [text, v] of options) s.createEl("option", { text, value: v });
  s.value = value;
  s.addEventListener("change", () => onChange(s.value));
  return s;
}

function input(parent: HTMLElement, label: string, type: string, value: string, onChange: (v: string) => void, attrs: Record<string, string> = {}, text = label): HTMLInputElement {
  const r = row(parent, label, text);
  const i = r.createEl("input", { type, value, attr: { "data-field": label, ...attrs } });
  i.addEventListener("change", () => onChange(i.value));
  if (type === "color") eyedropper(r, i);
  return i;
}

interface EyeDropperApi {
  open(): Promise<{ sRGBHex: string }>;
}

/** A pipette next to a color field, picking any pixel on screen (Chromium's EyeDropper; hidden without it). */
function eyedropper(parent: HTMLElement, field: HTMLInputElement): void {
  const Api = (window as unknown as { EyeDropper?: new () => EyeDropperApi }).EyeDropper;
  if (!Api) return;
  const b = parent.createEl("button", { cls: "pptx-eyedropper clickable-icon", attr: { type: "button", "aria-label": "Pick color from screen" } });
  // Lucide's "pipette" (ISC license), drawn inline: this module stays free of Obsidian imports.
  const svg = new DOMParser().parseFromString(
    `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="m2 22 1-1h3l9-9"/><path d="M3 21v-3l9-9"/><path d="m15 6 3.4-3.4a2.1 2.1 0 1 1 3 3L18 9l.4.4a2.1 2.1 0 1 1-3 3l-3.8-3.8a2.1 2.1 0 1 1 3-3l.4.4Z"/></svg>`,
    "image/svg+xml"
  ).documentElement;
  b.appendChild(document.importNode(svg, true));
  b.addEventListener("click", (e) => {
    e.preventDefault();
    new Api().open().then(({ sRGBHex }) => {
      field.value = sRGBHex.toLowerCase();
      field.dispatchEvent(new Event("change"));
    }, () => undefined); // Esc cancels
  });
}

/** A number field; `text` is a shorter visible label (the field keeps `label` as its name). */
const number = (parent: HTMLElement, label: string, value: number, onChange: (v: number) => void, step = "0.01", text = label) =>
  input(parent, label, "number", round2(value), (v) => Number.isFinite(Number(v)) && v !== "" && onChange(Number(v)), { step }, text);

/** Fill type, colors, transparency and gradient angle; `none` is the label/value for "no own fill". */
function fillControls(s: HTMLElement, fill: Fill | null, setFill: (f: Fill | null) => void, none: [string, string][], label = "Type", plain = DEFAULT_COLOR): void {
  const type = !fill ? none[0][1] : fill.type === "solid" ? "solid" : fill.radial ? "radial" : "linear";
  const first = !fill ? plain : fill.type === "solid" ? fill.color : fill.stops[0].color;
  select(s, label, [...none, ["Solid", "solid"], ["Gradient", "linear"], ["Radial gradient", "radial"]], type, (v) => {
    if (v === none[0][1]) return setFill(null);
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

const TABLE_FLAGS: [string, keyof TableFlags][] = [
  ["Header row", "firstRow"], ["Total row", "lastRow"], ["Banded rows", "bandRow"],
  ["First column", "firstCol"], ["Last column", "lastCol"], ["Banded columns", "bandCol"],
];

/** Row / column / merge actions for cells `range` of a table, shared by the format panel and the right-click menu. */
export function tableActions(t: NonNullable<ShapeModel["table"]>, range: CellRange): [string, (tbl: Element) => void][] {
  const [r1, c1, r2, c2] = range;
  const out: [string, (tbl: Element) => void][] = [
    ["Insert row above", (x) => insertRow(x, r1, false)],
    ["Insert row below", (x) => insertRow(x, r2, true)],
    ["Insert column left", (x) => insertCol(x, c1, false)],
    ["Insert column right", (x) => insertCol(x, c2, true)],
  ];
  if (t.rows.length > 1) out.push(["Delete row", (x) => { for (let r = r2; r >= r1; r--) deleteRow(x, r); }]);
  if (t.cols.length > 1) out.push(["Delete column", (x) => { for (let c = c2; c >= c1; c--) deleteCol(x, c); }]);
  if (r1 !== r2 || c1 !== c2) out.push(["Merge cells", (x) => mergeCells(x, [r1, c1, r2, c2])]);
  const here = t.cells[r1]?.[c1];
  if (here && r1 === r2 && c1 === c2 && (here.gridSpan > 1 || here.rowSpan > 1)) out.push(["Split cell", (x) => splitCell(x, r1, c1)]);
  return out;
}

/** Rows and columns, merging, cell fill and borders, and style options for a table. */
function tableControls(section: (title: string) => HTMLElement, shape: ShapeModel, ctl: PanelControl): void {
  const t = shape.table!;
  const tbl = (fn: (tbl: Element) => void) => ctl.edit((frame) => fn(frame.getElementsByTagNameNS(NS.a, "tbl")[0]));
  const all: CellRange = [0, 0, t.rows.length - 1, t.cols.length - 1];
  const range = ctl.cells ?? null;
  const [r1, c1, r2, c2] = range ?? all;

  const s = section("Table");
  const buttons = s.createDiv({ cls: "pptx-props-buttons" });
  const btn = (text: string, fn: (tbl: Element) => void) => buttons.createEl("button", { text }).addEventListener("click", () => tbl(fn));
  if (range) {
    for (const [text, fn] of tableActions(t, range)) btn(text, fn);
    number(s, "Column width (in)", t.cols[c1] / EMU_PER_INCH, (v) => tbl((x) => setColWidth(x, c1, Math.max(0.1, v) * EMU_PER_INCH)));
    number(s, "Row height (in)", t.rows[r1] / EMU_PER_INCH, (v) => tbl((x) => setRowHeight(x, r1, Math.max(0.1, v) * EMU_PER_INCH)));
  } else {
    s.createDiv({ cls: "pptx-props-hint", text: "Click a cell (Shift+click for a range) to add rows, columns or merge." });
  }

  const cells: [number, number][] = [];
  for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) cells.push([r, c]);
  const first = t.cells[r1]?.[c1];
  const ownFill = Array.from(kidOf(first?.el, "tcPr")?.children ?? []).some((c) => c.localName === "solidFill" || c.localName === "gradFill");
  const own = ownFill ? (first?.fill ?? null) : null;
  const fills = section(range ? "Cell fill" : "Table fill");
  fillControls(fills, own, (f) => tbl((x) => setCellFill(x, cells, f)), [["Table style", "style"]], "Cell fill", "#FFFFFF");

  const b = section("Borders");
  let color = "#000000";
  let width = 1;
  input(b, "Border color", "color", color, (v) => (color = v.toUpperCase()));
  number(b, "Border width (pt)", width, (v) => (width = Math.max(0.25, v)), "0.25");
  select(b, "Borders", [["Apply…", ""], ["All borders", "all"], ["Outer borders", "outer"], ["Inside borders", "inside"], ["No borders", "none"]], "", (v) => {
    if (!v) return;
    tbl((x) => setCellBorders(x, [r1, c1, r2, c2], v === "none" ? "all" : (v as BorderMode), v === "none" ? null : { color, width: width * EMU_PER_PT }));
  });

  const st = section("Table style");
  const styles = ctl.tableStyles ?? [];
  const options: [string, string][] = styles.map((x) => [x.name, x.id]);
  if (t.styleId && !styles.some((x) => x.id === t.styleId)) options.unshift(["Current style", t.styleId]);
  if (!t.styleId) options.unshift(["No style", ""]);
  select(st, "Table style", options, t.styleId ?? "", (id) => id && ctl.setTableStyle?.(id));
  for (const [label, key] of TABLE_FLAGS) {
    const box = row(st, label).createEl("input", { type: "checkbox", attr: { "data-field": label } });
    box.checked = t.flags[key];
    box.addEventListener("change", () => tbl((x) => setTableFlags(x, { [key]: box.checked })));
  }
}

const kidOf = (el: Element | undefined, local: string) => Array.from(el?.children ?? []).find((c) => c.localName === local) ?? null;

export interface SlideControl {
  layouts: { name: string; path: string }[];
  layout: string;
  setLayout(path: string): void;
  /** The slide's own background, null when it comes from the layout. */
  background: Fill | null;
  setBackground(fill: Fill | null): void;
  applyBackgroundToAll(): void;
  size: { cx: number; cy: number };
  setSize(cx: number, cy: number): void;
}

const SIZES: [string, number, number][] = [
  ["Widescreen 16:9", 12192000, 6858000],
  ["Standard 4:3", 9144000, 6858000],
  ["16:10", 10972800, 6858000],
  ["A4", 10691813, 7559675],
];

/** Panel shown when nothing is selected: layout, background and slide size. */
export function buildSlidePanel(host: HTMLElement, ctl: SlideControl): void {
  host.replaceChildren();
  host.createDiv({ cls: "pptx-props-heading", text: "Slide" });
  // Each part is a single labelled row, so no section headings repeating the labels.
  const s = host.createDiv({ cls: "pptx-props-section" });
  select(s, "Layout", ctl.layouts.map((l) => [l.name || "Untitled layout", l.path]), ctl.layout, (p) => ctl.setLayout(p));

  const bg = host.createDiv({ cls: "pptx-props-section" });
  fillControls(bg, ctl.background, (f) => ctl.setBackground(f), [["From layout", "layout"]], "Background", "#FFFFFF");
  bg.createEl("button", { text: "Apply to all slides" }).addEventListener("click", () => ctl.applyBackgroundToAll());

  const size = host.createDiv({ cls: "pptx-props-section" });
  const current = `${ctl.size.cx}x${ctl.size.cy}`;
  const options: [string, string][] = SIZES.map(([name, cx, cy]) => [name, `${cx}x${cy}`]);
  if (!options.some(([, v]) => v === current)) {
    options.push([`Custom (${round2(ctl.size.cx / EMU_PER_INCH)} × ${round2(ctl.size.cy / EMU_PER_INCH)} in)`, current]);
  }
  select(size, "Slide size", options, current, (v) => {
    const [cx, cy] = v.split("x").map(Number);
    ctl.setSize(cx, cy);
  });
}

export function buildShapePanel(host: HTMLElement, shape: ShapeModel, ctl: PanelControl): void {
  host.replaceChildren();
  const edit = (fn: (el: Element, spPr: Element) => void) => ctl.edit((el) => fn(el, spPrOf(el)));

  const section = (title: string) => sectionIn(host, title);

  // --- Fill ------------------------------------------------------------------
  const hasFill = shape.kind === "sp" && shape.el.localName === "sp";
  if (hasFill) {
    const fill = shape.fill;
    const s = section("Fill");
    fillControls(s, fill, (f) => edit((_, spPr) => writeFill(spPr, f)), [["No fill", "none"]]);
  }

  // --- Crop ------------------------------------------------------------------
  if (shape.kind === "pic" && !shape.media) {
    const s = section("Crop");
    const crop = shape.crop ?? { l: 0, t: 0, r: 0, b: 0 };
    for (const [label, key] of [["Crop left %", "l"], ["Crop top %", "t"], ["Crop right %", "r"], ["Crop bottom %", "b"]] as const) {
      number(s, label, crop[key] * 100, (v) => edit((el) => setCrop(el, { ...crop, [key]: v / 100 })), "1");
    }
    s.createEl("button", { text: "Reset crop" }).addEventListener("click", () => edit((el) => setCrop(el, { l: 0, t: 0, r: 0, b: 0 })));
  }

  // --- Adjustments (pictures) ---------------------------------------------------
  if (shape.kind === "pic" && !shape.media && shape.adjust) {
    const s = section("Adjustments");
    const adj = shape.adjust;
    const set = (patch: Partial<PictureAdjust>) => edit((el) => setPictureAdjust(el, { ...adj, ...patch }));
    number(s, "Picture transparency %", adj.transparency * 100, (v) => set({ transparency: Math.max(0, Math.min(100, v)) / 100 }), "1", "Transparency %");
    number(s, "Brightness %", adj.brightness * 100, (v) => set({ brightness: Math.max(-100, Math.min(100, v)) / 100 }), "1");
    number(s, "Contrast %", adj.contrast * 100, (v) => set({ contrast: Math.max(-100, Math.min(100, v)) / 100 }), "1");
    const options: [string, string][] = [["No recolor", "none"], ["Grayscale", "gray"]];
    ACCENTS.forEach((a, i) => options.push([`Accent ${i + 1} dark`, `dark:${a}`], [`Accent ${i + 1} light`, `light:${a}`]));
    if (adj.recolor === "duotone") options.push(["As saved", "duotone"]);
    select(s, "Recolor", options, adj.recolor, (v) => set({ recolor: v }));
    s.createEl("button", { text: "Reset adjustments" }).addEventListener("click", () => edit((el) => setPictureAdjust(el, NO_ADJUST)));
  }

  // --- Table -----------------------------------------------------------------
  if (shape.table) tableControls(section, shape, ctl);

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
    s.createDiv({ cls: "pptx-props-sub", text: "Margins (in)" });
    const margins = pair(s);
    for (const [label, key, short] of [["Left margin (in)", "l", "Left"], ["Right margin (in)", "r", "Right"], ["Top margin (in)", "t", "Top"], ["Bottom margin (in)", "b", "Bottom"]] as const) {
      number(margins, label, text.insets[key] / EMU_PER_INCH, (v) => body({ insets: { ...text.insets, [key]: Math.max(0, v) * EMU_PER_INCH } }), "0.01", short);
    }
  }

  // --- Position and size -----------------------------------------------------
  const s = section("Position and size");
  const r = shape.rect;
  const setR = (patch: Partial<typeof r>) => edit((el) => setRect(el, { ...r, ...patch }));
  s.createDiv({ cls: "pptx-props-sub", text: "Inches" });
  const box = pair(s);
  number(box, "X (in)", r.x / EMU_PER_INCH, (v) => setR({ x: v * EMU_PER_INCH }), "0.01", "X");
  number(box, "Y (in)", r.y / EMU_PER_INCH, (v) => setR({ y: v * EMU_PER_INCH }), "0.01", "Y");
  number(box, "Width (in)", r.w / EMU_PER_INCH, (v) => {
    const w = Math.max(0, v * EMU_PER_INCH);
    setR(ctl.lockAspect && r.w ? { w, h: (r.h * w) / r.w } : { w });
  }, "0.01", "W");
  number(box, "Height (in)", r.h / EMU_PER_INCH, (v) => {
    const h = Math.max(0, v * EMU_PER_INCH);
    setR(ctl.lockAspect && r.h ? { h, w: (r.w * h) / r.h } : { h });
  }, "0.01", "H");
  const lock = row(s, "Lock aspect ratio").createEl("input", { type: "checkbox", attr: { "data-field": "Lock aspect ratio" } });
  lock.checked = ctl.lockAspect;
  lock.addEventListener("change", () => (ctl.lockAspect = lock.checked));
  if (shape.el.localName !== "graphicFrame") {
    number(s, "Rotation (°)", shape.rot, (v) => edit((el) => setTransform(el, { rot: v })), "1");
    const flips = s.createDiv({ cls: "pptx-props-buttons is-grid" });
    const flipBtn = (text: string, fn: () => void) => flips.createEl("button", { text }).addEventListener("click", fn);
    flipBtn("Flip horizontal", () => edit((el) => setTransform(el, { flipH: !shape.flipH })));
    flipBtn("Flip vertical", () => edit((el) => setTransform(el, { flipV: !shape.flipV })));
  }

  // --- Arrange ---------------------------------------------------------------
  const order = title(host, "Arrange");
  for (const [label, to] of ORDER) {
    order.createEl("button", { text: label }).addEventListener("click", () => edit((el) => reorder(el, to)));
  }
  if (ctl.arrange) arrangeButtons(host, ctl.arrange, false);
}
