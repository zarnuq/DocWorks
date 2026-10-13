/*
 * Google Slides' paint format: copy a shape's look (fill, outline, shadow, and its first text run's style and
 * alignment) and apply it to other shapes. Writes only the properties copied; everything else stays.
 */
import { NS, childNS } from "../../core/utils/xml";
import { Fill, Line, Shadow, writeFill, writeLine, writeShadow } from "./shape-style";
import { Align, RunProps, ShapeModel } from "./slide-model";
import { writeParagraphs } from "./slide-edit";

type RunStyle = Pick<RunProps, "b" | "i" | "u" | "strike" | "sz" | "color" | "font">;

export interface ShapeFormat {
  /** Fill, outline and shadow (absent for pictures, tables and other frames). */
  look: { fill: Fill | null; line: Line | null; shadow: Shadow | null } | null;
  text: { run: RunStyle; align: Align } | null;
}

export function copyFormat(shape: ShapeModel): ShapeFormat {
  const look = shape.kind === "sp" ? { fill: shape.fill, line: shape.line, shadow: shape.shadow } : null;
  const para = shape.text?.paras.find((p) => p.runs.some((r) => r.text)) ?? shape.text?.paras[0];
  const r = para?.runs.find((x) => x.text) ?? para?.defaults;
  const text = para && r ? { run: { b: r.b, i: r.i, u: r.u, strike: r.strike, sz: r.sz, color: r.color, font: r.font }, align: para.align } : null;
  return { look, text };
}

export function applyFormat(shape: ShapeModel, f: ShapeFormat): void {
  const spPr = childNS(shape.el, NS.p, "spPr");
  if (f.look && spPr && shape.kind === "sp") {
    // Lines (connectors) have no fill.
    if (shape.el.localName === "sp") writeFill(spPr, f.look.fill);
    writeLine(spPr, f.look.line);
    writeShadow(spPr, f.look.shadow);
  }
  if (f.text && shape.text) {
    const { run, align } = f.text;
    writeParagraphs(shape.text, shape.text.paras.map((p) => ({ ...p, align, runs: p.runs.map((r) => ({ ...r, ...run })) })));
  }
}
