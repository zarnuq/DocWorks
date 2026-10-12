/*
 * Cells are painted on a canvas per pane, like Google Sheets does: hundreds of positioned elements per frame cost
 * more in style and layout (against Obsidian's large stylesheet) than drawing them. The grid builds a list of
 * paint operations (plain data, which tests read), and `paint` draws them.
 */
import type { Sparkline } from "../../domain/xlsx/sparklines";

export type Op =
  | { kind: "rect"; x: number; y: number; w: number; h: number; color: string }
  | { kind: "line"; x1: number; y1: number; x2: number; y2: number; color: string; width: number; dash?: number[] }
  | {
      kind: "text";
      /** Clip box (the cell, or the run of empty cells long text spills into). */
      x: number;
      y: number;
      w: number;
      h: number;
      text: string;
      font: string;
      /** Font size in px, for line height. */
      size: number;
      color: string;
      align: "left" | "center" | "right";
      valign: "top" | "center" | "bottom";
      wrap: boolean;
      underline: boolean;
      strike: boolean;
      /** Padding from the left / right edge (indent). */
      pad: number;
    }
  | { kind: "corner"; x: number; y: number; color: string }
  /** A sparkline in the cell box. */
  | { kind: "spark"; x: number; y: number; w: number; h: number; s: Sparkline };

const PAD = 3;

/** A sparkline: a line through the values, columns, or win/loss blocks (up for positive, down for negative). */
function drawSpark(ctx: CanvasRenderingContext2D, op: Extract<Op, { kind: "spark" }>): void {
  const { s } = op;
  const values = s.rtl ? [...s.values].reverse() : s.values;
  const nums = values.filter((v): v is number => v !== null);
  if (!nums.length) return;
  const x0 = op.x + PAD, y0 = op.y + PAD, w = op.w - 2 * PAD, h = op.h - 2 * PAD;
  const hi = Math.max(...nums), lo = Math.min(...nums);
  // Columns grow from zero; a line spans its own range.
  const min = s.min ?? (s.type === "column" ? Math.min(0, lo) : lo), max = s.max ?? (s.type === "column" ? Math.max(0, hi) : hi);
  const y = (v: number) => (max === min ? y0 + h / 2 : y0 + h - ((Math.min(max, Math.max(min, v)) - min) / (max - min)) * h);
  const firstAt = values.findIndex((v) => v !== null), lastAt = values.length - 1 - [...values].reverse().findIndex((v) => v !== null);
  const pointColor = (v: number, i: number, base: string | undefined) =>
    (v === hi && s.high) || (v === lo && s.low) || (i === firstAt && s.first) || (i === lastAt && s.last) || (v < 0 && s.negative) || base;
  ctx.save();
  ctx.beginPath();
  ctx.rect(op.x, op.y, op.w, op.h);
  ctx.clip();
  const n = values.length;
  if (s.type === "line") {
    const x = (i: number) => (n === 1 ? x0 + w / 2 : x0 + (i * w) / (n - 1));
    ctx.strokeStyle = s.color;
    ctx.lineWidth = s.weight;
    ctx.lineJoin = "round";
    ctx.beginPath();
    let pen = false;
    values.forEach((v, i) => {
      if (v === null) return void (pen = false); // blanks leave a gap
      if (pen) ctx.lineTo(x(i), y(v));
      else ctx.moveTo(x(i), y(v));
      pen = true;
    });
    ctx.stroke();
    values.forEach((v, i) => {
      const c = v === null ? undefined : pointColor(v, i, s.markers);
      if (v === null || !c) return;
      ctx.fillStyle = c;
      ctx.beginPath();
      ctx.arc(x(i), y(v), Math.max(1.5, s.weight), 0, Math.PI * 2);
      ctx.fill();
    });
  } else {
    const slot = w / n, bw = Math.max(1, slot * 0.75);
    const base = s.type === "stacked" ? y0 + h / 2 : y(Math.min(max, Math.max(min, 0)));
    values.forEach((v, i) => {
      if (v === null || (s.type === "stacked" && v === 0)) return;
      ctx.fillStyle = pointColor(v, i, s.color) ?? s.color;
      const top = s.type === "stacked" ? (v > 0 ? y0 : base) : Math.min(base, y(v));
      const height = s.type === "stacked" ? h / 2 : Math.max(1, Math.abs(base - y(v)));
      ctx.fillRect(x0 + i * slot + (slot - bw) / 2, top, bw, height);
    });
  }
  if (s.axis) {
    const ay = s.type === "stacked" ? y0 + h / 2 : y(0);
    if (min <= 0 && max >= 0) {
      ctx.strokeStyle = s.axis;
      ctx.lineWidth = 1;
      ctx.beginPath();
      ctx.moveTo(x0, Math.round(ay) + 0.5);
      ctx.lineTo(x0 + w, Math.round(ay) + 0.5);
      ctx.stroke();
    }
  }
  ctx.restore();
}

/** Lines of `text` wrapped to `width` at spaces (long words break anywhere). */
function wrapLines(ctx: CanvasRenderingContext2D, text: string, width: number): string[] {
  const out: string[] = [];
  for (const para of text.split("\n")) {
    let line = "";
    for (const word of para.split(/(?<= )/)) {
      if (ctx.measureText(line + word).width <= width || !line) {
        line += word;
        // A single word wider than the cell: break it by characters.
        while (ctx.measureText(line).width > width && line.length > 1) {
          let cut = line.length - 1;
          while (cut > 1 && ctx.measureText(line.slice(0, cut)).width > width) cut--;
          out.push(line.slice(0, cut));
          line = line.slice(cut);
        }
      } else {
        out.push(line.trimEnd());
        line = word;
      }
    }
    out.push(line.trimEnd());
  }
  return out;
}

function drawText(ctx: CanvasRenderingContext2D, op: Extract<Op, { kind: "text" }>): void {
  ctx.save();
  ctx.beginPath();
  ctx.rect(op.x, op.y, op.w, op.h);
  ctx.clip();
  ctx.font = op.font;
  ctx.fillStyle = op.color;
  ctx.strokeStyle = op.color;
  ctx.textAlign = op.align;
  ctx.textBaseline = "alphabetic";
  const inner = op.w - 2 * PAD - op.pad;
  const lines = op.wrap ? wrapLines(ctx, op.text, Math.max(1, inner)) : op.text.split("\n");
  const lh = Math.round(op.size * 1.25);
  const block = lh * lines.length;
  const top = op.valign === "top" ? op.y + 1 : op.valign === "center" ? op.y + (op.h - block) / 2 : op.y + op.h - block - 1;
  const x = op.align === "left" ? op.x + PAD + op.pad : op.align === "right" ? op.x + op.w - PAD - op.pad : op.x + op.w / 2;
  lines.forEach((line, i) => {
    // Baseline about 80% down the line box.
    const base = top + i * lh + Math.round(lh * 0.8);
    ctx.fillText(line, x, base);
    if (op.underline || op.strike) {
      const w = ctx.measureText(line).width;
      const left = op.align === "left" ? x : op.align === "right" ? x - w : x - w / 2;
      const y = op.underline ? base + 2 : base - op.size * 0.3;
      ctx.lineWidth = Math.max(1, op.size / 14);
      ctx.beginPath();
      ctx.moveTo(left, y);
      ctx.lineTo(left + w, y);
      ctx.stroke();
    }
  });
  ctx.restore();
}

/** Size `canvas` to `w` × `h` CSS pixels (sharp on high-DPI screens) and draw `ops`. */
export function paint(canvas: HTMLCanvasElement, w: number, h: number, ops: Op[], scale?: number): void {
  const dpr = scale ?? (window.devicePixelRatio || 1);
  const pw = Math.max(1, Math.round(w * dpr)), ph = Math.max(1, Math.round(h * dpr));
  if (canvas.width !== pw) canvas.width = pw;
  if (canvas.height !== ph) canvas.height = ph;
  canvas.setCssProps({ "--w": `${w}px`, "--h": `${h}px` });
  let ctx: CanvasRenderingContext2D | null = null;
  try {
    ctx = canvas.getContext("2d");
  } catch {
    // no canvas support (tests)
  }
  if (!ctx) return;
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  ctx.clearRect(0, 0, w, h);
  for (const op of ops) {
    if (op.kind === "rect") {
      ctx.fillStyle = op.color;
      ctx.fillRect(op.x, op.y, op.w, op.h);
    } else if (op.kind === "line") {
      ctx.strokeStyle = op.color;
      ctx.lineWidth = op.width;
      ctx.setLineDash(op.dash ?? []);
      ctx.beginPath();
      // Odd widths sit on half pixels to stay crisp.
      const o = op.width % 2 ? 0.5 : 0;
      ctx.moveTo(op.x1 + (op.x1 === op.x2 ? o : 0), op.y1 + (op.y1 === op.y2 ? o : 0));
      ctx.lineTo(op.x2 + (op.x1 === op.x2 ? o : 0), op.y2 + (op.y1 === op.y2 ? o : 0));
      ctx.stroke();
      ctx.setLineDash([]);
    } else if (op.kind === "corner") {
      ctx.fillStyle = op.color;
      ctx.beginPath();
      ctx.moveTo(op.x - 6, op.y);
      ctx.lineTo(op.x, op.y);
      ctx.lineTo(op.x, op.y + 6);
      ctx.closePath();
      ctx.fill();
    } else if (op.kind === "spark") drawSpark(ctx, op);
    else drawText(ctx, op);
  }
}
