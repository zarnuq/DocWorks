/*
 * Cells are painted on a canvas per pane, like Google Sheets does: hundreds of positioned elements per frame cost
 * more in style and layout (against Obsidian's large stylesheet) than drawing them. The grid builds a list of
 * paint operations (plain data, which tests read), and `paint` draws them.
 */

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
  | { kind: "corner"; x: number; y: number; color: string };

const PAD = 3;

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
    } else drawText(ctx, op);
  }
}
