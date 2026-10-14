/**
 * Copy as Markdown outline: each slide's title as a heading, its other text as nested bullets, tables as Markdown
 * tables and speaker notes quoted. A slide without a title placeholder takes its clearly largest text as the title.
 */
import type { PptxDocument } from "../../infrastructure/pptx/pptx-document";
import { Para, parseSlide, ShapeModel } from "./slide-model";

const isTitle = (shape: ShapeModel) => ["title", "ctrTitle"].includes(shape.el.getElementsByTagNameNS("*", "ph")[0]?.getAttribute("type") ?? "");
const paraText = (p: Para) => p.runs.map((r) => r.text).join("").replace(/\s*\n\s*/g, " ").trim();
const paraSize = (p: Para) => Math.max(0, ...p.runs.filter((r) => r.text.trim()).map((r) => r.sz));
const cellText = (shape: ShapeModel) => shape.text!.paras.map(paraText).filter(Boolean).join(" ").replace(/\|/g, "\\|");

/** Rows of a table's cells (the "cell" shapes listed after it), first row as the header. */
function table(cells: ShapeModel[]): string {
  const rows: string[][] = [];
  for (const c of cells) if (c.cell && !c.cell.hidden) (rows[c.cell.r] ??= [])[c.cell.c] = cellText(c);
  const grid = rows.filter(Boolean).map((r) => Array.from(r, (t) => t ?? ""));
  if (!grid.length) return "";
  const width = Math.max(...grid.map((r) => r.length));
  const line = (r: string[]) => `| ${Array.from({ length: width }, (_, i) => r[i] ?? "").join(" | ")} |`;
  return [line(grid[0]), line(Array(width).fill("---") as string[]), ...grid.slice(1).map(line)].join("\n");
}

export function deckToMarkdown(doc: PptxDocument): string {
  return doc
    .slides()
    .map((ref, i) => {
      const shapes = parseSlide(doc, ref.path).shapes.filter((s) => !s.inherited);
      const texts = shapes.filter((s) => s.text && s.kind !== "cell");
      let titleParas = new Set<Para>(texts.filter(isTitle).flatMap((s) => s.text!.paras));
      if (!titleParas.size) {
        // No title placeholder: the largest text, when it stands out (a quarter bigger than the size most of the
        // slide's characters have).
        const paras = texts.flatMap((s) => s.text!.paras).filter((p) => paraText(p));
        const sizes = paras.flatMap((p) => Array<number>(paraText(p).length).fill(paraSize(p))).sort((a, b) => a - b);
        const usual = sizes[Math.floor(sizes.length / 2)] ?? 0, top = sizes[sizes.length - 1] ?? 0;
        if (paras.length > 1 && top >= usual * 1.25) titleParas = new Set(paras.filter((p) => paraSize(p) === top));
      }
      const title = [...titleParas].map(paraText).filter(Boolean).join(" ");
      const blocks: string[] = [];
      let bullets: string[] = [];
      const flush = () => {
        if (bullets.length) blocks.push(bullets.join("\n"));
        bullets = [];
      };
      shapes.forEach((s, k) => {
        if (s.kind === "table") {
          const cells: ShapeModel[] = [];
          for (let j = k + 1; shapes[j]?.kind === "cell"; j++) cells.push(shapes[j]);
          flush();
          const md = table(cells);
          if (md) blocks.push(md);
          return;
        }
        if (!s.text || s.kind === "cell") return;
        for (const p of s.text.paras) {
          const text = paraText(p);
          if (text && !titleParas.has(p)) bullets.push(`${"  ".repeat(Math.max(0, p.lvl))}- ${text}`);
        }
      });
      flush();
      const notes = doc.notesText(i).trim();
      if (notes) blocks.push(notes.split(/\r?\n/).map((l) => `> ${l}`).join("\n"));
      return [`## ${title || `Slide ${i + 1}`}`, ...blocks].join("\n\n") + "\n";
    })
    .join("\n");
}
