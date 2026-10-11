/**
 * Smoke tests against real files: DOC_WORKS_DECKS=<dir> / DOC_WORKS_DOCS=<dir> npm test
 * Every .pptx is parsed, rendered and round-tripped; every .docx must write back
 * byte-identical when unchanged and keep its text through an edit + save.
 */
import { describe, it, expect } from "vitest";
import { readdirSync, readFileSync, writeFileSync } from "fs";
import { join } from "path";
import { PptxDocument } from "../src/infrastructure/pptx/pptx-document";
import { parseSlide } from "../src/domain/pptx/slide-model";
import { fillCss } from "../src/domain/pptx/shape-style";
import { writeParagraphs } from "../src/domain/pptx/slide-edit";
import { renderSlide } from "../src/presentation/pptx/slide-renderer";
import { Transform } from "prosemirror-transform";
import { DocxDocument } from "../src/infrastructure/docx/docx-document";
import { readDocx } from "../src/domain/docx/docx-reader";
import { writeBody } from "../src/domain/docx/docx-writer";
import { docxSchema } from "../src/domain/docx/schema";

const dir = process.env.DOC_WORKS_DECKS;
const decks = dir ? readdirSync(dir).filter((f) => f.endsWith(".pptx")) : [];

describe.skipIf(!dir)("real decks", () => {
  it.each(decks)("%s parses, renders, and survives an edit round trip", async (name) => {
    const doc = await PptxDocument.load(new Uint8Array(readFileSync(join(dir!, name))));
    const host = document.createElement("div");
    for (const slide of doc.slides()) {
      const model = parseSlide(doc, slide.path);
      renderSlide(host, model, doc.slideSize, { width: 960, editable: true, imageUrl: () => null });
      // Rewriting unchanged text must not alter the XML.
      for (const shape of model.shapes.filter((s) => s.text && !s.inherited)) {
        const before = new XMLSerializer().serializeToString(shape.text!.el);
        writeParagraphs(shape.text!, shape.text!.paras);
        expect(new XMLSerializer().serializeToString(shape.text!.el)).toBe(before);
      }
    }
    // Sanity checks that catch whole classes of rendering bugs on real decks.
    const media = new Set(doc.mediaPaths());
    const css = /^(#[0-9A-F]{6}([0-9A-F]{2})?|(linear|radial)-gradient\(.*\))$/;
    for (const slide of doc.slides()) {
      const model = parseSlide(doc, slide.path);
      for (const s of model.shapes) {
        const where = `${slide.path} ${s.name}`;
        expect(Object.values(s.rect).every(Number.isFinite), where).toBe(true);
        if (s.fill) expect(fillCss(s.fill), where).toMatch(css);
        if (s.el.getAttribute("useBgFill") === "1") expect(fillCss(s.fill), where).toBe(model.background.color);
        for (const p of s.paths ?? []) expect(p.d, where).toMatch(/^[MLACQZ0-9 .-]*$/);
        if (s.image) expect(media.has(s.image), where).toBe(true);
      }
      // SmartArt with a cached drawing must be drawn, not left as a placeholder frame.
      if (doc.rels(slide.path).some((r) => r.type === "diagramDrawing")) {
        expect(model.shapes.filter((s) => s.kind === "frame" && !s.diagram && s.el.getElementsByTagNameNS("*", "relIds").length), slide.path).toHaveLength(0);
      }
    }

    doc.duplicateSlide(0);
    const back = await PptxDocument.load(await doc.save());
    expect(back.slides()).toHaveLength(doc.slides().length);
  });
});


const docsDir = process.env.DOC_WORKS_DOCS;
const docs = docsDir ? readdirSync(docsDir).filter((f) => f.endsWith(".docx")) : [];

describe.skipIf(!docsDir)("real docx files", () => {
  it.each(docs)("%s writes back unchanged and survives an edit", async (name) => {
    const docx = await DocxDocument.load(new Uint8Array(readFileSync(join(docsDir!, name))));
    const before = new XMLSerializer().serializeToString(docx.body);
    const { doc, origins } = readDocx(docx.body, { imageSrc: () => null, hyperlinkHref: (id) => docx.hyperlinkHref(id) });
    writeBody(docx.body, doc, origins, docx);
    expect(new XMLSerializer().serializeToString(docx.body)).toBe(before);

    // Edit every paragraph, then check text survives save + reload.
    const tr = new Transform(doc);
    doc.descendants((node, pos) => {
      if (node.type === docxSchema.nodes.paragraph) tr.insert(tr.mapping.map(pos + 1), docxSchema.text("✎"));
      return node.type !== docxSchema.nodes.paragraph;
    });
    writeBody(docx.body, tr.doc, origins, docx);
    const back = await DocxDocument.load(await docx.save());
    const reread = readDocx(back.body, { imageSrc: () => null, hyperlinkHref: () => null }).doc;
    expect(reread.textContent).toBe(tr.doc.textContent);
  });
});

// ---------------------------------------------------------------------------
// Workbooks: DOC_WORKS_BOOKS=<dir>
// ---------------------------------------------------------------------------

import JSZip from "jszip";
import { XlsxWorkbook } from "../src/infrastructure/xlsx/xlsx-workbook";
import { Workbook } from "../src/domain/xlsx/workbook-model";
import { Calc } from "../src/domain/xlsx/calculation";
import { Edits } from "../src/domain/xlsx/commands";

const bookDir = process.env.DOC_WORKS_BOOKS;
const books = bookDir ? readdirSync(bookDir).filter((f) => f.endsWith(".xlsx")) : [];

async function zipTexts(bytes: Uint8Array): Promise<Map<string, string>> {
  const z = await JSZip.loadAsync(bytes);
  const out = new Map<string, string>();
  for (const f of Object.values(z.files)) if (!f.dir) out.set(f.name, await f.async("base64"));
  return out;
}

describe.skipIf(!bookDir)("real workbooks", () => {
  it.each(books)("%s keeps every part on a no-op save", async (name) => {
    const bytes = new Uint8Array(readFileSync(join(bookDir!, name)));
    const pkg = await XlsxWorkbook.load(bytes);
    Workbook.open(pkg);
    expect(await zipTexts(await pkg.save())).toEqual(await zipTexts(bytes));
  });

  it.each(books)("%s: supported formulas calculate to the cached results", async (name) => {
    const wb = Workbook.open(await XlsxWorkbook.load(new Uint8Array(readFileSync(join(bookDir!, name)))));
    const calc = new Calc(wb);
    const wrong: string[] = [];
    let checked = 0, unsupported = 0;
    for (const sheet of wb.sheets) {
      for (const cell of sheet.cells.values()) {
        if (cell.formula === undefined || cell.noCache || cell.array) continue;
        const v = calc.evaluate(sheet, cell.row, cell.col, cell.formula);
        if (v === Calc.UNSUPPORTED) {
          unsupported++;
          continue;
        }
        checked++;
        const same = typeof v === "number" && typeof cell.value === "number" ? Math.abs(v - cell.value) <= 1e-9 * Math.max(1, Math.abs(v)) : JSON.stringify(v) === JSON.stringify(cell.value);
        if (!same) wrong.push(`${sheet.name}!${cell.el.getAttribute("r")} =${cell.formula} -> ${JSON.stringify(v)} (cached ${JSON.stringify(cell.value)})`);
      }
    }
    console.log(`${name}: ${checked} formulas checked, ${unsupported} unsupported`);
    expect(wrong.slice(0, 20)).toEqual([]);
  });

  it.each(books)("%s survives an edit and reload", async (name) => {
    const e = new Edits(Workbook.open(await XlsxWorkbook.load(new Uint8Array(readFileSync(join(bookDir!, name))))));
    const s = e.wb.sheets[0];
    const used = s.used() ?? { r1: 0, c1: 0, r2: 0, c2: 0 };
    e.input(s, used.r2 + 2, 0, "edited");
    e.input(s, used.r2 + 2, 1, "=1+1");
    e.format(s, { r1: used.r2 + 2, c1: 0, r2: used.r2 + 2, c2: 1 }, { bold: true, fill: "#FFFF00" });
    const saved = await e.wb.pkg.save();
    if (process.env.DOC_WORKS_EXPORT) writeFileSync(join(process.env.DOC_WORKS_EXPORT, `edited-${name}`), saved);
    const back = Workbook.open(await XlsxWorkbook.load(saved));
    expect(back.sheets[0].at(used.r2 + 2, 0)!.value).toBe("edited");
    expect(back.sheets.map((x) => x.name)).toEqual(e.wb.sheets.map((x) => x.name));
  });
});
