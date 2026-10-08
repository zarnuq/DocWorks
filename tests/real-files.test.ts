/**
 * Smoke tests against real files: DOC_WORKS_DECKS=<dir> / DOC_WORKS_DOCS=<dir> npm test
 * Every .pptx is parsed, rendered and round-tripped; every .docx must write back
 * byte-identical when unchanged and keep its text through an edit + save.
 */
import { describe, it, expect } from "vitest";
import { readdirSync, readFileSync } from "fs";
import { join } from "path";
import { PptxDocument } from "../src/infrastructure/pptx/pptx-document";
import { parseSlide } from "../src/domain/pptx/slide-model";
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
        if (s.fill) expect(s.fill, where).toMatch(css);
        if (s.el.getAttribute("useBgFill") === "1") expect(s.fill, where).toBe(model.background.color);
        for (const p of s.paths ?? []) expect(p.d, where).toMatch(/^[MLCQZ0-9 .-]*$/);
        if (s.image) expect(media.has(s.image), where).toBe(true);
      }
      // SmartArt with a cached drawing must be drawn, not left as a placeholder frame.
      if (doc.rels(slide.path).some((r) => r.type === "diagramDrawing")) {
        expect(model.shapes.filter((s) => s.kind === "frame" && s.el.getElementsByTagNameNS("*", "relIds").length), slide.path).toHaveLength(0);
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
