import { addChart, chartWorkbook, prepareChartEdit } from "../src/domain/pptx/chart";
import { addMedia, setMediaPlayback } from "../src/domain/pptx/media";
/*
 * Writes heavily edited copies of the fixtures (and of real files, when
 * DOC_WORKS_DECKS / DOC_WORKS_DOCS are set) into DOC_WORKS_EXPORT, so another
 * program (LibreOffice, PowerPoint, Word) can check they open cleanly:
 *
 *   DOC_WORKS_EXPORT=/tmp/out npx vitest run tests/export-samples.test.ts
 *   soffice --headless --convert-to pdf --outdir /tmp/out/pdf /tmp/out/*.pptx /tmp/out/*.docx
 */
import { describe, it, expect } from "vitest";
import { mkdirSync, readdirSync, readFileSync, writeFileSync } from "fs";
import { join } from "path";
import { EditorState, TextSelection, Command } from "prosemirror-state";
import { PptxDocument } from "../src/infrastructure/pptx/pptx-document";
import { DocxDocument } from "../src/infrastructure/docx/docx-document";
import { PRESET_NAMES } from "../src/domain/pptx/geometry";
import { parseSlide } from "../src/domain/pptx/slide-model";
import { addShape, addTextBox, connect, CONNECTORS, rerouteConnectors, reorder, setAdjust, setTransform, shapeId, writeBodyPr, writeParagraphs } from "../src/domain/pptx/slide-edit";
import { writeFill, writeLine, writeShadow } from "../src/domain/pptx/shape-style";
import { alignRects, copyItems, distributeRects, groupItems, itemRect, pasteItems, setItemRect, ungroupItem } from "../src/domain/pptx/arrange";
import { addTable, insertCol, insertRow, mergeCells, setCellBorders, setCellFill, setTableFlags, TABLE_STYLES } from "../src/domain/pptx/table";
import { replaceAll } from "../src/domain/pptx/find";
import { setCrop } from "../src/domain/pptx/slide-edit";
import { readDocx } from "../src/domain/docx/docx-reader";
import { writeBody } from "../src/domain/docx/docx-writer";
import { docxSchema as schema } from "../src/domain/docx/schema";
import { DocxStyles } from "../src/domain/docx/docx-styles";
import { cmd, EditorContext, pageNumberNode } from "../src/presentation/docx/docx-editor";
import { acceptRevisions } from "../src/domain/docx/revisions";
import { writeTransition } from "../src/domain/pptx/animation";
import { addAnimation } from "../src/domain/pptx/animation-edit";
const fixture = (file: string) => join(__dirname, "fixtures", file);

const out = process.env.DOC_WORKS_EXPORT;
const P = "http://schemas.openxmlformats.org/presentationml/2006/main";
const spPr = (el: Element) => el.getElementsByTagNameNS(P, "spPr")[0];
const PNG = Uint8Array.from(
  atob("iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAFklEQVR4nGP8z8DAwMDAxMDAwMDAAAANHQEDasKb6QAAAABJRU5ErkJggg=="),
  (c) => c.charCodeAt(0)
);

/** Every pptx edit the editor can make, applied to one deck. */
async function editDeck(doc: PptxDocument): Promise<void> {
  const slides = doc.slides();
  const authoringIndex=doc.addSlide(slides.length-1);
  const authoring=doc.slides()[authoringIndex].path;
  const data={categories:["North","South"],series:[{name:"Sales",values:[10,20]}]};
  const chart=addChart(doc,authoring,data,"barChart",await chartWorkbook(data));
  const chartPath=parseSlide(doc,authoring).shapes.find(s=>s.el===chart)!.chart!.path;
  const update=await prepareChartEdit(doc,chartPath,{...data,series:[{name:"Income",values:[15,25]}]});
  doc.putXml(chartPath,update.xml);doc.putBinary(update.workbookPath!,update.workbook!);
  for(const [kind,ext] of [["audio","wav"],["video","mp4"]] as const){const media=addMedia(doc,authoring,new Uint8Array(readFileSync(fixture(`sample-${kind}.${ext}`))),kind,ext,kind);setMediaPlayback(doc.xml(authoring)!,media,{autoPlay:false,loop:true});}
  const sourceLayout=doc.layoutOf(authoring)!;const custom=doc.createLayout(sourceLayout,"Custom test layout");
  doc.addLayoutPlaceholder(custom,"body"); doc.setLayout(authoringIndex,custom);
  const theme=doc.themeOf(doc.masterOf(custom)!);if(theme)doc.setTheme(theme,{accent1:"4285f4"});
  // A slide with every preset shape, each styled differently.
  const index = doc.addSlide(slides.length - 1);
  const path = doc.slides()[index].path;
  const xml = doc.xml(path)!;
  const { cx, cy } = doc.slideSize;
  const cols = 17;
  const cell = { w: cx / cols, h: cy / Math.ceil(PRESET_NAMES.length / cols) };
  const added = PRESET_NAMES.map((prst, i) => {
    const rect = { x: (i % cols) * cell.w + cell.w * 0.1, y: Math.floor(i / cols) * cell.h + cell.h * 0.1, w: cell.w * 0.8, h: cell.h * 0.8 };
    const el = addShape(xml, prst, rect, { tailArrow: CONNECTORS.has(prst) });
    const pr = spPr(el);
    if (!CONNECTORS.has(prst)) {
      if (i % 4 === 1) writeFill(pr, { type: "solid", color: "#C0504D80" });
      if (i % 4 === 2) writeFill(pr, { type: "gradient", radial: false, angle: 45, stops: [{ pos: 0, color: "#FFFFFF" }, { pos: 1, color: "#1F497D" }] });
      if (i % 4 === 3) writeFill(pr, { type: "gradient", radial: true, angle: 0, stops: [{ pos: 0, color: "#FFFF00" }, { pos: 1, color: "#F79646" }] });
      if (i % 5 === 0) writeShadow(pr, { color: "#00000066", blur: 50800, dist: 38100, dir: 45 });
      if (i % 7 === 0) setTransform(el, { rot: 15, flipH: true });
    }
    writeLine(pr, { color: "#000000", width: 12700 + (i % 3) * 12700, dash: ["solid", "dash", "sysDot", "lgDashDot"][i % 4] });
    if (i % 6 === 0) writeLine(pr, { head: { type: "oval", w: "sm", len: "sm" }, tail: { type: "stealth", w: "lg", len: "lg" } });
    return el;
  });
  setAdjust(added[PRESET_NAMES.indexOf("chevron")], "adj", 30000);
  reorder(added[0], "front");
  reorder(added[1], "back");

  // Glued connectors between two shapes.
  const a = addShape(xml, "roundRect", { x: cx * 0.05, y: cy * 0.9, w: cx * 0.1, h: cy * 0.06 });
  const b = addShape(xml, "ellipse", { x: cx * 0.4, y: cy * 0.8, w: cx * 0.1, h: cy * 0.06 });
  for (const prst of ["straightConnector1", "bentConnector3", "curvedConnector3"]) {
    const c = addShape(xml, prst, { x: 0, y: 0, w: 1, h: 1 }, { tailArrow: true });
    connect(c, "st", { id: shapeId(a), idx: 3 });
    connect(c, "end", { id: shapeId(b), idx: 1 });
  }
  rerouteConnectors(xml, parseSlide(doc, path).shapes);

  // Text, notes, hiding, images, slide operations.
  const tb = addTextBox(xml, { x: cx * 0.6, y: cy * 0.85, w: cx * 0.35, h: cy * 0.1 }, "Edited by Doc Works");
  writeTransition(xml, { type: "push", dir: "l", dur: 850 });
  addAnimation(xml, shapeId(tb), { type: "fade-in", dur: 700 });
  addAnimation(xml, shapeId(a), { type: "fly-left", trigger: "after", dur: 900 });
  const model = parseSlide(doc, path).shapes.find((s) => s.el === tb)!;
  writeParagraphs(model.text!, [{ align: "ctr", lvl: 0, bullet: null, runs: [{ ...model.text!.paras[0].runs[0], text: "Edited ", b: true }, { ...model.text!.paras[0].runs[0], text: "by Doc Works", i: true, color: "#C0504D" }] }]);
  // Text: numbering, spacing, character spacing, a link, a field, text box settings.
  const box = addTextBox(xml, { x: cx * 0.05, y: cy * 0.05, w: cx * 0.4, h: cy * 0.3 }, "x");
  const tbody = parseSlide(doc, path).shapes.find((s) => s.el === box)!.text!;
  const base = tbody.paras[0].runs[0];
  writeParagraphs(tbody, [
    { align: "l", lvl: 0, bullet: "#arabicPeriod", lnSpc: { pct: 1.5 }, spcBef: { pts: 6 }, runs: [{ ...base, text: "Numbered " }, { ...base, text: "link", link: "https://example.com/" }] },
    { align: "l", lvl: 1, bullet: "#alphaLcParenR@3", spcAft: { pts: 12 }, runs: [{ ...base, text: "Spaced out", spc: 3 }] },
    { align: "ctr", lvl: 0, bullet: null, runs: [{ ...base, text: "Slide " }, { ...base, text: "1", field: "slidenum" }, { ...base, text: " on " }, { ...base, text: "today", field: "datetime1" }] },
  ]);
  writeBodyPr(tbody.el, { anchor: "ctr", insets: { l: 0, t: 0, r: 0, b: 0 }, wrap: true, autofit: { type: "shrink", fontScale: 0.9, lnSpcReduction: 0.1 } });

  // Arranging: group, ungroup, align, distribute, copy and paste across slides.
  const trio = [0, 1, 2].map((k) => addShape(xml, "ellipse", { x: cx * (0.55 + k * 0.1), y: cy * 0.05, w: cx * 0.05, h: cy * 0.05 }));
  const rects = alignRects(trio.map(itemRect), "top");
  trio.forEach((t, k) => setItemRect(t, rects[k]));
  distributeRects(trio.map(itemRect), "horizontal").forEach((r, k) => setItemRect(trio[k], r));
  const group = groupItems(xml, trio);
  setItemRect(group, { ...itemRect(group), w: itemRect(group).w * 1.5 });
  const pair = [addShape(xml, "rect", { x: 0, y: cy * 0.6, w: cx * 0.05, h: cy * 0.05 }), addShape(xml, "rect", { x: cx * 0.1, y: cy * 0.6, w: cx * 0.05, h: cy * 0.05 })];
  ungroupItem(groupItems(xml, pair));
  const clip = await copyItems(doc, path, [group, box, ...Array.from(xml.getElementsByTagNameNS(P, "pic"))]);
  pasteItems(doc, doc.slides()[0].path, clip, { dx: 0, dy: 0 });

  // Tables, crop, find and replace; slides: layout, background, size, sections.
  for (const [k, style] of [TABLE_STYLES[1], TABLE_STYLES[4], TABLE_STYLES[7]].entries()) {
    doc.ensureTableStyle(style.id);
    const frame = addTable(xml, 3, 3, { x: cx * (0.05 + k * 0.32), y: cy * 0.4, w: cx * 0.28, h: cy * 0.15 }, style.id);
    const tbl = frame.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/main", "tbl")[0];
    insertRow(tbl, 2, true);
    insertCol(tbl, 0, false);
    mergeCells(tbl, [1, 1, 2, 2]);
    setCellFill(tbl, [[0, 0]], { type: "solid", color: "#FFFF00" });
    setCellBorders(tbl, [0, 0, 3, 3], "outer", { color: "#FF0000", width: 25400 });
    setTableFlags(tbl, { firstCol: true, bandRow: k !== 1 });
  }
  for (const pic of Array.from(xml.getElementsByTagNameNS(P, "pic"))) setCrop(pic, { l: 0.1, t: 0.2, r: 0, b: 0.1 });
  replaceAll(doc, "Doc Works", "DocWorks", { matchCase: true });
  const layouts = doc.layouts();
  if (layouts.length > 1) doc.setLayout(0, layouts[1].path);
  doc.setBackground(index, { type: "gradient", radial: false, angle: 90, stops: [{ pos: 0, color: "#FFFFFF" }, { pos: 1, color: "#DDEEFF" }] });
  doc.addSection(0, "Intro");
  doc.addSection(Math.min(2, doc.slides().length - 1), "Body");

  doc.setNotes(index, "Notes written by the export test\nSecond line");
  doc.setNotes(0, "Changed notes on the first slide");
  doc.setHidden(Math.min(1, doc.slides().length - 1), true);
  doc.addImage(path, PNG, "png");
  doc.duplicateSlide(0);
  doc.moveSlide(0, 1);
  doc.setSlideSize(12192000, 6858000);
}

/** A run of Word edits through the editor's commands. */
function editDocx(docx: DocxDocument): void {
  const { doc, origins } = readDocx(docx.body, { imageSrc: () => null, hyperlinkHref: (id) => docx.hyperlinkHref(id) });
  const ctx: EditorContext = {
    styles: new DocxStyles(docx.stylesXml(), docx.numberingXml()),
    ensureList: (kind) => docx.ensureList(kind),
    listKind: (numId) => {
      const fmt = docx.listFormat(numId, 0);
      return fmt ? (fmt === "bullet" ? "bullet" : "ordered") : null;
    },
  };
  let st = EditorState.create({ doc });
  const run = (c: Command) => c(st, (tr) => (st = st.apply(tr)));
  const at = (pos: number, to = pos) => (st = st.apply(st.tr.setSelection(TextSelection.create(st.doc, pos, to))));

  // Type into every paragraph, then format the start of the document.
  const tr = st.tr;
  st.doc.descendants((node, pos) => {
    if (node.type === schema.nodes.paragraph) tr.insert(tr.mapping.map(pos + 1), schema.text("✎ "));
    return node.type !== schema.nodes.paragraph;
  });
  st = st.apply(tr);
  at(1, 3);
  run(cmd.setMark(schema.marks.bold, null));
  run(cmd.setMark(schema.marks.color, { color: "#C0504D" }));
  run(cmd.setMark(schema.marks.highlight, { color: "#FFFF00" }));
  run(cmd.setStyle("Heading2"));
  run(cmd.setAlign("center"));
  at(1);
  run(cmd.toggleList(ctx, "bullet"));
  run(cmd.shiftLevel(1));
  run(cmd.toggleList(ctx, "ordered"));
  run(cmd.insertPageBreak);
  run(cmd.setEquation(null, "\\frac{-b\\pm \\sqrt{b^2-4ac}}{2a}+\\sum_{i=1}^{n}{x_i}"));
  run(cmd.insertTable(3, 3));
  run(cmd.addRow(true));
  run(cmd.addColumn(true));
  run(cmd.setCellBackground("#DDEEFF"));
  const rId = docx.addImage(PNG, "png");
  st = st.apply(st.tr.replaceSelectionWith(schema.nodes.image.create({ rId, src: "", width: 40, height: 40, alt: "dot" })));
  run(cmd.insertRule);
  // Word features: links, spacing, merges, borders, widths, comments, footnotes, contents, header / footer.
  at(3, 6);
  run(cmd.setLink("https://example.com/"));
  run(cmd.setParagraph({ spaceBefore: 6, spaceAfter: 12, firstLine: -18 }));
  run(cmd.addCommentAnchors(docx.addComment("Exported comment", "Doc Works")));
  at(4);
  run(cmd.insertFootnote(docx.addFootnote("Exported footnote.")));
  docx.prepareToc();
  run(cmd.insertToc);
  let cellPos = -1;
  st.doc.descendants((n, pos) => {
    if (cellPos < 0 && n.type === schema.nodes.table_cell) cellPos = pos + 2;
    return cellPos < 0;
  });
  at(cellPos);
  run(cmd.mergeRight);
  at(cellPos);
  run(cmd.mergeDown);
  run(cmd.setTableBorders("outer"));
  run(cmd.setColumnWidth(2880));
  run(cmd.toggleHeaderRow);
  writeBody(docx.body, st.doc, origins, docx);
  const footer = docx.ensureHeaderFooter("footer");
  const read = readDocx(footer, { imageSrc: () => null, hyperlinkHref: () => null });
  writeBody(footer, schema.nodes.doc.create(null, schema.nodes.paragraph.create({ align: "center" }, [schema.text("Page "), pageNumberNode()])), read.origins, docx);
  const header = docx.ensureHeaderFooter("header");
  writeBody(header, schema.nodes.doc.create(null, schema.nodes.paragraph.create(null, schema.text("Exported header"))), readDocx(header, { imageSrc: () => null, hyperlinkHref: () => null }).origins, docx);
  acceptRevisions(docx.body);
}

const decks = (dir?: string) => (dir ? readdirSync(dir).filter((f) => f.endsWith(".pptx")).map((f) => join(dir, f)) : []);
const docs = (dir?: string) => (dir ? readdirSync(dir).filter((f) => f.endsWith(".docx")).map((f) => join(dir, f)) : []);
const name = (path: string) => `edited-${path.split("/").pop()}`;

describe.skipIf(!out)("export edited files", () => {
  if (out) mkdirSync(out, { recursive: true });

  it.each([fixture("sample.pptx"), ...decks(process.env.DOC_WORKS_DECKS)])("pptx %s", async (path) => {
    const doc = await PptxDocument.load(new Uint8Array(readFileSync(path)));
    await editDeck(doc);
    const bytes = await doc.save();
    writeFileSync(join(out!, name(path)), bytes);
    expect((await PptxDocument.load(bytes)).slides().length).toBeGreaterThan(1);
  });

  it.each([fixture("sample.docx"), ...docs(process.env.DOC_WORKS_DOCS)])("docx %s", async (path) => {
    const docx = await DocxDocument.load(new Uint8Array(readFileSync(path)));
    editDocx(docx);
    const bytes = await docx.save();
    writeFileSync(join(out!, name(path)), bytes);
    expect((await DocxDocument.load(bytes)).body).toBeTruthy();
  });
});
