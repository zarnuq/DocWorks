import { describe, it, expect } from "vitest";
import { Transform } from "prosemirror-transform";
import { Node as PMNode } from "prosemirror-model";
import { writeBody } from "../src/domain/docx/docx-writer";
import { docxSchema as schema } from "../src/domain/docx/schema";
import { DocxDocument } from "../src/infrastructure/docx/docx-document";
import { blockPos, bodyBlocks, readSample, serializeXml, W } from "./helpers";

const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";

async function setup() {
  const s = await readSample();
  const original = bodyBlocks(s.docx.body).map(serializeXml);
  const write = (doc: PMNode) => {
    writeBody(s.docx.body, doc, s.origins, s.docx);
    return bodyBlocks(s.docx.body);
  };
  return { ...s, original, write };
}

/** Text range [from, to) of `needle` inside top-level block `i`. */
function find(doc: PMNode, i: number, needle: string): [number, number] {
  const start = blockPos(doc, i) + 1;
  let found = -1;
  doc.child(i).forEach((child, offset) => {
    if (found < 0 && child.isText && child.text!.includes(needle)) found = start + offset + child.text!.indexOf(needle);
  });
  if (found < 0) throw new Error(`"${needle}" not in block ${i}`);
  return [found, found + needle.length];
}

const children = (el: Element) => Array.from(el.children).map((c) => c.localName);
const all = (el: Element | Document, local: string) => Array.from(el.getElementsByTagNameNS(W, local));

describe("writeBody", () => {
  it("writes an unchanged document back byte-identical", async () => {
    const { doc, write, original } = await setup();
    expect(write(doc).map(serializeXml)).toEqual(original);
  });

  it("keeps the final section properties last", async () => {
    const { doc, write, docx } = await setup();
    write(doc);
    expect(docx.body.lastElementChild!.localName).toBe("sectPr");
    expect(all(docx.body, "sectPr")).toHaveLength(2); // section break paragraph + final
  });

  it("only rewrites the edited paragraph", async () => {
    const { doc, write, original } = await setup();
    const out = write(new Transform(doc).insert(blockPos(doc, 0) + 1, schema.text("X")).doc).map(serializeXml);
    expect(out[0]).not.toBe(original[0]);
    expect(out.slice(1)).toEqual(original.slice(1));
  });

  it("keeps hyperlinks, bookmarks and fields when their paragraph is edited", async () => {
    const { doc, write } = await setup();
    const [from] = find(doc, 2, "See");
    const p = write(new Transform(doc).insert(from, schema.text("Now ")).doc)[2];
    expect(children(p)).toEqual(["bookmarkStart", "r", "hyperlink", "bookmarkEnd", "r", "fldSimple"]);
    const link = all(p, "hyperlink")[0];
    expect(link.getAttributeNS(R, "id")).toMatch(/^rId\d+$/);
    expect(link.textContent).toBe("example");
    expect(all(p, "rStyle")[0].getAttribute("w:val")).toBe("Hyperlink");
    expect(p.textContent).toContain("Now See");
  });

  it("adds bold in schema order and keeps other run properties", async () => {
    const { doc, write } = await setup();
    const [from, to] = find(doc, 1, " red20");
    const p = write(new Transform(doc).addMark(from, to, schema.marks.bold.create()).doc)[1];
    const run = all(p, "r").find((r) => r.textContent === " red20")!;
    expect(children(all(run, "rPr")[0])).toEqual(["b", "color", "sz"]);
  });

  it("removes a mark without touching the rest of rPr", async () => {
    const { doc, write } = await setup();
    const [from, to] = find(doc, 1, " red20");
    const p = write(new Transform(doc).removeMark(from, to, schema.marks.color).doc)[1];
    const run = all(p, "r").find((r) => r.textContent === " red20")!;
    expect(children(all(run, "rPr")[0])).toEqual(["sz"]);
  });

  it("writes palette highlights as w:highlight and other colors as shading", async () => {
    const { doc, write } = await setup();
    const [a, b] = find(doc, 1, "Plain");
    const [c, d] = find(doc, 1, " and ");
    const t = new Transform(doc)
      .addMark(a, b, schema.marks.highlight.create({ color: "#00FF00" }))
      .addMark(c, d, schema.marks.highlight.create({ color: "#123456" }));
    const p = write(t.doc)[1];
    const runs = all(p, "r");
    expect(all(runs[0], "highlight")[0].getAttribute("w:val")).toBe("green");
    const shaded = runs.find((r) => r.textContent === " and ")!;
    expect(all(shaded, "shd")[0].getAttribute("w:fill")).toBe("123456");
    expect(all(shaded, "highlight")).toHaveLength(0);
  });

  it("writes tabs as w:tab", async () => {
    const { doc, write } = await setup();
    const p = write(new Transform(doc).insert(blockPos(doc, 1) + 1, schema.text("Z")).doc)[1];
    expect(all(p, "tab")).toHaveLength(1);
    expect(all(p, "t").some((t) => t.textContent!.includes("\t"))).toBe(false);
  });

  it("changes the paragraph style and keeps other paragraph properties", async () => {
    const { doc, write } = await setup();
    const p = write(new Transform(doc).setNodeMarkup(blockPos(doc, 0), undefined, { ...doc.child(0).attrs, style: "Heading2" }).doc)[0];
    expect(all(p, "pStyle")[0].getAttribute("w:val")).toBe("Heading2");
  });

  it("writes alignment, justify as both", async () => {
    const { doc, write } = await setup();
    const p = write(new Transform(doc).setNodeMarkup(blockPos(doc, 1), undefined, { ...doc.child(1).attrs, align: "justify" }).doc)[1];
    expect(all(p, "jc")[0].getAttribute("w:val")).toBe("both");
  });

  it("keeps a section break on the second half of a split paragraph", async () => {
    const { doc, write } = await setup();
    const [from] = find(doc, 9, "section one");
    const out = write(new Transform(doc).split(from).doc);
    expect(out[9].textContent).toBe("End of ");
    expect(out[10].textContent).toBe("section one");
    expect(all(out[9], "sectPr")).toHaveLength(0);
    expect(all(out[10], "sectPr")).toHaveLength(1);
  });

  it("does not turn rendering hints into page breaks", async () => {
    const { doc, write } = await setup();
    const p = write(new Transform(doc).insert(blockPos(doc, 6) + 1, schema.text("!")).doc)[6];
    expect(all(p, "br").filter((b) => b.getAttribute("w:type") === "page")).toHaveLength(1);
    expect(all(p, "lastRenderedPageBreak")).toHaveLength(0);
  });

  it("starts and nests lists through the numbering part", async () => {
    const { doc, write, docx } = await setup();
    const numId = docx.ensureList("ordered");
    const attrs = { ...doc.child(1).attrs, numId, ilvl: 2 };
    const p = write(new Transform(doc).setNodeMarkup(blockPos(doc, 1), undefined, attrs).doc)[1];
    const numPr = all(p, "numPr")[0];
    expect(children(numPr)).toEqual(["ilvl", "numId"]);
    expect(all(numPr, "ilvl")[0].getAttribute("w:val")).toBe("2");
    expect(all(numPr, "numId")[0].getAttribute("w:val")).toBe(numId);
  });

  it("removes numbering from a list paragraph", async () => {
    const { doc, write } = await setup();
    const p = write(new Transform(doc).setNodeMarkup(blockPos(doc, 12), undefined, { ...doc.child(12).attrs, numId: null, ilvl: 0 }).doc)[12];
    expect(all(p, "numPr")).toHaveLength(0);
  });

  it("resizes images and updates alt text", async () => {
    const { doc, write } = await setup();
    const pos = blockPos(doc, 7) + 1;
    const img = doc.child(7).firstChild!;
    const p = write(new Transform(doc).setNodeMarkup(pos, undefined, { ...img.attrs, width: 192, height: 48, alt: "Blue box" }).doc)[7];
    expect(all(p, "drawing")).toHaveLength(1);
    const extent = p.getElementsByTagNameNS("*", "extent")[0];
    expect([extent.getAttribute("cx"), extent.getAttribute("cy")]).toEqual([String(192 * 9525), String(48 * 9525)]);
    expect(p.getElementsByTagNameNS("*", "docPr")[0].getAttribute("descr")).toBe("Blue box");
  });

  it("writes new images with a unique docPr id", async () => {
    const { doc, write, docx } = await setup();
    const rId = docx.addImage(new Uint8Array([1]), "png");
    const before = docx.nextDocPrId();
    const node = schema.nodes.image.create({ rId, width: 10, height: 20 });
    const p = write(new Transform(doc).insert(blockPos(doc, 1) + 1, node).doc)[1];
    const drawing = all(p, "drawing")[0];
    expect(drawing.getElementsByTagNameNS("*", "blip")[0].getAttributeNS(R, "embed")).toBe(rId);
    expect(Number(drawing.getElementsByTagNameNS("*", "docPr")[0].getAttribute("id"))).toBeGreaterThan(before);
  });

  it("keeps unknown run content when its paragraph is edited", async () => {
    const { doc, write } = await setup();
    const p = write(new Transform(doc).insert(blockPos(doc, 11) + 1, schema.text("!")).doc)[11];
    expect(all(p, "sym")).toHaveLength(1);
  });

  it("keeps cell spans and shading and grows the grid when a column is added", async () => {
    const { doc, write } = await setup();
    const table = doc.child(8);
    const cell = () => schema.nodes.table_cell.create(null, schema.nodes.paragraph.create());
    // Append a cell to each row (row ends are found from the table's own positions).
    let tr = new Transform(doc);
    const tablePos = blockPos(doc, 8);
    for (let r = table.childCount - 1; r >= 0; r--) {
      let rowEnd = tablePos + 1;
      for (let k = 0; k <= r; k++) rowEnd += table.child(k).nodeSize;
      tr = tr.insert(rowEnd - 1, cell());
    }
    const tbl = write(tr.doc)[8];
    expect(all(tbl, "gridCol")).toHaveLength(4);
    expect(all(tbl, "gridSpan")[0].getAttribute("w:val")).toBe("2");
    expect(all(tbl, "shd")[0].getAttribute("w:fill")).toBe("FFFF00");
    for (const tc of all(tbl, "tc")) expect(tc.lastElementChild!.localName).toBe("p");
  });

  it("builds new tables with borders and a grid", async () => {
    const { doc, write } = await setup();
    const table = schema.nodes.table.create(null, [
      schema.nodes.table_row.create(null, [
        schema.nodes.table_cell.create(null, schema.nodes.paragraph.create(null, schema.text("a"))),
        schema.nodes.table_cell.create({ background: "#ABCDEF" }, schema.nodes.paragraph.create()),
      ]),
    ]);
    const tbl = write(new Transform(doc).insert(blockPos(doc, 1), table).doc)[1];
    expect(children(tbl)).toEqual(["tblPr", "tblGrid", "tr"]);
    expect(all(tbl, "tblBorders")).toHaveLength(1);
    expect(all(tbl, "gridCol")).toHaveLength(2);
    expect(all(tbl, "shd")[0].getAttribute("w:fill")).toBe("ABCDEF");
  });

  it("writes horizontal rules as bottom-bordered paragraphs", async () => {
    const { doc, write } = await setup();
    const hr = schema.nodes.raw_block.create({ kind: "hr" });
    const p = write(new Transform(doc).insert(blockPos(doc, 1), hr).doc)[1];
    expect(p.localName).toBe("p");
    expect(all(p, "bottom")).toHaveLength(1);
  });

  it("writes page breaks as w:br type=page", async () => {
    const { doc, write } = await setup();
    const p = write(new Transform(doc).insert(blockPos(doc, 0) + 1, schema.nodes.page_break.create()).doc)[0];
    expect(all(p, "br")[0].getAttribute("w:type")).toBe("page");
  });

  it("does not duplicate paraIds when a paragraph is copied", async () => {
    const { doc, write, origins, docx } = await setup();
    origins.get(doc.child(0).attrs.key)!.el.setAttributeNS(W14, "w14:paraId", "1234ABCD");
    const out = write(new Transform(doc).insert(blockPos(doc, 1), doc.child(0)).doc);
    const ids = out.map((p) => p.getAttributeNS(W14, "paraId")).filter(Boolean);
    expect(ids).toEqual(["1234ABCD"]);
    // Serialized in context, the attribute keeps the w14 prefix the root declares.
    expect(serializeXml(docx.document)).toContain('w14:paraId="1234ABCD"');
  });

  it("survives save and reload", async () => {
    const { doc, write, docx } = await setup();
    write(new Transform(doc).insert(blockPos(doc, 0) + 1, schema.text("Saved ")).doc);
    const back = await DocxDocument.load(await docx.save());
    expect(back.body.textContent).toContain("Saved Doc Heading");
  });
});
