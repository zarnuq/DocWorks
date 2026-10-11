import { describe, it, expect } from "vitest";
import JSZip from "jszip";
import { DocxDocument } from "../src/infrastructure/docx/docx-document";
import { blankDocx } from "../src/infrastructure/blank";
import { fixture, W } from "./helpers";

const load = () => DocxDocument.load(fixture("sample.docx"));

describe("DocxDocument", () => {
  it("finds the main part through the package relationships", async () => {
    const d = await load();
    expect(d.mainPath).toBe("word/document.xml");
    expect(d.body.localName).toBe("body");
  });

  it("rejects packages without a document part", async () => {
    await expect(DocxDocument.load(fixture("sample.pptx"))).rejects.toThrow(/document\.xml is missing/);
  });

  it("loads images as data URLs keyed by relationship id", async () => {
    const urls = await (await load()).imageUrls();
    const [url] = Array.from(urls.values());
    expect(url).toMatch(/^data:image\/png;base64,/);
  });

  it("resolves external hyperlink targets", async () => {
    const d = await load();
    const rId = d.body.getElementsByTagNameNS(W, "hyperlink")[0].getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id")!;
    expect(d.hyperlinkHref(rId)).toBe("https://example.com");
    expect(d.hyperlinkHref("nope")).toBeNull();
  });

  it("adds images as media + relationship + content type", async () => {
    const d = await load();
    const rId = d.addImage(new Uint8Array([7, 7]), "gif");
    const back = await DocxDocument.load(await d.save());
    expect(back.relTarget(back.mainPath, rId)).toMatch(/^word\/media\/image\d+\.gif$/);
    const zip = await JSZip.loadAsync(await back.save());
    expect(await zip.file("[Content_Types].xml")!.async("text")).toContain('Extension="gif"');
  });

  it("hands out docPr ids above every existing one", async () => {
    const d = await load();
    const ids = Array.from(d.document.getElementsByTagNameNS("*", "docPr")).map((e) => Number(e.getAttribute("id")));
    const next = d.nextDocPrId();
    expect(ids.every((id) => next > id)).toBe(true);
    expect(d.nextDocPrId()).toBe(next + 1);
  });

  it("reuses a matching list definition when one exists", async () => {
    const d = await load();
    const bullet = d.ensureList("bullet");
    const ordered = d.ensureList("ordered");
    expect(d.listFormat(bullet, 0)).toBe("bullet");
    expect(d.listFormat(ordered, 0)).toBe("decimal");
    expect(d.ensureList("bullet")).toBe(bullet);
  });

  it("creates numbering.xml with its relationship and content type when missing", async () => {
    const zip = await JSZip.loadAsync(fixture("sample.docx"));
    zip.remove("word/numbering.xml");
    const rels = await zip.file("word/_rels/document.xml.rels")!.async("text");
    zip.file("word/_rels/document.xml.rels", rels.replace(/<Relationship [^>]*numbering\.xml"\/>/, ""));
    const ct = await zip.file("[Content_Types].xml")!.async("text");
    zip.file("[Content_Types].xml", ct.replace(/<Override PartName="\/word\/numbering\.xml"[^>]*\/>/, ""));

    const d = await DocxDocument.load(await zip.generateAsync({ type: "uint8array" }));
    expect(d.numberingXml()).toBeNull();
    const numId = d.ensureList("ordered");
    expect(d.listFormat(numId, 0)).toBe("decimal");
    expect(d.listFormat(numId, 1)).toBe("lowerLetter");

    const out = await JSZip.loadAsync(await d.save());
    expect(out.file("word/numbering.xml")).not.toBeNull();
    expect(await out.file("[Content_Types].xml")!.async("text")).toContain('PartName="/word/numbering.xml"');
    expect(await out.file("word/_rels/document.xml.rels")!.async("text")).toMatch(/Target="numbering\.xml"/);
  });
});

describe("ensureParagraphStyles", () => {
  const styleIds = (doc: Document) => Array.from(doc.getElementsByTagNameNS(W, "style")).map((s) => s.getAttributeNS(W, "styleId"));

  it("adds missing built-in heading styles, once", async () => {
    const docx = await DocxDocument.load(fixture("sample.docx"));
    const styles = docx.stylesXml()!;
    for (const s of Array.from(styles.getElementsByTagNameNS(W, "style"))) if (s.getAttributeNS(W, "styleId") === "Heading2") s.remove();
    docx.ensureParagraphStyles(["Heading2", "Heading2", "Heading1", "NotAStyle"]);
    const ids = styleIds(styles);
    expect(ids.filter((id) => id === "Heading2")).toHaveLength(1);
    expect(ids.filter((id) => id === "Heading1")).toHaveLength(1);
    expect(ids).not.toContain("NotAStyle");
    const back = await DocxDocument.load(await docx.save());
    const h2 = Array.from(back.stylesXml()!.getElementsByTagNameNS(W, "style")).find((s) => s.getAttributeNS(W, "styleId") === "Heading2")!;
    expect(h2.getElementsByTagNameNS(W, "name")[0].getAttributeNS(W, "val")).toBe("heading 2");
    expect(h2.getElementsByTagNameNS(W, "outlineLvl")[0].getAttributeNS(W, "val")).toBe("1");
  });
});

describe("comments without paragraph ids", () => {
  it("lists a comment whose paragraph has no w14:paraId as a thread of its own, not a reply", async () => {
    const docx = await DocxDocument.load(await blankDocx());
    const a = docx.addComment("First", "Ann");
    const b = docx.addComment("Second", "Bo");
    const list = docx.comments();
    expect(list.map((c) => [c.id, c.parentId])).toEqual([[a, undefined], [b, undefined]]);
  });
});
