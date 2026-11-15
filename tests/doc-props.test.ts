import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { blankDocx, blankPptx, blankXlsx } from "../src/infrastructure/blank";
import { DocxDocument } from "../src/infrastructure/docx/docx-document";
import { PptxDocument } from "../src/infrastructure/pptx/pptx-document";
import { XlsxWorkbook } from "../src/infrastructure/xlsx/xlsx-workbook";
import { corePropsOf, fromTemplate, readProps, writeProps } from "../src/infrastructure/ooxml/doc-props";

const CORE = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Plan &amp; notes</dc:title><dc:creator>Ada</dc:creator><cp:keywords>alpha, beta</cp:keywords><dcterms:created xsi:type="dcterms:W3CDTF">2024-01-02T03:04:05Z</dcterms:created></cp:coreProperties>`;

async function withCore(bytes: Uint8Array, core: string | null): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes);
  if (core === null) {
    zip.remove("docProps/core.xml");
    const rels = await zip.file("_rels/.rels")!.async("text");
    zip.file("_rels/.rels", rels.replace(/<Relationship[^>]*core-properties[^>]*\/>/, ""));
    const types = await zip.file("[Content_Types].xml")!.async("text");
    zip.file("[Content_Types].xml", types.replace(/<Override[^>]*docProps\/core\.xml[^>]*\/>/, ""));
  } else zip.file("docProps/core.xml", core);
  return zip.generateAsync({ type: "uint8array" });
}

describe("document properties (docProps/core.xml)", () => {
  it("reads title, author, keywords and dates", async () => {
    const docx = await DocxDocument.load(await withCore(await blankDocx(), CORE));
    expect(readProps(docx)).toMatchObject({ title: "Plan & notes", creator: "Ada", keywords: "alpha, beta", subject: "", created: "2024-01-02T03:04:05Z" });
  });

  it("writes edited fields, removes cleared ones and keeps the rest, in every format", async () => {
    for (const [load, blank] of [[DocxDocument.load, blankDocx], [PptxDocument.load, blankPptx], [XlsxWorkbook.load, blankXlsx]] as const) {
      const pkg = await (load as (b: Uint8Array) => Promise<DocxDocument>)(await withCore(await blank(), CORE));
      expect(writeProps(pkg, { title: "New", keywords: "", description: "About" })).toBe(true);
      expect(writeProps(pkg, { title: "New" })).toBe(false);
      const back = await (load as (b: Uint8Array) => Promise<DocxDocument>)(await pkg.save());
      expect(readProps(back)).toMatchObject({ title: "New", keywords: "", description: "About", creator: "Ada", created: "2024-01-02T03:04:05Z" });
    }
  });

  it("creates the part, its relationship and content type when the file has none", async () => {
    const docx = await DocxDocument.load(await withCore(await blankDocx(), null));
    expect(readProps(docx).title).toBe("");
    writeProps(docx, { title: "Made" });
    const zip = await JSZip.loadAsync(await docx.save());
    expect(await zip.file("_rels/.rels")!.async("text")).toContain("core-properties");
    expect(await zip.file("[Content_Types].xml")!.async("text")).toContain("/docProps/core.xml");
    expect(corePropsOf(await zip.file("docProps/core.xml")!.async("text")).title).toBe("Made");
  });

  it("makes a document from a Word, PowerPoint or Excel template with fresh dates", async () => {
    const zip = await JSZip.loadAsync(await withCore(await blankDocx(), CORE));
    const types = await zip.file("[Content_Types].xml")!.async("text");
    zip.file("[Content_Types].xml", types.replace("wordprocessingml.document.main+xml", "wordprocessingml.template.main+xml"));
    const now = new Date("2026-11-14T10:00:00Z");
    const out = await JSZip.loadAsync(await fromTemplate(await zip.generateAsync({ type: "uint8array" }), now));
    expect(await out.file("[Content_Types].xml")!.async("text")).toContain("wordprocessingml.document.main+xml");
    const props = corePropsOf(await out.file("docProps/core.xml")!.async("text"));
    expect(props).toMatchObject({ title: "Plan & notes", created: "2026-11-14T10:00:00Z", modified: "2026-11-14T10:00:00Z" });
  });
});
