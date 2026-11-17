import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { blankDocx, blankPptx, blankXlsx } from "../src/infrastructure/blank";
import { DocxDocument } from "../src/infrastructure/docx/docx-document";
import { PptxDocument } from "../src/infrastructure/pptx/pptx-document";
import { XlsxWorkbook } from "../src/infrastructure/xlsx/xlsx-workbook";
import { kindOfExtension } from "../src/presentation/common/settings";
import { DocxView } from "../src/presentation/docx/docx-view";
import { PptxView } from "../src/presentation/pptx/pptx-view";
import { XlsxView } from "../src/presentation/xlsx/xlsx-view";

/** `bytes` as Office saves a template: the main part's content type is the template one. */
async function asTemplate(bytes: Uint8Array, from: string, to: string): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes);
  const types = await zip.file("[Content_Types].xml")!.async("text");
  expect(types).toContain(from);
  zip.file("[Content_Types].xml", types.replace(from, to));
  return zip.generateAsync({ type: "uint8array" });
}
const typesOf = async (bytes: Uint8Array) => (await JSZip.loadAsync(bytes)).file("[Content_Types].xml")!.async("text");

describe("Office template files (.dotx / .potx / .xltx)", () => {
  it("are the same kind of file as the documents they make", () => {
    expect(["docx", "dotx", "pptx", "potx", "xlsx", "xltx", "DOTX"].map(kindOfExtension)).toEqual(["docx", "docx", "pptx", "pptx", "xlsx", "xlsx", "docx"]);
    expect(kindOfExtension("md")).toBeNull();
    // Macro-enabled files are the same kinds too (their macros are kept, never run).
    expect(["docm", "pptm", "xlsm"].map(kindOfExtension)).toEqual(["docx", "pptx", "xlsx"]);
  });

  it("offers documents and templates, not macro-enabled files, as templates for new files", async () => {
    const { templateKind } = await import("../src/presentation/common/settings");
    // A copy of a macro-enabled file would be a .docx that still says it holds macros: Office refuses it.
    expect(["docx", "dotx", "docm", "potx", "pptm", "XLSM", "xltx", "md"].map(templateKind)).toEqual(["docx", "docx", null, "pptx", null, null, "xlsx", null]);
  });

  it("keep a macro-enabled workbook's macros and content type through an edit", async () => {
    const { Workbook } = await import("../src/domain/xlsx/workbook-model");
    const { Edits } = await import("../src/domain/xlsx/commands");
    const zip = await JSZip.loadAsync(await blankXlsx());
    const types = await zip.file("[Content_Types].xml")!.async("text");
    zip.file("[Content_Types].xml", types.replace("spreadsheetml.sheet.main+xml", "ms-excel.sheet.macroEnabled.main+xml").replace("</Types>", '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>'));
    const vba = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 1, 2, 3, 4]);
    zip.file("xl/vbaProject.bin", vba);
    const rels = await zip.file("xl/_rels/workbook.xml.rels")!.async("text");
    zip.file("xl/_rels/workbook.xml.rels", rels.replace("</Relationships>", '<Relationship Id="rId99" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>'));
    const e = new Edits(Workbook.open(await XlsxWorkbook.load(await zip.generateAsync({ type: "uint8array" }))));
    e.input(e.wb.sheets[0], 0, 0, "42");
    const back = await JSZip.loadAsync(await e.wb.pkg.save());
    expect(Array.from(await back.file("xl/vbaProject.bin")!.async("uint8array"))).toEqual(Array.from(vba));
    expect(await back.file("[Content_Types].xml")!.async("text")).toContain("macroEnabled.main+xml");
    expect(await back.file("xl/_rels/workbook.xml.rels")!.async("text")).toContain("vbaProject.bin");
  });

  it("are searched, embedded and exported like documents", async () => {
    const { OFFICE_KINDS_OF } = await import("../src/presentation/common/office-export");
    expect(OFFICE_KINDS_OF("Plan.DOTX")).toBe("docx");
    expect(OFFICE_KINDS_OF("book.xlsm")).toBe("xlsx");
    expect(OFFICE_KINDS_OF("note.md")).toBeNull();
    const { embedExtensions } = await import("../src/presentation/common/embeds");
    expect(embedExtensions(["docx", "xlsx"])).toEqual([["docx", "docx"], ["dotx", "docx"], ["docm", "docx"], ["xlsx", "xlsx"], ["xltx", "xlsx"], ["xlsm", "xlsx"]]);
  });

  it("open in the editor of their kind", () => {
    const accepts = (view: { prototype: { canAcceptExtension(ext: string): boolean } }, exts: string[]) => exts.map((e) => view.prototype.canAcceptExtension(e));
    expect(accepts(DocxView, ["docx", "dotx", "docm", "potx"])).toEqual([true, true, true, false]);
    expect(accepts(PptxView, ["pptx", "potx", "xltx"])).toEqual([true, true, false]);
    expect(accepts(XlsxView, ["xlsx", "xltx", "dotx"])).toEqual([true, true, false]);
  });

  it("load and save as templates: the content type stays a template's", async () => {
    const dotx = await asTemplate(await blankDocx(), "wordprocessingml.document.main+xml", "wordprocessingml.template.main+xml");
    const docx = await DocxDocument.load(dotx);
    expect(docx.body.localName).toBe("body");
    expect(await typesOf(await docx.save())).toContain("wordprocessingml.template.main+xml");

    const potx = await asTemplate(await blankPptx(), "presentationml.presentation.main+xml", "presentationml.template.main+xml");
    const deck = await PptxDocument.load(potx);
    expect(deck.slides().length).toBeGreaterThan(0);
    expect(await typesOf(await deck.save())).toContain("presentationml.template.main+xml");

    const xltx = await asTemplate(await blankXlsx(), "spreadsheetml.sheet.main+xml", "spreadsheetml.template.main+xml");
    const book = await XlsxWorkbook.load(xltx);
    expect(book.sheets().length).toBe(1);
    expect(await typesOf(await book.save())).toContain("spreadsheetml.template.main+xml");
  });
});
