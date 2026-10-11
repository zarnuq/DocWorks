import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { blankDocx } from "../src/infrastructure/blank";
import { DocxDocument } from "../src/infrastructure/docx/docx-document";
import { deobfuscateFont } from "../src/infrastructure/fonts/odttf";

const KEY = "{01234567-89AB-CDEF-0123-456789ABCDEF}";
// ECMA-376 Part 2 §15.2.2: the key's 32 hex digits read as bytes from the end; the first 32 bytes XOR with it.
const keyBytes = Array.from({ length: 16 }, (_, i) => parseInt("0123456789ABCDEF0123456789ABCDEF".slice(30 - 2 * i, 32 - 2 * i), 16));
const font = Uint8Array.from({ length: 48 }, (_, i) => (i * 7) & 255);
const obfuscated = font.map((b, i) => (i < 32 ? b ^ keyBytes[i % 16] : b));

describe("fonts embedded in Word documents", () => {
  it("undoes Word's font obfuscation", () => {
    expect(Array.from(deobfuscateFont(obfuscated, KEY))).toEqual(Array.from(font));
  });

  it("lists the document's embedded fonts by family and style, ready to load", async () => {
    const zip = await JSZip.loadAsync(await blankDocx());
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    zip.file("word/fontTable.xml", `<w:fonts xmlns:w="${W}" xmlns:r="${R}"><w:font w:name="Fancy Sans"><w:embedRegular r:id="rId1" w:fontKey="${KEY}"/><w:embedBold r:id="rId2" w:fontKey="${KEY}"/></w:font><w:font w:name="Plain"/></w:fonts>`);
    zip.file("word/_rels/fontTable.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/font" Target="fonts/font1.odttf"/><Relationship Id="rId2" Type="${R}/font" Target="fonts/font2.odttf"/></Relationships>`);
    zip.file("word/fonts/font1.odttf", obfuscated);
    zip.file("word/fonts/font2.odttf", obfuscated);
    const rels = await zip.file("word/_rels/document.xml.rels")!.async("string");
    zip.file("word/_rels/document.xml.rels", rels.replace("</Relationships>", `<Relationship Id="rIdFonts" Type="${R}/fontTable" Target="fontTable.xml"/></Relationships>`));
    const doc = await DocxDocument.load(await zip.generateAsync({ type: "uint8array" }));
    const fonts = await doc.embeddedFonts();
    expect(fonts.map((f) => [f.family, f.bold, f.italic])).toEqual([["Fancy Sans", false, false], ["Fancy Sans", true, false]]);
    expect(Array.from(fonts[0].data)).toEqual(Array.from(font));
  });

  it("reads the page count Word saved, when it saved one", async () => {
    const zip = await JSZip.loadAsync(await blankDocx());
    expect((await DocxDocument.load(await zip.generateAsync({ type: "uint8array" }))).wordPages()).toBeNull();
    zip.file("docProps/app.xml", `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Pages>3</Pages></Properties>`);
    const rels = await zip.file("_rels/.rels")!.async("string");
    zip.file("_rels/.rels", rels.replace("</Relationships>", `<Relationship Id="rIdApp" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`));
    expect((await DocxDocument.load(await zip.generateAsync({ type: "uint8array" }))).wordPages()).toBe(3);
  });
});

describe("Word's saved page breaks", () => {
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const el = (xml: string) => new DOMParser().parseFromString(xml.replace("<w:p>", `<w:p xmlns:w="${W}">`).replace("<w:tbl>", `<w:tbl xmlns:w="${W}">`), "application/xml").documentElement;
  it("knows a paragraph or table started a page when Word's marker comes before its text", async () => {
    const { startsWordPage } = await import("../src/domain/docx/docx-reader");
    expect(startsWordPage(el(`<w:p><w:r><w:lastRenderedPageBreak/><w:t>Top of page 2</w:t></w:r></w:p>`))).toBe(true);
    expect(startsWordPage(el(`<w:p><w:r><w:t>Runs over</w:t><w:lastRenderedPageBreak/><w:t> onto page 2</w:t></w:r></w:p>`))).toBe(false);
    expect(startsWordPage(el(`<w:p><w:r><w:t>No break</w:t></w:r></w:p>`))).toBe(false);
    expect(startsWordPage(el(`<w:tbl><w:tr><w:tc><w:p><w:r><w:lastRenderedPageBreak/><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`))).toBe(true);
  });
});
