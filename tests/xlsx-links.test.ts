import { describe, it, expect } from "vitest";
import JSZip from "jszip";
import { XlsxWorkbook } from "../src/infrastructure/xlsx/xlsx-workbook";
import { Workbook } from "../src/domain/xlsx/workbook-model";
import { Edits } from "../src/domain/xlsx/commands";
import { linksOf, setLink } from "../src/domain/xlsx/links";
import { xlsx } from "./xlsx-fixtures";

const setup = async () => new Edits(Workbook.open(await XlsxWorkbook.load(await xlsx())));

describe("cell links", () => {
  it("adds a web link as a native hyperlink (external relationship), saves and reopens it, and removes it", async () => {
    const e = await setup();
    const path = e.wb.sheets[0].ref.path;
    e.structural(() => setLink(e.wb.pkg, path, "B2", "https://example.com/a"));
    expect([...linksOf(e.wb.pkg, path)]).toEqual([["B2", "https://example.com/a"]]);
    const bytes = await e.wb.pkg.save();
    const zip = await JSZip.loadAsync(bytes);
    const sheetXml = await zip.file(path)!.async("text");
    expect(sheetXml).toMatch(/<hyperlinks><hyperlink ref="B2" r:id="rId\d+"\/><\/hyperlinks>/);
    // Schema order: hyperlinks come before pageMargins.
    if (sheetXml.includes("<pageMargins")) expect(sheetXml.indexOf("<hyperlinks>")).toBeLessThan(sheetXml.indexOf("<pageMargins"));
    const rels = await zip.file(path.replace(/worksheets\//, "worksheets/_rels/") + ".rels")!.async("text");
    expect(rels).toContain('Target="https://example.com/a" TargetMode="External"');
    const back = new Edits(Workbook.open(await XlsxWorkbook.load(bytes)));
    expect(linksOf(back.wb.pkg, path).get("B2")).toBe("https://example.com/a");
    e.structural(() => setLink(e.wb.pkg, path, "B2", null));
    expect(linksOf(e.wb.pkg, path).size).toBe(0);
    expect(new XMLSerializer().serializeToString(e.wb.pkg.xml(path)!)).not.toContain("hyperlinks");
    e.undo();
    expect(linksOf(e.wb.pkg, path).get("B2")).toBe("https://example.com/a");
  });

  it("links to a place in the workbook with a location, and replaces an existing link", async () => {
    const e = await setup();
    const path = e.wb.sheets[0].ref.path;
    e.structural(() => setLink(e.wb.pkg, path, "A1", "https://old.example"));
    e.structural(() => setLink(e.wb.pkg, path, "A1", "#'Sheet 2'!C3"));
    expect([...linksOf(e.wb.pkg, path)]).toEqual([["A1", "#'Sheet 2'!C3"]]);
    expect(new XMLSerializer().serializeToString(e.wb.pkg.xml(path)!)).toContain(`location="'Sheet 2'!C3"`);
  });
});
