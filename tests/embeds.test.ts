import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { renderDocxEmbed, renderPptxEmbed, renderXlsxEmbed, sheetTarget, slideIndex } from "../src/presentation/common/embeds";
import { blankPptx } from "../src/infrastructure/blank";
import { sheetXml, xlsx } from "./xlsx-fixtures";

describe("embeds in notes", () => {
  it("picks the slide from #n, falling back to the first", () => {
    expect(slideIndex("#3", 5)).toBe(2);
    expect(slideIndex("", 5)).toBe(0);
    expect(slideIndex("#9", 5)).toBe(0);
    expect(slideIndex("#x", 5)).toBe(0);
  });

  it("shows a deck's slide with previous / next buttons", async () => {
    const el = document.body.createDiv();
    const done = await renderPptxEmbed(el, readFileSync("tests/fixtures/sample.pptx"), "#2");
    expect(el.querySelector(".dw-embed-nav span")!.textContent).toMatch(/^2 \/ \d+$/);
    expect(el.querySelector(".dw-embed-slide")!.childElementCount).toBeGreaterThan(0);
    (el.querySelector("button[aria-label='Previous slide']") as HTMLButtonElement).click();
    expect(el.querySelector(".dw-embed-nav span")!.textContent).toMatch(/^1 \//);
    done();
    const blank = document.body.createDiv();
    (await renderPptxEmbed(blank, await blankPptx(), ""))();
  });

  it("draws the named sheet, with tabs for the others", async () => {
    const bytes = await xlsx({ sheets: [sheetXml(`<row r="1"><c r="A1"><v>1</v></c></row>`), sheetXml(`<row r="1"><c r="B2" t="inlineStr"><is><t>two</t></is></c></row>`)], names: ["One", "Two"] });
    const el = document.body.createDiv();
    const done = await renderXlsxEmbed(el, bytes, "#Two");
    expect(Array.from(el.querySelectorAll(".dw-embed-tabs button")).map((b) => [b.textContent, b.classList.contains("is-active")])).toEqual([["One", false], ["Two", true]]);
    expect(el.querySelector(".dw-embed-sheet .dw-embed-grid canvas")).not.toBeNull();
    done();
  });

  it("reports the sheet a tab shows, so opening the file goes there", async () => {
    const bytes = await xlsx({ sheets: [sheetXml(""), sheetXml("")], names: ["One", "100% two"] });
    const el = document.body.createDiv();
    const places: string[] = [];
    const done = await renderXlsxEmbed(el, bytes, "", (p) => places.push(p));
    (el.querySelectorAll(".dw-embed-tabs button")[1] as HTMLButtonElement).click();
    expect(places).toEqual(["#100%25%20two"]);
    expect(sheetTarget(places[0], ["One", "100% two"])).toEqual({ sheet: "100% two", range: null });
    done();
  });

  it("reports the slide shown after previous / next", async () => {
    const el = document.body.createDiv();
    const places: string[] = [];
    const done = await renderPptxEmbed(el, readFileSync("tests/fixtures/sample.pptx"), "#2", (p) => places.push(p));
    (el.querySelector("button[aria-label='Previous slide']") as HTMLButtonElement).click();
    expect(places).toEqual(["#1"]);
    done();
  });

  it("reads a sheet and range from #Sheet!A1:B2, #A1:B2 or #Sheet", () => {
    const names = ["One", "My Sheet", "A1"];
    expect(sheetTarget("#One!B2:C3", names)).toEqual({ sheet: "One", range: { r1: 1, c1: 1, r2: 2, c2: 2 } });
    expect(sheetTarget("#'My Sheet'!A1:A4", names)).toEqual({ sheet: "My Sheet", range: { r1: 0, c1: 0, r2: 3, c2: 0 } });
    expect(sheetTarget("#My%20Sheet!B2", names)).toEqual({ sheet: "My Sheet", range: { r1: 1, c1: 1, r2: 1, c2: 1 } });
    expect(sheetTarget("#C1:D2", names)).toEqual({ sheet: null, range: { r1: 0, c1: 2, r2: 1, c2: 3 } });
    // A sheet whose name looks like a cell is still the sheet.
    expect(sheetTarget("#A1", names)).toEqual({ sheet: "A1", range: null });
    expect(sheetTarget("#My Sheet", names)).toEqual({ sheet: "My Sheet", range: null });
    expect(sheetTarget("", names)).toEqual({ sheet: null, range: null });
  });

  it("draws just the range asked for, without tabs, under a caption", async () => {
    const bytes = await xlsx({ sheets: [sheetXml(`<row r="1"><c r="A1"><v>1</v></c></row>`), sheetXml(`<row r="2"><c r="B2" t="inlineStr"><is><t>two</t></is></c></row>`)], names: ["One", "Two"] });
    const el = document.body.createDiv();
    const done = await renderXlsxEmbed(el, bytes, "#Two!B2:C3");
    expect(el.querySelector(".dw-embed-tabs")).toBeNull();
    expect(el.querySelector(".dw-embed-range")!.textContent).toBe("Two!B2:C3");
    expect(el.querySelector(".dw-embed-sheet .dw-embed-grid canvas")).not.toBeNull();
    done();
  });

  it("draws a table named by #Table, wherever it is, and follows the Obsidian dark theme", async () => {
    const S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const bytes = await xlsx({
      sheets: [sheetXml(`<row r="1"><c r="A1"><v>1</v></c></row>`), sheetXml(`<row r="2"><c r="B2" t="inlineStr"><is><t>Name</t></is></c></row>`)],
      names: ["One", "Two"],
      parts: {
        "xl/worksheets/_rels/sheet2.xml.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>',
        "xl/tables/table1.xml": `<table xmlns="${S}" id="1" name="Tabella1" displayName="Tabella1" ref="B2:C5"><tableColumns count="2"><tableColumn id="1" name="Name"/><tableColumn id="2" name="Amount"/></tableColumns></table>`,
      },
    });
    document.body.addClass("theme-dark");
    const el = document.body.createDiv();
    const done = await renderXlsxEmbed(el, bytes, "#Tabella1");
    document.body.removeClass("theme-dark");
    expect(el.querySelector(".dw-embed-range")!.textContent).toBe("Two!B2:C5");
    expect(el.querySelector(".dw-embed-sheet")!.classList.contains("xlsx-theme-dark")).toBe(true);
    done();
  });

  it("widens a whole-sheet embed so text running past the last used column isn't cut off", async () => {
    const { Workbook } = await import("../src/domain/xlsx/workbook-model");
    const { XlsxWorkbook } = await import("../src/infrastructure/xlsx/xlsx-workbook");
    const { embedRange } = await import("../src/presentation/common/embeds");
    const long = "a long line of text that runs well past its own column";
    const bytes = await xlsx({ sheets: [sheetXml(`<row r="1"><c r="A1"><v>1</v></c><c r="B1" t="inlineStr"><is><t>${long}</t></is></c></row><row r="2"><c r="B2"><v>2</v></c></row>`)] });
    const sheet = Workbook.open(await XlsxWorkbook.load(bytes)).sheets[0];
    const r = embedRange(sheet, null)!;
    expect(r.c2).toBeGreaterThan(1);
    expect(sheet.colX(r.c2 + 1)).toBeGreaterThanOrEqual(sheet.colX(1) + 200);
    // An explicit range is shown as asked.
    expect(embedRange(sheet, { r1: 0, c1: 0, r2: 1, c2: 1 })).toEqual({ r1: 0, c1: 0, r2: 1, c2: 1 });
  });

  it("shows just a heading's section for #Heading, or the whole document when there's no such heading", async () => {
    const { noteToDocx } = await import("../src/domain/docx/from-html");
    const { blankDocx } = await import("../src/infrastructure/blank");
    const html = document.createElement("div");
    html.innerHTML = "<p>Intro</p><h1>Results</h1><p>Clearer water</p><h2>Detail</h2><p>Sites</p><h1>Next</h1><p>Permits</p>";
    const bytes = await noteToDocx(await blankDocx(), html, () => null);
    const text = async (subpath: string) => {
      const el = document.body.createDiv();
      const done = await renderDocxEmbed(el, bytes, subpath);
      const out = Array.from(el.querySelectorAll(".ProseMirror > *")).map((p) => p.textContent);
      done();
      return out;
    };
    expect(await text("#Results")).toEqual(["Results", "Clearer water", "Detail", "Sites"]);
    expect(await text("#Missing")).toEqual(["Intro", "Results", "Clearer water", "Detail", "Sites", "Next", "Permits"]);
  });

  it("shows a document's text, read-only", async () => {
    const el = document.body.createDiv();
    const done = await renderDocxEmbed(el, readFileSync("tests/fixtures/sample.docx"));
    const pm = el.querySelector(".ProseMirror")!;
    expect(pm.getAttribute("contenteditable")).toBe("false");
    expect(pm.textContent!.length).toBeGreaterThan(10);
    done();
  });
});
