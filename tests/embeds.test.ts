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

  it("shows a document's text, read-only", async () => {
    const el = document.body.createDiv();
    const done = await renderDocxEmbed(el, readFileSync("tests/fixtures/sample.docx"));
    const pm = el.querySelector(".ProseMirror")!;
    expect(pm.getAttribute("contenteditable")).toBe("false");
    expect(pm.textContent!.length).toBeGreaterThan(10);
    done();
  });
});
