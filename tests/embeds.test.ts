import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { renderDocxEmbed, renderPptxEmbed, renderXlsxEmbed, slideIndex } from "../src/presentation/common/embeds";
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

  it("shows a document's text, read-only", async () => {
    const el = document.body.createDiv();
    const done = await renderDocxEmbed(el, readFileSync("tests/fixtures/sample.docx"));
    const pm = el.querySelector(".ProseMirror")!;
    expect(pm.getAttribute("contenteditable")).toBe("false");
    expect(pm.textContent!.length).toBeGreaterThan(10);
    done();
  });
});
