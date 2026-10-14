import { describe, expect, it } from "vitest";
import { XlsxWorkbook } from "../src/infrastructure/xlsx/xlsx-workbook";
import { Workbook } from "../src/domain/xlsx/workbook-model";
import { Edits } from "../src/domain/xlsx/commands";
import { rangeToMarkdown } from "../src/domain/xlsx/csv";
import { deckToMarkdown } from "../src/domain/pptx/outline";
import { loadSample } from "./helpers";
import { xlsx } from "./xlsx-fixtures";

describe("Copy as Markdown table", () => {
  it("turns the selected cells into a table: first row as header, numbers right-aligned, pipes escaped", async () => {
    const e = new Edits(Workbook.open(await XlsxWorkbook.load(await xlsx())));
    const s = e.wb.sheets[0];
    for (const [r, c, v] of [[30, 0, "Item"], [30, 1, "Cost"], [31, 0, "a|b"], [31, 1, "12"], [32, 0, "two\nlines"], [32, 1, "3.5"]] as const) e.input(s, r, c, v);
    expect(rangeToMarkdown(s, { r1: 30, c1: 0, r2: 32, c2: 1 })).toBe(
      ["| Item | Cost |", "| --- | ---: |", "| a\\|b | 12 |", "| two<br>lines | 3.5 |", ""].join("\n")
    );
  });

  it("stops at the used area for whole columns and gives nothing for empty cells", async () => {
    const e = new Edits(Workbook.open(await XlsxWorkbook.load(await xlsx())));
    const s = e.wb.sheets[0];
    const md = rangeToMarkdown(s, { r1: 0, c1: 0, r2: 1048575, c2: 0 });
    // Rows 1–6 are used and row 6 is hidden: five rows plus the separator line.
    expect(s.used()!.r2).toBe(5);
    expect(md.trim().split("\n").length).toBe(6);
    expect(md).not.toContain("x!");
    expect(rangeToMarkdown(s, { r1: 900, c1: 40, r2: 905, c2: 41 })).toBe("");
  });
});

describe("Copy as Markdown outline", () => {
  it("lists each slide's title as a heading and its text as nested bullets, with notes quoted", async () => {
    const doc = await loadSample();
    doc.setNotes(1, "Say hello\nthen bullets");
    const md = deckToMarkdown(doc);
    expect(md.startsWith("## Hello Deck\n\n- Subtitle here\n")).toBe(true);
    expect(md).toContain("## Bullets\n\n- First point\n  - Bold red italic\n\n> Say hello\n> then bullets\n");
    // A slide without a title is numbered; its table is a Markdown table.
    expect(md).toContain("## Slide 3\n");
    expect(md).toContain("| A1 | B1 |\n| --- | --- |\n| A2 | B2 |\n");
    expect(md).not.toContain("- A1");
  });

  it("takes a slide's clearly largest text as its title when it has no title placeholder", async () => {
    const doc = await loadSample();
    const xml = doc.xml(doc.slides()[2].path)!;
    const t = Array.from(xml.getElementsByTagNameNS("*", "t")).find((e) => e.textContent === "Box text")!;
    const r = t.parentElement!;
    let rPr = Array.from(r.children).find((c) => c.localName === "rPr");
    if (!rPr) rPr = r.insertBefore(xml.createElementNS("http://schemas.openxmlformats.org/drawingml/2006/main", "a:rPr"), t);
    rPr.setAttribute("sz", "4000");
    const md = deckToMarkdown(doc);
    expect(md).toContain("## Box text\n");
    expect(md).not.toContain("- Box text");
  });
});
