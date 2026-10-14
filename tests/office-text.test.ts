import { describe, expect, it } from "vitest";
import { officeText, searchText } from "../src/infrastructure/office-text";
import { fixture } from "./helpers";
import { sheetXml, xlsx } from "./xlsx-fixtures";

describe("text of Office files, for vault search", () => {
  it("reads a document's paragraphs", async () => {
    const parts = await officeText("docx", fixture("sample.docx"));
    expect(parts[0]).toEqual({ label: "", text: "Doc Heading" });
    expect(parts.some((p) => p.text.startsWith("Plain bold and italic"))).toBe(true);
  });

  it("reads slides in deck order, labelled by slide", async () => {
    const parts = await officeText("pptx", fixture("sample.pptx"));
    expect(parts.slice(0, 2)).toEqual([{ label: "Slide 1", text: "Hello Deck" }, { label: "Slide 1", text: "Subtitle here" }]);
    expect(parts).toContainEqual({ label: "Slide 3", text: "A1" });
  });

  it("reads cells by sheet and address: shared, inline, formula strings and numbers", async () => {
    const bytes = await xlsx({
      sheets: [sheetXml(`<row r="1"><c r="A1" t="inlineStr"><is><t>alpha &amp; beta</t></is></c><c r="B1"><v>42</v></c></row>`), sheetXml(`<row r="2"><c r="C2" t="str"><f>"x"</f><v>gamma</v></c></row>`)],
      names: ["One", "Two"],
    });
    const parts = await officeText("xlsx", bytes);
    expect(parts).toContainEqual({ label: "One!A1", text: "alpha & beta" });
    expect(parts).toContainEqual({ label: "One!B1", text: "42" });
    expect(parts).toContainEqual({ label: "Two!C2", text: "gamma" });
  });

  it("finds matches with a snippet around each", () => {
    const parts = [{ label: "Slide 2", text: "The quick brown fox jumps over the lazy dog" }, { label: "Slide 3", text: "nothing here" }];
    const hits = searchText(parts, "FOX", {});
    expect(hits).toEqual([{ label: "Slide 2", before: "The quick brown ", match: "fox", after: " jumps over the lazy dog" }]);
    expect(searchText(parts, "", {})).toEqual([]);
    const long = searchText([{ label: "", text: `${"a ".repeat(100)}needle${" b".repeat(100)}` }], "needle", {})[0];
    expect(long.before.length).toBeLessThanOrEqual(41);
    expect(long.before.startsWith("…")).toBe(true);
    expect(long.after.endsWith("…")).toBe(true);
  });
});
