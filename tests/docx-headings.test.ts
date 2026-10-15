import { describe, expect, it } from "vitest";
import { docxSchema as s } from "../src/domain/docx/schema";
import { findHeading, headingLinkText, headingSection } from "../src/domain/docx/headings";

const p = (text: string, style: string | null = null) => s.nodes.paragraph.create({ style }, text ? s.text(text) : undefined);
const doc = () =>
  s.nodes.doc.create(null, [
    p("Intro"),
    p("1. Background", "Heading1"),
    p("Some text"),
    p("Sites", "Heading2"),
    p("North ridge"),
    p("2. Method: tools [v2]", "Heading1"),
    p("Steps"),
  ]);

describe("headings in links", () => {
  it("writes a heading as Obsidian link text (no # | ^ : [ ] characters)", () => {
    expect(headingLinkText("2. Method: tools [v2]")).toBe("2. Method tools v2");
    expect(headingLinkText("  A   #b|c^d  ")).toBe("A b c d");
  });

  it("finds a heading from #Heading or #Parent#Child, ignoring case and dropped characters", () => {
    const d = doc();
    expect(findHeading(d, "#sites")?.text).toBe("Sites");
    expect(findHeading(d, "#1. Background#Sites")?.text).toBe("Sites");
    expect(findHeading(d, "#2. Method tools v2")?.text).toBe("2. Method: tools [v2]");
    expect(findHeading(d, "#Nope")).toBeNull();
  });

  it("cuts a heading's section: up to the next heading at the same or a higher level", () => {
    const d = doc();
    const texts = (sub: string) => {
      const r = headingSection(d, sub)!;
      const out: string[] = [];
      d.slice(r.from, r.to).content.forEach((n) => out.push(n.textContent));
      return out;
    };
    expect(texts("#1. Background")).toEqual(["1. Background", "Some text", "Sites", "North ridge"]);
    expect(texts("#Sites")).toEqual(["Sites", "North ridge"]);
    expect(texts("#2. Method tools v2")).toEqual(["2. Method: tools [v2]", "Steps"]);
    expect(headingSection(d, "#Nope")).toBeNull();
  });
});
