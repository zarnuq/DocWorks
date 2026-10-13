import { describe, expect, it } from "vitest";
import { blankDocx } from "../src/infrastructure/blank";
import { DocxDocument } from "../src/infrastructure/docx/docx-document";
import { docxSchema as schema } from "../src/domain/docx/schema";
import { updateStyleToMatch } from "../src/domain/docx/style-edit";
import { DocxStyles } from "../src/domain/docx/docx-styles";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const styleEl = (d: DocxDocument, id: string) => Array.from(d.stylesXml()!.getElementsByTagNameNS(W, "style")).find((s) => s.getAttributeNS(W, "styleId") === id)!;
const kids = (el: Element | undefined) => Array.from(el?.children ?? []).map((c) => c.localName);
const child = (el: Element, name: string) => Array.from(el.children).find((c) => c.localName === name);

describe("update style to match", () => {
  it("writes the paragraph's direct formatting into its style, in schema order", async () => {
    const d = await DocxDocument.load(await blankDocx());
    d.ensureParagraphStyles(["Heading1"]);
    const m = schema.marks;
    const para = schema.nodes.paragraph.create(
      { style: "Heading1", align: "center", spaceBefore: 18, spaceAfter: 6 },
      schema.text("Title", [m.italic.create(), m.font.create({ font: "Georgia" }), m.size.create({ pt: 20 }), m.color.create({ color: "#FF0000" })])
    );
    updateStyleToMatch(d, "Heading1", para);
    const st = styleEl(d, "Heading1");
    const pPr = child(st, "pPr")!, rPr = child(st, "rPr")!;
    expect(kids(pPr)).toEqual(expect.arrayContaining(["keepNext", "spacing", "jc", "outlineLvl"]));
    expect(kids(pPr).indexOf("jc")).toBeGreaterThan(kids(pPr).indexOf("spacing"));
    expect(child(pPr, "jc")!.getAttributeNS(W, "val")).toBe("center");
    expect([child(pPr, "spacing")!.getAttributeNS(W, "before"), child(pPr, "spacing")!.getAttributeNS(W, "after")]).toEqual(["360", "120"]);
    expect(kids(rPr)).toEqual(["rFonts", "b", "i", "color", "sz", "szCs"].filter((k) => kids(rPr).includes(k)));
    expect(child(rPr, "rFonts")!.getAttributeNS(W, "ascii")).toBe("Georgia");
    expect(child(rPr, "sz")!.getAttributeNS(W, "val")).toBe("40");
    expect(child(rPr, "b")).toBeTruthy(); // the style's own bold stays
    const styles = new DocxStyles(d.stylesXml(), d.numberingXml());
    expect(styles.fontFor("Heading1", null)).toBe("Georgia");
    expect(styles.sizeFor("Heading1", null)).toBe(20);
  });

  it("creates pPr / rPr in the style's child order when the style has none", async () => {
    const d = await DocxDocument.load(await blankDocx());
    const normal = styleEl(d, "Normal");
    for (const c of Array.from(normal.children)) if (c.localName === "pPr" || c.localName === "rPr") c.remove();
    updateStyleToMatch(d, "Normal", schema.nodes.paragraph.create({ align: "right" }, schema.text("x", [schema.marks.bold.create()])));
    const names = kids(normal);
    expect(names.indexOf("pPr")).toBeLessThan(names.indexOf("rPr"));
    expect(names.indexOf("pPr")).toBeGreaterThan(names.indexOf("name"));
    expect(child(child(normal, "rPr")!, "b")).toBeTruthy();
  });
});
