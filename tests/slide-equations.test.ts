import { describe, expect, it } from "vitest";
import { PptxDocument } from "../src/infrastructure/pptx/pptx-document";
import { blankPptx } from "../src/infrastructure/blank";
import { parseSlide, type ShapeModel } from "../src/domain/pptx/slide-model";
import { appendShape } from "../src/domain/pptx/slide-edit";
import { itemRect, setItemRect, topLevel } from "../src/domain/pptx/arrange";
import { A14, equationLatex, insertEquation, setEquation, shownBranch, slideOmml, syncEquationFallbacks } from "../src/domain/pptx/equation";
import { renderParagraphs } from "../src/presentation/pptx/text-html";
import { NS } from "../src/core/utils/xml";

const MC = NS.mc;
const M = NS.m;

/** An equation text box as PowerPoint writes it: the math in the a14 Choice, a picture in the Fallback. */
const POWERPOINT_EQUATION =
  `<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice xmlns:a14="${A14}" Requires="a14">` +
  `<p:sp><p:nvSpPr><p:cNvPr id="40" name="TextBox 39"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
  `<p:spPr><a:xfrm><a:off x="1000000" y="2000000"/><a:ext cx="3000000" cy="600000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>` +
  `<p:txBody><a:bodyPr wrap="none"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p>` +
  `<a14:m><m:oMathPara xmlns:m="${M}"><m:oMathParaPr><m:jc m:val="centerGroup"/></m:oMathParaPr><m:oMath>` +
  `<m:sSup><m:e><m:r><a:rPr lang="en-US" i="1"><a:latin typeface="Cambria Math"/></a:rPr><m:t>x</m:t></m:r></m:e><m:sup><m:r><a:rPr lang="en-US" i="1"><a:latin typeface="Cambria Math"/></a:rPr><m:t>2</m:t></m:r></m:sup></m:sSup>` +
  `</m:oMath></m:oMathPara></a14:m><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp></mc:Choice>` +
  `<mc:Fallback><p:sp><p:nvSpPr><p:cNvPr id="40" name="TextBox 39"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
  `<p:spPr><a:xfrm><a:off x="1000000" y="2000000"/><a:ext cx="3000000" cy="600000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:blipFill><a:blip r:embed="rId99"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></p:spPr>` +
  `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"><a:noFill/></a:rPr><a:t> </a:t></a:r></a:p></p:txBody></p:sp></mc:Fallback></mc:AlternateContent>`;

async function deck(shapeXml = POWERPOINT_EQUATION) {
  const doc = await PptxDocument.load(await blankPptx());
  const path = doc.slides()[0].path;
  const slide = doc.xml(path)!;
  if (shapeXml) appendShape(slide, shapeXml);
  return { doc, path, slide, model: () => parseSlide(doc, path) };
}
const equationShape = (shapes: ShapeModel[]) => shapes.find((s) => s.equation)!;
const ac = (slide: Document) => slide.getElementsByTagNameNS(MC, "AlternateContent")[0];
const branch = (slide: Document, name: "Choice" | "Fallback") => Array.from(ac(slide).children).find((c) => c.localName === name)!.firstElementChild!;

describe("reading PowerPoint equations", () => {
  it("draws the a14 Choice (the math), not the fallback picture", async () => {
    const { model, slide } = await deck();
    const shape = equationShape(model().shapes);
    expect(shape).toBeDefined();
    expect(shape.el).toBe(branch(slide, "Choice"));
    expect(shape.kind).toBe("sp");
    expect(shape.fill).toBeNull();
    expect(model().shapes.filter((s) => s.name === "TextBox 39")).toHaveLength(1);
  });

  it("the math is a run of its paragraph", async () => {
    const { model } = await deck();
    const runs = equationShape(model().shapes).text!.paras[0].runs;
    expect(runs).toHaveLength(1);
    expect(runs[0].math).toMatch(/^<m:oMathPara/);
    expect(runs[0].text).toBe("");
  });

  it("equationLatex reads the LaTeX", async () => {
    const { slide } = await deck();
    expect(equationLatex(branch(slide, "Choice"))).toBe("x^2");
  });

  it("other AlternateContent (no math) still shows its Fallback", async () => {
    const other = POWERPOINT_EQUATION.replace(/<a14:m>[\s\S]*<\/a14:m>/, `<a:r><a:t>p14 thing</a:t></a:r>`).replace('Requires="a14"', 'Requires="p14"').replace('xmlns:a14', "xmlns:p14");
    const { model, slide } = await deck(other);
    const shape = model().shapes.find((s) => s.name === "TextBox 39")!;
    expect(shape.el).toBe(branch(slide, "Fallback"));
    expect(shape.equation).toBeFalsy();
  });

  it("shownBranch picks the Choice only when it holds a14 math", async () => {
    const { slide } = await deck();
    expect(shownBranch(ac(slide))).toBe(branch(slide, "Choice"));
  });
});

describe("drawing math runs", () => {
  it("renders MathML in the paragraph, not editable", async () => {
    const { model } = await deck();
    const host = document.body.createDiv();
    renderParagraphs(host, equationShape(model().shapes).text!.paras, 1);
    const math = host.querySelector("math");
    expect(math).not.toBeNull();
    expect(math!.closest<HTMLElement>("[data-math]")!.contentEditable).toBe("false");
    expect(host.querySelector("msup")).not.toBeNull();
  });
});

describe("slideOmml (Word's OMML into PowerPoint's)", () => {
  it("swaps w:rPr for DrawingML run properties in Cambria Math, italic unless upright", () => {
    const xml = slideOmml(`<m:oMath xmlns:m="${M}" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><m:r><w:rPr><w:rFonts w:ascii="Cambria Math"/></w:rPr><m:t>x</m:t></m:r><m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>sin</m:t></m:r></m:oMath>`);
    const doc = new DOMParser().parseFromString(xml, "application/xml");
    expect(doc.getElementsByTagNameNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "rPr")).toHaveLength(0);
    const rPrs = Array.from(doc.getElementsByTagNameNS(NS.a, "rPr"));
    expect(rPrs.map((r) => r.getAttribute("i"))).toEqual(["1", "0"]);
    expect(rPrs[0].getElementsByTagNameNS(NS.a, "latin")[0].getAttribute("typeface")).toBe("Cambria Math");
    // a:rPr comes after m:rPr, before m:t.
    const runs = Array.from(doc.getElementsByTagNameNS(M, "r"));
    expect(Array.from(runs[1].children).map((c) => c.localName)).toEqual(["rPr", "rPr", "t"]);
    expect(runs[1].children[0].namespaceURI).toBe(M);
  });
});

describe("editing an equation", () => {
  it("rewrites the Choice's math and keeps it display math", async () => {
    const { slide } = await deck();
    setEquation(branch(slide, "Choice"), "\\frac{a}{b}");
    const choice = branch(slide, "Choice");
    expect(equationLatex(choice)).toBe("\\frac{a}{b}");
    const m = choice.getElementsByTagNameNS(A14, "m")[0];
    expect(m.firstElementChild!.localName).toBe("oMathPara");
    expect(m.getElementsByTagNameNS(M, "f")).toHaveLength(1);
  });

  it("replaces the stale fallback picture with the equation as text, in the same place", async () => {
    const { slide } = await deck();
    setEquation(branch(slide, "Choice"), "a+b");
    const fallback = branch(slide, "Fallback");
    expect(fallback.getElementsByTagNameNS(NS.a, "blipFill")).toHaveLength(0);
    expect(fallback.getElementsByTagNameNS(A14, "m")).toHaveLength(0);
    expect(fallback.textContent).toContain("a+b");
    expect(fallback.getElementsByTagNameNS(NS.a, "off")[0].getAttribute("x")).toBe("1000000");
    expect(fallback.getElementsByTagNameNS(NS.p, "cNvPr")[0].getAttribute("id")).toBe("40");
  });

  it("an empty equation is refused (Delete removes the shape instead)", async () => {
    const { slide } = await deck();
    expect(() => setEquation(branch(slide, "Choice"), "  ")).toThrow();
  });
});

describe("inserting an equation", () => {
  it("adds a PowerPoint equation text box: Choice with a14 math, text Fallback, a fresh id", async () => {
    const { doc, path } = await deck("");
    const slide = doc.xml(path)!;
    const sp = insertEquation(slide, "E=mc^2", { x: 100, y: 200, w: 3000, h: 400 });
    const wrapper = topLevel(sp);
    expect(wrapper.localName).toBe("AlternateContent");
    const choice = Array.from(wrapper.children).find((c) => c.localName === "Choice")!;
    expect(choice.getAttribute("Requires")).toBe("a14");
    expect(sp.parentElement).toBe(choice);
    expect(equationLatex(sp)).toBe("E=mc^2");
    const ids = Array.from(slide.getElementsByTagNameNS(NS.p, "cNvPr")).map((c) => c.getAttribute("id"));
    expect(new Set(ids.slice(0, -1)).size).toBe(ids.length - 1);
    const model = parseSlide(doc, path);
    expect(equationShape(model.shapes).text!.paras[0].runs[0].math).toBeTruthy();
  });

  it("survives saving and reopening, and the XML declares what it uses", async () => {
    const { doc, path } = await deck("");
    insertEquation(doc.xml(path)!, "\\alpha", { x: 0, y: 0, w: 1000, h: 400 });
    const again = await PptxDocument.load(await doc.save());
    const slide = again.xml(path)!;
    expect(equationLatex(shownBranch(ac(slide))!)).toBe("\\alpha");
    expect(equationShape(parseSlide(again, path).shapes)).toBeDefined();
  });
});

describe("moving an equation", () => {
  it("arrange reads and moves the shown shape inside the AlternateContent", async () => {
    const { slide } = await deck();
    const item = ac(slide);
    expect(itemRect(item)).toEqual({ x: 1000000, y: 2000000, w: 3000000, h: 600000 });
    setItemRect(item, { x: 5, y: 6, w: 7, h: 8 });
    expect(itemRect(item)).toEqual({ x: 5, y: 6, w: 7, h: 8 });
    expect(item.children[0].localName).toBe("Choice");
    expect(Array.from(item.children).every((c) => c.namespaceURI === MC)).toBe(true);
  });

  it("syncEquationFallbacks puts a moved equation's fallback in the same place, and nothing else", async () => {
    const { slide } = await deck();
    expect(syncEquationFallbacks(slide)).toBe(false);
    setItemRect(ac(slide), { x: 5, y: 6, w: 7, h: 8 });
    expect(syncEquationFallbacks(slide)).toBe(true);
    const off = branch(slide, "Fallback").getElementsByTagNameNS(NS.a, "off")[0];
    expect([off.getAttribute("x"), off.getAttribute("y")]).toEqual(["5", "6"]);
    // The picture is still right for unchanged math.
    expect(branch(slide, "Fallback").getElementsByTagNameNS(NS.a, "blipFill")).toHaveLength(1);
    expect(syncEquationFallbacks(slide)).toBe(false);
  });

  it("saving the deck syncs them", async () => {
    const { doc, path, slide } = await deck();
    setItemRect(ac(slide), { x: 5, y: 6, w: 7, h: 8 });
    const again = await PptxDocument.load(await doc.save());
    const fb = Array.from(again.xml(path)!.getElementsByTagNameNS(MC, "Fallback"))[0];
    expect(fb.getElementsByTagNameNS(NS.a, "off")[0].getAttribute("x")).toBe("5");
  });
});
