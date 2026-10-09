import { describe, it, expect } from "vitest";
import { parseSlide } from "../src/domain/pptx/slide-model";
import { writeParagraphs } from "../src/domain/pptx/slide-edit";
import { renderSlide } from "../src/presentation/pptx/slide-renderer";
import { readParagraphs } from "../src/presentation/pptx/text-html";
import { loadSample } from "./helpers";

const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const P = "http://schemas.openxmlformats.org/presentationml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const ASVG = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";

async function slideWith(spXml: string) {
  const doc = await loadSample();
  const path = doc.slides()[1].path;
  const xml = doc.xml(path)!;
  const tree = xml.getElementsByTagNameNS(P, "spTree")[0];
  const frag = new DOMParser().parseFromString(`<x xmlns:p="${P}" xmlns:a="${A}" xmlns:r="${R}" xmlns:asvg="${ASVG}">${spXml}</x>`, "application/xml").documentElement.firstElementChild!;
  tree.appendChild(xml.importNode(frag, true));
  const rels = doc.xml(path.replace("slides/", "slides/_rels/") + ".rels")!;
  for (const [id, target] of [["rIdPng", "../media/icon.png"], ["rIdSvg", "../media/icon.svg"]]) {
    const rel = rels.documentElement.appendChild(rels.createElementNS("http://schemas.openxmlformats.org/package/2006/relationships", "Relationship"));
    rel.setAttribute("Id", id);
    rel.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
    rel.setAttribute("Target", target);
  }
  return { doc, path, model: () => parseSlide(doc, path) };
}
const shape = (body: string, name = "Icon") =>
  `<p:sp><p:nvSpPr><p:cNvPr id="900" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${body}</p:spPr></p:sp>`;

describe("pictures and SmartArt icons", () => {
  it("draws a shape filled with a picture (SmartArt icons are shapes like this)", async () => {
    const { model } = await slideWith(shape(`<a:blipFill><a:blip r:embed="rIdPng"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>`));
    expect(model().shapes.find((s) => s.name === "Icon")!.image).toBe("ppt/media/icon.png");
  });

  it("prefers the SVG version of a picture, and uses it when there is no PNG at all", async () => {
    const svgOnly = `<a:blipFill><a:blip><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip r:embed="rIdSvg"/></a:ext></a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch></a:blipFill>`;
    const { model } = await slideWith(shape(svgOnly));
    expect(model().shapes.find((s) => s.name === "Icon")!.image).toBe("ppt/media/icon.svg");
    const both = svgOnly.replace("<a:blip>", `<a:blip r:embed="rIdPng">`);
    const again = await slideWith(shape(both));
    expect(again.model().shapes.find((s) => s.name === "Icon")!.image).toBe("ppt/media/icon.svg");
  });
});

describe("all caps text", () => {
  const capsBox = `<p:sp><p:nvSpPr><p:cNvPr id="901" name="Caps"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3000000" cy="500000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr cap="all"/></a:lvl1pPr></a:lstStyle><a:p><a:r><a:rPr lang="en-US"/><a:t>Decisions </a:t></a:r><a:r><a:rPr lang="en-US" cap="none"/><a:t>whether</a:t></a:r></a:p></p:txBody></p:sp>`;

  it("reads cap from the run and its style, and shows it in capitals", async () => {
    const { model, doc } = await slideWith(capsBox);
    const s = model().shapes.find((x) => x.name === "Caps")!;
    expect(s.text!.paras[0].runs.map((r) => r.cap)).toEqual(["all", "none"]);
    const host = document.createElement("div");
    const rendered = renderSlide(host, model(), doc.slideSize, { width: 960, editable: true, imageUrl: () => null });
    const text = rendered.find((r) => r.shape.name === "Caps")!.textEl!;
    const caps = (el: Element | null) => el && getComputedStyle(el).textTransform;
    const spans = Array.from(text.querySelectorAll("span"));
    expect(caps(spans.find((x) => x.textContent === "whether")!)).toBe("none");
    expect(caps(text.querySelector("p"))).toBe("uppercase");
  });

  it("keeps caps through an edit round trip, and writes nothing when unchanged", async () => {
    const { model, doc } = await slideWith(capsBox);
    const s = model().shapes.find((x) => x.name === "Caps")!;
    const host = document.createElement("div");
    const rs = renderSlide(host, model(), doc.slideSize, { width: 960, editable: true, imageUrl: () => null }).find((r) => r.shape.name === "Caps")!;
    const before = new XMLSerializer().serializeToString(s.text!.el);
    writeParagraphs(s.text!, readParagraphs(rs.textEl!, rs.pxPerPt!, (lvl) => s.text!.paras[0].defaults && { ...s.text!.paras[lvl]?.defaults ?? s.text!.paras[0].defaults }));
    expect(new XMLSerializer().serializeToString(s.text!.el)).toBe(before);
  });
});

describe("placeholder prompts", () => {
  it("never shows 'Click to add text' in an empty slide number, date or footer placeholder", async () => {
    const { model, doc } = await slideWith(
      `<p:sp><p:nvSpPr><p:cNvPr id="902" name="Slide Number"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldNum" sz="quarter" idx="12"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`
    );
    const host = document.createElement("div");
    const rs = renderSlide(host, model(), doc.slideSize, { width: 960, editable: true, imageUrl: () => null }).find((r) => r.shape.name === "Slide Number")!;
    expect(rs.textEl!.hasAttribute("data-prompt")).toBe(false);
  });

  it("still shows it in an empty title or body placeholder", async () => {
    const { model, doc } = await slideWith(
      `<p:sp><p:nvSpPr><p:cNvPr id="903" name="Body"/><p:cNvSpPr/><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`
    );
    const host = document.createElement("div");
    const rs = renderSlide(host, model(), doc.slideSize, { width: 960, editable: true, imageUrl: () => null }).find((r) => r.shape.name === "Body")!;
    expect(rs.textEl!.hasAttribute("data-prompt")).toBe(true);
    expect(rs.textEl!.classList.contains("is-empty")).toBe(true);
  });
});
