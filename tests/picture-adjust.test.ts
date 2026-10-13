import { describe, expect, it } from "vitest";
import { blankPptx } from "../src/infrastructure/blank";
import { PptxDocument } from "../src/infrastructure/pptx/pptx-document";
import { addPicture } from "../src/domain/pptx/slide-edit";
import { parseSlide } from "../src/domain/pptx/slide-model";
import { NO_ADJUST, pictureFilter, readAdjust, setPictureAdjust } from "../src/domain/pptx/picture-adjust";
import { roundTrip } from "./helpers";

const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PNG = Uint8Array.from(atob("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="), (c) => c.charCodeAt(0));

async function picture() {
  const doc = await PptxDocument.load(await blankPptx());
  const path = doc.slides()[0].path;
  const pic = addPicture(doc.xml(path)!, { x: 0, y: 0, w: 1000000, h: 1000000 }, doc.addImage(path, PNG, "png"), "Photo");
  const model = () => parseSlide(doc, path).shapes.find((s) => s.el === pic)!;
  const blip = () => pic.getElementsByTagNameNS(A, "blip")[0];
  return { doc, path, pic, model, blip };
}
const kids = (el: Element) => Array.from(el.children).map((c) => c.localName);

describe("picture adjustments", () => {
  it("reads no adjustments from a plain picture", async () => {
    const { model } = await picture();
    expect(model().adjust).toEqual(NO_ADJUST);
  });

  it("writes transparency, brightness and contrast as alphaModFix and lum, and reads them back", async () => {
    const { doc, path, pic, blip } = await picture();
    setPictureAdjust(pic, { ...NO_ADJUST, transparency: 0.25, brightness: 0.2, contrast: -0.1 });
    expect(kids(blip())).toEqual(["alphaModFix", "lum"]);
    expect(blip().getElementsByTagNameNS(A, "alphaModFix")[0].getAttribute("amt")).toBe("75000");
    expect([blip().getElementsByTagNameNS(A, "lum")[0].getAttribute("bright"), blip().getElementsByTagNameNS(A, "lum")[0].getAttribute("contrast")]).toEqual(["20000", "-10000"]);
    const back = await roundTrip(doc);
    const shape = parseSlide(back, path).shapes.find((s) => s.kind === "pic")!;
    expect(shape.adjust).toMatchObject({ transparency: 0.25, brightness: 0.2, contrast: -0.1, recolor: "none" });
  });

  it("recolors: grayscale, and light / dark accent variants as duotone with theme colors", async () => {
    const { pic, model, blip } = await picture();
    setPictureAdjust(pic, { ...NO_ADJUST, recolor: "gray" });
    expect(kids(blip())).toEqual(["grayscl"]);
    expect(model().adjust!.recolor).toBe("gray");
    setPictureAdjust(pic, { ...NO_ADJUST, recolor: "dark:accent2" });
    const duo = blip().getElementsByTagNameNS(A, "duotone")[0];
    expect(kids(duo)).toEqual(["prstClr", "schemeClr"]);
    expect(model().adjust!.recolor).toBe("dark:accent2");
    expect(model().adjust!.duotone![0]).toBe("#000000");
    setPictureAdjust(pic, { ...NO_ADJUST, recolor: "light:accent1" });
    expect(model().adjust!.recolor).toBe("light:accent1");
    expect(model().adjust!.duotone![1]).toBe("#FFFFFF");
  });

  it("removes everything on reset and keeps other blip children (extLst last)", async () => {
    const { pic, blip } = await picture();
    blip().appendChild(blip().ownerDocument.createElementNS(A, "a:extLst"));
    setPictureAdjust(pic, { ...NO_ADJUST, transparency: 0.5, recolor: "gray" });
    expect(kids(blip())).toEqual(["alphaModFix", "grayscl", "extLst"]);
    setPictureAdjust(pic, NO_ADJUST);
    expect(kids(blip())).toEqual(["extLst"]);
  });

  it("builds an SVG filter only when there is something to do", () => {
    expect(pictureFilter(NO_ADJUST)).toBeNull();
    const f = pictureFilter({ ...NO_ADJUST, recolor: "gray", contrast: 0.5, brightness: 0.1 })!;
    expect(f).toContain('type="saturate" values="0"');
    expect(f).toContain('slope="2"');
    const d = pictureFilter({ ...NO_ADJUST, recolor: "dark:accent1", duotone: ["#000000", "#FF8000"] })!;
    expect(d).toContain('feFuncR type="table" tableValues="0 1"');
    expect(d).toContain('feFuncG type="table" tableValues="0 0.502"');
  });
});

describe("drawing adjusted pictures", () => {
  it("applies opacity and an SVG filter to the picture", async () => {
    const { renderSlide } = await import("../src/presentation/pptx/slide-renderer");
    const { doc, path, pic } = await picture();
    setPictureAdjust(pic, { ...NO_ADJUST, transparency: 0.4, recolor: "gray" });
    const host = document.body.createDiv();
    renderSlide(host, parseSlide(doc, path), doc.slideSize, { width: 960, editable: false, imageUrl: () => "blob:x" });
    const img = host.querySelector<HTMLImageElement>(".pptx-kind-pic img")!;
    expect(img.style.opacity).toBe("0.6");
    const id = /url\("?#([^")]+)"?\)/.exec(img.style.filter)![1];
    const filter = host.querySelector(`filter[id="${id}"]`)!;
    expect(filter.getAttribute("color-interpolation-filters")).toBe("sRGB");
    expect(filter.firstElementChild!.localName).toBe("feColorMatrix");
    host.remove();
  });
});

describe("replacing a picture", () => {
  it("swaps the image, keeping size, crop and adjustments, and drops an SVG version of the old one", async () => {
    const { replacePicture } = await import("../src/domain/pptx/picture-adjust");
    const { doc, path, pic, model, blip } = await picture();
    setPictureAdjust(pic, { ...NO_ADJUST, recolor: "gray" });
    const svg = blip().ownerDocument.createElementNS(A, "a:extLst");
    svg.appendChild(blip().ownerDocument.createElementNS(A, "a:ext")).setAttribute("uri", "{96DAC541-7B7A-43D3-8B79-37D633B846F1}");
    blip().appendChild(svg);
    const before = model().rect;
    const JPG = new Uint8Array([0xff, 0xd8, 0xff, 0xd9]);
    replacePicture(doc, path, pic, JPG, "jpeg");
    expect(model().rect).toEqual(before);
    expect(model().adjust!.recolor).toBe("gray");
    expect(await doc.media(model().image!)).toEqual(JPG);
    expect(blip().getElementsByTagNameNS(A, "ext")).toHaveLength(0);
  });
});
