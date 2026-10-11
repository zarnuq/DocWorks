import { describe, expect, it } from "vitest";
import { blankPptx } from "../src/infrastructure/blank";
import { PptxDocument } from "../src/infrastructure/pptx/pptx-document";
import { addChart, chartData, chartWorkbook } from "../src/domain/pptx/chart";
import { addPicture, addTextBox } from "../src/domain/pptx/slide-edit";
import { itemRect } from "../src/domain/pptx/arrange";
import { parseSlide, plainText, slideColors } from "../src/domain/pptx/slide-model";

const PNG = Uint8Array.from(atob("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="), (c) => c.charCodeAt(0));
const text = (doc: PptxDocument, i: number) => parseSlide(doc, doc.slides()[i].path).shapes.map((s) => (s.text ? plainText(s.text.el) : "")).join("|");

async function source() {
  const doc = await PptxDocument.load(await blankPptx());
  const first = doc.slides()[0].path;
  addTextBox(doc.xml(first)!, { x: 100000, y: 100000, w: 2000000, h: 500000 }, "First slide");
  const index = doc.addSlide(0);
  const second = doc.slides()[index].path;
  addTextBox(doc.xml(second)!, { x: 100000, y: 100000, w: 2000000, h: 500000 }, "With a picture and a chart");
  addPicture(doc.xml(second)!, { x: 0, y: 1000000, w: 1000000, h: 1000000 }, doc.addImage(second, PNG, "png"), "Logo");
  const data = { categories: ["A", "B"], series: [{ name: "S", values: [1, 2] }] };
  addChart(doc, second, data, "barChart", await chartWorkbook(data));
  return doc;
}

describe("copying slides", () => {
  it("pastes a slide into another deck with its picture, chart and embedded workbook", async () => {
    const from = await source();
    const to = await PptxDocument.load(await blankPptx());
    const clip = await from.copySlides([1]);
    const at = to.pasteSlides(clip, 0);
    expect(at).toBe(1);
    expect(to.slides()).toHaveLength(2);
    expect(text(to, 1)).toContain("With a picture and a chart");
    const shapes = parseSlide(to, to.slides()[1].path).shapes;
    const pic = shapes.find((s) => s.image)!;
    expect(await to.media(pic.image!)).toEqual(PNG);
    const chart = shapes.find((s) => s.chart)!.chart!;
    expect(chart.data.series[0].values).toEqual([1, 2]);
    // The chart is the paste's own copy, with its workbook.
    const book = to.rels(chart.path).find((r) => r.type === "package");
    expect(book && (await to.media(book.target)).length).toBeGreaterThan(0);
    // It survives saving and reopening.
    const back = await PptxDocument.load(await to.save());
    expect(text(back, 1)).toContain("With a picture and a chart");
    expect(chartData(back.xml(parseSlide(back, back.slides()[1].path).shapes.find((s) => s.chart)!.chart!.path)!).series[0].values).toEqual([1, 2]);
  });

  it("keeps the source deck's look in another deck: its master, layouts and theme come along once", async () => {
    const from = await source();
    // Give the source a theme colour of its own so the paste can be told apart.
    const master = from.masterOf(from.layoutOf(from.slides()[0].path)!)!;
    from.setTheme(from.themeOf(master)!, { accent1: "FF00AA" });
    const to = await PptxDocument.load(await blankPptx());
    const before = to.layouts().length;
    to.pasteSlides(await from.copySlides([0, 1]), 0);
    const layoutA = to.layoutOf(to.slides()[1].path)!, layoutB = to.layoutOf(to.slides()[2].path)!;
    const masterA = to.masterOf(layoutA)!;
    // One imported master for both slides, registered with the presentation, with unique layout ids.
    expect(to.masterOf(layoutB)).toBe(masterA);
    expect(masterA).not.toBe(to.masterOf(to.layoutOf(to.slides()[0].path)!));
    const pres = to.xml("ppt/presentation.xml")!;
    expect(pres.getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "sldMasterId")).toHaveLength(2);
    const P = "http://schemas.openxmlformats.org/presentationml/2006/main";
    const masters = to.rels("ppt/presentation.xml").filter((r) => r.type === "slideMaster").map((r) => r.target);
    const ids = [pres, ...masters.map((m) => to.xml(m)!)].flatMap((d) => Array.from(d.getElementsByTagNameNS(P, "sldMasterId")).concat(Array.from(d.getElementsByTagNameNS(P, "sldLayoutId")))).map((e) => e.getAttribute("id"));
    expect(new Set(ids).size).toBe(ids.length);
    expect(to.layouts(masterA).length).toBe(before);
    expect(slideColors(to, to.slides()[1].path).scheme.accent1?.toUpperCase()).toBe("#FF00AA");
    // The destination's own slides keep its theme.
    expect(slideColors(to, to.slides()[0].path).scheme.accent1?.toUpperCase()).not.toBe("#FF00AA");
    const back = await PptxDocument.load(await to.save());
    expect(back.slides()).toHaveLength(3);
    expect(slideColors(back, back.slides()[1].path).scheme.accent1?.toUpperCase()).toBe("#FF00AA");
  });

  it("pastes several slides in order, and within the same deck, and leaves the clipboard unchanged by later edits", async () => {
    const from = await source();
    const clip = await from.copySlides([0, 1]);
    // Editing the source after copying doesn't change what pastes.
    addTextBox(from.xml(from.slides()[1].path)!, { x: 0, y: 0, w: 100, h: 100 }, "Added later");
    const last = from.pasteSlides(clip, 1);
    expect(last).toBe(3);
    expect(from.slides()).toHaveLength(4);
    expect(text(from, 2)).toContain("First slide");
    expect(text(from, 3)).toContain("With a picture and a chart");
    expect(text(from, 3)).not.toContain("Added later");
    const back = await PptxDocument.load(await from.save());
    expect(back.slides()).toHaveLength(4);
  });

  it("fits slides pasted from a deck of another size, text included, like Google's import", async () => {
    const from = await source();
    from.setSlideSize(12192000, 6858000);
    const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const firstBox = (doc: PptxDocument, i: number) => Array.from(doc.xml(doc.slides()[i].path)!.getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "sp")).find((sp) => sp.textContent?.includes("With a picture"))!;
    const box = firstBox(from, 1);
    const before = itemRect(box);
    const run = box.getElementsByTagNameNS(A, "rPr")[0] ?? box.getElementsByTagNameNS(A, "endParaRPr")[0];
    run.setAttribute("sz", "4000");
    const to = await PptxDocument.load(await blankPptx());
    to.setSlideSize(9144000, 5143500);
    const ownBefore = Number(to.xml(to.masterOf(to.layoutOf(to.slides()[0].path)!)!)!.getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "titleStyle")[0].getElementsByTagNameNS(A, "defRPr")[0].getAttribute("sz"));
    to.pasteSlides(await from.copySlides([1]), 0);
    const pasted = firstBox(to, 1);
    const after = itemRect(pasted);
    expect(after.w).toBeCloseTo(before.w * 0.75, -2);
    expect(after.x).toBeCloseTo(before.x * 0.75, -2);
    const size = (pasted.getElementsByTagNameNS(A, "rPr")[0] ?? pasted.getElementsByTagNameNS(A, "endParaRPr")[0]).getAttribute("sz");
    expect(size).toBe("3000");
    // The master it brought fits the new size too, and its title text shrinks with it.
    const master = to.masterOf(to.layoutOf(to.slides()[1].path)!)!;
    const own = to.masterOf(to.layoutOf(to.slides()[0].path)!)!;
    const right = (path: string) => Math.max(...Array.from(to.xml(path)!.getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "sp")).map((sp) => { const r = itemRect(sp); return r.x + r.w; }));
    expect(right(master)).toBeLessThanOrEqual(9144000 + 1);
    const title = (path: string) => Number(to.xml(path)!.getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "titleStyle")[0].getElementsByTagNameNS(A, "defRPr")[0].getAttribute("sz"));
    const srcTitle = Number(from.xml(from.masterOf(from.layoutOf(from.slides()[1].path)!)!)!.getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "titleStyle")[0].getElementsByTagNameNS(A, "defRPr")[0].getAttribute("sz"));
    expect(title(master)).toBe(Math.round(srcTitle * 0.75));
    expect(title(own)).toBe(ownBefore);
  });
});
