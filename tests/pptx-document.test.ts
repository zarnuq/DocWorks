import { describe, it, expect } from "vitest";
import JSZip from "jszip";
import { PptxDocument } from "../src/infrastructure/pptx/pptx-document";
import { loadSample, roundTrip, texts } from "./helpers";

describe("PptxDocument.load", () => {
  it("lists slides in presentation order with slide size", async () => {
    const doc = await loadSample();
    expect(doc.slides().map((s) => s.path)).toEqual([
      "ppt/slides/slide1.xml",
      "ppt/slides/slide2.xml",
      "ppt/slides/slide3.xml",
    ]);
    expect(doc.slideSize).toEqual({ cx: 9144000, cy: 6858000 });
  });

  it("rejects a zip without ppt/presentation.xml", async () => {
    const zip = new JSZip();
    zip.file("hello.txt", "hi");
    await expect(PptxDocument.load(await zip.generateAsync({ type: "uint8array" }))).rejects.toThrow(
      /presentation\.xml is missing/
    );
  });

  it("rejects non-zip bytes", async () => {
    await expect(PptxDocument.load(new Uint8Array([1, 2, 3]))).rejects.toThrow();
  });

  it("resolves layout, master, theme and notes relationships", async () => {
    const doc = await loadSample();
    const s1 = doc.slides()[0].path;
    expect(doc.layoutOf(s1)).toBe("ppt/slideLayouts/slideLayout1.xml");
    expect(doc.masterOf("ppt/slideLayouts/slideLayout1.xml")).toBe("ppt/slideMasters/slideMaster1.xml");
    expect(doc.themeOf("ppt/slideMasters/slideMaster1.xml")).toBe("ppt/theme/theme1.xml");
    expect(doc.notesOf(s1)).toBe("ppt/notesSlides/notesSlide1.xml");
    expect(doc.notesOf(doc.slides()[1].path)).toBeNull();
  });

  it("reads media bytes behind an image relationship", async () => {
    const doc = await loadSample();
    const s3 = doc.slides()[2].path;
    const target = doc.relTarget(s3, "rId2");
    expect(target).toBe("ppt/media/image1.png");
    const bytes = await doc.media(target!);
    expect(Array.from(bytes.slice(1, 4))).toEqual([0x50, 0x4e, 0x47]); // "PNG"
  });
});

describe("PptxDocument.save", () => {
  it("round-trips unchanged content", async () => {
    const doc = await roundTrip(await loadSample());
    expect(doc.slides()).toHaveLength(3);
    expect(texts(doc.xml(doc.slides()[0].path)!)).toEqual(["Hello Deck", "Subtitle here"]);
  });

  it("persists in-memory XML edits", async () => {
    const doc = await loadSample();
    doc.xml(doc.slides()[0].path)!.getElementsByTagNameNS("*", "t")[0].textContent = "Changed";
    const back = await roundTrip(doc);
    expect(texts(back.xml(back.slides()[0].path)!)[0]).toBe("Changed");
  });

  it("writes an XML declaration on every XML part", async () => {
    const doc = await loadSample();
    const zip = await JSZip.loadAsync(await doc.save());
    expect(await zip.file("ppt/presentation.xml")!.async("text")).toMatch(/^<\?xml version="1.0"/);
  });
});

describe("slide operations", () => {
  it("adds a slide after an index with the layout's placeholders", async () => {
    const doc = await loadSample();
    const idx = doc.addSlide(0, "ppt/slideLayouts/slideLayout2.xml");
    expect(idx).toBe(1);
    const back = await roundTrip(doc);
    const slides = back.slides();
    expect(slides).toHaveLength(4);
    const added = slides[1].path;
    expect(back.layoutOf(added)).toBe("ppt/slideLayouts/slideLayout2.xml");
    const phTypes = Array.from(back.xml(added)!.getElementsByTagNameNS("*", "ph")).map((p) => p.getAttribute("type"));
    expect(phTypes).toEqual(["title", null]); // title + body (idx=1, no type); dt/ftr/sldNum skipped
    expect(new Set(slides.map((s) => s.id)).size).toBe(4);
    expect(new Set(slides.map((s) => s.path)).size).toBe(4);
  });

  it("registers new slides in [Content_Types].xml", async () => {
    const doc = await loadSample();
    doc.addSlide(2);
    const zip = await JSZip.loadAsync(await doc.save());
    const ct = await zip.file("[Content_Types].xml")!.async("text");
    expect(ct).toContain('PartName="/ppt/slides/slide4.xml"');
  });

  it("defaults the new slide's layout to the neighbour's layout", async () => {
    const doc = await loadSample();
    const idx = doc.addSlide(1);
    expect(doc.layoutOf(doc.slides()[idx].path)).toBe(doc.layoutOf(doc.slides()[1].path));
  });

  it("duplicates a slide with its content but without its notes", async () => {
    const doc = await loadSample();
    const idx = doc.duplicateSlide(0);
    const back = await roundTrip(doc);
    const copy = back.slides()[idx].path;
    expect(idx).toBe(1);
    expect(texts(back.xml(copy)!)).toEqual(["Hello Deck", "Subtitle here"]);
    expect(back.notesOf(copy)).toBeNull();
    expect(back.notesOf(back.slides()[0].path)).toBe("ppt/notesSlides/notesSlide1.xml");
  });

  it("duplicated slide keeps image relationships", async () => {
    const doc = await loadSample();
    const idx = doc.duplicateSlide(2);
    expect(doc.relTarget(doc.slides()[idx].path, "rId2")).toBe("ppt/media/image1.png");
  });

  it("deletes a slide, its part, rels, notes and content type", async () => {
    const doc = await loadSample();
    doc.deleteSlide(0);
    const zip = await JSZip.loadAsync(await doc.save());
    const back = await PptxDocument.load(await zip.generateAsync({ type: "uint8array" }));
    expect(back.slides().map((s) => s.path)).toEqual(["ppt/slides/slide2.xml", "ppt/slides/slide3.xml"]);
    expect(zip.file("ppt/slides/slide1.xml")).toBeNull();
    expect(zip.file("ppt/slides/_rels/slide1.xml.rels")).toBeNull();
    expect(zip.file("ppt/notesSlides/notesSlide1.xml")).toBeNull();
    const ct = await zip.file("[Content_Types].xml")!.async("text");
    expect(ct).not.toContain("/ppt/slides/slide1.xml");
    expect(ct).not.toContain("/ppt/notesSlides/notesSlide1.xml");
    const presRels = await zip.file("ppt/_rels/presentation.xml.rels")!.async("text");
    expect(presRels).not.toContain("slides/slide1.xml");
  });

  it("refuses to delete the last remaining slide", async () => {
    const doc = await loadSample();
    doc.deleteSlide(0);
    doc.deleteSlide(0);
    expect(() => doc.deleteSlide(0)).toThrow(/last slide/);
  });

  it("moves slides", async () => {
    const doc = await loadSample();
    doc.moveSlide(2, 0);
    const back = await roundTrip(doc);
    expect(back.slides().map((s) => s.path)).toEqual([
      "ppt/slides/slide3.xml",
      "ppt/slides/slide1.xml",
      "ppt/slides/slide2.xml",
    ]);
  });

  it("ignores out-of-range moves", async () => {
    const doc = await loadSample();
    doc.moveSlide(0, 5);
    doc.moveSlide(-1, 0);
    expect(doc.slides().map((s) => s.path)[0]).toBe("ppt/slides/slide1.xml");
  });
});

describe("addImage", () => {
  it("stores media, adds a relationship and a content-type default", async () => {
    const doc = await loadSample();
    const s2 = doc.slides()[1].path;
    const rId = doc.addImage(s2, new Uint8Array([1, 2, 3]), "jpeg");
    const back = await roundTrip(doc);
    const target = back.relTarget(back.slides()[1].path, rId)!;
    expect(target).toMatch(/^ppt\/media\/image\d+\.jpeg$/);
    expect(target).not.toBe("ppt/media/image1.png");
    expect(Array.from(await back.media(target))).toEqual([1, 2, 3]);
    const zip = await JSZip.loadAsync(await back.save());
    expect(await zip.file("[Content_Types].xml")!.async("text")).toMatch(/Extension="jpeg"/);
  });

  it("uses a fresh rId on the slide", async () => {
    const doc = await loadSample();
    const s3 = doc.slides()[2].path;
    const rId = doc.addImage(s3, new Uint8Array([9]), "png");
    expect(doc.rels(s3).filter((r) => r.id === rId)).toHaveLength(1);
    expect(rId).not.toBe("rId1");
    expect(rId).not.toBe("rId2");
  });
});

describe("snapshot / restore", () => {
  it("undoes structural changes", async () => {
    const doc = await loadSample();
    const snap = doc.snapshot();
    doc.deleteSlide(1);
    doc.addSlide(0);
    doc.restore(snap);
    expect(doc.slides().map((s) => s.path)).toEqual([
      "ppt/slides/slide1.xml",
      "ppt/slides/slide2.xml",
      "ppt/slides/slide3.xml",
    ]);
    const back = await roundTrip(doc);
    expect(texts(back.xml(back.slides()[1].path)!)).toContain("Bullets");
  });

  it("undoes text edits", async () => {
    const doc = await loadSample();
    const snap = doc.snapshot();
    doc.xml(doc.slides()[0].path)!.getElementsByTagNameNS("*", "t")[0].textContent = "Changed";
    doc.restore(snap);
    expect(texts(doc.xml(doc.slides()[0].path)!)[0]).toBe("Hello Deck");
  });
});

describe("mediaPaths", () => {
  it("lists media parts including newly added ones", async () => {
    const doc = await loadSample();
    expect(doc.mediaPaths()).toEqual(["ppt/media/image1.png"]);
    doc.addImage(doc.slides()[0].path, new Uint8Array([1]), "gif");
    expect(doc.mediaPaths()).toEqual(["ppt/media/image1.png", "ppt/media/image2.gif"]);
  });
});
