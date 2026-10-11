import { expect, it } from "vitest";
import { blankPptx } from "../src/infrastructure/blank";
import { PptxDocument } from "../src/infrastructure/pptx/pptx-document";
import { parseXml, relsPathOf } from "../src/infrastructure/ooxml/ooxml-package";
import { NS } from "../src/core/utils/xml";
it("copies a chart resource graph with an independent embedded workbook and keeps undo binary-safe", async () => {
  const a = await PptxDocument.load(await blankPptx());
  const b = await PptxDocument.load(await blankPptx());
  a.putXml("ppt/charts/chart1.xml", parseXml('<chart/>')!, "application/test-chart");
  a.putBinary("ppt/embeddings/data.xlsx", new Uint8Array([1, 2, 3]), "application/test-workbook");
  a.addSlideRel("ppt/charts/chart1.xml", "package", "ppt/embeddings/data.xlsx");
  const before = b.snapshot();
  const bundle = await a.captureResources(["ppt/charts/chart1.xml"]);
  const copied = b.importResources(bundle);
  const path = copied.get("ppt/charts/chart1.xml")!;
  const workbook = b.relOfType(path, "package")!;
  expect(await b.media(workbook)).toEqual(new Uint8Array([1, 2, 3]));
  const edited = b.snapshot();
  await b.save();
  b.restore(before);
  const undone = await PptxDocument.load(await b.save());
  expect(undone.xml(path)).toBeNull();
  await expect(undone.media(workbook)).rejects.toThrow();
  b.restore(edited);
  expect(await b.media(workbook)).toEqual(new Uint8Array([1, 2, 3]));
});
it("creates and registers a custom layout without changing the source", async () => {
  const doc = await PptxDocument.load(await blankPptx());
  const base = doc.layoutOf(doc.slides()[0].path)!;
  const original = new XMLSerializer().serializeToString(doc.xml(base)!);
  const path = doc.createLayout(base, "My layout");
  expect(doc.layouts(doc.masterOf(base)!).find(l => l.path === path)?.name).toBe("My layout");
  expect(new XMLSerializer().serializeToString(doc.xml(base)!)).toBe(original);
  doc.setTheme(doc.themeOf(doc.masterOf(base)!)!, { accent1: "ff0000" }, "Arial", "Verdana");
  const reloaded = await PptxDocument.load(await doc.save());
  expect(reloaded.xml(path)).not.toBeNull();
  expect(reloaded.xml(relsPathOf(path))!.getElementsByTagNameNS(NS.pkgRels, "Relationship").length).toBeGreaterThan(0);
});

it("does not enumerate media that was undone after a save",async()=>{
 const doc=await PptxDocument.load(await blankPptx());const before=doc.snapshot();
 doc.putBinary("ppt/media/undone.mp4",new Uint8Array([1]),"video/mp4");await doc.save();
 doc.restore(before);expect(doc.mediaPaths()).not.toContain("ppt/media/undone.mp4");
});
