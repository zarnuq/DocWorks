import { describe, expect, it } from "vitest";
import { NS } from "../src/core/utils/xml";
import { readTransition, writeTransition } from "../src/domain/pptx/animation";
import { PptxDocument } from "../src/infrastructure/pptx/pptx-document";
import { blankPptx } from "../src/infrastructure/blank";

const P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main";
const slide = (inner = "") => new DOMParser().parseFromString(
  `<p:sld xmlns:p="${NS.p}" xmlns:mc="${NS.mc}" xmlns:p14="${P14}"><p:cSld><p:spTree/></p:cSld><p:clrMapOvr/>${inner}</p:sld>`, "application/xml"
);
const xml = (doc: Document | Element) => new XMLSerializer().serializeToString(doc);

describe("writeTransition", () => {
  it.each(["fade", "push", "wipe", "cover", "pull", "zoom"])("adds %s with an exact duration and valid slide order", (type) => {
    const doc = slide('<p:timing/><p:extLst/>');
    writeTransition(doc, { type, dir: type === "zoom" ? "in" : "l", dur: 1250 });
    expect(readTransition(doc)).toEqual({ type, dir: type === "fade" ? null : type === "zoom" ? "in" : "l", dur: 1250, advanceAfter: null });
    expect(Array.from(doc.documentElement.children).map((e) => e.localName)).toEqual(["cSld", "clrMapOvr", "AlternateContent", "timing", "extLst"]);
    const fallback = doc.getElementsByTagNameNS(NS.mc, "Fallback")[0].firstElementChild!;
    expect(fallback.getAttribute("spd")).toBe("slow");
    expect(fallback.hasAttributeNS(P14, "dur")).toBe(false);
    const reparsed = new DOMParser().parseFromString(xml(doc), "application/xml");
    expect(reparsed.querySelector("parsererror")).toBeNull();
    expect(readTransition(reparsed)?.dur).toBe(1250);
    const choice = reparsed.getElementsByTagNameNS(NS.mc, "Choice")[0];
    expect(choice.lookupNamespaceURI(choice.getAttribute("Requires")!)).toBe(P14);
  });

  it("preserves timing, sound, extensions and unrelated XML while changing the effect", () => {
    const doc = slide('<p:transition advClick="0" advTm="7000"><p:push dir="u"/><p:sndAc><p:endSnd/></p:sndAc><p:extLst/></p:transition><p:timing><p:tnLst/></p:timing><p:extLst/>');
    const timing = xml(doc.getElementsByTagNameNS(NS.p, "timing")[0]);
    const content = xml(doc.getElementsByTagNameNS(NS.p, "cSld")[0]);
    writeTransition(doc, { type: "fade", dur: 650 });
    expect(readTransition(doc)).toEqual({ type: "fade", dir: null, dur: 650, advanceAfter: 7000 });
    for (const t of Array.from(doc.getElementsByTagNameNS(NS.p, "transition"))) {
      expect(t.getAttribute("advClick")).toBe("0");
      expect(Array.from(t.children).map((e) => e.localName)).toEqual(["fade", "sndAc", "extLst"]);
    }
    expect(xml(doc.getElementsByTagNameNS(NS.p, "timing")[0])).toBe(timing);
    expect(xml(doc.getElementsByTagNameNS(NS.p, "cSld")[0])).toBe(content);
  });

  it("keeps alternate representations in step without replacing unknown effects on a speed edit", () => {
    const doc = slide('<mc:AlternateContent><mc:Choice Requires="p14"><p:transition p14:dur="1500"><p14:vortex dir="r"/></p:transition></mc:Choice><mc:Fallback><p:transition><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>');
    writeTransition(doc, { dur: 2000 });
    expect(readTransition(doc)?.type).toBe("vortex");
    expect(doc.getElementsByTagNameNS(P14, "vortex")).toHaveLength(1);
    writeTransition(doc, { type: "wipe", dir: "d" });
    for (const t of Array.from(doc.getElementsByTagNameNS(NS.p, "transition"))) {
      expect(t.firstElementChild?.localName).toBe("wipe");
      expect(t.firstElementChild?.getAttribute("dir")).toBe("d");
    }
  });

  it("removes only the visual effect when None is selected", () => {
    const doc = slide('<p:transition advTm="3000"><p:fade/></p:transition><p:timing/>');
    writeTransition(doc, { type: "cut" });
    expect(readTransition(doc)).toEqual({ type: "cut", dir: null, dur: 0, advanceAfter: 3000 });
    expect(doc.getElementsByTagNameNS(NS.p, "timing")).toHaveLength(1);
  });

  it("does not create XML for None on a slide without transitions", () => {
    const doc = slide();
    const before = xml(doc);
    writeTransition(doc, { type: "cut" });
    expect(xml(doc)).toBe(before);
  });

  it.each([NaN, Infinity, -1, 0])("rejects invalid durations (%s) before editing XML", (dur) => {
    const doc = slide();
    const before = xml(doc);
    expect(() => writeTransition(doc, { type: "fade", dur })).toThrow();
    expect(xml(doc)).toBe(before);
  });

  it("saves and reopens transitions in a deck that had no extension namespaces", async () => {
    const doc = await PptxDocument.load(await blankPptx());
    const path = doc.slides()[0].path;
    writeTransition(doc.xml(path)!, { type: "push", dir: "r", dur: 850 });
    const reloaded = await PptxDocument.load(await doc.save());
    expect(readTransition(reloaded.xml(path)!)).toEqual({ type: "push", dir: "r", dur: 850, advanceAfter: null });
  });
});
