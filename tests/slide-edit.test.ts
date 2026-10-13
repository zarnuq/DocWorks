import { describe, it, expect } from "vitest";
import { parseSlide, ShapeModel, Para } from "../src/domain/pptx/slide-model";
import { writeParagraphs, setRect, addTextBox, addPicture } from "../src/domain/pptx/slide-edit";
import { PptxDocument } from "../src/infrastructure/pptx/pptx-document";
import { loadSample, roundTrip } from "./helpers";

const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const ser = (n: Node) => new XMLSerializer().serializeToString(n);

async function shapeOn(slideIdx: number, name: string, doc?: PptxDocument) {
  doc ??= await loadSample();
  const path = doc.slides()[slideIdx].path;
  const find = () => parseSlide(doc!, path).shapes.find((s) => s.name === name)!;
  return { doc, path, shape: find(), reparse: find };
}

/** Clone paragraphs and let the test tweak them. */
const edit = (sh: ShapeModel, fn: (p: Para[]) => void) => {
  const paras: Para[] = structuredClone(sh.text!.paras);
  fn(paras);
  writeParagraphs(sh.text!, paras);
};
const rPrs = (sh: ShapeModel) => Array.from(sh.text!.el.getElementsByTagNameNS(A, "rPr"));
const pPrs = (sh: ShapeModel) => Array.from(sh.text!.el.getElementsByTagNameNS(A, "pPr"));

describe("writeParagraphs", () => {
  it("is a no-op for unchanged paragraphs", async () => {
    for (const [i, name] of [[1, "Content Placeholder 2"], [0, "Subtitle 2"], [2, "TextBox 1"], [2, "Rectangle 2"]] as const) {
      const { shape } = await shapeOn(i, name);
      const before = ser(shape.text!.el);
      edit(shape, () => {});
      expect(ser(shape.text!.el)).toBe(before);
    }
  });

  it("keeps an explicit lvl=\"0\" when unchanged", async () => {
    const { shape } = await shapeOn(2, "TextBox 1");
    const p = shape.text!.el.getElementsByTagNameNS(A, "p")[0];
    if (!pPrs(shape).length) p.prepend(p.ownerDocument.createElementNS(A, "a:pPr"));
    pPrs(shape)[0].setAttribute("lvl", "0");
    const before = ser(shape.text!.el);
    edit(shape, () => {});
    expect(ser(shape.text!.el)).toBe(before);
  });

  it("keeps line breaks and fields byte-identical when unchanged", async () => {
    const { reparse, shape: original } = await shapeOn(2, "TextBox 1");
    const p = original.text!.el.getElementsByTagNameNS(A, "p")[0];
    const replacement = new DOMParser().parseFromString(
      `<a:p xmlns:a="${A}"><a:br><a:rPr lang="en-US" sz="4000"/></a:br>` +
        `<a:r><a:rPr lang="en-US" sz="4000" b="1"/><a:t>Title</a:t></a:r>` +
        `<a:fld id="{X}" type="slidenum"><a:rPr lang="en-US"/><a:t>7</a:t></a:fld></a:p>`,
      "application/xml"
    ).documentElement;
    p.replaceWith(p.ownerDocument.importNode(replacement, true));
    const shape = reparse();
    const before = ser(shape.text!.el);
    edit(shape, () => {});
    expect(ser(shape.text!.el)).toBe(before);
  });

  it("changes text but keeps the run's existing properties", async () => {
    const { shape, reparse } = await shapeOn(1, "Content Placeholder 2");
    edit(shape, (p) => (p[1].runs[0].text = "Still bold red"));
    const run = reparse().text!.paras[1].runs[0];
    expect(run).toMatchObject({ text: "Still bold red", b: true, sz: 20, color: "#FF0000" });
  });

  it("writes explicit b=0 to un-bold inherited bold", async () => {
    const { shape, reparse } = await shapeOn(1, "Content Placeholder 2");
    edit(shape, (p) => (p[1].runs[0].b = false));
    expect(rPrs(shape)[0].getAttribute("b")).toBe("0");
    expect(reparse().text!.paras[1].runs[0].b).toBe(false);
  });

  it("writes only properties that differ from inherited defaults", async () => {
    const { shape } = await shapeOn(1, "Content Placeholder 2");
    edit(shape, (p) => {
      p[0].runs[0].i = true;
      p[0].runs[0].u = true;
      p[0].runs[0].strike = true;
    });
    const rPr = rPrs(shape)[0];
    expect(rPr.getAttribute("i")).toBe("1");
    expect(rPr.getAttribute("u")).toBe("sng");
    expect(rPr.getAttribute("strike")).toBe("sngStrike");
    expect(rPr.hasAttribute("sz")).toBe(false);
    expect(rPr.hasAttribute("b")).toBe(false);
  });

  it("writes size, color and font changes", async () => {
    const { shape, reparse } = await shapeOn(2, "TextBox 1");
    edit(shape, (p) => Object.assign(p[0].runs[0], { sz: 24, color: "#00ff00", font: "Georgia" }));
    expect(reparse().text!.paras[0].runs[0]).toMatchObject({ sz: 24, color: "#00FF00", font: "Georgia" });
    expect(ser(rPrs(shape)[0])).toMatch(/<a:solidFill><a:srgbClr val="00FF00"\/><\/a:solidFill><a:latin typeface="Georgia"\/>/);
  });

  it("replaces an existing fill instead of adding a second one", async () => {
    const { shape } = await shapeOn(1, "Content Placeholder 2");
    edit(shape, (p) => (p[1].runs[0].color = "#0000FF"));
    expect(rPrs(shape)[0].getElementsByTagNameNS(A, "solidFill")).toHaveLength(1);
    expect(ser(rPrs(shape)[0])).toContain('val="0000FF"');
  });

  it("writes line breaks as a:br", async () => {
    const { shape, reparse } = await shapeOn(2, "TextBox 1");
    edit(shape, (p) => (p[0].runs[0].text = "one\ntwo"));
    expect(shape.text!.el.getElementsByTagNameNS(A, "br")).toHaveLength(1);
    expect(reparse().text!.paras[0].runs.map((r) => r.text)).toEqual(["one", "\n", "two"]);
  });

  it("only writes algn when it differs from the inherited alignment", async () => {
    const { shape, reparse } = await shapeOn(1, "Content Placeholder 2");
    edit(shape, (p) => (p[0].align = "r"));
    expect(pPrs(shape)[0].getAttribute("algn")).toBe("r");
    expect(reparse().text!.paras[0].align).toBe("r");
    expect(pPrs(shape)[1].hasAttribute("algn")).toBe(false);
  });

  it("toggles bullets with buNone / buChar in schema order", async () => {
    const { shape, reparse } = await shapeOn(1, "Content Placeholder 2");
    edit(shape, (p) => (p[0].bullet = null));
    expect(reparse().text!.paras[0].bullet).toBeNull();
    expect(pPrs(shape)[0].getElementsByTagNameNS(A, "buNone")).toHaveLength(1);

    const tb = await shapeOn(2, "TextBox 1");
    const pPr = tb.shape.text!.el.getElementsByTagNameNS(A, "pPr")[0];
    pPr.appendChild(pPr.ownerDocument.createElementNS(A, "a:defRPr"));
    edit(tb.shape, (p) => (p[0].bullet = "•"));
    expect(tb.reparse().text!.paras[0].bullet).toBe("•");
    expect(Array.from(pPrs(tb.shape)[0].children).map((c) => c.localName)).toEqual(["buChar", "defRPr"]);
  });

  it("changes levels and picks up the new level's defaults", async () => {
    const { shape, reparse } = await shapeOn(1, "Content Placeholder 2");
    edit(shape, (p) => {
      p[1].lvl = 0;
      p[0].lvl = 2;
    });
    const [p0, p1] = reparse().text!.paras;
    expect(p0.lvl).toBe(2);
    expect(p1.lvl).toBe(0);
    expect(pPrs(shape)[1].hasAttribute("lvl")).toBe(false);
  });

  it("adds new paragraphs and removes deleted ones", async () => {
    const { shape, reparse } = await shapeOn(1, "Content Placeholder 2");
    edit(shape, (p) => {
      p.splice(0, 1);
      p.push({ ...structuredClone(p[0]), runs: [{ ...p[0].runs[1], text: "Added" }] });
    });
    expect(reparse().text!.paras.map((p) => p.runs.map((r) => r.text).join(""))).toEqual(["Bold red italic", "Added"]);
  });

  it("keeps one empty paragraph when all text is deleted", async () => {
    const { shape, reparse } = await shapeOn(2, "TextBox 1");
    edit(shape, (p) => p.splice(0));
    expect(shape.text!.el.getElementsByTagNameNS(A, "p")).toHaveLength(1);
    expect(reparse().text!.paras[0].runs).toEqual([]);
  });

  it("escapes XML special characters through save", async () => {
    const { doc, shape } = await shapeOn(2, "TextBox 1");
    edit(shape, (p) => (p[0].runs[0].text = `<b> & "q"`));
    const back = await roundTrip(doc);
    const again = parseSlide(back, back.slides()[2].path).shapes.find((s) => s.name === "TextBox 1")!;
    expect(again.text!.paras[0].runs[0].text).toBe(`<b> & "q"`);
  });
});

describe("shape geometry and insertion", () => {
  it("writes an xfrm for placeholders that inherited their geometry", async () => {
    const { shape, reparse } = await shapeOn(0, "Title 1");
    setRect(shape.el, { x: 10, y: 20, w: 30, h: 40.6 });
    expect(reparse().rect).toEqual({ x: 10, y: 20, w: 30, h: 41 });
    const spPr = shape.el.getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "spPr")[0];
    expect(spPr.firstElementChild!.localName).toBe("xfrm");
  });

  it("moves graphic frames via p:xfrm", async () => {
    const { shape, reparse } = await shapeOn(2, "Table 5");
    setRect(shape.el, { x: 1, y: 2, w: 3, h: 4 });
    expect(reparse().rect).toEqual({ x: 1, y: 2, w: 3, h: 4 });
  });

  it("adds a text box with a unique id that survives save", async () => {
    const doc = await loadSample();
    const path = doc.slides()[2].path;
    const el = addTextBox(doc.xml(path)!, { x: 0, y: 0, w: 100, h: 50 }, "New <text>");
    const ids = Array.from(doc.xml(path)!.getElementsByTagNameNS("*", "cNvPr")).map((c) => c.getAttribute("id"));
    expect(new Set(ids).size).toBe(ids.length);
    expect(el.parentElement!.localName).toBe("spTree");
    const back = await roundTrip(doc);
    const tb = parseSlide(back, path).shapes.at(-1)!;
    expect(tb.rect).toEqual({ x: 0, y: 0, w: 100, h: 50 });
    expect(tb.text!.paras[0].runs[0].text).toBe("New <text>");
  });

  it("inserts shapes before a trailing p:extLst", async () => {
    const doc = await loadSample();
    const xml = doc.xml(doc.slides()[0].path)!;
    const tree = xml.getElementsByTagNameNS("*", "spTree")[0];
    tree.appendChild(xml.createElementNS("http://schemas.openxmlformats.org/presentationml/2006/main", "p:extLst"));
    addTextBox(xml, { x: 0, y: 0, w: 1, h: 1 }, "x");
    expect(tree.lastElementChild!.localName).toBe("extLst");
  });

  it("adds a picture bound to an image relationship", async () => {
    const doc = await loadSample();
    const path = doc.slides()[1].path;
    const rId = doc.addImage(path, new Uint8Array([1]), "png");
    addPicture(doc.xml(path)!, { x: 5, y: 6, w: 7, h: 8 }, rId, "photo.png");
    const back = await roundTrip(doc);
    const pic = parseSlide(back, path).shapes.find((s) => s.kind === "pic")!;
    expect(pic.image).toMatch(/^ppt\/media\/image\d+\.png$/);
    expect(pic.rect).toEqual({ x: 5, y: 6, w: 7, h: 8 });
  });
});

describe("superscript, subscript and highlight", () => {
  it("writes a:rPr baseline and a:highlight (in schema order) and reads them back", async () => {
    const { shape, reparse } = await shapeOn(2, "TextBox 1");
    edit(shape, (p) => {
      p[0].runs[0] = { ...p[0].runs[0], baseline: 30, highlight: "#FFFF00", color: "#FF0000" };
    });
    const rPr = rPrs(shape)[0];
    expect(rPr.getAttribute("baseline")).toBe("30000");
    const names = Array.from(rPr.children).map((c) => c.localName);
    expect(names.indexOf("highlight")).toBeGreaterThan(names.indexOf("solidFill"));
    if (names.includes("latin")) expect(names.indexOf("highlight")).toBeLessThan(names.indexOf("latin"));
    expect(rPr.getElementsByTagNameNS(A, "highlight")[0].firstElementChild!.getAttribute("val")).toBe("FFFF00");
    expect(reparse().text!.paras[0].runs[0]).toMatchObject({ baseline: 30, highlight: "#FFFF00" });
    edit(reparse(), (p) => {
      p[0].runs[0] = { ...p[0].runs[0], baseline: 0, highlight: null };
    });
    expect(rPrs(reparse())[0].hasAttribute("baseline")).toBe(false);
    expect(rPrs(reparse())[0].getElementsByTagNameNS(A, "highlight")).toHaveLength(0);
  });
});
