import { describe, it, expect } from "vitest";
import { DOMSerializer, Node as PMNode } from "prosemirror-model";
import { readDocx } from "../src/domain/docx/docx-reader";
import { docxSchema as schema } from "../src/domain/docx/schema";
import { readSample, W } from "./helpers";

const inlineTypes = (n: PMNode) => {
  const out: string[] = [];
  n.forEach((c) => out.push(c.type.name));
  return out;
};
const marked = (n: PMNode, mark: string) => {
  let s = "";
  n.descendants((c) => {
    if (c.isText && c.marks.some((m) => m.type.name === mark)) s += c.text;
  });
  return s;
};
const markAttrs = (n: PMNode, mark: string) => {
  let attrs = null as Record<string, unknown> | null;
  n.descendants((c) => {
    attrs ??= c.marks.find((m) => m.type.name === mark)?.attrs ?? null;
  });
  return attrs;
};

describe("readDocx", () => {
  it("keeps paragraph styles and renders headings as h1-h6", async () => {
    const { doc } = await readSample();
    const h = doc.child(0);
    expect(h.type.name).toBe("paragraph");
    expect(h.attrs.style).toBe("Heading1");
    expect(h.textContent).toBe("Doc Heading");
    const dom = DOMSerializer.fromSchema(schema).serializeNode(h) as HTMLElement;
    expect(dom.tagName).toBe("H1");
  });

  it("reads direct run formatting as marks", async () => {
    const { doc } = await readSample();
    const p = doc.child(1);
    expect(marked(p, "bold")).toBe("bold");
    expect(marked(p, "italic")).toBe("italic");
    expect(marked(p, "color")).toBe(" red20");
    expect(markAttrs(p, "color")).toEqual({ color: "#FF0000" });
    expect(markAttrs(p, "size")).toEqual({ pt: 20 });
    expect(marked(p, "highlight")).toBe(" yellow");
    expect(markAttrs(p, "highlight")).toEqual({ color: "#FFFF00" });
    expect(p.textContent).toContain("\tafter tab");
  });

  it("keeps the original run properties on every run", async () => {
    const { doc } = await readSample();
    const xml = markAttrs(doc.child(1), "rpr")?.xml as string;
    expect(xml).toMatch(/^<w:rPr/);
  });

  it("reads hyperlinks as link marks and keeps bookmarks and fields as atoms", async () => {
    const { doc } = await readSample();
    const p = doc.child(2);
    expect(marked(p, "link")).toBe("example");
    expect(markAttrs(p, "link")).toMatchObject({ href: "https://example.com" });
    const opaque: PMNode[] = [];
    p.forEach((c) => c.type.name === "opaque" && opaque.push(c));
    expect(opaque.map((o) => o.attrs.xml.match(/^<w:(\w+)/)?.[1])).toEqual(["bookmarkStart", "bookmarkEnd", "fldSimple"]);
    expect(opaque[2].attrs.text).toBe("1");
  });

  it("keeps list numbering on the paragraph", async () => {
    const { doc } = await readSample();
    expect(doc.child(3).attrs).toMatchObject({ style: "ListBullet", numId: null });
    expect(doc.child(12).attrs).toMatchObject({ numId: "1", ilvl: 1 });
  });

  it("reads page breaks inline and ignores rendering hints", async () => {
    const { doc } = await readSample();
    const p = doc.child(6);
    expect(p.textContent).toBe("Before breakafter break");
    expect(inlineTypes(p).filter((t) => t === "page_break")).toHaveLength(1);
    expect(doc.child(7).firstChild?.type.name).toBe("image");
  });

  it("reads images with size, relationship and preview source", async () => {
    const { doc } = await readSample();
    const img = doc.child(7).firstChild!;
    expect(img.attrs).toMatchObject({ width: 96, height: 96 });
    expect(img.attrs.rId).toMatch(/^rId\d+$/);
    expect(img.attrs.src).toMatch(/^data:image\/png/);
    expect(img.attrs.xml).toMatch(/^<w:r[\s>]/);
  });

  it("reads tables with spans and shading", async () => {
    const { doc } = await readSample();
    const table = doc.child(8);
    expect(table.type.name).toBe("table");
    const row0 = table.child(0);
    const row1 = table.child(1);
    expect(row0.childCount).toBe(2);
    expect(row0.child(0).attrs.colspan).toBe(2);
    expect(row0.child(0).textContent).toBe("Merged");
    expect(row1.childCount).toBe(3);
    expect(row1.child(0).attrs.background).toBe("#FFFF00");
  });

  it("keeps content controls as read-only blocks", async () => {
    const { doc } = await readSample();
    const raw = doc.child(10);
    expect(raw.type.name).toBe("raw_block");
    expect(raw.attrs.text).toBe("Inside control");
  });

  it("keeps unknown run content as an atom", async () => {
    const { doc } = await readSample();
    const p = doc.child(11);
    expect(inlineTypes(p)).toEqual(["text", "opaque", "text"]);
  });

  it("records an origin for every paragraph and table", async () => {
    const { doc, origins } = await readSample();
    const keys: string[] = [];
    doc.descendants((n) => {
      if (n.attrs.key) keys.push(n.attrs.key);
    });
    expect(new Set(keys).size).toBe(keys.length);
    expect(keys.every((k) => origins.has(k))).toBe(true);
    expect(origins.get(doc.child(0).attrs.key)!.el.localName).toBe("p");
  });

  it("keeps run formatting on line breaks", () => {
    const body = new DOMParser().parseFromString(
      `<w:body xmlns:w="${W}"><w:p><w:r><w:rPr><w:b/></w:rPr><w:br/></w:r></w:p></w:body>`,
      "application/xml"
    ).documentElement;
    const br = readDocx(body, { imageSrc: () => null, hyperlinkHref: () => null }).doc.child(0).child(0);
    expect(br.type.name).toBe("hard_break");
    expect(br.marks.map((m) => m.type.name)).toEqual(["rpr", "bold"]);
  });

  it("returns one empty paragraph for an empty body", () => {
    const body = new DOMParser().parseFromString(`<w:body xmlns:w="${W}"/>`, "application/xml").documentElement;
    const { doc } = readDocx(body, { imageSrc: () => null, hyperlinkHref: () => null });
    expect(doc.childCount).toBe(1);
    expect(doc.child(0).type.name).toBe("paragraph");
  });
});
