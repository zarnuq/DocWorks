import { describe, expect, it } from "vitest";
import { blankDocx } from "../src/infrastructure/blank";
import { DocxDocument } from "../src/infrastructure/docx/docx-document";
import { readWatermark, setWatermark, TextWatermark } from "../src/domain/docx/watermark";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const V = "urn:schemas-microsoft-com:vml";
const PNG = Uint8Array.from(atob("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="), (c) => c.charCodeAt(0));
const load = async () => DocxDocument.load(await blankDocx());
const roundTrip = async (d: DocxDocument) => DocxDocument.load(await d.save());
const text: TextWatermark = { kind: "text", text: "DRAFT", font: "Calibri", width: 468, height: 117, color: "#C0C0C0", opacity: 0.5, diagonal: true };

describe("watermarks", () => {
  it("has none in a blank document", async () => {
    expect(readWatermark(await load())).toBeNull();
  });

  it("writes a text watermark as Word's VML text path in the header and reads it back", async () => {
    const d = await load();
    setWatermark(d, text);
    const back = await roundTrip(d);
    expect(readWatermark(back)).toEqual(text);
    const hdr = back.headerFooter("header")!;
    const shape = hdr.getElementsByTagNameNS(V, "shape")[0];
    expect(shape.getAttribute("id")).toMatch(/^PowerPlusWaterMarkObject/);
    expect(shape.getAttribute("style")).toContain("rotation:315");
    expect(shape.getAttribute("style")).toContain("mso-position-horizontal:center");
    expect(shape.getElementsByTagNameNS(V, "textpath")[0].getAttribute("string")).toBe("DRAFT");
    // In the header's first paragraph, not a new one (which would make the header taller).
    expect(hdr.getElementsByTagNameNS(W, "p")).toHaveLength(1);
  });

  it("replaces and removes the watermark, leaving the rest of the header alone", async () => {
    const d = await load();
    const hdr = d.ensureHeaderFooter("header");
    hdr.getElementsByTagNameNS(W, "p")[0].appendChild(d.document.createElementNS(W, "w:r")).appendChild(d.document.createElementNS(W, "w:t")).textContent = "Header text";
    setWatermark(d, text);
    setWatermark(d, { ...text, text: "SECRET", diagonal: false });
    expect(readWatermark(d)).toMatchObject({ text: "SECRET", diagonal: false });
    expect(d.headerFooter("header")!.getElementsByTagNameNS(V, "shape")).toHaveLength(1);
    setWatermark(d, null);
    expect(readWatermark(d)).toBeNull();
    expect(d.headerFooter("header")!.textContent).toBe("Header text");
  });

  it("puts the watermark in every header of every section, first-page ones included", async () => {
    const d = await load();
    d.ensureHeaderFooter("header", "default");
    d.ensureHeaderFooter("header", "first");
    setWatermark(d, text);
    for (const type of ["default", "first"] as const) expect(d.headerFooter("header", type)!.getElementsByTagNameNS(V, "shape")).toHaveLength(1);
  });

  it("writes a picture watermark with its image in the header's relationships, faded or not", async () => {
    const d = await load();
    setWatermark(d, { kind: "image", data: PNG, ext: "png", width: 300, height: 200, washout: true });
    const back = await roundTrip(d);
    const wm = readWatermark(back)!;
    expect(wm).toMatchObject({ kind: "image", width: 300, height: 200, washout: true });
    expect(wm.kind === "image" && wm.path && (await back.media(wm.path))).toEqual(PNG);
    const data = back.headerFooter("header")!.getElementsByTagNameNS(V, "imagedata")[0];
    expect(data.getAttribute("gain")).toBe("19661f");
    setWatermark(back, { ...(wm as any), washout: false });
    expect(readWatermark(back)).toMatchObject({ washout: false, path: (wm as any).path });
  });
});
