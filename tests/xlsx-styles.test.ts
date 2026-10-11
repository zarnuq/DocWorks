import { describe, it, expect } from "vitest";
import { Styles, themeColors } from "../src/domain/xlsx/styles";
import { parseXml } from "../src/infrastructure/ooxml/ooxml-package";
import { STYLES, S } from "./xlsx-fixtures";

const styles = (xml = STYLES, theme: Record<string, string> = {}) => new Styles(parseXml(xml)!, theme);
const children = (doc: Document, tag: string) => Array.from(doc.getElementsByTagNameNS(S, tag)[0].children);

describe("cell styles", () => {
  it("turns a cell format into CSS: font, color, fill, borders, alignment", () => {
    const css = styles().css(1);
    expect(css).toMatchObject({
      "font-weight": "bold",
      color: "#FF0000",
      "background-color": "#FFFF00",
      "border-bottom": "1px solid #000000",
      "text-align": "center",
      "font-size": "11pt",
    });
    expect(css["font-family"]).toMatch(/^"Calibri"/);
  });

  it("reads number formats: built-in ids and custom codes", () => {
    const s = styles();
    expect(s.numFmt(0)).toBe("General");
    expect(s.numFmt(2)).toBe("yyyy-mm-dd");
    expect(s.numFmt(99)).toBe("General");
  });

  it("resolves theme and indexed colors with tints", () => {
    const xml = STYLES.replace('<color rgb="FFFF0000"/>', '<color theme="4" tint="0.4"/>').replace('<fgColor rgb="FFFFFF00"/>', '<fgColor indexed="10"/>');
    const s = styles(xml, { accent1: "#4472C4" });
    expect(s.css(1).color).toBe("#8FAADC");
    expect(s.css(1)["background-color"]).toBe("#FF0000");
  });

  it("reads theme colors in spreadsheet index order (light 1 first)", () => {
    const theme = parseXml(
      `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="x">` +
        `<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>` +
        `<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1>` +
        `</a:clrScheme></a:themeElements></a:theme>`
    )!;
    expect(themeColors(theme)).toMatchObject({ lt1: "#FFFFFF", dk1: "#000000", accent1: "#4472C4" });
  });

  it("makes a new format for a change and reuses it the next time", () => {
    const s = styles();
    const bold = s.derive(0, { bold: true });
    expect(bold).toBe(3);
    expect(s.css(bold)["font-weight"]).toBe("bold");
    expect(s.derive(0, { bold: true })).toBe(3);
    expect(children(s.doc, "fonts")).toHaveLength(3);
    expect(s.doc.getElementsByTagNameNS(S, "cellXfs")[0].getAttribute("count")).toBe("4");
    // The original format and font are untouched.
    expect(s.css(0)["font-weight"]).toBeUndefined();
  });

  it("writes font properties in schema order", () => {
    const s = styles();
    const xf = s.derive(0, { italic: true, color: "#00FF00", underline: true, fontSize: 14 });
    const font = children(s.doc, "fonts")[s.xf(xf).fontId];
    expect(Array.from(font.children).map((c) => c.localName)).toEqual(["i", "u", "sz", "color", "name"]);
    expect(font.getElementsByTagNameNS(S, "color")[0].getAttribute("rgb")).toBe("FF00FF00");
  });

  it("changes fills, number formats, alignment and borders", () => {
    const s = styles();
    const xf = s.derive(1, { fill: "#0000FF", numFmt: "0.00%", hAlign: "right", wrap: true, borders: { top: "thin" } });
    const css = s.css(xf);
    expect(css["background-color"]).toBe("#0000FF");
    expect(css["text-align"]).toBe("right");
    expect(css["white-space"]).toBe("pre-wrap");
    expect(css["border-top"]).toBe("1px solid #000000");
    expect(css["border-bottom"]).toBe("1px solid #000000");
    expect(s.numFmt(xf)).toBe("0.00%");
    expect(s.derive(0, { numFmt: "yyyy-mm-dd" })).toBe(s.derive(0, { numFmt: "yyyy-mm-dd" }));
    expect(s.xf(s.derive(0, { numFmt: "#,##0.000" })).numFmtId).toBe(165);
    expect(s.xf(s.derive(0, { numFmt: "0.00" })).numFmtId).toBe(2);
    expect(s.css(s.derive(1, { fill: null }))["background-color"]).toBeUndefined();
  });

  it("reports the current format for toolbar state", () => {
    const f = styles().format(1);
    expect(f).toMatchObject({ bold: true, italic: false, fontName: "Calibri", fontSize: 11, color: "#FF0000", fill: "#FFFF00", hAlign: "center" });
  });
});
