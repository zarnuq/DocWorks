import { readFileSync } from "fs";
import { join } from "path";
import { describe, it, expect, beforeAll } from "vitest";
import { DOMSerializer } from "prosemirror-model";
import { docxSchema as schema } from "../src/domain/docx/schema";
import { readDocx } from "../src/domain/docx/docx-reader";
import { W } from "./helpers";

// A run's own text colour must survive its highlight / shading: a light-grey
// command on a black highlight (Word's "code" look) or white heading text on a
// maroon bar. Only highlighted text with no colour of its own is kept dark, so
// it stays readable on the dark paper theme.
const runs =
  `<w:r><w:rPr><w:color w:val="D9D9D9" w:themeColor="background1" w:themeShade="D9"/><w:highlight w:val="black"/></w:rPr><w:t>show interfaces</w:t></w:r>` +
  `<w:r><w:rPr><w:color w:val="ffffff"/><w:shd w:fill="660033" w:val="clear"/></w:rPr><w:t>SYSTEM OVERVIEW</w:t></w:r>` +
  `<w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t>plain highlight</w:t></w:r>`;

function render(theme: "light" | "dark"): (text: string) => string {
  const el = new DOMParser().parseFromString(`<w:body xmlns:w="${W}"><w:p>${runs}</w:p></w:body>`, "application/xml").documentElement;
  const { doc } = readDocx(el, { imageSrc: () => null, hyperlinkHref: () => null });
  const root = document.body.createDiv({ cls: `docx-theme-${theme}` }).createDiv({ cls: "docx-editor-host" }).createDiv({ cls: "ProseMirror" });
  root.append(DOMSerializer.fromSchema(schema).serializeFragment(doc.content));
  return (text) => {
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    for (let n = walker.nextNode(); n; n = walker.nextNode()) if (n.textContent === text) return getComputedStyle(n.parentElement!).color;
    throw new Error(`no text ${text}`);
  };
}

describe("text colour inside highlights and run shading", () => {
  beforeAll(() => {
    const style = document.createElement("style");
    style.textContent = readFileSync(join(__dirname, "..", "styles.css"), "utf8");
    document.head.append(style);
  });

  for (const theme of ["light", "dark"] as const) {
    it(`keeps a run's own colour on its highlight (${theme} paper)`, () => {
      const color = render(theme);
      expect(color("show interfaces")).toBe("rgb(217, 217, 217)");
      expect(color("SYSTEM OVERVIEW")).toBe("rgb(255, 255, 255)");
    });
  }

  it("keeps highlighted text without a colour dark on dark paper", () => {
    expect(render("dark")("plain highlight")).toBe("rgb(26, 26, 26)");
  });
});
