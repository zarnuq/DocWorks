import { describe, it, expect } from "vitest";
import { DOMSerializer } from "prosemirror-model";
import { latexToOmml, mathSpec, ommlToLatex } from "../src/domain/docx/math";

const M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
const MML = "http://www.w3.org/1998/Math/MathML";
const omath = (inner: string, tag = "oMath") => `<m:${tag} xmlns:m="${M}">${inner}</m:${tag}>`;
const r = (t: string) => `<m:r><m:t>${t}</m:t></m:r>`;
const render = (xml: string) => DOMSerializer.renderSpec(document, mathSpec(xml)).dom as HTMLElement;
const tags = (el: Element) => Array.from(el.querySelectorAll("*")).map((e) => e.localName);

describe("mathSpec (MathML)", () => {
  it("draws an equation as MathML, letters as identifiers, digits as numbers, the rest as operators", () => {
    const el = render(omath(r("x+12")));
    const math = el.querySelector("math")!;
    expect(math.namespaceURI).toBe(MML);
    expect(Array.from(math.querySelectorAll("mi, mn, mo")).map((e) => `${e.localName}:${e.textContent}`)).toEqual(["mi:x", "mo:+", "mn:12"]);
  });

  it("stacks fractions, scripts, roots and big operators", () => {
    const xml = omath(
      `<m:f><m:num>${r("a")}</m:num><m:den>${r("b")}</m:den></m:f>` +
        `<m:sSubSup><m:e>${r("x")}</m:e><m:sub>${r("i")}</m:sub><m:sup>${r("2")}</m:sup></m:sSubSup>` +
        `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>${r("y")}</m:e></m:rad>` +
        `<m:rad><m:deg>${r("3")}</m:deg><m:e>${r("z")}</m:e></m:rad>` +
        `<m:nary><m:naryPr><m:chr m:val="∑"/></m:naryPr><m:sub>${r("i=1")}</m:sub><m:sup>${r("n")}</m:sup><m:e>${r("i")}</m:e></m:nary>`
    );
    const t = tags(render(xml));
    for (const tag of ["mfrac", "msubsup", "msqrt", "mroot", "munderover"]) expect(t).toContain(tag);
    expect(render(xml).querySelector("munderover > mo")!.textContent).toBe("∑");
  });

  it("draws delimiters, accents, matrices and functions", () => {
    const xml = omath(
      `<m:d><m:dPr><m:begChr m:val="["/><m:endChr m:val="]"/></m:dPr><m:e>${r("a")}</m:e><m:e>${r("b")}</m:e></m:d>` +
        `<m:acc><m:accPr><m:chr m:val="⃗"/></m:accPr><m:e>${r("v")}</m:e></m:acc>` +
        `<m:m><m:mr><m:e>${r("1")}</m:e><m:e>${r("0")}</m:e></m:mr><m:mr><m:e>${r("0")}</m:e><m:e>${r("1")}</m:e></m:mr></m:m>` +
        `<m:func><m:fName>${r("sin")}</m:fName><m:e>${r("x")}</m:e></m:func>`
    );
    const el = render(xml);
    expect(Array.from(el.querySelectorAll("mo")).map((e) => e.textContent)).toEqual(expect.arrayContaining(["[", "|", "]"]));
    expect(el.querySelector("mover")).not.toBeNull();
    expect(el.querySelectorAll("mtr")).toHaveLength(2);
    expect(el.querySelector("mi[mathvariant='normal']")?.textContent).toBe("sin");
  });

  it("shows a display equation as a block", () => {
    expect(render(omath(omath(r("x")), "oMathPara")).querySelector("math")!.getAttribute("display")).toBe("block");
  });
});

describe("ommlToLatex", () => {
  it("writes equations as LaTeX", () => {
    expect(ommlToLatex(omath(`<m:f><m:num>${r("a")}</m:num><m:den>${r("b+c")}</m:den></m:f>`))).toBe("\\frac{a}{b+c}");
    expect(ommlToLatex(omath(`<m:sSup><m:e>${r("x")}</m:e><m:sup>${r("2")}</m:sup></m:sSup>`))).toBe("x^2");
    expect(ommlToLatex(omath(`<m:sSub><m:e>${r("x")}</m:e><m:sub>${r("ij")}</m:sub></m:sSub>`))).toBe("x_{ij}");
    expect(ommlToLatex(omath(r("α≤β")))).toBe("\\alpha \\leq \\beta ");
    expect(ommlToLatex(omath(`<m:rad><m:deg>${r("3")}</m:deg><m:e>${r("x")}</m:e></m:rad>`))).toBe("\\sqrt[3]{x}");
  });
});

describe("latexToOmml", () => {
  it("attaches summand and function scripts to their operands", () => {
    for (const latex of [String.raw`\sum_{i=1}^{n} x_i`, String.raw`\prod_{i=1}^{n} x_i^2`, String.raw`\sin x^2`]) {
      const xml = new DOMParser().parseFromString(latexToOmml(latex), "application/xml");
      const operator = xml.getElementsByTagNameNS(M, latex.includes("sin") ? "func" : "nary")[0];
      const operand = Array.from(operator.children).find(e => e.localName === "e")!;
      expect(Array.from(operand.children).some(e => ["sSub", "sSup", "sSubSup"].includes(e.localName))).toBe(true);
      expect(operator.parentElement?.localName).toBe("oMath");
    }
  });
  const parse = (latex: string) => new DOMParser().parseFromString(latexToOmml(latex), "application/xml").documentElement;

  it("builds Office Math XML", () => {
    const el = parse("\\frac{a}{b}");
    expect(el.localName).toBe("oMath");
    expect(el.namespaceURI).toBe(M);
    const f = el.getElementsByTagNameNS(M, "f")[0];
    expect(f.getElementsByTagNameNS(M, "num")[0].textContent).toBe("a");
    expect(f.getElementsByTagNameNS(M, "den")[0].textContent).toBe("b");
  });

  it("turns symbol commands into characters", () => {
    expect(parse("\\alpha + \\infty").textContent).toBe("α+∞");
  });

  it("builds a display equation", () => {
    expect(parse("x").localName).toBe("oMath");
    expect(new DOMParser().parseFromString(latexToOmml("x", true), "application/xml").documentElement.localName).toBe("oMathPara");
  });

  it("round-trips through LaTeX", () => {
    for (const src of [
      "x^2+y^2=z^2",
      "x_i^2",
      "\\frac{-b\\pm \\sqrt{b^2-4ac}}{2a}",
      "\\sum_{i=1}^{n}{i^2}",
      "\\int_{0}^{\\infty }{e^{-x}}",
      "\\left(a+b\\right)",
      "\\left[a,b\\right]",
      "\\hat{x}+\\vec{v}",
      "\\overline{AB}",
      "\\sin{x}+\\lim_{x\\to 0}{f}",
      "\\begin{matrix}1&0\\\\0&1\\end{matrix}",
      "\\sqrt[3]{8}",
      "\\text{if }x",
    ]) {
      expect(ommlToLatex(latexToOmml(src)), src).toBe(src);
    }
  });

  it("escapes XML characters", () => {
    expect(parse("a<b&c").textContent).toBe("a<b&c");
  });

  it("keeps unbalanced input as text instead of failing", () => {
    expect(() => latexToOmml("\\frac{a")).not.toThrow();
    expect(parse("x^").textContent).toContain("x");
  });
});
