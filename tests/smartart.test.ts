import { describe, expect, it } from "vitest";
import { setSmartArtText, smartArtItems } from "../src/domain/pptx/smartart";

const DGM = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const DSP = "http://schemas.microsoft.com/office/drawing/2008/diagram";
const t = (text: string) => `<dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" b="1"/><a:t>${text}</a:t></a:r></a:p></dgm:t>`;
const pt = (id: string, text: string, type = "") => `<dgm:pt modelId="${id}"${type ? ` type="${type}"` : ""}><dgm:prSet/><dgm:spPr/>${t(text)}</dgm:pt>`;
const pres = (id: string, of: string) => `<dgm:pt modelId="${id}" type="pres"><dgm:prSet presAssocID="${of}" presName="text"/><dgm:spPr/></dgm:pt>`;
const cxn = (src: string, dest: string, ord: number) => `<dgm:cxn modelId="c${src}${dest}" srcId="${src}" destId="${dest}" srcOrd="${ord}" destOrd="0"/>`;
// A list "Plan" (with child "Budget") and "Build", in that order though stored out of order.
const data = () => new DOMParser().parseFromString(
  `<dgm:dataModel xmlns:dgm="${DGM}" xmlns:a="${A}"><dgm:ptLst>${pt("doc", "", "doc")}${pt("b", "Build")}${pt("a", "Plan")}${pt("a1", "Budget")}${pres("pa", "a")}${pres("pb", "b")}</dgm:ptLst>` +
  `<dgm:cxnLst>${cxn("doc", "b", 1)}${cxn("doc", "a", 0)}${cxn("a", "a1", 0)}</dgm:cxnLst></dgm:dataModel>`, "application/xml");
const drawing = () => new DOMParser().parseFromString(
  `<dsp:drawing xmlns:dsp="${DSP}" xmlns:a="${A}"><dsp:spTree><dsp:sp modelId="pa"><dsp:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="2400"/><a:t>Plan</a:t></a:r></a:p></dsp:txBody></dsp:sp><dsp:sp modelId="pb"><dsp:txBody><a:bodyPr/><a:p><a:r><a:t>Build</a:t></a:r></a:p></dsp:txBody></dsp:sp></dsp:spTree></dsp:drawing>`, "application/xml");

describe("SmartArt text", () => {
  it("lists the items in order, with their level", () => {
    expect(smartArtItems(data())).toEqual([
      { id: "a", text: "Plan", depth: 0 },
      { id: "a1", text: "Budget", depth: 1 },
      { id: "b", text: "Build", depth: 0 },
    ]);
  });

  it("changes an item's text in the data and in the drawing shown, keeping its formatting", () => {
    const d = data(), dr = drawing();
    setSmartArtText(d, dr, "a", "Plan it\nthen check");
    expect(smartArtItems(d)[0].text).toBe("Plan it\nthen check");
    const shown = dr.getElementsByTagNameNS(DSP, "sp")[0];
    expect(Array.from(shown.getElementsByTagNameNS(A, "p")).map((p) => p.textContent)).toEqual(["Plan it", "then check"]);
    // The run keeps its size; the other item is untouched.
    expect(shown.getElementsByTagNameNS(A, "rPr")[0].getAttribute("sz")).toBe("2400");
    expect(dr.getElementsByTagNameNS(DSP, "sp")[1].textContent).toBe("Build");
    const model = Array.from(d.getElementsByTagNameNS(DGM, "pt")).find((p) => p.getAttribute("modelId") === "a")!;
    expect(model.getElementsByTagNameNS(A, "rPr")[0].getAttribute("b")).toBe("1");
  });
});
