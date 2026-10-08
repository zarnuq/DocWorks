import { describe, it, expect } from "vitest";
import { ArrangeControl, buildArrangePanel, buildShapePanel, PanelControl } from "../src/presentation/pptx/shape-panel";
import { addShape } from "../src/domain/pptx/slide-edit";
import { parseSlide, ShapeModel } from "../src/domain/pptx/slide-model";
import { loadSample } from "./helpers";

const A = "http://schemas.openxmlformats.org/drawingml/2006/main";

/** A panel for a fresh ellipse; edits apply straight to the XML and rebuild the panel. */
async function setup(prst = "ellipse") {
  const doc = await loadSample();
  const path = doc.slides()[1].path;
  const el = addShape(doc.xml(path)!, prst, { x: 914400, y: 914400, w: 1828800, h: 914400 });
  const host = document.createElement("div");
  const model = (): ShapeModel => parseSlide(doc, path).shapes.find((s) => s.el === el)!;
  const ctl: PanelControl = {
    lockAspect: false,
    edit(fn) {
      fn(el);
      buildShapePanel(host, model(), ctl);
    },
  };
  buildShapePanel(host, model(), ctl);
  const field = <T extends HTMLElement>(name: string, nth = 0) => host.querySelectorAll<T>(`[data-field="${name}"]`)[nth];
  const change = (name: string, value: string, nth = 0) => {
    const f = field<HTMLInputElement>(name, nth);
    if (f.type === "checkbox") f.checked = value === "true";
    else f.value = value;
    f.dispatchEvent(new Event("change"));
  };
  const button = (text: string) => Array.from(host.querySelectorAll("button")).find((b) => b.textContent === text)!.click();
  return { el, model, change, button, field };
}

describe("shape format panel", () => {
  it("switches fills and sets color and transparency", async () => {
    const { el, model, change } = await setup();
    change("Type", "none");
    expect(model().fill).toBeNull();
    expect(el.getElementsByTagNameNS(A, "noFill")).toHaveLength(1);
    change("Type", "solid");
    change("Color", "#ff0000");
    change("Transparency %", "50");
    expect(model().fill).toEqual({ type: "solid", color: "#FF000080" });
    change("Type", "linear");
    change("Angle", "45");
    expect(model().fill).toMatchObject({ type: "gradient", radial: false, angle: 45, stops: [{ color: "#FF000080" }, { color: "#FFFFFF" }] });
  });

  it("edits the outline", async () => {
    const { model, change } = await setup();
    change("Color", "#00ff00", 1);
    change("Width (pt)", "3");
    change("Dash", "dash");
    change("End arrow", "stealth");
    expect(model().line).toMatchObject({ color: "#00FF00", width: 38100, dash: "dash", tail: { type: "stealth" } });
    change("Type", "none", 1);
    expect(model().line).toBeNull();
  });

  it("toggles the shadow", async () => {
    const { model, change } = await setup();
    change("Shadow", "true");
    expect(model().shadow).toMatchObject({ dist: 38100 });
    change("Shadow", "false");
    expect(model().shadow).toBeNull();
  });

  it("moves, sizes, rotates and flips", async () => {
    const { model, change, button } = await setup();
    change("X (in)", "2");
    change("Width (in)", "3");
    expect(model().rect).toEqual({ x: 1828800, y: 914400, w: 2743200, h: 914400 });
    change("Lock aspect ratio", "true");
    change("Height (in)", "2");
    expect(model().rect).toMatchObject({ w: 5486400, h: 1828800 });
    change("Rotation (°)", "45");
    button("Flip horizontal");
    expect(model()).toMatchObject({ rot: 45, flipH: true, flipV: false });
  });

  it("changes the stacking order", async () => {
    const { el, button } = await setup();
    button("Send to back");
    const tree = el.parentElement!;
    expect(Array.from(tree.children).filter((c) => !["nvGrpSpPr", "grpSpPr"].includes(c.localName))[0]).toBe(el);
    button("Bring to front");
    expect(tree.lastElementChild).toBe(el);
  });

  it("has no fill section for lines", async () => {
    const { field } = await setup("line");
    expect(field("Transparency %")).toBeUndefined();
    expect(field("End arrow")).toBeDefined();
  });
});

describe("text box settings", () => {
  it("edits anchor, autofit, wrapping and margins", async () => {
    const { model, change } = await setup("rect");
    change("Vertical alignment", "b");
    change("Autofit", "resize");
    change("Wrap text", "false");
    change("Left margin (in)", "0");
    expect(model().text).toMatchObject({ anchor: "b", autofit: "resize", wrap: false, insets: { l: 0, t: 45720, r: 91440, b: 45720 } });
  });
});

describe("arrange panel", () => {
  it("offers align, distribute and grouping for a multi-selection", () => {
    const calls: unknown[][] = [];
    const host = document.createElement("div");
    const ctl: ArrangeControl = {
      count: 3,
      groups: 1,
      align: (m) => calls.push(["align", m]),
      distribute: (a) => calls.push(["distribute", a]),
      group: () => calls.push(["group"]),
      ungroup: () => calls.push(["ungroup"]),
      order: (to) => calls.push(["order", to]),
    };
    buildArrangePanel(host, ctl);
    for (const text of ["Left", "Middle", "Horizontally", "Group", "Ungroup", "Send to back"]) {
      Array.from(host.querySelectorAll("button")).find((b) => b.textContent === text)!.click();
    }
    expect(calls).toEqual([["align", "left"], ["align", "middle"], ["distribute", "horizontal"], ["group"], ["ungroup"], ["order", "back"]]);
  });

  it("hides distribute and group for a single shape", () => {
    const host = document.createElement("div");
    const noop = () => undefined;
    buildArrangePanel(host, { count: 1, groups: 0, align: noop, distribute: noop, group: noop, ungroup: noop, order: noop });
    const labels = Array.from(host.querySelectorAll("button")).map((b) => b.textContent);
    expect(labels).not.toContain("Horizontally");
    expect(labels).not.toContain("Group");
    expect(host.textContent).toContain("Align to slide");
  });
});
