import { describe, it, expect } from "vitest";
import { buildShapeGallery, shapeLabel } from "../src/presentation/pptx/shape-gallery";
import { PRESET_NAMES } from "../src/domain/pptx/geometry";

describe("shape gallery", () => {
  it("offers every preset once, with a preview", () => {
    const host = document.createElement("div");
    buildShapeGallery(host, () => undefined);
    const items = Array.from(host.querySelectorAll<HTMLElement>(".pptx-gallery-item"));
    const names = items.map((b) => b.dataset.prst);
    expect(new Set(names).size).toBe(names.length);
    expect(new Set(names)).toEqual(new Set(PRESET_NAMES));
    expect(items.every((b) => b.querySelector("svg path"))).toBe(true);
  });

  it("reports the picked shape, with an arrowhead for the arrow line", () => {
    const host = document.createElement("div");
    const picked: unknown[] = [];
    buildShapeGallery(host, (c) => picked.push(c));
    host.querySelector<HTMLElement>('[data-prst="straightConnector1"]')!.click();
    host.querySelector<HTMLElement>('[data-prst="ellipse"]')!.click();
    expect(picked).toEqual([{ prst: "straightConnector1", tailArrow: true }, { prst: "ellipse", tailArrow: false }]);
  });

  it("labels shapes readably", () => {
    expect(shapeLabel("flowChartDecision")).toBe("Flow chart decision");
    expect(shapeLabel("star5")).toBe("Star 5");
  });
});
