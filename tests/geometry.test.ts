import { describe, it, expect } from "vitest";
import { customGeometry, presetGeometry, PRESET_NAMES, solveAdjust } from "../src/domain/pptx/geometry";

const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const custGeom = (gd: string, path: string) =>
  new DOMParser().parseFromString(
    `<a:custGeom xmlns:a="${A}"><a:avLst/><a:gdLst>${gd}</a:gdLst><a:rect l="l" t="t" r="r" b="b"/><a:pathLst>${path}</a:pathLst></a:custGeom>`,
    "application/xml"
  ).documentElement;

describe("presetGeometry", () => {
  it("knows every ECMA-376 preset", () => {
    expect(PRESET_NAMES.length).toBe(187);
    for (const name of PRESET_NAMES) {
      const g = presetGeometry(name, 1000, 600)!;
      expect(g.paths.length, name).toBeGreaterThan(0);
      for (const p of g.paths) expect(p.d, name).toMatch(/^M[MLACQZ0-9 .-]*$/);
    }
  });

  it("draws simple outlines", () => {
    expect(presetGeometry("rect", 10, 10)!.paths[0].d).toBe("M0 0 L10 0 L10 10 L0 10 Z");
    expect(presetGeometry("triangle", 100, 50)!.paths[0].d).toBe("M0 50 L50 0 L100 50 Z");
  });

  it("applies adjust values over the defaults", () => {
    expect(presetGeometry("chevron", 1000, 400, { adj: 30000 })!.paths[0].d).toBe("M0 0 L880 0 L1000 200 L880 400 L0 400 L120 200 Z");
    expect(presetGeometry("chevron", 1000, 400)!.paths[0].d).toBe("M0 0 L800 0 L1000 200 L800 400 L0 400 L200 200 Z");
  });

  it("draws arcs as SVG elliptical arcs", () => {
    const d = presetGeometry("ellipse", 100, 50)!.paths[0].d;
    expect(d.startsWith("M0 25 A50 25 0 0 1 50 0")).toBe(true);
    expect(d.endsWith("Z")).toBe(true);
  });

  it("keeps path fill modes for shaded parts", () => {
    expect(presetGeometry("can", 100, 100)!.paths.map((p) => p.fill)).toEqual(["norm", "lighten", "none"]);
    expect(presetGeometry("can", 100, 100)!.paths.map((p) => p.stroke)).toEqual([false, false, true]);
  });

  it("returns the text rectangle, connection sites and adjust handles", () => {
    const e = presetGeometry("ellipse", 100, 50)!;
    expect(e.textRect.x).toBeCloseTo(14.645, 2);
    expect(e.textRect.y).toBeCloseTo(7.322, 2);
    expect(e.textRect.w).toBeCloseTo(70.711, 2);
    expect(presetGeometry("rect", 10, 20)!.sites).toEqual([
      { x: 5, y: 0, ang: 270 },
      { x: 0, y: 10, ang: 180 },
      { x: 5, y: 20, ang: 90 },
      { x: 10, y: 10, ang: 0 },
    ]);
    const [h] = presetGeometry("chevron", 1000, 400)!.handles;
    expect(h).toMatchObject({ x: 800, y: 0, refX: "adj", minX: 0, maxX: 250000, refY: null });
  });

  it("returns null for unknown presets", () => {
    expect(presetGeometry("notAShape", 10, 10)).toBeNull();
  });
});

describe("customGeometry", () => {
  it("evaluates trigonometric guide formulas", () => {
    const g = customGeometry(
      custGeom(
        `<a:gd name="a" fmla="at2 1 1"/><a:gd name="x" fmla="cos 100 a"/><a:gd name="y" fmla="mod 3 4 0"/><a:gd name="z" fmla="sat2 10 1 1"/>`,
        `<a:path w="200" h="200"><a:moveTo><a:pt x="x" y="y"/></a:moveTo><a:lnTo><a:pt x="z" y="0"/></a:lnTo></a:path>`
      ),
      200,
      200
    );
    expect(g.paths[0].d).toBe("M70.711 5 L7.071 0");
  });

  it("continues arcs from the current point", () => {
    const g = customGeometry(
      custGeom("", `<a:path w="100" h="100"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:arcTo wR="10" hR="10" stAng="0" swAng="5400000"/></a:path>`),
      100,
      100
    );
    expect(g.paths[0].d).toBe("M0 0 A10 10 0 0 1 -10 10");
  });
});

describe("solveAdjust", () => {
  it("finds the adjust value that puts a handle under the pointer", () => {
    const adj = solveAdjust("chevron", 1000, 400, {}, 0, { x: 880, y: 0 });
    expect(adj.adj).toBeCloseTo(30000, -1);
  });

  it("clamps to the handle's range", () => {
    expect(solveAdjust("chevron", 1000, 400, {}, 0, { x: -5000, y: 0 }).adj).toBeCloseTo(250000, -1);
    expect(solveAdjust("chevron", 1000, 400, {}, 0, { x: 5000, y: 0 }).adj).toBeCloseTo(0, -1);
  });
});
