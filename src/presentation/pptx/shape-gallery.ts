/*
 * Shape picker: every DrawingML preset, grouped like PowerPoint's Shapes menu,
 * each with a small preview drawn by the geometry engine. Plain DOM.
 */
import { PRESET_NAMES, presetGeometry } from "../../domain/pptx/geometry";

export interface ShapeChoice {
  prst: string;
  /** Line with an arrowhead at its end. */
  tailArrow?: boolean;
}

const CATEGORIES: [string, string[]][] = [
  ["Lines", ["line", "straightConnector1", "bentConnector3", "curvedConnector3"]],
  ["Rectangles", ["rect", "roundRect", "snip1Rect", "snip2SameRect", "snip2DiagRect", "snipRoundRect", "round1Rect", "round2SameRect", "round2DiagRect"]],
  [
    "Basic shapes",
    [
      "ellipse", "triangle", "rtTriangle", "parallelogram", "trapezoid", "nonIsoscelesTrapezoid", "diamond", "pentagon", "hexagon", "heptagon",
      "octagon", "decagon", "dodecagon", "pie", "chord", "teardrop", "frame", "halfFrame", "corner", "diagStripe", "plus", "plaque", "can",
      "cube", "bevel", "donut", "noSmoking", "blockArc", "foldedCorner", "smileyFace", "heart", "lightningBolt", "sun", "moon", "cloud", "arc",
      "bracketPair", "bracePair", "leftBracket", "rightBracket", "leftBrace", "rightBrace",
    ],
  ],
  [
    "Block arrows",
    [
      "rightArrow", "leftArrow", "upArrow", "downArrow", "leftRightArrow", "upDownArrow", "quadArrow", "leftRightUpArrow", "bentArrow",
      "uturnArrow", "leftUpArrow", "bentUpArrow", "curvedRightArrow", "curvedLeftArrow", "curvedUpArrow", "curvedDownArrow",
      "stripedRightArrow", "notchedRightArrow", "homePlate", "chevron", "rightArrowCallout", "downArrowCallout", "leftArrowCallout",
      "upArrowCallout", "leftRightArrowCallout", "quadArrowCallout", "circularArrow",
    ],
  ],
  ["Equation shapes", ["mathPlus", "mathMinus", "mathMultiply", "mathDivide", "mathEqual", "mathNotEqual"]],
  ["Flowchart", PRESET_NAMES.filter((n) => n.startsWith("flowChart"))],
  [
    "Stars and banners",
    [
      "irregularSeal1", "irregularSeal2", "star4", "star5", "star6", "star7", "star8", "star10", "star12", "star16", "star24", "star32",
      "ribbon2", "ribbon", "ellipseRibbon2", "ellipseRibbon", "verticalScroll", "horizontalScroll", "wave", "doubleWave",
    ],
  ],
  [
    "Callouts",
    [
      "wedgeRectCallout", "wedgeRoundRectCallout", "wedgeEllipseCallout", "cloudCallout", "borderCallout1", "borderCallout2", "borderCallout3",
      "accentCallout1", "accentCallout2", "accentCallout3", "callout1", "callout2", "callout3", "accentBorderCallout1", "accentBorderCallout2",
      "accentBorderCallout3",
    ],
  ],
  ["Action buttons", PRESET_NAMES.filter((n) => n.startsWith("actionButton"))],
];

// Everything else (extra connector variants, tabs, gears...) so every preset is reachable.
const listed = new Set(CATEGORIES.flatMap(([, names]) => names));
CATEGORIES.push(["Other", PRESET_NAMES.filter((n) => !listed.has(n))]);

/** "flowChartDecision" -> "Flow chart decision" */
export const shapeLabel = (name: string) =>
  name === "straightConnector1"
    ? "Arrow"
    : name
        .replace(/([a-z])([A-Z0-9])/g, "$1 $2")
        .toLowerCase()
        .replace(/^./, (c) => c.toUpperCase());

const SVG_NS = "http://www.w3.org/2000/svg";

function preview(name: string): SVGSVGElement {
  const svg = document.createElementNS(SVG_NS, "svg");
  svg.setAttribute("viewBox", "-2 -2 36 28");
  for (const p of presetGeometry(name, 32, 24)?.paths ?? []) {
    const path = svg.appendChild(document.createElementNS(SVG_NS, "path"));
    path.setAttribute("d", p.d);
    path.setAttribute("class", p.fill === "none" ? "is-outline" : "is-filled");
    if (!p.stroke) path.setAttribute("stroke", "none");
  }
  return svg;
}

/** Build the picker; `onPick` gets the chosen shape. */
export function buildShapeGallery(host: HTMLElement, onPick: (choice: ShapeChoice) => void): void {
  for (const [title, names] of CATEGORIES) {
    if (!names.length) continue;
    host.createDiv({ cls: "pptx-gallery-title", text: title });
    const grid = host.createDiv({ cls: "pptx-gallery-grid" });
    for (const name of names) {
      const btn = grid.createEl("button", { cls: "pptx-gallery-item", attr: { "aria-label": shapeLabel(name), "data-prst": name } });
      btn.appendChild(preview(name));
      btn.addEventListener("click", () => onPick({ prst: name, tailArrow: name === "straightConnector1" }));
    }
  }
}
