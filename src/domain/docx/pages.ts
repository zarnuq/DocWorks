/** Page geometry and the editor's page layout (an estimate: blocks never split across pages). */

const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

export interface PageSetup {
  /** px at 96 dpi */
  width: number;
  height: number;
  top: number;
  right: number;
  bottom: number;
  left: number;
  /** Distance of the header / footer from the page edge. */
  header: number;
  footer: number;
}

/** Page size and margins from the body's final w:sectPr (Letter, 1" margins by default). */
export function pageSetup(body: Element): PageSetup {
  const sectPr = Array.from(body.children).filter((c) => c.localName === "sectPr").pop();
  const get = (local: string, name: string, fallback: number) => {
    const v = sectPr?.getElementsByTagNameNS(NS_W, local)[0]?.getAttributeNS(NS_W, name);
    return Math.round(Number(v ?? fallback) / 15);
  };
  return {
    width: get("pgSz", "w", 12240),
    height: get("pgSz", "h", 15840),
    top: Math.abs(get("pgMar", "top", 1440)),
    right: get("pgMar", "right", 1440),
    bottom: Math.abs(get("pgMar", "bottom", 1440)),
    left: get("pgMar", "left", 1440),
    header: get("pgMar", "header", 720),
    footer: get("pgMar", "footer", 720),
  };
}

/** Indices of the blocks that start a new page, given each block's height and whether a page break ends it. */
export function paginate(blocks: { height: number; breakAfter: boolean }[], pageHeight: number): number[] {
  const starts: number[] = [];
  let used = 0;
  blocks.forEach((b, i) => {
    if (i > 0 && (blocks[i - 1].breakAfter || (used > 0 && used + b.height > pageHeight))) {
      starts.push(i);
      used = 0;
    }
    used += b.height;
  });
  return starts;
}
