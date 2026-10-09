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

export type PageSize = "document" | "letter";

/** Page size and margins from the body's final w:sectPr (Letter, 1" margins by default); `letter` forces 8.5 x 11 in pages. */
export function pageSetup(body: Element, size: PageSize = "document"): PageSetup {
  const sectPr = Array.from(body.children).filter((c) => c.localName === "sectPr").pop();
  const get = (local: string, name: string, fallback: number) => {
    const v = sectPr?.getElementsByTagNameNS(NS_W, local)[0]?.getAttributeNS(NS_W, name);
    return Math.round(Number(v ?? fallback) / 15);
  };
  return {
    width: size === "letter" ? 816 : get("pgSz", "w", 12240),
    height: size === "letter" ? 1056 : get("pgSz", "h", 15840),
    top: Math.abs(get("pgMar", "top", 1440)),
    right: get("pgMar", "right", 1440),
    bottom: Math.abs(get("pgMar", "bottom", 1440)),
    left: get("pgMar", "left", 1440),
    header: get("pgMar", "header", 720),
    footer: get("pgMar", "footer", 720),
  };
}

export interface PageBlock {
  height: number;
  /** A manual page break ends the block. */
  breakAfter: boolean;
  /** Word can split it across pages (a table's rows); it starts where it is and runs on instead of jumping. */
  splittable?: boolean;
}

export interface PageLayout {
  /** Indices of the blocks that start a new drawn page. */
  starts: number[];
  /** Page number of each drawn page (a table running on past a page bottom skips numbers). */
  pageNumbers: number[];
  /** Blank space left at the bottom of each drawn page. */
  fills: number[];
  pages: number;
}

/** Lay blocks out on pages (an estimate: only splittable blocks cross a page bottom, and no page break is drawn inside them). */
export function layoutPages(blocks: PageBlock[], pageHeight: number): PageLayout {
  const starts: number[] = [];
  const pageNumbers = [1];
  const fills: number[] = [];
  let used = 0;
  let page = 1;
  blocks.forEach((b, i) => {
    if (i > 0 && (blocks[i - 1].breakAfter || (used > 0 && used + b.height > pageHeight && !b.splittable))) {
      fills.push(Math.max(0, pageHeight - used));
      starts.push(i);
      pageNumbers.push(++page);
      used = 0;
    }
    used += b.height;
    if (b.splittable) {
      while (used > pageHeight) {
        used -= pageHeight;
        page++;
      }
    }
  });
  fills.push(Math.max(0, pageHeight - used));
  return { starts, pageNumbers, fills, pages: page };
}
