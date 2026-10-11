import { afterEach, describe, it, expect, vi } from "vitest";
import { XlsxWorkbook } from "../src/infrastructure/xlsx/xlsx-workbook";
import { Workbook } from "../src/domain/xlsx/workbook-model";
import { Edits } from "../src/domain/xlsx/commands";
import { Grid, GridHost, HEADER_H, HEADER_W, editText } from "../src/presentation/xlsx/grid";
import type { Op } from "../src/presentation/xlsx/paint";
import { xlsx, sheetXml } from "./xlsx-fixtures";

afterEach(() => document.body.empty());

async function mount(bytes?: Uint8Array) {
  const edits = new Edits(Workbook.open(await XlsxWorkbook.load(bytes ?? (await xlsx()))));
  const commits: [number, number, string][] = [];
  const host: GridHost = {
    sheet: () => edits.wb.sheets[0],
    isStale: () => false,
    isCircular: () => false,
    commit: (r, c, raw) => {
      commits.push([r, c, raw]);
      if (raw === "=bad(") return false;
      edits.input(edits.wb.sheets[0], r, c, raw);
      return true;
    },
    selectionChanged: vi.fn(),
    editing: vi.fn(),
    clear: (r) => edits.clear(edits.wb.sheets[0], r),
    fill: vi.fn(),
    resize: vi.fn(),
    autofit: vi.fn(),
    contextMenu: vi.fn(),
    filterMenu: vi.fn(),
  };
  const grid = new Grid(document.body.createDiv(), host);
  grid.render();
  return { grid, host, edits, commits };
}

const key = (grid: Grid, k: string, opts: KeyboardEventInit = {}) => grid.input.dispatchEvent(new KeyboardEvent("keydown", { key: k, bubbles: true, cancelable: true, ...opts }));
type TextOp = Extract<Op, { kind: "text" }>;
const textOps = (grid: Grid, quad: string) => (grid.ops.get(quad) ?? []).filter((o): o is TextOp => o.kind === "text");
const texts = (grid: Grid, quad: string) => textOps(grid, quad).map((o) => o.text);
const allTexts = (grid: Grid) => ["corner", "top", "left", "main"].flatMap((q) => textOps(grid, q));

describe("grid rendering", () => {
  it("draws cell text with formats, frozen rows in their own pane", async () => {
    const { grid } = await mount();
    expect(texts(grid, "top")).toEqual(["Name", "Notes", "Double"]);
    expect(texts(grid, "main")).toContain("2024-01-01");
    expect(texts(grid, "main")).toContain("#DIV/0!");
    expect((grid.ops.get("top") ?? []).filter((o) => o.kind === "rect")).toHaveLength(2);
    expect(textOps(grid, "top")[0]).toMatchObject({ font: expect.stringMatching(/^bold 14.6+7?px "Calibri"/), color: "#FF0000", align: "center" });
  });

  it("only creates elements for the visible part of a huge sheet", async () => {
    const rows = Array.from({ length: 20000 }, (_, i) => `<row r="${i + 1}"><c r="A${i + 1}"><v>${i}</v></c><c r="B${i + 1}"><v>${i}</v></c></row>`).join("");
    const { grid } = await mount(await xlsx({ sheets: [sheetXml(rows)] }));
    const n = allTexts(grid).length;
    expect(n).toBeGreaterThan(20);
    expect(n).toBeLessThan(200);
    expect(grid.el.querySelectorAll("*").length).toBeLessThan(2000);
  });

  it("labels headers and highlights the selected ones", async () => {
    const { grid } = await mount();
    grid.select({ row: 1, col: 1 });
    const heads = (k: string) => (grid.ops.get(k) ?? []).filter((o): o is TextOp => o.kind === "text");
    expect(heads("colheads1").slice(0, 3).map((o) => o.text)).toEqual(["A", "B", "C"]);
    const bold = [...heads("colheads1"), ...heads("rowheads0"), ...heads("rowheads1")].filter((o) => o.font.startsWith("bold")).map((o) => o.text);
    expect(bold).toEqual(["B", "2"]);
  });

  it("lets text run over empty cells to its right", async () => {
    const { grid, edits } = await mount();
    edits.input(edits.wb.sheets[0], 8, 0, "a long piece of text");
    grid.render();
    const cell = allTexts(grid).find((o) => o.text === "a long piece of text")!;
    expect(cell.w).toBeGreaterThan(64 * 3);
  });
});

describe("grid pointer", () => {
  it("finds cells and headers under the pointer, past frozen rows and hidden columns", async () => {
    const { grid } = await mount();
    expect(grid.hit(HEADER_W + 5, HEADER_H + 5)).toMatchObject({ row: 0, col: 0, area: "cell" });
    expect(grid.hit(HEADER_W + 64 + 140 + 70, HEADER_H + 25)).toMatchObject({ row: 1, col: 4, area: "cell" });
    expect(grid.hit(HEADER_W + 30, 5)).toMatchObject({ col: 0, area: "colHeader", resize: undefined });
    expect(grid.hit(HEADER_W + 62, 5)).toMatchObject({ col: 0, area: "colHeader", resize: 0 });
    expect(grid.hit(5, HEADER_H + 30)).toMatchObject({ row: 1, area: "rowHeader" });
    expect(grid.hit(5, 5)).toMatchObject({ area: "corner" });
  });
});

describe("grid keyboard", () => {
  it("moves with arrows, extends with Shift and jumps with Ctrl", async () => {
    const { grid } = await mount();
    key(grid, "ArrowDown");
    key(grid, "ArrowRight");
    expect(grid.active).toEqual({ row: 1, col: 1 });
    key(grid, "ArrowDown", { shiftKey: true });
    expect(grid.selection()).toEqual({ r1: 1, c1: 1, r2: 2, c2: 1 });
    key(grid, "ArrowDown", { ctrlKey: true });
    expect(grid.active).toEqual({ row: 4, col: 1 }); // B2..B5 are filled
    key(grid, "ArrowRight", { ctrlKey: true });
    expect(grid.active).toEqual({ row: 4, col: 2 });
    key(grid, "Home", { ctrlKey: true });
    expect(grid.active).toEqual({ row: 0, col: 0 });
  });

  it("skips hidden columns", async () => {
    const { grid } = await mount();
    grid.select({ row: 1, col: 2 });
    key(grid, "ArrowRight");
    expect(grid.active.col).toBe(4);
  });

  it("starts editing by typing, commits with Enter and moves down", async () => {
    const { grid, commits, edits } = await mount();
    grid.select({ row: 7, col: 0 });
    grid.input.value = "4";
    grid.input.dispatchEvent(new Event("input"));
    expect(grid.edit).toMatchObject({ text: "4", typed: true });
    grid.input.value = "42";
    grid.input.dispatchEvent(new Event("input"));
    key(grid, "Enter");
    expect(commits).toEqual([[7, 0, "42"]]);
    expect(edits.wb.sheets[0].cell("A8")!.value).toBe(42);
    expect(grid.active).toEqual({ row: 8, col: 0 });
    expect(grid.edit).toBeNull();
  });

  it("returns to the starting column after Tab, Tab, Enter", async () => {
    const { grid } = await mount();
    grid.select({ row: 10, col: 0 });
    for (const v of ["a", "b"]) {
      grid.input.value = v;
      grid.input.dispatchEvent(new Event("input"));
      key(grid, "Tab");
    }
    grid.input.value = "c";
    grid.input.dispatchEvent(new Event("input"));
    key(grid, "Enter");
    expect(grid.active).toEqual({ row: 11, col: 0 });
  });

  it("keeps editing when the host refuses a formula, and Escape cancels", async () => {
    const { grid, edits } = await mount();
    grid.select({ row: 1, col: 0 });
    grid.beginEdit("=bad(", false);
    key(grid, "Enter");
    expect(grid.edit).not.toBeNull();
    key(grid, "Escape");
    expect(grid.edit).toBeNull();
    expect(edits.wb.sheets[0].cell("A2")!.value).toBe(10);
  });

  it("edits an existing cell with F2, showing its formula", async () => {
    const { grid } = await mount();
    grid.select({ row: 4, col: 0 });
    key(grid, "F2");
    expect(grid.input.value).toBe("=SUM(A2:A4)");
  });

  it("clears the selection with Delete", async () => {
    const { grid, edits } = await mount();
    grid.select({ row: 1, col: 0 }, { row: 2, col: 0 });
    key(grid, "Delete");
    expect(edits.wb.sheets[0].cell("A2")).toBeUndefined();
  });

  it("grows the selection to cover merged cells", async () => {
    const { grid } = await mount();
    grid.select({ row: 4, col: 1 }, { row: 5, col: 1 });
    expect(grid.selection()).toEqual({ r1: 4, c1: 1, r2: 5, c2: 2 });
  });
});

describe("edit text", () => {
  it("shows dates, percentages and numbers so typing them back gives the same value", () => {
    const cell = (value: number) => ({ row: 0, col: 0, el: document.createElement("c"), value, style: 0 });
    expect(editText(cell(45292), "yyyy-mm-dd", false)).toBe("2024-01-01");
    expect(editText(cell(45292.5), "m/d/yyyy h:mm", false)).toBe("2024-01-01 12:00:00");
    expect(editText(cell(0.25), "0%", false)).toBe("25%");
    expect(editText(cell(1 / 3), "0.00", false)).toBe("0.333333333333333");
  });
});

describe("grid dark theme", () => {
  it("shows black text on unfilled cells and white fills in theme colors, keeping text on colored fills dark", async () => {
    const { grid, edits } = await mount();
    const s = edits.wb.sheets[0];
    edits.input(s, 8, 0, "plain");
    edits.format(s, { r1: 8, c1: 0, r2: 8, c2: 0 }, { color: "#000000" });
    edits.input(s, 9, 0, "white");
    edits.format(s, { r1: 9, c1: 0, r2: 9, c2: 0 }, { fill: "#FFFFFF" });
    grid.dark = true;
    grid.invalidate();
    const cell = (t: string) => allTexts(grid).find((o) => o.text === t)!;
    expect(cell("plain").color).toBe("#e3e3e3");
    expect(cell("Name").color).toBe("#FF0000");
    const fills = ["top", "main"].flatMap((q) => grid.ops.get(q) ?? []).filter((o) => o.kind === "rect").map((o) => (o as { color: string }).color);
    expect(fills).not.toContain("#FFFFFF");
    grid.dark = false;
    grid.invalidate();
    expect(cell("plain").color).toBe("#000000");
  });
});

describe("grid header context menu", () => {
  it("selects the row under a right-clicked header when it isn't selected", async () => {
    const { grid, host } = await mount();
    grid.select({ row: 3, col: 7 }, { row: 5, col: 7 });
    grid.el.querySelector(".xlsx-canvas")!.dispatchEvent(new MouseEvent("contextmenu", { clientX: 5, clientY: 24 + 160 + 10 /* row 9: row 4 is 40px tall, row 6 hidden */, bubbles: true, cancelable: true }));
    expect(grid.selection()).toEqual({ r1: 8, c1: 0, r2: 8, c2: 16383 });
    expect(host.contextMenu).toHaveBeenCalled();
  });
});

describe("grid pictures", () => {
  it("draws visible pictures over the cells with their image URL", async () => {
    const { grid, host } = await mount();
    host.pictures = () => [
      { name: "logo", alt: "", image: "xl/media/a.png", x: 70, y: 30, w: 40, h: 20 },
      { name: "far", alt: "", image: "xl/media/b.png", x: 90000, y: 90000, w: 40, h: 20 },
    ];
    host.imageUrl = (p) => `blob:${p}`;
    grid.render();
    const imgs = Array.from(grid.el.querySelectorAll<HTMLImageElement>(".xlsx-picture img"));
    expect(imgs.map((i) => i.getAttribute("src"))).toEqual(["blob:xl/media/a.png"]);
    expect(imgs[0].alt).toBe("logo");
  });
});

describe("charts in the grid", () => {
  it("opens a chart's menu once the click is over (opening it on press lets the release close it)", async () => {
    const { grid, host } = await mount();
    const chartMenu = vi.fn();
    Object.assign(host, { charts: () => [{ key: "xl/charts/chart1.xml", x: 10, y: 10, w: 100, h: 80, draw: (el: HTMLElement) => el.createDiv({ cls: "svg-stub" }) }], chartMenu });
    grid.render();
    const el = grid.el.querySelector<HTMLElement>(".xlsx-chart .svg-stub")!;
    el.dispatchEvent(new PointerEvent("pointerdown", { bubbles: true, button: 0 }));
    expect(chartMenu).not.toHaveBeenCalled();
    el.dispatchEvent(new MouseEvent("click", { bubbles: true, button: 0 }));
    expect(chartMenu).toHaveBeenCalledWith("xl/charts/chart1.xml", expect.anything());
  });
});

describe("notes in the grid", () => {
  it("marks cells with a note in their top right corner, empty cells too", async () => {
    const { grid, host } = await mount();
    // B2 has a value; A5 is empty.
    host.note = (row, col) => (row === 1 && col === 1) || (row === 4 && col === 0) ? "a note" : null;
    grid.invalidate();
    grid.render();
    const corners = ["corner", "top", "left", "main"].flatMap((q) => (grid.ops.get(q) ?? []).filter((o) => o.kind === "corner" && o.color === "#000000"));
    expect(corners).toHaveLength(2);
  });
});

describe("zoom", () => {
  it("finds the cell under the pointer at any zoom, and paints sharply", async () => {
    const { grid, edits } = await mount();
    const s = edits.wb.sheets[0];
    // The middle of B3 at 100%...
    const x = HEADER_W + s.colX(1) + s.colWidth(1) / 2, y = HEADER_H + s.rowY(2) + s.rowHeight(2) / 2;
    expect(grid.hit(x, y)).toMatchObject({ area: "cell", row: 2, col: 1 });
    // ...is twice as far from the corner at 200%.
    grid.setZoom(2);
    expect(grid.zoom).toBe(2);
    expect(grid.hit(x * 2, y * 2)).toMatchObject({ area: "cell", row: 2, col: 1 });
    expect(grid.el.style.getPropertyValue("zoom")).toBe("2");
  });
});
