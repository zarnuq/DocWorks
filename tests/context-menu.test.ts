import { afterEach, describe, expect, it, vi } from "vitest";
import { closeContextMenu, showContextMenu, type CtxEntry } from "../src/presentation/common/context-menu";

afterEach(() => document.body.empty());

const titles = () => Array.from(document.querySelectorAll(".dw-ctx > .menu .menu-item, .dw-ctx > .menu .menu-separator")).map((e) => (e.classList.contains("menu-separator") ? "---" : e.querySelector(".menu-item-title")!.textContent));
const item = (label: string, root: ParentNode = document) => Array.from(root.querySelectorAll<HTMLElement>(".menu-item")).find((e) => e.querySelector(".menu-item-title")!.textContent === label)!;
const key = (k: string) => document.querySelector(".dw-ctx .menu")!.dispatchEvent(new KeyboardEvent("keydown", { key: k, bubbles: true }));

describe("own context menu", () => {
  it("lists items and separators in order, dropping leading, trailing and doubled separators", () => {
    showContextMenu({ x: 10, y: 10 }, ["---", { label: "Cut" }, "---", "---", { label: "Copy" }, "---"]);
    expect(titles()).toEqual(["Cut", "---", "Copy"]);
  });

  it("runs a clicked item and closes, menu and mini toolbar together", () => {
    const run = vi.fn(), onClose = vi.fn();
    showContextMenu({ x: 10, y: 10 }, [{ label: "Cut", run }], { mini: (el) => el.createEl("button", { text: "B" }), onClose });
    expect(document.querySelector(".dw-ctx-mini button")!.textContent).toBe("B");
    item("Cut").click();
    expect(run).toHaveBeenCalledOnce();
    expect(document.querySelector(".dw-ctx")).toBeNull();
    expect(document.querySelector(".dw-ctx-mini")).toBeNull();
    expect(onClose).toHaveBeenCalledOnce();
  });

  it("doesn't run disabled items", () => {
    const run = vi.fn();
    showContextMenu({ x: 0, y: 0 }, [{ label: "Paste", run, disabled: true }]);
    item("Paste").click();
    expect(run).not.toHaveBeenCalled();
    expect(item("Paste").classList.contains("is-disabled")).toBe(true);
  });

  it("moves with the arrow keys past disabled items, runs with Enter, and opens submenus with the right arrow", () => {
    const run = vi.fn(), sub = vi.fn();
    const entries: CtxEntry[] = [{ label: "Cut", run }, { label: "Copy", disabled: true }, { label: "Insert", sub: [{ label: "Rows Above", run: sub }] }];
    showContextMenu({ x: 0, y: 0 }, entries);
    key("ArrowDown");
    expect(item("Cut").classList.contains("selected")).toBe(true);
    key("ArrowDown");
    expect(item("Insert").classList.contains("selected")).toBe(true);
    key("ArrowRight");
    const submenu = document.querySelector<HTMLElement>(".dw-ctx-sub")!;
    expect(submenu.textContent).toContain("Rows Above");
    expect(item("Rows Above", submenu).classList.contains("selected")).toBe(true);
    submenu.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowLeft", bubbles: true }));
    expect(document.querySelector(".dw-ctx-sub")).toBeNull();
    key("ArrowUp");
    key("ArrowUp");
    key("Enter");
    expect(run).toHaveBeenCalledOnce();
  });

  it("Escape closes everything", () => {
    showContextMenu({ x: 0, y: 0 }, [{ label: "Cut" }], { mini: () => {} });
    key("Escape");
    expect(document.querySelector(".dw-ctx")).toBeNull();
  });

  it("closes on a mousedown outside, on window blur and on scroll", () => {
    for (const fire of [() => document.body.dispatchEvent(new MouseEvent("mousedown", { bubbles: true })), () => window.dispatchEvent(new Event("blur")), () => document.dispatchEvent(new Event("scroll"))]) {
      showContextMenu({ x: 0, y: 0 }, [{ label: "Cut" }]);
      fire();
      expect(document.querySelector(".dw-ctx")).toBeNull();
    }
  });

  it("stays open on a mousedown inside it or its mini toolbar", () => {
    showContextMenu({ x: 0, y: 0 }, [{ label: "Cut" }], { mini: (el) => el.createEl("button", { text: "B" }) });
    document.querySelector(".dw-ctx-mini button")!.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
    item("Cut").dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
    expect(document.querySelector(".dw-ctx")).not.toBeNull();
  });

  it("is placed at the click, moved left / up when it would leave the window", () => {
    showContextMenu({ x: 20, y: 30 }, [{ label: "Cut" }]);
    const menu = document.querySelector<HTMLElement>(".dw-ctx .menu")!;
    expect([menu.style.left, menu.style.top]).toEqual(["20px", "30px"]);
    vi.spyOn(HTMLElement.prototype, "getBoundingClientRect").mockReturnValue({ width: 200, height: 300, left: 0, top: 0, right: 200, bottom: 300, x: 0, y: 0, toJSON: () => ({}) });
    showContextMenu({ x: window.innerWidth - 50, y: window.innerHeight - 10 }, [{ label: "Cut" }]);
    const moved = document.querySelector<HTMLElement>(".dw-ctx .menu")!;
    expect(parseFloat(moved.style.left)).toBeLessThanOrEqual(window.innerWidth - 200);
    expect(parseFloat(moved.style.top)).toBeLessThanOrEqual(window.innerHeight - 300);
  });

  it("closes the menu that was open when another opens", () => {
    showContextMenu({ x: 0, y: 0 }, [{ label: "One" }]);
    showContextMenu({ x: 0, y: 0 }, [{ label: "Two" }]);
    expect(document.querySelectorAll(".dw-ctx").length).toBe(1);
    expect(titles()).toEqual(["Two"]);
  });

  it("can be closed by its owner", () => {
    const m = showContextMenu({ x: 0, y: 0 }, [{ label: "One" }]);
    m.close();
    expect(document.querySelector(".dw-ctx")).toBeNull();
  });

  it("keeps the document's selection: mousedown on the menu or mini toolbar does nothing by default", () => {
    showContextMenu({ x: 0, y: 0 }, [{ label: "Cut" }], { mini: (el) => el.createEl("button", { text: "B" }) });
    for (const target of [document.querySelector(".dw-ctx-mini button")!, item("Cut")]) {
      const e = new MouseEvent("mousedown", { bubbles: true, cancelable: true });
      target.dispatchEvent(e);
      expect(e.defaultPrevented).toBe(true);
    }
  });
});

describe("own context menu, review fixes", () => {
  it("Escape closes it even when focus went back to the document", () => {
    showContextMenu({ x: 0, y: 0 }, [{ label: "Cut" }], { mini: (el) => el.createEl("button", { text: "B" }) });
    document.body.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
    expect(document.querySelector(".dw-ctx")).toBeNull();
  });

  it("keeps the mini toolbar inside the window", () => {
    vi.spyOn(HTMLElement.prototype, "getBoundingClientRect").mockReturnValue({ width: 300, height: 40, left: 0, top: 0, right: 300, bottom: 40, x: 0, y: 0, toJSON: () => ({}) });
    showContextMenu({ x: window.innerWidth - 50, y: 200 }, [{ label: "Cut" }], { mini: (el) => el.createEl("button", { text: "B" }) });
    const mini = document.querySelector<HTMLElement>(".dw-ctx-mini")!;
    expect(parseFloat(mini.style.left) + 300).toBeLessThanOrEqual(window.innerWidth);
  });

  it("opens in the document it's given (popout windows)", () => {
    const frame = document.body.createEl("iframe");
    const doc = frame.contentDocument!;
    showContextMenu({ x: 0, y: 0, doc }, [{ label: "Cut" }]);
    expect(doc.querySelector(".dw-ctx")).not.toBeNull();
    expect(document.querySelector(".dw-ctx")).toBeNull();
    closeContextMenu();
    expect(doc.querySelector(".dw-ctx")).toBeNull();
  });
});
