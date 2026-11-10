/** The Office-style interface's status bar: facts on the left, view buttons and the zoom slider on the right. */
import { Menu, setIcon } from "obsidian";
import type { AppStorage } from "./commands";
import { zoomLabel } from "./toolbar";

export interface StatusItem {
  id: string;
  /** Its name in the Customize Status Bar menu. */
  label: string;
  side: "left" | "right";
  render: (el: HTMLElement) => void;
}

export class StatusBar {
  readonly root: HTMLElement;
  /** Where the save status goes. */
  readonly save: HTMLElement;
  private readonly cells = new Map<string, HTMLElement>();
  private readonly right: HTMLElement;
  private zoomSync: (() => void) | null = null;

  constructor(
    el: HTMLElement,
    private readonly items: StatusItem[],
    private readonly storage: AppStorage,
    private readonly key: string
  ) {
    this.root = el.createDiv({ cls: "dw-statusbar", attr: { role: "status" } });
    const left = this.root.createDiv({ cls: "dw-sb-left" });
    this.right = this.root.createDiv({ cls: "dw-sb-right" });
    for (const item of items) this.cells.set(item.id, (item.side === "left" ? left : this.right).createDiv({ cls: "dw-sb-item", attr: { "data-item": item.id } }));
    this.save = left.createDiv({ cls: "dw-sb-save" });
    this.root.addEventListener("contextmenu", (e) => {
      e.preventDefault();
      const hidden = this.hidden();
      const m = new Menu();
      for (const item of items) m.addItem((i) => i.setTitle(item.label).setChecked(!hidden.includes(item.id)).onClick(() => this.toggle(item.id)));
      m.showAtMouseEvent(e);
    });
    this.update();
  }

  private hidden(): string[] {
    const saved = this.storage.loadLocalStorage(`${this.key}-status`);
    return Array.isArray(saved) ? saved.filter((x): x is string => typeof x === "string") : [];
  }

  private toggle(id: string): void {
    const hidden = this.hidden();
    this.storage.saveLocalStorage(`${this.key}-status`, hidden.includes(id) ? hidden.filter((x) => x !== id) : [...hidden, id]);
    this.update();
  }

  /** Draw every shown item again. */
  update(): void {
    const hidden = this.hidden();
    for (const item of this.items) {
      const cell = this.cells.get(item.id)!;
      cell.hidden = hidden.includes(item.id);
      cell.empty();
      if (!cell.hidden) item.render(cell);
    }
    this.zoomSync?.();
  }

  /** − / slider / + and the percentage (opening the zoom menu) at the right end, as Word's. */
  zoom(opts: { get: () => string; set: (z: string) => void; levels: string[] }): void {
    const steps = opts.levels.filter((z) => z !== "fit");
    const index = () => {
      const i = steps.indexOf(opts.get());
      return i >= 0 ? i : Math.max(0, steps.indexOf("1"));
    };
    const box = this.right.createDiv({ cls: "dw-sb-zoom" });
    const set = (z: string) => {
      opts.set(z);
      sync();
    };
    const out = box.createEl("button", { cls: "clickable-icon", attr: { "aria-label": "Zoom out" } });
    setIcon(out, "minus");
    out.addEventListener("click", () => set(steps[Math.max(0, index() - 1)]));
    const slider = box.createEl("input", { type: "range", attr: { min: "0", max: String(steps.length - 1), step: "1", "aria-label": "Zoom" } });
    slider.addEventListener("input", () => set(steps[Number(slider.value)]));
    const inBtn = box.createEl("button", { cls: "clickable-icon", attr: { "aria-label": "Zoom in" } });
    setIcon(inBtn, "plus");
    inBtn.addEventListener("click", () => set(steps[Math.min(steps.length - 1, index() + 1)]));
    const pct = box.createEl("button", { cls: "dw-sb-zoom-pct", attr: { "aria-label": "Zoom level", "aria-haspopup": "menu" } });
    pct.addEventListener("click", () => {
      const m = new Menu();
      for (const z of opts.levels) m.addItem((i) => i.setTitle(zoomLabel(z)).setChecked(opts.get() === z).onClick(() => set(z)));
      const r = pct.getBoundingClientRect();
      m.showAtPosition({ x: r.left, y: r.top });
    });
    const sync = (this.zoomSync = () => {
      slider.value = String(index());
      pct.setText(zoomLabel(opts.get()));
    });
    sync();
  }
}
