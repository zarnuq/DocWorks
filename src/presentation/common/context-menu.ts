/**
 * The Office-style interface's right-click menu, drawn in the page (with Obsidian's menu classes, so it looks like
 * Obsidian's own menus) because an OS menu (Obsidian's "native menus") can't sit beside the mini toolbar.
 */
import { setIcon } from "obsidian";

export interface CtxItem {
  label: string;
  run?: () => void;
  icon?: string;
  checked?: boolean;
  disabled?: boolean;
  sub?: CtxEntry[];
}
export type CtxEntry = CtxItem | "---";

let open: { close(): void } | null = null;

/** Separators only between items. */
function tidy(entries: CtxEntry[]): CtxEntry[] {
  const out: CtxEntry[] = [];
  for (const e of entries) if (e !== "---" || (out.length && out[out.length - 1] !== "---")) out.push(e);
  while (out[out.length - 1] === "---") out.pop();
  return out;
}

/** Put `el` at x / y, moved back inside its window when it would leave it. */
function place(el: HTMLElement, x: number, y: number): void {
  const win = el.ownerDocument.defaultView ?? window;
  el.style.left = `${x}px`;
  el.style.top = `${y}px`;
  const r = el.getBoundingClientRect();
  if (x + r.width > win.innerWidth) el.style.left = `${Math.max(0, win.innerWidth - r.width)}px`;
  if (y + r.height > win.innerHeight) el.style.top = `${Math.max(0, win.innerHeight - r.height)}px`;
}

/** Close the open menu, if any (its view is closing). */
export function closeContextMenu(): void {
  open?.close();
}

/**
 * Show a menu at `at` (with the mini toolbar above it when `mini` fills one); returns a handle to close it. `at.doc`:
 * the document clicked in (a popout window's), else the main one.
 */
export function showContextMenu(at: { x: number; y: number; doc?: Document }, entries: CtxEntry[], opts: { mini?: (el: HTMLElement) => void; onClose?: () => void } = {}): { close(): void } {
  open?.close();
  const doc = at.doc ?? document;
  const win = doc.defaultView ?? window;
  // Built here and moved into the clicked window's document (Obsidian's helpers work on it either way).
  const root = createDiv({ cls: "dw-ctx" });
  doc.body.appendChild(root);
  // The clicked window's own AbortController: its listeners take only its own signals.
  const abort = new (win as typeof globalThis).AbortController();
  let closed = false;
  const handle = {
    close: () => {
      if (closed) return;
      closed = true;
      abort.abort();
      root.remove();
      if (open === handle) open = null;
      opts.onClose?.();
    },
  };
  open = handle;
  const close = handle.close;

  // Clicks here keep the document's selection (the commands act on it).
  root.addEventListener("mousedown", (e) => {
    if (!(e.target as HTMLElement).closest("input, select, textarea")) e.preventDefault();
  });
  let mini: HTMLElement | null = null;
  if (opts.mini) {
    mini = root.createDiv({ cls: "dw-ctx-mini" });
    opts.mini(mini);
  }
  let sub: HTMLElement | null = null;
  /** Selects a menu's first enabled item (keyboard into a submenu). */
  const selectFirst = new WeakMap<HTMLElement, () => void>();
  const closeSub = () => {
    sub?.remove();
    sub = null;
  };

  const build = (list: CtxEntry[], depth: number): HTMLElement => {
    const menu = root.createDiv({ cls: depth ? "menu dw-ctx-sub" : "menu", attr: { tabindex: "-1", role: "menu" } });
    const rows: { el: HTMLElement; item: CtxItem }[] = [];
    for (const e of tidy(list)) {
      if (e === "---") {
        menu.createDiv({ cls: "menu-separator" });
        continue;
      }
      const el = menu.createDiv({ cls: "menu-item", attr: { role: "menuitem" } });
      if (e.disabled) el.addClass("is-disabled");
      const icon = el.createDiv({ cls: "menu-item-icon" });
      if (e.icon || e.checked) setIcon(icon, e.checked ? "check" : e.icon!);
      el.createDiv({ cls: "menu-item-title", text: e.label });
      if (e.sub) setIcon(el.createDiv({ cls: "menu-item-icon mod-submenu" }), "chevron-right");
      rows.push({ el, item: e });
      el.addEventListener("click", () => {
        if (e.disabled) return;
        if (e.sub) return openSub(el, e.sub);
        close();
        e.run?.();
      });
      el.addEventListener("mouseenter", () => {
        if (!depth) (e.sub && !e.disabled ? openSub(el, e.sub) : closeSub());
      });
    }
    const enabled = () => rows.filter((r) => !r.item.disabled);
    const select = (el: HTMLElement | undefined) => {
      for (const r of rows) r.el.toggleClass("selected", r.el === el);
    };
    const selected = () => rows.find((r) => r.el.hasClass("selected"));
    menu.addEventListener("keydown", (e) => {
      const list = enabled(), now = selected(), i = now ? list.indexOf(now) : -1;
      if (e.key === "ArrowDown" || e.key === "ArrowUp") select(list[Math.min(list.length - 1, Math.max(0, i < 0 ? 0 : i + (e.key === "ArrowDown" ? 1 : -1)))]?.el);
      else if (e.key === "Enter" || e.key === " ") now?.el.click();
      else if (e.key === "ArrowRight" && now?.item.sub && !depth) openSub(now.el, now.item.sub);
      else if (e.key === "ArrowLeft" && depth) {
        closeSub();
        root.querySelector<HTMLElement>(".menu:not(.dw-ctx-sub)")?.focus();
      } else if (e.key === "Escape") close();
      else return;
      e.preventDefault();
      e.stopPropagation();
    });
    selectFirst.set(menu, () => select(enabled()[0]?.el));
    return menu;
  };

  const openSub = (itemEl: HTMLElement, list: CtxEntry[]) => {
    closeSub();
    sub = build(list, 1);
    const r = itemEl.getBoundingClientRect();
    place(sub, r.right, r.top);
    selectFirst.get(sub)?.();
    sub.focus();
  };

  const menu = build(entries, 0);
  place(menu, at.x, at.y);
  if (mini) {
    const m = menu.getBoundingClientRect(), h = mini.getBoundingClientRect().height;
    const left = parseFloat(menu.style.left) || at.x, top = parseFloat(menu.style.top) || at.y;
    // Above the menu, as Word; below it when there's no room; inside the window either way.
    place(mini, left, top - h - 6 >= 0 ? top - h - 6 : top + m.height + 6);
  }
  doc.addEventListener("mousedown", (e) => !root.contains(e.target as Node) && close(), { capture: true, signal: abort.signal });
  doc.addEventListener("scroll", (e) => !root.contains(e.target as Node) && close(), { capture: true, signal: abort.signal });
  // Escape closes it wherever focus is (a mini toolbar button hands focus back to the document).
  doc.addEventListener("keydown", (e) => {
    if (e.key !== "Escape" || root.contains(e.target as Node)) return;
    e.preventDefault();
    close();
  }, { capture: true, signal: abort.signal });
  win.addEventListener("blur", close, { signal: abort.signal });
  menu.focus();
  return handle;
}
