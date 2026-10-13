/*
 * Keyboard access shared by the three editors: each toolbar is one Tab stop with arrow keys inside it (the ARIA
 * toolbar pattern Google Docs uses), F6 / Shift+F6 move between the view's regions, Escape in a side panel goes back
 * to the document, and focus returns to the document when a dialog or menu closes.
 */
import type { Scope } from "obsidian";

const FOCUSABLE = 'button, select, input, textarea, [tabindex]:not([tabindex="-1"]), [contenteditable="true"]';

/** Shown on screen: neither it nor an ancestor is display: none or hidden. */
function shown(el: HTMLElement): boolean {
  for (let n: HTMLElement | null = el; n; n = n.parentElement) {
    if (n.hidden || n.style.display === "none" || getComputedStyle(n).display === "none") return false;
  }
  return true;
}

const enabled = (el: HTMLElement) => !(el as HTMLButtonElement).disabled;

function toolbarItems(bar: HTMLElement): HTMLElement[] {
  return Array.from(bar.querySelectorAll<HTMLElement>("button, select, input")).filter((el) => enabled(el) && shown(el));
}

/** Make `bar` one Tab stop whose controls are reached with Left / Right / Home / End. */
export function rovingToolbar(bar: HTMLElement): void {
  bar.setAttribute("role", "toolbar");
  const setStop = (current: HTMLElement) => {
    for (const el of Array.from(bar.querySelectorAll<HTMLElement>("button, select, input"))) el.tabIndex = el === current ? 0 : -1;
  };
  const first = toolbarItems(bar)[0];
  if (first) setStop(first);
  bar.addEventListener("focusin", (e) => {
    const target = e.target as HTMLElement;
    if (bar.contains(target) && target.matches("button, select, input")) setStop(target);
  });
  bar.addEventListener("keydown", (e) => {
    const target = e.target as HTMLElement;
    // Text fields keep their own arrow keys.
    if (target.instanceOf(HTMLInputElement) && !["color", "checkbox", "button"].includes(target.type)) return;
    if (!["ArrowLeft", "ArrowRight", "Home", "End"].includes(e.key) || e.altKey || e.ctrlKey || e.metaKey) return;
    const items = toolbarItems(bar);
    const i = items.indexOf(target);
    if (i < 0) return;
    e.preventDefault();
    const next = e.key === "Home" ? 0 : e.key === "End" ? items.length - 1 : (i + (e.key === "ArrowRight" ? 1 : -1) + items.length) % items.length;
    setStop(items[next]);
    items[next].focus();
  });
}

/** Where focus goes in a region: a toolbar's current stop, the region itself if it takes focus, or its first control. */
function focusTarget(region: HTMLElement): HTMLElement | null {
  if (["toolbar", "menubar"].includes(region.getAttribute("role") ?? "")) return toolbarItems(region).find((el) => el.tabIndex === 0) ?? toolbarItems(region)[0] ?? null;
  if (region.hasAttribute("tabindex") || region.isContentEditable || region.matches("button, select, input, textarea")) return region;
  return Array.from(region.querySelectorAll<HTMLElement>(FOCUSABLE)).find((el) => enabled(el) && shown(el)) ?? null;
}

/** F6 / Shift+F6 move focus to the next / previous shown region of a view (while focus is inside `root`). */
export function focusRegions(scope: Scope, root: HTMLElement, regions: () => (HTMLElement | null | undefined)[]): void {
  const move = (dir: 1 | -1) => () => {
    const active = document.activeElement;
    if (!active || !root.contains(active)) return true;
    const list = regions().filter((r): r is HTMLElement => !!r && shown(r) && !!focusTarget(r));
    if (!list.length) return true;
    const at = list.findIndex((r) => r.contains(active));
    const next = list[(at + dir + list.length) % list.length];
    focusTarget(next)?.focus();
    return false;
  };
  scope.register([], "F6", move(1));
  scope.register(["Shift"], "F6", move(-1));
}

/** Escape inside `panel` (that nothing else used) returns focus to the document; the panel stays open. */
export function escapeToEditor(panel: HTMLElement, focusEditor: () => void): void {
  panel.addEventListener("keydown", (e) => {
    if (e.key !== "Escape" || e.defaultPrevented) return;
    e.preventDefault();
    focusEditor();
  });
}

/**
 * When a dialog or menu closes and focus is left on nothing, give it back to the document of the active view.
 * Returns a function that stops watching.
 */
export function returnFocusAfterOverlays(isActive: () => boolean, focusEditor: () => void): () => void {
  const observer = new MutationObserver((records) => {
    const closed = records.some((r) => Array.from(r.removedNodes).some((n) => n.instanceOf(HTMLElement) && (n.classList.contains("modal-container") || n.classList.contains("menu"))));
    if (!closed) return;
    // After the overlay's own focus handling.
    window.setTimeout(() => {
      const a = document.activeElement;
      if ((!a || a === document.body) && isActive()) focusEditor();
    }, 0);
  });
  observer.observe(document.body, { childList: true });
  return () => observer.disconnect();
}
