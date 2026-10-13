/*
 * Narrow panes, as Google's editors on a small window: below NARROW px the view gets `is-narrow`, docked side panels
 * are hidden by CSS, and their toggles open them as drawers over the document instead.
 */
export const NARROW = 700;

/** Toggle `is-narrow` on `root` as it resizes; `onChange` runs when it flips. Returns a stop function. */
export function watchNarrow(root: HTMLElement, onChange: (narrow: boolean) => void, width = () => root.clientWidth): (() => void) & { check: () => void } {
  // A file opened in a pane that was narrow keeps the class until the first measure says otherwise.
  let narrow = root.hasClass("is-narrow");
  const check = () => {
    const w = width();
    // 0 = not laid out yet (hidden tab): leave the layout as it is.
    const now = w > 0 ? w < NARROW : narrow;
    if (now === narrow) return;
    narrow = now;
    root.toggleClass("is-narrow", now);
    onChange(now);
  };
  const observer = typeof ResizeObserver === "undefined" ? null : new ResizeObserver(check);
  observer?.observe(root);
  return Object.assign(() => observer?.disconnect(), { check });
}

/** Show `panel` as a drawer over the document until Escape, a click outside it, or the returned function. */
export function openDrawer(panel: HTMLElement, opts: { root: HTMLElement; toggle?: HTMLElement | null; onClose: () => void }): () => void {
  let open = true;
  const outside = (e: Event) => {
    const target = e.target as Node;
    if (!panel.contains(target) && !opts.toggle?.contains(target)) close();
  };
  // Capture phase: runs before the panel's escapeToEditor, which then moves focus back to the document.
  const escape = (e: KeyboardEvent) => {
    if (e.key === "Escape") close();
  };
  const close = () => {
    if (!open) return;
    open = false;
    panel.removeClass("is-drawer");
    opts.root.removeEventListener("pointerdown", outside, true);
    panel.removeEventListener("keydown", escape, true);
    opts.onClose();
  };
  panel.addClass("is-drawer");
  panel.show();
  opts.root.addEventListener("pointerdown", outside, true);
  panel.addEventListener("keydown", escape, true);
  return close;
}
