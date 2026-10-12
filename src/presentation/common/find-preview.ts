/** A search result stays readable while navigating annotations without opening a modal for every match. */
export function showFindPreview(el: HTMLElement, label: string, text: string, start: number, end: number, action: string, open: () => void): void {
  el.empty();
  el.createSpan({ cls: "dw-find-source", text: label });
  const snippet = el.createSpan({ cls: "dw-find-snippet" });
  const from = Math.max(0, start - 80), to = Math.min(text.length, end + 80);
  snippet.appendText(`${from ? "…" : ""}${text.slice(from, start)}`);
  snippet.createEl("mark", { text: text.slice(start, end) });
  snippet.appendText(`${text.slice(end, to)}${to < text.length ? "…" : ""}`);
  el.createEl("button", { cls: "dw-btn", text: action }).addEventListener("click", open);
}
