import { App, Modal, setIcon, Setting } from "obsidian";

export const FONTS = ["Aptos", "Calibri", "Cambria", "Arial", "Times New Roman", "Georgia", "Verdana", "Segoe UI", "Trebuchet MS", "Courier New"];

export type PaperTheme = "light" | "dark";

/** Thin helpers for the editor toolbars; all controls are plain buttons, selects and color inputs, boxed in groups split by gap(). */
export class Toolbar {
  private group: HTMLElement;

  constructor(readonly el: HTMLElement) {
    this.group = el.createDiv({ cls: "dw-group" });
  }

  button(label: string, title: string, onClick: () => void, cls = ""): HTMLButtonElement {
    const b = this.group.createEl("button", { text: label, cls: `dw-btn ${cls}`.trim(), attr: { "aria-label": title } });
    b.addEventListener("click", (e) => {
      e.preventDefault();
      onClick();
    });
    return b;
  }

  /** A button showing a Lucide icon instead of a label. */
  icon(icon: string, title: string, onClick: () => void): HTMLButtonElement {
    const b = this.button("", title, onClick, "dw-icon");
    setIcon(b, icon);
    return b;
  }

  /** A select whose first option is a placeholder; `onChange` only fires for real values. */
  select(title: string, placeholder: string, options: [label: string, value: string][], onChange: (value: string) => void): HTMLSelectElement {
    return fillSelect(this.group.createEl("select", { cls: "dropdown dw-select", attr: { "aria-label": title } }), placeholder, options, onChange);
  }

  /** A select drawn as an icon button; clicking it opens the native menu. */
  iconSelect(icon: string, title: string, options: [label: string, value: string][], onChange: (value: string) => void): HTMLSelectElement {
    const wrap = this.group.createEl("label", { cls: "dw-btn dw-icon-select", attr: { "aria-label": title } });
    const glyph = wrap.createSpan();
    setIcon(glyph, icon);
    if (!glyph.querySelector("svg")) glyph.setText(title);
    wrap.createSpan({ cls: "dw-caret", text: "▾" });
    return fillSelect(wrap.createEl("select", { attr: { "aria-label": title } }), title, options, onChange);
  }

  /** A button-like label over a hidden color input; the last color picked shows as an underline. */
  color(label: string, title: string, initial: string, onChange: (hex: string) => void): void {
    const wrap = this.group.createEl("label", { text: label, cls: "dw-btn dw-color", attr: { "aria-label": title } });
    wrap.setCssProps({ "--dw-swatch": initial });
    const input = wrap.createEl("input", { type: "color", value: initial });
    input.addEventListener("change", () => {
      wrap.setCssProps({ "--dw-swatch": input.value });
      onChange(input.value.toUpperCase());
    });
  }

  /** Start the next group. */
  gap(): void {
    this.group = this.el.createDiv({ cls: "dw-group" });
  }

  /** Light/dark switch plus save status, pushed to the end of the bar. Returns the theme syncer and status element. */
  end(theme: PaperTheme, onTheme: (t: PaperTheme) => void, what: string): { syncTheme: (t: PaperTheme) => void; status: HTMLElement } {
    const end = this.el.createDiv({ cls: "dw-end" });
    const seg = end.createDiv({ cls: "dw-segmented" });
    const btns = (["light", "dark"] as const).map((t) => {
      const b = seg.createEl("button", { text: t === "light" ? "Light" : "Dark", attr: { "aria-label": `${t === "light" ? "Light" : "Dark"} ${what}` } });
      b.addEventListener("click", () => onTheme(t));
      return [t, b] as const;
    });
    const syncTheme = (current: PaperTheme) => btns.forEach(([t, b]) => b.toggleClass("is-active", t === current));
    syncTheme(theme);
    return { syncTheme, status: end.createSpan({ cls: "dw-status", text: "Saved" }) };
  }
}

/** Show a select's current value: tick the matching option, or name a value that isn't listed in its placeholder. */
export function showValue(s: HTMLSelectElement, value: string | null): void {
  const listed = value !== null && Array.from(s.options).some((o) => o.value === value && o.value !== "");
  s.options[0].text = listed || value === null ? (s.dataset.placeholder ?? "") : value;
  s.value = listed ? value : "";
}

function fillSelect(s: HTMLSelectElement, placeholder: string, options: [label: string, value: string][], onChange: (value: string) => void): HTMLSelectElement {
  s.dataset.placeholder = placeholder;
  // The placeholder labels the closed control only; macOS would otherwise list it with a check mark and anchor the menu on it.
  s.createEl("option", { text: placeholder, value: "" }).hidden = true;
  for (const [text, value] of options) s.createEl("option", { text, value });
  s.addEventListener("change", () => s.value && onChange(s.value));
  return s;
}

/** Ask for one line of text; resolves null when cancelled. */
export function promptText(app: App, title: string, value: string): Promise<string | null> {
  return new Promise((resolve) => {
    let result: string | null = null;
    const modal = new Modal(app);
    modal.titleEl.setText(title);
    const submit = () => {
      result = value;
      modal.close();
    };
    new Setting(modal.contentEl).addText((text) => {
      text.setValue(value).onChange((v) => (value = v));
      // Cancel the key too: focus returns to the editor before the key's input, which would add a line break there.
      text.inputEl.addEventListener("keydown", (e) => {
        if (e.key !== "Enter" || e.isComposing) return;
        e.preventDefault();
        submit();
      });
      window.setTimeout(() => text.inputEl.select());
    });
    new Setting(modal.contentEl)
      .addButton((b) => b.setButtonText("Save").setCta().onClick(submit))
      .addButton((b) => b.setButtonText("Cancel").onClick(() => modal.close()));
    modal.onClose = () => resolve(result);
    modal.open();
  });
}
