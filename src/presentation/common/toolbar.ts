import { App, Modal, Setting } from "obsidian";

export const FONTS = ["Aptos", "Calibri", "Cambria", "Arial", "Times New Roman", "Georgia", "Verdana", "Segoe UI", "Trebuchet MS", "Courier New"];

export type PaperTheme = "light" | "dark";

/** Thin helpers for the editor toolbars; all controls are plain buttons, selects and color inputs. */
export class Toolbar {
  constructor(readonly el: HTMLElement) {}

  button(label: string, title: string, onClick: () => void, cls = ""): HTMLButtonElement {
    const b = this.el.createEl("button", { text: label, cls: `dw-btn ${cls}`.trim(), attr: { "aria-label": title } });
    b.addEventListener("click", (e) => {
      e.preventDefault();
      onClick();
    });
    return b;
  }

  /** A select whose first option is a placeholder; `onChange` only fires for real values. */
  select(title: string, placeholder: string, options: [label: string, value: string][], onChange: (value: string) => void): HTMLSelectElement {
    const s = this.el.createEl("select", { cls: "dropdown dw-select", attr: { "aria-label": title } });
    for (const [text, value] of [[placeholder, ""], ...options]) s.createEl("option", { text, value });
    s.addEventListener("change", () => s.value && onChange(s.value));
    return s;
  }

  /** A button-like label over a hidden color input; the last color picked shows as an underline. */
  color(label: string, title: string, initial: string, onChange: (hex: string) => void): void {
    const wrap = this.el.createEl("label", { text: label, cls: "dw-btn dw-color", attr: { "aria-label": title } });
    wrap.setCssProps({ "--dw-swatch": initial });
    const input = wrap.createEl("input", { type: "color", value: initial });
    input.addEventListener("change", () => {
      wrap.setCssProps({ "--dw-swatch": input.value });
      onChange(input.value.toUpperCase());
    });
  }

  gap(): void {
    this.el.createDiv({ cls: "dw-gap" });
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
      text.inputEl.addEventListener("keydown", (e) => e.key === "Enter" && submit());
      window.setTimeout(() => text.inputEl.select());
    });
    new Setting(modal.contentEl)
      .addButton((b) => b.setButtonText("Save").setCta().onClick(submit))
      .addButton((b) => b.setButtonText("Cancel").onClick(() => modal.close()));
    modal.onClose = () => resolve(result);
    modal.open();
  });
}
