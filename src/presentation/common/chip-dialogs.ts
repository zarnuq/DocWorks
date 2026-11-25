/*
 * Dialogs for smart chips: a date chip's day and format, a dropdown chip's items (Google's "Dropdown" editor with
 * its two ready-made lists).
 */
import { App, Modal, Setting } from "obsidian";
import { DATE_FORMATS, formatDate } from "../../domain/docx/chips";

/** Google's ready-made dropdowns. */
export const DROPDOWN_PRESETS: [name: string, items: string[]][] = [
  ["Project status", ["Not started", "In progress", "Blocked", "Done"]],
  ["Review status", ["Not started", "In review", "Changes requested", "Approved"]],
];

function buttons(modal: Modal, ok: string, done: () => void): void {
  new Setting(modal.contentEl)
    .addButton((b) => b.setButtonText("Cancel").onClick(() => modal.close()))
    .addButton((b) =>
      b
        .setButtonText(ok)
        .setCta()
        .onClick(() => {
          done();
          modal.close();
        })
    );
}

/** Day and format of a date chip; null when cancelled. */
export function dateDialog(app: App, date: string, format: string, ok = "Insert"): Promise<{ date: string; format: string } | null> {
  return new Promise((resolve) => {
    let result: { date: string; format: string } | null = null;
    const now = { date, format };
    const modal = new Modal(app);
    modal.titleEl.setText("Date");
    let formats: HTMLSelectElement | null = null;
    /** The formats show the picked day. */
    const relabel = () => {
      for (const option of Array.from(formats?.options ?? [])) option.text = formatDate(now.date, option.value);
    };
    new Setting(modal.contentEl).setName("Date").addText((t) => {
      t.inputEl.type = "date";
      t.setValue(date).onChange((v) => {
        if (!/^\d{4}-\d\d-\d\d$/.test(v)) return;
        now.date = v;
        relabel();
      });
    });
    new Setting(modal.contentEl).setName("Format").addDropdown((d) => {
      const options = Object.fromEntries(DATE_FORMATS.map(([, f]) => [f, f]));
      if (!options[format]) options[format] = format;
      d.addOptions(options).setValue(format).onChange((v) => (now.format = v));
      formats = d.selectEl;
      relabel();
    });
    buttons(modal, ok, () => (result = { ...now }));
    modal.onClose = () => resolve(result);
    modal.open();
  });
}

/** A dropdown chip's items, one per line; null when cancelled. */
export function dropdownDialog(app: App, items: string[], ok = "Insert"): Promise<string[] | null> {
  return new Promise((resolve) => {
    let result: string[] | null = null;
    const modal = new Modal(app);
    modal.titleEl.setText("Dropdown");
    let area: HTMLTextAreaElement | null = null;
    new Setting(modal.contentEl).setName("Start from").addDropdown((d) =>
      d
        .addOptions({ "": "Your own list", ...Object.fromEntries(DROPDOWN_PRESETS.map(([name]) => [name, name])) })
        .setValue("")
        .onChange((v) => {
          const preset = DROPDOWN_PRESETS.find(([name]) => name === v);
          if (preset && area) area.value = preset[1].join("\n");
        })
    );
    new Setting(modal.contentEl).setName("Options").setDesc("One per line.").addTextArea((t) => {
      t.setValue(items.join("\n"));
      area = t.inputEl;
      area.rows = 6;
    });
    buttons(modal, ok, () => (result = (area?.value ?? "").split("\n")));
    modal.onClose = () => resolve(result);
    modal.open();
  });
}
