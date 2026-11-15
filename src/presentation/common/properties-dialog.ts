/* File › Properties: a file's title, subject, author, keywords, description and category (docProps/core.xml). */
import { App, Modal, Setting } from "obsidian";
import type { OoxmlPackage } from "../../infrastructure/ooxml/ooxml-package";
import { EditableProp, EDITABLE_PROPS, readProps, writeProps } from "../../infrastructure/ooxml/doc-props";

const LABELS: Record<EditableProp, string> = { title: "Title", subject: "Subject", creator: "Author", keywords: "Keywords", description: "Description", category: "Category" };

const when = (iso: string) => {
  const d = new Date(iso);
  return Number.isNaN(d.getTime()) ? iso : d.toLocaleString();
};

/** Show (and unless read-only, edit) the properties; `changed` runs after edits were written to the package. */
export function propertiesDialog(app: App, pkg: OoxmlPackage, readOnly: boolean, changed: () => void): Modal {
  const modal = new Modal(app);
  modal.titleEl.setText("Properties");
  const props = readProps(pkg);
  const edits: Partial<Record<EditableProp, string>> = {};
  const el = modal.contentEl;
  for (const key of EDITABLE_PROPS) {
    const row = new Setting(el).setName(LABELS[key]);
    const control = (c: { setValue(v: string): unknown; onChange(fn: (v: string) => void): unknown; setDisabled?(d: boolean): unknown; inputEl: HTMLInputElement | HTMLTextAreaElement }) => {
      c.setValue(props[key]);
      c.onChange((v) => (edits[key] = v));
      c.inputEl.disabled = readOnly;
      c.inputEl.dataset.prop = key;
    };
    if (key === "description") row.addTextArea(control);
    else row.addText(control);
  }
  const facts: [string, string][] = [["Created", props.created && when(props.created)], ["Modified", props.modified && when(props.modified)], ["Last modified by", props.lastModifiedBy]];
  for (const [name, value] of facts) if (value) new Setting(el).setName(name).setDesc(value);
  const buttons = new Setting(el);
  if (!readOnly) {
    buttons.addButton((b) => b.setButtonText("Save").setCta().onClick(() => {
      if (writeProps(pkg, edits)) changed();
      modal.close();
    }));
  }
  buttons.addButton((b) => b.setButtonText(readOnly ? "Close" : "Cancel").onClick(() => modal.close()));
  modal.open();
  return modal;
}
