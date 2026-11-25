/*
 * Google Docs' Tools › Citations: the citation format, the document's sources (cite, edit, delete), adding a source
 * and inserting the bibliography. The view stores and applies the changes (docx-view `citations`).
 */
import { App, Modal, Notice, Setting } from "obsidian";
import { CITATION_STYLES, emptySource, newTag, parseAuthors, SOURCE_TYPES, type CitationStyle, type Source, type SourceType } from "../../domain/docx/citations";

export interface CitationsActions {
  style: CitationStyle;
  sources: Source[];
  /** The format or the sources changed: store them and update the document's citations. */
  save(style: CitationStyle, sources: Source[]): void;
  cite(source: Source): void;
  bibliography(): void;
}

/** The dialog; it closes after Cite or Insert bibliography. */
export function citationsDialog(app: App, actions: CitationsActions): Modal {
  let { style, sources } = actions;
  const modal = new Modal(app);
  modal.titleEl.setText("Citations");
  modal.modalEl.addClass("dw-citations");
  const render = () => {
    modal.contentEl.empty();
    new Setting(modal.contentEl).setName("Citation format").addDropdown((d) =>
      d
        .addOptions(Object.fromEntries(CITATION_STYLES))
        .setValue(style)
        .onChange((v) => {
          style = v as CitationStyle;
          actions.save(style, sources);
        })
    );
    const list = modal.contentEl.createDiv({ cls: "dw-citations-list" });
    if (!sources.length) list.createDiv({ cls: "dw-citations-empty", text: "No sources yet. Add one to cite it." });
    for (const source of sources) {
      const row = new Setting(list).setName(source.title || "(Untitled)").setDesc(describe(source));
      row.settingEl.addClass("dw-citations-source");
      row
        .addButton((b) =>
          b
            .setButtonText("Cite")
            .setCta()
            .onClick(() => {
              modal.close();
              actions.cite(source);
            })
        )
        .addButton((b) =>
          b.setButtonText("Edit").onClick(() =>
            void sourceForm(app, source).then((edited) => {
              if (!edited) return;
              sources = sources.map((s) => (s === source ? { ...edited, tag: source.tag } : s));
              actions.save(style, sources);
              render();
            })
          )
        )
        .addButton((b) =>
          b.setButtonText("Delete").onClick(() => {
            sources = sources.filter((s) => s !== source);
            actions.save(style, sources);
            render();
          })
        );
    }
    new Setting(modal.contentEl)
      .addButton((b) =>
        b.setButtonText("Add citation source").onClick(() =>
          void sourceForm(app, emptySource()).then((added) => {
            if (!added) return;
            sources = [...sources, { ...added, tag: newTag(added, sources.map((s) => s.tag)) }];
            actions.save(style, sources);
            render();
          })
        )
      )
      .addButton((b) => {
        b.setButtonText("Insert bibliography").onClick(() => {
          modal.close();
          actions.bibliography();
        });
        if (!sources.length) b.setDisabled(true);
      });
  };
  render();
  modal.open();
  return modal;
}

const describe = (s: Source) =>
  [SOURCE_TYPES.find(([t]) => t === s.type)?.[1], s.authors.map((a) => [a.first, a.last].filter(Boolean).join(" ")).join(", "), s.year].filter(Boolean).join(" · ");

/** Add or edit a source; null when cancelled or untitled. */
export function sourceForm(app: App, source: Source): Promise<Source | null> {
  return new Promise((resolve) => {
    let result: Source | null = null;
    const now: Source = { ...source, authors: [...source.authors] };
    let authors = source.authors.map((a) => [a.first, a.last].filter(Boolean).join(" ")).join("\n");
    const modal = new Modal(app);
    modal.titleEl.setText(source.title ? "Edit source" : "Add citation source");
    const field = (name: string, key: keyof Omit<Source, "authors" | "type" | "tag">, types: SourceType[], desc = "") => {
      const s = new Setting(modal.contentEl).setName(name).setDesc(desc).addText((t) => t.setValue(now[key]).onChange((v) => (now[key] = v.trim())));
      s.settingEl.dataset.types = types.join(" ");
      return s;
    };
    new Setting(modal.contentEl).setName("Source type").addDropdown((d) =>
      d
        .addOptions(Object.fromEntries(SOURCE_TYPES))
        .setValue(now.type)
        .onChange((v) => {
          now.type = v as SourceType;
          showFields();
        })
    );
    new Setting(modal.contentEl)
      .setName("Contributors")
      .setDesc("One person per line, first name first or last name, comma, first name.")
      .addTextArea((t) => t.setValue(authors).onChange((v) => (authors = v)));
    field("Title", "title", ["book", "website", "journal"]);
    field("Year", "year", ["book", "website", "journal"]);
    field("Publisher", "publisher", ["book"]);
    field("Journal", "container", ["journal"]);
    field("Website name", "container", ["website"]);
    field("Volume", "volume", ["journal"]);
    field("Issue", "issue", ["journal"]);
    field("Pages", "pages", ["journal"], "For example 45-67.");
    field("URL", "url", ["book", "website", "journal"]);
    const showFields = () => {
      for (const el of Array.from(modal.contentEl.querySelectorAll<HTMLElement>("[data-types]"))) el.toggle(el.dataset.types!.split(" ").includes(now.type));
    };
    showFields();
    new Setting(modal.contentEl)
      .addButton((b) => b.setButtonText("Cancel").onClick(() => modal.close()))
      .addButton((b) =>
        b
          .setButtonText(source.title ? "Save" : "Add source")
          .setCta()
          .onClick(() => {
            if (!now.title) return void new Notice("A source needs a title.");
            result = { ...now, authors: parseAuthors(authors) };
            // Fields of other source types aren't kept.
            if (now.type !== "book") result.publisher = "";
            if (now.type === "book") result.container = "";
            if (now.type !== "journal") result.volume = result.issue = result.pages = "";
            modal.close();
          })
      );
    modal.onClose = () => resolve(result);
    modal.open();
  });
}
