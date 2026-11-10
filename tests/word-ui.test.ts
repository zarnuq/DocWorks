import { describe, expect, it } from "vitest";
import { controlIds, type Tab } from "../src/presentation/common/ribbon";
import { WORD_COMMAND_IDS, WORD_FILE_MENU, WORD_NESTED, WORD_QAT, WORD_SHORTCUTS, wordShortcutSections, wordTabs, type WordControls } from "../src/presentation/docx/word-ui";

const controls: WordControls = {
  fonts: ["Calibri", "Arial"],
  sizes: [10, 11, 12],
  styles: [["Normal", "Normal"], ["Heading 1", "Heading1"]],
  font: { value: () => "Calibri", set: () => {} },
  size: { value: () => "11", set: () => {} },
  style: { value: () => "Normal", set: () => {} },
  highlight: () => {},
  color: () => {},
  shading: () => {},
  inTable: () => false,
  onPicture: () => false,
  onDrawing: () => false,
};

const tabIds = (tabs: Tab[]) => tabs.flatMap((t) => t.groups.flatMap((g) => [...g.controls.flatMap(controlIds), ...(g.launcher ? [g.launcher] : [])]));

describe("Word's ribbon layout", () => {
  const tabs = wordTabs(controls);

  it("has Word's tabs in Word's order, contextual ones after", () => {
    expect(tabs.filter((t) => !t.context).map((t) => t.label)).toEqual(["Home", "Insert", "Design", "Layout", "References", "Review", "View", "Help"]);
    expect(tabs.filter((t) => t.context).map((t) => t.label)).toEqual(["Table Design", "Layout", "Picture Format", "Shape Format"]);
  });

  it("has Word's groups, in order, on each tab", () => {
    const groups = (label: string) => tabs.find((t) => t.label === label && !t.context)!.groups.map((g) => g.label);
    expect(groups("Home")).toEqual(["Clipboard", "Font", "Paragraph", "Styles", "Editing"]);
    expect(groups("Insert")).toEqual(["Pages", "Tables", "Illustrations", "Links", "Comments", "Header & Footer", "Text", "Symbols"]);
    expect(groups("Design")).toEqual(["Page Background"]);
    expect(groups("Layout")).toEqual(["Page Setup", "Paragraph", "Arrange"]);
    expect(groups("References")).toEqual(["Table of Contents", "Footnotes", "Notes"]);
    expect(groups("Review")).toEqual(["Proofing", "Comments", "Tracking", "Changes", "Compare"]);
    expect(groups("View")).toEqual(["Views", "Dark Mode", "Show", "Zoom"]);
    expect(groups("Help")).toEqual(["Help"]);
  });

  it("shows the table tabs in a table, Picture Format on a picture, Shape Format on a drawing", () => {
    const state = { table: false, picture: false, drawing: false };
    const live = wordTabs({ ...controls, inTable: () => state.table, onPicture: () => state.picture, onDrawing: () => state.drawing });
    const shown = () => live.filter((t) => t.context?.()).map((t) => t.label);
    expect(shown()).toEqual([]);
    state.table = true;
    expect(shown()).toEqual(["Table Design", "Layout"]);
    state.table = false;
    state.picture = true;
    expect(shown()).toEqual(["Picture Format"]);
    state.picture = false;
    state.drawing = true;
    expect(shown()).toEqual(["Shape Format"]);
  });

  it("names only commands the view provides", () => {
    const known = new Set<string>(WORD_COMMAND_IDS);
    const used = [...tabIds(tabs), ...WORD_QAT.defaults, ...WORD_QAT.choices, ...WORD_SHORTCUTS.map(([, id]) => id), ...Object.keys(WORD_NESTED), ...Object.values(WORD_NESTED).flat(), ...WORD_FILE_MENU];
    expect(used.filter((id) => !known.has(id))).toEqual([]);
  });

  it("reaches every command the Google Docs layout has", () => {
    const reachable = new Set([...tabIds(tabs), ...WORD_QAT.choices, ...WORD_FILE_MENU, ...Object.values(WORD_NESTED).flat()]);
    // Toolbar, Edit / View / Insert / Format / Tools / Help menus, side buttons, File and the table menu of the Google layout.
    const GOOGLE_PARITY = [
      "Undo", "Redo", "ZoomDialog", "ViewPrintLayoutView", "ViewWebLayoutView", "QuickStylesGallery", "UpdateStyle", "Font", "FontSize",
      "Bold", "Italic", "Underline", "FontColorPicker", "TextHighlightColorPicker", "HyperlinkInsert", "ReviewNewComment", "PictureInsertFromFile",
      "AlignLeft", "AlignCenter", "AlignRight", "AlignJustify", "LineSpacingGallery", "BulletsGalleryWord", "NumberingGalleryWord",
      "IndentDecreaseWord", "IndentIncreaseWord", "FormatPainter", "ClearFormatting", "TextWrapGallery", "PictureCrop", "AltText", "PictureChange",
      "TableShading", "TableBorders", "ReviewTrackChanges", "Cut", "Copy", "Paste", "PasteTextOnly", "SelectAll", "CopyLinkToHeading",
      "CreateNoteFromSelection", "Delete", "FindDialog", "ReplaceDialog", "ViewRulerWord", "NavigationPaneShowHide", "ShowComments", "LinkedFrom",
      "DarkModeOn", "TableInsertGallery", "ShapesInsertGallery", "ChartInsert", "EquationInsert", "SymbolInsert", "HorizontalLineInsert",
      "BreaksGallery", "PageBreakInsertWord", "ColumnBreakInsert", "SectionBreakNextPage", "SectionBreakContinuous", "FootnoteInsert",
      "TableOfContentsInsert", "TableOfContentsUpdate", "HeaderInsertGallery", "WatermarkGallery", "Strikethrough", "Subscript", "Superscript",
      "InlineCode", "ParagraphDialog", "TableColumnsGallery", "WordCount", "KeyboardShortcuts", "VersionHistory", "CompanionNote", "ExportFile",
      "TableRowsInsertAboveWord", "TableRowsInsertBelowWord", "TableColumnsInsertLeft", "TableColumnsInsertRight", "MergeCellsRight", "MergeCellsDown",
      "SplitCells", "TableColumnWidth", "TableRepeatHeaderRows", "TableAllowRowBreak", "TableDeleteRows", "TableDeleteColumns", "TableDelete",
      "ReviewDeleteComment", "ReviewAcceptAll", "ReviewRejectAll", "EditShape",
    ];
    expect(GOOGLE_PARITY.filter((id) => !reachable.has(id))).toEqual([]);
  });

  it("starts the Quick Access Toolbar with Save, Undo, Redo", () => {
    expect(WORD_QAT.defaults).toEqual(["FileSave", "Undo", "Redo"]);
    expect(WORD_QAT.choices.slice(0, 3)).toEqual(["FileSave", "Undo", "Redo"]);
  });
});

describe("Word's shortcuts", () => {
  it("are unique", () => {
    const keys = WORD_SHORTCUTS.map(([k]) => k);
    expect(new Set(keys).size).toBe(keys.length);
  });

  it("include Word's alignment, font size, list and heading keys", () => {
    const map = new Map(WORD_SHORTCUTS);
    expect(["Mod-l", "Mod-e", "Mod-r", "Mod-j"].map((k) => map.get(k))).toEqual(["AlignLeft", "AlignCenter", "AlignRight", "AlignJustify"]);
    expect(map.get("Mod-]")).toBe("FontSizeIncreaseWord");
    expect(map.get("Mod-[")).toBe("FontSizeDecreaseWord");
    expect(map.get("Mod-Shift-l")).toBe("BulletsGalleryWord");
    expect(map.get("Mod-Alt-1")).toBe("StyleHeading1");
    expect(map.get("Mod-Shift-n")).toBe("StyleNormal");
    expect(map.get("Mod-Shift-e")).toBe("ReviewTrackChanges");
    expect(map.has("Mod-Shift-8")).toBe(false);
  });

  it("are listed for Help › Keyboard Shortcuts with Ctrl+ names", () => {
    const rows = wordShortcutSections((id) => id).flatMap(([, list]) => list);
    expect(rows).toContainEqual(["AlignCenter", "Ctrl+E"]);
    expect(rows).toContainEqual(["Superscript", "Ctrl+Shift+="]);
    expect(rows).toContainEqual(["ClearFormatting", "Ctrl+Space"]);
    expect(rows).toContainEqual(["EquationInsert", "Alt+="]);
  });
});
