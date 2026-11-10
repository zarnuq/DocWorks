/**
 * Word's interface (beta), as data: desktop Word's ribbon tabs and groups in Word's order (Microsoft's command ids
 * where Word has one, see docs/office-ribbon/word.md), the Quick Access Toolbar, and Word's shortcuts. The view
 * supplies a command for every id in WORD_COMMAND_IDS.
 */
import type { Control, Preview, Tab } from "../common/ribbon";
import type { ShortcutSection } from "../common/shortcuts";

/** Value controls the view provides (the font box, colours…) and the tests of the contextual tabs. */
export interface WordControls {
  fonts: string[];
  sizes: number[];
  styles: [label: string, id: string][];
  font: { value: () => string | null; set: (v: string) => void };
  size: { value: () => string | null; set: (v: string) => void };
  style: { value: () => string | null; set: (v: string) => void; preview?: (id: string) => Preview };
  highlight: (hex: string) => void;
  color: (hex: string) => void;
  shading: (hex: string) => void;
  inTable: () => boolean;
  onPicture: () => boolean;
  onDrawing: () => boolean;
}

/** Every command id this file names. */
export const WORD_COMMAND_IDS = [
  "FileSave", "Undo", "Redo", "FilePrint", "ExportFile",
  "Paste", "PasteTextOnly", "Cut", "Copy", "FormatPainter",
  "Font", "FontSize", "FontSizeIncreaseWord", "FontSizeDecreaseWord", "ClearFormatting", "Bold", "Italic", "Underline", "Strikethrough",
  "Subscript", "Superscript", "InlineCode", "TextHighlightColorPicker", "FontColorPicker",
  "BulletsGalleryWord", "NumberingGalleryWord", "IndentDecreaseWord", "IndentIncreaseWord", "AlignLeft", "AlignCenter", "AlignRight",
  "AlignJustify", "LineSpacingGallery", "ParagraphDialog", "QuickStylesGallery", "UpdateStyle", "FindDialog", "ReplaceDialog", "SelectMenu",
  "SelectAll", "Delete",
  "PageBreakInsertWord", "TableInsertGallery", "PictureInsertFromFile", "ShapesInsertGallery", "ChartInsert", "HyperlinkInsert",
  "ReviewNewComment", "HeaderInsertGallery", "FooterInsertGallery", "TextBoxInsert", "HorizontalLineInsert", "EquationInsert", "SymbolInsert",
  "WatermarkGallery",
  "TableColumnsGallery", "BreaksGallery", "ColumnBreakInsert", "SectionBreakNextPage", "SectionBreakContinuous", "TextWrapGallery",
  "TableOfContentsInsert", "TableOfContentsUpdate", "FootnoteInsert", "CreateNoteFromSelection", "CopyLinkToHeading", "CompanionNote", "LinkedFrom",
  "WordCount", "ReviewDeleteComment", "ShowComments", "ReviewTrackChanges", "ReviewingPane", "ReviewAcceptAll", "ReviewRejectAll",
  "ReviewCompareMenu", "VersionHistory",
  "ViewPrintLayoutView", "ViewWebLayoutView", "DarkModeOn", "ViewRulerWord", "NavigationPaneShowHide", "ZoomDialog", "ZoomCurrent100",
  "ZoomPageWidth", "KeyboardShortcuts",
  "TableShading", "TableBorders", "TableDeleteMenu", "TableDeleteRows", "TableDeleteColumns", "TableDelete", "TableRowsInsertAboveWord",
  "TableRowsInsertBelowWord", "TableColumnsInsertLeft", "TableColumnsInsertRight", "MergeCellsRight", "MergeCellsDown", "SplitCells",
  "TableColumnWidth", "TableAllowRowBreak", "TableRepeatHeaderRows",
  "PictureChange", "AltText", "PictureCrop", "EditShape",
  // Keyboard only.
  "CopyFormat", "PasteFormat", "LineSpacingSingle", "LineSpacingDouble", "LineSpacingOneHalf", "StyleHeading1", "StyleHeading2",
  "StyleHeading3", "StyleNormal",
] as const;

export type WordCommandId = (typeof WORD_COMMAND_IDS)[number];

/** Word's File drop-down, before Obsidian's own file items (Rename, Delete…). */
export const WORD_FILE_MENU: WordCommandId[] = ["FileSave", "ExportFile", "FilePrint", "VersionHistory"];

/** Commands reached through another command's menu (parent → children); the view's menus must offer them. */
export const WORD_NESTED: Partial<Record<WordCommandId, WordCommandId[]>> = {
  Paste: ["PasteTextOnly"],
  SelectMenu: ["SelectAll", "Delete"],
  QuickStylesGallery: ["UpdateStyle"],
  BreaksGallery: ["PageBreakInsertWord", "ColumnBreakInsert", "SectionBreakNextPage", "SectionBreakContinuous"],
  TableDeleteMenu: ["TableDeleteRows", "TableDeleteColumns", "TableDelete"],
  ReviewCompareMenu: ["VersionHistory"],
};

const btn = (kind: "large" | "small" | "icon", id: WordCommandId): Control => ({ kind, id });
const col = (...items: Control[]): Control => ({ kind: "column", items });
const row = (...items: Control[]): Control => ({ kind: "row", items });

export function wordTabs(c: WordControls): Tab[] {
  const [L, S, I] = [(id: WordCommandId) => btn("large", id), (id: WordCommandId) => btn("small", id), (id: WordCommandId) => btn("icon", id)];
  return [
    { id: "home", label: "Home", groups: [
      { id: "clipboard", label: "Clipboard", icon: "clipboard", controls: [L("Paste"), col(S("Cut"), S("Copy"), S("FormatPainter"))] },
      { id: "font", label: "Font", icon: "type", controls: [col(
        row(
          { kind: "select", id: "Font", width: 9, options: () => c.fonts.map((f) => [f, f]), value: c.font.value, set: c.font.set },
          { kind: "select", id: "FontSize", width: 4, options: () => c.sizes.map((s) => [String(s), String(s)]), value: c.size.value, set: c.size.set },
          I("FontSizeIncreaseWord"), I("FontSizeDecreaseWord"), I("ClearFormatting"),
        ),
        row(
          I("Bold"), I("Italic"), I("Underline"), I("Strikethrough"), I("Subscript"), I("Superscript"), I("InlineCode"),
          { kind: "color", id: "TextHighlightColorPicker", initial: "#FFFF00", set: c.highlight },
          { kind: "color", id: "FontColorPicker", initial: "#FF0000", set: c.color },
        ),
      )] },
      { id: "paragraph", label: "Paragraph", icon: "pilcrow", launcher: "ParagraphDialog", controls: [col(
        row(I("BulletsGalleryWord"), I("NumberingGalleryWord"), I("IndentDecreaseWord"), I("IndentIncreaseWord")),
        row(I("AlignLeft"), I("AlignCenter"), I("AlignRight"), I("AlignJustify"), I("LineSpacingGallery")),
      )] },
      { id: "styles", label: "Styles", icon: "heading", controls: [{ kind: "gallery", id: "QuickStylesGallery", items: () => c.styles, value: c.style.value, set: c.style.set, preview: c.style.preview }] },
      { id: "editing", label: "Editing", icon: "search", controls: [col(S("FindDialog"), S("ReplaceDialog"), S("SelectMenu"))] },
    ] },
    { id: "insert", label: "Insert", groups: [
      { id: "pages", label: "Pages", icon: "file-plus", controls: [L("PageBreakInsertWord")] },
      { id: "tables", label: "Tables", icon: "table", controls: [L("TableInsertGallery")] },
      { id: "illustrations", label: "Illustrations", icon: "image", controls: [L("PictureInsertFromFile"), L("ShapesInsertGallery"), L("ChartInsert")] },
      { id: "links", label: "Links", icon: "link", controls: [L("HyperlinkInsert")] },
      { id: "comments", label: "Comments", icon: "message-square-plus", controls: [L("ReviewNewComment")] },
      { id: "headerFooter", label: "Header & Footer", icon: "panel-top", controls: [col(S("HeaderInsertGallery"), S("FooterInsertGallery"))] },
      { id: "text", label: "Text", icon: "type", controls: [L("TextBoxInsert"), col(S("HorizontalLineInsert"))] },
      { id: "symbols", label: "Symbols", icon: "omega", controls: [L("EquationInsert"), L("SymbolInsert")] },
    ] },
    { id: "design", label: "Design", groups: [
      { id: "pageBackground", label: "Page Background", icon: "droplet", controls: [L("WatermarkGallery")] },
    ] },
    { id: "layout", label: "Layout", groups: [
      { id: "pageSetup", label: "Page Setup", icon: "file-text", controls: [L("TableColumnsGallery"), col(S("BreaksGallery"))] },
      { id: "paragraph", label: "Paragraph", icon: "pilcrow", launcher: "ParagraphDialog", controls: [L("ParagraphDialog")] },
      { id: "arrange", label: "Arrange", icon: "layers", controls: [L("TextWrapGallery")] },
    ] },
    { id: "references", label: "References", groups: [
      { id: "toc", label: "Table of Contents", icon: "book-open-text", controls: [L("TableOfContentsInsert"), col(S("TableOfContentsUpdate"))] },
      { id: "footnotes", label: "Footnotes", icon: "superscript", controls: [L("FootnoteInsert")] },
      { id: "notes", label: "Notes", icon: "notebook-pen", controls: [L("CreateNoteFromSelection"), col(S("CopyLinkToHeading"), S("CompanionNote"), S("LinkedFrom"))] },
    ] },
    { id: "review", label: "Review", groups: [
      { id: "proofing", label: "Proofing", icon: "spell-check", controls: [L("WordCount")] },
      { id: "comments", label: "Comments", icon: "messages-square", controls: [L("ReviewNewComment"), L("ReviewDeleteComment"), L("ShowComments")] },
      { id: "tracking", label: "Tracking", icon: "file-pen", controls: [L("ReviewTrackChanges"), L("ReviewingPane")] },
      { id: "changes", label: "Changes", icon: "check-check", controls: [L("ReviewAcceptAll"), L("ReviewRejectAll")] },
      { id: "compare", label: "Compare", icon: "git-compare", controls: [L("ReviewCompareMenu")] },
    ] },
    { id: "view", label: "View", groups: [
      { id: "views", label: "Views", icon: "book-open", controls: [L("ViewPrintLayoutView"), L("ViewWebLayoutView")] },
      { id: "darkMode", label: "Dark Mode", icon: "moon", controls: [L("DarkModeOn")] },
      { id: "show", label: "Show", icon: "eye", controls: [col({ kind: "check", id: "ViewRulerWord" }, { kind: "check", id: "NavigationPaneShowHide" })] },
      { id: "zoom", label: "Zoom", icon: "zoom-in", controls: [L("ZoomDialog"), L("ZoomCurrent100"), col(S("ZoomPageWidth"))] },
    ] },
    { id: "help", label: "Help", groups: [
      { id: "help", label: "Help", icon: "circle-help", controls: [L("KeyboardShortcuts")] },
    ] },
    // Contextual tabs.
    { id: "tableDesign", label: "Table Design", context: c.inTable, groups: [
      { id: "tableStyles", label: "Table Styles", icon: "paint-bucket", controls: [{ kind: "color", id: "TableShading", initial: "#EBF3FB", set: c.shading }] },
      { id: "borders", label: "Borders", icon: "square-dashed", controls: [L("TableBorders")] },
    ] },
    { id: "tableLayout", label: "Layout", context: c.inTable, groups: [
      { id: "rowsColumns", label: "Rows & Columns", icon: "table-2", controls: [L("TableDeleteMenu"), L("TableRowsInsertAboveWord"), col(S("TableRowsInsertBelowWord"), S("TableColumnsInsertLeft"), S("TableColumnsInsertRight"))] },
      { id: "merge", label: "Merge", icon: "table-cells-merge", controls: [col(S("MergeCellsRight"), S("MergeCellsDown"), S("SplitCells"))] },
      { id: "cellSize", label: "Cell Size", icon: "move-horizontal", controls: [col(S("TableColumnWidth"), S("TableAllowRowBreak"))] },
      { id: "data", label: "Data", icon: "rows-3", controls: [S("TableRepeatHeaderRows")] },
    ] },
    { id: "pictureFormat", label: "Picture Format", context: c.onPicture, groups: [
      { id: "adjust", label: "Adjust", icon: "image", controls: [L("PictureChange")] },
      { id: "accessibility", label: "Accessibility", icon: "text", controls: [L("AltText")] },
      { id: "arrange", label: "Arrange", icon: "layers", controls: [L("TextWrapGallery")] },
      { id: "size", label: "Size", icon: "crop", controls: [L("PictureCrop")] },
    ] },
    { id: "shapeFormat", label: "Shape Format", context: c.onDrawing, groups: [
      { id: "insertShapes", label: "Insert Shapes", icon: "shapes", controls: [L("EditShape")] },
    ] },
  ];
}

/** Quick Access Toolbar: what it starts with, and what its ▾ menu offers first (Word's list, limited to what Doc Works has). */
export const WORD_QAT: { defaults: WordCommandId[]; choices: WordCommandId[] } = {
  defaults: ["FileSave", "Undo", "Redo"],
  choices: ["FileSave", "Undo", "Redo", "FilePrint", "FindDialog", "WordCount", "ReviewTrackChanges"],
};

/** Word's shortcuts (ProseMirror key names) for the commands Doc Works has; active while the Word interface is on. */
export const WORD_SHORTCUTS: [key: string, id: WordCommandId][] = [
  ["Mod-b", "Bold"], ["Mod-i", "Italic"], ["Mod-u", "Underline"],
  ["Mod-=", "Subscript"], ["Mod-Shift-=", "Superscript"],
  ["Mod-]", "FontSizeIncreaseWord"], ["Mod-[", "FontSizeDecreaseWord"], ["Mod-Shift-.", "FontSizeIncreaseWord"], ["Mod-Shift-,", "FontSizeDecreaseWord"],
  ["Mod-Space", "ClearFormatting"],
  ["Mod-Shift-c", "CopyFormat"], ["Mod-Shift-v", "PasteFormat"],
  ["Mod-l", "AlignLeft"], ["Mod-e", "AlignCenter"], ["Mod-r", "AlignRight"], ["Mod-j", "AlignJustify"],
  ["Mod-1", "LineSpacingSingle"], ["Mod-2", "LineSpacingDouble"], ["Mod-5", "LineSpacingOneHalf"],
  ["Mod-m", "IndentIncreaseWord"], ["Mod-Shift-m", "IndentDecreaseWord"],
  ["Mod-Shift-l", "BulletsGalleryWord"],
  ["Mod-Alt-1", "StyleHeading1"], ["Mod-Alt-2", "StyleHeading2"], ["Mod-Alt-3", "StyleHeading3"], ["Mod-Shift-n", "StyleNormal"],
  ["Mod-Enter", "PageBreakInsertWord"], ["Mod-Shift-Enter", "ColumnBreakInsert"],
  ["Mod-k", "HyperlinkInsert"], ["Mod-Alt-m", "ReviewNewComment"], ["Mod-Alt-f", "FootnoteInsert"],
  ["Mod-Shift-e", "ReviewTrackChanges"],
  ["Mod-f", "FindDialog"], ["Mod-h", "ReplaceDialog"],
  ["Alt-=", "EquationInsert"],
];

/** "Mod-Shift-l" → "Ctrl+Shift+L", as Help › Keyboard Shortcuts writes keys. */
export function keyName(key: string): string {
  return key.split("-").map((p) => (p === "Mod" ? "Ctrl" : p.length === 1 ? p.toUpperCase() : p)).join("+");
}

/** Help › Keyboard Shortcuts for the Word interface, captions from the view's commands. */
export function wordShortcutSections(label: (id: WordCommandId) => string): ShortcutSection[] {
  const rows = new Map<WordCommandId, string[]>();
  for (const [key, id] of WORD_SHORTCUTS) rows.set(id, [...(rows.get(id) ?? []), keyName(key)]);
  return [["Word shortcuts", [...rows].map(([id, keys]) => [label(id), keys.join(", ")])]];
}
