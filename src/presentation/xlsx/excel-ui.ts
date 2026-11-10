/**
 * Excel's interface (beta), as data: desktop Excel's ribbon tabs and groups in its order (Microsoft's command ids
 * where it has one, see docs/office-ribbon/excel.md), the Quick Access Toolbar and its shortcuts. The view supplies
 * a command for every id in XL_COMMAND_IDS.
 */
import { col, large as L, row, small as S, icon as I, type Tab } from "../common/ribbon";

/** Value controls the view provides, and the test of the contextual tab. */
export interface ExcelControls {
  fonts: string[];
  sizes: number[];
  formats: [label: string, code: string][];
  font: { value: () => string | null; set: (v: string) => void };
  size: { value: () => string | null; set: (v: string) => void };
  format: { value: () => string | null; set: (v: string) => void };
  fill: (hex: string) => void;
  color: (hex: string) => void;
  inPivot: () => boolean;
}

export const XL_COMMAND_IDS = [
  "FileSave", "Undo", "Redo", "FilePrint", "ExportFile", "VersionHistory",
  "Paste", "PasteValues", "Cut", "Copy", "FormatPainter",
  "Font", "FontSize", "FontSizeIncrease", "FontSizeDecrease", "Bold", "Italic", "Underline", "Strikethrough", "BordersMenu", "FillColorPicker",
  "FontColorPicker",
  "AlignTop", "AlignMiddle", "AlignBottom", "WrapText", "AlignLeft", "AlignCenter", "AlignRight", "MergeCenter",
  "NumberFormatGallery", "AccountingFormat", "PercentStyle", "DecimalsIncrease", "DecimalsDecrease",
  "ConditionalFormatting",
  "CellsInsertMenu", "CellsDeleteMenu", "CellsFormatMenu",
  "AutoSum", "ClearMenu", "ClearFormats", "ClearContents", "SortFilterMenu", "FindSelectMenu", "FindDialog", "ReplaceDialog", "SelectAll",
  "PivotTableInsert", "PictureInsertFromFile", "ChartInsert", "HyperlinkInsert", "ReviewNewComment", "CheckboxInsert", "DropdownInsert",
  "SymbolInsert", "SheetInsert",
  "PageSetup",
  "InsertFunctionMenu", "NameManager",
  "ImportCsv", "SortAscending", "SortDescending", "FilterToggle", "TextToColumns", "RemoveDuplicates", "DataValidation", "OutlineGroup",
  "OutlineUngroup",
  "ShowComments", "NewNote", "ProtectSheet", "CopyAsMarkdownTable", "CopyLinkToCells", "CreateNoteFromSelection", "CompanionNote", "LinkedFrom",
  "DarkModeOn", "ZoomDialog", "Zoom100", "FreezePanesMenu", "KeyboardShortcuts",
  "PivotRefresh", "PivotEdit",
  // Keyboard only.
  "GeneralFormat", "BorderOutside", "BorderNone", "HideRows", "HideColumns", "CopyFormat",
] as const;

export type XlCommandId = (typeof XL_COMMAND_IDS)[number];

export const XL_FILE_MENU: XlCommandId[] = ["FileSave", "ExportFile", "FilePrint", "VersionHistory"];

/** Commands reached through another command's menu (parent → children). */
export const XL_NESTED: Partial<Record<XlCommandId, XlCommandId[]>> = {
  Paste: ["PasteValues"],
  ClearMenu: ["ClearFormats", "ClearContents"],
  SortFilterMenu: ["SortAscending", "SortDescending", "FilterToggle"],
  FindSelectMenu: ["FindDialog", "ReplaceDialog", "SelectAll"],
  CellsInsertMenu: ["SheetInsert"],
  CellsFormatMenu: ["ProtectSheet"],
};

export function excelTabs(c: ExcelControls): Tab[] {
  return [
    { id: "home", label: "Home", groups: [
      { id: "clipboard", label: "Clipboard", icon: "clipboard", controls: [L("Paste"), col(S("Cut"), S("Copy"), S("FormatPainter"))] },
      { id: "font", label: "Font", icon: "type", controls: [col(
        row(
          { kind: "select", id: "Font", width: 9, options: () => c.fonts.map((f) => [f, f]), value: c.font.value, set: c.font.set },
          { kind: "select", id: "FontSize", width: 4, options: () => c.sizes.map((s) => [String(s), String(s)]), value: c.size.value, set: c.size.set },
          I("FontSizeIncrease"), I("FontSizeDecrease"),
        ),
        row(
          I("Bold"), I("Italic"), I("Underline"), I("Strikethrough"), I("BordersMenu"),
          { kind: "color", id: "FillColorPicker", initial: "#FFFF00", set: c.fill },
          { kind: "color", id: "FontColorPicker", initial: "#FF0000", set: c.color },
        ),
      )] },
      { id: "alignment", label: "Alignment", icon: "align-left", controls: [col(
        row(I("AlignTop"), I("AlignMiddle"), I("AlignBottom"), S("WrapText")),
        row(I("AlignLeft"), I("AlignCenter"), I("AlignRight"), S("MergeCenter")),
      )] },
      { id: "number", label: "Number", icon: "hash", controls: [col(
        { kind: "select", id: "NumberFormatGallery", width: 9, options: () => c.formats, value: c.format.value, set: c.format.set },
        row(I("AccountingFormat"), I("PercentStyle"), I("DecimalsIncrease"), I("DecimalsDecrease")),
      )] },
      { id: "styles", label: "Styles", icon: "palette", controls: [L("ConditionalFormatting")] },
      { id: "cells", label: "Cells", icon: "table-2", controls: [L("CellsInsertMenu"), L("CellsDeleteMenu"), L("CellsFormatMenu")] },
      { id: "editing", label: "Editing", icon: "sigma", controls: [col(S("AutoSum"), S("ClearMenu")), L("SortFilterMenu"), L("FindSelectMenu")] },
    ] },
    { id: "insert", label: "Insert", groups: [
      { id: "tables", label: "Tables", icon: "table-2", controls: [L("PivotTableInsert")] },
      { id: "illustrations", label: "Illustrations", icon: "image", controls: [L("PictureInsertFromFile")] },
      { id: "charts", label: "Charts", icon: "chart-column", controls: [L("ChartInsert")] },
      { id: "links", label: "Links", icon: "link", controls: [L("HyperlinkInsert")] },
      { id: "comments", label: "Comments", icon: "message-square-plus", controls: [L("ReviewNewComment")] },
      { id: "cellControls", label: "Cell Controls", icon: "square-check", controls: [L("CheckboxInsert"), L("DropdownInsert")] },
      { id: "symbols", label: "Symbols", icon: "omega", controls: [L("SymbolInsert")] },
    ] },
    { id: "pageLayout", label: "Page Layout", groups: [
      { id: "pageSetup", label: "Page Setup", icon: "file-cog", controls: [L("PageSetup")] },
    ] },
    { id: "formulas", label: "Formulas", groups: [
      { id: "functionLibrary", label: "Function Library", icon: "sigma", controls: [L("InsertFunctionMenu"), L("AutoSum")] },
      { id: "definedNames", label: "Defined Names", icon: "tag", controls: [L("NameManager")] },
    ] },
    { id: "data", label: "Data", groups: [
      { id: "getData", label: "Get & Transform Data", icon: "file-input", controls: [L("ImportCsv")] },
      { id: "sortFilter", label: "Sort & Filter", icon: "filter", controls: [col(S("SortAscending"), S("SortDescending")), L("FilterToggle")] },
      { id: "dataTools", label: "Data Tools", icon: "wrench", controls: [L("TextToColumns"), col(S("RemoveDuplicates"), S("DataValidation"))] },
      { id: "outline", label: "Outline", icon: "list-tree", controls: [L("OutlineGroup"), L("OutlineUngroup")] },
    ] },
    { id: "review", label: "Review", groups: [
      { id: "comments", label: "Comments", icon: "messages-square", controls: [L("ReviewNewComment"), L("ShowComments")] },
      { id: "notes", label: "Notes", icon: "sticky-note", controls: [L("NewNote")] },
      { id: "protect", label: "Protect", icon: "lock", controls: [L("ProtectSheet")] },
      { id: "linkedNotes", label: "Linked Notes", icon: "notebook-pen", controls: [L("CreateNoteFromSelection"), col(S("CopyAsMarkdownTable"), S("CopyLinkToCells"), S("CompanionNote")), col(S("LinkedFrom"))] },
    ] },
    { id: "view", label: "View", groups: [
      { id: "darkMode", label: "Dark Mode", icon: "moon", controls: [L("DarkModeOn")] },
      { id: "zoom", label: "Zoom", icon: "zoom-in", controls: [L("ZoomDialog"), L("Zoom100")] },
      { id: "window", label: "Window", icon: "snowflake", controls: [L("FreezePanesMenu")] },
    ] },
    { id: "help", label: "Help", groups: [{ id: "help", label: "Help", icon: "circle-help", controls: [L("KeyboardShortcuts")] }] },
    { id: "pivotAnalyze", label: "PivotTable Analyze", context: c.inPivot, groups: [
      { id: "data", label: "Data", icon: "refresh-cw", controls: [L("PivotRefresh")] },
      { id: "show", label: "Show", icon: "table-properties", controls: [L("PivotEdit")] },
    ] },
  ];
}

export const XL_QAT: { defaults: XlCommandId[]; choices: XlCommandId[] } = {
  defaults: ["FileSave", "Undo", "Redo"],
  choices: ["FileSave", "Undo", "Redo", "FilePrint", "SortAscending", "SortDescending", "FilterToggle", "AutoSum"],
};

/** Excel's shortcuts (ProseMirror-style names) for the commands Doc Works has. */
export const XL_SHORTCUTS: [key: string, id: XlCommandId][] = [
  ["Mod-b", "Bold"], ["Mod-i", "Italic"], ["Mod-u", "Underline"], ["Mod-5", "Strikethrough"],
  ["Mod-Shift-$", "AccountingFormat"], ["Mod-Shift-%", "PercentStyle"], ["Mod-Shift-~", "GeneralFormat"],
  ["Mod-Shift-&", "BorderOutside"], ["Mod-Shift-_", "BorderNone"],
  ["Alt-=", "AutoSum"], ["Mod-Shift-l", "FilterToggle"],
  ["Mod-9", "HideRows"], ["Mod-0", "HideColumns"],
  ["Mod-k", "HyperlinkInsert"], ["Mod-Alt-m", "ReviewNewComment"], ["Shift-F2", "NewNote"],
  ["Mod-f", "FindDialog"], ["Mod-h", "ReplaceDialog"],
];
