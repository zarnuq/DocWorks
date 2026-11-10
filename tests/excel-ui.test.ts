import { describe, expect, it } from "vitest";
import { controlIds, type Tab } from "../src/presentation/common/ribbon";
import { XL_COMMAND_IDS, XL_FILE_MENU, XL_NESTED, XL_QAT, XL_SHORTCUTS, excelTabs, type ExcelControls } from "../src/presentation/xlsx/excel-ui";

const controls: ExcelControls = {
  fonts: ["Calibri"],
  sizes: [10, 11],
  formats: [["General", "General"]],
  font: { value: () => "Calibri", set: () => {} },
  size: { value: () => "11", set: () => {} },
  format: { value: () => "General", set: () => {} },
  fill: () => {},
  color: () => {},
  inPivot: () => false,
};
const tabIds = (tabs: Tab[]) => tabs.flatMap((t) => t.groups.flatMap((g) => [...g.controls.flatMap(controlIds), ...(g.launcher ? [g.launcher] : [])]));

describe("Excel's ribbon layout", () => {
  const tabs = excelTabs(controls);

  it("has Excel's tabs in order, PivotTable Analyze in a pivot table", () => {
    expect(tabs.filter((t) => !t.context).map((t) => t.label)).toEqual(["Home", "Insert", "Page Layout", "Formulas", "Data", "Review", "View", "Help"]);
    expect(tabs.filter((t) => t.context).map((t) => t.label)).toEqual(["PivotTable Analyze"]);
    let pivot = false;
    const live = excelTabs({ ...controls, inPivot: () => pivot });
    expect(live.filter((t) => t.context?.()).length).toBe(0);
    pivot = true;
    expect(live.filter((t) => t.context?.()).map((t) => t.label)).toEqual(["PivotTable Analyze"]);
  });

  it("has Excel's groups in order", () => {
    const groups = (label: string) => tabs.find((t) => t.label === label && !t.context)!.groups.map((g) => g.label);
    expect(groups("Home")).toEqual(["Clipboard", "Font", "Alignment", "Number", "Styles", "Cells", "Editing"]);
    expect(groups("Insert")).toEqual(["Tables", "Illustrations", "Charts", "Links", "Comments", "Cell Controls", "Symbols"]);
    expect(groups("Page Layout")).toEqual(["Page Setup"]);
    expect(groups("Formulas")).toEqual(["Function Library", "Defined Names"]);
    expect(groups("Data")).toEqual(["Get & Transform Data", "Sort & Filter", "Data Tools", "Outline"]);
    expect(groups("Review")).toEqual(["Comments", "Notes", "Protect", "Linked Notes"]);
    expect(groups("View")).toEqual(["Dark Mode", "Zoom", "Window"]);
  });

  it("names only commands the view provides", () => {
    const known = new Set<string>(XL_COMMAND_IDS);
    const used = [...tabIds(tabs), ...XL_QAT.defaults, ...XL_QAT.choices, ...XL_SHORTCUTS.map(([, id]) => id), ...Object.keys(XL_NESTED), ...Object.values(XL_NESTED).flat(), ...XL_FILE_MENU];
    expect(used.filter((id) => !known.has(id))).toEqual([]);
  });

  it("reaches every command the Google Sheets layout's toolbar and menus have", () => {
    const reachable = new Set([...tabIds(tabs), ...XL_QAT.choices, ...XL_FILE_MENU, ...Object.values(XL_NESTED).flat()]);
    const GOOGLE_PARITY = [
      "Undo", "Redo", "ZoomDialog", "AccountingFormat", "PercentStyle", "DecimalsDecrease", "DecimalsIncrease", "NumberFormatGallery", "Font",
      "FontSize", "Bold", "Italic", "Strikethrough", "Underline", "FontColorPicker", "HyperlinkInsert", "ReviewNewComment", "FillColorPicker",
      "BordersMenu", "MergeCenter", "AlignLeft", "AlignCenter", "AlignRight", "AlignTop", "AlignMiddle", "AlignBottom", "WrapText", "FormatPainter",
      "ClearFormats", "ChartInsert", "FilterToggle", "InsertFunctionMenu", "ExportFile", "Cut", "Copy", "CopyAsMarkdownTable", "CopyLinkToCells",
      "CreateNoteFromSelection", "Paste", "PasteValues", "SelectAll", "ClearContents", "FindDialog", "ReplaceDialog", "FreezePanesMenu",
      "ShowComments", "LinkedFrom", "DarkModeOn", "SheetInsert", "PictureInsertFromFile", "PivotTableInsert", "CheckboxInsert", "NewNote",
      "DropdownInsert", "SymbolInsert", "ConditionalFormatting", "SortAscending", "SortDescending", "DataValidation", "RemoveDuplicates",
      "TextToColumns", "NameManager", "OutlineGroup", "OutlineUngroup", "ImportCsv", "KeyboardShortcuts", "VersionHistory", "CompanionNote",
    ];
    expect(GOOGLE_PARITY.filter((id) => !reachable.has(id))).toEqual([]);
  });

  it("starts the Quick Access Toolbar with Save, Undo, Redo", () => {
    expect(XL_QAT.defaults).toEqual(["FileSave", "Undo", "Redo"]);
  });
});

describe("Excel's shortcuts", () => {
  it("are unique and include Excel's keys", () => {
    const keys = XL_SHORTCUTS.map(([k]) => k);
    expect(new Set(keys).size).toBe(keys.length);
    const map = new Map(XL_SHORTCUTS);
    expect(map.get("Mod-5")).toBe("Strikethrough");
    expect(map.get("Mod-Shift-l")).toBe("FilterToggle");
    expect(map.get("Alt-=")).toBe("AutoSum");
    expect(map.get("Mod-9")).toBe("HideRows");
    expect(map.get("Mod-0")).toBe("HideColumns");
    expect(map.has("Mod-Shift-x")).toBe(false);
  });
});
