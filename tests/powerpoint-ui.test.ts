import { describe, expect, it } from "vitest";
import { controlIds, type Tab } from "../src/presentation/common/ribbon";
import { PPT_COMMAND_IDS, PPT_FILE_MENU, PPT_NESTED, PPT_QAT, PPT_SHORTCUTS, powerpointTabs, type PowerPointControls } from "../src/presentation/pptx/powerpoint-ui";

const controls: PowerPointControls = {
  fonts: ["Calibri"],
  sizes: [18, 24],
  font: { value: () => "Calibri", set: () => {} },
  size: { value: () => "18", set: () => {} },
  color: () => {},
  highlight: () => {},
  shapeFill: () => {},
  shapeOutline: () => {},
  onShape: () => false,
  onPicture: () => false,
};
const tabIds = (tabs: Tab[]) => tabs.flatMap((t) => t.groups.flatMap((g) => [...g.controls.flatMap(controlIds), ...(g.launcher ? [g.launcher] : [])]));

describe("PowerPoint's ribbon layout", () => {
  const tabs = powerpointTabs(controls);

  it("has PowerPoint's tabs in order, contextual ones after", () => {
    expect(tabs.filter((t) => !t.context).map((t) => t.label)).toEqual(["Home", "Insert", "Design", "Transitions", "Animations", "Slide Show", "Review", "View", "Help"]);
    expect(tabs.filter((t) => t.context).map((t) => t.label)).toEqual(["Shape Format", "Picture Format"]);
  });

  it("has PowerPoint's groups in order", () => {
    const groups = (label: string) => tabs.find((t) => t.label === label && !t.context)!.groups.map((g) => g.label);
    expect(groups("Home")).toEqual(["Clipboard", "Slides", "Font", "Paragraph", "Drawing", "Editing"]);
    expect(groups("Insert")).toEqual(["Slides", "Tables", "Images", "Illustrations", "Links", "Comments", "Text", "Symbols", "Media"]);
    expect(groups("Design")).toEqual(["Themes", "Customize"]);
    expect(groups("Slide Show")).toEqual(["Start Slide Show", "Set Up"]);
    expect(groups("Review")).toEqual(["Comments", "Notes"]);
    expect(groups("View")).toEqual(["Presentation Views", "Show", "Zoom", "Dark Mode"]);
  });

  it("shows Shape Format on a shape and Picture Format on a picture", () => {
    const state = { shape: false, picture: false };
    const live = powerpointTabs({ ...controls, onShape: () => state.shape, onPicture: () => state.picture });
    const shown = () => live.filter((t) => t.context?.()).map((t) => t.label);
    expect(shown()).toEqual([]);
    state.shape = true;
    expect(shown()).toEqual(["Shape Format"]);
    state.shape = false;
    state.picture = true;
    expect(shown()).toEqual(["Picture Format"]);
  });

  it("names only commands the view provides", () => {
    const known = new Set<string>(PPT_COMMAND_IDS);
    const used = [...tabIds(tabs), ...PPT_QAT.defaults, ...PPT_QAT.choices, ...PPT_SHORTCUTS.map(([, id]) => id), ...Object.keys(PPT_NESTED), ...Object.values(PPT_NESTED).flat(), ...PPT_FILE_MENU];
    expect(used.filter((id) => !known.has(id))).toEqual([]);
  });

  it("reaches every command the Google Slides layout's toolbar and menus have", () => {
    const reachable = new Set([...tabIds(tabs), ...PPT_QAT.choices, ...PPT_FILE_MENU, ...Object.values(PPT_NESTED).flat()]);
    const GOOGLE_PARITY = [
      "Undo", "Redo", "ZoomDialog", "TextBoxInsert", "ShapesInsertGallery", "TableInsert", "Font", "FontSize", "Bold", "Italic", "Underline",
      "FontColorPicker", "TextHighlightColorPicker", "HyperlinkInsert", "ReviewNewComment", "PictureInsertFromFile", "AlignLeft", "AlignCenter",
      "AlignRight", "AlignJustify", "LineSpacing", "Bullets", "Numbering", "IndentDecrease", "IndentIncrease", "FormatPainter", "ClearFormatting",
      "ShapeFillColorPicker", "ShapeOutlineColorPicker", "ShapeOutlineWeight", "ShapeOutlineDashes", "PictureCrop", "PictureChange", "PictureReset",
      "SlideShowFromCurrent", "SlideShowFromBeginning", "ExportFile", "Cut", "Copy", "CopyAsMarkdownOutline", "CopyLinkToSlide",
      "CreateNoteFromSelection", "Paste", "PasteTextOnly", "SelectAll", "Delete", "FindDialog", "ViewSlideSorter", "FormatPane", "ShowComments",
      "LinkedFrom", "DarkModeOn", "MediaInsert", "ChartInsert", "FieldInsert", "SymbolInsert", "SlideNew", "Strikethrough", "Superscript",
      "Subscript", "CharacterSpacing", "ThemesGallery", "TransitionsPane", "AnimationPane", "ObjectBringToFront", "ObjectBringForward",
      "ObjectSendBackward", "ObjectSendToBack", "ObjectsAlignMenu", "ObjectsGroup", "ObjectsUngroup", "KeyboardShortcuts", "VersionHistory",
      "CompanionNote", "SlideDuplicate", "SlideDelete", "SlideHide",
    ];
    expect(GOOGLE_PARITY.filter((id) => !reachable.has(id))).toEqual([]);
  });

  it("starts the Quick Access Toolbar with Save, Undo, Redo, From Beginning", () => {
    expect(PPT_QAT.defaults).toEqual(["FileSave", "Undo", "Redo", "SlideShowFromBeginning"]);
  });
});

describe("PowerPoint's shortcuts", () => {
  it("are unique and include PowerPoint's keys, not Google's list keys", () => {
    const keys = PPT_SHORTCUTS.map(([k]) => k);
    expect(new Set(keys).size).toBe(keys.length);
    const map = new Map(PPT_SHORTCUTS);
    expect(map.get("Mod-e")).toBe("AlignCenter");
    expect(map.get("Mod-m")).toBe("SlideNew");
    expect(map.get("Mod-]")).toBe("FontSizeIncrease");
    expect(map.has("Mod-Shift-8")).toBe(false);
  });
});
