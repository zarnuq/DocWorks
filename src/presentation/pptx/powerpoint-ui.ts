/**
 * PowerPoint's interface (beta), as data: desktop PowerPoint's ribbon tabs and groups in its order (Microsoft's
 * command ids where it has one, see docs/office-ribbon/powerpoint.md), the Quick Access Toolbar and its shortcuts.
 * The view supplies a command for every id in PPT_COMMAND_IDS.
 */
import { col, large as L, row, small as S, icon as I, type Tab } from "../common/ribbon";

/** Value controls the view provides, and the tests of the contextual tabs. */
export interface PowerPointControls {
  fonts: string[];
  sizes: number[];
  font: { value: () => string | null; set: (v: string) => void };
  size: { value: () => string | null; set: (v: string) => void };
  color: (hex: string) => void;
  highlight: (hex: string) => void;
  shapeFill: (hex: string) => void;
  shapeOutline: (hex: string) => void;
  onShape: () => boolean;
  onPicture: () => boolean;
}

export const PPT_COMMAND_IDS = [
  "FileSave", "Undo", "Redo", "FilePrint", "ExportFile", "VersionHistory",
  "Paste", "PasteTextOnly", "Cut", "Copy", "FormatPainter",
  "SlideNew", "SlideDuplicate", "SlideDelete", "SlideHide", "SlideLayout", "SectionAdd",
  "Font", "FontSize", "FontSizeIncrease", "FontSizeDecrease", "ClearFormatting", "Bold", "Italic", "Underline", "Strikethrough", "Subscript",
  "Superscript", "CharacterSpacing", "FontColorPicker", "TextHighlightColorPicker",
  "Bullets", "Numbering", "IndentDecrease", "IndentIncrease", "AlignLeft", "AlignCenter", "AlignRight", "AlignJustify", "LineSpacing",
  "ShapesInsertGallery", "ArrangeMenu", "ShapeFillColorPicker", "ShapeOutlineColorPicker", "ShapeOutlineWeight", "ShapeOutlineDashes",
  "ObjectBringToFront", "ObjectBringForward", "ObjectSendBackward", "ObjectSendToBack", "ObjectsAlignMenu", "ObjectsGroup", "ObjectsUngroup",
  "FindDialog", "ReplaceDialog", "SelectMenu", "SelectAll", "Delete",
  "TableInsert", "PictureInsertFromFile", "ChartInsert", "HyperlinkInsert", "ReviewNewComment", "TextBoxInsert", "FieldInsert", "SymbolInsert",
  "MediaInsert",
  "ThemesGallery", "SlideSize", "FormatBackground", "TransitionsPane", "AnimationPane",
  "SlideShowFromBeginning", "SlideShowFromCurrent",
  "ShowComments", "CopyAsMarkdownOutline", "CopyLinkToSlide", "CreateNoteFromSelection", "CompanionNote", "LinkedFrom",
  "ViewNormal", "ViewSlideSorter", "FormatPane", "DarkModeOn", "ZoomDialog", "ZoomFitToWindow", "KeyboardShortcuts",
  "PictureCrop", "PictureChange", "PictureReset",
  // Keyboard only.
  "CopyFormat", "PasteFormat",
] as const;

export type PptCommandId = (typeof PPT_COMMAND_IDS)[number];

export const PPT_FILE_MENU: PptCommandId[] = ["FileSave", "ExportFile", "FilePrint", "VersionHistory"];

/** Commands reached through another command's menu (parent → children). */
export const PPT_NESTED: Partial<Record<PptCommandId, PptCommandId[]>> = {
  Paste: ["PasteTextOnly"],
  SelectMenu: ["SelectAll", "Delete"],
  ArrangeMenu: ["ObjectBringToFront", "ObjectBringForward", "ObjectSendBackward", "ObjectSendToBack", "ObjectsAlignMenu", "ObjectsGroup", "ObjectsUngroup"],
};

export function powerpointTabs(c: PowerPointControls): Tab[] {
  const arrange = [L("ArrangeMenu"), col(S("ObjectBringForward"), S("ObjectSendBackward"), S("ObjectsGroup"))];
  return [
    { id: "home", label: "Home", groups: [
      { id: "clipboard", label: "Clipboard", icon: "clipboard", controls: [L("Paste"), col(S("Cut"), S("Copy"), S("FormatPainter"))] },
      { id: "slides", label: "Slides", icon: "square-plus", controls: [L("SlideNew"), col(S("SlideLayout"), S("SlideDuplicate"), S("SlideDelete")), col(S("SectionAdd"))] },
      { id: "font", label: "Font", icon: "type", controls: [col(
        row(
          { kind: "select", id: "Font", width: 9, options: () => c.fonts.map((f) => [f, f]), value: c.font.value, set: c.font.set },
          { kind: "select", id: "FontSize", width: 4, options: () => c.sizes.map((s) => [String(s), String(s)]), value: c.size.value, set: c.size.set },
          I("FontSizeIncrease"), I("FontSizeDecrease"), I("ClearFormatting"),
        ),
        row(
          I("Bold"), I("Italic"), I("Underline"), I("Strikethrough"), I("Subscript"), I("Superscript"), I("CharacterSpacing"),
          { kind: "color", id: "TextHighlightColorPicker", initial: "#FFFF00", set: c.highlight },
          { kind: "color", id: "FontColorPicker", initial: "#FF0000", set: c.color },
        ),
      )] },
      { id: "paragraph", label: "Paragraph", icon: "pilcrow", controls: [col(
        row(I("Bullets"), I("Numbering"), I("IndentDecrease"), I("IndentIncrease"), I("LineSpacing")),
        row(I("AlignLeft"), I("AlignCenter"), I("AlignRight"), I("AlignJustify")),
      )] },
      { id: "drawing", label: "Drawing", icon: "shapes", controls: [L("ShapesInsertGallery"), L("ArrangeMenu"), col(
        { kind: "color", id: "ShapeFillColorPicker", initial: "#4472C4", set: c.shapeFill },
        { kind: "color", id: "ShapeOutlineColorPicker", initial: "#000000", set: c.shapeOutline },
      )] },
      { id: "editing", label: "Editing", icon: "search", controls: [col(S("FindDialog"), S("ReplaceDialog"), S("SelectMenu"))] },
    ] },
    { id: "insert", label: "Insert", groups: [
      { id: "slides", label: "Slides", icon: "square-plus", controls: [L("SlideNew")] },
      { id: "tables", label: "Tables", icon: "table", controls: [L("TableInsert")] },
      { id: "images", label: "Images", icon: "image", controls: [L("PictureInsertFromFile")] },
      { id: "illustrations", label: "Illustrations", icon: "shapes", controls: [L("ShapesInsertGallery"), L("ChartInsert")] },
      { id: "links", label: "Links", icon: "link", controls: [L("HyperlinkInsert")] },
      { id: "comments", label: "Comments", icon: "message-square-plus", controls: [L("ReviewNewComment")] },
      { id: "text", label: "Text", icon: "type", controls: [L("TextBoxInsert"), col(S("FieldInsert"))] },
      { id: "symbols", label: "Symbols", icon: "omega", controls: [L("SymbolInsert")] },
      { id: "media", label: "Media", icon: "film", controls: [L("MediaInsert")] },
    ] },
    { id: "design", label: "Design", groups: [
      { id: "themes", label: "Themes", icon: "palette", controls: [L("ThemesGallery")] },
      { id: "customize", label: "Customize", icon: "settings-2", controls: [L("SlideSize"), L("FormatBackground")] },
    ] },
    { id: "transitions", label: "Transitions", groups: [
      { id: "transitionToThisSlide", label: "Transition to This Slide", icon: "sparkles", controls: [L("TransitionsPane")] },
    ] },
    { id: "animations", label: "Animations", groups: [
      { id: "advancedAnimation", label: "Advanced Animation", icon: "sparkles", controls: [L("AnimationPane")] },
    ] },
    { id: "slideShow", label: "Slide Show", groups: [
      { id: "start", label: "Start Slide Show", icon: "play", controls: [L("SlideShowFromBeginning"), L("SlideShowFromCurrent")] },
      { id: "setUp", label: "Set Up", icon: "eye-off", controls: [L("SlideHide")] },
    ] },
    { id: "review", label: "Review", groups: [
      { id: "comments", label: "Comments", icon: "messages-square", controls: [L("ReviewNewComment"), L("ShowComments")] },
      { id: "notes", label: "Notes", icon: "notebook-pen", controls: [L("CreateNoteFromSelection"), col(S("CopyAsMarkdownOutline"), S("CopyLinkToSlide"), S("CompanionNote")), col(S("LinkedFrom"))] },
    ] },
    { id: "view", label: "View", groups: [
      { id: "presentationViews", label: "Presentation Views", icon: "presentation", controls: [L("ViewNormal"), L("ViewSlideSorter")] },
      { id: "show", label: "Show", icon: "eye", controls: [{ kind: "check", id: "FormatPane" }] },
      { id: "zoom", label: "Zoom", icon: "zoom-in", controls: [L("ZoomDialog"), L("ZoomFitToWindow")] },
      { id: "darkMode", label: "Dark Mode", icon: "moon", controls: [L("DarkModeOn")] },
    ] },
    { id: "help", label: "Help", groups: [{ id: "help", label: "Help", icon: "circle-help", controls: [L("KeyboardShortcuts")] }] },
    // Contextual tabs.
    { id: "shapeFormat", label: "Shape Format", context: c.onShape, groups: [
      { id: "shapeStyles", label: "Shape Styles", icon: "paint-bucket", controls: [col(
        { kind: "color", id: "ShapeFillColorPicker", initial: "#4472C4", set: c.shapeFill },
        { kind: "color", id: "ShapeOutlineColorPicker", initial: "#000000", set: c.shapeOutline },
      ), col(S("ShapeOutlineWeight"), S("ShapeOutlineDashes"))] },
      { id: "arrange", label: "Arrange", icon: "layers", controls: arrange },
    ] },
    { id: "pictureFormat", label: "Picture Format", context: c.onPicture, groups: [
      { id: "adjust", label: "Adjust", icon: "image", controls: [L("PictureChange"), col(S("PictureReset"))] },
      { id: "arrange", label: "Arrange", icon: "layers", controls: arrange },
      { id: "size", label: "Size", icon: "crop", controls: [L("PictureCrop")] },
    ] },
  ];
}

export const PPT_QAT: { defaults: PptCommandId[]; choices: PptCommandId[] } = {
  defaults: ["FileSave", "Undo", "Redo", "SlideShowFromBeginning"],
  choices: ["FileSave", "Undo", "Redo", "SlideShowFromBeginning", "FilePrint", "SlideNew", "FindDialog"],
};

/** PowerPoint's shortcuts (ProseMirror-style names) for the commands Doc Works has. */
export const PPT_SHORTCUTS: [key: string, id: PptCommandId][] = [
  ["Mod-b", "Bold"], ["Mod-i", "Italic"], ["Mod-u", "Underline"],
  ["Mod-=", "Subscript"], ["Mod-Shift-=", "Superscript"],
  ["Mod-]", "FontSizeIncrease"], ["Mod-[", "FontSizeDecrease"], ["Mod-Shift-.", "FontSizeIncrease"], ["Mod-Shift-,", "FontSizeDecrease"],
  ["Mod-Space", "ClearFormatting"],
  ["Mod-Shift-c", "CopyFormat"], ["Mod-Shift-v", "PasteFormat"],
  ["Mod-l", "AlignLeft"], ["Mod-e", "AlignCenter"], ["Mod-r", "AlignRight"], ["Mod-j", "AlignJustify"],
  // Ctrl+D (duplicate), Ctrl+G / Ctrl+Shift+G (group / ungroup) already act as PowerPoint's in both layouts.
  ["Mod-m", "SlideNew"],
  ["Mod-k", "HyperlinkInsert"], ["Mod-Alt-m", "ReviewNewComment"],
  ["Mod-f", "FindDialog"], ["Mod-h", "ReplaceDialog"],
];
