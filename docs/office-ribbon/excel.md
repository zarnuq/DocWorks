# excel ribbon controls (Microsoft 365 Current Channel)

Outline of Microsoft's Office Fluent UI command identifiers (excel), from github.com/OfficeDev/office-fluent-ui-command-identifiers (MIT License, Copyright (c) Microsoft Corporation). Rows: tab (tab set), group, controls in ribbon order; indented controls are inside the one after '<'. Dev reference for the Office-style interface; not shipped.

## Quick Access Toolbar  (None (Quick Access Toolbar))
### 
- AutoSaveSwitch [toggleButton]
- FileNewDefault [button]
- FileOpenUsingBackstage [button]
- FileSave [button]
- FileSendAsAttachment [button]
- FilePrintQuick [button]
- PrintPreviewAndPrint [button]
- Spelling [button]
- Undo [gallery]
- Redo [gallery]
- SortAscendingExcel [button]
- SortDescendingExcel [button]
- PointerModeOptions [gallery]

##   (None (Core Tab))
### 
- TabHome [tab]
- TabInsert [tab]
- TabDrawInk [tab]
- TabPageLayoutExcel [tab]
- TabFormulas [tab]
- TabData [tab]
- TabReview [tab]
- TabView [tab]
- TabAutomate [tab]
- TabDeveloper [tab]
- TabAddIns [tab]
- HelpTab [tab]
- TabPrintPreview [tab]
- TabBackgroundRemoval [tab]

## TabHome  (None (Core Tab))
### 
- GroupClipboard [group]
- GroupFont [group]
- GroupAlignmentExcel [group]
- GroupNumber [group]
- GroupStyles [group]
- GroupCells [group]
- GroupEditingExcel [group]
- GroupActivation [group]
- GroupIdeas [group]
- GroupClassifyLabelProtect [group]
- GroupOfficeExtensionsAddinFlyout [group]
- GroupHomePowerOptions [group]
- GroupAIAssistance [group]
### GroupClipboard
- PasteMenu [splitButton]
-   Paste [button] < PasteMenu
-   PasteGallery [gallery] < PasteMenu
-     PasteUsingTextImportWizard [button] < PasteMenu
-     PasteRefreshableWebQuery [button] < PasteMenu
-     PasteSpecialDialog [button] < PasteMenu
- Cut [button]
- CopySplitButton [splitButton]
-   Copy [button] < CopySplitButton
-   CopyAsPicture [button] < CopySplitButton
- FormatPainter [control]
- ShowClipboard [button (dialogBoxLauncher)]
### GroupFont
- Font [comboBox]
- FontSize [comboBox]
- FontSizeIncrease [button]
- FontSizeDecrease [button]
- Bold [toggleButton]
- Italic [toggleButton]
- UnderlineGallery [splitButton]
-   Underline [toggleButton] < UnderlineGallery
-   UnderlineDouble [toggleButton] < UnderlineGallery
- BordersGallery [splitButton]
-   BorderBottomNoToggle [button] < BordersGallery
-   BorderTopNoToggle [button] < BordersGallery
-   BorderLeftNoToggle [button] < BordersGallery
-   BorderRightNoToggle [button] < BordersGallery
-   BorderNone [button] < BordersGallery
-   BordersAll [button] < BordersGallery
-   BorderOutside [button] < BordersGallery
-   BorderThickOutside [button] < BordersGallery
-   BorderDoubleBottom [button] < BordersGallery
-   BorderThickBottom [button] < BordersGallery
-   BorderTopAndBottom [button] < BordersGallery
-   BorderTopAndThickBottom [button] < BordersGallery
-   BorderTopAndDoubleBottom [button] < BordersGallery
-   BorderDrawLine [toggleButton] < BordersGallery
-   BorderDrawGrid [toggleButton] < BordersGallery
-   BorderErase [toggleButton] < BordersGallery
-   BorderColorPickerExcel [gallery] < BordersGallery
-     BorderMoreColorsDialog [button] < BordersGallery
-   BorderStyle [dropDown] < BordersGallery
-   BordersMoreDialog [button] < BordersGallery
- CellFillColorPicker [gallery]
-   FontShadingColorMoreColorsDialog [button] < CellFillColorPicker
- FontColorPicker [gallery]
-   FontColorMoreColorsDialogExcel [button] < FontColorPicker
- PhoneticGuideMenu [splitButton]
-   PhoneticGuideFieldShow [toggleButton] < PhoneticGuideMenu
-   PhoneticGuideEdit [button] < PhoneticGuideMenu
-   PhoneticGuideSettings [button] < PhoneticGuideMenu
- FormatCellsFontDialog [button (dialogBoxLauncher)]
### GroupAlignmentExcel
- AlignTopExcel [toggleButton]
- AlignMiddleExcel [toggleButton]
- AlignBottomExcel [toggleButton]
- OrientationMenu [menu]
-   TextOrientationAngleCounterclockwise [toggleButton] < OrientationMenu
-   TextOrientationAngleClockwise [toggleButton] < OrientationMenu
-   TextOrientationVertical [toggleButton] < OrientationMenu
-   TextOrientationRotateUp [toggleButton] < OrientationMenu
-   TextOrientationRotateDown [toggleButton] < OrientationMenu
-   CellAlignmentOptions [button] < OrientationMenu
- AlignLeftToRightMenu [splitButton]
-   TextDirectionLeftToRight [toggleButton] < AlignLeftToRightMenu
-   TextDirectionRightToLeft [toggleButton] < AlignLeftToRightMenu
-   TextDirectionContext [toggleButton] < AlignLeftToRightMenu
- AlignLeft [toggleButton]
- AlignCenter [toggleButton]
- AlignRight [toggleButton]
- IndentDecreaseExcel [button]
- IndentIncreaseExcel [button]
- WrapText [toggleButton]
- MergeCenterMenu [splitButton]
-   MergeCenter [toggleButton] < MergeCenterMenu
-   MergeCellsAcross [button] < MergeCenterMenu
-   MergeCells [button] < MergeCenterMenu
-   UnmergeCells [button] < MergeCenterMenu
-   CenterAcrossSelection [toggleButton] < MergeCenterMenu
- CellAlignmentOptions [button (dialogBoxLauncher)]
### GroupNumber
- NumberFormatGallery [comboBox]
-   NumberFormatsDialog [button] < NumberFormatGallery
- AccountingFormatMenu [splitButton]
-   AccountingFormat [button] < AccountingFormatMenu
-   AccountingFormatExcel [button] < AccountingFormatMenu
-   AccountingFormatMoreExcel [button] < AccountingFormatMenu
- PercentStyle [button]
- CommaStyle [button]
- DecimalsIncrease [button]
- DecimalsDecrease [button]
- FormatCellsNumberDialog [button (dialogBoxLauncher)]
### GroupStyles
- ConditionalFormattingMenu [menu]
-   ConditionalFormattingHighlightCellsMenu [menu] < ConditionalFormattingMenu
-     ConditionalFormattingHighlightGreaterThan [button] < ConditionalFormattingMenu
-     ConditionalFormattingHighlightLessThan [button] < ConditionalFormattingMenu
-     ConditionalFormattingHighlightBetween [button] < ConditionalFormattingMenu
-     ConditionalFormattingHighlightEqualTo [button] < ConditionalFormattingMenu
-     ConditionalFormattingHighlightTextContaining [button] < ConditionalFormattingMenu
-     ConditionalFormattingHighlightDateOccuring [button] < ConditionalFormattingMenu
-     ConditionalFormattingHighlightDuplicateValues [button] < ConditionalFormattingMenu
-     ConditionalFormattingHighlightRulesMore [button] < ConditionalFormattingMenu
-   ConditionalFormattingTopBottomMenu [menu] < ConditionalFormattingMenu
-     ConditionalFormattingTopNItems [button] < ConditionalFormattingMenu
-     ConditionalFormattingTopNPercent [button] < ConditionalFormattingMenu
-     ConditionalFormattingBottomNItems [button] < ConditionalFormattingMenu
-     ConditionalFormattingBottomNPercent [button] < ConditionalFormattingMenu
-     ConditionalFormattingAboveAverage [button] < ConditionalFormattingMenu
-     ConditionalFormattingBelowAverage [button] < ConditionalFormattingMenu
-     ConditionalFormattingTopBottomMore [button] < ConditionalFormattingMenu
-   ConditionalFormattingDataBarsGallery [gallery] < ConditionalFormattingMenu
-     ConditionalFormattingDataBarsMoreOptions [button] < ConditionalFormattingMenu
-   ConditionalFormattingColorScalesGallery [gallery] < ConditionalFormattingMenu
-     ConditionalFormattingColorScalesMore [button] < ConditionalFormattingMenu
-   ConditionalFormattingIconSetsGallery [gallery] < ConditionalFormattingMenu
-     ConditionalFormattingIconSetsMore [button] < ConditionalFormattingMenu
-   ConditionalFormattingNewRule [button] < ConditionalFormattingMenu
-   ConditionalFormattingClearMenu [menu] < ConditionalFormattingMenu
-     ConditionalFormattingClearSelectedCells [button] < ConditionalFormattingMenu
-     ConditionalFormattingClearSheet [button] < ConditionalFormattingMenu
-     ConditionalFormattingClearTable [button] < ConditionalFormattingMenu
-     ConditionalFormattingClearPivotTable [button] < ConditionalFormattingMenu
-   ConditionalFormattingsManage [button] < ConditionalFormattingMenu
- FormatAsTableGallery [gallery]
-   TableStyleNew [button] < FormatAsTableGallery
-   PivotTableNewStyle [button] < FormatAsTableGallery
- CellStylesGallery [gallery]
-   CellStyleNew [button] < CellStylesGallery
-   CellStylesMerge [button] < CellStylesGallery
### GroupCells
- InsertCellstMenu [splitButton]
-   CellsInsertSmart [button] < InsertCellstMenu
-   CellsInsertDialog [button] < InsertCellstMenu
-   SheetRowsInsert [button] < InsertCellstMenu
-   SheetColumnsInsert [button] < InsertCellstMenu
-   TableRowsInsertAboveExcel [button] < InsertCellstMenu
-   TableRowsInsertBelowExcel [button] < InsertCellstMenu
-   TableColumnsInsertLeftExcel [button] < InsertCellstMenu
-   TableColumnsInsertRightExcel [button] < InsertCellstMenu
-   SheetInsert [button] < InsertCellstMenu
- TableDeleteRowsAndColumnsMenu [splitButton]
-   CellsDeleteSmart [button] < TableDeleteRowsAndColumnsMenu
-   CellsDelete [button] < TableDeleteRowsAndColumnsMenu
-   SheetRowsDelete [button] < TableDeleteRowsAndColumnsMenu
-   SheetColumnsDelete [button] < TableDeleteRowsAndColumnsMenu
-   TableRowsDeleteExcel [button] < TableDeleteRowsAndColumnsMenu
-   TableColumnsDeleteExcel [button] < TableDeleteRowsAndColumnsMenu
-   SheetDelete [button] < TableDeleteRowsAndColumnsMenu
- FormatCellsMenu [menu]
-   RowHeight [button] < FormatCellsMenu
-   RowHeightAutoFit [button] < FormatCellsMenu
-   ColumnWidth [button] < FormatCellsMenu
-   ColumnWidthAutoFit [button] < FormatCellsMenu
-   ColumnWidthDefault [button] < FormatCellsMenu
-   HideAndUnhideMenu [menu] < FormatCellsMenu
-     RowsHide [button] < FormatCellsMenu
-     ColumnsHide [button] < FormatCellsMenu
-     SheetHide [button] < FormatCellsMenu
-     RowsUnhide [button] < FormatCellsMenu
-     ColumnsUnhide [button] < FormatCellsMenu
-     SheetUnhide [button] < FormatCellsMenu
-   SheetRename [button] < FormatCellsMenu
-   SheetMoveOrCopy [button] < FormatCellsMenu
-   SheetTabColorGallery [gallery] < FormatCellsMenu
-     SheetTabColorMoreColorsDialog [button] < FormatCellsMenu
-   CollapseArrays [toggleButton] < FormatCellsMenu
-   SheetProtect [button] < FormatCellsMenu
-   Lock [toggleButton] < FormatCellsMenu
-   FormatCellsDialog [button] < FormatCellsMenu
### GroupEditingExcel
- AutoSumMenu [splitButton]
-   AutoSum [button] < AutoSumMenu
-   AutoSumAverage [button] < AutoSumMenu
-   AutoSumCount [button] < AutoSumMenu
-   AutoSumMax [button] < AutoSumMenu
-   AutoSumMin [button] < AutoSumMenu
-   AutoSumMoreFunctions [button] < AutoSumMenu
- FillMenu [menu]
-   FillDown [button] < FillMenu
-   FillRight [button] < FillMenu
-   FillUp [button] < FillMenu
-   FillLeft [button] < FillMenu
-   FillAcrossWorksheets [button] < FillMenu
-   FillSeries [button] < FillMenu
-   FillJustify [button] < FillMenu
-   QuickCode [button] < FillMenu
- ClearMenu [menu]
-   ClearAll [button] < ClearMenu
-   ClearFormats [button] < ClearMenu
-   ClearContents [button] < ClearMenu
-   ClearComments [button] < ClearMenu
-   ClearThreadedCommentsAndNotes [button] < ClearMenu
-   HyperlinksClear [button] < ClearMenu
-   HyperlinksRemove [button] < ClearMenu
-   CleanupWorkbook [button] < ClearMenu
- SortFilterMenu [menu]
-   SortAscendingExcel [button] < SortFilterMenu
-   SortDescendingExcel [button] < SortFilterMenu
-   SortCustomExcel [button] < SortFilterMenu
-   Filter [toggleButton] < SortFilterMenu
-   SortClear [button] < SortFilterMenu
-   FilterReapply [button] < SortFilterMenu
- SelectMenuExcel [menu]
-   FindDialogExcel [button] < SelectMenuExcel
-   ReplaceDialog [button] < SelectMenuExcel
-   GoTo [button] < SelectMenuExcel
-   GoToSpecial [button] < SelectMenuExcel
-   GoToFormulas [button] < SelectMenuExcel
-   GoToComments [button] < SelectMenuExcel
-   GoToCommentsLegacy [button] < SelectMenuExcel
-   GoToConditionalFormatting [button] < SelectMenuExcel
-   GoToConstants [button] < SelectMenuExcel
-   GoToDataValidation [button] < SelectMenuExcel
-   ObjectsSelect [toggleButton] < SelectMenuExcel
-   SelectionPane [toggleButton] < SelectMenuExcel
### GroupActivation
- ActivateOffice [button]
### GroupIdeas
- DataCleansing [button]
- AnalyzeData [button]
### GroupClassifyLabelProtect
- ClassifyLabelProtect [menu]
-   None [gallery] < ClassifyLabelProtect
-   None [toggleButton] < ClassifyLabelProtect
-   None [gallery] < ClassifyLabelProtect
-   None [gallery] < ClassifyLabelProtect
-   None [button] < ClassifyLabelProtect
-   None [button] < ClassifyLabelProtect
### GroupOfficeExtensionsAddinFlyout
- OfficeExtensionsShowAddinFlyout [button]
- OfficeExtensionsShowAddinFlyout [toggleButton]
### GroupHomePowerOptions
- UXPPPowerOptions [button]
### GroupAIAssistance
- DataCleansing [button]
- AnalyzeData [button]
- Copilot [button]

## TabInsert  (None (Core Tab))
### 
- GroupInsertTablesExcel [group]
- GroupInsertIllustrations [group]
- GroupOfficeExtension [group]
- GroupInsertChartsExcel [group]
- GroupPowerMap [group]
- GroupSparklinesInsert [group]
- GroupSlicerInsert [group]
- GroupInsertLinks [group]
- GroupInsertComments [group]
- GroupCellControls [group]
- GroupInsertText [group]
- GroupInsertSymbols [group]
### GroupInsertTablesExcel
- PivotTableInsert [button]
- InsertPivotTableSplit [gallery]
-   NewPivotTable [button] < InsertPivotTableSplit
-   PowerBIDataInsert [button] < InsertPivotTableSplit
- InsertPivotTableDropdown [gallery]
-   CreatePivotTableFromTableRange [button] < InsertPivotTableDropdown
-   CreatePivotTableFromExternalDataSource [button] < InsertPivotTableDropdown
-   CreatePivotTableFromDataModel [button] < InsertPivotTableDropdown
-   PowerBIDataInsert [button] < InsertPivotTableDropdown
- PivotTableSuggestion [button]
- TableInsertExcel [button]
- SplitButtonMsForm [gallery]
-   MsFormInsert [button] < SplitButtonMsForm
-   MsFormPreview [button] < SplitButtonMsForm
-   MsFormEdit [button] < SplitButtonMsForm
-   MsFormSend [button] < SplitButtonMsForm
-   MsFormDelete [button] < SplitButtonMsForm
### GroupInsertIllustrations
- FlyoutAnchorInsertPictures [gallery]
-   PictureInsertFromFile [button] < FlyoutAnchorInsertPictures
-   InsertMobilePicture [button] < FlyoutAnchorInsertPictures
-   InsertM365Picture [button] < FlyoutAnchorInsertPictures
-   ClipArtInsertDialog [button] < FlyoutAnchorInsertPictures
- ShapesInsertGallery [gallery]
- IconInsertFromFile [button]
- Insert3DModelDropdown [splitButton]
-   Insert3DModelDefault [button] < Insert3DModelDropdown
-   Insert3DModelFromFile [button] < Insert3DModelDropdown
-   Insert3DModelFromFileConsumer [button] < Insert3DModelDropdown
-   Insert3DModelFromOnline [button] < Insert3DModelDropdown
-   Insert3DModelFromOnlineConsumer [button] < Insert3DModelDropdown
- SmartArtInsert [button]
- Insert3DModelFallback [button]
- ScreenshotInsertGallery [gallery]
-   ScreenClipping [button] < ScreenshotInsertGallery
- InsertPicturesInCellOverCellsMenu [menu]
-   InsertPicturePlaceInCellsMenu [menu] < InsertPicturesInCellOverCellsMenu
-     InsertCellPictureFromFile [button] < InsertPicturesInCellOverCellsMenu
-     InsertCellMobilePicture [button] < InsertPicturesInCellOverCellsMenu
-     InsertCellM365Picture [button] < InsertPicturesInCellOverCellsMenu
-     InsertCellClipArtDialog [button] < InsertPicturesInCellOverCellsMenu
-   InsertPicturePlaceOverCellsMenu [menu] < InsertPicturesInCellOverCellsMenu
-     PictureInsertFromFile [button] < InsertPicturesInCellOverCellsMenu
-     InsertMobilePicture [button] < InsertPicturesInCellOverCellsMenu
-     InsertM365Picture [button] < InsertPicturesInCellOverCellsMenu
-     ClipArtInsertDialog [button] < InsertPicturesInCellOverCellsMenu
### GroupOfficeExtension
- OfficeExtensionsAppStore [button]
- OfficeExtensionsGallery2 [gallery]
-   OfficeExtensionsDialog [button] < OfficeExtensionsGallery2
-   OfficeExtensionsManageOtherAddins [button] < OfficeExtensionsGallery2
-   ExtensibilityExperimentation [button] < OfficeExtensionsGallery2
- OfficeExtensionsFeaturedApp1 [button]
- OfficeExtensionsFeaturedApp2 [button]
- OfficeExtensionsFeaturedApp3 [button]
- OfficeExtensionsFeaturedApp4 [button]
- OfficeExtensionsFeaturedApp5 [button]
- OfficeExtensionsFeaturedApp6 [button]
### GroupInsertChartsExcel
- ChartInsertGalleryNew [button]
- ChartTypeColumnInsertGallery [gallery]
-   ChartTypeColumnInsertDialog [button] < ChartTypeColumnInsertGallery
- ChartTypeLineInsertGallery [gallery]
-   ChartTypeLineInsertDialog [button] < ChartTypeLineInsertGallery
- ChartTypePieInsertGallery [gallery]
-   ChartTypePieInsertDialog [button] < ChartTypePieInsertGallery
- ChartTypeHierarchyInsertGallery [gallery]
-   ChartTypeHierarchyInsertDialog [button] < ChartTypeHierarchyInsertGallery
- ChartTypeStatisticalInsertGallery [gallery]
-   ChartTypeStatisticalInsertDialog [button] < ChartTypeStatisticalInsertGallery
- ChartTypeXYScatterInsertGallery [gallery]
-   ChartTypeScatterInsertDialog [button] < ChartTypeXYScatterInsertGallery
- ChartTypeWaterfallInsertChartGallery [gallery]
-   ChartTypeStockInsertDialog [button] < ChartTypeWaterfallInsertChartGallery
- ChartTypeComboInsertGallery [gallery]
-   ChartTypeComboInsertDialog [button] < ChartTypeComboInsertGallery
- ChartTypeCustomVisualInsertGallery [gallery]
-   ChartTypeCustomVisualInsertDialog [button] < ChartTypeCustomVisualInsertGallery
- ChartTypeMapsInsertGallery [gallery]
-   ChartTypeMapsInsertDialog [button] < ChartTypeMapsInsertGallery
- NewPivotChartSplitButton [splitButton]
-   PivotChartInsert [button] < NewPivotChartSplitButton
-   SxNewPivotChartAndTable [button] < NewPivotChartSplitButton
- ChartDialogLauncherAll [button (dialogBoxLauncher)]
### GroupPowerMap
- SplitButtonPowerMap [gallery]
-   PowerMapLaunch [button] < SplitButtonPowerMap
-   PowerMapAddData [button] < SplitButtonPowerMap
### GroupSparklinesInsert
- SparklineLineInsert [button]
- SparklineColumnInsert [button]
- SparklineWinLossInsert [button]
### GroupSlicerInsert
- SlicerInsert [button]
- InsertTimeline [button]
### GroupInsertLinks
- InsertLinkGallery [gallery]
-   None [button] < InsertLinkGallery
- HyperlinkInsert [button]
### GroupInsertComments
- InsertNewComment [button]
### GroupCellControls
- InsertInCellCheckbox [button]
- InCellStarRating [button]
- InCellPictureCellFormat [button]
### GroupInsertText
- TextBoxInsertExcel [toggleButton]
- TextBoxInsertMenu [splitButton]
-   TextBoxInsertHorizontal [toggleButton] < TextBoxInsertMenu
-   TextBoxInsertVertical [toggleButton] < TextBoxInsertMenu
- HeaderFooterInsert [button]
- WordArtInsertGallery [gallery]
- SignatureLineInsertMenu [splitButton]
-   SignatureLineInsert [button] < SignatureLineInsertMenu
-   SignatureServicesAdd [button] < SignatureLineInsertMenu
- OleObjectctInsert [button]
### GroupInsertSymbols
- InsertBuildingBlocksEquationsGallery [gallery]
-   EquationInsertNew [button] < InsertBuildingBlocksEquationsGallery
-   InkEquation [button] < InsertBuildingBlocksEquationsGallery
- SymbolInsert [button]

## TabDrawInk  (None (Core Tab))
### 
- GroupInputMode [group]
- GroupWrite [group]
- GroupDrawingTools [group]
- GroupPensOneNote [group]
- GroupPens2 [group]
- GroupDrawWithTouch [group]
- GroupInkConvertOneNote [group]
- InkReplay [group]
- GroupInkClose [group]
- GroupPenAndInkHelp [group]
### GroupInputMode
- InkSelect [toggleButton]
- FingerPaintingMode [toggleButton]
- DrawInkMode [toggleButton]
### GroupWrite
- InkSelect [toggleButton]
- FingerPaintingMode [toggleButton]
- DrawInkMode [toggleButton]
- InkEraser [toggleButton]
- LassoSelect [toggleButton]
### GroupDrawingTools
- InkToolboxWithEraserAndLasso [control]
### GroupPensOneNote
- InkToolbox [control]
### GroupPens2
- PensGallery2 [gallery]
-   PenAddFavorite [button] < PensGallery2
- InkColorPicker2 [gallery]
-   InkColorMoreColorsDialog [button] < InkColorPicker2
- LineThickness2 [gallery]
-   LineStylesDialog [button] < LineThickness2
### GroupDrawWithTouch
- FingerPaintingMode [toggleButton]
### GroupInkConvertOneNote
- DiagrammingOnline [button]
- SelectionToMathConvert [button]
### InkReplay
- Replay [toggleButton]
### GroupInkClose
- InkStopInkingReadingView [button]
### GroupPenAndInkHelp
- PenAndInkHelp [button]

## TabPageLayoutExcel  (None (Core Tab))
### 
- GroupThemesExcel [group]
- GroupPageSetup [group]
- GroupPageLayoutScaleToFit [group]
- GroupPageLayoutSheetOptions [group]
- GroupArrange [group]
### GroupThemesExcel
- ThemesGallery [gallery]
-   ThemeSearchOfficeOnline [button] < ThemesGallery
-   ThemeBrowseForThemes [button] < ThemesGallery
-   ThemeSaveCurrent [button] < ThemesGallery
- ThemeColorsGallery [gallery]
-   ThemeColorsCreateNew [button] < ThemeColorsGallery
- ThemeFontsGallery [gallery]
-   ThemeFontsCreateNew [button] < ThemeFontsGallery
- ThemeEffectsGallery [gallery]
### GroupPageSetup
- PageMarginsGallery [gallery]
-   MarginsCustomMargins [button] < PageMarginsGallery
- PageOrientationGallery [gallery]
- PageSizeGallery [gallery]
-   PageSizeMorePaperSizesDialogExcel [button] < PageSizeGallery
- PrintAreaMenu [menu]
-   PrintAreaSetPrintArea [button] < PrintAreaMenu
-   PrintAreaClearPrintArea [button] < PrintAreaMenu
-   PrintAreaAddToPrintArea [button] < PrintAreaMenu
- PageBreakMenu [menu]
-   PageBreakInsertExcel [button] < PageBreakMenu
-   PageBreakRemove [button] < PageBreakMenu
-   PageBreaksResetAll [button] < PageBreakMenu
- SheetBackground [button]
- PrintTitles [button]
- PageSetupPageDialog [button (dialogBoxLauncher)]
### GroupPageLayoutScaleToFit
- PageScaleToFitWidth [comboBox]
-   PageScaleToFitOptionsDialog [button] < PageScaleToFitWidth
- PageScaleToFitHeight [comboBox]
-   PageScaleToFitOptionsDialog [button] < PageScaleToFitHeight
- PageScaleToFitScale [control]
- PageSetupPageDialog [button (dialogBoxLauncher)]
### GroupPageLayoutSheetOptions
- RightToLeftDocument [toggleButton]
- GridlinesLabel [labelControl]
- ViewSheetGridlines [checkBox]
- SheetPrintGridlines [checkBox]
- HeadingsLabel [labelControl]
- ViewSheetHeadings [checkBox]
- SheetPrintHeadings [checkBox]
- PageSetupSheetDialog [button (dialogBoxLauncher)]
### GroupArrange
- ObjectBringForwardMenu [splitButton]
-   ObjectBringForward [button] < ObjectBringForwardMenu
-   ObjectBringToFront [button] < ObjectBringForwardMenu
- ObjectSendBackwardMenu [splitButton]
-   ObjectSendBackward [button] < ObjectSendBackwardMenu
-   ObjectSendToBack [button] < ObjectSendBackwardMenu
- SelectionPane [toggleButton]
- ObjectAlignMenu [menu]
-   ObjectsAlignLeft [button] < ObjectAlignMenu
-   ObjectsAlignCenterHorizontal [button] < ObjectAlignMenu
-   ObjectsAlignRight [button] < ObjectAlignMenu
-   ObjectsAlignTop [button] < ObjectAlignMenu
-   ObjectsAlignMiddleVertical [button] < ObjectAlignMenu
-   ObjectsAlignBottom [button] < ObjectAlignMenu
-   AlignDistributeHorizontallyClassic [button] < ObjectAlignMenu
-   AlignDistributeVerticallyClassic [button] < ObjectAlignMenu
-   SnapToGrid [toggleButton] < ObjectAlignMenu
-   SnapToShapes [toggleButton] < ObjectAlignMenu
-   ViewGridlinesToggleExcel [toggleButton] < ObjectAlignMenu
- ObjectsGroupMenu [menu]
-   ObjectsGroup [button] < ObjectsGroupMenu
-   ObjectsRegroup [button] < ObjectsGroupMenu
-   ObjectsUngroup [button] < ObjectsGroupMenu
- ObjectRotateGallery [gallery]
-   ObjectRotationOptionsDialog [button] < ObjectRotateGallery
- ObjectsArrangeMenu [menu]
-   ObjectBringToFront [button] < ObjectsArrangeMenu
-   ObjectBringForward [button] < ObjectsArrangeMenu
-   ObjectSendToBack [button] < ObjectsArrangeMenu
-   ObjectSendBackward [button] < ObjectsArrangeMenu
-   ObjectsGroup [button] < ObjectsArrangeMenu
-   ObjectsUngroup [button] < ObjectsArrangeMenu
-   ObjectsRegroup [button] < ObjectsArrangeMenu
-   ObjectAlignMenu [menu] < ObjectsArrangeMenu
-     ObjectsAlignLeft [button] < ObjectsArrangeMenu
-     ObjectsAlignCenterHorizontal [button] < ObjectsArrangeMenu
-     ObjectsAlignRight [button] < ObjectsArrangeMenu
-     ObjectsAlignTop [button] < ObjectsArrangeMenu
-     ObjectsAlignMiddleVertical [button] < ObjectsArrangeMenu
-     ObjectsAlignBottom [button] < ObjectsArrangeMenu
-     AlignDistributeHorizontallyClassic [button] < ObjectsArrangeMenu
-     AlignDistributeVerticallyClassic [button] < ObjectsArrangeMenu
-     SnapToGrid [toggleButton] < ObjectsArrangeMenu
-     SnapToShapes [toggleButton] < ObjectsArrangeMenu
-     ViewGridlinesToggleExcel [toggleButton] < ObjectsArrangeMenu
-   ObjectRotateGallery [gallery] < ObjectsArrangeMenu
-     ObjectRotationOptionsDialog [button] < ObjectsArrangeMenu

## TabFormulas  (None (Core Tab))
### 
- GroupFunctionLibrary [group]
- GroupPythonChunk [group]
- GroupPythonChunkCommercialPreview [group]
- GroupPythonChunkLFT [group]
- GroupPythonChunkConsumerPreview [group]
- GroupNamedCells [group]
- GroupFormulaAuditing [group]
- GroupCalculation [group]
### GroupFunctionLibrary
- FunctionWizard [button]
- AutoSumMenu [splitButton]
-   AutoSum [button] < AutoSumMenu
-   AutoSumAverage [button] < AutoSumMenu
-   AutoSumCount [button] < AutoSumMenu
-   AutoSumMax [button] < AutoSumMenu
-   AutoSumMin [button] < AutoSumMenu
-   AutoSumMoreFunctions [button] < AutoSumMenu
- FunctionsRecentlyUsedtInsertGallery [gallery]
-   FunctionWizard [button] < FunctionsRecentlyUsedtInsertGallery
- FunctionsFinancialInsertGallery [gallery]
-   FunctionWizard [button] < FunctionsFinancialInsertGallery
- FunctionsLogicalInsertGallery [gallery]
-   FunctionWizard [button] < FunctionsLogicalInsertGallery
- FunctionsTextInsertGallery [gallery]
-   FunctionWizard [button] < FunctionsTextInsertGallery
- FunctionsDateTimeInsertGallery [gallery]
-   FunctionWizard [button] < FunctionsDateTimeInsertGallery
- FunctionsLookupReferenceInsertGallery [gallery]
-   FunctionWizard [button] < FunctionsLookupReferenceInsertGallery
- FunctionsMathTrigInsertGallery [gallery]
-   FunctionWizard [button] < FunctionsMathTrigInsertGallery
- FormulaMoreFunctionsMenu [menu]
-   FunctionsStatisticalInsertGallery [gallery] < FormulaMoreFunctionsMenu
-     FunctionWizard [button] < FormulaMoreFunctionsMenu
-   FunctionsEngineeringInsertGallery [gallery] < FormulaMoreFunctionsMenu
-     FunctionWizard [button] < FormulaMoreFunctionsMenu
-   FunctionsCubeInsertGallery [gallery] < FormulaMoreFunctionsMenu
-     FunctionWizard [button] < FormulaMoreFunctionsMenu
-   FunctionsInformationInsertGallery [gallery] < FormulaMoreFunctionsMenu
-     FunctionWizard [button] < FormulaMoreFunctionsMenu
-   FunctionsCompatibility [gallery] < FormulaMoreFunctionsMenu
-     FunctionWizard [button] < FormulaMoreFunctionsMenu
-   FunctionsWebInsertGallery [gallery] < FormulaMoreFunctionsMenu
-     FunctionWizard [button] < FormulaMoreFunctionsMenu
-   FunctionsAzureInsertGallery [gallery] < FormulaMoreFunctionsMenu
-     FunctionWizard [button] < FormulaMoreFunctionsMenu
### GroupPythonChunk
- InsertPythonMenu [splitButton]
-   InsertPython [button] < InsertPythonMenu
-   PythonInExcelOption [button] < InsertPythonMenu
-   TryPythonSamples [button] < InsertPythonMenu
- PythonResetMenu [splitButton]
-   PythonClearGlobals [button] < PythonResetMenu
-   PythonResetRuntime [button] < PythonResetMenu
- ShowExternalCodeDiagnostics [button]
- InitPyEditor [button]
### GroupPythonChunkCommercialPreview
- InsertPythonMenu [splitButton]
-   InsertPython [button] < InsertPythonMenu
-   PythonInExcelOption [button] < InsertPythonMenu
-   TryPythonSamples [button] < InsertPythonMenu
- PythonResetMenu [splitButton]
-   PythonClearGlobals [button] < PythonResetMenu
-   PythonResetRuntime [button] < PythonResetMenu
- ShowExternalCodeDiagnostics [button]
- InitPyEditor [button]
### GroupPythonChunkLFT
- InsertPythonMenu [splitButton]
-   InsertPython [button] < InsertPythonMenu
-   PythonInExcelOption [button] < InsertPythonMenu
-   TryPythonSamples [button] < InsertPythonMenu
- PythonResetMenu [splitButton]
-   PythonClearGlobals [button] < PythonResetMenu
-   PythonResetRuntime [button] < PythonResetMenu
- ShowExternalCodeDiagnostics [button]
- InitPyEditor [button]
### GroupPythonChunkConsumerPreview
- InsertPythonMenu [splitButton]
-   InsertPython [button] < InsertPythonMenu
-   PythonInExcelOption [button] < InsertPythonMenu
-   TryPythonSamples [button] < InsertPythonMenu
- PythonResetMenu [splitButton]
-   PythonClearGlobals [button] < PythonResetMenu
-   PythonResetRuntime [button] < PythonResetMenu
- ShowExternalCodeDiagnostics [button]
- InitPyEditor [button]
### GroupNamedCells
- NameManager [button]
- NameDefineMenu [splitButton]
-   NameDefine [button] < NameDefineMenu
-   NamesApply [button] < NameDefineMenu
- NameUseInFormula [gallery]
-   NamePasteName [button] < NameUseInFormula
- NameCreateFromSelection [button]
### GroupFormulaAuditing
- TracePrecedents [button]
- TraceDependents [button]
- TraceRemoveArrowsMenu [splitButton]
-   TraceRemoveAllArrows [button] < TraceRemoveArrowsMenu
-   TracePrecedentsRemoveArrows [button] < TraceRemoveArrowsMenu
-   TraceDependentRemoveArrows [button] < TraceRemoveArrowsMenu
- ShowFormulas [toggleButton]
- ErrorCheckingMenu [splitButton]
-   ErrorChecking [button] < ErrorCheckingMenu
-   TraceError [button] < ErrorCheckingMenu
-   CircularReferences [gallery] < ErrorCheckingMenu
- FormulaEvaluate [button]
- WatchWindow [toggleButton]
### GroupCalculation
- CalculationOptionsMenu [menu]
-   CalculationOptionsAutomatically [toggleButton] < CalculationOptionsMenu
-   CalculationOptionsAutomaticallyExceptDataTables [toggleButton] < CalculationOptionsMenu
-   CalculationOptionsManuallly [toggleButton] < CalculationOptionsMenu
-   StaleValueNotification [toggleButton] < CalculationOptionsMenu
- CalculateNow [button]
- CalculateSheet [button]

## TabData  (None (Core Tab))
### 
- GroupGetExternalData [group]
- GroupPowerQueryGetTransformDataSubscription19 [group]
- GroupPowerQueryGetTransformData [group]
- GroupPowerQueryDataSubscription [group]
- GroupPowerQueryData [group]
- GroupConnections [group]
- GroupDataQueriesAndConnections [group]
- GroupDataQueriesAndConnectionsWorkbookLinks [group]
- GroupLinkedEntityConvert [group]
- GroupSortFilter [group]
- GroupDataTools [group]
- GroupForecast [group]
- GroupOutline [group]
### GroupGetExternalData
- GetExternalDataFromAccess [button]
- GetExternalDataFromWeb [button]
- GetExternalDataFromText [button]
- GetExternalDataFromOtherSources [gallery]
- GetExternalDataExistingConnections [button]
### GroupPowerQueryGetTransformDataSubscription19
- PowerQueryGetTransformDataMenuSubscription19 [menu]
-   PowerQueryModernGetData [button] < PowerQueryGetTransformDataMenuSubscription19
-   PowerQueryGetTransformDataFromFileSubscription19 [menu] < PowerQueryGetTransformDataMenuSubscription19
-     GetPowerQueryDataFromExcel [button] < PowerQueryGetTransformDataMenuSubscription19
-     GetPowerQueryGetTransformDataFromText [button] < PowerQueryGetTransformDataMenuSubscription19
-     GetPowerQueryDataFromXml [button] < PowerQueryGetTransformDataMenuSubscription19
-     GetPowerQueryDataFromJson [button] < PowerQueryGetTransformDataMenuSubscription19
-     PowerQueryDataFromPDFFile [button] < PowerQueryGetTransformDataMenuSubscription19
-     GetPowerQueryDataFromFolder [button] < PowerQueryGetTransformDataMenuSubscription19
-     GetPowerQueryDataFromSharePointFolder [button] < PowerQueryGetTransformDataMenuSubscription19
-   PowerQueryGetTransformDataFromDatabaseSubscription19 [menu] < PowerQueryGetTransformDataMenuSubscription19
-     GetPowerQueryDataFromSqlServerDatabase [button] < PowerQueryGetTransformDataMenuSubscription19
-     GetPowerQueryDataFromAccessDatabase [button] < PowerQueryGetTransformDataMenuSubscription19
-     GetPowerQueryGetTransformDataFromAnalysisServices [button] < PowerQueryGetTransformDataMenuSubscription19
-     GetPowerQueryGetTransformDataFromSqlServerASServices [button] < PowerQueryGetTransformDataMenuSubscription19
-     GetPowerQueryDataFromOracleDatabase [button] < PowerQueryGetTransformDataMenuSubscription19
-     GetPowerQueryDataFromIbmDb2Database [button] < PowerQueryGetTransformDataMenuSubscription19
-     GetPowerQueryDataFromMySqlDatabase [button] < PowerQueryGetTransformDataMenuSubscription19
-     GetPowerQueryDataFromPostgreSqlDatabase [button] < PowerQueryGetTransformDataMenuSubscription19
-     GetPowerQueryDataFromSybaseDatabase [button] < PowerQueryGetTransformDataMenuSubscription19
-     GetPowerQueryDataFromTeradataDatabase [button] < PowerQueryGetTransformDataMenuSubscription19
-     GetPowerQueryDataFromSapHanaDatabase [button] < PowerQueryGetTransformDataMenuSubscription19
-     PowerQueryDataFromSnowflake [button] < PowerQueryGetTransformDataMenuSubscription19
-   PowerQueryGetTransformDataFromAzureSubscription19 [menu] < PowerQueryGetTransformDataMenuSubscription19
-     GetPowerQueryDataFromAzureSqlDatabase [button] < PowerQueryGetTransformDataMenuSubscription19
-     GetPowerQueryDataFromAzureDataWarehouse [button] < PowerQueryGetTransformDataMenuSubscription19
-     GetPowerQueryDataFromWindowsAzureHdInsight [button] < PowerQueryGetTransformDataMenuSubscription19
-     GetPowerQueryDataFromWindowsAzureBlob [button] < PowerQueryGetTransformDataMenuSubscription19
-     GetPowerQueryDataFromWindowsAzureTable [button] < PowerQueryGetTransformDataMenuSubscription19
-     GetPowerQueryDataFromAzureDataLakeStore [button] < PowerQueryGetTransformDataMenuSubscription19
-     GetPowerQueryDataFromAzureDataLakeStoreGen2 [button] < PowerQueryGetTransformDataMenuSubscription19
-     PowerQueryDataFromKusto [button] < PowerQueryGetTransformDataMenuSubscription19
-   PowerQueryDataFromPowerPlatform [menu] < PowerQueryGetTransformDataMenuSubscription19
-     PowerBIData [button] < PowerQueryGetTransformDataMenuSubscription19
-     PowerQueryDataFromPowerBIDataFlows [button] < PowerQueryGetTransformDataMenuSubscription19
-     PowerQueryDataFromPowerBIDataVerse [button] < PowerQueryGetTransformDataMenuSubscription19
-   PowerQueryGetTransformDataFromOnlineServicesSubscription19 [menu] < PowerQueryGetTransformDataMenuSubscription19
-     GetPowerQueryDataFromSharePointOnlineList [button] < PowerQueryGetTransformDataMenuSubscription19
-     GetPowerQueryDataFromMicrosoftExchangeOnline [button] < PowerQueryGetTransformDataMenuSubscription19
-     GetPowerQueryDataFromDynamicsCRM [button] < PowerQueryGetTransformDataMenuSubscription19
-     PowerQueryDataFromMicrosoftDynamics365BusinessCentral [button] < PowerQueryGetTransformDataMenuSubscription19
-     GetPowerQueryDataFromSalesforceObjects [button] < PowerQueryGetTransformDataMenuSubscription19
-     GetPowerQueryDataFromSalesforceReports [button] < PowerQueryGetTransformDataMenuSubscription19
-     PowerQueryDataFromAdobeAnalytics [button] < PowerQueryGetTransformDataMenuSubscription19
-   PowerQueryGetTransformDataFromOtherSourceSubscription19 [menu] < PowerQueryGetTransformDataMenuSubscription19
-     GetPowerQueryGetTransformDataFromTable [button] < PowerQueryGetTransformDataMenuSubscription19
-     GetPowerQueryDataFromWeb [button] < PowerQueryGetTransformDataMenuSubscription19
-     GetPowerQueryGetTransformDataFromMicrosoftQuery [button] < PowerQueryGetTransformDataMenuSubscription19
-     GetPowerQueryDataFromSharePointList [button] < PowerQueryGetTransformDataMenuSubscription19
-     GetPowerQueryDataFromODataFeed [button] < PowerQueryGetTransformDataMenuSubscription19
-     GetPowerQueryDataFromHdfsFile [button] < PowerQueryGetTransformDataMenuSubscription19
-     GetPowerQueryDataFromActiveDirectory [button] < PowerQueryGetTransformDataMenuSubscription19
-     GetPowerQueryDataFromMicrosoftExchange [button] < PowerQueryGetTransformDataMenuSubscription19
-     GetPowerQueryDataFromODBC [button] < PowerQueryGetTransformDataMenuSubscription19
-     GetPowerQueryDataFromOLEDB [button] < PowerQueryGetTransformDataMenuSubscription19
-     DataFromPicture [menu] < PowerQueryGetTransformDataMenuSubscription19
-     DataFromScreenshot [button] < PowerQueryGetTransformDataMenuSubscription19
-     DataFromClipboardImage [button] < PowerQueryGetTransformDataMenuSubscription19
-     GetPowerQueryDataFromEmptyFormula [button] < PowerQueryGetTransformDataMenuSubscription19
-   GetPowerQueryGetTransformDataLegacyWizard [gallery] < PowerQueryGetTransformDataMenuSubscription19
-   PowerQueryGetTransformDataCombineSubscription19 [menu] < PowerQueryGetTransformDataMenuSubscription19
-     GetPowerQueryMerge [button] < PowerQueryGetTransformDataMenuSubscription19
-     GetPowerQueryAppend [button] < PowerQueryGetTransformDataMenuSubscription19
-   PowerQueryGetTransformDataLaunchQueryEditor [button] < PowerQueryGetTransformDataMenuSubscription19
-   GetPowerQueryDataSourceSettings [button] < PowerQueryGetTransformDataMenuSubscription19
-   PowerQueryShowQuerySettings [button] < PowerQueryGetTransformDataMenuSubscription19
- GetPowerQueryGetTransformDataFromText [button]
- GetPowerQueryDataFromWeb [button]
- GetPowerQueryGetTransformDataFromTable [button]
- DataFromPicture [menu]
-   DataFromScreenshot [button] < DataFromPicture
-   DataFromClipboardImage [button] < DataFromPicture
- PowerQueryRecentSources [button]
- GetTransformExistingConnections [button]
### GroupPowerQueryGetTransformData
- PowerQueryGetTransformDataMenu [menu]
-   GetPowerQueryGetTransformDataFromFile [menu] < PowerQueryGetTransformDataMenu
-     GetPowerQueryDataFromExcel [button] < PowerQueryGetTransformDataMenu
-     GetPowerQueryGetTransformDataFromText [button] < PowerQueryGetTransformDataMenu
-     GetPowerQueryDataFromXml [button] < PowerQueryGetTransformDataMenu
-     GetPowerQueryDataFromJson [button] < PowerQueryGetTransformDataMenu
-     GetPowerQueryDataFromFolder [button] < PowerQueryGetTransformDataMenu
-     GetPowerQueryDataFromSharePointFolder [button] < PowerQueryGetTransformDataMenu
-   GetPowerQueryGetTransformDataFromDatabaseSubscription [menu] < PowerQueryGetTransformDataMenu
-     GetPowerQueryDataFromSqlServerDatabase [button] < PowerQueryGetTransformDataMenu
-     GetPowerQueryDataFromAccessDatabase [button] < PowerQueryGetTransformDataMenu
-     GetPowerQueryGetTransformDataFromAnalysisServices [button] < PowerQueryGetTransformDataMenu
-     GetPowerQueryGetTransformDataFromSqlServerASServices [button] < PowerQueryGetTransformDataMenu
-     GetPowerQueryDataFromOracleDatabase [button] < PowerQueryGetTransformDataMenu
-     GetPowerQueryDataFromIbmDb2Database [button] < PowerQueryGetTransformDataMenu
-     GetPowerQueryDataFromMySqlDatabase [button] < PowerQueryGetTransformDataMenu
-     GetPowerQueryDataFromPostgreSqlDatabase [button] < PowerQueryGetTransformDataMenu
-     GetPowerQueryDataFromSybaseDatabase [button] < PowerQueryGetTransformDataMenu
-     GetPowerQueryDataFromTeradataDatabase [button] < PowerQueryGetTransformDataMenu
-     GetPowerQueryDataFromSapHanaDatabase [button] < PowerQueryGetTransformDataMenu
-     PowerQueryDataFromSnowflake [button] < PowerQueryGetTransformDataMenu
-   PowerQueryGetTransformDataFromAzure [menu] < PowerQueryGetTransformDataMenu
-     GetPowerQueryDataFromAzureSqlDatabase [button] < PowerQueryGetTransformDataMenu
-     GetPowerQueryDataFromAzureDataWarehouse [button] < PowerQueryGetTransformDataMenu
-     GetPowerQueryDataFromWindowsAzureHdInsight [button] < PowerQueryGetTransformDataMenu
-     GetPowerQueryDataFromWindowsAzureBlob [button] < PowerQueryGetTransformDataMenu
-     GetPowerQueryDataFromWindowsAzureTable [button] < PowerQueryGetTransformDataMenu
-     GetPowerQueryDataFromAzureDataLakeStore [button] < PowerQueryGetTransformDataMenu
-   PowerQueryGetTransformDataFromOnlineServices [menu] < PowerQueryGetTransformDataMenu
-     GetPowerQueryDataFromSharePointOnlineList [button] < PowerQueryGetTransformDataMenu
-     GetPowerQueryDataFromMicrosoftExchangeOnline [button] < PowerQueryGetTransformDataMenu
-     GetPowerQueryDataFromDynamicsCRM [button] < PowerQueryGetTransformDataMenu
-     PowerQueryDataFromMicrosoftDynamics365BusinessCentral [button] < PowerQueryGetTransformDataMenu
-     GetPowerQueryDataFromSalesforceObjects [button] < PowerQueryGetTransformDataMenu
-     GetPowerQueryDataFromSalesforceReports [button] < PowerQueryGetTransformDataMenu
-     PowerQueryDataFromAdobeAnalytics [button] < PowerQueryGetTransformDataMenu
-   GetPowerQueryGetTransformDataFromOtherSource [menu] < PowerQueryGetTransformDataMenu
-     GetPowerQueryGetTransformDataFromTable [button] < PowerQueryGetTransformDataMenu
-     GetPowerQueryDataFromWeb [button] < PowerQueryGetTransformDataMenu
-     GetPowerQueryGetTransformDataFromMicrosoftQuery [button] < PowerQueryGetTransformDataMenu
-     GetPowerQueryDataFromSharePointList [button] < PowerQueryGetTransformDataMenu
-     GetPowerQueryDataFromODataFeed [button] < PowerQueryGetTransformDataMenu
-     GetPowerQueryDataFromHdfsFile [button] < PowerQueryGetTransformDataMenu
-     GetPowerQueryDataFromActiveDirectory [button] < PowerQueryGetTransformDataMenu
-     GetPowerQueryDataFromMicrosoftExchange [button] < PowerQueryGetTransformDataMenu
-     GetPowerQueryDataFromODBC [button] < PowerQueryGetTransformDataMenu
-     GetPowerQueryDataFromOLEDB [button] < PowerQueryGetTransformDataMenu
-     GetPowerQueryDataFromEmptyFormula [button] < PowerQueryGetTransformDataMenu
-   GetPowerQueryGetTransformDataLegacyWizard [gallery] < PowerQueryGetTransformDataMenu
-   PowerQueryGetTransformDataCombine [menu] < PowerQueryGetTransformDataMenu
-     GetPowerQueryMerge [button] < PowerQueryGetTransformDataMenu
-     GetPowerQueryAppend [button] < PowerQueryGetTransformDataMenu
-   PowerQueryGetTransformDataLaunchQueryEditor [button] < PowerQueryGetTransformDataMenu
-   GetPowerQueryDataSourceSettings [button] < PowerQueryGetTransformDataMenu
-   PowerQueryShowQuerySettings [button] < PowerQueryGetTransformDataMenu
- GetPowerQueryGetTransformDataFromText [button]
- GetPowerQueryDataFromWeb [button]
- GetPowerQueryGetTransformDataFromTable [button]
- PowerQueryRecentSources [button]
- GetTransformExistingConnections [button]
### GroupPowerQueryDataSubscription
- FlyoutAnchorPowerQueryNewQuerySubscription [menu]
-   GetPowerQueryDataFromFileSubscription [menu] < FlyoutAnchorPowerQueryNewQuerySubscription
-     GetPowerQueryDataFromExcel [button] < FlyoutAnchorPowerQueryNewQuerySubscription
-     GetPowerQueryDataFromCsv [button] < FlyoutAnchorPowerQueryNewQuerySubscription
-     GetPowerQueryDataFromXml [button] < FlyoutAnchorPowerQueryNewQuerySubscription
-     GetPowerQueryDataFromText [button] < FlyoutAnchorPowerQueryNewQuerySubscription
-     GetPowerQueryDataFromJson [button] < FlyoutAnchorPowerQueryNewQuerySubscription
-     GetPowerQueryDataFromFolder [button] < FlyoutAnchorPowerQueryNewQuerySubscription
-     GetPowerQueryDataFromSharePointFolder [button] < FlyoutAnchorPowerQueryNewQuerySubscription
-   GetPowerQueryDataFromDatabaseSubscription [menu] < FlyoutAnchorPowerQueryNewQuerySubscription
-     GetPowerQueryDataFromSqlServerDatabase [button] < FlyoutAnchorPowerQueryNewQuerySubscription
-     GetPowerQueryDataFromAccessDatabase [button] < FlyoutAnchorPowerQueryNewQuerySubscription
-     GetPowerQueryDataFromSqlServerASServices [button] < FlyoutAnchorPowerQueryNewQuerySubscription
-     GetPowerQueryDataFromOracleDatabase [button] < FlyoutAnchorPowerQueryNewQuerySubscription
-     GetPowerQueryDataFromIbmDb2Database [button] < FlyoutAnchorPowerQueryNewQuerySubscription
-     GetPowerQueryDataFromMySqlDatabase [button] < FlyoutAnchorPowerQueryNewQuerySubscription
-     GetPowerQueryDataFromPostgreSqlDatabase [button] < FlyoutAnchorPowerQueryNewQuerySubscription
-     GetPowerQueryDataFromSybaseDatabase [button] < FlyoutAnchorPowerQueryNewQuerySubscription
-     GetPowerQueryDataFromTeradataDatabase [button] < FlyoutAnchorPowerQueryNewQuerySubscription
-     GetPowerQueryDataFromSapHanaDatabase [button] < FlyoutAnchorPowerQueryNewQuerySubscription
-   GetPowerQueryDataFromAzureSubscription [menu] < FlyoutAnchorPowerQueryNewQuerySubscription
-     GetPowerQueryDataFromAzureSqlDatabase [button] < FlyoutAnchorPowerQueryNewQuerySubscription
-     GetPowerQueryDataFromAzureDataWarehouse [button] < FlyoutAnchorPowerQueryNewQuerySubscription
-     GetPowerQueryDataFromWindowsAzureHdInsight [button] < FlyoutAnchorPowerQueryNewQuerySubscription
-     GetPowerQueryDataFromWindowsAzureBlob [button] < FlyoutAnchorPowerQueryNewQuerySubscription
-     GetPowerQueryDataFromWindowsAzureTable [button] < FlyoutAnchorPowerQueryNewQuerySubscription
-     GetPowerQueryDataFromAzureDataLakeStore [button] < FlyoutAnchorPowerQueryNewQuerySubscription
-   GetPowerQueryDataFromOnlineServicesSubscription [menu] < FlyoutAnchorPowerQueryNewQuerySubscription
-     GetPowerQueryDataFromSharePointOnlineList [button] < FlyoutAnchorPowerQueryNewQuerySubscription
-     GetPowerQueryDataFromMicrosoftExchangeOnline [button] < FlyoutAnchorPowerQueryNewQuerySubscription
-     GetPowerQueryDataFromDynamicsCRM [button] < FlyoutAnchorPowerQueryNewQuerySubscription
-     PowerQueryDataFromMicrosoftDynamics365BusinessCentral [button] < FlyoutAnchorPowerQueryNewQuerySubscription
-     GetPowerQueryDataFromSalesforceObjects [button] < FlyoutAnchorPowerQueryNewQuerySubscription
-     GetPowerQueryDataFromSalesforceReports [button] < FlyoutAnchorPowerQueryNewQuerySubscription
-   GetPowerQueryDataFromOtherSourceSubscription [menu] < FlyoutAnchorPowerQueryNewQuerySubscription
-     GetPowerQueryDataFromWeb [button] < FlyoutAnchorPowerQueryNewQuerySubscription
-     GetPowerQueryDataFromSharePointList [button] < FlyoutAnchorPowerQueryNewQuerySubscription
-     GetPowerQueryDataFromODataFeed [button] < FlyoutAnchorPowerQueryNewQuerySubscription
-     GetPowerQueryDataFromHdfsFile [button] < FlyoutAnchorPowerQueryNewQuerySubscription
-     GetPowerQueryDataFromActiveDirectory [button] < FlyoutAnchorPowerQueryNewQuerySubscription
-     GetPowerQueryDataFromMicrosoftExchange [button] < FlyoutAnchorPowerQueryNewQuerySubscription
-     GetPowerQueryDataFromODBC [button] < FlyoutAnchorPowerQueryNewQuerySubscription
-     GetPowerQueryDataFromOLEDB [button] < FlyoutAnchorPowerQueryNewQuerySubscription
-     GetPowerQueryDataFromEmptyFormula [button] < FlyoutAnchorPowerQueryNewQuerySubscription
-   FlyoutAnchorPowerQueryCombineSubscription [menu] < FlyoutAnchorPowerQueryNewQuerySubscription
-     GetPowerQueryMerge [button] < FlyoutAnchorPowerQueryNewQuerySubscription
-     GetPowerQueryAppend [button] < FlyoutAnchorPowerQueryNewQuerySubscription
-     PowerQueryLaunchQueryEditor [button] < FlyoutAnchorPowerQueryNewQuerySubscription
-   GetPowerQueryDataSourceSettings [button] < FlyoutAnchorPowerQueryNewQuerySubscription
-   PowerQueryShowQuerySettings [button] < FlyoutAnchorPowerQueryNewQuerySubscription
- PowerQueryManageWorkbookQueries [toggleButton]
- GetPowerQueryExcelDataFromTable [button]
- PowerQueryRecentSources [button]
### GroupPowerQueryData
- FlyoutAnchorPowerQueryNewQuery [menu]
-   GetPowerQueryDataFromFile [menu] < FlyoutAnchorPowerQueryNewQuery
-     GetPowerQueryDataFromExcel [button] < FlyoutAnchorPowerQueryNewQuery
-     GetPowerQueryDataFromCsv [button] < FlyoutAnchorPowerQueryNewQuery
-     GetPowerQueryDataFromXml [button] < FlyoutAnchorPowerQueryNewQuery
-     GetPowerQueryDataFromText [button] < FlyoutAnchorPowerQueryNewQuery
-     GetPowerQueryDataFromFolder [button] < FlyoutAnchorPowerQueryNewQuery
-   GetPowerQueryDataFromDatabase [menu] < FlyoutAnchorPowerQueryNewQuery
-     GetPowerQueryDataFromSqlServerDatabase [button] < FlyoutAnchorPowerQueryNewQuery
-     GetPowerQueryDataFromAccessDatabase [button] < FlyoutAnchorPowerQueryNewQuery
-     GetPowerQueryDataFromSqlServerASServices [button] < FlyoutAnchorPowerQueryNewQuery
-     GetPowerQueryDataFromOracleDatabase [button] < FlyoutAnchorPowerQueryNewQuery
-     GetPowerQueryDataFromIbmDb2Database [button] < FlyoutAnchorPowerQueryNewQuery
-     GetPowerQueryDataFromMySqlDatabase [button] < FlyoutAnchorPowerQueryNewQuery
-     GetPowerQueryDataFromPostgreSqlDatabase [button] < FlyoutAnchorPowerQueryNewQuery
-     GetPowerQueryDataFromSybaseDatabase [button] < FlyoutAnchorPowerQueryNewQuery
-     GetPowerQueryDataFromTeradataDatabase [button] < FlyoutAnchorPowerQueryNewQuery
-   GetPowerQueryDataFromAzure [menu] < FlyoutAnchorPowerQueryNewQuery
-     GetPowerQueryDataFromAzureSqlDatabase [button] < FlyoutAnchorPowerQueryNewQuery
-     GetPowerQueryDataFromWindowsAzureHdInsight [button] < FlyoutAnchorPowerQueryNewQuery
-     GetPowerQueryDataFromWindowsAzureBlob [button] < FlyoutAnchorPowerQueryNewQuery
-     GetPowerQueryDataFromWindowsAzureTable [button] < FlyoutAnchorPowerQueryNewQuery
-   GetPowerQueryDataFromOnlineServices [menu] < FlyoutAnchorPowerQueryNewQuery
-     GetPowerQueryDataFromSharePointOnlineList [button] < FlyoutAnchorPowerQueryNewQuery
-     GetPowerQueryDataFromMicrosoftExchangeOnline [button] < FlyoutAnchorPowerQueryNewQuery
-     GetPowerQueryDataFromDynamicsCRM [button] < FlyoutAnchorPowerQueryNewQuery
-     GetPowerQueryDataFromSalesforceObjects [button] < FlyoutAnchorPowerQueryNewQuery
-     GetPowerQueryDataFromSalesforceReports [button] < FlyoutAnchorPowerQueryNewQuery
-   GetPowerQueryDataFromOtherSource [menu] < FlyoutAnchorPowerQueryNewQuery
-     GetPowerQueryDataFromWeb [button] < FlyoutAnchorPowerQueryNewQuery
-     GetPowerQueryDataFromSharePointList [button] < FlyoutAnchorPowerQueryNewQuery
-     GetPowerQueryDataFromODataFeed [button] < FlyoutAnchorPowerQueryNewQuery
-     GetPowerQueryDataFromHdfsFile [button] < FlyoutAnchorPowerQueryNewQuery
-     GetPowerQueryDataFromActiveDirectory [button] < FlyoutAnchorPowerQueryNewQuery
-     GetPowerQueryDataFromMicrosoftExchange [button] < FlyoutAnchorPowerQueryNewQuery
-     GetPowerQueryDataFromODBC [button] < FlyoutAnchorPowerQueryNewQuery
-     GetPowerQueryDataFromEmptyFormula [button] < FlyoutAnchorPowerQueryNewQuery
-   FlyoutAnchorPowerQueryCombine [menu] < FlyoutAnchorPowerQueryNewQuery
-     GetPowerQueryMerge [button] < FlyoutAnchorPowerQueryNewQuery
-     GetPowerQueryAppend [button] < FlyoutAnchorPowerQueryNewQuery
-     PowerQueryLaunchQueryEditor [button] < FlyoutAnchorPowerQueryNewQuery
-   GetPowerQueryDataSourceSettings [button] < FlyoutAnchorPowerQueryNewQuery
-   PowerQueryShowQuerySettings [button] < FlyoutAnchorPowerQueryNewQuery
- PowerQueryManageWorkbookQueries [toggleButton]
- GetPowerQueryExcelDataFromTable [button]
- PowerQueryRecentSources [button]
### GroupConnections
- RefreshAllMenu [splitButton]
-   RefreshAll [button] < RefreshAllMenu
-   Refresh [button] < RefreshAllMenu
-   RefreshStatus [button] < RefreshAllMenu
-   RefreshCancel [button] < RefreshAllMenu
-   ConnectionProperties [button] < RefreshAllMenu
- Connections [button]
- DataRangeProperties [button]
- EditLinks [button]
### GroupDataQueriesAndConnections
- RefreshAllMenu [splitButton]
-   RefreshAll [button] < RefreshAllMenu
-   Refresh [button] < RefreshAllMenu
-   RefreshStatus [button] < RefreshAllMenu
-   RefreshCancel [button] < RefreshAllMenu
-   ConnectionProperties [button] < RefreshAllMenu
- ShowQueriesAndConnections [toggleButton]
- DataRangeProperties [button]
- EditLinks [button]
### GroupDataQueriesAndConnectionsWorkbookLinks
- RefreshAllMenu [splitButton]
-   RefreshAll [button] < RefreshAllMenu
-   Refresh [button] < RefreshAllMenu
-   RefreshStatus [button] < RefreshAllMenu
-   RefreshCancel [button] < RefreshAllMenu
-   ConnectionProperties [button] < RefreshAllMenu
- ShowQueriesAndConnections [toggleButton]
- DataRangeProperties [button]
- WorkbookLinks [button]
### GroupLinkedEntityConvert
- ConvertToLinkedEntity [gallery]
-   LinkedEntityMoreDataBtn [button] < ConvertToLinkedEntity
### GroupSortFilter
- SortAscendingExcel [button]
- SortDescendingExcel [button]
- SortDialog [button]
- Filter [toggleButton]
- SortClear [button]
- FilterReapply [button]
- AdvancedFilterDialog [button]
### GroupDataTools
- ConvertTextToTable [button]
- QuickCode [button]
- RemoveDuplicates [button]
- DataValidationMenu [splitButton]
-   DataValidation [button] < DataValidationMenu
-   DataValidationCircleInvalid [button] < DataValidationMenu
-   DataValidationClearValidationCircles [button] < DataValidationMenu
- Consolidate [button]
- AsppRelationships [button]
- DataModelManage [button]
### GroupForecast
- WhatIfAnalysisMenu [menu]
-   ScenarioManager [button] < WhatIfAnalysisMenu
-   GoalSeek [button] < WhatIfAnalysisMenu
-   DataTable [button] < WhatIfAnalysisMenu
- ForecastInsert [button]
### GroupOutline
- OutlineGroupMenu [splitButton]
-   OutlineGroup [button] < OutlineGroupMenu
-   OutlineAuto [button] < OutlineGroupMenu
- OutlineUngroupMenu [splitButton]
-   OutlineUngroup [button] < OutlineUngroupMenu
-   OutlineClear [button] < OutlineUngroupMenu
- OutlineSubtotals [button]
- OutlineShowDetail [button]
- OutlineHideDetail [button]
- OutlineSettings [button (dialogBoxLauncher)]

## TabReview  (None (Core Tab))
### 
- GroupProofing [group]
- GroupPerformance [group]
- GroupAccessibility [group]
- GroupInsights [group]
- GroupLanguage [group]
- GroupLineage [group]
- GroupComments [group]
- GroupThreadedComments [group]
- GroupCommentsLegacy [group]
- GroupChangesExcel [group]
- GroupProtectExcel [group]
- GroupInk [group]
- GroupDebug [group]
### GroupProofing
- Spelling [button]
- Thesaurus [button]
- StatusWorkbookStatistics [button]
### GroupPerformance
- CheckPerformance [button]
### GroupAccessibility
- MenuAccessibility [splitButton]
-   AccessibilityChecker [button] < MenuAccessibility
-   AltTextPaneRibbon [toggleButton] < MenuAccessibility
-   FormatAsTableGallery [gallery] < MenuAccessibility
-     TableStyleNew [button] < MenuAccessibility
-     PivotTableNewStyle [button] < MenuAccessibility
-   UnmergeCells [button] < MenuAccessibility
-   OpenOptionsEaseOfAcess [button] < MenuAccessibility
-   AccessibilityHelp [button] < MenuAccessibility
### GroupInsights
- Insights [button]
### GroupLanguage
- Translate [button]
- Translator [button]
- HangulHanjaConversion [button]
### GroupLineage
- LineageShowChanges [toggleButton]
### GroupComments
- ReviewNewComment [button]
- ReviewDeleteComment [button]
- ReviewPreviousComment [button]
- ReviewNextComment [button]
- ReviewShowOrHideComment [button]
- ReviewShowAllComments [toggleButton]
### GroupThreadedComments
- NewThreadedComment [button]
- ReviewDeleteComment [button]
- PreviousThreadedComment [button]
- NextThreadedComment [button]
- ShowThreadedComments [toggleButton]
### GroupCommentsLegacy
- ReviewNoteSplit [menu]
-   NewCommentLegacy [button] < ReviewNoteSplit
-   PreviousCommentLegacy [button] < ReviewNoteSplit
-   NextCommentLegacy [button] < ReviewNoteSplit
-   ShowCommentLegacy [button] < ReviewNoteSplit
-   ShowAllCommentsLegacy [toggleButton] < ReviewNoteSplit
-   ConvertComments [button] < ReviewNoteSplit
### GroupChangesExcel
- SheetProtect [button]
- ReviewRestrictEditing [toggleButton]
- ReviewShareWorkbook [button]
- ReviewProtectAndShareWorkbook [button]
- ReviewAllowUsersToEditRanges [button]
- ReviewTrackChangesMenu [menu]
-   ReviewHighlightChanges [button] < ReviewTrackChangesMenu
-   ReviewAcceptOrRejectChangeDialog [button] < ReviewTrackChangesMenu
### GroupProtectExcel
- SheetProtect [button]
- ReviewRestrictEditing [toggleButton]
- ReviewAllowUsersToEditRangesCoauth [button]
- SameAsName [button]
### GroupInk
- InkingStart [button]
- HideInkSplitButton [splitButton]
-   HideInk [toggleButton] < HideInkSplitButton
-   InkDeleteAllOnSheet [button] < HideInkSplitButton
-   InkDeleteAll [button] < HideInkSplitButton
### GroupDebug
- CollaborateChanges [button]

## TabView  (None (Core Tab))
### 
- GroupNamedSheetView [group]
- GroupWorkbookViews [group]
- GroupNightMode [group]
- GroupViewShowHide [group]
- GroupZoom [group]
- GroupWindow [group]
- GroupMacros [group]
- GroupViewDebug [group]
### GroupNamedSheetView
- SheetViewComboBox [comboBox]
- KeepTemporarySheetView [button]
- ExitSheetView [button]
- NewSheetView [button]
- SheetViewOptions [button]
### GroupWorkbookViews
- ViewNormalViewExcel [toggleButton]
- ViewPageBreakPreviewView [toggleButton]
- ViewPageLayoutView [toggleButton]
- ViewCustomViews [button]
### GroupNightMode
- DarkModeOn [toggleButton]
### GroupViewShowHide
- ViewDocumentActionsPane [toggleButton]
- NavigationPaneExcelFeature [toggleButton]
- FocusActiveCell [toggleButton]
- ViewRulerExcel [checkBox]
- GridlinesExcel [checkBox]
- ViewFormulaBar [checkBox]
- ViewHeadings [checkBox]
- DataTypeIcons [checkBox]
### GroupZoom
- ZoomDialog [button]
- ZoomCurrent100 [button]
- ZoomToSelection [toggleButton]
### GroupWindow
- WindowNew [button]
- WindowsArrangeAll [button]
- ViewFreezePanesGallery [gallery]
- WindowSplitToggle [toggleButton]
- WindowHide [button]
- WindowUnhide [button]
- ViewSideBySide [toggleButton]
- WindowSideBySideSynchronousScrolling [toggleButton]
- WindowResetPosition [button]
- WindowSwitchWindowsMenuExcel [menu]
### GroupMacros
- MenuMacros [splitButton]
-   PlayMacro [button] < MenuMacros
-   MacroRecord [button] < MenuMacros
-   MacroRelativeReferences [toggleButton] < MenuMacros
### GroupViewDebug
- DisableGdiHWAcceleration [toggleButton]

## TabAutomate  (None (Core Tab))
### 
- GroupScriptingTools [group]
- GroupOfficeScripts [group]
- GroupPowerAutomate [group]
### GroupScriptingTools
- RecordActions [button]
- NewScript [button]
### GroupOfficeScripts
- OfficeScriptsGallery [gallery]
-   SharedScripts [button] < OfficeScriptsGallery
-   AllScripts [button] < OfficeScriptsGallery
- AllScripts [button]
### GroupPowerAutomate
- PowerAutomate [button]

## TabDeveloper  (None (Core Tab))
### 
- GroupCode [group]
- GroupAddins [group]
- GroupControls [group]
- GroupXml [group]
### GroupCode
- VisualBasic [button]
- MacroPlay [button]
- MacroRecord [button]
- MacroRelativeReferences [toggleButton]
- MacroSecurity [button]
### GroupAddins
- OfficeExtensionsGallery3 [button]
- AddInManager [button]
- ComAddInsDialog [button]
### GroupControls
- ControlsGallery [gallery]
- DesignMode [toggleButton]
- ControlProperties [button]
- ViewCode [button]
- RunDialog [button]
### GroupXml
- XmlSource [toggleButton]
- XmlMapProperties [button]
- XmlExpansionPacksExcel [button]
- XmlDataRefresh [button]
- XmlImport [button]
- XmlExport [button]

## TabAddIns  (None (Core Tab))
### 
- GroupAddInsMenuCommands [group]
- GroupAddInsToolbarCommands [group]
- GroupAddInsCustomToolbars [group]

## HelpTab  (None (Core Tab))
### 
- GroupHelpAndSupport [group]
- GroupExcelCommunity [group]
### GroupHelpAndSupport
- Help [button]
- ContactUs [button]
- OfficeFeedbackHelpTab [button]
- OfficeFeedbackIcon [button]
- ShowTraining [button]
- WhatsNewRecentUpdates [button]
### GroupExcelCommunity
- AskCommunity [button]
- BlogSite [button]

## TabPrintPreview  (None (Core Tab))
### 
- GroupPrintPreviewPrint [group]
- GroupPrintPreviewZoom [group]
- GroupPrintPreviewPreview [group]
### GroupPrintPreviewPrint
- FilePrint [button]
- PageSetupPageDialog [button]
### GroupPrintPreviewZoom
- ZoomPrintPreviewExcel [toggleButton]
### GroupPrintPreviewPreview
- PageNext [button]
- PagePrevious [button]
- MarginsShowHide [checkBox]
- PrintPreviewClose [button]

## TabBackgroundRemoval  (None (Core Tab))
### 
- GroupBackgroundRemovalMode [group]
- GroupBackgroundRemovalClose [group]
### GroupBackgroundRemovalMode
- PictureBackgroundRemovalMarkForeground [toggleButton]
- PictureBackgroundRemovalMarkBackground [toggleButton]
- PictureBackgroundRemovalMarkDelete [toggleButton]
### GroupBackgroundRemovalClose
- BackgroundRemovalClose [button]
- KeepBackgroundRemoval [button]

## None (Not in the Ribbon)  (None (Not in the Ribbon))
### 
- FileNewDefault [button]
- FileOpenUsingBackstage [button]
- FileSave [button]
- FileSendAsAttachment [button]
- FilePrintQuick [button]
- PrintPreviewAndPrint [button]
- Spelling [button]
- Superscript [toggleButton]
- Subscript [toggleButton]
- Undo [gallery]
- Redo [gallery]
- SortAscendingExcel [button]
- SortDescendingExcel [button]
- FileOpenRecentFile [button]
- PointerModeOptions [gallery]
- FilePermissionRestrictAs [button]
- EditLinks [button]
- MdiChildSystemMenu [menu]
- MinimizeRibbon [toggleButton]
- WindowMinimize [button]
- WindowRestore [button]
- WindowClose [button]
- HideRibbon [toggleButton]
- PowerViewInsert [button]
- FileCreateDocumentWorkspace [toggleButton]
- FileDocumentManagementInformation [toggleButton]
- ViewMessageBar [checkBox]
- SplitCells [button]
- SortDialogClassic [button]
- ViewFullScreenView [toggleButton]
- OpenOptionsAddIns [button]
- OpenOptionsAdvanced [button]
- OpenOptionsCustomization [button]
- OpenOptionsExcelCustomizeRibbon [button]
- OpenOptionsData [button]
- OpenOptionsEaseOfAcess [button]
- OpenOptionsExperimentSettings [button]
- OpenOptionsFormulas [button]
- OpenOptionsLanguage [button]
- OpenOptionsSaving [button]
- OpenOptionsSetup [button]
- OpenOptionsSpelling [button]
- OpenOptionsTrustCenterSettings [button]
- OpenOptionsExcelTrustCenter [button]
- FindTellMe [control]
- SmartLookup [control]
- ZoomClassic [gallery]
- FreezePanes [button]
- FreezeTopRow [button]
- FreezeFirstColumn [button]
- PivotTableReport [button]
- ShapeRectangle [toggleButton]
- PageBreakInsertOrRemove [button]
- EqualSign [button]
- SendCopySendToMailRecipient [toggleButton]
- Repeat [button]
- AutoFilterClassic [button]
- AutoFormatDialog [button]
- GetExternalDataImportClassic [button]
- ShapeOval [toggleButton]
- BrightnessMore [button]
- ContrastMore [button]
- LightShading [button]
- StylesDialogClassic [button]
- EditQuery [button]
- FilePermission [button]
- PropertySheet [button]
- ContrastLess [button]
- Strikethrough [toggleButton]
- BrightnessLess [button]
- FileSaveAsWebPage [button]
- AlignJustify [toggleButton]
- DataFormExcel [button]
- TableSelectVisibleCells [button]
- BorderBottom [toggleButton]
- ObjectRotateLeft90 [button]
- Camera [button]
- ShapeElbowConnectorArrow [toggleButton]
- ShapeStraightConnectorArrow [toggleButton]
- ObjectSetShapeDefaults [button]
- ShapeFreeform [toggleButton]
- ShapeConvertToFreeform [button]
- ShapeCurve [toggleButton]
- ParagraphDistributed [toggleButton]
- PlusSign [button]
- DatabaseQueryNew [button]
- BorderTop [toggleButton]
- BorderLeft [toggleButton]
- BorderRight [toggleButton]
- ShapeStraightConnector [toggleButton]
- ShapeRoundedRectangularCallout [toggleButton]
- Calculator [button]
- MinusSign [button]
- ActiveXButton [button]
- FormControlButton [button]
- ShapeRightBrace [toggleButton]
- WebPagePreview [button]
- BorderInsideHorizontal [button]
- ZoomOut [button]
- SpeakCells [button]
- ResetAddinShortcuts [button]
- ZoomIn [button]
- BorderInside [button]
- ShapeRoundedRectangle [toggleButton]
- DarkShading [button]
- SelectCurrentRegion [button]
- DivisionSign [button]
- ShapeArc [toggleButton]
- ObjectFlipHorizontal [button]
- MultiplicationSign [button]
- BorderInsideVertical [button]
- CheckForUpdates [button]
- SpeakStop [button]
- ShapeRightArrow [toggleButton]
- ShapeScribble [toggleButton]
- ReviewSendForReview [button]
- WebGoBack [button]
- ObjectRotateFree [button]
- ShapeLeftBrace [toggleButton]
- WordArtVerticalText [button]
- ObjectNudgeLeft [button]
- ShapeDownArrow [toggleButton]
- FormControlListBox [button]
- ParenthesisRight [button]
- ParenthesisLeft [button]
- FormControlComboBox [button]
- OutlookTaskCreate [button]
- ObjectRotateRight90 [button]
- ListSynchronize [button]
- ReviewReplyWithChanges [button]
- _3DSurfaceMaterialClassic [menu]
- AutoCorrect [button]
- ShapeElbowConnector [toggleButton]
- FormControlRadioButton [button]
- FormControlCheckBox [button]
- ActiveXCheckBox [button]
- BorderDiagonalUp [button]
- ShapeIsoscelesTriangle [toggleButton]
- FileUpdate [button]
- PrivacyOptionsDialog [button]
- ObjectNudgeDown [button]
- ObjectNudgeRight [button]
- ObjectFlipVertical [button]
- FormControlSpinner [button]
- ActiveXComboBox [button]
- VerticallyDistributed [button]
- StyleGalleryClassic [gallery]
- SpeakOnEnter [toggleButton]
- ActiveXTextBox [button]
- FormControlLabel [button]
- ActiveXListBox [button]
- MoreControlsDialog [button]
- FormControlGroupBox [button]
- ChangesDiscardAndRefresh [button]
- FormControlScrollBar [button]
- ObjectNudgeUp [button]
- ActiveXImage [button]
- ActiveXLabel [button]
- DocumentLocation [comboBox]
- ActiveXSpinButton [button]
- ActiveXRadioButton [button]
- FileCloseAll [button]
- WebGoForward [button]
- BorderDiagonalDown [button]
- QueryParameters [button]
- DollarSign [button]
- ShapeStar [toggleButton]
- ActiveXScrollBar [button]
- ObjectsMultiSelect [button]
- ReviewEndReview [button]
- ExchangeFolder [button]
- OutlineSymbolsShowHide [button]
- SpeakByRows [toggleButton]
- SpeakByColumns [toggleButton]
- SelectAll [button]
- ReadOnly [button]
- TabOrder [button]
- ActiveXToggleButton [button]
- FormControlCombinationDropDownEdit [button]
- CodeEdit [button]
- CompareAndMergeWorkbooks [button]
- ExponentiationSign [button]
- PrintListRange [button]
- FormControlEditBox [button]
- PercentSign [button]
- FormControlCombinationListEdit [button]
- CommaSign [button]
- DrawingCanvasExpand [button]
- SendCopySendNow [button]
- DrawingCanvasFit [button]
- ScenarioGallery [gallery]
- FilePublishAsWebPage [button]
- InkCopyAsText [button]
- ColonSign [button]
- ConstrainNumeric [button]
- WebOptionsDialog [button]
- CalculateFull [button]
- SendCopySelectNames [button]
- DrawingCanvasResize [button]
- FontColorCycle [button]
- ChartInsert [button]
- CombineCharacters [toggleButton]
- BulletsAndNumberingBulletsDialog [button]
- QuickAccessToolbarCustomization [button]
- RibbonCustomization [button]
- AlternativeText [button]
- AdvancedFileProperties [button]
- CustomViews [comboBox]
- PictureBrightnessGallery [gallery]
- PictureContrastGallery [gallery]
- PictureCorrectionsReset [button]
- PictureColorReset [button]
- PictureArtisticEffectsReset [button]
- FileDiscardCachedChanges [button]
- RefreshServerFile [button]
- ReopenServerFile [button]
- ReopenFromServer [button]
- FileSaveACopy [button]
- None [button]
- PivotChartFilterShow [toggleButton]
- TranslationPane [button]
- ObjectSizeAndPositionDialog [button]
- About [button]
- ReviewProtectWorkbook [toggleButton]
- InkDeleteAllInk [button]
- ObjectBringToFrontMenu [splitButton]
- ObjectSendToBackMenu [splitButton]
- PrintPreviewFullScreen [button]
- FilePrintPreview [button]
- ClipArtInsert [toggleButton]
- SlicerInsertMenu [splitButton]
- PivotChartInsert [button]
- FileOpen [button]
- WindowSaveWorkspace [button]
- Collaborate [toggleButton]
- CollaborateComments [toggleButton]
- CollaborateActivity [toggleButton]
- None [toggleButton]
- WhatsNewRecentUpdates [button]
- GroupOfficeInsider [button]
- OfficeFeedbackBackstage [button]
- HyperlinkInsert [button]
- MarkAsDecorative [button]
- GenerateAutoAltText [button]
- ManagePrivacy [button]
- None [button]
- UpgradeDocument [button]
- PasteInk [button]
- PasteFormatting [button]
- PasteFormulasAndNumberFormatting [button]
- PasteValuesAndNumberFormatting [button]
- PasteTextOnly [button]
- PasteDestinationFormatting [button]
- PasteSourceTheme [button]
- PasteDestinationTheme [button]
- PasteValuesAndSourceFormatting [button]
- PasteWithColumnWidths [button]
- PasteSourceFormatting [button]
- PasteFormulas [button]
- PasteValues [button]
- PasteNoBorders [button]
- PasteTranspose [button]
- PasteLink [button]
- PasteAsHyperlink [button]
- PasteMergeConditionalFormatting [button]
- PasteAsPictureMenu [menu]
- InkingStart [button]
- InkToolsClose [button]
- ChartTypeAllInsertDialog [button]
- ReviewShareWorkbookLegacy [button]
- ReviewProtectAndShareWorkbookLegacy [button]
- ReviewTrackChangesMenuLegacy [menu]
- NewWebQueryLegacy [button]
- ImportTextFileLegacy [button]
- ImportAccessFileLegacy [button]
- XlNewOdataConnLegacy [button]
- XlNewDcwConnLegacy [button]
- XlNewXmlConnLegacy [button]
- XlNewSqlConnLegacy [button]
- OfficeExtensionsAppStore [button]
- OfficeExtensionsGallery2 [gallery]
- OfficeExtensionsFeaturedApp1 [button]
- OfficeExtensionsFeaturedApp2 [button]
- OfficeExtensionsFeaturedApp3 [button]
- OfficeExtensionsFeaturedApp4 [button]
- OfficeExtensionsFeaturedApp5 [button]
- OfficeExtensionsFeaturedApp6 [button]
- Group3DEffects [group]
- GroupShadowEffects [group]
- GroupDrawBorders [group]
- GroupInkSelect [group]
- PictureInsertFromFile [button]
- TableInsertExcel [button]
- PivotTableInsert [button]
- SortCustomExcel [button]
- FormatCellsDialog [button]
- PrintAreaSetPrintArea [button]
- CalculateNow [button]
- MacroPlay [button]
- SheetRowsInsert [button]
- SheetColumnsInsert [button]
- SheetRowsDelete [button]
- SheetColumnsDelete [button]
- CellsInsertDialog [button]
- CellsDelete [button]
- RefreshAll [button]
- Connections [button]
- Filter [toggleButton]
- NameManager [button]
- FunctionWizard [button]
- MergeCenter [toggleButton]
- AlignCenter [toggleButton]
- Cut [button]
- Copy [button]
- FormatPainter [control]
- AutoSum [button]
- FontSizeIncrease [button]
- FontSizeDecrease [button]
- FileSaveAs [button]
- PageSetupPageDialog [button]
- ViewFreezePanesGallery [gallery]
- ShapesInsertGallery [gallery]
- FontSize [comboBox]
- Font [comboBox]
- PasteMenu [splitButton]
- CellFillColorPicker [gallery]
- FontColorPicker [gallery]
- BordersGallery [splitButton]
- ConditionalFormattingMenu [menu]
- UpgradeWorkbook [button]
- FileExit [button]
- FileNew [button]
- ApplicationOptionsDialog [button]
- None [button]
- FileStartWorkflow [button]
- LeaveReader [button]
- SheetProtect [button]
- MenuAccessibility [splitButton]
- OfficeFeedback [button]
- AppGuardTrustPromotion [button]
- OfficeExtensionsGetAddinsButtonOnBackstageBelowOpenDivider [button]
- FilePrepareMenu [menu]
- FilePrintMenu [splitButton]
- MenuPublish [menu]
- FileSaveAsMenu [splitButton]
- FileSendMenu [menu]
- FileServerTasksMenu [menu]
- ReviewProtectWorkbookMenu [menu]
### None (Not in the Ribbon)
-   _3DSurfaceWireFrameClassic [toggleButton] < _3DSurfaceMaterialClassic
-   _3DSurfaceMatteClassic [toggleButton] < _3DSurfaceMaterialClassic
-   _3DSurfacePlasticClassic [toggleButton] < _3DSurfaceMaterialClassic
-   _3DSurfaceMetalClassic [toggleButton] < _3DSurfaceMaterialClassic
-   PictureCorrectionsDialog [button] < PictureBrightnessGallery
-   PictureCorrectionsDialog [button] < PictureContrastGallery
-   ObjectBringToFront [button] < ObjectBringToFrontMenu
-   ObjectBringForward [button] < ObjectBringToFrontMenu
-   ObjectSendToBack [button] < ObjectSendToBackMenu
-   ObjectSendBackward [button] < ObjectSendToBackMenu
-   SlicerCreate [button] < SlicerInsertMenu
-   SlicerShare [button] < SlicerInsertMenu
-   CopyAsPicture [button] < PasteAsPictureMenu
-   PasteAsPicture [button] < PasteAsPictureMenu
-   PastePictureLink [button] < PasteAsPictureMenu
-   ReviewHighlightChanges [button] < ReviewTrackChangesMenuLegacy
-   ReviewAcceptOrRejectChangeDialog [button] < ReviewTrackChangesMenuLegacy
-   OfficeExtensionsDialog [button] < OfficeExtensionsGallery2
-   OfficeExtensionsManageOtherAddins [button] < OfficeExtensionsGallery2
-   ExtensibilityExperimentation [button] < OfficeExtensionsGallery2
-   Paste [button] < PasteMenu
-   PasteGallery [gallery] < PasteMenu
-     PasteUsingTextImportWizard [button] < PasteMenu
-     PasteRefreshableWebQuery [button] < PasteMenu
-     PasteSpecialDialog [button] < PasteMenu
-   FontShadingColorMoreColorsDialog [button] < CellFillColorPicker
-   FontColorMoreColorsDialogExcel [button] < FontColorPicker
-   BorderBottomNoToggle [button] < BordersGallery
-   BorderTopNoToggle [button] < BordersGallery
-   BorderLeftNoToggle [button] < BordersGallery
-   BorderRightNoToggle [button] < BordersGallery
-   BorderNone [button] < BordersGallery
-   BordersAll [button] < BordersGallery
-   BorderOutside [button] < BordersGallery
-   BorderThickOutside [button] < BordersGallery
-   BorderDoubleBottom [button] < BordersGallery
-   BorderThickBottom [button] < BordersGallery
-   BorderTopAndBottom [button] < BordersGallery
-   BorderTopAndThickBottom [button] < BordersGallery
-   BorderTopAndDoubleBottom [button] < BordersGallery
-   BorderDrawLine [toggleButton] < BordersGallery
-   BorderDrawGrid [toggleButton] < BordersGallery
-   BorderErase [toggleButton] < BordersGallery
-   BorderColorPickerExcel [gallery] < BordersGallery
-     BorderMoreColorsDialog [button] < BordersGallery
-   BorderStyle [dropDown] < BordersGallery
-   BordersMoreDialog [button] < BordersGallery
-   ConditionalFormattingHighlightCellsMenu [menu] < ConditionalFormattingMenu
-     ConditionalFormattingHighlightGreaterThan [button] < ConditionalFormattingMenu
-     ConditionalFormattingHighlightLessThan [button] < ConditionalFormattingMenu
-     ConditionalFormattingHighlightBetween [button] < ConditionalFormattingMenu
-     ConditionalFormattingHighlightEqualTo [button] < ConditionalFormattingMenu
-     ConditionalFormattingHighlightTextContaining [button] < ConditionalFormattingMenu
-     ConditionalFormattingHighlightDateOccuring [button] < ConditionalFormattingMenu
-     ConditionalFormattingHighlightDuplicateValues [button] < ConditionalFormattingMenu
-     ConditionalFormattingHighlightRulesMore [button] < ConditionalFormattingMenu
-   ConditionalFormattingTopBottomMenu [menu] < ConditionalFormattingMenu
-     ConditionalFormattingTopNItems [button] < ConditionalFormattingMenu
-     ConditionalFormattingTopNPercent [button] < ConditionalFormattingMenu
-     ConditionalFormattingBottomNItems [button] < ConditionalFormattingMenu
-     ConditionalFormattingBottomNPercent [button] < ConditionalFormattingMenu
-     ConditionalFormattingAboveAverage [button] < ConditionalFormattingMenu
-     ConditionalFormattingBelowAverage [button] < ConditionalFormattingMenu
-     ConditionalFormattingTopBottomMore [button] < ConditionalFormattingMenu
-   ConditionalFormattingDataBarsGallery [gallery] < ConditionalFormattingMenu
-     ConditionalFormattingDataBarsMoreOptions [button] < ConditionalFormattingMenu
-   ConditionalFormattingColorScalesGallery [gallery] < ConditionalFormattingMenu
-     ConditionalFormattingColorScalesMore [button] < ConditionalFormattingMenu
-   ConditionalFormattingIconSetsGallery [gallery] < ConditionalFormattingMenu
-     ConditionalFormattingIconSetsMore [button] < ConditionalFormattingMenu
-   ConditionalFormattingNewRule [button] < ConditionalFormattingMenu
-   ConditionalFormattingClearMenu [menu] < ConditionalFormattingMenu
-     ConditionalFormattingClearSelectedCells [button] < ConditionalFormattingMenu
-     ConditionalFormattingClearSheet [button] < ConditionalFormattingMenu
-     ConditionalFormattingClearTable [button] < ConditionalFormattingMenu
-     ConditionalFormattingClearPivotTable [button] < ConditionalFormattingMenu
-   ConditionalFormattingsManage [button] < ConditionalFormattingMenu
-   AccessibilityChecker [button] < MenuAccessibility
-   AltTextPaneRibbon [toggleButton] < MenuAccessibility
-   FormatAsTableGallery [gallery] < MenuAccessibility
-     TableStyleNew [button] < MenuAccessibility
-     PivotTableNewStyle [button] < MenuAccessibility
-   UnmergeCells [button] < MenuAccessibility
-   OpenOptionsEaseOfAcess [button] < MenuAccessibility
-   AccessibilityHelp [button] < MenuAccessibility
-   FileProperties [toggleButton] < FilePrepareMenu
-   FileDocumentInspect [button] < FilePrepareMenu
-   FileDocumentEncrypt [toggleButton] < FilePrepareMenu
-   FilePermissionView [button] < FilePrepareMenu
-   FilePermissionRestrictMenu [menu] < FilePrepareMenu
-     FilePermissionUnrestricted [toggleButton] < FilePrepareMenu
-     FilePermissionDoNotDistribute [toggleButton] < FilePrepareMenu
-   FileAddDigitalSignature [button] < FilePrepareMenu
-   FileViewDigitalSignatures [toggleButton] < FilePrepareMenu
-   FileMarkAsFinal [toggleButton] < FilePrepareMenu
-   FileMarkAsReadOnlyRecommended [toggleButton] < FilePrepareMenu
-   FileCompatibilityChecker [button] < FilePrepareMenu
-   FileLinksToFiles [button] < FilePrepareMenu
-   PrintPreviewAndPrint [button] < FilePrintMenu
-   FilePrintQuick [button] < FilePrintMenu
-   PrintPreviewFullScreen [button] < FilePrintMenu
-   FilePublishExcelServices [button] < MenuPublish
-   FileSaveToDocumentManagementServer [button] < MenuPublish
-   FileSaveAs [button] < FileSaveAsMenu
-   FileSaveAsExcelXlsx [button] < FileSaveAsMenu
-   FileSaveAsExcelXlsxMacro [button] < FileSaveAsMenu
-   FileSaveAsExcelXlsb [button] < FileSaveAsMenu
-   FileSaveAsExcel97_2003 [button] < FileSaveAsMenu
-   FileSaveAsExcelOpenDocumentSpreadsheet [button] < FileSaveAsMenu
-   AdvertisePublishAs [button] < FileSaveAsMenu
-   FileSaveAsPdfOrXps [button] < FileSaveAsMenu
-   FileSaveAsOtherFormats [button] < FileSaveAsMenu
-   FileSendAsAttachment [button] < FileSendMenu
-   FileEmailAsPdfEmailAttachment [button] < FileSendMenu
-   FileEmailAsXpsEmailAttachment [button] < FileSendMenu
-   FileInternetFax [button] < FileSendMenu
-   FileCheckIn [button] < FileServerTasksMenu
-   FileCheckOut [button] < FileServerTasksMenu
-   FileCheckOutDiscard [button] < FileServerTasksMenu
-   FileVersionHistory [button] < FileServerTasksMenu
-   FileWorkflowTasks [button] < FileServerTasksMenu
-   FileExcelServicesOptions [button] < FileServerTasksMenu
-   LineageClearHistory [button] < FileServerTasksMenu
-   ReviewRestrictEditing [toggleButton] < ReviewProtectWorkbookMenu
-   FilePermissionUnrestricted [toggleButton] < ReviewProtectWorkbookMenu
-   FilePermissionDoNotDistribute [toggleButton] < ReviewProtectWorkbookMenu
### Group3DEffects
- _3DEffectsGalleryClassic [gallery]
-   _3DEffectColorPickerClassic [gallery] < _3DEffectsGalleryClassic
-     _3DEffectColorPickerMoreClassic [button] < _3DEffectsGalleryClassic
-   _3DExtrusionDepthGalleryClassic [gallery] < _3DEffectsGalleryClassic
-   _3DDirectionGalleryClassic [gallery] < _3DEffectsGalleryClassic
-     _3DExtrusionPerspectiveClassic [toggleButton] < _3DEffectsGalleryClassic
-     _3DExtrusionParallelClassic [toggleButton] < _3DEffectsGalleryClassic
-   _3DLightingGalleryClassic [gallery] < _3DEffectsGalleryClassic
-     _3DLightingFlatClassic [toggleButton] < _3DEffectsGalleryClassic
-     _3DLightingNormalClassic [toggleButton] < _3DEffectsGalleryClassic
-     _3DLightingDimClassic [toggleButton] < _3DEffectsGalleryClassic
-   _3DSurfaceMaterialGalleryClassic [gallery] < _3DEffectsGalleryClassic
- _3DTiltUpClassic [button]
- _3DTiltLeftClassic [button]
- _3DEffectsOnOffClassic [toggleButton]
- _3DTiltRightClassic [button]
- _3DTiltDownClassic [button]
### GroupShadowEffects
- ShadowStyleGalleryClassic [gallery]
-   ShadowColorPickerClassic [gallery] < ShadowStyleGalleryClassic
-     ShadowSemitransparentClassic [toggleButton] < ShadowStyleGalleryClassic
-     ObjectShadowColorMoreColorsDialog [button] < ShadowStyleGalleryClassic
- ShadowNudgeUpClassic [button]
- ShadowNudgeLeftClassic [button]
- ShadowOnOrOffClassic [toggleButton]
- ShadowNudgeRightClassic [button]
- ShadowNudgeDownClassic [button]
### GroupDrawBorders
- BorderStyle [dropDown]
- BorderColorPickerExcel [gallery]
-   BorderMoreColorsDialog [button] < BorderColorPickerExcel
- BorderDrawMenu [splitButton]
-   BorderDrawLine [toggleButton] < BorderDrawMenu
-   BorderDrawGrid [toggleButton] < BorderDrawMenu
- BorderErase [toggleButton]
### GroupInkSelect
- ObjectsSelect [toggleButton]

##   ()
### 
- TabSetSmartArtTools [tabSet]
- TabSetChartTools [tabSet]
- TabSetDrawingTools [tabSet]
- TabSetPictureTools [tabSet]
- TabSetPivotTableTools [tabSet]
- TabSetHeaderAndFooterTools [tabSet]
- TabSetTableToolsExcel [tabSet]
- TabSetPivotChartTools [tabSet]
- TabSetInkTools [tabSet]
- TabSetSparkline [tabSet]
- TabSetTimeSlicerTools [tabSet]
- TabSetSlicerTools [tabSet]
- TabSetPowerQueryEdit [tabSet]
- TabSetEquationTools [tabSet]
- TabSetSVGTools [tabSet]
- TabSet3DModelTools [tabSet]
- TabSetBIVisualTools [tabSet]
- TabSetAccessibleAuthoring [tabSet]

##   (TabSetSmartArtTools)
### 
- TabSmartArtToolsDesign [tab]
- TabSmartArtToolsFormat [tab]

## TabSmartArtToolsDesign  (TabSetSmartArtTools)
### 
- GroupSmartArtCreateGraphic [group]
- GroupSmartArtLayouts [group]
- GroupSmartArtQuickStyles [group]
- GroupSmartArtReset [group]
### GroupSmartArtCreateGraphic
- SmartArtAddShapeSplitMenu [splitButton]
-   SmartArtAddShape [button] < SmartArtAddShapeSplitMenu
-   SmartArtAddShapeAfter [button] < SmartArtAddShapeSplitMenu
-   SmartArtAddShapeBefore [button] < SmartArtAddShapeSplitMenu
-   SmartArtAddShapeAbove [button] < SmartArtAddShapeSplitMenu
-   SmartArtAddShapeBelow [button] < SmartArtAddShapeSplitMenu
-   SmartArtAddAssistant [button] < SmartArtAddShapeSplitMenu
- SmartArtAddBullet [button]
- SmartArtTextPane [toggleButton]
- SmartArtPromote [button]
- SmartArtDemote [button]
- SmartArtRightToLeft [toggleButton]
- SmartArtReorderUp [button]
- SmartArtReorderDown [button]
- SmartArtOrganizationChartMenu [menu]
-   SmartArtOrganizationChartStandard [button] < SmartArtOrganizationChartMenu
-   SmartArtOrganizationChartBoth [button] < SmartArtOrganizationChartMenu
-   SmartArtOrganizationChartLeftHanging [button] < SmartArtOrganizationChartMenu
-   SmartArtOrganizationChartRightHanging [button] < SmartArtOrganizationChartMenu
### GroupSmartArtLayouts
- SmartArtLayoutGallery [gallery]
-   SmartArtMoreLayoutsDialog [button] < SmartArtLayoutGallery
### GroupSmartArtQuickStyles
- SmartArtChangeColorsGallery [gallery]
-   SmartArtRecolorImages [toggleButton] < SmartArtChangeColorsGallery
- SmartArtStylesGallery [gallery]
### GroupSmartArtReset
- SmartArtResetGraphic [button]
- SmartArtConvertToShapes [button]

## TabSmartArtToolsFormat  (TabSetSmartArtTools)
### 
- GroupSmartArtShapes [group]
- GroupShapeStyles [group]
- GroupWordArtStyles [group]
- GroupAltText [group]
- GroupArrangeWith3DEditor [group]
- GroupSmartArtSize [group]
### GroupSmartArtShapes
- SmartArtEditIn2D [toggleButton]
- ShapeChangeShapeGallery [gallery]
- SmartArtLargerShape [button]
- SmartArtSmallerShape [button]
### GroupShapeStyles
- ShapeStylesGallery [gallery]
-   ShapeStylesOtherThemeFillsGallery [gallery] < ShapeStylesGallery
- ShapeFillColorPicker [gallery]
-   ObjectFillMoreColorsDialog [button] < ShapeFillColorPicker
-   ObjectPictureFill [button] < ShapeFillColorPicker
-   GradientGallery [gallery] < ShapeFillColorPicker
-     ShapeFillMoreGradientsDialog [button] < ShapeFillColorPicker
-   ShapeFillTextureGallery [gallery] < ShapeFillColorPicker
-     MoreTextureOptions [button] < ShapeFillColorPicker
- ShapeOutlineColorPicker [gallery]
-   ObjectBorderOutlineColorMoreColorsDialog [button] < ShapeOutlineColorPicker
-   OutlineWeightGallery [gallery] < ShapeOutlineColorPicker
-     LineStylesDialog [button] < ShapeOutlineColorPicker
-   GallerySketchyShapes [gallery] < ShapeOutlineColorPicker
-     MoreLines2 [button] < ShapeOutlineColorPicker
-   OutlineDashesGallery [gallery] < ShapeOutlineColorPicker
-     LineStylesDialog [button] < ShapeOutlineColorPicker
-   ArrowStyleGallery [gallery] < ShapeOutlineColorPicker
-     ArrowsMore [button] < ShapeOutlineColorPicker
- ShapeEffectsMenu [menu]
-   ObjectEffectPresetGallery [gallery] < ShapeEffectsMenu
-     _3DBevelOptionsDialog [button] < ShapeEffectsMenu
-   ObjectEffectShadowGallery [gallery] < ShapeEffectsMenu
-     ShadowOptionsDialog [button] < ShapeEffectsMenu
-   ReflectionGallery [gallery] < ShapeEffectsMenu
-     ReflectionsMoreOptions [button] < ShapeEffectsMenu
-   ObjectEffectGlowGallery [gallery] < ShapeEffectsMenu
-     GlowColorPicker [gallery] < ShapeEffectsMenu
-     GlowColorMoreColorsDialog [button] < ShapeEffectsMenu
-     GlowsMoreOptions [button] < ShapeEffectsMenu
-   ObjectEffectSoftEdgesGallery [gallery] < ShapeEffectsMenu
-     SoftEdgesMoreOptions [button] < ShapeEffectsMenu
-   BevelShapeGallery [gallery] < ShapeEffectsMenu
-     _3DBevelOptionsDialog [button] < ShapeEffectsMenu
-   _3DRotationGallery [gallery] < ShapeEffectsMenu
-     _3DRotationOptionsDialog [button] < ShapeEffectsMenu
- ObjectFormatDialog [button (dialogBoxLauncher)]
### GroupWordArtStyles
- TextStylesGallery [gallery]
-   WordArtClear [button] < TextStylesGallery
- TextFillColorPicker [gallery]
-   TextFillColorMoreColorsDialog [button] < TextFillColorPicker
-   TextPictureFill [button] < TextFillColorPicker
-   TextFillGradientGallery [gallery] < TextFillColorPicker
-     TextFillMoreGradientsDialog [button] < TextFillColorPicker
-   TextFillTextureGallery [gallery] < TextFillColorPicker
-     TextFillMoreTextures [button] < TextFillColorPicker
- TextOutlineColorPicker [gallery]
-   TextOutlineColorMoreColorsDialog [button] < TextOutlineColorPicker
-   TextOutlineWeightGallery [gallery] < TextOutlineColorPicker
-     TextOutlineMoreLinesDialog [button] < TextOutlineColorPicker
-   TextOutlineDashesGallery [gallery] < TextOutlineColorPicker
-     TextOutlineMoreLinesDialog [button] < TextOutlineColorPicker
- TextEffectsMenu [menu]
-   TextEffectShadowGallery [gallery] < TextEffectsMenu
-     TextEffectsMoreShadowsDialog [button] < TextEffectsMenu
-   TextReflectionGallery [gallery] < TextEffectsMenu
-     TextEffectsReflectionsMoreOptions [button] < TextEffectsMenu
-   TextEffectGlowGallery [gallery] < TextEffectsMenu
-     TextGlowColorPicker [gallery] < TextEffectsMenu
-     TextGlowColorMoreColorsDialog [button] < TextEffectsMenu
-     TextEffectsGlowsMoreOptions [button] < TextEffectsMenu
-   BevelTextGallery [gallery] < TextEffectsMenu
-     TextEffectsBevelMore3DOptionsDialog [button] < TextEffectsMenu
-   TextEffects3DRotationGallery [gallery] < TextEffectsMenu
-     TextEffects3DRotationOptionsDialog [button] < TextEffectsMenu
-   TextEffectTransformGallery [gallery] < TextEffectsMenu
- WordArtFormatDialog [button (dialogBoxLauncher)]
### GroupAltText
- AltTextPaneRibbon [toggleButton]
### GroupArrangeWith3DEditor
- ObjectBringForwardMenu [splitButton]
-   ObjectBringForward [button] < ObjectBringForwardMenu
-   ObjectBringToFront [button] < ObjectBringForwardMenu
- ObjectSendBackwardMenu [splitButton]
-   ObjectSendBackward [button] < ObjectSendBackwardMenu
-   ObjectSendToBack [button] < ObjectSendBackwardMenu
- SelectionPane [toggleButton]
- ObjectAlignMenu [menu]
-   ObjectsAlignLeft [button] < ObjectAlignMenu
-   ObjectsAlignCenterHorizontal [button] < ObjectAlignMenu
-   ObjectsAlignRight [button] < ObjectAlignMenu
-   ObjectsAlignTop [button] < ObjectAlignMenu
-   ObjectsAlignMiddleVertical [button] < ObjectAlignMenu
-   ObjectsAlignBottom [button] < ObjectAlignMenu
-   AlignDistributeHorizontallyClassic [button] < ObjectAlignMenu
-   AlignDistributeVerticallyClassic [button] < ObjectAlignMenu
-   SnapToGrid [toggleButton] < ObjectAlignMenu
-   SnapToShapes [toggleButton] < ObjectAlignMenu
-   ViewGridlinesToggleExcel [toggleButton] < ObjectAlignMenu
- ObjectsGroupMenu [menu]
-   ObjectsGroup [button] < ObjectsGroupMenu
-   ObjectsRegroup [button] < ObjectsGroupMenu
-   ObjectsUngroup [button] < ObjectsGroupMenu
- ObjectRotateGallery [gallery]
-   ObjectRotationOptionsDialog [button] < ObjectRotateGallery
- ObjectsArrangeMenu [menu]
-   ObjectBringToFront [button] < ObjectsArrangeMenu
-   ObjectBringForward [button] < ObjectsArrangeMenu
-   ObjectSendToBack [button] < ObjectsArrangeMenu
-   ObjectSendBackward [button] < ObjectsArrangeMenu
-   ObjectsGroup [button] < ObjectsArrangeMenu
-   ObjectsUngroup [button] < ObjectsArrangeMenu
-   ObjectsRegroup [button] < ObjectsArrangeMenu
-   ObjectAlignMenu [menu] < ObjectsArrangeMenu
-     ObjectsAlignLeft [button] < ObjectsArrangeMenu
-     ObjectsAlignCenterHorizontal [button] < ObjectsArrangeMenu
-     ObjectsAlignRight [button] < ObjectsArrangeMenu
-     ObjectsAlignTop [button] < ObjectsArrangeMenu
-     ObjectsAlignMiddleVertical [button] < ObjectsArrangeMenu
-     ObjectsAlignBottom [button] < ObjectsArrangeMenu
-     AlignDistributeHorizontallyClassic [button] < ObjectsArrangeMenu
-     AlignDistributeVerticallyClassic [button] < ObjectsArrangeMenu
-     SnapToGrid [toggleButton] < ObjectsArrangeMenu
-     SnapToShapes [toggleButton] < ObjectsArrangeMenu
-     ViewGridlinesToggleExcel [toggleButton] < ObjectsArrangeMenu
-   ObjectRotateGallery [gallery] < ObjectsArrangeMenu
-     ObjectRotationOptionsDialog [button] < ObjectsArrangeMenu
### GroupSmartArtSize
- ShapeHeight [control]
- ShapeWidth [control]
- ObjectSizeAndPropertiesDialog [button (dialogBoxLauncher)]

##   (TabSetChartTools)
### 
- TabChartToolsDesignNew [tab]
- TabChartToolsFormatNew [tab]
- TabChartToolsDesign [tab]
- TabChartToolsLayout [tab]
- TabChartToolsFormat [tab]

## TabChartToolsDesignNew  (TabSetChartTools)
### 
- GroupChartLayouts [group]
- GroupChartStyles [group]
- GroupChartData [group]
- GroupChartType [group]
- GroupChartLocation [group]
### GroupChartLayouts
- AddChartElementMenu [menu]
-   ChartAxesGallery [gallery] < AddChartElementMenu
-     ChartAxesOptionsDialog [button] < AddChartElementMenu
-   ChartAxisTitlesGallery [gallery] < AddChartElementMenu
-     ChartAxesTitlesOptionsDialog [button] < AddChartElementMenu
-   ChartTitle [gallery] < AddChartElementMenu
-     ChartTitleOptionsDialog [button] < AddChartElementMenu
-   ChartDataLabel [gallery] < AddChartElementMenu
-     ChartDataLabelDialog [button] < AddChartElementMenu
-   ChartDataTable [gallery] < AddChartElementMenu
-     ChartDataTableOptionsDialog [button] < AddChartElementMenu
-   ChartErrorBars [gallery] < AddChartElementMenu
-     ChartErrorBarsOptionsDialog [button] < AddChartElementMenu
-   ChartGridlinesGallery [gallery] < AddChartElementMenu
-     ChartGridlineOptionsDialog [button] < AddChartElementMenu
-   ChartLegend [gallery] < AddChartElementMenu
-     ChartLegendOptionsDialogDialog [button] < AddChartElementMenu
-   ChartLines [gallery] < AddChartElementMenu
-   ChartTrendline [gallery] < AddChartElementMenu
-     ChartTrendlineOptionsDialog [button] < AddChartElementMenu
-   ChartUpDownBars [gallery] < AddChartElementMenu
-     ChartUpDownBarsOptionsDialog [button] < AddChartElementMenu
- ChartLayoutGallery [gallery]
### GroupChartStyles
- ChartColorsGallery [gallery]
- ChartStylesGallery [gallery]
- ChartDesigner [button]
### GroupChartData
- ChartSwitchRowColumn [button]
- ChartEditDataSource [button]
### GroupChartType
- ChartChangeType [button]
### GroupChartLocation
- ChartPlacement [button]

## TabChartToolsFormatNew  (TabSetChartTools)
### 
- GroupChartCurrentSelection [group]
- GroupShapesChart [group]
- GroupChartShapeStyles [group]
- GroupWordArtStyles [group]
- GroupAltText [group]
- GroupArrange [group]
- GroupSize [group]
### GroupChartCurrentSelection
- ChartElementSelector [comboBox]
- ChartFormatSelection [button]
- ChartResetToMatchStyle [button]
### GroupShapesChart
- ShapesInsertGallery [gallery]
- ShapeChangeShapeGallery [gallery]
### GroupChartShapeStyles
- ShapeStylesGallery [gallery]
-   ShapeStylesOtherThemeFillsGallery [gallery] < ShapeStylesGallery
- ShapeFillColorPicker [gallery]
-   ObjectFillMoreColorsDialog [button] < ShapeFillColorPicker
-   ObjectPictureFill [button] < ShapeFillColorPicker
-   GradientGallery [gallery] < ShapeFillColorPicker
-     ShapeFillMoreGradientsDialog [button] < ShapeFillColorPicker
-   ShapeFillTextureGallery [gallery] < ShapeFillColorPicker
-     MoreTextureOptions [button] < ShapeFillColorPicker
- ColorPickerChartLineStyles [gallery]
-   ObjectBorderOutlineColorMoreColorsDialog [button] < ColorPickerChartLineStyles
-   OutlineWeightGallery [gallery] < ColorPickerChartLineStyles
-     LineStylesDialog [button] < ColorPickerChartLineStyles
-   OutlineDashesGallery [gallery] < ColorPickerChartLineStyles
-     LineStylesDialog [button] < ColorPickerChartLineStyles
-   ArrowStyleGallery [gallery] < ColorPickerChartLineStyles
-     ArrowsMore [button] < ColorPickerChartLineStyles
- ShapeEffectsMenu [menu]
-   ObjectEffectPresetGallery [gallery] < ShapeEffectsMenu
-     _3DBevelOptionsDialog [button] < ShapeEffectsMenu
-   ObjectEffectShadowGallery [gallery] < ShapeEffectsMenu
-     ShadowOptionsDialog [button] < ShapeEffectsMenu
-   ReflectionGallery [gallery] < ShapeEffectsMenu
-     ReflectionsMoreOptions [button] < ShapeEffectsMenu
-   ObjectEffectGlowGallery [gallery] < ShapeEffectsMenu
-     GlowColorPicker [gallery] < ShapeEffectsMenu
-     GlowColorMoreColorsDialog [button] < ShapeEffectsMenu
-     GlowsMoreOptions [button] < ShapeEffectsMenu
-   ObjectEffectSoftEdgesGallery [gallery] < ShapeEffectsMenu
-     SoftEdgesMoreOptions [button] < ShapeEffectsMenu
-   BevelShapeGallery [gallery] < ShapeEffectsMenu
-     _3DBevelOptionsDialog [button] < ShapeEffectsMenu
-   _3DRotationGallery [gallery] < ShapeEffectsMenu
-     _3DRotationOptionsDialog [button] < ShapeEffectsMenu
- ObjectFormatDialog [button (dialogBoxLauncher)]
### GroupWordArtStyles
- TextStylesGallery [gallery]
-   WordArtClear [button] < TextStylesGallery
- TextFillColorPicker [gallery]
-   TextFillColorMoreColorsDialog [button] < TextFillColorPicker
-   TextPictureFill [button] < TextFillColorPicker
-   TextFillGradientGallery [gallery] < TextFillColorPicker
-     TextFillMoreGradientsDialog [button] < TextFillColorPicker
-   TextFillTextureGallery [gallery] < TextFillColorPicker
-     TextFillMoreTextures [button] < TextFillColorPicker
- TextOutlineColorPicker [gallery]
-   TextOutlineColorMoreColorsDialog [button] < TextOutlineColorPicker
-   TextOutlineWeightGallery [gallery] < TextOutlineColorPicker
-     TextOutlineMoreLinesDialog [button] < TextOutlineColorPicker
-   TextOutlineDashesGallery [gallery] < TextOutlineColorPicker
-     TextOutlineMoreLinesDialog [button] < TextOutlineColorPicker
- TextEffectsMenu [menu]
-   TextEffectShadowGallery [gallery] < TextEffectsMenu
-     TextEffectsMoreShadowsDialog [button] < TextEffectsMenu
-   TextReflectionGallery [gallery] < TextEffectsMenu
-     TextEffectsReflectionsMoreOptions [button] < TextEffectsMenu
-   TextEffectGlowGallery [gallery] < TextEffectsMenu
-     TextGlowColorPicker [gallery] < TextEffectsMenu
-     TextGlowColorMoreColorsDialog [button] < TextEffectsMenu
-     TextEffectsGlowsMoreOptions [button] < TextEffectsMenu
-   BevelTextGallery [gallery] < TextEffectsMenu
-     TextEffectsBevelMore3DOptionsDialog [button] < TextEffectsMenu
-   TextEffects3DRotationGallery [gallery] < TextEffectsMenu
-     TextEffects3DRotationOptionsDialog [button] < TextEffectsMenu
-   TextEffectTransformGallery [gallery] < TextEffectsMenu
- WordArtFormatDialog [button (dialogBoxLauncher)]
### GroupAltText
- AltTextPaneRibbon [toggleButton]
### GroupArrange
- ObjectBringForwardMenu [splitButton]
-   ObjectBringForward [button] < ObjectBringForwardMenu
-   ObjectBringToFront [button] < ObjectBringForwardMenu
- ObjectSendBackwardMenu [splitButton]
-   ObjectSendBackward [button] < ObjectSendBackwardMenu
-   ObjectSendToBack [button] < ObjectSendBackwardMenu
- SelectionPane [toggleButton]
- ObjectAlignMenu [menu]
-   ObjectsAlignLeft [button] < ObjectAlignMenu
-   ObjectsAlignCenterHorizontal [button] < ObjectAlignMenu
-   ObjectsAlignRight [button] < ObjectAlignMenu
-   ObjectsAlignTop [button] < ObjectAlignMenu
-   ObjectsAlignMiddleVertical [button] < ObjectAlignMenu
-   ObjectsAlignBottom [button] < ObjectAlignMenu
-   AlignDistributeHorizontallyClassic [button] < ObjectAlignMenu
-   AlignDistributeVerticallyClassic [button] < ObjectAlignMenu
-   SnapToGrid [toggleButton] < ObjectAlignMenu
-   SnapToShapes [toggleButton] < ObjectAlignMenu
-   ViewGridlinesToggleExcel [toggleButton] < ObjectAlignMenu
- ObjectsGroupMenu [menu]
-   ObjectsGroup [button] < ObjectsGroupMenu
-   ObjectsRegroup [button] < ObjectsGroupMenu
-   ObjectsUngroup [button] < ObjectsGroupMenu
- ObjectRotateGallery [gallery]
-   ObjectRotationOptionsDialog [button] < ObjectRotateGallery
- ObjectsArrangeMenu [menu]
-   ObjectBringToFront [button] < ObjectsArrangeMenu
-   ObjectBringForward [button] < ObjectsArrangeMenu
-   ObjectSendToBack [button] < ObjectsArrangeMenu
-   ObjectSendBackward [button] < ObjectsArrangeMenu
-   ObjectsGroup [button] < ObjectsArrangeMenu
-   ObjectsUngroup [button] < ObjectsArrangeMenu
-   ObjectsRegroup [button] < ObjectsArrangeMenu
-   ObjectAlignMenu [menu] < ObjectsArrangeMenu
-     ObjectsAlignLeft [button] < ObjectsArrangeMenu
-     ObjectsAlignCenterHorizontal [button] < ObjectsArrangeMenu
-     ObjectsAlignRight [button] < ObjectsArrangeMenu
-     ObjectsAlignTop [button] < ObjectsArrangeMenu
-     ObjectsAlignMiddleVertical [button] < ObjectsArrangeMenu
-     ObjectsAlignBottom [button] < ObjectsArrangeMenu
-     AlignDistributeHorizontallyClassic [button] < ObjectsArrangeMenu
-     AlignDistributeVerticallyClassic [button] < ObjectsArrangeMenu
-     SnapToGrid [toggleButton] < ObjectsArrangeMenu
-     SnapToShapes [toggleButton] < ObjectsArrangeMenu
-     ViewGridlinesToggleExcel [toggleButton] < ObjectsArrangeMenu
-   ObjectRotateGallery [gallery] < ObjectsArrangeMenu
-     ObjectRotationOptionsDialog [button] < ObjectsArrangeMenu
### GroupSize
- ShapeHeight [control]
- ShapeWidth [control]
- ObjectSizeAndPropertiesDialog [button (dialogBoxLauncher)]

## TabChartToolsDesign  (TabSetChartTools)
### 
- GroupChartData [group]
- GroupChartStyles [group]
- GroupChartLocation [group]
### GroupChartData
- ChartSwitchRowColumn [button]
- ChartEditDataSource [button]
### GroupChartStyles
- ChartColorsGallery [gallery]
- ChartStylesGallery [gallery]
- ChartDesigner [button]
### GroupChartLocation
- ChartPlacement [button]

## TabChartToolsLayout  (TabSetChartTools)
### 
- GroupChartCurrentSelection [group]
- GroupChartShapes [group]
- GroupChartLabels [group]
- GroupChartAxes [group]
- GroupChartBackground [group]
- GroupChartAnalysis [group]
- GroupChartProperties [group]
### GroupChartCurrentSelection
- ChartElementSelector [comboBox]
- ChartFormatSelection [button]
- ChartResetToMatchStyle [button]
### GroupChartShapes
- PictureInsertFromFile [button]
- ShapesInsertGallery [gallery]
- TextBoxInsertExcel [toggleButton]
- TextBoxInsertMenu [splitButton]
-   TextBoxInsertHorizontal [toggleButton] < TextBoxInsertMenu
-   TextBoxInsertVertical [toggleButton] < TextBoxInsertMenu
### GroupChartLabels
- ChartTitle [gallery]
-   ChartTitleOptionsDialog [button] < ChartTitle
- ChartAxisTitles [menu]
-   ChartPrimaryHorizontalAxisTitle [gallery] < ChartAxisTitles
-     ChartPrimaryHorizontalAxisTitleOptionsDialog [button] < ChartAxisTitles
-   ChartPrimaryVerticalAxisTitle [gallery] < ChartAxisTitles
-     ChartPrimaryVerticalAxisTitleOptionsDialog [button] < ChartAxisTitles
-   ChartSecondaryHorizontalAxisTitle [gallery] < ChartAxisTitles
-     ChartSecondaryHorizontalAxisTitleOptionsDialog [button] < ChartAxisTitles
-   ChartSecondaryVerticalAxisTitle [gallery] < ChartAxisTitles
-     ChartSecondaryVerticalAxisTitleOptionsDialog [button] < ChartAxisTitles
-   ChartDepthAxisTitle [gallery] < ChartAxisTitles
-     ChartDepthAxisTitleOptionsDialog [button] < ChartAxisTitles
- ChartLegend [gallery]
-   ChartLegendOptionsDialogDialog [button] < ChartLegend
- ChartDataLabel [gallery]
-   ChartDataLabelDialog [button] < ChartDataLabel
- ChartDataTable [gallery]
-   ChartDataTableOptionsDialog [button] < ChartDataTable
### GroupChartAxes
- ChartAxes [menu]
-   ChartPrimaryHorizontalAxis [gallery] < ChartAxes
-     ChartPrimaryHorizontalAxisOptionsDialog [button] < ChartAxes
-   ChartPrimaryVerticalAxis [gallery] < ChartAxes
-     ChartPrimaryVerticalAxisOptionsDialog [button] < ChartAxes
-   ChartSecondaryHorizontalAxis [gallery] < ChartAxes
-     ChartSecondaryHorizontalAxisOption [button] < ChartAxes
-   ChartSecondaryVerticalAxis [gallery] < ChartAxes
-     ChartSecondaryVerticalAxisOptionsDialog [button] < ChartAxes
-   ChartDepthAxis [gallery] < ChartAxes
-     ChartDepthAxisOptionsDialog [button] < ChartAxes
- ChartGridlines [menu]
-   ChartPrimaryHorizontalGridlines [gallery] < ChartGridlines
-     ChartPrimaryHorizontalGridlinesOptionsDialog [button] < ChartGridlines
-   ChartPrimaryVerticalGridlines [gallery] < ChartGridlines
-     ChartPrimaryVerticalGridlinesOptionsDialog [button] < ChartGridlines
-   ChartSecondaryHorizontalGridlines [gallery] < ChartGridlines
-     ChartSecondaryHorizontalGridlinesOptionsDialog [button] < ChartGridlines
-   ChartSecondaryVerticalGridlines [gallery] < ChartGridlines
-     ChartSecondaryVerticalGridlinesOptionsDialog [button] < ChartGridlines
-   ChartDepthGridlines [gallery] < ChartGridlines
-     ChartDepthGridlinesOptionsDialog [button] < ChartGridlines
### GroupChartBackground
- ChartPlotArea [gallery]
-   ChartPlotAreaOptionsDialog [button] < ChartPlotArea
- ChartWall [gallery]
-   ChartWallOptionsDialog [button] < ChartWall
- ChartFloor [gallery]
-   ChartFloorOptionsDialog [button] < ChartFloor
- Chart3DView [button]
### GroupChartAnalysis
- ChartTrendline [gallery]
-   ChartTrendlineOptionsDialog [button] < ChartTrendline
- ChartLines [gallery]
- ChartUpDownBars [gallery]
-   ChartUpDownBarsOptionsDialog [button] < ChartUpDownBars
- ChartErrorBars [gallery]
-   ChartErrorBarsOptionsDialog [button] < ChartErrorBars
### GroupChartProperties
- ChartNameLabel [labelControl]
- ChartName [control]

## TabChartToolsFormat  (TabSetChartTools)
### 
- GroupChartCurrentSelection [group]
- GroupShapeStyles [group]
- GroupWordArtStyles [group]
- GroupArrange [group]
- GroupSize [group]
### GroupChartCurrentSelection
- ChartElementSelector [comboBox]
- ChartFormatSelection [button]
- ChartResetToMatchStyle [button]
### GroupShapeStyles
- ShapeStylesGallery [gallery]
-   ShapeStylesOtherThemeFillsGallery [gallery] < ShapeStylesGallery
- ShapeFillColorPicker [gallery]
-   ObjectFillMoreColorsDialog [button] < ShapeFillColorPicker
-   ObjectPictureFill [button] < ShapeFillColorPicker
-   GradientGallery [gallery] < ShapeFillColorPicker
-     ShapeFillMoreGradientsDialog [button] < ShapeFillColorPicker
-   ShapeFillTextureGallery [gallery] < ShapeFillColorPicker
-     MoreTextureOptions [button] < ShapeFillColorPicker
- ShapeOutlineColorPicker [gallery]
-   ObjectBorderOutlineColorMoreColorsDialog [button] < ShapeOutlineColorPicker
-   OutlineWeightGallery [gallery] < ShapeOutlineColorPicker
-     LineStylesDialog [button] < ShapeOutlineColorPicker
-   GallerySketchyShapes [gallery] < ShapeOutlineColorPicker
-     MoreLines2 [button] < ShapeOutlineColorPicker
-   OutlineDashesGallery [gallery] < ShapeOutlineColorPicker
-     LineStylesDialog [button] < ShapeOutlineColorPicker
-   ArrowStyleGallery [gallery] < ShapeOutlineColorPicker
-     ArrowsMore [button] < ShapeOutlineColorPicker
- ShapeEffectsMenu [menu]
-   ObjectEffectPresetGallery [gallery] < ShapeEffectsMenu
-     _3DBevelOptionsDialog [button] < ShapeEffectsMenu
-   ObjectEffectShadowGallery [gallery] < ShapeEffectsMenu
-     ShadowOptionsDialog [button] < ShapeEffectsMenu
-   ReflectionGallery [gallery] < ShapeEffectsMenu
-     ReflectionsMoreOptions [button] < ShapeEffectsMenu
-   ObjectEffectGlowGallery [gallery] < ShapeEffectsMenu
-     GlowColorPicker [gallery] < ShapeEffectsMenu
-     GlowColorMoreColorsDialog [button] < ShapeEffectsMenu
-     GlowsMoreOptions [button] < ShapeEffectsMenu
-   ObjectEffectSoftEdgesGallery [gallery] < ShapeEffectsMenu
-     SoftEdgesMoreOptions [button] < ShapeEffectsMenu
-   BevelShapeGallery [gallery] < ShapeEffectsMenu
-     _3DBevelOptionsDialog [button] < ShapeEffectsMenu
-   _3DRotationGallery [gallery] < ShapeEffectsMenu
-     _3DRotationOptionsDialog [button] < ShapeEffectsMenu
- ObjectFormatDialog [button (dialogBoxLauncher)]
### GroupWordArtStyles
- TextStylesGallery [gallery]
-   WordArtClear [button] < TextStylesGallery
- TextFillColorPicker [gallery]
-   TextFillColorMoreColorsDialog [button] < TextFillColorPicker
-   TextPictureFill [button] < TextFillColorPicker
-   TextFillGradientGallery [gallery] < TextFillColorPicker
-     TextFillMoreGradientsDialog [button] < TextFillColorPicker
-   TextFillTextureGallery [gallery] < TextFillColorPicker
-     TextFillMoreTextures [button] < TextFillColorPicker
- TextOutlineColorPicker [gallery]
-   TextOutlineColorMoreColorsDialog [button] < TextOutlineColorPicker
-   TextOutlineWeightGallery [gallery] < TextOutlineColorPicker
-     TextOutlineMoreLinesDialog [button] < TextOutlineColorPicker
-   TextOutlineDashesGallery [gallery] < TextOutlineColorPicker
-     TextOutlineMoreLinesDialog [button] < TextOutlineColorPicker
- TextEffectsMenu [menu]
-   TextEffectShadowGallery [gallery] < TextEffectsMenu
-     TextEffectsMoreShadowsDialog [button] < TextEffectsMenu
-   TextReflectionGallery [gallery] < TextEffectsMenu
-     TextEffectsReflectionsMoreOptions [button] < TextEffectsMenu
-   TextEffectGlowGallery [gallery] < TextEffectsMenu
-     TextGlowColorPicker [gallery] < TextEffectsMenu
-     TextGlowColorMoreColorsDialog [button] < TextEffectsMenu
-     TextEffectsGlowsMoreOptions [button] < TextEffectsMenu
-   BevelTextGallery [gallery] < TextEffectsMenu
-     TextEffectsBevelMore3DOptionsDialog [button] < TextEffectsMenu
-   TextEffects3DRotationGallery [gallery] < TextEffectsMenu
-     TextEffects3DRotationOptionsDialog [button] < TextEffectsMenu
-   TextEffectTransformGallery [gallery] < TextEffectsMenu
- WordArtFormatDialog [button (dialogBoxLauncher)]
### GroupArrange
- ObjectBringForwardMenu [splitButton]
-   ObjectBringForward [button] < ObjectBringForwardMenu
-   ObjectBringToFront [button] < ObjectBringForwardMenu
- ObjectSendBackwardMenu [splitButton]
-   ObjectSendBackward [button] < ObjectSendBackwardMenu
-   ObjectSendToBack [button] < ObjectSendBackwardMenu
- SelectionPane [toggleButton]
- ObjectAlignMenu [menu]
-   ObjectsAlignLeft [button] < ObjectAlignMenu
-   ObjectsAlignCenterHorizontal [button] < ObjectAlignMenu
-   ObjectsAlignRight [button] < ObjectAlignMenu
-   ObjectsAlignTop [button] < ObjectAlignMenu
-   ObjectsAlignMiddleVertical [button] < ObjectAlignMenu
-   ObjectsAlignBottom [button] < ObjectAlignMenu
-   AlignDistributeHorizontallyClassic [button] < ObjectAlignMenu
-   AlignDistributeVerticallyClassic [button] < ObjectAlignMenu
-   SnapToGrid [toggleButton] < ObjectAlignMenu
-   SnapToShapes [toggleButton] < ObjectAlignMenu
-   ViewGridlinesToggleExcel [toggleButton] < ObjectAlignMenu
- ObjectsGroupMenu [menu]
-   ObjectsGroup [button] < ObjectsGroupMenu
-   ObjectsRegroup [button] < ObjectsGroupMenu
-   ObjectsUngroup [button] < ObjectsGroupMenu
- ObjectRotateGallery [gallery]
-   ObjectRotationOptionsDialog [button] < ObjectRotateGallery
- ObjectsArrangeMenu [menu]
-   ObjectBringToFront [button] < ObjectsArrangeMenu
-   ObjectBringForward [button] < ObjectsArrangeMenu
-   ObjectSendToBack [button] < ObjectsArrangeMenu
-   ObjectSendBackward [button] < ObjectsArrangeMenu
-   ObjectsGroup [button] < ObjectsArrangeMenu
-   ObjectsUngroup [button] < ObjectsArrangeMenu
-   ObjectsRegroup [button] < ObjectsArrangeMenu
-   ObjectAlignMenu [menu] < ObjectsArrangeMenu
-     ObjectsAlignLeft [button] < ObjectsArrangeMenu
-     ObjectsAlignCenterHorizontal [button] < ObjectsArrangeMenu
-     ObjectsAlignRight [button] < ObjectsArrangeMenu
-     ObjectsAlignTop [button] < ObjectsArrangeMenu
-     ObjectsAlignMiddleVertical [button] < ObjectsArrangeMenu
-     ObjectsAlignBottom [button] < ObjectsArrangeMenu
-     AlignDistributeHorizontallyClassic [button] < ObjectsArrangeMenu
-     AlignDistributeVerticallyClassic [button] < ObjectsArrangeMenu
-     SnapToGrid [toggleButton] < ObjectsArrangeMenu
-     SnapToShapes [toggleButton] < ObjectsArrangeMenu
-     ViewGridlinesToggleExcel [toggleButton] < ObjectsArrangeMenu
-   ObjectRotateGallery [gallery] < ObjectsArrangeMenu
-     ObjectRotationOptionsDialog [button] < ObjectsArrangeMenu
### GroupSize
- ShapeHeight [control]
- ShapeWidth [control]
- ObjectSizeAndPropertiesDialog [button (dialogBoxLauncher)]

##   (TabSetDrawingTools)
### 
- TabDrawingToolsFormat [tab]

## TabDrawingToolsFormat  (TabSetDrawingTools)
### 
- GroupShapes [group]
- GroupShapeStyles [group]
- GroupWordArtStyles [group]
- GroupAltText [group]
- GroupArrangeWith3DEditor [group]
- GroupSize [group]
### GroupShapes
- ShapesInsertGallery [gallery]
- ObjectEditShapeMenu [menu]
-   ShapeChangeShapeGallery [gallery] < ObjectEditShapeMenu
-   ObjectEditPoints [toggleButton] < ObjectEditShapeMenu
-   ShapeRerouteConnectors [toggleButton] < ObjectEditShapeMenu
- TextBoxInsertExcel [toggleButton]
- TextBoxInsertMenu [splitButton]
-   TextBoxInsertHorizontal [toggleButton] < TextBoxInsertMenu
-   TextBoxInsertVertical [toggleButton] < TextBoxInsertMenu
- TextboxFlyoutMenuAnchor [menu]
-   TextBoxInsertHorizontal [toggleButton] < TextboxFlyoutMenuAnchor
-   TextBoxInsertVertical [toggleButton] < TextboxFlyoutMenuAnchor
### GroupShapeStyles
- ShapeStylesGallery [gallery]
-   ShapeStylesOtherThemeFillsGallery [gallery] < ShapeStylesGallery
- ShapeFillColorPicker [gallery]
-   ObjectFillMoreColorsDialog [button] < ShapeFillColorPicker
-   ObjectPictureFill [button] < ShapeFillColorPicker
-   GradientGallery [gallery] < ShapeFillColorPicker
-     ShapeFillMoreGradientsDialog [button] < ShapeFillColorPicker
-   ShapeFillTextureGallery [gallery] < ShapeFillColorPicker
-     MoreTextureOptions [button] < ShapeFillColorPicker
- ShapeOutlineColorPicker [gallery]
-   ObjectBorderOutlineColorMoreColorsDialog [button] < ShapeOutlineColorPicker
-   OutlineWeightGallery [gallery] < ShapeOutlineColorPicker
-     LineStylesDialog [button] < ShapeOutlineColorPicker
-   GallerySketchyShapes [gallery] < ShapeOutlineColorPicker
-     MoreLines2 [button] < ShapeOutlineColorPicker
-   OutlineDashesGallery [gallery] < ShapeOutlineColorPicker
-     LineStylesDialog [button] < ShapeOutlineColorPicker
-   ArrowStyleGallery [gallery] < ShapeOutlineColorPicker
-     ArrowsMore [button] < ShapeOutlineColorPicker
- ShapeEffectsMenu [menu]
-   ObjectEffectPresetGallery [gallery] < ShapeEffectsMenu
-     _3DBevelOptionsDialog [button] < ShapeEffectsMenu
-   ObjectEffectShadowGallery [gallery] < ShapeEffectsMenu
-     ShadowOptionsDialog [button] < ShapeEffectsMenu
-   ReflectionGallery [gallery] < ShapeEffectsMenu
-     ReflectionsMoreOptions [button] < ShapeEffectsMenu
-   ObjectEffectGlowGallery [gallery] < ShapeEffectsMenu
-     GlowColorPicker [gallery] < ShapeEffectsMenu
-     GlowColorMoreColorsDialog [button] < ShapeEffectsMenu
-     GlowsMoreOptions [button] < ShapeEffectsMenu
-   ObjectEffectSoftEdgesGallery [gallery] < ShapeEffectsMenu
-     SoftEdgesMoreOptions [button] < ShapeEffectsMenu
-   BevelShapeGallery [gallery] < ShapeEffectsMenu
-     _3DBevelOptionsDialog [button] < ShapeEffectsMenu
-   _3DRotationGallery [gallery] < ShapeEffectsMenu
-     _3DRotationOptionsDialog [button] < ShapeEffectsMenu
- ObjectFormatDialog [button (dialogBoxLauncher)]
### GroupWordArtStyles
- TextStylesGallery [gallery]
-   WordArtClear [button] < TextStylesGallery
- TextFillColorPicker [gallery]
-   TextFillColorMoreColorsDialog [button] < TextFillColorPicker
-   TextPictureFill [button] < TextFillColorPicker
-   TextFillGradientGallery [gallery] < TextFillColorPicker
-     TextFillMoreGradientsDialog [button] < TextFillColorPicker
-   TextFillTextureGallery [gallery] < TextFillColorPicker
-     TextFillMoreTextures [button] < TextFillColorPicker
- TextOutlineColorPicker [gallery]
-   TextOutlineColorMoreColorsDialog [button] < TextOutlineColorPicker
-   TextOutlineWeightGallery [gallery] < TextOutlineColorPicker
-     TextOutlineMoreLinesDialog [button] < TextOutlineColorPicker
-   TextOutlineDashesGallery [gallery] < TextOutlineColorPicker
-     TextOutlineMoreLinesDialog [button] < TextOutlineColorPicker
- TextEffectsMenu [menu]
-   TextEffectShadowGallery [gallery] < TextEffectsMenu
-     TextEffectsMoreShadowsDialog [button] < TextEffectsMenu
-   TextReflectionGallery [gallery] < TextEffectsMenu
-     TextEffectsReflectionsMoreOptions [button] < TextEffectsMenu
-   TextEffectGlowGallery [gallery] < TextEffectsMenu
-     TextGlowColorPicker [gallery] < TextEffectsMenu
-     TextGlowColorMoreColorsDialog [button] < TextEffectsMenu
-     TextEffectsGlowsMoreOptions [button] < TextEffectsMenu
-   BevelTextGallery [gallery] < TextEffectsMenu
-     TextEffectsBevelMore3DOptionsDialog [button] < TextEffectsMenu
-   TextEffects3DRotationGallery [gallery] < TextEffectsMenu
-     TextEffects3DRotationOptionsDialog [button] < TextEffectsMenu
-   TextEffectTransformGallery [gallery] < TextEffectsMenu
- WordArtFormatDialog [button (dialogBoxLauncher)]
### GroupAltText
- AltTextPaneRibbon [toggleButton]
### GroupArrangeWith3DEditor
- ObjectBringForwardMenu [splitButton]
-   ObjectBringForward [button] < ObjectBringForwardMenu
-   ObjectBringToFront [button] < ObjectBringForwardMenu
- ObjectSendBackwardMenu [splitButton]
-   ObjectSendBackward [button] < ObjectSendBackwardMenu
-   ObjectSendToBack [button] < ObjectSendBackwardMenu
- SelectionPane [toggleButton]
- ObjectAlignMenu [menu]
-   ObjectsAlignLeft [button] < ObjectAlignMenu
-   ObjectsAlignCenterHorizontal [button] < ObjectAlignMenu
-   ObjectsAlignRight [button] < ObjectAlignMenu
-   ObjectsAlignTop [button] < ObjectAlignMenu
-   ObjectsAlignMiddleVertical [button] < ObjectAlignMenu
-   ObjectsAlignBottom [button] < ObjectAlignMenu
-   AlignDistributeHorizontallyClassic [button] < ObjectAlignMenu
-   AlignDistributeVerticallyClassic [button] < ObjectAlignMenu
-   SnapToGrid [toggleButton] < ObjectAlignMenu
-   SnapToShapes [toggleButton] < ObjectAlignMenu
-   ViewGridlinesToggleExcel [toggleButton] < ObjectAlignMenu
- ObjectsGroupMenu [menu]
-   ObjectsGroup [button] < ObjectsGroupMenu
-   ObjectsRegroup [button] < ObjectsGroupMenu
-   ObjectsUngroup [button] < ObjectsGroupMenu
- ObjectRotateGallery [gallery]
-   ObjectRotationOptionsDialog [button] < ObjectRotateGallery
- ObjectsArrangeMenu [menu]
-   ObjectBringToFront [button] < ObjectsArrangeMenu
-   ObjectBringForward [button] < ObjectsArrangeMenu
-   ObjectSendToBack [button] < ObjectsArrangeMenu
-   ObjectSendBackward [button] < ObjectsArrangeMenu
-   ObjectsGroup [button] < ObjectsArrangeMenu
-   ObjectsUngroup [button] < ObjectsArrangeMenu
-   ObjectsRegroup [button] < ObjectsArrangeMenu
-   ObjectAlignMenu [menu] < ObjectsArrangeMenu
-     ObjectsAlignLeft [button] < ObjectsArrangeMenu
-     ObjectsAlignCenterHorizontal [button] < ObjectsArrangeMenu
-     ObjectsAlignRight [button] < ObjectsArrangeMenu
-     ObjectsAlignTop [button] < ObjectsArrangeMenu
-     ObjectsAlignMiddleVertical [button] < ObjectsArrangeMenu
-     ObjectsAlignBottom [button] < ObjectsArrangeMenu
-     AlignDistributeHorizontallyClassic [button] < ObjectsArrangeMenu
-     AlignDistributeVerticallyClassic [button] < ObjectsArrangeMenu
-     SnapToGrid [toggleButton] < ObjectsArrangeMenu
-     SnapToShapes [toggleButton] < ObjectsArrangeMenu
-     ViewGridlinesToggleExcel [toggleButton] < ObjectsArrangeMenu
-   ObjectRotateGallery [gallery] < ObjectsArrangeMenu
-     ObjectRotationOptionsDialog [button] < ObjectsArrangeMenu
### GroupSize
- ShapeHeight [control]
- ShapeWidth [control]
- ObjectSizeAndPropertiesDialog [button (dialogBoxLauncher)]

##   (TabSetPictureTools)
### 
- TabPictureToolsFormat [tab]

## TabPictureToolsFormat  (TabSetPictureTools)
### 
- GroupImagePlay [group]
- GroupPictureTools [group]
- GroupPictureStyles [group]
- GroupAltText [group]
- GroupArrangeWith3DEditor [group]
- GroupPictureSize [group]
### GroupImagePlay
- ImagePlay [button]
### GroupPictureTools
- PicturePlaceInCell [button]
- PictureBackgroundRemoval [toggleButton]
- PictureCorrectionsMenu [menu]
-   PictureSharpenSoftenGallery [gallery] < PictureCorrectionsMenu
-   PictureBrightnessAndContrastGallery [gallery] < PictureCorrectionsMenu
-   PictureCorrectionsDialog [button] < PictureCorrectionsMenu
- PictureColorMenu [menu]
-   PictureSaturationGallery [gallery] < PictureColorMenu
-   PictureColorTempertatureGallery [gallery] < PictureColorMenu
-   PictureRecolorGallery [gallery] < PictureColorMenu
-   RecolorColorPicker [gallery] < PictureColorMenu
-     PictureRecolorMoreColorsDialog [button] < PictureColorMenu
-   PictureSetTransparentColor [toggleButton] < PictureColorMenu
-   PictureColorDialog [button] < PictureColorMenu
- PictureArtisticEffectsGallery [gallery]
-   ArtisticEffectsDialog [button] < PictureArtisticEffectsGallery
- PictureTransparencyGallery [gallery]
-   PictureMoreTransparency [button] < PictureTransparencyGallery
- PicturesCompress [button]
- PictureChange [button]
- PictureChangeMenu [menu]
-   PictureChangeFromFile [button] < PictureChangeMenu
-   PictureChangeFromMobile [button] < PictureChangeMenu
-   PictureChangeFromM365 [button] < PictureChangeMenu
-   PictureChangeFromOnline [button] < PictureChangeMenu
-   ChangeIcon [button] < PictureChangeMenu
-   PictureChangeFromClipboard [button] < PictureChangeMenu
- PictureResetSplitButton [splitButton]
-   PictureReset [button] < PictureResetSplitButton
-   PictureResetAndSize [button] < PictureResetSplitButton
### GroupPictureStyles
- PictureStylesGallery [gallery]
- OutlineColorPicker [gallery]
-   ObjectBorderOutlineColorMoreColorsDialog [button] < OutlineColorPicker
-   OutlineWeightGallery [gallery] < OutlineColorPicker
-     LineStylesDialog [button] < OutlineColorPicker
-   GallerySketchyShapes [gallery] < OutlineColorPicker
-     MoreLines2 [button] < OutlineColorPicker
-   OutlineDashesGallery [gallery] < OutlineColorPicker
-     LineStylesDialog [button] < OutlineColorPicker
- PictureEffectsMenu [menu]
-   PictureEffectsPresetGallery [gallery] < PictureEffectsMenu
-     _3DBevelOptionsDialog [button] < PictureEffectsMenu
-   PictureEffectsShadowGallery [gallery] < PictureEffectsMenu
-     ShadowOptionsDialog [button] < PictureEffectsMenu
-   PictureReflectionGallery [gallery] < PictureEffectsMenu
-     ReflectionsMoreOptions [button] < PictureEffectsMenu
-   PictureEffectsGlowGallery [gallery] < PictureEffectsMenu
-     GlowColorPicker [gallery] < PictureEffectsMenu
-     GlowColorMoreColorsDialog [button] < PictureEffectsMenu
-     GlowsMoreOptions [button] < PictureEffectsMenu
-   PictureEffectsSoftEdgesGallery [gallery] < PictureEffectsMenu
-     SoftEdgesMoreOptions [button] < PictureEffectsMenu
-   _3DBevelPictureTopGallery [gallery] < PictureEffectsMenu
-     _3DBevelOptionsDialog [button] < PictureEffectsMenu
-   PictureRotationGallery [gallery] < PictureEffectsMenu
-     _3DRotationOptionsDialog [button] < PictureEffectsMenu
- PicturesConvertToSmartArt [gallery]
- PictureFormatDialog [button (dialogBoxLauncher)]
### GroupAltText
- AltTextPaneRibbon [toggleButton]
### GroupArrangeWith3DEditor
- ObjectBringForwardMenu [splitButton]
-   ObjectBringForward [button] < ObjectBringForwardMenu
-   ObjectBringToFront [button] < ObjectBringForwardMenu
- ObjectSendBackwardMenu [splitButton]
-   ObjectSendBackward [button] < ObjectSendBackwardMenu
-   ObjectSendToBack [button] < ObjectSendBackwardMenu
- SelectionPane [toggleButton]
- ObjectAlignMenu [menu]
-   ObjectsAlignLeft [button] < ObjectAlignMenu
-   ObjectsAlignCenterHorizontal [button] < ObjectAlignMenu
-   ObjectsAlignRight [button] < ObjectAlignMenu
-   ObjectsAlignTop [button] < ObjectAlignMenu
-   ObjectsAlignMiddleVertical [button] < ObjectAlignMenu
-   ObjectsAlignBottom [button] < ObjectAlignMenu
-   AlignDistributeHorizontallyClassic [button] < ObjectAlignMenu
-   AlignDistributeVerticallyClassic [button] < ObjectAlignMenu
-   SnapToGrid [toggleButton] < ObjectAlignMenu
-   SnapToShapes [toggleButton] < ObjectAlignMenu
-   ViewGridlinesToggleExcel [toggleButton] < ObjectAlignMenu
- ObjectsGroupMenu [menu]
-   ObjectsGroup [button] < ObjectsGroupMenu
-   ObjectsRegroup [button] < ObjectsGroupMenu
-   ObjectsUngroup [button] < ObjectsGroupMenu
- ObjectRotateGallery [gallery]
-   ObjectRotationOptionsDialog [button] < ObjectRotateGallery
- ObjectsArrangeMenu [menu]
-   ObjectBringToFront [button] < ObjectsArrangeMenu
-   ObjectBringForward [button] < ObjectsArrangeMenu
-   ObjectSendToBack [button] < ObjectsArrangeMenu
-   ObjectSendBackward [button] < ObjectsArrangeMenu
-   ObjectsGroup [button] < ObjectsArrangeMenu
-   ObjectsUngroup [button] < ObjectsArrangeMenu
-   ObjectsRegroup [button] < ObjectsArrangeMenu
-   ObjectAlignMenu [menu] < ObjectsArrangeMenu
-     ObjectsAlignLeft [button] < ObjectsArrangeMenu
-     ObjectsAlignCenterHorizontal [button] < ObjectsArrangeMenu
-     ObjectsAlignRight [button] < ObjectsArrangeMenu
-     ObjectsAlignTop [button] < ObjectsArrangeMenu
-     ObjectsAlignMiddleVertical [button] < ObjectsArrangeMenu
-     ObjectsAlignBottom [button] < ObjectsArrangeMenu
-     AlignDistributeHorizontallyClassic [button] < ObjectsArrangeMenu
-     AlignDistributeVerticallyClassic [button] < ObjectsArrangeMenu
-     SnapToGrid [toggleButton] < ObjectsArrangeMenu
-     SnapToShapes [toggleButton] < ObjectsArrangeMenu
-     ViewGridlinesToggleExcel [toggleButton] < ObjectsArrangeMenu
-   ObjectRotateGallery [gallery] < ObjectsArrangeMenu
-     ObjectRotationOptionsDialog [button] < ObjectsArrangeMenu
### GroupPictureSize
- PictureCropTools [splitButton]
-   PictureCrop [toggleButton] < PictureCropTools
-   PictureShapeGallery [gallery] < PictureCropTools
-   PictureCropAspectRatioMenu [menu] < PictureCropTools
-     PictureCropAspectRatio1To1 [button] < PictureCropTools
-     PictureCropAspectRatio2To3 [button] < PictureCropTools
-     PictureCropAspectRatio3To4 [button] < PictureCropTools
-     PictureCropAspectRatio3To5 [button] < PictureCropTools
-     PictureCropAspectRatio4To5 [button] < PictureCropTools
-     PictureCropAspectRatio3To2 [button] < PictureCropTools
-     PictureCropAspectRatio4To3 [button] < PictureCropTools
-     PictureCropAspectRatio5To3 [button] < PictureCropTools
-     PictureCropAspectRatio5To4 [button] < PictureCropTools
-     PictureCropAspectRatio16To9 [button] < PictureCropTools
-     PictureCropAspectRatio16To10 [button] < PictureCropTools
-   PictureFillCrop [button] < PictureCropTools
-   PictureFitCrop [button] < PictureCropTools
- ShapeHeight [control]
- ShapeWidth [control]
- ObjectSizeAndPropertiesDialog [button (dialogBoxLauncher)]

##   (TabSetPivotTableTools)
### 
- TabPivotTableToolsOptions [tab]
- TabPivotTableToolsDesign [tab]
- TabPivotTableToolsDesignDeprecated [tab]

## TabPivotTableToolsOptions  (TabSetPivotTableTools)
### 
- GroupPivotTableOptions [group]
- GroupPivotTableActiveField [group]
- GroupPivotTableGroup [group]
- GroupPivotTableSort [group]
- GroupPivotTableData [group]
- GroupPivotActions [group]
- GroupPivotTableCalculations [group]
- GroupPivotTableTools [group]
- GroupPivotTableShowHide [group]
### GroupPivotTableOptions
- PivotTableNameLabel [labelControl]
- PivotTableName [control]
- PivotTableOptionsMenu [splitButton]
-   PivotTableOptions [button] < PivotTableOptionsMenu
-   PivotTableShowPages [button] < PivotTableOptionsMenu
-   PivotTableGenerateGetPivotData [checkBox] < PivotTableOptionsMenu
### GroupPivotTableActiveField
- LabelActiveField [labelControl]
- PivotFieldName [control]
- PivotTableFieldSettings [button]
- SelectedItemNavDrillDown [button]
- SelectedItemNavDrillUpSplitButton [splitButton]
-   SelectedItemNavDrillUp [button] < SelectedItemNavDrillUpSplitButton
-   None [button] < SelectedItemNavDrillUpSplitButton
- PivotTableExpandField [button]
- PivotCollapseField [button]
- PivotTableShowDetails [button]
### GroupPivotTableGroup
- PivotTableGroupSelection [button]
- OutlineUngroup [button]
- PivotTableGroupField [button]
### GroupPivotTableSort
- SlicerCreate [button]
- TimeSlicerCreate [button]
- SlicerShare [button]
### GroupPivotTableData
- RefreshMenu [splitButton]
-   Refresh [button] < RefreshMenu
-   RefreshAll [button] < RefreshMenu
-   RefreshStatus [button] < RefreshMenu
-   RefreshCancel [button] < RefreshMenu
-   ConnectionProperties [button] < RefreshMenu
- PivotTableEditDataSource [splitButton]
-   PivotTableChangeDataSource [button] < PivotTableEditDataSource
-   ConnectionProperties [button] < PivotTableEditDataSource
### GroupPivotActions
- PivotTableClearMenu [menu]
-   PivotClearAll [button] < PivotTableClearMenu
-   PivotClearFilters [button] < PivotTableClearMenu
- PivotTableSelectFlyout [menu]
-   PivotTableSelectLabelAndData [button] < PivotTableSelectFlyout
-   PivotTableSelectData [button] < PivotTableSelectFlyout
-   PivotTableSelectLabel [button] < PivotTableSelectFlyout
-   PivotTableSelectEntireTable [button] < PivotTableSelectFlyout
-   PivotTableEnableSelection [toggleButton] < PivotTableSelectFlyout
- PivotTableMove [button]
### GroupPivotTableCalculations
- PivotTableFormulasMenu [menu]
-   PivotTableFieldInsert [button] < PivotTableFormulasMenu
-   PivotTableCalculatedItem [button] < PivotTableFormulasMenu
-   PivotTableSolveOrder [button] < PivotTableFormulasMenu
-   PivotTableListFormulas [button] < PivotTableFormulasMenu
-   PivotTableCreateTableBasedOnRow [button] < PivotTableFormulasMenu
-   PivotTableCreateTableBasedOnColumn [button] < PivotTableFormulasMenu
-   NamedSetManager [button] < PivotTableFormulasMenu
- PivotTableOlapTools [menu]
-   PivotTableOlapPropertyFields [button] < PivotTableOlapTools
-   PivotTableOlapOffline [button] < PivotTableOlapTools
-   PivotTableOlapConvertToFormulas [button] < PivotTableOlapTools
-   PivotTableWhatIfAnalysisMenu [menu] < PivotTableOlapTools
-     PivotTableEnableWhatIfAnalysis [toggleButton] < PivotTableOlapTools
-     PivotTableCalculateChanges [button] < PivotTableOlapTools
-     PivotTableAutoCalculate [toggleButton] < PivotTableOlapTools
-     PivotTableDiscardChanges [button] < PivotTableOlapTools
-     PivotTablePublishChanges [button] < PivotTableOlapTools
-     PivotTableRefreshDataSourceValue [button] < PivotTableOlapTools
-     PivotTableWhatIfAnalysisSettings [button] < PivotTableOlapTools
-   PivotTableCreateMeasures [button] < PivotTableOlapTools
-   PivotTableCreateMembers [button] < PivotTableOlapTools
-   PivotTableManageCalculations [button] < PivotTableOlapTools
- AsppRelationships [button]
### GroupPivotTableTools
- PivotChartInsertClassic [button]
- SuggestPivot [button]
### GroupPivotTableShowHide
- PivotFieldListShowHide [toggleButton]
- PivotPlusMinusButtonsShowHide [toggleButton]
- PivotPlusMinusFieldHeadersShowHide [toggleButton]

## TabPivotTableToolsDesign  (TabSetPivotTableTools)
### 
- GroupPivotTableLayout [group]
- GroupPivotTableStyleOptions [group]
- GroupPivotTableStyles [group]
### GroupPivotTableLayout
- PivotTableLayoutSubtotals [menu]
-   PivotTableSubtotalsDoNotShow [button] < PivotTableLayoutSubtotals
-   PivotTableSubtotalsOnBottom [button] < PivotTableLayoutSubtotals
-   PivotTableSubtotalsOnTop [button] < PivotTableLayoutSubtotals
-   PivotTableNonVisualTotals [toggleButton] < PivotTableLayoutSubtotals
- PivotTableLayoutGrandTotals [menu]
-   PivotTableGrandTotalsOffForRowsAndColumns [button] < PivotTableLayoutGrandTotals
-   PivotTableGrandTotalsOnForRowsAndColumns [button] < PivotTableLayoutGrandTotals
-   PivotTableGrandTotalsOnForRowsOnly [button] < PivotTableLayoutGrandTotals
-   PivotTableGrandTotalsOnForColumnsOnly [button] < PivotTableLayoutGrandTotals
- PivotTableLayoutReportLayout [menu]
-   PivotTableLayoutShowInCompactForm [button] < PivotTableLayoutReportLayout
-   PivotTableLayoutShowInOutlineForm [button] < PivotTableLayoutReportLayout
-   PivotTableLayoutShowInTabularForm [button] < PivotTableLayoutReportLayout
-   PivotTableRepeatItemLabels [button] < PivotTableLayoutReportLayout
-   PivotTableDoNotRepeatItemLabels [button] < PivotTableLayoutReportLayout
- PivotTableLayoutBlankRows [menu]
-   PivotTableBlankRowsInsert [button] < PivotTableLayoutBlankRows
-   PivotTableBlankRowsRemove [button] < PivotTableLayoutBlankRows
### GroupPivotTableStyleOptions
- TableStyleRowHeaders [checkBox]
- TableStyleColumnHeaders [checkBox]
- TableStyleBandedRows [checkBox]
- TableStyleBandedColumns [checkBox]
### GroupPivotTableStyles
- PivotTableStylesGallery [gallery]
-   PivotTableNewStyle [button] < PivotTableStylesGallery
-   TableStyleClear [button] < PivotTableStylesGallery

## TabPivotTableToolsDesignDeprecated  (TabSetPivotTableTools)
### 
- GroupPivotTableCalculationsDeprecated [group]
### GroupPivotTableCalculationsDeprecated
- PivotTableSummarizeByMenu [menu]
-   PivotTableSummarizeBySum [toggleButton] < PivotTableSummarizeByMenu
-   PivotTableSummarizeByCount [toggleButton] < PivotTableSummarizeByMenu
-   PivotTableSummarizeByAverage [toggleButton] < PivotTableSummarizeByMenu
-   PivotTableSummarizeByMaximum [toggleButton] < PivotTableSummarizeByMenu
-   PivotTableSummarizeByMinimum [toggleButton] < PivotTableSummarizeByMenu
-   PivotTableSummarizeByProduct [toggleButton] < PivotTableSummarizeByMenu
-   PivotTableSummarizeByMore [button] < PivotTableSummarizeByMenu
- PivotTableShowAsMenu [menu]
-   PivotTableShowAs [toggleButton] < PivotTableShowAsMenu
-   PivotTableShowAsPercentOfTotal [toggleButton] < PivotTableShowAsMenu
-   PivotTableShowAsPercentOfColumn [toggleButton] < PivotTableShowAsMenu
-   PivotTableShowAsPercentOfRow [toggleButton] < PivotTableShowAsMenu
-   PivotTableShowAsPercentOf [toggleButton] < PivotTableShowAsMenu
-   PivotTableShowAsPercentOfParentRow [toggleButton] < PivotTableShowAsMenu
-   PivotTableShowAsPercentOfParentColumn [toggleButton] < PivotTableShowAsMenu
-   PivotTableShowAsPercentOfParent [toggleButton] < PivotTableShowAsMenu
-   PivotTableShowAsDifference [toggleButton] < PivotTableShowAsMenu
-   PivotTableShowAsPercentDifference [toggleButton] < PivotTableShowAsMenu
-   PivotTableShowAsRunningTotal [toggleButton] < PivotTableShowAsMenu
-   PivotTableShowAsPercentRunningTotal [toggleButton] < PivotTableShowAsMenu
-   PivotTableShowAsRangAscending [toggleButton] < PivotTableShowAsMenu
-   PivotTableShowAsRankDescending [toggleButton] < PivotTableShowAsMenu
-   PivotTableShowAsIndex [toggleButton] < PivotTableShowAsMenu
-   PivotTableShowAsMore [button] < PivotTableShowAsMenu

##   (TabSetHeaderAndFooterTools)
### 
- TabHeaderAndFooterToolsDesign [tab]

## TabHeaderAndFooterToolsDesign  (TabSetHeaderAndFooterTools)
### 
- GroupHeaderFooter [group]
- GroupHeaderFooterElements [group]
- GroupHeaderFooterNavigation [group]
- GroupHeaderFooterOptions [group]
### GroupHeaderFooter
- HeaderFooterHeaderGallery [gallery]
- HeaderFooterFooterGallery [gallery]
### GroupHeaderFooterElements
- HeaderFooterPageNumberInsertExcel [button]
- HeaderFooterNumberOfPagesInsert [button]
- HeaderFooterCurrentDate [button]
- HeaderFooterCurrentTimeInsert [button]
- HeaderFooterFilePathInsert [button]
- HeaderFooterFileNameInsert [button]
- HeaderFooterSheetNameInsert [button]
- HeaderFooterPictureInsert [button]
- HeaderFooterFormatPicture [button]
### GroupHeaderFooterNavigation
- GoToHeader [button]
- GoToFooter [button]
### GroupHeaderFooterOptions
- HeaderFooterDifferentFirstPageExcel [checkBox]
- HeaderFooterDifferentOddEvenPageExcel [checkBox]
- HeaderFooterScaleWithDocument [checkBox]
- HeaderFooterAlignMargins [checkBox]

##   (TabSetTableToolsExcel)
### 
- TabTableToolsDesignExcel [tab]

## TabTableToolsDesignExcel  (TabSetTableToolsExcel)
### 
- GroupTableProperties [group]
- GroupTableTools [group]
- GroupTableExternalData [group]
- GroupTableStyleOptions [group]
- GroupTableStylesExcel [group]
### GroupTableProperties
- TableNameLabel [labelControl]
- TableName [control]
- TableResize [button]
### GroupTableTools
- TableSummarizeWithPivot [button]
- RemoveDuplicates [button]
- TableConvertToRange [button]
- TableSlicerInsert [button]
### GroupTableExternalData
- TableExportMenu [menu]
-   TableExportTableToSharePointList [button] < TableExportMenu
-   TableExportTableToVisioPivotDiagram [button] < TableExportMenu
-   CreateVisioDataVisualizer [button] < TableExportMenu
- RefreshMenu [splitButton]
-   Refresh [button] < RefreshMenu
-   RefreshAll [button] < RefreshMenu
-   RefreshStatus [button] < RefreshMenu
-   RefreshCancel [button] < RefreshMenu
-   ConnectionProperties [button] < RefreshMenu
- DataRangeProperties [button]
- TableOpenInBrowser [button]
- TableUnlinkExternalData [button]
### GroupTableStyleOptions
- TableStyleHeaderRow [checkBox]
- TableStyleTotalsRow [checkBox]
- TableStyleBandedRows [checkBox]
- TableStyleFirstColumn [checkBox]
- TableStyleLastColumn [checkBox]
- TableStyleBandedColumns [checkBox]
- TableStyleFilterButton [checkBox]
### GroupTableStylesExcel
- TableStylesGalleryExcel [gallery]
-   TableStyleNew [button] < TableStylesGalleryExcel
-   TableStyleClear [button] < TableStylesGalleryExcel

##   (TabSetPivotChartTools)
### 
- TabPivotChartToolsAnalyze [tab]
- TabChartToolsDesignPivotChart [tab]
- TabChartToolsFormatPivotChart [tab]
- TabPivotChartToolsDesign [tab]
- TabPivotChartToolsLayout [tab]
- TabPivotChartToolsFormat [tab]

## TabPivotChartToolsAnalyze  (TabSetPivotChartTools)
### 
- GroupPivotChartOptions [group]
- GroupPivotChartActiveField [group]
- GroupPivotChartFilter [group]
- GroupPivotChartData [group]
- GroupPivotChartActions [group]
- GroupPivotChartCalculations [group]
- GroupPivotChartShowOrHide [group]
### GroupPivotChartOptions
- ChartNameLabel [labelControl]
- ChartName [control]
- PivotTableOptions [button]
### GroupPivotChartActiveField
- LabelActiveField [labelControl]
- PivotFieldName [control]
- PivotTableFieldSettings [button]
- SelectedItemNavDrillDown [button]
- SelectedItemNavDrillUpSplitButton [splitButton]
-   SelectedItemNavDrillUp [button] < SelectedItemNavDrillUpSplitButton
-   None [button] < SelectedItemNavDrillUpSplitButton
- PivotTableExpandField [button]
- PivotCollapseField [button]
### GroupPivotChartFilter
- SlicerCreate [button]
- TimeSlicerCreate [button]
- PivotChartSlicerShare [button]
### GroupPivotChartData
- RefreshMenu [splitButton]
-   Refresh [button] < RefreshMenu
-   RefreshAll [button] < RefreshMenu
-   RefreshStatus [button] < RefreshMenu
-   RefreshCancel [button] < RefreshMenu
-   ConnectionProperties [button] < RefreshMenu
- PivotTableEditDataSource [splitButton]
-   PivotTableChangeDataSource [button] < PivotTableEditDataSource
-   ConnectionProperties [button] < PivotTableEditDataSource
### GroupPivotChartActions
- PivotTableClearMenu [menu]
-   PivotClearAll [button] < PivotTableClearMenu
-   PivotClearFilters [button] < PivotTableClearMenu
- ChartPlacement [button]
### GroupPivotChartCalculations
- PivotChartFormulasMenu [menu]
-   PivotTableFieldInsert [button] < PivotChartFormulasMenu
-   PivotTableCalculatedItem [button] < PivotChartFormulasMenu
-   PivotTableSolveOrder [button] < PivotChartFormulasMenu
-   PivotTableListFormulas [button] < PivotChartFormulasMenu
-   PivotChartCreateTableBasedOnAxisMenu [button] < PivotChartFormulasMenu
-   PivotTableCreateTableBasedOnLegendMenu [button] < PivotChartFormulasMenu
-   NamedSetManager [button] < PivotChartFormulasMenu
- PivotChartOlapTools [menu]
-   PivotTableOlapPropertyFields [button] < PivotChartOlapTools
-   PivotTableOlapOffline [button] < PivotChartOlapTools
-   PivotTableCreateMeasures [button] < PivotChartOlapTools
-   PivotTableCreateMembers [button] < PivotChartOlapTools
-   PivotTableManageCalculations [button] < PivotChartOlapTools
- AsppRelationships [button]
### GroupPivotChartShowOrHide
- PivotFieldListShowHide [toggleButton]
- PivotChartShowFieldButtonsMenu [splitButton]
-   PivotChartShowFieldButtons [toggleButton] < PivotChartShowFieldButtonsMenu
-   PivotChartShowReportFilterFieldButtons [toggleButton] < PivotChartShowFieldButtonsMenu
-   PivotChartShowLegendFieldButtons [toggleButton] < PivotChartShowFieldButtonsMenu
-   PivotChartShowAxisFieldButtons [toggleButton] < PivotChartShowFieldButtonsMenu
-   PivotChartShowValueFieldButtons [toggleButton] < PivotChartShowFieldButtonsMenu
-   PivotChartShowExpandCollapseButtons [toggleButton] < PivotChartShowFieldButtonsMenu
-   PivotChartHideButtons [toggleButton] < PivotChartShowFieldButtonsMenu

## TabChartToolsDesignPivotChart  (TabSetPivotChartTools)
### 
- GroupChartLayouts [group]
- GroupChartStyles [group]
- GroupChartData [group]
- GroupChartType [group]
- GroupChartLocation [group]
### GroupChartLayouts
- AddChartElementMenu [menu]
-   ChartAxesGallery [gallery] < AddChartElementMenu
-     ChartAxesOptionsDialog [button] < AddChartElementMenu
-   ChartAxisTitlesGallery [gallery] < AddChartElementMenu
-     ChartAxesTitlesOptionsDialog [button] < AddChartElementMenu
-   ChartTitle [gallery] < AddChartElementMenu
-     ChartTitleOptionsDialog [button] < AddChartElementMenu
-   ChartDataLabel [gallery] < AddChartElementMenu
-     ChartDataLabelDialog [button] < AddChartElementMenu
-   ChartDataTable [gallery] < AddChartElementMenu
-     ChartDataTableOptionsDialog [button] < AddChartElementMenu
-   ChartErrorBars [gallery] < AddChartElementMenu
-     ChartErrorBarsOptionsDialog [button] < AddChartElementMenu
-   ChartGridlinesGallery [gallery] < AddChartElementMenu
-     ChartGridlineOptionsDialog [button] < AddChartElementMenu
-   ChartLegend [gallery] < AddChartElementMenu
-     ChartLegendOptionsDialogDialog [button] < AddChartElementMenu
-   ChartLines [gallery] < AddChartElementMenu
-   ChartTrendline [gallery] < AddChartElementMenu
-     ChartTrendlineOptionsDialog [button] < AddChartElementMenu
-   ChartUpDownBars [gallery] < AddChartElementMenu
-     ChartUpDownBarsOptionsDialog [button] < AddChartElementMenu
- ChartLayoutGallery [gallery]
### GroupChartStyles
- ChartColorsGallery [gallery]
- ChartStylesGallery [gallery]
- ChartDesigner [button]
### GroupChartData
- ChartSwitchRowColumn [button]
- ChartEditDataSource [button]
### GroupChartType
- ChartChangeType [button]
### GroupChartLocation
- ChartPlacement [button]

## TabChartToolsFormatPivotChart  (TabSetPivotChartTools)
### 
- GroupChartCurrentSelection [group]
- GroupShapesChart [group]
- GroupShapeStyles [group]
- GroupWordArtStyles [group]
- GroupArrange [group]
- GroupSize [group]
### GroupChartCurrentSelection
- ChartElementSelector [comboBox]
- ChartFormatSelection [button]
- ChartResetToMatchStyle [button]
### GroupShapesChart
- ShapesInsertGallery [gallery]
- ShapeChangeShapeGallery [gallery]
### GroupShapeStyles
- ShapeStylesGallery [gallery]
-   ShapeStylesOtherThemeFillsGallery [gallery] < ShapeStylesGallery
- ShapeFillColorPicker [gallery]
-   ObjectFillMoreColorsDialog [button] < ShapeFillColorPicker
-   ObjectPictureFill [button] < ShapeFillColorPicker
-   GradientGallery [gallery] < ShapeFillColorPicker
-     ShapeFillMoreGradientsDialog [button] < ShapeFillColorPicker
-   ShapeFillTextureGallery [gallery] < ShapeFillColorPicker
-     MoreTextureOptions [button] < ShapeFillColorPicker
- ShapeOutlineColorPicker [gallery]
-   ObjectBorderOutlineColorMoreColorsDialog [button] < ShapeOutlineColorPicker
-   OutlineWeightGallery [gallery] < ShapeOutlineColorPicker
-     LineStylesDialog [button] < ShapeOutlineColorPicker
-   GallerySketchyShapes [gallery] < ShapeOutlineColorPicker
-     MoreLines2 [button] < ShapeOutlineColorPicker
-   OutlineDashesGallery [gallery] < ShapeOutlineColorPicker
-     LineStylesDialog [button] < ShapeOutlineColorPicker
-   ArrowStyleGallery [gallery] < ShapeOutlineColorPicker
-     ArrowsMore [button] < ShapeOutlineColorPicker
- ShapeEffectsMenu [menu]
-   ObjectEffectPresetGallery [gallery] < ShapeEffectsMenu
-     _3DBevelOptionsDialog [button] < ShapeEffectsMenu
-   ObjectEffectShadowGallery [gallery] < ShapeEffectsMenu
-     ShadowOptionsDialog [button] < ShapeEffectsMenu
-   ReflectionGallery [gallery] < ShapeEffectsMenu
-     ReflectionsMoreOptions [button] < ShapeEffectsMenu
-   ObjectEffectGlowGallery [gallery] < ShapeEffectsMenu
-     GlowColorPicker [gallery] < ShapeEffectsMenu
-     GlowColorMoreColorsDialog [button] < ShapeEffectsMenu
-     GlowsMoreOptions [button] < ShapeEffectsMenu
-   ObjectEffectSoftEdgesGallery [gallery] < ShapeEffectsMenu
-     SoftEdgesMoreOptions [button] < ShapeEffectsMenu
-   BevelShapeGallery [gallery] < ShapeEffectsMenu
-     _3DBevelOptionsDialog [button] < ShapeEffectsMenu
-   _3DRotationGallery [gallery] < ShapeEffectsMenu
-     _3DRotationOptionsDialog [button] < ShapeEffectsMenu
- ObjectFormatDialog [button (dialogBoxLauncher)]
### GroupWordArtStyles
- TextStylesGallery [gallery]
-   WordArtClear [button] < TextStylesGallery
- TextFillColorPicker [gallery]
-   TextFillColorMoreColorsDialog [button] < TextFillColorPicker
-   TextPictureFill [button] < TextFillColorPicker
-   TextFillGradientGallery [gallery] < TextFillColorPicker
-     TextFillMoreGradientsDialog [button] < TextFillColorPicker
-   TextFillTextureGallery [gallery] < TextFillColorPicker
-     TextFillMoreTextures [button] < TextFillColorPicker
- TextOutlineColorPicker [gallery]
-   TextOutlineColorMoreColorsDialog [button] < TextOutlineColorPicker
-   TextOutlineWeightGallery [gallery] < TextOutlineColorPicker
-     TextOutlineMoreLinesDialog [button] < TextOutlineColorPicker
-   TextOutlineDashesGallery [gallery] < TextOutlineColorPicker
-     TextOutlineMoreLinesDialog [button] < TextOutlineColorPicker
- TextEffectsMenu [menu]
-   TextEffectShadowGallery [gallery] < TextEffectsMenu
-     TextEffectsMoreShadowsDialog [button] < TextEffectsMenu
-   TextReflectionGallery [gallery] < TextEffectsMenu
-     TextEffectsReflectionsMoreOptions [button] < TextEffectsMenu
-   TextEffectGlowGallery [gallery] < TextEffectsMenu
-     TextGlowColorPicker [gallery] < TextEffectsMenu
-     TextGlowColorMoreColorsDialog [button] < TextEffectsMenu
-     TextEffectsGlowsMoreOptions [button] < TextEffectsMenu
-   BevelTextGallery [gallery] < TextEffectsMenu
-     TextEffectsBevelMore3DOptionsDialog [button] < TextEffectsMenu
-   TextEffects3DRotationGallery [gallery] < TextEffectsMenu
-     TextEffects3DRotationOptionsDialog [button] < TextEffectsMenu
-   TextEffectTransformGallery [gallery] < TextEffectsMenu
- WordArtFormatDialog [button (dialogBoxLauncher)]
### GroupArrange
- ObjectBringForwardMenu [splitButton]
-   ObjectBringForward [button] < ObjectBringForwardMenu
-   ObjectBringToFront [button] < ObjectBringForwardMenu
- ObjectSendBackwardMenu [splitButton]
-   ObjectSendBackward [button] < ObjectSendBackwardMenu
-   ObjectSendToBack [button] < ObjectSendBackwardMenu
- SelectionPane [toggleButton]
- ObjectAlignMenu [menu]
-   ObjectsAlignLeft [button] < ObjectAlignMenu
-   ObjectsAlignCenterHorizontal [button] < ObjectAlignMenu
-   ObjectsAlignRight [button] < ObjectAlignMenu
-   ObjectsAlignTop [button] < ObjectAlignMenu
-   ObjectsAlignMiddleVertical [button] < ObjectAlignMenu
-   ObjectsAlignBottom [button] < ObjectAlignMenu
-   AlignDistributeHorizontallyClassic [button] < ObjectAlignMenu
-   AlignDistributeVerticallyClassic [button] < ObjectAlignMenu
-   SnapToGrid [toggleButton] < ObjectAlignMenu
-   SnapToShapes [toggleButton] < ObjectAlignMenu
-   ViewGridlinesToggleExcel [toggleButton] < ObjectAlignMenu
- ObjectsGroupMenu [menu]
-   ObjectsGroup [button] < ObjectsGroupMenu
-   ObjectsRegroup [button] < ObjectsGroupMenu
-   ObjectsUngroup [button] < ObjectsGroupMenu
- ObjectRotateGallery [gallery]
-   ObjectRotationOptionsDialog [button] < ObjectRotateGallery
- ObjectsArrangeMenu [menu]
-   ObjectBringToFront [button] < ObjectsArrangeMenu
-   ObjectBringForward [button] < ObjectsArrangeMenu
-   ObjectSendToBack [button] < ObjectsArrangeMenu
-   ObjectSendBackward [button] < ObjectsArrangeMenu
-   ObjectsGroup [button] < ObjectsArrangeMenu
-   ObjectsUngroup [button] < ObjectsArrangeMenu
-   ObjectsRegroup [button] < ObjectsArrangeMenu
-   ObjectAlignMenu [menu] < ObjectsArrangeMenu
-     ObjectsAlignLeft [button] < ObjectsArrangeMenu
-     ObjectsAlignCenterHorizontal [button] < ObjectsArrangeMenu
-     ObjectsAlignRight [button] < ObjectsArrangeMenu
-     ObjectsAlignTop [button] < ObjectsArrangeMenu
-     ObjectsAlignMiddleVertical [button] < ObjectsArrangeMenu
-     ObjectsAlignBottom [button] < ObjectsArrangeMenu
-     AlignDistributeHorizontallyClassic [button] < ObjectsArrangeMenu
-     AlignDistributeVerticallyClassic [button] < ObjectsArrangeMenu
-     SnapToGrid [toggleButton] < ObjectsArrangeMenu
-     SnapToShapes [toggleButton] < ObjectsArrangeMenu
-     ViewGridlinesToggleExcel [toggleButton] < ObjectsArrangeMenu
-   ObjectRotateGallery [gallery] < ObjectsArrangeMenu
-     ObjectRotationOptionsDialog [button] < ObjectsArrangeMenu
### GroupSize
- ShapeHeight [control]
- ShapeWidth [control]
- ObjectSizeAndPropertiesDialog [button (dialogBoxLauncher)]

## TabPivotChartToolsDesign  (TabSetPivotChartTools)
### 
- GroupChartData [group]
- GroupChartStyles [group]
- GroupChartLocation [group]
### GroupChartData
- ChartSwitchRowColumn [button]
- ChartEditDataSource [button]
### GroupChartStyles
- ChartColorsGallery [gallery]
- ChartStylesGallery [gallery]
- ChartDesigner [button]
### GroupChartLocation
- ChartPlacement [button]

## TabPivotChartToolsLayout  (TabSetPivotChartTools)
### 
- GroupChartCurrentSelection [group]
- GroupChartShapes [group]
- GroupChartLabels [group]
- GroupChartAxes [group]
- GroupChartBackground [group]
- GroupChartAnalysis [group]
- GroupChartProperties [group]
### GroupChartCurrentSelection
- ChartElementSelector [comboBox]
- ChartFormatSelection [button]
- ChartResetToMatchStyle [button]
### GroupChartShapes
- PictureInsertFromFile [button]
- ShapesInsertGallery [gallery]
- TextBoxInsertExcel [toggleButton]
- TextBoxInsertMenu [splitButton]
-   TextBoxInsertHorizontal [toggleButton] < TextBoxInsertMenu
-   TextBoxInsertVertical [toggleButton] < TextBoxInsertMenu
### GroupChartLabels
- ChartTitle [gallery]
-   ChartTitleOptionsDialog [button] < ChartTitle
- ChartAxisTitles [menu]
-   ChartPrimaryHorizontalAxisTitle [gallery] < ChartAxisTitles
-     ChartPrimaryHorizontalAxisTitleOptionsDialog [button] < ChartAxisTitles
-   ChartPrimaryVerticalAxisTitle [gallery] < ChartAxisTitles
-     ChartPrimaryVerticalAxisTitleOptionsDialog [button] < ChartAxisTitles
-   ChartSecondaryHorizontalAxisTitle [gallery] < ChartAxisTitles
-     ChartSecondaryHorizontalAxisTitleOptionsDialog [button] < ChartAxisTitles
-   ChartSecondaryVerticalAxisTitle [gallery] < ChartAxisTitles
-     ChartSecondaryVerticalAxisTitleOptionsDialog [button] < ChartAxisTitles
-   ChartDepthAxisTitle [gallery] < ChartAxisTitles
-     ChartDepthAxisTitleOptionsDialog [button] < ChartAxisTitles
- ChartLegend [gallery]
-   ChartLegendOptionsDialogDialog [button] < ChartLegend
- ChartDataLabel [gallery]
-   ChartDataLabelDialog [button] < ChartDataLabel
- ChartDataTable [gallery]
-   ChartDataTableOptionsDialog [button] < ChartDataTable
### GroupChartAxes
- ChartAxes [menu]
-   ChartPrimaryHorizontalAxis [gallery] < ChartAxes
-     ChartPrimaryHorizontalAxisOptionsDialog [button] < ChartAxes
-   ChartPrimaryVerticalAxis [gallery] < ChartAxes
-     ChartPrimaryVerticalAxisOptionsDialog [button] < ChartAxes
-   ChartSecondaryHorizontalAxis [gallery] < ChartAxes
-     ChartSecondaryHorizontalAxisOption [button] < ChartAxes
-   ChartSecondaryVerticalAxis [gallery] < ChartAxes
-     ChartSecondaryVerticalAxisOptionsDialog [button] < ChartAxes
-   ChartDepthAxis [gallery] < ChartAxes
-     ChartDepthAxisOptionsDialog [button] < ChartAxes
- ChartGridlines [menu]
-   ChartPrimaryHorizontalGridlines [gallery] < ChartGridlines
-     ChartPrimaryHorizontalGridlinesOptionsDialog [button] < ChartGridlines
-   ChartPrimaryVerticalGridlines [gallery] < ChartGridlines
-     ChartPrimaryVerticalGridlinesOptionsDialog [button] < ChartGridlines
-   ChartSecondaryHorizontalGridlines [gallery] < ChartGridlines
-     ChartSecondaryHorizontalGridlinesOptionsDialog [button] < ChartGridlines
-   ChartSecondaryVerticalGridlines [gallery] < ChartGridlines
-     ChartSecondaryVerticalGridlinesOptionsDialog [button] < ChartGridlines
-   ChartDepthGridlines [gallery] < ChartGridlines
-     ChartDepthGridlinesOptionsDialog [button] < ChartGridlines
### GroupChartBackground
- ChartPlotArea [gallery]
-   ChartPlotAreaOptionsDialog [button] < ChartPlotArea
- ChartWall [gallery]
-   ChartWallOptionsDialog [button] < ChartWall
- ChartFloor [gallery]
-   ChartFloorOptionsDialog [button] < ChartFloor
- Chart3DView [button]
### GroupChartAnalysis
- ChartTrendline [gallery]
-   ChartTrendlineOptionsDialog [button] < ChartTrendline
- ChartLines [gallery]
- ChartUpDownBars [gallery]
-   ChartUpDownBarsOptionsDialog [button] < ChartUpDownBars
- ChartErrorBars [gallery]
-   ChartErrorBarsOptionsDialog [button] < ChartErrorBars
### GroupChartProperties
- ChartNameLabel [labelControl]
- ChartName [control]

## TabPivotChartToolsFormat  (TabSetPivotChartTools)
### 
- GroupChartCurrentSelection [group]
- GroupShapeStyles [group]
- GroupWordArtStyles [group]
- GroupArrange [group]
- GroupSize [group]
### GroupChartCurrentSelection
- ChartElementSelector [comboBox]
- ChartFormatSelection [button]
- ChartResetToMatchStyle [button]
### GroupShapeStyles
- ShapeStylesGallery [gallery]
-   ShapeStylesOtherThemeFillsGallery [gallery] < ShapeStylesGallery
- ShapeFillColorPicker [gallery]
-   ObjectFillMoreColorsDialog [button] < ShapeFillColorPicker
-   ObjectPictureFill [button] < ShapeFillColorPicker
-   GradientGallery [gallery] < ShapeFillColorPicker
-     ShapeFillMoreGradientsDialog [button] < ShapeFillColorPicker
-   ShapeFillTextureGallery [gallery] < ShapeFillColorPicker
-     MoreTextureOptions [button] < ShapeFillColorPicker
- ShapeOutlineColorPicker [gallery]
-   ObjectBorderOutlineColorMoreColorsDialog [button] < ShapeOutlineColorPicker
-   OutlineWeightGallery [gallery] < ShapeOutlineColorPicker
-     LineStylesDialog [button] < ShapeOutlineColorPicker
-   GallerySketchyShapes [gallery] < ShapeOutlineColorPicker
-     MoreLines2 [button] < ShapeOutlineColorPicker
-   OutlineDashesGallery [gallery] < ShapeOutlineColorPicker
-     LineStylesDialog [button] < ShapeOutlineColorPicker
-   ArrowStyleGallery [gallery] < ShapeOutlineColorPicker
-     ArrowsMore [button] < ShapeOutlineColorPicker
- ShapeEffectsMenu [menu]
-   ObjectEffectPresetGallery [gallery] < ShapeEffectsMenu
-     _3DBevelOptionsDialog [button] < ShapeEffectsMenu
-   ObjectEffectShadowGallery [gallery] < ShapeEffectsMenu
-     ShadowOptionsDialog [button] < ShapeEffectsMenu
-   ReflectionGallery [gallery] < ShapeEffectsMenu
-     ReflectionsMoreOptions [button] < ShapeEffectsMenu
-   ObjectEffectGlowGallery [gallery] < ShapeEffectsMenu
-     GlowColorPicker [gallery] < ShapeEffectsMenu
-     GlowColorMoreColorsDialog [button] < ShapeEffectsMenu
-     GlowsMoreOptions [button] < ShapeEffectsMenu
-   ObjectEffectSoftEdgesGallery [gallery] < ShapeEffectsMenu
-     SoftEdgesMoreOptions [button] < ShapeEffectsMenu
-   BevelShapeGallery [gallery] < ShapeEffectsMenu
-     _3DBevelOptionsDialog [button] < ShapeEffectsMenu
-   _3DRotationGallery [gallery] < ShapeEffectsMenu
-     _3DRotationOptionsDialog [button] < ShapeEffectsMenu
- ObjectFormatDialog [button (dialogBoxLauncher)]
### GroupWordArtStyles
- TextStylesGallery [gallery]
-   WordArtClear [button] < TextStylesGallery
- TextFillColorPicker [gallery]
-   TextFillColorMoreColorsDialog [button] < TextFillColorPicker
-   TextPictureFill [button] < TextFillColorPicker
-   TextFillGradientGallery [gallery] < TextFillColorPicker
-     TextFillMoreGradientsDialog [button] < TextFillColorPicker
-   TextFillTextureGallery [gallery] < TextFillColorPicker
-     TextFillMoreTextures [button] < TextFillColorPicker
- TextOutlineColorPicker [gallery]
-   TextOutlineColorMoreColorsDialog [button] < TextOutlineColorPicker
-   TextOutlineWeightGallery [gallery] < TextOutlineColorPicker
-     TextOutlineMoreLinesDialog [button] < TextOutlineColorPicker
-   TextOutlineDashesGallery [gallery] < TextOutlineColorPicker
-     TextOutlineMoreLinesDialog [button] < TextOutlineColorPicker
- TextEffectsMenu [menu]
-   TextEffectShadowGallery [gallery] < TextEffectsMenu
-     TextEffectsMoreShadowsDialog [button] < TextEffectsMenu
-   TextReflectionGallery [gallery] < TextEffectsMenu
-     TextEffectsReflectionsMoreOptions [button] < TextEffectsMenu
-   TextEffectGlowGallery [gallery] < TextEffectsMenu
-     TextGlowColorPicker [gallery] < TextEffectsMenu
-     TextGlowColorMoreColorsDialog [button] < TextEffectsMenu
-     TextEffectsGlowsMoreOptions [button] < TextEffectsMenu
-   BevelTextGallery [gallery] < TextEffectsMenu
-     TextEffectsBevelMore3DOptionsDialog [button] < TextEffectsMenu
-   TextEffects3DRotationGallery [gallery] < TextEffectsMenu
-     TextEffects3DRotationOptionsDialog [button] < TextEffectsMenu
-   TextEffectTransformGallery [gallery] < TextEffectsMenu
- WordArtFormatDialog [button (dialogBoxLauncher)]
### GroupArrange
- ObjectBringForwardMenu [splitButton]
-   ObjectBringForward [button] < ObjectBringForwardMenu
-   ObjectBringToFront [button] < ObjectBringForwardMenu
- ObjectSendBackwardMenu [splitButton]
-   ObjectSendBackward [button] < ObjectSendBackwardMenu
-   ObjectSendToBack [button] < ObjectSendBackwardMenu
- SelectionPane [toggleButton]
- ObjectAlignMenu [menu]
-   ObjectsAlignLeft [button] < ObjectAlignMenu
-   ObjectsAlignCenterHorizontal [button] < ObjectAlignMenu
-   ObjectsAlignRight [button] < ObjectAlignMenu
-   ObjectsAlignTop [button] < ObjectAlignMenu
-   ObjectsAlignMiddleVertical [button] < ObjectAlignMenu
-   ObjectsAlignBottom [button] < ObjectAlignMenu
-   AlignDistributeHorizontallyClassic [button] < ObjectAlignMenu
-   AlignDistributeVerticallyClassic [button] < ObjectAlignMenu
-   SnapToGrid [toggleButton] < ObjectAlignMenu
-   SnapToShapes [toggleButton] < ObjectAlignMenu
-   ViewGridlinesToggleExcel [toggleButton] < ObjectAlignMenu
- ObjectsGroupMenu [menu]
-   ObjectsGroup [button] < ObjectsGroupMenu
-   ObjectsRegroup [button] < ObjectsGroupMenu
-   ObjectsUngroup [button] < ObjectsGroupMenu
- ObjectRotateGallery [gallery]
-   ObjectRotationOptionsDialog [button] < ObjectRotateGallery
- ObjectsArrangeMenu [menu]
-   ObjectBringToFront [button] < ObjectsArrangeMenu
-   ObjectBringForward [button] < ObjectsArrangeMenu
-   ObjectSendToBack [button] < ObjectsArrangeMenu
-   ObjectSendBackward [button] < ObjectsArrangeMenu
-   ObjectsGroup [button] < ObjectsArrangeMenu
-   ObjectsUngroup [button] < ObjectsArrangeMenu
-   ObjectsRegroup [button] < ObjectsArrangeMenu
-   ObjectAlignMenu [menu] < ObjectsArrangeMenu
-     ObjectsAlignLeft [button] < ObjectsArrangeMenu
-     ObjectsAlignCenterHorizontal [button] < ObjectsArrangeMenu
-     ObjectsAlignRight [button] < ObjectsArrangeMenu
-     ObjectsAlignTop [button] < ObjectsArrangeMenu
-     ObjectsAlignMiddleVertical [button] < ObjectsArrangeMenu
-     ObjectsAlignBottom [button] < ObjectsArrangeMenu
-     AlignDistributeHorizontallyClassic [button] < ObjectsArrangeMenu
-     AlignDistributeVerticallyClassic [button] < ObjectsArrangeMenu
-     SnapToGrid [toggleButton] < ObjectsArrangeMenu
-     SnapToShapes [toggleButton] < ObjectsArrangeMenu
-     ViewGridlinesToggleExcel [toggleButton] < ObjectsArrangeMenu
-   ObjectRotateGallery [gallery] < ObjectsArrangeMenu
-     ObjectRotationOptionsDialog [button] < ObjectsArrangeMenu
### GroupSize
- ShapeHeight [control]
- ShapeWidth [control]
- ObjectSizeAndPropertiesDialog [button (dialogBoxLauncher)]

##   (TabSetInkTools)
### 
- TabInkToolsPens [tab]

## TabInkToolsPens  (TabSetInkTools)
### 
- GroupPensWrite [group]
- GroupPens [group]
- GroupInkSelect [group]
- GroupInkFormat [group]
- GroupInkPens [group]
- GroupInkClose [group]
### GroupPensWrite
- InkBallpointPen [toggleButton]
- InkHighlighter [toggleButton]
- InkEraser [toggleButton]
- ObjectsSelect [toggleButton]
### GroupPens
- PensGallery [gallery]
-   PenAddFavorite [button] < PensGallery
- InkColorPicker [gallery]
-   InkColorMoreColorsDialog [button] < InkColorPicker
- LineThickness [gallery]
-   LineStylesDialog [button] < LineThickness
### GroupInkSelect
- ObjectsSelect [toggleButton]
### GroupInkFormat
- ObjectsSelect [toggleButton]
### GroupInkPens
- InkFeltTipPen [toggleButton]
### GroupInkClose
- InkStopInkingReadingView [button]

##   (TabSetSparkline)
### 
- TabSparklineDesign [tab]

## TabSparklineDesign  (TabSetSparkline)
### 
- GroupSparklineDataEdit [group]
- GroupSparklineType [group]
- GroupSparklineDetailsShow [group]
- GroupSparklineStyle [group]
- GroupSparklineGroup [group]
### GroupSparklineDataEdit
- SparklineEditLocationAndDataMenu [splitButton]
-   SparklineEditLocationAndData [button] < SparklineEditLocationAndDataMenu
-   SparklineEditSingleData [button] < SparklineEditLocationAndDataMenu
-   SparklineHiddenAndEmptyCells [button] < SparklineEditLocationAndDataMenu
-   SwitchSparklineRowColumn [button] < SparklineEditLocationAndDataMenu
### GroupSparklineType
- SparklineConvertToLine [toggleButton]
- SparklineConvertToColumn [toggleButton]
- SparklineConvertToWinLoss [toggleButton]
### GroupSparklineDetailsShow
- SparklineToggleHighPoint [checkBox]
- SparklineToggleLowPoint [checkBox]
- SparklineToggleNegativePoints [checkBox]
- SparklineToggleFirstPoint [checkBox]
- SparklineToggleLastPoint [checkBox]
- SparklineToggleMarkers [checkBox]
### GroupSparklineStyle
- SparklineColorStyleGallery [gallery]
- SparklineColorPicker [gallery]
-   SparklineMoreColors [button] < SparklineColorPicker
-   SparklineLineWeightGallery [gallery] < SparklineColorPicker
-     SparklineCustomWeight [toggleButton] < SparklineColorPicker
- SparklineAccentColorMenu [menu]
-   SparkLineNegativePointColorPicker [gallery] < SparklineAccentColorMenu
-     SparklineNegativePointMoreColors [button] < SparklineAccentColorMenu
-   SparklineMarkerColorPicker [gallery] < SparklineAccentColorMenu
-     SparklineMarkerMoreColors [button] < SparklineAccentColorMenu
-   SparklineHighPointColorPicker [gallery] < SparklineAccentColorMenu
-     SparklineHighPointMoreColors [button] < SparklineAccentColorMenu
-   SparklineLowPointColorPicker [gallery] < SparklineAccentColorMenu
-     SparklineLowPointMoreColors [button] < SparklineAccentColorMenu
-   SparklineFirstPointColorPicker [gallery] < SparklineAccentColorMenu
-     SparklineFirstPointMoreColors [button] < SparklineAccentColorMenu
-   SparklineLastPointColorPicker [gallery] < SparklineAccentColorMenu
-     SparklineLastPointMoreColors [button] < SparklineAccentColorMenu
### GroupSparklineGroup
- SparklineAxisMenu [menu]
-   SparklineAxisGeneralType [checkBox] < SparklineAxisMenu
-   SparklineAxisDateType [checkBox] < SparklineAxisMenu
-   SparklineHorizontalAxisToggle [checkBox] < SparklineAxisMenu
-   SparklinePlotDataRightToLeft [checkBox] < SparklineAxisMenu
-   SparklineVerticalAxisAutomaticMinimumValue [checkBox] < SparklineAxisMenu
-   SparklineVerticalAxisSameMinimumValue [checkBox] < SparklineAxisMenu
-   SparklineVerticalAxisCustomMinimumValue [checkBox] < SparklineAxisMenu
-   SparklineVerticalAxisAutomaticMaximumValue [checkBox] < SparklineAxisMenu
-   SparklineVerticalAxisSameMaximumValue [checkBox] < SparklineAxisMenu
-   SparklineVerticalAxisCustomMaximumValue [checkBox] < SparklineAxisMenu
- SparklinesGroupSelected [button]
- SparklinesUngroupSelected [button]
- SparklineClearMenu [splitButton]
-   SparklineClearSelected [button] < SparklineClearMenu
-   SparklineClearSelectedGroups [button] < SparklineClearMenu

##   (TabSetTimeSlicerTools)
### 
- TabTimeSlicerDesign [tab]

## TabTimeSlicerDesign  (TabSetTimeSlicerTools)
### 
- GroupTimeSlicerOptions [group]
- GroupTimelineStyles [group]
- GroupArrange [group]
- GroupSize [group]
- GroupTimeSlicerShowOptions [group]
### GroupTimeSlicerOptions
- TimeSlicerNameLabel [labelControl]
- TimeSlicerName [control]
- SlicerConnectionsMenu [button]
### GroupTimelineStyles
- TimelineStyles [gallery]
-   TimelineCreateNewStyle [button] < TimelineStyles
### GroupArrange
- ObjectBringForwardMenu [splitButton]
-   ObjectBringForward [button] < ObjectBringForwardMenu
-   ObjectBringToFront [button] < ObjectBringForwardMenu
- ObjectSendBackwardMenu [splitButton]
-   ObjectSendBackward [button] < ObjectSendBackwardMenu
-   ObjectSendToBack [button] < ObjectSendBackwardMenu
- SelectionPane [toggleButton]
- ObjectAlignMenu [menu]
-   ObjectsAlignLeft [button] < ObjectAlignMenu
-   ObjectsAlignCenterHorizontal [button] < ObjectAlignMenu
-   ObjectsAlignRight [button] < ObjectAlignMenu
-   ObjectsAlignTop [button] < ObjectAlignMenu
-   ObjectsAlignMiddleVertical [button] < ObjectAlignMenu
-   ObjectsAlignBottom [button] < ObjectAlignMenu
-   AlignDistributeHorizontallyClassic [button] < ObjectAlignMenu
-   AlignDistributeVerticallyClassic [button] < ObjectAlignMenu
-   SnapToGrid [toggleButton] < ObjectAlignMenu
-   SnapToShapes [toggleButton] < ObjectAlignMenu
-   ViewGridlinesToggleExcel [toggleButton] < ObjectAlignMenu
- ObjectsGroupMenu [menu]
-   ObjectsGroup [button] < ObjectsGroupMenu
-   ObjectsRegroup [button] < ObjectsGroupMenu
-   ObjectsUngroup [button] < ObjectsGroupMenu
- ObjectRotateGallery [gallery]
-   ObjectRotationOptionsDialog [button] < ObjectRotateGallery
- ObjectsArrangeMenu [menu]
-   ObjectBringToFront [button] < ObjectsArrangeMenu
-   ObjectBringForward [button] < ObjectsArrangeMenu
-   ObjectSendToBack [button] < ObjectsArrangeMenu
-   ObjectSendBackward [button] < ObjectsArrangeMenu
-   ObjectsGroup [button] < ObjectsArrangeMenu
-   ObjectsUngroup [button] < ObjectsArrangeMenu
-   ObjectsRegroup [button] < ObjectsArrangeMenu
-   ObjectAlignMenu [menu] < ObjectsArrangeMenu
-     ObjectsAlignLeft [button] < ObjectsArrangeMenu
-     ObjectsAlignCenterHorizontal [button] < ObjectsArrangeMenu
-     ObjectsAlignRight [button] < ObjectsArrangeMenu
-     ObjectsAlignTop [button] < ObjectsArrangeMenu
-     ObjectsAlignMiddleVertical [button] < ObjectsArrangeMenu
-     ObjectsAlignBottom [button] < ObjectsArrangeMenu
-     AlignDistributeHorizontallyClassic [button] < ObjectsArrangeMenu
-     AlignDistributeVerticallyClassic [button] < ObjectsArrangeMenu
-     SnapToGrid [toggleButton] < ObjectsArrangeMenu
-     SnapToShapes [toggleButton] < ObjectsArrangeMenu
-     ViewGridlinesToggleExcel [toggleButton] < ObjectsArrangeMenu
-   ObjectRotateGallery [gallery] < ObjectsArrangeMenu
-     ObjectRotationOptionsDialog [button] < ObjectsArrangeMenu
### GroupSize
- ShapeHeight [control]
- ShapeWidth [control]
- ObjectSizeAndPropertiesDialog [button (dialogBoxLauncher)]
### GroupTimeSlicerShowOptions
- TimeSlicerShowHeader [checkBox]
- TimeSlicerShowSelectionLabel [checkBox]
- TimeSlicerShowScrollbar [checkBox]
- TimeSlicerShowTimeLevel [checkBox]

##   (TabSetSlicerTools)
### 
- TabSlicerDesign [tab]

## TabSlicerDesign  (TabSetSlicerTools)
### 
- GroupSlicerOptions [group]
- GroupSlicerStyles [group]
- GroupArrange [group]
- GroupSlicerRowColumnSize [group]
- GroupSize [group]
### GroupSlicerOptions
- SlicerNameLabel [labelControl]
- SlicerName [control]
- SlicerSettings [button]
- SlicerConnectionsMenu [button]
### GroupSlicerStyles
- SlicerStyles [gallery]
-   SlicerCreateNewStyle [button] < SlicerStyles
### GroupArrange
- ObjectBringForwardMenu [splitButton]
-   ObjectBringForward [button] < ObjectBringForwardMenu
-   ObjectBringToFront [button] < ObjectBringForwardMenu
- ObjectSendBackwardMenu [splitButton]
-   ObjectSendBackward [button] < ObjectSendBackwardMenu
-   ObjectSendToBack [button] < ObjectSendBackwardMenu
- SelectionPane [toggleButton]
- ObjectAlignMenu [menu]
-   ObjectsAlignLeft [button] < ObjectAlignMenu
-   ObjectsAlignCenterHorizontal [button] < ObjectAlignMenu
-   ObjectsAlignRight [button] < ObjectAlignMenu
-   ObjectsAlignTop [button] < ObjectAlignMenu
-   ObjectsAlignMiddleVertical [button] < ObjectAlignMenu
-   ObjectsAlignBottom [button] < ObjectAlignMenu
-   AlignDistributeHorizontallyClassic [button] < ObjectAlignMenu
-   AlignDistributeVerticallyClassic [button] < ObjectAlignMenu
-   SnapToGrid [toggleButton] < ObjectAlignMenu
-   SnapToShapes [toggleButton] < ObjectAlignMenu
-   ViewGridlinesToggleExcel [toggleButton] < ObjectAlignMenu
- ObjectsGroupMenu [menu]
-   ObjectsGroup [button] < ObjectsGroupMenu
-   ObjectsRegroup [button] < ObjectsGroupMenu
-   ObjectsUngroup [button] < ObjectsGroupMenu
- ObjectRotateGallery [gallery]
-   ObjectRotationOptionsDialog [button] < ObjectRotateGallery
- ObjectsArrangeMenu [menu]
-   ObjectBringToFront [button] < ObjectsArrangeMenu
-   ObjectBringForward [button] < ObjectsArrangeMenu
-   ObjectSendToBack [button] < ObjectsArrangeMenu
-   ObjectSendBackward [button] < ObjectsArrangeMenu
-   ObjectsGroup [button] < ObjectsArrangeMenu
-   ObjectsUngroup [button] < ObjectsArrangeMenu
-   ObjectsRegroup [button] < ObjectsArrangeMenu
-   ObjectAlignMenu [menu] < ObjectsArrangeMenu
-     ObjectsAlignLeft [button] < ObjectsArrangeMenu
-     ObjectsAlignCenterHorizontal [button] < ObjectsArrangeMenu
-     ObjectsAlignRight [button] < ObjectsArrangeMenu
-     ObjectsAlignTop [button] < ObjectsArrangeMenu
-     ObjectsAlignMiddleVertical [button] < ObjectsArrangeMenu
-     ObjectsAlignBottom [button] < ObjectsArrangeMenu
-     AlignDistributeHorizontallyClassic [button] < ObjectsArrangeMenu
-     AlignDistributeVerticallyClassic [button] < ObjectsArrangeMenu
-     SnapToGrid [toggleButton] < ObjectsArrangeMenu
-     SnapToShapes [toggleButton] < ObjectsArrangeMenu
-     ViewGridlinesToggleExcel [toggleButton] < ObjectsArrangeMenu
-   ObjectRotateGallery [gallery] < ObjectsArrangeMenu
-     ObjectRotationOptionsDialog [button] < ObjectsArrangeMenu
### GroupSlicerRowColumnSize
- SlicerNumberOfColumns [control]
- SlicerRowHeight [control]
- SlicerColumnWidth [control]
### GroupSize
- ShapeHeight [control]
- ShapeWidth [control]
- ObjectSizeAndPropertiesDialog [button (dialogBoxLauncher)]

##   (TabSetPowerQueryEdit)
### 
- TabPowerQueryQueryEdit [tab]

## TabPowerQueryQueryEdit  (TabSetPowerQueryEdit)
### 
- GroupPowerQueryQueryData [group]
- GroupPowerQueryQueryLoad [group]
- GroupPowerQueryQueryManage [group]
- GroupPowerQueryQueryCombine [group]
- GroupPowerQueryQueryExport [group]
### GroupPowerQueryQueryData
- PowerQueryEditQuery [button]
- PowerQueryQueryProperties [button]
- PowerQueryDeleteQuery [button]
### GroupPowerQueryQueryLoad
- FlyoutAnchorPowerQueryRefresh [splitButton]
-   GetPowerQueryRefresh [button] < FlyoutAnchorPowerQueryRefresh
-   PowerQueryCancelRefresh [button] < FlyoutAnchorPowerQueryRefresh
- PowerQueryQueryLoadOptions [button]
### GroupPowerQueryQueryManage
- PowerQueryDuplicateQuery [button]
- PowerQueryReferenceQuery [button]
### GroupPowerQueryQueryCombine
- PowerQueryMergeQuery [button]
- PowerQueryAppendQuery [button]
### GroupPowerQueryQueryExport
- PowerQueryExportConnectionFile [button]

##   (TabSetEquationTools)
### 
- TabEquationToolsDesign [tab]

## TabEquationToolsDesign  (TabSetEquationTools)
### 
- GroupEquationTools [group]
- GroupEquationSymbols [group]
- GroupEquationStructures [group]
### GroupEquationTools
- EquationInsertGallery [gallery]
- InkEquation [button]
- EquationProfessional [button]
- EquationLinearFormat [button]
- EquationNormalText [toggleButton]
- EquationOptions [button (dialogBoxLauncher)]
### GroupEquationSymbols
- EquationSymbolsInsertGallery [gallery]
### GroupEquationStructures
- EquationFractionGallery [gallery]
- EquationScriptGallery [gallery]
- EquationRadicalGallery [gallery]
- EquationIntegralGallery [gallery]
- EquationLargeOperatorGallery [gallery]
- EquationDelimiterGallery [gallery]
- EquationFunctionGallery [gallery]
- EquationAccentGallery [gallery]
- EquationLimitGallery [gallery]
- EquationOperatorGallery [gallery]
- EquationMatrixGallery [gallery]

##   (TabSetSVGTools)
### 
- TabGraphicsToolsFormat [tab]

## TabGraphicsToolsFormat  (TabSetSVGTools)
### 
- GroupIcons [group]
- GroupGraphicsStyles [group]
- GroupAltText [group]
- GroupArrange [group]
- GroupGraphicsSize [group]
### GroupIcons
- GraphicChangeMenu [menu]
-   PictureChangeFromFile [button] < GraphicChangeMenu
-   PictureChangeFromMobile [button] < GraphicChangeMenu
-   PictureChangeFromM365 [button] < GraphicChangeMenu
-   PictureChangeFromOnline [button] < GraphicChangeMenu
-   ChangeIcon [button] < GraphicChangeMenu
-   PictureChangeFromClipboard [button] < GraphicChangeMenu
- SVGEdit [button]
### GroupGraphicsStyles
- GraphicsStylesGallery [gallery]
- GraphicsColorPicker [gallery]
-   ObjectFillMoreColorsDialog [button] < GraphicsColorPicker
- GraphicsOutlineColorPicker [gallery]
-   ObjectBorderOutlineColorMoreColorsDialog [button] < GraphicsOutlineColorPicker
-   OutlineWeightGallery [gallery] < GraphicsOutlineColorPicker
-     LineStylesDialog [button] < GraphicsOutlineColorPicker
-   OutlineDashesGallery [gallery] < GraphicsOutlineColorPicker
-     LineStylesDialog [button] < GraphicsOutlineColorPicker
- GraphicsEffectsMenu [menu]
-   ObjectEffectPresetGallery [gallery] < GraphicsEffectsMenu
-     _3DBevelOptionsDialog [button] < GraphicsEffectsMenu
-   ObjectEffectShadowGallery [gallery] < GraphicsEffectsMenu
-     ShadowOptionsDialog [button] < GraphicsEffectsMenu
-   ReflectionGallery [gallery] < GraphicsEffectsMenu
-     ReflectionsMoreOptions [button] < GraphicsEffectsMenu
-   ObjectEffectGlowGallery [gallery] < GraphicsEffectsMenu
-     GlowColorPicker [gallery] < GraphicsEffectsMenu
-     GlowColorMoreColorsDialog [button] < GraphicsEffectsMenu
-     GlowsMoreOptions [button] < GraphicsEffectsMenu
-   ObjectEffectSoftEdgesGallery [gallery] < GraphicsEffectsMenu
-     SoftEdgesMoreOptions [button] < GraphicsEffectsMenu
-   BevelShapeGallery [gallery] < GraphicsEffectsMenu
-     _3DBevelOptionsDialog [button] < GraphicsEffectsMenu
-   _3DRotationGallery [gallery] < GraphicsEffectsMenu
-     _3DRotationOptionsDialog [button] < GraphicsEffectsMenu
- GraphicsFormatDialog [button (dialogBoxLauncher)]
### GroupAltText
- AltTextPaneRibbon [toggleButton]
### GroupArrange
- ObjectBringForwardMenu [splitButton]
-   ObjectBringForward [button] < ObjectBringForwardMenu
-   ObjectBringToFront [button] < ObjectBringForwardMenu
- ObjectSendBackwardMenu [splitButton]
-   ObjectSendBackward [button] < ObjectSendBackwardMenu
-   ObjectSendToBack [button] < ObjectSendBackwardMenu
- SelectionPane [toggleButton]
- ObjectAlignMenu [menu]
-   ObjectsAlignLeft [button] < ObjectAlignMenu
-   ObjectsAlignCenterHorizontal [button] < ObjectAlignMenu
-   ObjectsAlignRight [button] < ObjectAlignMenu
-   ObjectsAlignTop [button] < ObjectAlignMenu
-   ObjectsAlignMiddleVertical [button] < ObjectAlignMenu
-   ObjectsAlignBottom [button] < ObjectAlignMenu
-   AlignDistributeHorizontallyClassic [button] < ObjectAlignMenu
-   AlignDistributeVerticallyClassic [button] < ObjectAlignMenu
-   SnapToGrid [toggleButton] < ObjectAlignMenu
-   SnapToShapes [toggleButton] < ObjectAlignMenu
-   ViewGridlinesToggleExcel [toggleButton] < ObjectAlignMenu
- ObjectsGroupMenu [menu]
-   ObjectsGroup [button] < ObjectsGroupMenu
-   ObjectsRegroup [button] < ObjectsGroupMenu
-   ObjectsUngroup [button] < ObjectsGroupMenu
- ObjectRotateGallery [gallery]
-   ObjectRotationOptionsDialog [button] < ObjectRotateGallery
- ObjectsArrangeMenu [menu]
-   ObjectBringToFront [button] < ObjectsArrangeMenu
-   ObjectBringForward [button] < ObjectsArrangeMenu
-   ObjectSendToBack [button] < ObjectsArrangeMenu
-   ObjectSendBackward [button] < ObjectsArrangeMenu
-   ObjectsGroup [button] < ObjectsArrangeMenu
-   ObjectsUngroup [button] < ObjectsArrangeMenu
-   ObjectsRegroup [button] < ObjectsArrangeMenu
-   ObjectAlignMenu [menu] < ObjectsArrangeMenu
-     ObjectsAlignLeft [button] < ObjectsArrangeMenu
-     ObjectsAlignCenterHorizontal [button] < ObjectsArrangeMenu
-     ObjectsAlignRight [button] < ObjectsArrangeMenu
-     ObjectsAlignTop [button] < ObjectsArrangeMenu
-     ObjectsAlignMiddleVertical [button] < ObjectsArrangeMenu
-     ObjectsAlignBottom [button] < ObjectsArrangeMenu
-     AlignDistributeHorizontallyClassic [button] < ObjectsArrangeMenu
-     AlignDistributeVerticallyClassic [button] < ObjectsArrangeMenu
-     SnapToGrid [toggleButton] < ObjectsArrangeMenu
-     SnapToShapes [toggleButton] < ObjectsArrangeMenu
-     ViewGridlinesToggleExcel [toggleButton] < ObjectsArrangeMenu
-   ObjectRotateGallery [gallery] < ObjectsArrangeMenu
-     ObjectRotationOptionsDialog [button] < ObjectsArrangeMenu
### GroupGraphicsSize
- PictureCropTools [splitButton]
-   PictureCrop [toggleButton] < PictureCropTools
-   PictureShapeGallery [gallery] < PictureCropTools
-   PictureCropAspectRatioMenu [menu] < PictureCropTools
-     PictureCropAspectRatio1To1 [button] < PictureCropTools
-     PictureCropAspectRatio2To3 [button] < PictureCropTools
-     PictureCropAspectRatio3To4 [button] < PictureCropTools
-     PictureCropAspectRatio3To5 [button] < PictureCropTools
-     PictureCropAspectRatio4To5 [button] < PictureCropTools
-     PictureCropAspectRatio3To2 [button] < PictureCropTools
-     PictureCropAspectRatio4To3 [button] < PictureCropTools
-     PictureCropAspectRatio5To3 [button] < PictureCropTools
-     PictureCropAspectRatio5To4 [button] < PictureCropTools
-     PictureCropAspectRatio16To9 [button] < PictureCropTools
-     PictureCropAspectRatio16To10 [button] < PictureCropTools
-   PictureFillCrop [button] < PictureCropTools
-   PictureFitCrop [button] < PictureCropTools
- ShapeHeight [control]
- ShapeWidth [control]
- ObjectSizeAndPropertiesDialog [button (dialogBoxLauncher)]

##   (TabSet3DModelTools)
### 
- Tab3DModelToolsFormat [tab]

## Tab3DModelToolsFormat  (TabSet3DModelTools)
### 
- Group3DModelPlay [group]
- GroupAdjust3DModel [group]
- Group3DModelStyles [group]
- GroupAltText [group]
- GroupArrange [group]
- Group3DModelSize [group]
### Group3DModelPlay
- _3DModelPlay [button]
- _3DModelSceneGallery [gallery]
### GroupAdjust3DModel
- Insert3DModelFallback [button]
- Insert3DModelDropdown [splitButton]
-   Insert3DModelDefault [button] < Insert3DModelDropdown
-   Insert3DModelFromFile [button] < Insert3DModelDropdown
-   Insert3DModelFromFileConsumer [button] < Insert3DModelDropdown
-   Insert3DModelFromOnline [button] < Insert3DModelDropdown
-   Insert3DModelFromOnlineConsumer [button] < Insert3DModelDropdown
- Reset3DModelSplitButton [splitButton]
-   Reset3DModel [button] < Reset3DModelSplitButton
-   Reset3DModelAndSize [button] < Reset3DModelSplitButton
### Group3DModelStyles
- Gallery3DModelViewPresets [gallery]
- FormatDialog3DModel [button (dialogBoxLauncher)]
### GroupAltText
- AltTextPaneRibbon [toggleButton]
### GroupArrange
- ObjectBringForwardMenu [splitButton]
-   ObjectBringForward [button] < ObjectBringForwardMenu
-   ObjectBringToFront [button] < ObjectBringForwardMenu
- ObjectSendBackwardMenu [splitButton]
-   ObjectSendBackward [button] < ObjectSendBackwardMenu
-   ObjectSendToBack [button] < ObjectSendBackwardMenu
- SelectionPane [toggleButton]
- ObjectAlignMenu [menu]
-   ObjectsAlignLeftSmart [button] < ObjectAlignMenu
-   ObjectsAlignCenterHorizontalSmart [button] < ObjectAlignMenu
-   ObjectsAlignRightSmart [button] < ObjectAlignMenu
-   ObjectsAlignTopSmart [button] < ObjectAlignMenu
-   ObjectsAlignMiddleVerticalSmart [button] < ObjectAlignMenu
-   ObjectsAlignBottomSmart [button] < ObjectAlignMenu
-   AlignDistributeHorizontally [button] < ObjectAlignMenu
-   AlignDistributeVertically [button] < ObjectAlignMenu
- ObjectsGroupMenu [menu]
-   ObjectsGroup [button] < ObjectsGroupMenu
-   ObjectsRegroup [button] < ObjectsGroupMenu
-   ObjectsUngroup [button] < ObjectsGroupMenu
### Group3DModelSize
- AdvancedEdit3DModel [toggleButton]
- ShapeHeight [control]
- ShapeWidth [control]
- ObjectSizeAndPositionDialog [button (dialogBoxLauncher)]

##   (TabSetBIVisualTools)
### 
- TabBIVisualTools [tab]

## TabBIVisualTools  (TabSetBIVisualTools)
### 
- GroupBIVisualTools [group]
- GroupAltText [group]
### GroupBIVisualTools
- BIVisualChangeType [button]
- BIVisualSourceDataGrid [button]
- BIVisualFormat [button]
### GroupAltText
- AltTextPaneRibbon [toggleButton]

##   (TabSetAccessibleAuthoring)
### 
- TabAccessibleAuthoring [tab]

## TabAccessibleAuthoring  (TabSetAccessibleAuthoring)
### 
- GroupAccessibleAuthoringReview [group]
- GroupAccessibleAuthoringColor [group]
- GroupAccessibleAuthoringStyle [group]
- GroupAccessibleAuthoringFormat [group]
- GroupAccessibleAuthoringNames [group]
- GroupAccessibleAuthoringResources [group]
### GroupAccessibleAuthoringReview
- AccessibilityChecker [button]
- Spelling [button]
- UpgradeWorkbook [button]
### GroupAccessibleAuthoringColor
- CellFillColorPicker [gallery]
-   FontShadingColorMoreColorsDialog [button] < CellFillColorPicker
- FontColorPicker [gallery]
-   FontColorMoreColorsDialogExcel [button] < FontColorPicker
### GroupAccessibleAuthoringStyle
- GalleryAccessibleCellStyles [gallery]
- GalleryAccessibleTableStyles [gallery]
- GalleryAccessiblePivotTableStyles [gallery]
- GalleryAccessibleThemeStyles [gallery]
-   ShapeStylesOtherThemeFillsGallery [gallery] < GalleryAccessibleThemeStyles
- AccessibleChartColorsGallery [gallery]
- NumberFormatGallery [gallery]
-   NumberFormatsDialog [button] < NumberFormatGallery
### GroupAccessibleAuthoringFormat
- FormatAsTableAccessibleGallery [gallery]
- UnmergeCells [button]
- WrapText [toggleButton]
- HyperlinkEdit [button]
### GroupAccessibleAuthoringNames
- AltTextPaneRibbon [toggleButton]
- NameDefine [button]
- SheetRename [button]
- TableName [control]
- PivotTableName [control]
### GroupAccessibleAuthoringResources
- AccessibilityHelp [button]

##   (None (Backstage View))
### 
- NoTcid [control]
- None [button]
- PlaceTabHome [tab]
- TabInfo [tab]
- None [button]
- TabOfficeStart [tab]
- None [button]
- TabRecent [tab]
- OfficeExtensionsGetAddinsButtonOnBackstageBelowOpenDivider [button]
- FileSave [button]
- TabSave [tab]
- MoveDocument [button]
- HistoryTab [button]
- TabPrint [tab]
- ShareDocument [button]
- TabShare [tab]
- TabPublish [tab]
- Publish2Tab [tab]
- FileClose [button]
- TabHelp [tab]
- None [button]
- TabOfficeFeedback [tab]
- OfficeFeedback [button]
- OfficeFeedbackBackstage [button]
- ApplicationOptionsDialog [button]

## PlaceTabHome  (None (Backstage View))
### 
- HomePageSlabCreateNew [group]
- None [group]
- None [control]
- HomePlaceRecommendedDocuments [group]
- GroupRecentDocuments [group]
### GroupRecentDocuments
- None [button]

## TabInfo  (None (Backstage View))
### 
- GroupUploadErrors [group]
- None [group]
- GroupARMPreview [group]
- GroupWorkflows [group]
- GrupPolicyWarning [group]
- GroupSecurityInformation [group]
- GroupSecurityWarning [group]
- None [group]
- AppGuardTrustPromotion [button]
- GroupProtectedView [group]
- GroupReadOnly [group]
- GroupReadOnlyServer [group]
- GroupReadOnlyOffline [group]
- GroupReadOnlyCheckOutRequired [group]
- GroupReadOnlyCachedOffline [group]
- GroupReadOnlyCached [group]
- GroupCheckIn [group]
- GroupFileConvert [group]
- OfficePolicyTips [group]
- GroupSignatures [group]
- GroupPermissions [group]
- GroupPrepareForSharing [group]
- DocuemntVersionHistory [group]
- GroupVersions [group]
- GroupLineageClearHistory [group]
- GroupPublishToExcelServices [group]
- GroupThumbnail [group]
- GroupDocumentProperties [group]
### GroupReadOnlyCachedOffline
- ReadOnlyCachedOffline [button]
### GroupCheckIn
- FileCheckOutDiscard [button]
### GroupFileConvert
- UpgradeDocument [button]
### GroupPermissions
- FileMarkAsFinal [control]
- FileMarkAsReadOnlyRecommended [control]
- FileDocumentEncrypt [control]
- UsagePermissions [button]
- FileAddDigitalSignature [button]
### GroupPrepareForSharing
- AllowPersonallyIdentifiableInformation [control]

## TabOfficeStart  (None (Backstage View))
### 
- GroupOfficeStart [group]

## TabRecent  (None (Backstage View))
### 
- GroupOpenPlaces [taskFormGroup]
- GroupModernOpenPlaces [group]
### GroupOpenPlaces
- ButtonCategoryRecentDocuments [category]
-   ButtonTaskRecentDocuments [task] < ButtonCategoryRecentDocuments
-     GroupRecentDocuments [group] < ButtonCategoryRecentDocuments
-     None [control] < ButtonCategoryRecentDocuments
-     ButtonOpenBrowse [button] < ButtonCategoryRecentDocuments
-     GroupRecoverUnsavedFiles [group] < ButtonCategoryRecentDocuments
-   ButtonTaskFixItHub [task] < ButtonCategoryRecentDocuments
-     FixItHub [group] < ButtonCategoryRecentDocuments
-   ButtonTaskSharedWithMe [task] < ButtonCategoryRecentDocuments
-     GroupSharedWithMeDocuments [group] < ButtonCategoryRecentDocuments
-     None [button] < ButtonCategoryRecentDocuments
-   ButtonTaskDynamicServiceProvider [task] < ButtonCategoryRecentDocuments
-     GroupRecentPlaces [group] < ButtonCategoryRecentDocuments
-     None [control] < ButtonCategoryRecentDocuments
-     None [group] < ButtonCategoryRecentDocuments
-     GroupGettingStartedOpenMRUPlacesOrNavSlab [group] < ButtonCategoryRecentDocuments
-     None [button] < ButtonCategoryRecentDocuments
-     SkyDriveUpsellGroup [group] < ButtonCategoryRecentDocuments
-     LoadingServicesGroup [group] < ButtonCategoryRecentDocuments
-     GroupAddLocation [group] < ButtonCategoryRecentDocuments
### GroupModernOpenPlaces
- ButtonGroupRecentDocuments [control]
- ButtonGroupCloudServices [control]
- ButtonCategoryCloudServices [category]
-   ButtonTaskRecentDocuments [task] < ButtonCategoryCloudServices
-     GroupRecentDocumentsNew [group] < ButtonCategoryCloudServices
-     None [control] < ButtonCategoryCloudServices
-     ButtonOpenBrowse [button] < ButtonCategoryCloudServices
-   ButtonTaskFixItHub [task] < ButtonCategoryCloudServices
-     CloudFixItHub [group] < ButtonCategoryCloudServices
-   ButtonTaskSharedWithMe [task] < ButtonCategoryCloudServices
-     GroupSharedWithMeDocumentsNew [group] < ButtonCategoryCloudServices
-     None [button] < ButtonCategoryCloudServices
-   ButtonTaskDynamicServiceProvider [task] < ButtonCategoryCloudServices
-     GroupRecentCloudPlaces [group] < ButtonCategoryCloudServices
-     None [control] < ButtonCategoryCloudServices
-     None [group] < ButtonCategoryCloudServices
-     GroupCloudGettingStartedOpenMRUPlacesOrNavSlab [group] < ButtonCategoryCloudServices
-     None [button] < ButtonCategoryCloudServices
-   ButtonTaskDynamicServiceProvider [control] < ButtonCategoryCloudServices
-     CloudSkyDriveUpsellGroup [group] < ButtonCategoryCloudServices
-     LoadingCloudServicesGroup [group] < ButtonCategoryCloudServices
- ButtonTaskRecentDocuments [task]
-   GroupRecentDocumentsNew [group] < ButtonTaskRecentDocuments
-     None [control] < ButtonTaskRecentDocuments
-     ButtonOpenBrowse [button] < ButtonTaskRecentDocuments
- GroupRecentDocumentsNew [group]
-   None [control] < GroupRecentDocumentsNew
-   ButtonOpenBrowse [button] < GroupRecentDocumentsNew
- None [control]
- ButtonOpenBrowse [button]
- ButtonTaskFixItHub [task]
-   CloudFixItHub [group] < ButtonTaskFixItHub
- CloudFixItHub [group]
- ButtonTaskSharedWithMe [task]
-   GroupSharedWithMeDocumentsNew [group] < ButtonTaskSharedWithMe
-     None [button] < ButtonTaskSharedWithMe
- GroupSharedWithMeDocumentsNew [group]
-   None [button] < GroupSharedWithMeDocumentsNew
- None [button]
- ButtonTaskDynamicServiceProvider [task]
-   GroupRecentCloudPlaces [group] < ButtonTaskDynamicServiceProvider
-     None [control] < ButtonTaskDynamicServiceProvider
-   None [group] < ButtonTaskDynamicServiceProvider
-   GroupCloudGettingStartedOpenMRUPlacesOrNavSlab [group] < ButtonTaskDynamicServiceProvider
-     None [button] < ButtonTaskDynamicServiceProvider
- GroupRecentCloudPlaces [group]
-   None [control] < GroupRecentCloudPlaces
- None [group]
- GroupCloudGettingStartedOpenMRUPlacesOrNavSlab [group]
-   None [button] < GroupCloudGettingStartedOpenMRUPlacesOrNavSlab
- ButtonTaskDynamicServiceProvider [control]
- GroupRecentPinnableCloudPlaces [group]
-   None [control] < GroupRecentPinnableCloudPlaces
- None [group]
- PinnableGroupCloudGettingStartedOpenMRUPlacesOrNavSlab [group]
-   None [button] < PinnableGroupCloudGettingStartedOpenMRUPlacesOrNavSlab
-   CloudSkyDriveUpsellGroup [group] < ButtonTaskDynamicServiceProvider
-     None [button] < ButtonTaskDynamicServiceProvider
- CloudSkyDriveUpsellGroup [group]
-   None [button] < CloudSkyDriveUpsellGroup
-   LoadingCloudServicesGroup [group] < ButtonTaskDynamicServiceProvider
- LoadingCloudServicesGroup [group]
- ButtonGroupLocalServices [control]
- ButtonCategoryLocalServices [category]
-   ButtonTaskDynamicServiceProvider [task] < ButtonCategoryLocalServices
-     GroupRecentLocalPlaces [group] < ButtonCategoryLocalServices
-     None [control] < ButtonCategoryLocalServices
-     GroupLocalGettingStartedOpenMRUPlacesOrNavSlab [group] < ButtonCategoryLocalServices
-     None [button] < ButtonCategoryLocalServices
-     GroupAddLocationLocal [group] < ButtonCategoryLocalServices
-   GroupRecentLocalPlaces [group] < ButtonTaskDynamicServiceProvider
-     None [control] < ButtonTaskDynamicServiceProvider
-   GroupLocalGettingStartedOpenMRUPlacesOrNavSlab [group] < ButtonTaskDynamicServiceProvider
-     None [button] < ButtonTaskDynamicServiceProvider
- GroupRecentLocalPlaces [group]
-   None [control] < GroupRecentLocalPlaces
- GroupLocalGettingStartedOpenMRUPlacesOrNavSlab [group]
-   None [button] < GroupLocalGettingStartedOpenMRUPlacesOrNavSlab
-   GroupAddLocationLocal [group] < ButtonTaskDynamicServiceProvider
- GroupAddLocationLocal [group]

## TabSave  (None (Backstage View))
### 
- SaveGroup [taskFormGroup]
- ModernSaveGroup [group]
### SaveGroup
- Save [category]
-   None [task] < Save
-   None [group] < Save
-   AppGuardTrustPromotion [button] < Save
-   ButtonTaskDynamicServiceProvider [task] < Save
-     GroupDynamicServiceProvider [group] < Save
-     TabSave [control] < Save
-     SlabSaveAsAccessibility [group] < Save
-     None [group] < Save
-     GroupGettingStartedSaveAsMRUPlacesOrNavSlab [group] < Save
-     None [button] < Save
-     SkyDriveUpsellGroup [group] < Save
-     LoadingServicesGroup [group] < Save
-     GroupAddLocation [group] < Save
### ModernSaveGroup
- ButtonGroupsSave [control]
- ButtonGroupItemSaveCloudServices [control]
- SaveCloudServices [category]
-   None [task] < SaveCloudServices
-   None [group] < SaveCloudServices
-   AppGuardTrustPromotion [button] < SaveCloudServices
-   ButtonTaskDynamicServiceProvider [task] < SaveCloudServices
-     GroupDynamicSaveCloudServiceProvider [group] < SaveCloudServices
-     TabSave [control] < SaveCloudServices
-     SlabCloudSaveAsAccessibility [group] < SaveCloudServices
-     None [group] < SaveCloudServices
-     GroupGettingStartedSaveAsMRUCloudPlacesOrNavSlab [group] < SaveCloudServices
-     None [button] < SaveCloudServices
-   ButtonTaskDynamicServiceProvider [control] < SaveCloudServices
-     CloudSkyDriveUpsellGroup [group] < SaveCloudServices
-     LoadingCloudServicesGroup [group] < SaveCloudServices
- None [task]
- None [group]
- AppGuardTrustPromotion [button]
- ButtonTaskDynamicServiceProvider [task]
-   GroupDynamicSaveCloudServiceProvider [group] < ButtonTaskDynamicServiceProvider
-     TabSave [control] < ButtonTaskDynamicServiceProvider
-   SlabCloudSaveAsAccessibility [group] < ButtonTaskDynamicServiceProvider
-   None [group] < ButtonTaskDynamicServiceProvider
-   GroupGettingStartedSaveAsMRUCloudPlacesOrNavSlab [group] < ButtonTaskDynamicServiceProvider
-     None [button] < ButtonTaskDynamicServiceProvider
- GroupDynamicSaveCloudServiceProvider [group]
-   TabSave [control] < GroupDynamicSaveCloudServiceProvider
- TabSave [control]
- SlabCloudSaveAsAccessibility [group]
- None [group]
- GroupGettingStartedSaveAsMRUCloudPlacesOrNavSlab [group]
-   None [button] < GroupGettingStartedSaveAsMRUCloudPlacesOrNavSlab
- None [button]
- ButtonTaskDynamicServiceProvider [control]
- GroupDynamicSavePinnableCloudServiceProvider [group]
-   TabSave [control] < GroupDynamicSavePinnableCloudServiceProvider
- GroupPinnableCloudSaveAsAccessibility [group]
- None [group]
- GroupPinnableGettingStartedSaveAsMRUCloudPlacesOrNav [group]
-   None [button] < GroupPinnableGettingStartedSaveAsMRUCloudPlacesOrNav
-   CloudSkyDriveUpsellGroup [group] < ButtonTaskDynamicServiceProvider
-     None [button] < ButtonTaskDynamicServiceProvider
- CloudSkyDriveUpsellGroup [group]
-   None [button] < CloudSkyDriveUpsellGroup
-   LoadingCloudServicesGroup [group] < ButtonTaskDynamicServiceProvider
- LoadingCloudServicesGroup [group]
- ButtonGroupItemSaveLocalServices [control]
- LocalSave [category]
-   ButtonTaskDynamicServiceProvider [task] < LocalSave
-     GroupDynamicSaveLocalServiceProvider [group] < LocalSave
-     TabSave [control] < LocalSave
-     SlabLocalSaveAsAccessibility [group] < LocalSave
-     None [group] < LocalSave
-     GroupGettingStartedSaveAsMRULocalPlacesOrNavSlab [group] < LocalSave
-     None [button] < LocalSave
-     GroupAddLocationLocal [group] < LocalSave
-   GroupDynamicSaveLocalServiceProvider [group] < ButtonTaskDynamicServiceProvider
-     TabSave [control] < ButtonTaskDynamicServiceProvider
-   SlabLocalSaveAsAccessibility [group] < ButtonTaskDynamicServiceProvider
-   None [group] < ButtonTaskDynamicServiceProvider
-   GroupGettingStartedSaveAsMRULocalPlacesOrNavSlab [group] < ButtonTaskDynamicServiceProvider
-     None [button] < ButtonTaskDynamicServiceProvider
- GroupDynamicSaveLocalServiceProvider [group]
-   TabSave [control] < GroupDynamicSaveLocalServiceProvider
- SlabLocalSaveAsAccessibility [group]
- None [group]
- GroupGettingStartedSaveAsMRULocalPlacesOrNavSlab [group]
-   None [button] < GroupGettingStartedSaveAsMRULocalPlacesOrNavSlab
-   GroupAddLocationLocal [group] < ButtonTaskDynamicServiceProvider
- GroupAddLocationLocal [group]

## TabPrint  (None (Backstage View))
### 
- GroupPrintProtectedView [group]
- GroupPrintSettings [group]
- GroupPrintPreview [group]

## TabShare  (None (Backstage View))
### 
- GroupShare [taskFormGroup]
### GroupShare
- Share [category]
-   ShareWithPeopleFromSharePane [task] < Share
-     SlabShareWithPeopleFromSharePane [group] < Share
-   SendUsingEmail [task] < Share
-     GroupSendUsingEmail [group] < Share
-     GroupSendAsAttachment [group] < Share
-     GroupSendAsLink [group] < Share
-     GroupSendAsPdf [group] < Share
-     GroupSendAsXps [group] < Share
-     GroupSendAsInternetFax [group] < Share
-   AppShare [task] < Share
-     GroupAppShare [group] < Share
-   SendByInstantMessage [task] < Share
-     GroupSendByInstantMessage [group] < Share
- Workflow [category]
-   WorkflowStart [task] < Workflow
-     GroupWorkflowStart [group] < Workflow
-   WorkflowsUnavailable [task] < Workflow
-     GroupWorkflowsUnavailable [group] < Workflow

## TabPublish  (None (Backstage View))
### 
- PublishGroup [taskFormGroup]
### PublishGroup
- FileTypes [category]
-   FileSaveAsPdfOrXps [task] < FileTypes
-     GroupCreatePdfOrXpsDocumentSlab [group] < FileTypes
-     GroupCreatePdfOrXpsDocument [group] < FileTypes
-   FileTypeChange [task] < FileTypes
-     GroupFileSave [group] < FileTypes

## Publish2Tab  (None (Backstage View))
### 
- Publish2Group [taskFormGroup]
### Publish2Group
- Publish2ButtonCategory [category]
-   PublishToPowerBI [task] < Publish2ButtonCategory
-     PublishToPowerBI [group] < Publish2ButtonCategory
-     None [button] < Publish2ButtonCategory

## TabHelp  (None (Backstage View))
### 
- GroupBackstageIdentity [group]
- GroupBackstagePrivacy [group]
- Personalization [group]
- Theming [group]
- ManageServices [group]
- GroupBackstageSignInDisabled [group]
- GroupBackstageSignIn [group]
- GroupOfficeBranding [group]
- GroupClickToRunUpdateStatus [group]
- GroupUpdateChannel [group]
- GroupOfficeInsider [group]
- GroupAboutOfficeProducts [group]
- GroupWhatsNewInOfficeProducts [group]
### GroupBackstagePrivacy
- ManagePrivacy [button]
### GroupBackstageSignIn
- None [button]
### GroupUpdateChannel
- UpdateChannel [button]
### GroupOfficeInsider
- GroupOfficeInsiderRegistration [button]
- GroupOfficeInsiderChannel [button]
- GroupOfficeInsiderAbout [button]
- GroupOfficeInsiderHelp [button]

## TabOfficeFeedback  (None (Backstage View))
### 
- GroupOfficeSendSmile [group]
- GroupOfficeSendFrown [group]
- GroupOfficeSendSuggestion [group]

## None (Context Menu)  (None (Context Menu))
### 
- ContextMenuWorkbook [contextMenu]
- ContextMenuWorkbookPly [contextMenu]
- ContextMenuFormulaBar [contextMenu]
- ContextMenuFormulaDropdown [contextMenu]
- ContextMenuDesktop [contextMenu]
- ContextMenuTitleBar [contextMenu]
- ContextMenuPreviewer [contextMenu]
- ContextMenuCellLayout [contextMenu]
- ContextMenuCell [contextMenu]
- ContextMenuRowLayout [contextMenu]
- ContextMenuRow [contextMenu]
- ContextMenuColumnLayout [contextMenu]
- ContextMenuColumn [contextMenu]
- ContextMenuDialogPly [contextMenu]
- ContextMenuMacroCell [contextMenu]
- ContextMenuXmlRangeLayout [contextMenu]
- ContextMenuXmlRange [contextMenu]
- ContextMenuListRangeLayout [contextMenu]
- ContextMenuListRange [contextMenu]
- ContextMenuPivotTable [contextMenu]
- ContextMenuDrawnObject [contextMenu]
- ContextMenuActiveXControl [contextMenu]
- ContextMenuFormControl [contextMenu]
- ContextMenuOleObject [contextMenu]
- ContextMenuCurve [contextMenu]
- ContextMenuCurveNode [contextMenu]
- ContextMenuCurveSegment [contextMenu]
- ContextMenuConnectorClassic [contextMenu]
- ContextMenuQueryLayout [contextMenu]
- ContextMenuQuery [contextMenu]
- ContextMenuAutoFill [contextMenu]
- ContextMenuRightDragDrop [contextMenu]
- ContextMenuPhoneticEdit [contextMenu]
- ContextMenuPictureClassic [contextMenu]
- ContextMenuRowTouch [contextMenu]
- ContextMenuColumnTouch [contextMenu]
- ContextMenuPivotTableTouch [contextMenu]
- ContextMenuPivotTableSimpleTouch [contextMenu]
- ContextMenuPicture [contextMenu]
- ContextMenuGraphic [contextMenu]
- ContextMenu3DModel [contextMenu]
- ContextMenuShape [contextMenu]
- ContextMenuInk [contextMenu]
- ContextMenuObjectsGroup [contextMenu]
- ContextMenuObjectEditPoint [contextMenu]
- ContextMenuObjectEditSegment [contextMenu]
- ContextMenuTextEdit [contextMenu]
- ContextMenuShapeConnector [contextMenu]
- ContextMenuShapeFreeform [contextMenu]
- ContextMenuChartArea [contextMenu]
- ContextMenuChartAxis [contextMenu]
- ContextMenuChartAxisLabel [contextMenu]
- ContextMenuChartAxisTitle [contextMenu]
- ContextMenuChartBackWall [contextMenu]
- ContextMenuChartTitle [contextMenu]
- ContextMenuChartDataLabel [contextMenu]
- ContextMenuChartDataLabels [contextMenu]
- ContextMenuChartDataPoint [contextMenu]
- ContextMenuChartDataSeries [contextMenu]
- ContextMenuChartParetoLine [contextMenu]
- ContextMenuChartDataTable [contextMenu]
- ContextMenuChartDisplayUnit [contextMenu]
- ContextMenuChartDownBars [contextMenu]
- ContextMenuChartDropLines [contextMenu]
- ContextMenuChartErrorBars [contextMenu]
- ContextMenuChartFloor [contextMenu]
- ContextMenuChartGridlines [contextMenu]
- ContextMenuChartHighLowLine [contextMenu]
- ContextMenuChartLeaderLines [contextMenu]
- ContextMenuChartLegend [contextMenu]
- ContextMenuChartLegendEntry [contextMenu]
- ContextMenuChartPlotArea [contextMenu]
- ContextMenuChartSeriesLine [contextMenu]
- ContextMenuChartSideWall [contextMenu]
- ContextMenuChartTrendline [contextMenu]
- ContextMenuChartTrendlineLabel [contextMenu]
- ContextMenuChartUpBars [contextMenu]
- ContextMenuChartWalls [contextMenu]
- ContextMenuGraphicsCompatibility [contextMenu]
- ContextMenuSmartArtContentPane [contextMenu]
- ContextMenuSmartArtBackground [contextMenu]
- ContextMenuSmartArtEditSmartArt [contextMenu]
- ContextMenuSmartArtEdit1DShape [contextMenu]
- ContextMenuSmartArtEditText [contextMenu]
- ContextMenuSmartArtEdit1DShapeText [contextMenu]
- ContextMenuTimeSlicerLevel [contextMenu]
- ContextMenuTimeSlicer [contextMenu]
- ContextMenuSlicer [contextMenu]
- ContextMenuOfficeWebExtension [contextMenu]
- None [contextMenu]
- None [gallery]
- None [toggleButton]
- None [toggleButton]
### ContextMenuWorkbookPly
- SheetInsertPage [button]
- SheetDelete [button]
- SheetRename [button]
- SheetMoveOrCopy [button]
- ViewVisualBasicCode [button]
- SheetProtect [button]
- SheetTabColorGallery [gallery]
-   SheetTabColorMoreColorsDialog [button] < SheetTabColorGallery
- SheetHide [button]
- SheetUnhide [button]
- SelectAllSheets [button]
- UngroupSheets [button]
- LinkToSheet [button]
- LineageShowChangesFromSheetTab [button]
### ContextMenuFormulaDropdown
- PythonObject [toggleButton]
- ConvertToExcelValues [toggleButton]
### ContextMenuPicture
- Cut [button]
- Copy [button]
- PasteGalleryMini [gallery]
- PicturePlaceInCell [button]
- PictureEdit [button]
- PictureChange [button]
- PictureChangeMenu [menu]
-   PictureChangeFromFile [button] < PictureChangeMenu
-   PictureChangeFromMobile [button] < PictureChangeMenu
-   PictureChangeFromM365 [button] < PictureChangeMenu
-   PictureChangeFromOnline [button] < PictureChangeMenu
-   ChangeIcon [button] < PictureChangeMenu
-   PictureChangeFromClipboard [button] < PictureChangeMenu
- ObjectsGroupMenu [menu]
-   ObjectsGroup [button] < ObjectsGroupMenu
-   ObjectsRegroup [button] < ObjectsGroupMenu
-   ObjectsUngroup [button] < ObjectsGroupMenu
- ObjectBringToFrontMenu [splitButton]
-   ObjectBringToFront [button] < ObjectBringToFrontMenu
-   ObjectBringForward [button] < ObjectBringToFrontMenu
- ObjectSendToBackMenu [splitButton]
-   ObjectSendToBack [button] < ObjectSendToBackMenu
-   ObjectSendBackward [button] < ObjectSendToBackMenu
- InsertLinkGallery [gallery]
-   None [button] < InsertLinkGallery
- HyperlinkOpenInMenu [menu]
-   HyperlinkOpenInWeb [button] < HyperlinkOpenInMenu
-   HyperlinkOpenInNative [button] < HyperlinkOpenInMenu
- HyperlinkOpenDefaultSettingMenu [menu]
-   OpenDefaultWeb [button] < HyperlinkOpenDefaultSettingMenu
-   OpenDefaultNative [button] < HyperlinkOpenDefaultSettingMenu
- HyperlinkInsert [button]
- HyperlinkEdit [button]
- HyperlinkOpenExcel [button]
- HyperlinkRemove [button]
- ObjectSaveAsPicture [button]
- PictureAttributionNavigateToSource [button]
- MacroAssign [button]
- AutoAlternativeText [button]
- ObjectSizeAndPropertiesDialog [button]
- ObjectFormatDialog [button]
### ContextMenuGraphic
- Cut [button]
- Copy [button]
- PasteGalleryMini [gallery]
- SVGEdit [button]
- GraphicChangeMenu [menu]
-   PictureChangeFromFile [button] < GraphicChangeMenu
-   PictureChangeFromMobile [button] < GraphicChangeMenu
-   PictureChangeFromM365 [button] < GraphicChangeMenu
-   PictureChangeFromOnline [button] < GraphicChangeMenu
-   ChangeIcon [button] < GraphicChangeMenu
-   PictureChangeFromClipboard [button] < GraphicChangeMenu
- ObjectsGroupMenu [menu]
-   ObjectsGroup [button] < ObjectsGroupMenu
-   ObjectsRegroup [button] < ObjectsGroupMenu
-   ObjectsUngroup [button] < ObjectsGroupMenu
- ObjectBringToFrontMenu [splitButton]
-   ObjectBringToFront [button] < ObjectBringToFrontMenu
-   ObjectBringForward [button] < ObjectBringToFrontMenu
- ObjectSendToBackMenu [splitButton]
-   ObjectSendToBack [button] < ObjectSendToBackMenu
-   ObjectSendBackward [button] < ObjectSendToBackMenu
- InsertLinkGallery [gallery]
-   None [button] < InsertLinkGallery
- HyperlinkOpenInMenu [menu]
-   HyperlinkOpenInWeb [button] < HyperlinkOpenInMenu
-   HyperlinkOpenInNative [button] < HyperlinkOpenInMenu
- HyperlinkOpenDefaultSettingMenu [menu]
-   OpenDefaultWeb [button] < HyperlinkOpenDefaultSettingMenu
-   OpenDefaultNative [button] < HyperlinkOpenDefaultSettingMenu
- HyperlinkInsert [button]
- HyperlinkEdit [button]
- HyperlinkOpenExcel [button]
- HyperlinkRemove [button]
- ObjectSaveAsPicture [button]
- MacroAssign [button]
- AutoAlternativeText [button]
- ObjectSizeAndPropertiesDialog [button]
- GraphicsFormatDialog [button]
### ContextMenu3DModel
- Cut [button]
- Copy [button]
- PasteGalleryMini [gallery]
- ObjectsGroupMenu [menu]
-   ObjectsGroup [button] < ObjectsGroupMenu
-   ObjectsRegroup [button] < ObjectsGroupMenu
-   ObjectsUngroup [button] < ObjectsGroupMenu
- ObjectBringToFrontMenu [splitButton]
-   ObjectBringToFront [button] < ObjectBringToFrontMenu
-   ObjectBringForward [button] < ObjectBringToFrontMenu
- ObjectSendToBackMenu [splitButton]
-   ObjectSendToBack [button] < ObjectSendToBackMenu
-   ObjectSendBackward [button] < ObjectSendToBackMenu
- InsertLinkGallery [gallery]
-   None [button] < InsertLinkGallery
- HyperlinkOpenInMenu [menu]
-   HyperlinkOpenInWeb [button] < HyperlinkOpenInMenu
-   HyperlinkOpenInNative [button] < HyperlinkOpenInMenu
- HyperlinkOpenDefaultSettingMenu [menu]
-   OpenDefaultWeb [button] < HyperlinkOpenDefaultSettingMenu
-   OpenDefaultNative [button] < HyperlinkOpenDefaultSettingMenu
- HyperlinkInsert [button]
- HyperlinkEdit [button]
- HyperlinkOpenExcel [button]
- HyperlinkRemove [button]
- ObjectSaveAsPicture [button]
- PictureAttributionNavigateToSource [button]
- MacroAssign [button]
- AutoAlternativeText [button]
- ObjectSizeAndPropertiesDialog [button]
- FormatDialog3DModel [button]
### ContextMenuShape
- Cut [button]
- Copy [button]
- PasteGalleryMini [gallery]
- ObjectAddText [button]
- ObjectEditText [button]
- ObjectEditPoints [toggleButton]
- ObjectsGroupMenu [menu]
-   ObjectsGroup [button] < ObjectsGroupMenu
-   ObjectsRegroup [button] < ObjectsGroupMenu
-   ObjectsUngroup [button] < ObjectsGroupMenu
- ObjectBringToFrontMenu [splitButton]
-   ObjectBringToFront [button] < ObjectBringToFrontMenu
-   ObjectBringForward [button] < ObjectBringToFrontMenu
- ObjectSendToBackMenu [splitButton]
-   ObjectSendToBack [button] < ObjectSendToBackMenu
-   ObjectSendBackward [button] < ObjectSendToBackMenu
- InsertLinkGallery [gallery]
-   None [button] < InsertLinkGallery
- HyperlinkOpenInMenu [menu]
-   HyperlinkOpenInWeb [button] < HyperlinkOpenInMenu
-   HyperlinkOpenInNative [button] < HyperlinkOpenInMenu
- HyperlinkOpenDefaultSettingMenu [menu]
-   OpenDefaultWeb [button] < HyperlinkOpenDefaultSettingMenu
-   OpenDefaultNative [button] < HyperlinkOpenDefaultSettingMenu
- HyperlinkInsert [button]
- HyperlinkEdit [button]
- HyperlinkOpenExcel [button]
- HyperlinkRemove [button]
- ObjectSaveAsPicture [button]
- Insights [button]
- ViewScript [button]
- MacroAssign [button]
- AutoAlternativeText [button]
- ObjectSetShapeDefaults [button]
- ObjectSizeAndPropertiesDialog [button]
- ObjectFormatDialog [button]
- PictureAttributionNavigateToSource [button]
### ContextMenuInk
- Cut [button]
- Copy [button]
- InkCopyAsText [button]
- PasteGalleryMini [gallery]
- ObjectsGroupMenu [menu]
-   ObjectsGroup [button] < ObjectsGroupMenu
-   ObjectsRegroup [button] < ObjectsGroupMenu
-   ObjectsUngroup [button] < ObjectsGroupMenu
- ObjectBringToFrontMenu [splitButton]
-   ObjectBringToFront [button] < ObjectBringToFrontMenu
-   ObjectBringForward [button] < ObjectBringToFrontMenu
- ObjectSendToBackMenu [splitButton]
-   ObjectSendToBack [button] < ObjectSendToBackMenu
-   ObjectSendBackward [button] < ObjectSendToBackMenu
- ObjectSaveAsPicture [button]
- MacroAssign [button]
- AutoAlternativeText [button]
- ObjectSetShapeDefaults [button]
- ObjectSizeAndPropertiesDialog [button]
- ObjectFormatDialog [button]
### ContextMenuObjectsGroup
- Cut [button]
- Copy [button]
- PasteGalleryMini [gallery]
- ObjectsGroupMenu [menu]
-   ObjectsGroup [button] < ObjectsGroupMenu
-   ObjectsRegroup [button] < ObjectsGroupMenu
-   ObjectsUngroup [button] < ObjectsGroupMenu
- ObjectBringToFrontMenu [splitButton]
-   ObjectBringToFront [button] < ObjectBringToFrontMenu
-   ObjectBringForward [button] < ObjectBringToFrontMenu
- ObjectSendToBackMenu [splitButton]
-   ObjectSendToBack [button] < ObjectSendToBackMenu
-   ObjectSendBackward [button] < ObjectSendToBackMenu
- InsertLinkGallery [gallery]
-   None [button] < InsertLinkGallery
- HyperlinkOpenInMenu [menu]
-   HyperlinkOpenInWeb [button] < HyperlinkOpenInMenu
-   HyperlinkOpenInNative [button] < HyperlinkOpenInMenu
- HyperlinkOpenDefaultSettingMenu [menu]
-   OpenDefaultWeb [button] < HyperlinkOpenDefaultSettingMenu
-   OpenDefaultNative [button] < HyperlinkOpenDefaultSettingMenu
- HyperlinkInsert [button]
- HyperlinkEdit [button]
- HyperlinkOpenExcel [button]
- HyperlinkRemove [button]
- ObjectSaveAsPicture [button]
- MacroAssign [button]
- AutoAlternativeText [button]
- ObjectSetShapeDefaults [button]
- ObjectSizeAndPropertiesDialog [button]
- ObjectFormatDialog [button]
### ContextMenuObjectEditPoint
- AddNode [button]
- DeleteNode [button]
- OpenPath [toggleButton]
- ClosePath [toggleButton]
- SmoothVertex [toggleButton]
- StraightVertex [toggleButton]
- CornerVertex [toggleButton]
- ExitEditCurve [button]
### ContextMenuObjectEditSegment
- AddNode [button]
- DeleteSegment [button]
- OpenPath [toggleButton]
- ClosePath [toggleButton]
- MakeSegmentStraight [button]
- MakeSegmentCurved [button]
- ExitEditCurve [button]
### ContextMenuTextEdit
- FirstCorrectionAlternative [button]
- SecondCorrectionAlternative [button]
- ThirdCorrectionAlternative [button]
- FourthCorrectionAlternative [button]
- FifthCorrectionAlternative [button]
- MoreCorrectionAlternatives [button]
- Cut [button]
- Copy [button]
- PasteGalleryMini [gallery]
- EquationOptionsGallery [gallery]
- TextEditModeExit [button]
- FontDialog [button]
- ParagraphDialog [button]
- BulletsGallery [gallery]
-   BulletsAndNumberingBulletsDialog [button] < BulletsGallery
- ImeReconvert [button]
- InsertLinkGallery [gallery]
-   None [button] < InsertLinkGallery
- HyperlinkOpenInMenu [menu]
-   HyperlinkOpenInWeb [button] < HyperlinkOpenInMenu
-   HyperlinkOpenInNative [button] < HyperlinkOpenInMenu
- HyperlinkOpenDefaultSettingMenu [menu]
-   OpenDefaultWeb [button] < HyperlinkOpenDefaultSettingMenu
-   OpenDefaultNative [button] < HyperlinkOpenDefaultSettingMenu
- HyperlinkInsert [button]
- HyperlinkEdit [button]
- HyperlinkOpenExcel [button]
- HyperlinkRemove [button]
- Insights [button]
- ViewScript [button]
- MacroAssign [button]
- WordArtFormatDialog [button]
- ObjectFormatDialog [button]
### ContextMenuShapeConnector
- Cut [button]
- Copy [button]
- PasteGalleryMini [gallery]
- ShapeConnectorStyleMenu [menu]
-   ShapeConnectorStyleStraight [toggleButton] < ShapeConnectorStyleMenu
-   ShapeConnectorStyleElbow [toggleButton] < ShapeConnectorStyleMenu
-   ShapeConnectorStyleCurved [toggleButton] < ShapeConnectorStyleMenu
- ShapeRerouteConnectors [toggleButton]
- ObjectsGroupMenu [menu]
-   ObjectsGroup [button] < ObjectsGroupMenu
-   ObjectsRegroup [button] < ObjectsGroupMenu
-   ObjectsUngroup [button] < ObjectsGroupMenu
- ObjectBringToFrontMenu [splitButton]
-   ObjectBringToFront [button] < ObjectBringToFrontMenu
-   ObjectBringForward [button] < ObjectBringToFrontMenu
- ObjectSendToBackMenu [splitButton]
-   ObjectSendToBack [button] < ObjectSendToBackMenu
-   ObjectSendBackward [button] < ObjectSendToBackMenu
- InsertLinkGallery [gallery]
-   None [button] < InsertLinkGallery
- HyperlinkOpenInMenu [menu]
-   HyperlinkOpenInWeb [button] < HyperlinkOpenInMenu
-   HyperlinkOpenInNative [button] < HyperlinkOpenInMenu
- HyperlinkOpenDefaultSettingMenu [menu]
-   OpenDefaultWeb [button] < HyperlinkOpenDefaultSettingMenu
-   OpenDefaultNative [button] < HyperlinkOpenDefaultSettingMenu
- HyperlinkInsert [button]
- HyperlinkEdit [button]
- HyperlinkOpenExcel [button]
- HyperlinkRemove [button]
- ObjectSaveAsPicture [button]
- MacroAssign [button]
- AutoAlternativeText [button]
- ObjectSetShapeDefaults [button]
- ObjectSizeAndPositionDialog [button]
- ObjectFormatDialog [button]
### ContextMenuShapeFreeform
- Cut [button]
- Copy [button]
- PasteGalleryMini [gallery]
- ObjectAddText [button]
- ObjectEditText [button]
- ObjectEditPoints [toggleButton]
- ObjectsGroupMenu [menu]
-   ObjectsGroup [button] < ObjectsGroupMenu
-   ObjectsRegroup [button] < ObjectsGroupMenu
-   ObjectsUngroup [button] < ObjectsGroupMenu
- ObjectBringToFrontMenu [splitButton]
-   ObjectBringToFront [button] < ObjectBringToFrontMenu
-   ObjectBringForward [button] < ObjectBringToFrontMenu
- ObjectSendToBackMenu [splitButton]
-   ObjectSendToBack [button] < ObjectSendToBackMenu
-   ObjectSendBackward [button] < ObjectSendToBackMenu
- InsertLinkGallery [gallery]
-   None [button] < InsertLinkGallery
- HyperlinkOpenInMenu [menu]
-   HyperlinkOpenInWeb [button] < HyperlinkOpenInMenu
-   HyperlinkOpenInNative [button] < HyperlinkOpenInMenu
- HyperlinkOpenDefaultSettingMenu [menu]
-   OpenDefaultWeb [button] < HyperlinkOpenDefaultSettingMenu
-   OpenDefaultNative [button] < HyperlinkOpenDefaultSettingMenu
- HyperlinkInsert [button]
- HyperlinkEdit [button]
- HyperlinkOpenExcel [button]
- HyperlinkRemove [button]
- ObjectSaveAsPicture [button]
- MacroAssign [button]
- AutoAlternativeText [button]
- ObjectSetShapeDefaults [button]
- ObjectSizeAndPositionDialog [button]
- ObjectFormatDialog [button]
- PictureAttributionNavigateToSource [button]
### ContextMenuChartArea
- PivotChartRefresh [button]
- Cut [button]
- Copy [button]
- PasteGalleryMini [gallery]
- None [button]
- ChartResetToMatchStyle [button]
- FontDialog [button]
- ChartChangeType [button]
- ChartSaveTemplates [button]
- ChartEditDataSource [button]
- ChartPlacement [button]
- Chart3DView [button]
- ObjectsGroupMenu [menu]
-   ObjectsGroup [button] < ObjectsGroupMenu
-   ObjectsRegroup [button] < ObjectsGroupMenu
-   ObjectsUngroup [button] < ObjectsGroupMenu
- ObjectBringToFrontMenu [splitButton]
-   ObjectBringToFront [button] < ObjectBringToFrontMenu
-   ObjectBringForward [button] < ObjectBringToFrontMenu
- ObjectSendToBackMenu [splitButton]
-   ObjectSendToBack [button] < ObjectSendToBackMenu
-   ObjectSendBackward [button] < ObjectSendToBackMenu
- ObjectSaveAsPicture [button]
- MacroAssign [button]
- AutoAlternativeText [button]
- ChartFormatChartArea [button]
- PictureAttributionNavigateToSource [button]
- PivotChartOptionsPopup [button]
- PivotFieldListShowHideContextual [toggleButton]
- ChartShowBigPicture [button]
### ContextMenuChartAxis
- ChartSortRightClickMenu [menu]
-   ChartSortAscending [button] < ChartSortRightClickMenu
-   ChartSortDescending [button] < ChartSortRightClickMenu
-   ChartSortReset [button] < ChartSortRightClickMenu
- ChartDelete [button]
- ChartResetToMatchStyle [button]
- FontDialog [button]
- ChartChangeType [button]
- ChartSourceDataContextualMenu [button]
- ChartEditDataMenu [splitButton]
-   ChartShowDataGrid [button] < ChartEditDataMenu
-   ChartShowDataContextualMenu [button] < ChartEditDataMenu
- Chart3DView [button]
- ChartMajorGridlinesAdd [button]
- ChartMinorGridlinesAdd [button]
- ChartFormatMajorGridlinesAndSelect [button]
- ChartFormatMinorGridlinesAndSelect [button]
- ChartFormatAxis [button]
- PictureAttributionNavigateToSource [button]
- None [button]
### ContextMenuChartAxisLabel
- PivotTableSortMenu [menu]
- PivotTableFilterMenu [menu]
- ChartShowHideLevelsMenu [menu]
- ChartExpandCollapseMenu [menu]
- ChartNavDrillDownDrillUpMenu [menu]
- ChartQuickExplore [button]
- PivotTableReorderingMenu [menu]
- ChartPropertiesInTooltips [menu]
- PivotActionPopup [menu]
- ChartDelete [button]
- ChartResetToMatchStyle [button]
- FontDialog [button]
- ChartChangeType [button]
- ChartSourceDataContextualMenu [button]
- ChartEditDataMenu [splitButton]
-   ChartShowDataGrid [button] < ChartEditDataMenu
-   ChartShowDataContextualMenu [button] < ChartEditDataMenu
- Chart3DView [button]
- ChartMajorGridlinesAdd [button]
- ChartMinorGridlinesAdd [button]
- ChartFormatMajorGridlinesAndSelect [button]
- ChartFormatMinorGridlinesAndSelect [button]
- ChartFormatAxis [button]
### ContextMenuChartAxisTitle
- FirstCorrectionAlternative [button]
- SecondCorrectionAlternative [button]
- ThirdCorrectionAlternative [button]
- FourthCorrectionAlternative [button]
- FifthCorrectionAlternative [button]
- MoreCorrectionAlternatives [button]
- ChartDelete [button]
- ChartResetToMatchStyle [button]
- ChartEditText [button]
- ChartExitEditText [button]
- FontDialog [button]
- ChartChangeType [button]
- ChartSourceDataContextualMenu [button]
- ChartEditDataMenu [splitButton]
-   ChartShowDataGrid [button] < ChartEditDataMenu
-   ChartShowDataContextualMenu [button] < ChartEditDataMenu
- Chart3DView [button]
- ChartFormatAxisTitle [button]
- None [button]
### ContextMenuChartBackWall
- ChartDelete [button]
- ChartResetToMatchStyle [button]
- ChartChangeType [button]
- ChartSourceDataContextualMenu [button]
- ChartEditDataMenu [splitButton]
-   ChartShowDataGrid [button] < ChartEditDataMenu
-   ChartShowDataContextualMenu [button] < ChartEditDataMenu
- Chart3DView [button]
- ChartFormatBackWall [button]
- PictureAttributionNavigateToSource [button]
- None [button]
### ContextMenuChartTitle
- FirstCorrectionAlternative [button]
- SecondCorrectionAlternative [button]
- ThirdCorrectionAlternative [button]
- FourthCorrectionAlternative [button]
- FifthCorrectionAlternative [button]
- MoreCorrectionAlternatives [button]
- ChartDelete [button]
- ChartResetToMatchStyle [button]
- ChartEditText [button]
- ChartExitEditText [button]
- FontDialog [button]
- ChartChangeType [button]
- ChartSourceDataContextualMenu [button]
- ChartEditDataMenu [splitButton]
-   ChartShowDataGrid [button] < ChartEditDataMenu
-   ChartShowDataContextualMenu [button] < ChartEditDataMenu
- Chart3DView [button]
- ChartFormatChartTitle [button]
- PictureAttributionNavigateToSource [button]
- None [button]
### ContextMenuChartDataLabel
- FirstCorrectionAlternative [button]
- SecondCorrectionAlternative [button]
- ThirdCorrectionAlternative [button]
- FourthCorrectionAlternative [button]
- FifthCorrectionAlternative [button]
- MoreCorrectionAlternatives [button]
- ChartDelete [button]
- ChartResetToMatchStyle [button]
- ChartEditText [button]
- ChartExitEditText [button]
- FontDialog [button]
- ChartInsertDataField [button]
- ChartChangeType [button]
- ChartSourceDataContextualMenu [button]
- ChartEditDataMenu [splitButton]
-   ChartShowDataGrid [button] < ChartEditDataMenu
-   ChartShowDataContextualMenu [button] < ChartEditDataMenu
- Chart3DView [button]
- ChartFormatDataPointAndSelect [button]
- ChartChangeDataLabelShapeGallery [gallery]
- ChartFormatDataLabel [button]
- PictureAttributionNavigateToSource [button]
- None [button]
### ContextMenuChartDataLabels
- ChartDelete [button]
- ChartResetToMatchStyle [button]
- FontDialog [button]
- ChartChangeType [button]
- ChartSourceDataContextualMenu [button]
- ChartEditDataMenu [splitButton]
-   ChartShowDataGrid [button] < ChartEditDataMenu
-   ChartShowDataContextualMenu [button] < ChartEditDataMenu
- Chart3DView [button]
- ChartFormatDataSeriesAndSelect [button]
- ChartChangeDataLabelShapesGallery [gallery]
- ChartFormatDataLabels [button]
- None [button]
### ContextMenuChartDataPoint
- PivotTableSortMenu [menu]
- PivotTableFilterMenu [menu]
- ChartExpandCollapseMenu [menu]
- ChartNavDrillDownDrillUpMenu [menu]
- ChartQuickExplore [button]
- PivotChartDatapointShowDetail [button]
- PivotActionPopup [menu]
- ChartDeleteSeries [button]
- ChartResetToMatchStyle [button]
- ChartSeriesTypeChange [button]
- ChartSourceDataContextualMenu [button]
- ChartEditDataMenu [splitButton]
-   ChartShowDataGrid [button] < ChartEditDataMenu
-   ChartShowDataContextualMenu [button] < ChartEditDataMenu
- Chart3DView [button]
- AddDataLabelShapesMenu [splitButton]
-   ChartDataLabelAdd [button] < AddDataLabelShapesMenu
-   ChartCalloutDataLabelAdd [button] < AddDataLabelShapesMenu
- ChartTrendlineAdd [button]
- ChartFormatDataLabelAndSelect [button]
- ChartWaterfallSetAsTotal [button]
- ChartWaterfallUnsetTotal [button]
- ChartFormatDataPoint [button]
- PictureAttributionNavigateToSource [button]
- None [button]
### ContextMenuChartDataSeries
- ChartSortRightClickMenu [menu]
-   ChartSortAscending [button] < ChartSortRightClickMenu
-   ChartSortDescending [button] < ChartSortRightClickMenu
-   ChartSortReset [button] < ChartSortRightClickMenu
- PivotTableSortMenu [menu]
- PivotTableFilterMenu [menu]
- ChartShowHideLevelsMenu [menu]
- ChartExpandCollapseMenu [menu]
- ChartNavDrillDownDrillUpMenu [menu]
- ChartQuickExplore [button]
- ChartPropertiesInTooltips [menu]
- PivotActionPopup [menu]
- ChartDelete [button]
- ChartResetToMatchStyle [button]
- ChartSeriesTypeChange [button]
- ChartSourceDataContextualMenu [button]
- ChartEditDataMenu [splitButton]
-   ChartShowDataGrid [button] < ChartEditDataMenu
-   ChartShowDataContextualMenu [button] < ChartEditDataMenu
- Chart3DView [button]
- AddDataLabelsShapesMenu [splitButton]
-   ChartDataLabelsAdd [button] < AddDataLabelsShapesMenu
-   ChartCalloutDataLabelsAdd [button] < AddDataLabelsShapesMenu
- ChartTrendlineAdd [button]
- ChartFormatDataLabelsAndSelect [button]
- ChartFormatDataSeries [button]
- PictureAttributionNavigateToSource [button]
- None [button]
### ContextMenuChartParetoLine
- ChartFormatParetoLine [button]
- None [button]
### ContextMenuChartDataTable
- ChartDelete [button]
- ChartResetToMatchStyle [button]
- FontDialog [button]
- ChartChangeType [button]
- ChartSourceDataContextualMenu [button]
- ChartEditDataMenu [splitButton]
-   ChartShowDataGrid [button] < ChartEditDataMenu
-   ChartShowDataContextualMenu [button] < ChartEditDataMenu
- Chart3DView [button]
- ChartFormatDataTable [button]
- ChartFormatAxisAndSelect [button]
- None [button]
### ContextMenuChartDisplayUnit
- FirstCorrectionAlternative [button]
- SecondCorrectionAlternative [button]
- ThirdCorrectionAlternative [button]
- FourthCorrectionAlternative [button]
- FifthCorrectionAlternative [button]
- MoreCorrectionAlternatives [button]
- ChartDelete [button]
- ChartResetToMatchStyle [button]
- ChartEditText [button]
- ChartExitEditText [button]
- FontDialog [button]
- ChartChangeType [button]
- ChartSourceDataContextualMenu [button]
- ChartEditDataMenu [splitButton]
-   ChartShowDataGrid [button] < ChartEditDataMenu
-   ChartShowDataContextualMenu [button] < ChartEditDataMenu
- Chart3DView [button]
- ChartFormatDisplayUnit [button]
- ChartFormatAxisAndSelect [button]
- ChartAddAxis [button]
- None [button]
### ContextMenuChartDownBars
- ChartDelete [button]
- ChartResetToMatchStyle [button]
- ChartChangeType [button]
- ChartSourceDataContextualMenu [button]
- ChartEditDataMenu [splitButton]
-   ChartShowDataGrid [button] < ChartEditDataMenu
-   ChartShowDataContextualMenu [button] < ChartEditDataMenu
- Chart3DView [button]
- ChartFormatDownBars [button]
- None [button]
### ContextMenuChartDropLines
- ChartDelete [button]
- ChartResetToMatchStyle [button]
- ChartChangeType [button]
- ChartSourceDataContextualMenu [button]
- ChartEditDataMenu [splitButton]
-   ChartShowDataGrid [button] < ChartEditDataMenu
-   ChartShowDataContextualMenu [button] < ChartEditDataMenu
- Chart3DView [button]
- ChartFormatDropLines [button]
- None [button]
### ContextMenuChartErrorBars
- ChartDelete [button]
- ChartResetToMatchStyle [button]
- ChartChangeType [button]
- ChartSourceDataContextualMenu [button]
- ChartEditDataMenu [splitButton]
-   ChartShowDataGrid [button] < ChartEditDataMenu
-   ChartShowDataContextualMenu [button] < ChartEditDataMenu
- Chart3DView [button]
- ChartFormatErrorBars [button]
- None [button]
### ContextMenuChartFloor
- ChartDelete [button]
- ChartResetToMatchStyle [button]
- ChartChangeType [button]
- ChartSourceDataContextualMenu [button]
- ChartEditDataMenu [splitButton]
-   ChartShowDataGrid [button] < ChartEditDataMenu
-   ChartShowDataContextualMenu [button] < ChartEditDataMenu
- Chart3DView [button]
- ChartFormatFloor [button]
- None [button]
### ContextMenuChartGridlines
- ChartDelete [button]
- ChartResetToMatchStyle [button]
- ChartChangeType [button]
- ChartSourceDataContextualMenu [button]
- ChartEditDataMenu [splitButton]
-   ChartShowDataGrid [button] < ChartEditDataMenu
-   ChartShowDataContextualMenu [button] < ChartEditDataMenu
- Chart3DView [button]
- ChartFormatGridlines [button]
- ChartFormatAxisAndSelect [button]
- ChartAddAxis [button]
- None [button]
### ContextMenuChartHighLowLine
- ChartDelete [button]
- ChartResetToMatchStyle [button]
- ChartChangeType [button]
- ChartSourceDataContextualMenu [button]
- ChartEditDataMenu [splitButton]
-   ChartShowDataGrid [button] < ChartEditDataMenu
-   ChartShowDataContextualMenu [button] < ChartEditDataMenu
- Chart3DView [button]
- ChartFormatHighLowLine [button]
- None [button]
### ContextMenuChartLeaderLines
- ChartDelete [button]
- ChartResetToMatchStyle [button]
- ChartChangeType [button]
- ChartSourceDataContextualMenu [button]
- ChartEditDataMenu [splitButton]
-   ChartShowDataGrid [button] < ChartEditDataMenu
-   ChartShowDataContextualMenu [button] < ChartEditDataMenu
- Chart3DView [button]
- ChartFormatLeaderLines [button]
- None [button]
### ContextMenuChartLegend
- ChartDelete [button]
- ChartResetToMatchStyle [button]
- FontDialog [button]
- ChartChangeType [button]
- ChartSourceDataContextualMenu [button]
- ChartEditDataMenu [splitButton]
-   ChartShowDataGrid [button] < ChartEditDataMenu
-   ChartShowDataContextualMenu [button] < ChartEditDataMenu
- Chart3DView [button]
- ChartFormatLegend [button]
- PictureAttributionNavigateToSource [button]
- None [button]
### ContextMenuChartLegendEntry
- PivotTableSortMenu [menu]
- PivotTableFilterMenu [menu]
- ChartShowHideLevelsMenu [menu]
- ChartExpandCollapseMenu [menu]
- ChartNavDrillDownDrillUpMenu [menu]
- ChartQuickExplore [button]
- PivotTableReorderingMenu [menu]
- ChartPropertiesInTooltips [menu]
- PivotActionPopup [menu]
- ChartDelete [button]
- ChartDeleteSeries [button]
- ChartResetToMatchStyle [button]
- FontDialog [button]
- ChartSeriesTypeChange [button]
- ChartSourceDataContextualMenu [button]
- ChartEditDataMenu [splitButton]
-   ChartShowDataGrid [button] < ChartEditDataMenu
-   ChartShowDataContextualMenu [button] < ChartEditDataMenu
- Chart3DView [button]
- ChartFormatDataSeriesAndSelect [button]
- ChartFormatLegendEntry [button]
- None [button]
### ContextMenuChartPlotArea
- ChartDelete [button]
- ChartResetToMatchStyle [button]
- ChartChangeType [button]
- ChartSaveTemplates [button]
- ChartSourceDataContextualMenu [button]
- ChartEditDataMenu [splitButton]
-   ChartShowDataGrid [button] < ChartEditDataMenu
-   ChartShowDataContextualMenu [button] < ChartEditDataMenu
- Chart3DView [button]
- ChartFormatPlotArea [button]
- PictureAttributionNavigateToSource [button]
- None [button]
### ContextMenuChartSeriesLine
- ChartDelete [button]
- ChartResetToMatchStyle [button]
- ChartChangeType [button]
- ChartSourceDataContextualMenu [button]
- ChartEditDataMenu [splitButton]
-   ChartShowDataGrid [button] < ChartEditDataMenu
-   ChartShowDataContextualMenu [button] < ChartEditDataMenu
- Chart3DView [button]
- ChartFormatSeriesLine [button]
- None [button]
### ContextMenuChartSideWall
- ChartDelete [button]
- ChartResetToMatchStyle [button]
- ChartChangeType [button]
- ChartSourceDataContextualMenu [button]
- ChartEditDataMenu [splitButton]
-   ChartShowDataGrid [button] < ChartEditDataMenu
-   ChartShowDataContextualMenu [button] < ChartEditDataMenu
- Chart3DView [button]
- ChartFormatSideWall [button]
- PictureAttributionNavigateToSource [button]
- None [button]
### ContextMenuChartTrendline
- ChartDelete [button]
- ChartResetToMatchStyle [button]
- ChartChangeType [button]
- ChartSourceDataContextualMenu [button]
- ChartEditDataMenu [splitButton]
-   ChartShowDataGrid [button] < ChartEditDataMenu
-   ChartShowDataContextualMenu [button] < ChartEditDataMenu
- Chart3DView [button]
- ChartFormatTrendline [button]
- None [button]
### ContextMenuChartTrendlineLabel
- FirstCorrectionAlternative [button]
- SecondCorrectionAlternative [button]
- ThirdCorrectionAlternative [button]
- FourthCorrectionAlternative [button]
- FifthCorrectionAlternative [button]
- MoreCorrectionAlternatives [button]
- ChartDelete [button]
- ChartResetToMatchStyle [button]
- ChartEditText [button]
- ChartExitEditText [button]
- FontDialog [button]
- ChartChangeType [button]
- ChartSourceDataContextualMenu [button]
- ChartEditDataMenu [splitButton]
-   ChartShowDataGrid [button] < ChartEditDataMenu
-   ChartShowDataContextualMenu [button] < ChartEditDataMenu
- Chart3DView [button]
- ChartFormatTrendlineLabel [button]
- ChartFormatTrendlineAndSelect [button]
- ChartFormatDataSeriesAndSelect [button]
- None [button]
### ContextMenuChartUpBars
- ChartDelete [button]
- ChartResetToMatchStyle [button]
- ChartChangeType [button]
- ChartSourceDataContextualMenu [button]
- ChartEditDataMenu [splitButton]
-   ChartShowDataGrid [button] < ChartEditDataMenu
-   ChartShowDataContextualMenu [button] < ChartEditDataMenu
- Chart3DView [button]
- ChartFormatUpBars [button]
- None [button]
### ContextMenuChartWalls
- ChartDelete [button]
- ChartResetToMatchStyle [button]
- ChartChangeType [button]
- ChartSourceDataContextualMenu [button]
- ChartEditDataMenu [splitButton]
-   ChartShowDataGrid [button] < ChartEditDataMenu
-   ChartShowDataContextualMenu [button] < ChartEditDataMenu
- Chart3DView [button]
- ChartFormatWalls [button]
- None [button]
### ContextMenuGraphicsCompatibility
- Cut [button]
- Copy [button]
- PasteGalleryMini [gallery]
- ObjectsGroupMenu [menu]
-   ObjectsGroup [button] < ObjectsGroupMenu
-   ObjectsRegroup [button] < ObjectsGroupMenu
-   ObjectsUngroup [button] < ObjectsGroupMenu
- ObjectBringToFrontMenu [splitButton]
-   ObjectBringToFront [button] < ObjectBringToFrontMenu
-   ObjectBringForward [button] < ObjectBringToFrontMenu
- ObjectSendToBackMenu [splitButton]
-   ObjectSendToBack [button] < ObjectSendToBackMenu
-   ObjectSendBackward [button] < ObjectSendToBackMenu
- ObjectSaveAsPicture [button]
- MacroAssign [button]
### ContextMenuSmartArtContentPane
- FirstCorrectionAlternative [button]
- SecondCorrectionAlternative [button]
- ThirdCorrectionAlternative [button]
- FourthCorrectionAlternative [button]
- FifthCorrectionAlternative [button]
- MoreCorrectionAlternatives [button]
- Cut [button]
- Copy [button]
- PasteGalleryMini [gallery]
- SmartArtPromote [button]
- SmartArtDemote [button]
- SmartArtReorderUp [button]
- SmartArtReorderDown [button]
- FontDialog [button]
- ImeReconvert [button]
### ContextMenuSmartArtBackground
- Cut [button]
- Copy [button]
- PasteGalleryMini [gallery]
- SmartArtChangeLayout [button]
- SmartArtTextPane [toggleButton]
- ObjectsGroupMenu [menu]
-   ObjectsGroup [button] < ObjectsGroupMenu
-   ObjectsRegroup [button] < ObjectsGroupMenu
-   ObjectsUngroup [button] < ObjectsGroupMenu
- ObjectBringToFrontMenu [splitButton]
-   ObjectBringToFront [button] < ObjectBringToFrontMenu
-   ObjectBringForward [button] < ObjectBringToFrontMenu
- ObjectSendToBackMenu [splitButton]
-   ObjectSendToBack [button] < ObjectSendToBackMenu
-   ObjectSendBackward [button] < ObjectSendToBackMenu
- InsertLinkGallery [gallery]
-   None [button] < InsertLinkGallery
- HyperlinkOpenInMenu [menu]
-   HyperlinkOpenInWeb [button] < HyperlinkOpenInMenu
-   HyperlinkOpenInNative [button] < HyperlinkOpenInMenu
- HyperlinkOpenDefaultSettingMenu [menu]
-   OpenDefaultWeb [button] < HyperlinkOpenDefaultSettingMenu
-   OpenDefaultNative [button] < HyperlinkOpenDefaultSettingMenu
- HyperlinkInsert [button]
- HyperlinkEdit [button]
- HyperlinkOpenExcel [button]
- HyperlinkRemove [button]
- ObjectSaveAsPicture [button]
- MacroAssign [button]
- AutoAlternativeText [button]
- SmartArtResetGraphic [button]
- SmartArtConvertToShapes [button]
- ObjectSizeAndPropertiesDialog [button]
- ObjectFormatDialog [button]
- PictureAttributionNavigateToSource [button]
### ContextMenuSmartArtEditSmartArt
- Cut [button]
- Copy [button]
- PasteGalleryMini [gallery]
- ObjectAddText [button]
- ObjectEditText [button]
- FontDialog [button]
- ShapeChangeShapeGallery [gallery]
- SmartArtAddShapeMenu [menu]
-   SmartArtAddShapeAfter [button] < SmartArtAddShapeMenu
-   SmartArtAddShapeBefore [button] < SmartArtAddShapeMenu
-   SmartArtAddShapeAbove [button] < SmartArtAddShapeMenu
-   SmartArtAddShapeBelow [button] < SmartArtAddShapeMenu
-   SmartArtAddAssistant [button] < SmartArtAddShapeMenu
- GraphicChangeMenu [menu]
-   PictureChangeFromFile [button] < GraphicChangeMenu
-   PictureChangeFromMobile [button] < GraphicChangeMenu
-   PictureChangeFromM365 [button] < GraphicChangeMenu
-   PictureChangeFromOnline [button] < GraphicChangeMenu
-   ChangeIcon [button] < GraphicChangeMenu
-   PictureChangeFromClipboard [button] < GraphicChangeMenu
- InsertLinkGallery [gallery]
-   None [button] < InsertLinkGallery
- HyperlinkOpenInMenu [menu]
-   HyperlinkOpenInWeb [button] < HyperlinkOpenInMenu
-   HyperlinkOpenInNative [button] < HyperlinkOpenInMenu
- HyperlinkOpenDefaultSettingMenu [menu]
-   OpenDefaultWeb [button] < HyperlinkOpenDefaultSettingMenu
-   OpenDefaultNative [button] < HyperlinkOpenDefaultSettingMenu
- HyperlinkInsert [button]
- HyperlinkEdit [button]
- HyperlinkOpenExcel [button]
- HyperlinkRemove [button]
- ObjectSaveAsPicture [button]
- SmartArtResetShape [button]
- ObjectSizeAndPropertiesDialog [button]
- ObjectFormatDialog [button]
- PictureAttributionNavigateToSource [button]
### ContextMenuSmartArtEdit1DShape
- Cut [button]
- Copy [button]
- PasteGalleryMini [gallery]
- FontDialog [button]
- SmartArtAddShapeMenu [menu]
-   SmartArtAddShapeAfter [button] < SmartArtAddShapeMenu
-   SmartArtAddShapeBefore [button] < SmartArtAddShapeMenu
-   SmartArtAddShapeAbove [button] < SmartArtAddShapeMenu
-   SmartArtAddShapeBelow [button] < SmartArtAddShapeMenu
-   SmartArtAddAssistant [button] < SmartArtAddShapeMenu
- ObjectSaveAsPicture [button]
- SmartArtResetShape [button]
- ObjectSizeAndPropertiesDialog [button]
- ObjectFormatDialog [button]
### ContextMenuSmartArtEditText
- FirstCorrectionAlternative [button]
- SecondCorrectionAlternative [button]
- ThirdCorrectionAlternative [button]
- FourthCorrectionAlternative [button]
- FifthCorrectionAlternative [button]
- MoreCorrectionAlternatives [button]
- Cut [button]
- Copy [button]
- PasteGalleryMini [gallery]
- EquationOptionsGallery [gallery]
- TextEditModeExit [button]
- FontDialog [button]
- ParagraphDialog [button]
- ImeReconvert [button]
- BulletsGallery [gallery]
-   BulletsAndNumberingBulletsDialog [button] < BulletsGallery
- NumberingGallery [gallery]
-   BulletsAndNumberingNumberingDialog [button] < NumberingGallery
- ShapeChangeShapeGallery [gallery]
- SmartArtAddShapeMenu [menu]
-   SmartArtAddShapeAfter [button] < SmartArtAddShapeMenu
-   SmartArtAddShapeBefore [button] < SmartArtAddShapeMenu
-   SmartArtAddShapeAbove [button] < SmartArtAddShapeMenu
-   SmartArtAddShapeBelow [button] < SmartArtAddShapeMenu
-   SmartArtAddAssistant [button] < SmartArtAddShapeMenu
- InsertLinkGallery [gallery]
-   None [button] < InsertLinkGallery
- HyperlinkOpenInMenu [menu]
-   HyperlinkOpenInWeb [button] < HyperlinkOpenInMenu
-   HyperlinkOpenInNative [button] < HyperlinkOpenInMenu
- HyperlinkOpenDefaultSettingMenu [menu]
-   OpenDefaultWeb [button] < HyperlinkOpenDefaultSettingMenu
-   OpenDefaultNative [button] < HyperlinkOpenDefaultSettingMenu
- HyperlinkInsert [button]
- HyperlinkEdit [button]
- HyperlinkOpenExcel [button]
- HyperlinkRemove [button]
- SmartArtResetShape [button]
- WordArtFormatDialog [button]
- ObjectSizeAndPropertiesDialog [button]
- ObjectFormatDialog [button]
### ContextMenuSmartArtEdit1DShapeText
- FirstCorrectionAlternative [button]
- SecondCorrectionAlternative [button]
- ThirdCorrectionAlternative [button]
- FourthCorrectionAlternative [button]
- FifthCorrectionAlternative [button]
- MoreCorrectionAlternatives [button]
- Cut [button]
- Copy [button]
- PasteGalleryMini [gallery]
- FontDialog [button]
- ParagraphDialog [button]
- ImeReconvert [button]
- SmartArtAddShapeMenu [menu]
-   SmartArtAddShapeAfter [button] < SmartArtAddShapeMenu
-   SmartArtAddShapeBefore [button] < SmartArtAddShapeMenu
-   SmartArtAddShapeAbove [button] < SmartArtAddShapeMenu
-   SmartArtAddShapeBelow [button] < SmartArtAddShapeMenu
-   SmartArtAddAssistant [button] < SmartArtAddShapeMenu
- SmartArtResetShape [button]
- WordArtFormatDialog [button]
- ObjectSizeAndPropertiesDialog [button]
- ObjectFormatDialog [button]
### ContextMenuTimeSlicerLevel
- None [checkBox]
### ContextMenuTimeSlicer
- Cut [button]
- Copy [button]
- PasteGalleryMini [gallery]
- TimeSlicerRefresh [button]
- TimeSlicerClearFilter [button]
- SlicerConnectionsMenu [button]
- TimeSlicerDelete [button]
- ObjectsGroupMenu [menu]
-   ObjectsGroup [button] < ObjectsGroupMenu
-   ObjectsRegroup [button] < ObjectsGroupMenu
-   ObjectsUngroup [button] < ObjectsGroupMenu
- ObjectBringToFrontMenu [splitButton]
-   ObjectBringToFront [button] < ObjectBringToFrontMenu
-   ObjectBringForward [button] < ObjectBringToFrontMenu
- ObjectSendToBackMenu [splitButton]
-   ObjectSendToBack [button] < ObjectSendToBackMenu
-   ObjectSendBackward [button] < ObjectSendToBackMenu
- MacroAssign [button]
- AutoAlternativeText [button]
- ObjectSizeAndPropertiesDialog [button]
### ContextMenuSlicer
- Cut [button]
- Copy [button]
- PasteGalleryMini [gallery]
- SlicerRefresh [button]
- SlicerSortUp [toggleButton]
- SlicerSortDown [toggleButton]
- SlicerDataSourceOrder [toggleButton]
- SlicerClearFilter [button]
- SlicerMultiSelect [button]
- SlicerConnectionsMenu [button]
- SlicerDelete [button]
- ObjectsGroupMenu [menu]
-   ObjectsGroup [button] < ObjectsGroupMenu
-   ObjectsRegroup [button] < ObjectsGroupMenu
-   ObjectsUngroup [button] < ObjectsGroupMenu
- ObjectBringToFrontMenu [splitButton]
-   ObjectBringToFront [button] < ObjectBringToFrontMenu
-   ObjectBringForward [button] < ObjectBringToFrontMenu
- ObjectSendToBackMenu [splitButton]
-   ObjectSendToBack [button] < ObjectSendToBackMenu
-   ObjectSendBackward [button] < ObjectSendToBackMenu
- MacroAssign [button]
- AutoAlternativeText [button]
- ObjectSizeAndPropertiesDialog [button]
- SlicerSettings [button]
### ContextMenuOfficeWebExtension
- Cut [button]
- Copy [button]
- PasteGalleryMini [gallery]
- ObjectsGroupMenu [menu]
-   ObjectsGroup [button] < ObjectsGroupMenu
-   ObjectsRegroup [button] < ObjectsGroupMenu
-   ObjectsUngroup [button] < ObjectsGroupMenu
- ObjectBringToFrontMenu [splitButton]
-   ObjectBringToFront [button] < ObjectBringToFrontMenu
-   ObjectBringForward [button] < ObjectBringToFrontMenu
- ObjectSendToBackMenu [splitButton]
-   ObjectSendToBack [button] < ObjectSendToBackMenu
-   ObjectSendBackward [button] < ObjectSendToBackMenu
- ObjectSetShapeDefaults [button]
- ObjectSizeAndPropertiesDialog [button]
- ObjectFormatDialog [button]
