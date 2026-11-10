# powerpoint ribbon controls (Microsoft 365 Current Channel)

Outline of Microsoft's Office Fluent UI command identifiers (powerpoint), from github.com/OfficeDev/office-fluent-ui-command-identifiers (MIT License, Copyright (c) Microsoft Corporation). Rows: tab (tab set), group, controls in ribbon order; indented controls are inside the one after '<'. Dev reference for the Office-style interface; not shipped.

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
- RedoOrRepeat [button]
- SlideShowFromBeginning [button]
- PointerModeOptions [gallery]

##   (None (Core Tab))
### 
- TabHome [tab]
- TabInsert [tab]
- TabDrawInk [tab]
- TabDesign [tab]
- TabTransitions [tab]
- TabAnimations [tab]
- TabSlideShow [tab]
- TabRecording [tab]
- TabReview [tab]
- TabView [tab]
- TabDeveloper [tab]
- TabAddIns [tab]
- HelpTab [tab]
- TabPrintPreview [tab]
- TabMerge [tab]
- TabGrayscale [tab]
- TabBlackAndWhite [tab]
- TabBroadcastPresentation [tab]
- TabSlideMaster [tab]
- TabHandoutMaster [tab]
- TabNotesMaster [tab]
- TabBackgroundRemoval [tab]
- TabSlideMasterHome [tab]

## TabHome  (None (Core Tab))
### 
- GroupClipboard [group]
- GroupSlides [group]
- GroupFont [group]
- GroupParagraph [group]
- GroupDrawing [group]
- GroupEditing [group]
- GroupVoiceTools [group]
- GroupActivation [group]
- GroupClassifyLabelProtect [group]
- GroupDesignerOptions [group]
- GroupOfficeExtensionsAddinFlyout [group]
- GroupAIAssistance [group]
### GroupClipboard
- PasteMenu [splitButton]
-   Paste [button] < PasteMenu
-   PasteGallery [gallery] < PasteMenu
-     PasteSpecialDialog [button] < PasteMenu
- Cut [button]
- CopySplitButton [splitButton]
-   Copy [button] < CopySplitButton
-   PasteDuplicate [button] < CopySplitButton
- FormatPainter [control]
- ShowClipboard [button (dialogBoxLauncher)]
### GroupSlides
- SlideNewGallery [gallery]
-   DuplicateSelectedSlides [button] < SlideNewGallery
-   SlidesFromOutline [button] < SlideNewGallery
-   SlidesReuseSlides [toggleButton] < SlideNewGallery
-   ReuseSlides [toggleButton] < SlideNewGallery
-   PasteAsSlide [toggleButton] < SlideNewGallery
- ReuseSlides [toggleButton]
- SlideLayoutGallery [gallery]
- SlideReset [button]
- SectionMenu [menu]
-   SectionAdd [button] < SectionMenu
-   SectionRename [button] < SectionMenu
-   SectionMergeWithPrevious [button] < SectionMenu
-   SectionRemoveAll [button] < SectionMenu
-   SectionCollapseAll [button] < SectionMenu
-   SectionExpandAll [button] < SectionMenu
### GroupFont
- Font [comboBox]
- FontSize [comboBox]
- FontSizeIncrease [button]
- FontSizeDecrease [button]
- ClearFormatting [button]
- Bold [toggleButton]
- Italic [toggleButton]
- Underline [toggleButton]
- Shadow [toggleButton]
- Strikethrough [toggleButton]
- CharacterSpacingGallery [gallery]
- ChangeCaseGallery [gallery]
- TextHighlightColorPickerLicensed [gallery]
-   HighlightingStop [button] < TextHighlightColorPickerLicensed
- FontColorPicker [gallery]
-   FontColorMoreColorsDialogPowerPoint [button] < FontColorPicker
-   EyedropperFillText [button] < FontColorPicker
- FontDialogPowerPoint [button (dialogBoxLauncher)]
### GroupParagraph
- BulletsGallery [gallery]
-   BulletsAndNumberingBulletsDialog [button] < BulletsGallery
- NumberingGallery [gallery]
-   BulletsAndNumberingNumberingDialog [button] < NumberingGallery
- IndentDecrease [button]
- IndentIncrease [button]
- LineSpacingGalleryPowerPoint [gallery]
- AlignLeft [toggleButton]
- AlignCenter [toggleButton]
- AlignRight [toggleButton]
- AlignJustify [toggleButton]
- AlignJustifyMenu [menu]
-   AlignJustifyWithMixedLanguages [toggleButton] < AlignJustifyMenu
-   AlignJustifyLow [toggleButton] < AlignJustifyMenu
- ParagraphDistributed [toggleButton]
- AlignJustifyThai [toggleButton]
- TextDirectionLeftToRight [toggleButton]
- TextDirectionRightToLeft [toggleButton]
- TableColumnsGallery [gallery]
-   ParagraphMoreColumnsDialog [button] < TableColumnsGallery
- TextDirectionGallery [gallery]
-   TextDirectionMoreOptionsDialog [button] < TextDirectionGallery
- TextAlignGallery [gallery]
-   TextAlignMoreOptionsDialog [button] < TextAlignGallery
- ConvertToSmartArt [gallery]
-   ConvertToSmartArtMoreSmartArtGraphicsDialog [button] < ConvertToSmartArt
- PowerPointParagraphDialog [button (dialogBoxLauncher)]
### GroupDrawing
- GalleryAllShapesAndTextboxes [gallery]
- ObjectsArrangeMenu [menu]
-   ObjectBringToFront [button] < ObjectsArrangeMenu
-   ObjectSendToBack [button] < ObjectsArrangeMenu
-   ObjectBringForward [button] < ObjectsArrangeMenu
-   ObjectSendBackward [button] < ObjectsArrangeMenu
-   ObjectsGroup [button] < ObjectsArrangeMenu
-   ObjectsUngroup [button] < ObjectsArrangeMenu
-   ObjectsRegroup [button] < ObjectsArrangeMenu
-   ObjectAlignMenu [menu] < ObjectsArrangeMenu
-     ObjectsAlignLeftSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignCenterHorizontalSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignRightSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignTopSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignMiddleVerticalSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignBottomSmart [button] < ObjectsArrangeMenu
-     AlignDistributeHorizontally [button] < ObjectsArrangeMenu
-     AlignDistributeVertically [button] < ObjectsArrangeMenu
-     ObjectsAlignRelativeToContainerSmart [toggleButton] < ObjectsArrangeMenu
-     ObjectsAlignSelectedSmart [toggleButton] < ObjectsArrangeMenu
-   ObjectRotateGallery [gallery] < ObjectsArrangeMenu
-     ObjectRotationOptionsDialog [button] < ObjectsArrangeMenu
-   SelectionPane [toggleButton] < ObjectsArrangeMenu
- ShapeQuickStylesHome [gallery]
-   ShapeStylesOtherThemeFillsGallery [gallery] < ShapeQuickStylesHome
- ShapeFillColorPicker [gallery]
-   ObjectFillMoreColorsDialog [button] < ShapeFillColorPicker
-   EyedropperFill [button] < ShapeFillColorPicker
-   ObjectPictureFill [button] < ShapeFillColorPicker
-   GradientGallery [gallery] < ShapeFillColorPicker
-     ShapeFillMoreGradientsDialog [button] < ShapeFillColorPicker
-   ShapeFillTextureGallery [gallery] < ShapeFillColorPicker
-     MoreTextureOptions [button] < ShapeFillColorPicker
- ShapeOutlineColorPicker [gallery]
-   ObjectBorderOutlineColorMoreColorsDialog [button] < ShapeOutlineColorPicker
-   EyedropperOutline [button] < ShapeOutlineColorPicker
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
-     EyedropperGlow [button] < ShapeEffectsMenu
-     GlowsMoreOptions [button] < ShapeEffectsMenu
-   ObjectEffectSoftEdgesGallery [gallery] < ShapeEffectsMenu
-     SoftEdgesMoreOptions [button] < ShapeEffectsMenu
-   BevelShapeGallery [gallery] < ShapeEffectsMenu
-     _3DBevelOptionsDialog [button] < ShapeEffectsMenu
-   _3DRotationGallery [gallery] < ShapeEffectsMenu
-     _3DRotationOptionsDialog [button] < ShapeEffectsMenu
- ObjectFormatDialog [button (dialogBoxLauncher)]
### GroupEditing
- FindDialog [button]
- ReplaceMenu [splitButton]
-   ReplaceDialog [button] < ReplaceMenu
-   FontsReplaceFonts [button] < ReplaceMenu
- SelectMenu [menu]
-   SelectAll [button] < SelectMenu
-   ObjectsSelect [toggleButton] < SelectMenu
-   SelectionPane [toggleButton] < SelectMenu
### GroupVoiceTools
- DictationMenu [splitButton]
-   Dictate [toggleButton] < DictationMenu
-   DictateInMenu [button] < DictationMenu
-   SetDictationLanguageGallery [gallery] < DictationMenu
### GroupActivation
- ActivateOffice [button]
### GroupClassifyLabelProtect
- ClassifyLabelProtect [menu]
-   None [gallery] < ClassifyLabelProtect
-   None [toggleButton] < ClassifyLabelProtect
-   None [gallery] < ClassifyLabelProtect
-   None [gallery] < ClassifyLabelProtect
-   None [button] < ClassifyLabelProtect
-   None [button] < ClassifyLabelProtect
### GroupDesignerOptions
- DesignerPane [toggleButton]
### GroupOfficeExtensionsAddinFlyout
- OfficeExtensionsShowAddinFlyout [button]
- OfficeExtensionsShowAddinFlyout [toggleButton]
### GroupAIAssistance
- DesignerPane [toggleButton]
- CopilotDraftSlides [button]
- Copilot [button]

## TabInsert  (None (Core Tab))
### 
- GroupSlides2 [group]
- GroupInsertTables [group]
- GroupImages [group]
- GroupChunkCameoCamera [group]
- GroupInsertIllustrations [group]
- GroupContent [group]
- GroupPowerBI [group]
- GroupMSForms [group]
- GroupOfficeExtension [group]
- GroupInsertLinks [group]
- GroupInsertComments [group]
- GroupInsertText [group]
- GroupInsertSymbols [group]
- GroupInsertMediaClips [group]
### GroupSlides2
- SlideNewGalleryInsert [gallery]
-   DuplicateSelectedSlides [button] < SlideNewGalleryInsert
-   SlidesFromOutline [button] < SlideNewGalleryInsert
-   SlidesReuseSlides [toggleButton] < SlideNewGalleryInsert
-   ReuseSlides [toggleButton] < SlideNewGalleryInsert
-   PasteAsSlide [toggleButton] < SlideNewGalleryInsert
- ReuseSlides [toggleButton]
### GroupInsertTables
- TableInsertGallery [gallery]
-   TableInsert [button] < TableInsertGallery
-   TableDrawTable [toggleButton] < TableInsertGallery
-   ExcelSpreadsheetInsert [button] < TableInsertGallery
### GroupImages
- FlyoutAnchorInsertPictures [gallery]
-   PictureInsertFromFilePowerPoint [button] < FlyoutAnchorInsertPictures
-   InsertMobilePicture [button] < FlyoutAnchorInsertPictures
-   InsertM365Picture [button] < FlyoutAnchorInsertPictures
-   ClipArtInsertDialog [button] < FlyoutAnchorInsertPictures
- ScreenshotInsertGallery [gallery]
-   ScreenClipping [button] < ScreenshotInsertGallery
- PhotoAlbumInsertMenu [splitButton]
-   PhotoAlbumInsert [button] < PhotoAlbumInsertMenu
-   PhotoAlbumEdit [button] < PhotoAlbumInsertMenu
### GroupChunkCameoCamera
- DropdownInsertCameo [splitButton]
-   InsertCameo [button] < DropdownInsertCameo
-   InsertCameoAllSlides [button] < DropdownInsertCameo
### GroupInsertIllustrations
- ShapesInsertGallery [gallery]
- IconInsertFromFile [button]
- Insert3DModelFallback [button]
- Insert3DModelDropdown [splitButton]
-   Insert3DModelDefault [button] < Insert3DModelDropdown
-   Insert3DModelFromFile [button] < Insert3DModelDropdown
-   Insert3DModelFromFileConsumer [button] < Insert3DModelDropdown
-   Insert3DModelFromOnline [button] < Insert3DModelDropdown
-   Insert3DModelFromOnlineConsumer [button] < Insert3DModelDropdown
- SmartArtInsert [button]
- ChartInsert [button]
### GroupContent
- Collector [toggleButton]
### GroupPowerBI
- InsertPowerBI [button]
### GroupMSForms
- InsertMSForms [button]
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
### GroupInsertLinks
- InsertZoomMenu [menu]
-   MSPPTInsertTableofContents [button] < InsertZoomMenu
-   SectionZoomInsert [button] < InsertZoomMenu
-   SlideZoomInsert [button] < InsertZoomMenu
- InsertLinkGallery [gallery]
-   None [button] < InsertLinkGallery
- HyperlinkInsert [button]
- ActionInsert [button]
### GroupInsertComments
- InsertNewComment [button]
### GroupInsertText
- TextBoxInsert [toggleButton]
- TextBoxInsertMenu [splitButton]
-   TextBoxInsertHorizontal [toggleButton] < TextBoxInsertMenu
-   TextBoxInsertVertical [toggleButton] < TextBoxInsertMenu
- HeaderFooterInsert [button]
- WordArtInsertGallery [gallery]
- DateAndTimeInsert [button]
- NumberInsert [button]
- OleObjectctInsert [button]
### GroupInsertSymbols
- InsertBuildingBlocksEquationsGallery [gallery]
-   EquationInsertNew [button] < InsertBuildingBlocksEquationsGallery
-   InkEquation [button] < InsertBuildingBlocksEquationsGallery
- SymbolInsert [button]
### GroupInsertMediaClips
- MovieInsert02 [menu]
-   InsertVideoFromDevice [button] < MovieInsert02
-   MovieFromFileInsert [button] < MovieInsert02
-   InsertM365Video [button] < MovieInsert02
-   MovieFromClipOrganizerInsert [button] < MovieInsert02
- SoundInsertMenu02 [menu]
-   SoundInsertFromFile [button] < SoundInsertMenu02
-   SoundRecord [button] < SoundInsertMenu02
- ObjectScreenRecording [button]

## TabDrawInk  (None (Core Tab))
### 
- GroupInputMode [group]
- GroupWrite [group]
- GroupDrawingTools [group]
- GroupPensOneNote [group]
- GroupPens2 [group]
- GroupDrawWithTouch [group]
- GroupStencils [group]
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
- InkEraserMenu [splitButton]
-   CurrentEraser [toggleButton] < InkEraserMenu
-   InkEraser [toggleButton] < InkEraserMenu
-   PointEraserSmall [toggleButton] < InkEraserMenu
-   PointEraserMedium [toggleButton] < InkEraserMenu
-   SegmentEraser [toggleButton] < InkEraserMenu
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
-   EyedropperOutline [button] < InkColorPicker2
- LineThickness2 [gallery]
-   LineStylesDialog [button] < LineThickness2
### GroupDrawWithTouch
- FingerPaintingMode [toggleButton]
### GroupStencils
- ShowRulerStencil [toggleButton]
### GroupInkConvertOneNote
- DiagrammingOnline [toggleButton]
- InkToTextAnalysis [toggleButton]
- InkToShapeAnalysis [toggleButton]
- InkToTextAnalysis [button]
- InkToShapeAnalysis [button]
- SelectionToMathConvert [button]
- InkToMathSplitButton [splitButton]
-   InkToMathAnalysis [toggleButton] < InkToMathSplitButton
-   InkToMathAnalysis [button] < InkToMathSplitButton
-   MathInputEditor [button] < InkToMathSplitButton
### InkReplay
- Replay [toggleButton]
### GroupInkClose
- InkStopInkingReadingView [button]
### GroupPenAndInkHelp
- PenAndInkHelp [button]

## TabDesign  (None (Core Tab))
### 
- GroupSlideThemes [group]
- GroupThemeVariants [group]
- GroupCustomizeThemeOptions [group]
- GroupDesignerOptions [group]
### GroupSlideThemes
- SlideThemesGallery [gallery]
-   ThemeSearchOfficeOnlinePowerPoint [button] < SlideThemesGallery
-   ThemeBrowseForThemesPowerPoint [button] < SlideThemesGallery
-   ThemeSaveCurrentPowerPoint [button] < SlideThemesGallery
### GroupThemeVariants
- ThemeVariantsGallery [gallery]
-   ThemeColorsGallery [gallery] < ThemeVariantsGallery
-     ThemeColorsCreateNew [button] < ThemeVariantsGallery
-     ThemeColorsReset [button] < ThemeVariantsGallery
-   ThemeFontsGallery [gallery] < ThemeVariantsGallery
-     ThemeFontsCreateNew [button] < ThemeVariantsGallery
-   ThemeEffectsGallery [gallery] < ThemeVariantsGallery
-   SlideBackgroundStylesGallery [gallery] < ThemeVariantsGallery
-     SlideBackgroundFormatDialog [button] < ThemeVariantsGallery
-     SlideBackgroundReset [button] < ThemeVariantsGallery
### GroupCustomizeThemeOptions
- PowerPointPageSetup [gallery]
-   CustomSlideSize [button] < PowerPointPageSetup
- LaunchFormatBackground [button]
### GroupDesignerOptions
- DesignerPane [toggleButton]

## TabTransitions  (None (Core Tab))
### 
- GroupPreviewTransitions [group]
- GroupTransitionToThisSlide [group]
- GroupTransitionStyles [group]
### GroupPreviewTransitions
- TransitionPreview [button]
### GroupTransitionToThisSlide
- AnimationTransitionGallery [gallery]
- AnimationTransitionVariantGallery [gallery]
### GroupTransitionStyles
- AnimationTransitionSoundGallery [dropDown]
-   TransitionSoundLoopUntilNextSound [toggleButton] < AnimationTransitionSoundGallery
- TransitionDuration [control]
- SlideTransitionApplyToAll [button]
- AdvanceSlideLabel [labelControl]
- SlideTransitionOnMouseClick [checkBox]
- SlideTransitionAutomaticallyAfter [checkBox]
- TransitionTimeAutomaticallyAfter [control]

## TabAnimations  (None (Core Tab))
### 
- GroupPreview [group]
- GroupAnimations [group]
- GroupAnimationCustom [group]
- GroupAnimationTiming [group]
### GroupPreview
- AnimationPreviewMenu [splitButton]
-   AnimationPreview [button] < AnimationPreviewMenu
-   AnimationAutoPreview [toggleButton] < AnimationPreviewMenu
### GroupAnimations
- AnimationGallery [gallery]
-   AnimationCustomEntranceDialog [button] < AnimationGallery
-   AnimationCustomEmphasisDialog [button] < AnimationGallery
-   AnimationCustomExitDialog [button] < AnimationGallery
-   AnimationCustomPathDialog [button] < AnimationGallery
-   AnimationCustomActionVerbDialog [button] < AnimationGallery
- EffectOptionsMenu [menu]
-   EffectOptionsGallery [gallery] < EffectOptionsMenu
-   EffectOptionsColorsGallery [gallery] < EffectOptionsMenu
-   AnimationChangeGallery [gallery] < EffectOptionsMenu
-   BuildOrderChangeGallery [gallery] < EffectOptionsMenu
- EffectOptionsDialog [button (dialogBoxLauncher)]
### GroupAnimationCustom
- AnimationAddGallery [gallery]
-   AnimationCustomAddEntranceDialog [button] < AnimationAddGallery
-   AnimationCustomAddEmphasisDialog [button] < AnimationAddGallery
-   AnimationCustomAddExitDialog [button] < AnimationAddGallery
-   AnimationCustomAddPathDialog [button] < AnimationAddGallery
-   AnimationCustomAddActionVerbDialog [button] < AnimationAddGallery
- AnimationCustom [toggleButton]
- AnimationTriggerAddMenu [menu]
-   AnimationTriggerAddOnClick [gallery] < AnimationTriggerAddMenu
-   AnimationTriggerAddOnMediaBookmark [gallery] < AnimationTriggerAddMenu
- AnimationPainter [control]
### GroupAnimationTiming
- AnimationStartDropdown [dropDown]
- AnimationDuration [control]
- AnimationDelay [control]
- AnimationReorderLabel [labelControl]
- AnimationMoveEarlier [button]
- AnimationMoveLater [button]

## TabSlideShow  (None (Core Tab))
### 
- GroupSlideShowStart [group]
- GroupRehearse [group]
- GroupSlideShowSetup [group]
- GroupMonitors [group]
- GroupLiveSubtitles [group]
### GroupSlideShowStart
- SlideShowFromBeginning [button]
- SlideShowFromCurrent [button]
- PresentInTeamsMeeting [button]
- SlideshowBroadcastMenu [splitButton]
-   BroadcastSlideShow [button] < SlideshowBroadcastMenu
-   BroadcastSlideShowLync [button] < SlideshowBroadcastMenu
-   BroadcastSlideShowOfficeHosted [button] < SlideshowBroadcastMenu
-   BroadcastSlideShowGeneric [button] < SlideshowBroadcastMenu
- SlideShowCustomMenu [menu]
### GroupRehearse
- Rehearse [button]
### GroupSlideShowSetup
- SlideShowSetUpDialog [button]
- SlideHide [toggleButton]
- SlideShowRehearseTimings [button]
- RecordSlideShowMenu [splitButton]
-   RecordNarrationFromCurrentSlide [button] < RecordSlideShowMenu
-   RecordNarration [button] < RecordSlideShowMenu
-   SlideShowClearMenu [menu] < RecordSlideShowMenu
-     SlideShowClearCurrentSlideTiming [button] < RecordSlideShowMenu
-     SlideShowClearAllSlidesTimings [button] < RecordSlideShowMenu
-     SlideShowClearCurrentSlideNarration [button] < RecordSlideShowMenu
-     SlideShowClearAllSlidesNarrations [button] < RecordSlideShowMenu
- SlideShowPlayNarrations [checkBox]
- SlideShowUseRehearsedTimings [checkBox]
- MediaControlsShow [checkBox]
- HideSlide [button]
- KeepSlidesUpdated [checkBox]
### GroupMonitors
- SlideShowShowPresentationOnGallery [dropDown]
- SlideShowUsePresenterView [checkBox]
### GroupLiveSubtitles
- UseSubtitles [checkBox]
- SubtitleSettingsMenu [menu]
-   SubtitlesSpokenLanguage [gallery] < SubtitleSettingsMenu
-   SubtitlesDisplayLanguage [gallery] < SubtitleSettingsMenu
-   SubtitlesMicrophoneSettings [gallery] < SubtitleSettingsMenu
-   SubtitlePositionBottomOverlayed [toggleButton] < SubtitleSettingsMenu
-   SubtitlePositionTopOverlayed [toggleButton] < SubtitleSettingsMenu
-   SubtitlePositionBelowSlide [toggleButton] < SubtitleSettingsMenu
-   SubtitlePositionAboveSlide [toggleButton] < SubtitleSettingsMenu
-   SubtitleMoreSettings [button] < SubtitleSettingsMenu

## TabRecording  (None (Core Tab))
### 
- GroupPreviewTabRecord [group]
- GroupChunkCameoCamera [group]
- GroupRecordTabRecord [group]
- GroupEditTabRecord [group]
- GroupExportTabRecord [group]
- GroupRecord [group]
- GroupContentRecording [group]
- GroupAutoPlayMediaRecording [group]
- GroupSaveRecording [group]
- GroupHelpTabRecord [group]
### GroupPreviewTabRecord
- PreviewRecordedVideo [button]
### GroupChunkCameoCamera
- DropdownInsertCameo [splitButton]
-   InsertCameo [button] < DropdownInsertCameo
-   InsertCameoAllSlides [button] < DropdownInsertCameo
### GroupRecordTabRecord
- RecordNarration [button]
- RecordNarrationFromCurrentSlide [button]
- ObjectScreenRecordingAutoPlay [button]
- SoundRecordAutoPlay [button]
### GroupEditTabRecord
- ClearRecordingMenu [menu]
-   ClearRecordingOnCurrentSlide [button] < ClearRecordingMenu
-   ClearRecordingOnAllSlides [button] < ClearRecordingMenu
- ResetToCameoMenu [menu]
-   ResetToCameoOnCurrentSlide [button] < ResetToCameoMenu
-   ResetToCameoOnAllSlides [button] < ResetToCameoMenu
### GroupExportTabRecord
- FileSaveAsPowerPointSlideShow [button]
- ExportVideoRecording [button]
### GroupRecord
- RecordSlideShowMenu [splitButton]
-   RecordNarrationFromCurrentSlide [button] < RecordSlideShowMenu
-   RecordNarration [button] < RecordSlideShowMenu
-   SlideShowClearMenu [menu] < RecordSlideShowMenu
-     SlideShowClearCurrentSlideTiming [button] < RecordSlideShowMenu
-     SlideShowClearAllSlidesTimings [button] < RecordSlideShowMenu
-     SlideShowClearCurrentSlideNarration [button] < RecordSlideShowMenu
-     SlideShowClearAllSlidesNarrations [button] < RecordSlideShowMenu
### GroupContentRecording
- InsertMSForms [button]
- ScreenshotInsertGallery [gallery]
-   ScreenClipping [button] < ScreenshotInsertGallery
### GroupAutoPlayMediaRecording
- ObjectScreenRecordingAutoPlay [button]
- MovieFromFileRecordingAutoPlay [button]
- SoundInsertMenuAutoPlay [menu]
-   SoundInsertFromFileAutoPlay [button] < SoundInsertMenuAutoPlay
-   SoundRecordAutoPlay [button] < SoundInsertMenuAutoPlay
### GroupSaveRecording
- FileSaveAsPowerPointSlideShow [button]
- ExportToVideoRecording [button]
- UploadToMSStreamRecording [button]
### GroupHelpTabRecord
- LearnAboutRecord [button]

## TabReview  (None (Core Tab))
### 
- GroupProofing [group]
- GroupAccessibility [group]
- GroupInsights [group]
- GroupLanguage [group]
- GroupChineseTranslation [group]
- GroupActivity [group]
- GroupComments [group]
- GroupReviewCompare [group]
- GroupInk [group]
### GroupProofing
- Spelling [button]
- Thesaurus [button]
### GroupAccessibility
- MenuAccessibility [splitButton]
-   AccessibilityChecker [button] < MenuAccessibility
-   AltTextPaneRibbon [toggleButton] < MenuAccessibility
-   ReadingOrderPane [toggleButton] < MenuAccessibility
-   OpenOptionsEaseOfAcess [button] < MenuAccessibility
-   AccessibilityHelp [button] < MenuAccessibility
### GroupInsights
- Insights [button]
### GroupLanguage
- TranslateMenu [menu]
-   TranslateSelected [button] < TranslateMenu
-   MiniTranslator [toggleButton] < TranslateMenu
-   TranslationLanguageOptions [button] < TranslateMenu
- Translator [button]
- HangulHanjaConversionPowerPoint [button]
- LanguageCommands [menu]
-   SetLanguage [button] < LanguageCommands
-   LanguagePreferences [button] < LanguageCommands
### GroupChineseTranslation
- TranslateToSimplifiedChinese [button]
- TranslateToTraditionalChinese [button]
- ChineseTranslationDialog [button]
### GroupActivity
- MarkAllRevTrackHighlightsAsRead [button]
- ToggleRevTrackCanvasHighlights [toggleButton]
### GroupComments
- ReviewNewComment [button]
- ReviewDeleteCommentsMenuPowerPoint [splitButton]
-   ReviewDeleteCommentPowerPoint [button] < ReviewDeleteCommentsMenuPowerPoint
-   ReviewDeleteAllMarkupOnSlide [button] < ReviewDeleteCommentsMenuPowerPoint
-   ReviewDeleteAllMarkupInPresentation [button] < ReviewDeleteCommentsMenuPowerPoint
- ReviewPreviousCommentPowerPoint [button]
- ReviewNextCommentPowerPoint [button]
- ShowCommentsMenu [splitButton]
-   CommentsPaneSplitButton [toggleButton] < ShowCommentsMenu
-   CommentsPane [toggleButton] < ShowCommentsMenu
-   ReviewShowOrHideMarkup [toggleButton] < ShowCommentsMenu
-   ShowAllCommentMarkers [toggleButton] < ShowCommentsMenu
### GroupReviewCompare
- ReviewCompareAndMerge [button]
- ReviewRevisionAcceptMenu [splitButton]
-   ReviewRevisionAccept [button] < ReviewRevisionAcceptMenu
-   ReviewSlideRevisionAccept [button] < ReviewRevisionAcceptMenu
-   ReviewPresentationRevisionAccept [button] < ReviewRevisionAcceptMenu
- ReviewRevisionRejectMenu [splitButton]
-   ReviewRevisionReject [button] < ReviewRevisionRejectMenu
-   ReviewSlideRevisionReject [button] < ReviewRevisionRejectMenu
-   ReviewPresentationRevisionReject [button] < ReviewRevisionRejectMenu
- ReviewPreviousChange [button]
- ReviewNextChange [button]
- ReviewReviewingPanePowerPoint [toggleButton]
- ReviewEndReviewPowerPoint [button]
### GroupInk
- InkingStart [button]
- HideInkSplitButton [splitButton]
-   HideInk [toggleButton] < HideInkSplitButton
-   InkDeleteAllOnSlide [button] < HideInkSplitButton
-   InkDeleteAll [button] < HideInkSplitButton

## TabView  (None (Core Tab))
### 
- GroupPresentationViews [group]
- GroupMasterViews [group]
- GroupViewShowHide [group]
- GroupViewDirection [group]
- GroupZoom [group]
- GroupColorGrayscale [group]
- GroupWindow [group]
- GroupMacros [group]
### GroupPresentationViews
- ViewThumbnailViewPowerPoint [toggleButton]
- ViewOutlineViewPowerPoint [toggleButton]
- ViewSlideSorterView [toggleButton]
- ViewNotesPageView [toggleButton]
- ViewSlideShowReadingView [button]
### GroupMasterViews
- ViewSlideMasterView [toggleButton]
- ViewHandoutMasterView [toggleButton]
- ViewNotesMasterView [toggleButton]
### GroupViewShowHide
- ViewRulerPowerPoint [checkBox]
- ViewGridlinesPowerPoint [checkBox]
- GuidesShowHide [checkBox]
- ShowNotes [toggleButton]
- GridSettings [button (dialogBoxLauncher)]
### GroupViewDirection
- ViewDirectionMenu [menu]
-   ViewDirectionLeftToRight [toggleButton] < ViewDirectionMenu
-   ViewDirectionRightToLeft [toggleButton] < ViewDirectionMenu
### GroupZoom
- ZoomDialog [button]
- ZoomFitToWindow [button]
### GroupColorGrayscale
- ViewDisplayInHighContrast [toggleButton]
- ViewDisplayInColor [toggleButton]
- ViewDisplayInGrayscale [toggleButton]
- ViewDisplayInPureBlackAndWhite [toggleButton]
### GroupWindow
- WindowNew [button]
- WindowsArrangeAll [button]
- WindowsCascade [button]
- WindowMoveSplit [button]
- WindowSwitchWindowsMenuPowerPoint [menu]
-   WindowMoreWindowsDialog [toggleButton] < WindowSwitchWindowsMenuPowerPoint
### GroupMacros
- MacroPlay [button]

## TabDeveloper  (None (Core Tab))
### 
- GroupCode [group]
- GroupAddins [group]
- GroupControls [group]
### GroupCode
- VisualBasic [button]
- MacroPlay [button]
- MacroSecurity [button]
### GroupAddins
- OfficeExtensionsGallery3 [button]
- AddInManager [button]
- ComAddInsDialog [button]
### GroupControls
- ActiveXLabel [button]
- ActiveXTextBox [button]
- ActiveXSpinButton [button]
- ActiveXButton [button]
- ActiveXImage [button]
- ActiveXScrollBar [button]
- ActiveXCheckBox [button]
- ActiveXRadioButton [button]
- ActiveXComboBox [button]
- ActiveXListBox [button]
- ActiveXToggleButton [button]
- MoreControlsDialog [button]
- ControlProperties [button]
- ViewVisualBasicCode [button]

## TabAddIns  (None (Core Tab))
### 
- GroupAddInsMenuCommands [group]
- GroupAddInsToolbarCommands [group]
- GroupAddInsCustomToolbars [group]

## HelpTab  (None (Core Tab))
### 
- GroupHelpAndSupport [group]
### GroupHelpAndSupport
- Help [button]
- ContactUs [button]
- OfficeFeedbackHelpTab [button]
- OfficeFeedbackIcon [button]
- ShowTraining [button]
- WhatsNewRecentUpdates [button]

## TabPrintPreview  (None (Core Tab))
### 
- GroupPrintPreviewPrint [group]
- GroupPrintPreviewPageSetup [group]
- GroupZoom [group]
- GroupPrintPreviewPreview [group]
### GroupPrintPreviewPrint
- PrintPreviewAndPrint [button]
- PrintOptionsMenu [menu]
### GroupPrintPreviewPageSetup
- PrintWhatLabel [labelControl]
- PrintWhat [dropDown]
- PageOrientationGallery [menu]
-   PageOrientationPortrait [toggleButton] < PageOrientationGallery
-   PageOrientationLandscape [toggleButton] < PageOrientationGallery
### GroupZoom
- ZoomDialog [button]
- ZoomFitToWindow [button]
### GroupPrintPreviewPreview
- PageNext [button]
- PagePrevious [button]
- PrintPreviewClose [button]

## TabMerge  (None (Core Tab))
### 
- GroupResolutions [group]
- GroupShowHideChanges [group]
- GroupAnimationPreview [group]
- GroupMergeViewClose [group]
### GroupResolutions
- ReviewRevisionAcceptMenu [splitButton]
-   ReviewRevisionAccept [button] < ReviewRevisionAcceptMenu
-   ReviewSlideRevisionAccept [button] < ReviewRevisionAcceptMenu
-   ReviewPresentationRevisionAccept [button] < ReviewRevisionAcceptMenu
- ReviewRevisionRejectMenu [splitButton]
-   ReviewRevisionReject [button] < ReviewRevisionRejectMenu
-   ReviewSlideRevisionReject [button] < ReviewRevisionRejectMenu
-   ReviewPresentationRevisionReject [button] < ReviewRevisionRejectMenu
- ReviewPreviousChange [button]
- ReviewNextChange [button]
### GroupShowHideChanges
- ReviewShowChanges [menu]
-   ReviewShowOnlyConflicts [toggleButton] < ReviewShowChanges
-   ReviewShowAllChanges [toggleButton] < ReviewShowChanges
- ReviewReviewingPanePowerPoint [toggleButton]
### GroupAnimationPreview
- AnimationPreview [button]
### GroupMergeViewClose
- MergeViewClose [button]

## TabGrayscale  (None (Core Tab))
### 
- GroupColorModeSetting [group]
- GroupColorModeClose [group]
### GroupColorModeSetting
- BlackAndWhiteAutomatic [toggleButton]
- BlackAndWhiteGrayscale [toggleButton]
- BlackAndWhiteLightGrayscale [toggleButton]
- BlackAndWhiteInverseGrayscale [toggleButton]
- BlackAndWhiteGrayWithWhiteFill [toggleButton]
- BlackAndWhiteBlackWithGrayscaleFill [toggleButton]
- BlackAndWhiteBlackWithWhiteFill [toggleButton]
- BlackAndWhiteBlack [toggleButton]
- BlackAndWhiteWhite [toggleButton]
- BlackAndWhiteDontShow [toggleButton]
### GroupColorModeClose
- ViewBackToColorView [toggleButton]

## TabBlackAndWhite  (None (Core Tab))
### 
- GroupColorModeSetting [group]
- GroupColorModeClose [group]
### GroupColorModeSetting
- BlackAndWhiteAutomatic [toggleButton]
- BlackAndWhiteGrayscale [toggleButton]
- BlackAndWhiteLightGrayscale [toggleButton]
- BlackAndWhiteInverseGrayscale [toggleButton]
- BlackAndWhiteGrayWithWhiteFill [toggleButton]
- BlackAndWhiteBlackWithGrayscaleFill [toggleButton]
- BlackAndWhiteBlackWithWhiteFill [toggleButton]
- BlackAndWhiteBlack [toggleButton]
- BlackAndWhiteWhite [toggleButton]
- BlackAndWhiteDontShow [toggleButton]
### GroupColorModeClose
- ViewBackToColorView [toggleButton]

## TabBroadcastPresentation  (None (Core Tab))
### 
- GroupBroadcastResume [group]
- GroupMonitors [group]
- GroupBroadcastBroadcast [group]
### GroupBroadcastResume
- BroadcastFromBeginning [button]
- BroadcastFromCurrent [button]
### GroupMonitors
- SlideShowShowPresentationOnGallery [dropDown]
- SlideShowUsePresenterView [checkBox]
### GroupBroadcastBroadcast
- BroadcastNotesSplitButton [splitButton]
-   BroadcastNotes [button] < BroadcastNotesSplitButton
-   OpenMeetingNotes [button] < BroadcastNotesSplitButton
-   ShareNotesWithMeeting [button] < BroadcastNotesSplitButton
- BroadcastSendInvitation [button]
- BroadcastEnd [button]

## TabSlideMaster  (None (Core Tab))
### 
- GroupMasterEdit [group]
- GroupMasterLayout [group]
- GroupMasterEditTheme [group]
- GroupBackground [group]
- GroupSlideSize [group]
- GroupMasterClose [group]
### GroupMasterEdit
- SlideNew [button]
- SlideMasterInsertLayout [button]
- SlideDelete [button]
- SlideMasterRenameMaster [button]
- SlideMasterPreserveMaster [toggleButton]
### GroupMasterLayout
- SlideMasterMasterLayout [button]
- SlideMasterInsertPlaceholderMenu [splitButton]
-   SlideMasterContentPlaceholderInsert [button] < SlideMasterInsertPlaceholderMenu
-   SlideMasterVerticalContentPlaceholderInsert [button] < SlideMasterInsertPlaceholderMenu
-   SlideMasterTextPlaceholderInsert [button] < SlideMasterInsertPlaceholderMenu
-   SlideMasterVerticalTextPlaceholderInsert [button] < SlideMasterInsertPlaceholderMenu
-   SlideMasterPicturePlaceholderInsert [button] < SlideMasterInsertPlaceholderMenu
-   SlideMasterChartPlaceholderInsert [button] < SlideMasterInsertPlaceholderMenu
-   SlideMasterTablePlaceholderInsert [button] < SlideMasterInsertPlaceholderMenu
-   SlideMasterDiagramPlaceholderInsert [button] < SlideMasterInsertPlaceholderMenu
-   SlideMasterMediaPlaceholderInsert [button] < SlideMasterInsertPlaceholderMenu
-   SlideMasterClipArtPlaceholderInsert [button] < SlideMasterInsertPlaceholderMenu
-   AddCameoPlaceholder [button] < SlideMasterInsertPlaceholderMenu
- SlideMasterShowTitle [checkBox]
- SlideMasterShowFooters [checkBox]
### GroupMasterEditTheme
- SlideThemesGallery [gallery]
-   ThemeSearchOfficeOnlinePowerPoint [button] < SlideThemesGallery
-   ThemeBrowseForThemesPowerPoint [button] < SlideThemesGallery
-   ThemeSaveCurrentPowerPoint [button] < SlideThemesGallery
### GroupBackground
- ThemeColorsGallery [gallery]
-   ThemeColorsCreateNew [button] < ThemeColorsGallery
-   ThemeColorsReset [button] < ThemeColorsGallery
- ThemeFontsGallery [gallery]
-   ThemeFontsCreateNew [button] < ThemeFontsGallery
- ThemeEffectsGallery [gallery]
- SlideBackgroundStylesGallery [gallery]
-   SlideBackgroundFormatDialog [button] < SlideBackgroundStylesGallery
-   SlideBackgroundReset [button] < SlideBackgroundStylesGallery
- SlideBackgroundHideGraphics [checkBox]
- SlideBackgroundFormatDialog [button (dialogBoxLauncher)]
### GroupSlideSize
- PowerPointPageSetup [gallery]
-   CustomSlideSize [button] < PowerPointPageSetup
### GroupMasterClose
- MasterViewClose [button]

## TabHandoutMaster  (None (Core Tab))
### 
- GroupPageSetupHandoutMaster [group]
- GroupPlaceholdersHandoutMaster [group]
- GroupMasterEditTheme [group]
- GroupBackground [group]
- GroupMasterClose [group]
### GroupPageSetupHandoutMaster
- HandoutOrientation [gallery]
- PowerPointPageSetup [gallery]
-   CustomSlideSize [button] < PowerPointPageSetup
- SlidesPerPageGallery [menu]
-   SlidesPerPage1Slide [toggleButton] < SlidesPerPageGallery
-   SlidesPerPage2Slides [toggleButton] < SlidesPerPageGallery
-   SlidesPerPage3Slides [toggleButton] < SlidesPerPageGallery
-   SlidesPerPage4Slides [toggleButton] < SlidesPerPageGallery
-   SlidesPerPage6Slides [toggleButton] < SlidesPerPageGallery
-   SlidesPerPage9Slides [toggleButton] < SlidesPerPageGallery
-   SlidesPerPageSlideOutline [toggleButton] < SlidesPerPageGallery
### GroupPlaceholdersHandoutMaster
- MasterShowHeader [checkBox]
- MasterShowFooter [checkBox]
- MasterShowDate [checkBox]
- MasterShowPageNumber [checkBox]
### GroupMasterEditTheme
- SlideThemesGallery [gallery]
-   ThemeSearchOfficeOnlinePowerPoint [button] < SlideThemesGallery
-   ThemeBrowseForThemesPowerPoint [button] < SlideThemesGallery
-   ThemeSaveCurrentPowerPoint [button] < SlideThemesGallery
### GroupBackground
- ThemeColorsGallery [gallery]
-   ThemeColorsCreateNew [button] < ThemeColorsGallery
-   ThemeColorsReset [button] < ThemeColorsGallery
- ThemeFontsGallery [gallery]
-   ThemeFontsCreateNew [button] < ThemeFontsGallery
- ThemeEffectsGallery [gallery]
- SlideBackgroundStylesGallery [gallery]
-   SlideBackgroundFormatDialog [button] < SlideBackgroundStylesGallery
-   SlideBackgroundReset [button] < SlideBackgroundStylesGallery
- SlideBackgroundHideGraphics [checkBox]
- SlideBackgroundFormatDialog [button (dialogBoxLauncher)]
### GroupMasterClose
- MasterViewClose [button]

## TabNotesMaster  (None (Core Tab))
### 
- GroupPageSetupNotesMaster [group]
- GroupPlaceholdersNotesMaster [group]
- GroupMasterEditTheme [group]
- GroupBackground [group]
- GroupMasterClose [group]
### GroupPageSetupNotesMaster
- MasterNotesPageOrientation [gallery]
- PowerPointPageSetup [gallery]
-   CustomSlideSize [button] < PowerPointPageSetup
### GroupPlaceholdersNotesMaster
- MasterShowHeader [checkBox]
- MasterShowSlideImage [checkBox]
- MasterShowFooter [checkBox]
- MasterShowDate [checkBox]
- MasterShowBody [checkBox]
- MasterShowPageNumber [checkBox]
### GroupMasterEditTheme
- SlideThemesGallery [gallery]
-   ThemeSearchOfficeOnlinePowerPoint [button] < SlideThemesGallery
-   ThemeBrowseForThemesPowerPoint [button] < SlideThemesGallery
-   ThemeSaveCurrentPowerPoint [button] < SlideThemesGallery
### GroupBackground
- ThemeColorsGallery [gallery]
-   ThemeColorsCreateNew [button] < ThemeColorsGallery
-   ThemeColorsReset [button] < ThemeColorsGallery
- ThemeFontsGallery [gallery]
-   ThemeFontsCreateNew [button] < ThemeFontsGallery
- ThemeEffectsGallery [gallery]
- SlideBackgroundStylesGallery [gallery]
-   SlideBackgroundFormatDialog [button] < SlideBackgroundStylesGallery
-   SlideBackgroundReset [button] < SlideBackgroundStylesGallery
- SlideBackgroundHideGraphics [checkBox]
- SlideBackgroundFormatDialog [button (dialogBoxLauncher)]
### GroupMasterClose
- MasterViewClose [button]

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

## TabSlideMasterHome  (None (Core Tab))
### 
- GroupClipboard [group]
- GroupMasterSlides [group]
- GroupFont [group]
- GroupParagraph [group]
- GroupDrawing [group]
- GroupEditing [group]
### GroupClipboard
- PasteMenu [splitButton]
-   Paste [button] < PasteMenu
-   PasteGallery [gallery] < PasteMenu
-     PasteSpecialDialog [button] < PasteMenu
- Cut [button]
- CopySplitButton [splitButton]
-   Copy [button] < CopySplitButton
-   PasteDuplicate [button] < CopySplitButton
- FormatPainter [control]
- ShowClipboard [button (dialogBoxLauncher)]
### GroupMasterSlides
- SlideNewGallery [gallery]
-   DuplicateSelectedSlides [button] < SlideNewGallery
-   SlidesFromOutline [button] < SlideNewGallery
-   SlidesReuseSlides [toggleButton] < SlideNewGallery
-   ReuseSlides [toggleButton] < SlideNewGallery
-   PasteAsSlide [toggleButton] < SlideNewGallery
- SlideLayoutGallery [gallery]
- SlideReset [button]
- SlideDelete [button]
### GroupFont
- Font [comboBox]
- FontSize [comboBox]
- FontSizeIncrease [button]
- FontSizeDecrease [button]
- ClearFormatting [button]
- Bold [toggleButton]
- Italic [toggleButton]
- Underline [toggleButton]
- Shadow [toggleButton]
- Strikethrough [toggleButton]
- CharacterSpacingGallery [gallery]
- ChangeCaseGallery [gallery]
- TextHighlightColorPickerLicensed [gallery]
-   HighlightingStop [button] < TextHighlightColorPickerLicensed
- FontColorPicker [gallery]
-   FontColorMoreColorsDialogPowerPoint [button] < FontColorPicker
-   EyedropperFillText [button] < FontColorPicker
- FontDialogPowerPoint [button (dialogBoxLauncher)]
### GroupParagraph
- BulletsGallery [gallery]
-   BulletsAndNumberingBulletsDialog [button] < BulletsGallery
- NumberingGallery [gallery]
-   BulletsAndNumberingNumberingDialog [button] < NumberingGallery
- IndentDecrease [button]
- IndentIncrease [button]
- LineSpacingGalleryPowerPoint [gallery]
- AlignLeft [toggleButton]
- AlignCenter [toggleButton]
- AlignRight [toggleButton]
- AlignJustify [toggleButton]
- AlignJustifyMenu [menu]
-   AlignJustifyWithMixedLanguages [toggleButton] < AlignJustifyMenu
-   AlignJustifyLow [toggleButton] < AlignJustifyMenu
- ParagraphDistributed [toggleButton]
- AlignJustifyThai [toggleButton]
- TextDirectionLeftToRight [toggleButton]
- TextDirectionRightToLeft [toggleButton]
- TableColumnsGallery [gallery]
-   ParagraphMoreColumnsDialog [button] < TableColumnsGallery
- TextDirectionGallery [gallery]
-   TextDirectionMoreOptionsDialog [button] < TextDirectionGallery
- TextAlignGallery [gallery]
-   TextAlignMoreOptionsDialog [button] < TextAlignGallery
- ConvertToSmartArt [gallery]
-   ConvertToSmartArtMoreSmartArtGraphicsDialog [button] < ConvertToSmartArt
- PowerPointParagraphDialog [button (dialogBoxLauncher)]
### GroupDrawing
- GalleryAllShapesAndTextboxes [gallery]
- ObjectsArrangeMenu [menu]
-   ObjectBringToFront [button] < ObjectsArrangeMenu
-   ObjectSendToBack [button] < ObjectsArrangeMenu
-   ObjectBringForward [button] < ObjectsArrangeMenu
-   ObjectSendBackward [button] < ObjectsArrangeMenu
-   ObjectsGroup [button] < ObjectsArrangeMenu
-   ObjectsUngroup [button] < ObjectsArrangeMenu
-   ObjectsRegroup [button] < ObjectsArrangeMenu
-   ObjectAlignMenu [menu] < ObjectsArrangeMenu
-     ObjectsAlignLeftSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignCenterHorizontalSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignRightSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignTopSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignMiddleVerticalSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignBottomSmart [button] < ObjectsArrangeMenu
-     AlignDistributeHorizontally [button] < ObjectsArrangeMenu
-     AlignDistributeVertically [button] < ObjectsArrangeMenu
-     ObjectsAlignRelativeToContainerSmart [toggleButton] < ObjectsArrangeMenu
-     ObjectsAlignSelectedSmart [toggleButton] < ObjectsArrangeMenu
-   ObjectRotateGallery [gallery] < ObjectsArrangeMenu
-     ObjectRotationOptionsDialog [button] < ObjectsArrangeMenu
-   SelectionPane [toggleButton] < ObjectsArrangeMenu
- ShapeQuickStylesHome [gallery]
-   ShapeStylesOtherThemeFillsGallery [gallery] < ShapeQuickStylesHome
- ShapeFillColorPicker [gallery]
-   ObjectFillMoreColorsDialog [button] < ShapeFillColorPicker
-   EyedropperFill [button] < ShapeFillColorPicker
-   ObjectPictureFill [button] < ShapeFillColorPicker
-   GradientGallery [gallery] < ShapeFillColorPicker
-     ShapeFillMoreGradientsDialog [button] < ShapeFillColorPicker
-   ShapeFillTextureGallery [gallery] < ShapeFillColorPicker
-     MoreTextureOptions [button] < ShapeFillColorPicker
- ShapeOutlineColorPicker [gallery]
-   ObjectBorderOutlineColorMoreColorsDialog [button] < ShapeOutlineColorPicker
-   EyedropperOutline [button] < ShapeOutlineColorPicker
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
-     EyedropperGlow [button] < ShapeEffectsMenu
-     GlowsMoreOptions [button] < ShapeEffectsMenu
-   ObjectEffectSoftEdgesGallery [gallery] < ShapeEffectsMenu
-     SoftEdgesMoreOptions [button] < ShapeEffectsMenu
-   BevelShapeGallery [gallery] < ShapeEffectsMenu
-     _3DBevelOptionsDialog [button] < ShapeEffectsMenu
-   _3DRotationGallery [gallery] < ShapeEffectsMenu
-     _3DRotationOptionsDialog [button] < ShapeEffectsMenu
- ObjectFormatDialog [button (dialogBoxLauncher)]
### GroupEditing
- FindDialog [button]
- ReplaceMenu [splitButton]
-   ReplaceDialog [button] < ReplaceMenu
-   FontsReplaceFonts [button] < ReplaceMenu
- SelectMenu [menu]
-   SelectAll [button] < SelectMenu
-   ObjectsSelect [toggleButton] < SelectMenu
-   SelectionPane [toggleButton] < SelectMenu

## None (Not in the Ribbon)  (None (Not in the Ribbon))
### 
- FileNewDefault [button]
- FileOpenUsingBackstage [button]
- FileSave [button]
- SaveAndReview [button]
- FileSendAsAttachment [button]
- FilePrintQuick [button]
- PrintPreviewAndPrint [button]
- Spelling [button]
- Undo [gallery]
- RedoOrRepeat [button]
- SlideShowFromBeginning [button]
- FileOpenRecentFile [button]
- PointerModeOptions [gallery]
- FilePermissionRestrictAs [button]
- MdiChildSystemMenu [menu]
- MinimizeRibbon [toggleButton]
- WindowMinimize [button]
- WindowRestore [button]
- WindowClose [button]
- HideRibbon [toggleButton]
- PasteAsHyperlink [button]
- PasteDuplicate [button]
- PasteBitmap [button]
- PasteGif [button]
- PastePng [button]
- PasteJpeg [button]
- PasteInk [button]
- PasteTextOnly [button]
- PasteAsPicture [button]
- PasteSourceFormatting [button]
- PasteDestinationTheme [button]
- PasteExcelChartSourceFormatting [button]
- PasteExcelChartDestinationTheme [button]
- PasteLinkedExcelChartSourceFormatting [button]
- PasteLinkedExcelChartDestinationTheme [button]
- PasteAsEmbedded [button]
- PasteExcelTableSourceFormatting [button]
- PasteExcelTableDestinationTableStyle [button]
- FileSaveACopy [button]
- FileDiscardCachedChanges [button]
- ReopenFromServer [button]
- None [button]
- RefreshServerFile [button]
- ReopenServerFile [button]
- AnimationTransitionSpeedGallery [dropDown]
- ViewMessageBar [checkBox]
- ExportToVideo [button]
- ExportToAnimatedGif [button]
- ZoomClassic [gallery]
- ShapeRectangle [toggleButton]
- ContrastMore [button]
- SlideShowPreviousSlide [button]
- Redo [gallery]
- BrightnessMore [button]
- BordersShadingDialog [button]
- TableInsertDialog [button]
- ShapeOval [toggleButton]
- SlideShowNextSlide [button]
- BrightnessLess [button]
- ContrastLess [button]
- ShapeElbowConnectorArrow [toggleButton]
- ObjectRotateLeft90 [button]
- ShapeStraightConnectorArrow [toggleButton]
- ShadowNudgeLeftClassic [button]
- ShadowNudgeDownClassic [button]
- ShapeElbowConnector [toggleButton]
- ShadowNudgeRightClassic [button]
- ShadowNudgeUpClassic [button]
- OpenOptionsAddIns [button]
- OpenOptionsAdvanced [button]
- OpenOptionsPPTTrustCenter [button]
- OpenOptionsCustomization [button]
- OpenOptionsEaseOfAcess [button]
- OpenOptionsCustomizeRibbon [button]
- OpenOptionsExperimentSettings [button]
- OpenOptionsLanguage [button]
- OpenOptionsSaving [button]
- OpenOptionsSetup [button]
- OpenOptionsSpelling [button]
- OpenOptionsTrustCenterSettings [button]
- FindTellMe [control]
- ObjectSetShapeDefaults [button]
- ShapeRoundedRectangle [toggleButton]
- ShapeStraightConnector [toggleButton]
- ObjectFlipHorizontal [button]
- ShapeRightArrow [toggleButton]
- LinksUpdate [button]
- OutlineDemote [button]
- ObjectRotateFree [button]
- Subscript [toggleButton]
- SlideShowCustom [button]
- ShapeDownArrow [toggleButton]
- ObjectFlipVertical [button]
- Superscript [toggleButton]
- ObjectRotateRight90 [button]
- OutlinePromote [button]
- OutlineShowTextFormatting [toggleButton]
- ShadowOnOrOffClassic [button]
- ShapeRoundedRectangularCallout [toggleButton]
- Clear [button]
- PasteApplyStyle [button]
- ShapeIsoscelesTriangle [toggleButton]
- WordArtVerticalText [button]
- SlideShowResumeShow [button]
- ShapesMoreShapes [button]
- AutoCorrect [button]
- ViewFullScreenView [button]
- OutlineMoveUp [button]
- ShapeLeftBrace [toggleButton]
- ObjectNudgeDown [button]
- ShapeRightBrace [toggleButton]
- PickUpStyle [button]
- InsertTab [button]
- About [button]
- PropertySheet [button]
- ObjectNudgeLeft [button]
- CreateHandoutsInWord [button]
- ObjectNudgeUp [button]
- ShapeArc [toggleButton]
- OutlineMoveDown [button]
- OleConvert [button]
- ObjectNudgeRight [button]
- ShapeStar [toggleButton]
- WebPagePreview [button]
- OutlineCollapse [button]
- OutlineExpandAll [button]
- SnapToGrid [toggleButton]
- ObjectsMultiSelect [button]
- CheckForUpdates [button]
- NextPane [button]
- WebGoBack [button]
- OutlineExpand [button]
- SnapToShapes [toggleButton]
- ViewSlideView [toggleButton]
- CollapseAll [button]
- ExchangeFolder [button]
- WebGoForward [button]
- DocumentLocation [comboBox]
- DrawingCanvasFit [button]
- CombineCharacters [toggleButton]
- ChangeCaseToggle [button]
- ViewOutlineView [toggleButton]
- DrawingCanvasScale [button]
- DrawingCanvasExpand [button]
- BaselineRaise [button]
- OutlineThumbnailsShowHide [button]
- BaselineLower [button]
- SlideShowInAWindow [button]
- WindowFitToPage [button]
- OutlookTaskCreate [button]
- FindNext [button]
- GoToProperty [button]
- LinkBreak [button]
- LinkChange [button]
- CharacterFormattingReset [button]
- WebOptionsDialog [button]
- DrawingCanvasResize [button]
- QuickAccessToolbarCustomization [button]
- RibbonCustomization [button]
- AlternativeText [button]
- AdvancedFileProperties [button]
- HyperlinkOpen [button]
- HyperlinkCopy [button]
- HyperlinkEdit [button]
- HyperlinkRemove [button]
- SoundMaximumFileSize [control]
- InkingStart [button]
- InkToolsClose [button]
- ShapesUnion [button]
- ShapesIntersect [button]
- ShapesFragment [button]
- ShapesCombine [button]
- ShapesSubtract [button]
- ClipArtInsert [toggleButton]
- ReviewEditComment [button]
- CombineShapesMenu [menu]
- PictureBrightnessGallery [gallery]
- PictureContrastGallery [gallery]
- PictureCorrectionsReset [button]
- PictureColorReset [button]
- PictureArtisticEffectsReset [button]
- ShapeConvertToFreeform [button]
- ViewSlideShowView [button]
- CDAudioPlayTrack [button]
- PlaybackTipPlay [button]
- ObjectBringToFrontMenu [splitButton]
- ObjectSendToBackMenu [splitButton]
- LockObject [button]
- UnlockObject [button]
- MovieInsert [splitButton]
- SoundInsertMenu [splitButton]
- SlideShowResolutionGallery [dropDown]
- InsertMovieDialog [button]
- FileOpen [button]
- Collaborate [toggleButton]
- None [toggleButton]
- CollaborateActivity [toggleButton]
- CollaborateComments [toggleButton]
- TransitionByObject [button]
- WhatsNewRecentUpdates [button]
- GroupOfficeInsider [button]
- OfficeFeedbackBackstage [button]
- MarkAsDecorative [button]
- GenerateAutoAltText [button]
- SmartLookup [control]
- ReuseSlidesFromSelection [control]
- AddSlideTitle [button]
- AddOffSlideTitle [button]
- EditSlideTitle [button]
- SetAsSlideTitle [button]
- ManagePrivacy [button]
- ReadingOrderPane [toggleButton]
- SlideHide [toggleButton]
- None [button]
- UpgradeDocument [button]
- OfficeExtensionsAppStore [button]
- OfficeExtensionsGallery2 [gallery]
- OfficeExtensionsFeaturedApp1 [button]
- OfficeExtensionsFeaturedApp2 [button]
- OfficeExtensionsFeaturedApp3 [button]
- OfficeExtensionsFeaturedApp4 [button]
- OfficeExtensionsFeaturedApp5 [button]
- OfficeExtensionsFeaturedApp6 [button]
- GroupPermission [group]
- GroupInsertBarcode [group]
- PictureInsertFromFilePowerPoint [button]
- TableDrawTable [toggleButton]
- SlideNewGallery [gallery]
- ObjectFormatDialog [button]
- ObjectSizeAndPositionDialog [button]
- AnimationCustom [toggleButton]
- TextBoxInsert [toggleButton]
- TextBoxInsertVertical [toggleButton]
- SlideBackgroundFormatDialog [button]
- SlideLayoutGallery [gallery]
- ActionInsert [button]
- ObjectBringToFront [button]
- ObjectSendToBack [button]
- ObjectBringForward [button]
- ObjectSendBackward [button]
- ObjectsGroup [button]
- ObjectsUngroup [button]
- ObjectsRegroup [button]
- FileSaveAs [button]
- MacroPlay [button]
- FormatPainter [control]
- FontSize [comboBox]
- TextHighlightColorPickerLicensed [gallery]
- FontColorPicker [gallery]
- Font [comboBox]
- FontSizeDecrease [button]
- FontSizeIncrease [button]
- Cut [button]
- Paste [button]
- Copy [button]
- ShapesInsertGallery [gallery]
- WordArtInsertGallery [gallery]
- AnimationTransitionGallery [gallery]
- SlideShowFromCurrent [button]
- ViewSlideMasterView [toggleButton]
- ViewSlideSorterView [toggleButton]
- BulletsGallery [gallery]
- AnimationGallery [gallery]
- AnimationAddGallery [gallery]
- TableInsertGallery [gallery]
- UpgradePresentation [button]
- FileExit [button]
- FileNew [button]
- ApplicationOptionsDialog [button]
- None [button]
- FileStartWorkflow [button]
- LeaveReader [button]
- MenuAccessibility [splitButton]
- FilePrint [button]
- OfficeFeedback [button]
- AppGuardTrustPromotion [button]
- OfficeExtensionsGetAddinsButtonOnBackstageBelowOpenDivider [button]
- FilePrepareMenu [menu]
- MenuPublish [menu]
- FilePrintMenu [splitButton]
- FileSaveAsMenu [splitButton]
- FileSendMenu [menu]
- FileServerTasksMenu [menu]
- ReviewProtectPresentationMenu [menu]
- ChartShowData [button]
- None [group]
- Insert3DModelDefaultPlaceholder [button]
- InsertIconPlaceholder [button]
### None (Not in the Ribbon)
-   ShapesUnion [button] < CombineShapesMenu
-   ShapesCombine [button] < CombineShapesMenu
-   ShapesFragment [button] < CombineShapesMenu
-   ShapesIntersect [button] < CombineShapesMenu
-   ShapesSubtract [button] < CombineShapesMenu
-   PictureCorrectionsDialog [button] < PictureBrightnessGallery
-   PictureCorrectionsDialog [button] < PictureContrastGallery
-   ObjectBringToFront [button] < ObjectBringToFrontMenu
-   ObjectBringForward [button] < ObjectBringToFrontMenu
-   ObjectSendToBack [button] < ObjectSendToBackMenu
-   ObjectSendBackward [button] < ObjectSendToBackMenu
-   VideoInsert [button] < MovieInsert
-   MovieFromFileInsert [button] < MovieInsert
-   VideoInsertFromOnline [button] < MovieInsert
-   MovieFromClipOrganizerInsert [button] < MovieInsert
-   AudioInsert [button] < SoundInsertMenu
-   SoundInsertFromFile [button] < SoundInsertMenu
-   SoundInsertFromClipOrganizer [button] < SoundInsertMenu
-   SoundRecord [button] < SoundInsertMenu
-   OfficeExtensionsDialog [button] < OfficeExtensionsGallery2
-   OfficeExtensionsManageOtherAddins [button] < OfficeExtensionsGallery2
-   ExtensibilityExperimentation [button] < OfficeExtensionsGallery2
-   DuplicateSelectedSlides [button] < SlideNewGallery
-   SlidesFromOutline [button] < SlideNewGallery
-   SlidesReuseSlides [toggleButton] < SlideNewGallery
-   ReuseSlides [toggleButton] < SlideNewGallery
-   PasteAsSlide [toggleButton] < SlideNewGallery
-   HighlightingStop [button] < TextHighlightColorPickerLicensed
-   FontColorMoreColorsDialogPowerPoint [button] < FontColorPicker
-   EyedropperFillText [button] < FontColorPicker
-   BulletsAndNumberingBulletsDialog [button] < BulletsGallery
-   AnimationCustomEntranceDialog [button] < AnimationGallery
-   AnimationCustomEmphasisDialog [button] < AnimationGallery
-   AnimationCustomExitDialog [button] < AnimationGallery
-   AnimationCustomPathDialog [button] < AnimationGallery
-   AnimationCustomActionVerbDialog [button] < AnimationGallery
-   AnimationCustomAddEntranceDialog [button] < AnimationAddGallery
-   AnimationCustomAddEmphasisDialog [button] < AnimationAddGallery
-   AnimationCustomAddExitDialog [button] < AnimationAddGallery
-   AnimationCustomAddPathDialog [button] < AnimationAddGallery
-   AnimationCustomAddActionVerbDialog [button] < AnimationAddGallery
-   TableInsert [button] < TableInsertGallery
-   TableDrawTable [toggleButton] < TableInsertGallery
-   ExcelSpreadsheetInsert [button] < TableInsertGallery
-   AccessibilityChecker [button] < MenuAccessibility
-   AltTextPaneRibbon [toggleButton] < MenuAccessibility
-   ReadingOrderPane [toggleButton] < MenuAccessibility
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
-   FileCompatibilityCheckerPowerPoint [button] < FilePrepareMenu
-   FileLinksToFiles [button] < FilePrepareMenu
-   FilePackageForCD [button] < MenuPublish
-   FilePublishSlides [button] < MenuPublish
-   CreateHandoutsInWord [button] < MenuPublish
-   FileSaveToDocumentManagementServer [button] < MenuPublish
-   PrintPreviewAndPrint [button] < FilePrintMenu
-   FilePrintQuick [button] < FilePrintMenu
-   FilePrintPreview [toggleButton] < FilePrintMenu
-   FileSaveAs [button] < FileSaveAsMenu
-   FileSaveAsPowerPointPptx [button] < FileSaveAsMenu
-   FileSaveAsPowerPointPpsx [button] < FileSaveAsMenu
-   FileSaveAsPowerPoint97_2003 [button] < FileSaveAsMenu
-   FileSaveAsPowerPointOpenDocumentPresentation [button] < FileSaveAsMenu
-   AdvertisePublishAs [button] < FileSaveAsMenu
-   FileSaveAsPdfOrXps [button] < FileSaveAsMenu
-   FileSaveAsOtherFormats [button] < FileSaveAsMenu
-   FileSendAsAttachment [button] < FileSendMenu
-   FileEmailAsPdfEmailAttachment [button] < FileSendMenu
-   FileEmailAsXpsEmailAttachment [button] < FileSendMenu
-   FileInternetFax [button] < FileSendMenu
-   BroadcastSlideShowBackstage [button] < FileSendMenu
-   BroadcastResume [button] < FileSendMenu
-   FileCheckIn [button] < FileServerTasksMenu
-   FileCheckOut [button] < FileServerTasksMenu
-   FileCheckOutDiscard [button] < FileServerTasksMenu
-   FileVersionHistory [button] < FileServerTasksMenu
-   FileWorkflowTasks [button] < FileServerTasksMenu
-   FilePermissionUnrestricted [toggleButton] < ReviewProtectPresentationMenu
-   FilePermissionDoNotDistribute [toggleButton] < ReviewProtectPresentationMenu
### GroupPermission
- ReviewProtectPresentationMenu [menu]
-   FilePermissionUnrestricted [toggleButton] < ReviewProtectPresentationMenu
-   FilePermissionDoNotDistribute [toggleButton] < ReviewProtectPresentationMenu
### GroupInsertBarcode
- BarcodeInsert [button]
- LabelInsert [button]

##   ()
### 
- TabSetSmartArtTools [tabSet]
- TabSetChartTools [tabSet]
- TabSetDrawingTools [tabSet]
- TabSetPictureTools [tabSet]
- TabSetInteractiveIndexTools [tabSet]
- TabSetSVGTools [tabSet]
- TabSet3DModelTools [tabSet]
- TabSetTableTools [tabSet]
- TabSetSoundTools [tabSet]
- TabSetAudioTools [tabSet]
- TabSetVideoTools [tabSet]
- TabSetCameoTools [tabSet]
- TabSetMovieTools [tabSet]
- TabSetCDAudioTools [tabSet]
- TabSetInkTools [tabSet]
- TabSetEquationTools [tabSet]
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
- ConvertSmartArtMenu [menu]
-   SmartArtConvedrtToText [button] < ConvertSmartArtMenu
-   SmartArtConvertToShapes [button] < ConvertSmartArtMenu

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
-   EyedropperFill [button] < ShapeFillColorPicker
-   ObjectPictureFill [button] < ShapeFillColorPicker
-   GradientGallery [gallery] < ShapeFillColorPicker
-     ShapeFillMoreGradientsDialog [button] < ShapeFillColorPicker
-   ShapeFillTextureGallery [gallery] < ShapeFillColorPicker
-     MoreTextureOptions [button] < ShapeFillColorPicker
- ShapeOutlineColorPicker [gallery]
-   ObjectBorderOutlineColorMoreColorsDialog [button] < ShapeOutlineColorPicker
-   EyedropperOutline [button] < ShapeOutlineColorPicker
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
-     EyedropperGlow [button] < ShapeEffectsMenu
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
-   EyedropperFillText [button] < TextFillColorPicker
-   TextPictureFill [button] < TextFillColorPicker
-   TextFillGradientGallery [gallery] < TextFillColorPicker
-     TextFillMoreGradientsDialog [button] < TextFillColorPicker
-   TextFillTextureGallery [gallery] < TextFillColorPicker
-     TextFillMoreTextures [button] < TextFillColorPicker
- TextOutlineColorPicker [gallery]
-   TextOutlineColorMoreColorsDialog [button] < TextOutlineColorPicker
-   EyedropperOutlineText [button] < TextOutlineColorPicker
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
-     EyedropperGlowText [button] < TextEffectsMenu
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
-   ObjectsAlignLeftSmart [button] < ObjectAlignMenu
-   ObjectsAlignCenterHorizontalSmart [button] < ObjectAlignMenu
-   ObjectsAlignRightSmart [button] < ObjectAlignMenu
-   ObjectsAlignTopSmart [button] < ObjectAlignMenu
-   ObjectsAlignMiddleVerticalSmart [button] < ObjectAlignMenu
-   ObjectsAlignBottomSmart [button] < ObjectAlignMenu
-   AlignDistributeHorizontally [button] < ObjectAlignMenu
-   AlignDistributeVertically [button] < ObjectAlignMenu
-   ObjectsAlignRelativeToContainerSmart [toggleButton] < ObjectAlignMenu
-   ObjectsAlignSelectedSmart [toggleButton] < ObjectAlignMenu
- ObjectsGroupMenu [menu]
-   ObjectsGroup [button] < ObjectsGroupMenu
-   ObjectsRegroup [button] < ObjectsGroupMenu
-   ObjectsUngroup [button] < ObjectsGroupMenu
- ObjectRotateGallery [gallery]
-   ObjectRotationOptionsDialog [button] < ObjectRotateGallery
### GroupSmartArtSize
- ShapeHeight [control]
- ShapeWidth [control]
- ObjectSizeAndPositionDialog [button (dialogBoxLauncher)]

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
### GroupChartData
- ChartSwitchRowColumn [button]
- ChartEditDataSource [button]
- ChartEditDataMenu [splitButton]
-   ChartShowDataGrid [button] < ChartEditDataMenu
-   ChartShowDataContextualMenu [button] < ChartEditDataMenu
- ChartRefresh [button]
### GroupChartType
- ChartChangeType [button]

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
-   EyedropperFill [button] < ShapeFillColorPicker
-   ObjectPictureFill [button] < ShapeFillColorPicker
-   GradientGallery [gallery] < ShapeFillColorPicker
-     ShapeFillMoreGradientsDialog [button] < ShapeFillColorPicker
-   ShapeFillTextureGallery [gallery] < ShapeFillColorPicker
-     MoreTextureOptions [button] < ShapeFillColorPicker
- ColorPickerChartLineStyles [gallery]
-   ObjectBorderOutlineColorMoreColorsDialog [button] < ColorPickerChartLineStyles
-   EyedropperOutline [button] < ColorPickerChartLineStyles
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
-     EyedropperGlow [button] < ShapeEffectsMenu
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
-   EyedropperFillText [button] < TextFillColorPicker
-   TextPictureFill [button] < TextFillColorPicker
-   TextFillGradientGallery [gallery] < TextFillColorPicker
-     TextFillMoreGradientsDialog [button] < TextFillColorPicker
-   TextFillTextureGallery [gallery] < TextFillColorPicker
-     TextFillMoreTextures [button] < TextFillColorPicker
- TextOutlineColorPicker [gallery]
-   TextOutlineColorMoreColorsDialog [button] < TextOutlineColorPicker
-   EyedropperOutlineText [button] < TextOutlineColorPicker
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
-     EyedropperGlowText [button] < TextEffectsMenu
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
-   ObjectsAlignLeftSmart [button] < ObjectAlignMenu
-   ObjectsAlignCenterHorizontalSmart [button] < ObjectAlignMenu
-   ObjectsAlignRightSmart [button] < ObjectAlignMenu
-   ObjectsAlignTopSmart [button] < ObjectAlignMenu
-   ObjectsAlignMiddleVerticalSmart [button] < ObjectAlignMenu
-   ObjectsAlignBottomSmart [button] < ObjectAlignMenu
-   AlignDistributeHorizontally [button] < ObjectAlignMenu
-   AlignDistributeVertically [button] < ObjectAlignMenu
-   ObjectsAlignRelativeToContainerSmart [toggleButton] < ObjectAlignMenu
-   ObjectsAlignSelectedSmart [toggleButton] < ObjectAlignMenu
- ObjectsGroupMenu [menu]
-   ObjectsGroup [button] < ObjectsGroupMenu
-   ObjectsRegroup [button] < ObjectsGroupMenu
-   ObjectsUngroup [button] < ObjectsGroupMenu
- ObjectRotateGallery [gallery]
-   ObjectRotationOptionsDialog [button] < ObjectRotateGallery
- ObjectsArrangeMenu [menu]
-   ObjectBringToFront [button] < ObjectsArrangeMenu
-   ObjectSendToBack [button] < ObjectsArrangeMenu
-   ObjectBringForward [button] < ObjectsArrangeMenu
-   ObjectSendBackward [button] < ObjectsArrangeMenu
-   ObjectsGroup [button] < ObjectsArrangeMenu
-   ObjectsUngroup [button] < ObjectsArrangeMenu
-   ObjectsRegroup [button] < ObjectsArrangeMenu
-   ObjectAlignMenu [menu] < ObjectsArrangeMenu
-     ObjectsAlignLeftSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignCenterHorizontalSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignRightSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignTopSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignMiddleVerticalSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignBottomSmart [button] < ObjectsArrangeMenu
-     AlignDistributeHorizontally [button] < ObjectsArrangeMenu
-     AlignDistributeVertically [button] < ObjectsArrangeMenu
-     ObjectsAlignRelativeToContainerSmart [toggleButton] < ObjectsArrangeMenu
-     ObjectsAlignSelectedSmart [toggleButton] < ObjectsArrangeMenu
-   ObjectRotateGallery [gallery] < ObjectsArrangeMenu
-     ObjectRotationOptionsDialog [button] < ObjectsArrangeMenu
-   SelectionPane [toggleButton] < ObjectsArrangeMenu
### GroupSize
- ShapeHeight [control]
- ShapeWidth [control]
- ObjectSizeAndPositionDialog [button (dialogBoxLauncher)]

## TabChartToolsDesign  (TabSetChartTools)
### 
- GroupChartData [group]
- GroupChartStyles [group]
### GroupChartData
- ChartSwitchRowColumn [button]
- ChartEditDataSource [button]
- ChartEditDataMenu [splitButton]
-   ChartShowDataGrid [button] < ChartEditDataMenu
-   ChartShowDataContextualMenu [button] < ChartEditDataMenu
- ChartRefresh [button]
### GroupChartStyles
- ChartColorsGallery [gallery]
- ChartStylesGallery [gallery]

## TabChartToolsLayout  (TabSetChartTools)
### 
- GroupChartCurrentSelection [group]
- GroupChartShapes [group]
- GroupChartLabels [group]
- GroupChartAxes [group]
- GroupChartBackground [group]
- GroupChartAnalysis [group]
### GroupChartCurrentSelection
- ChartElementSelector [comboBox]
- ChartFormatSelection [button]
- ChartResetToMatchStyle [button]
### GroupChartShapes
- PictureInsertFromFile [button]
- ShapesInsertGallery [gallery]
- TextBoxInsert [toggleButton]
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
-   EyedropperFill [button] < ShapeFillColorPicker
-   ObjectPictureFill [button] < ShapeFillColorPicker
-   GradientGallery [gallery] < ShapeFillColorPicker
-     ShapeFillMoreGradientsDialog [button] < ShapeFillColorPicker
-   ShapeFillTextureGallery [gallery] < ShapeFillColorPicker
-     MoreTextureOptions [button] < ShapeFillColorPicker
- ShapeOutlineColorPicker [gallery]
-   ObjectBorderOutlineColorMoreColorsDialog [button] < ShapeOutlineColorPicker
-   EyedropperOutline [button] < ShapeOutlineColorPicker
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
-     EyedropperGlow [button] < ShapeEffectsMenu
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
-   EyedropperFillText [button] < TextFillColorPicker
-   TextPictureFill [button] < TextFillColorPicker
-   TextFillGradientGallery [gallery] < TextFillColorPicker
-     TextFillMoreGradientsDialog [button] < TextFillColorPicker
-   TextFillTextureGallery [gallery] < TextFillColorPicker
-     TextFillMoreTextures [button] < TextFillColorPicker
- TextOutlineColorPicker [gallery]
-   TextOutlineColorMoreColorsDialog [button] < TextOutlineColorPicker
-   EyedropperOutlineText [button] < TextOutlineColorPicker
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
-     EyedropperGlowText [button] < TextEffectsMenu
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
-   ObjectsAlignLeftSmart [button] < ObjectAlignMenu
-   ObjectsAlignCenterHorizontalSmart [button] < ObjectAlignMenu
-   ObjectsAlignRightSmart [button] < ObjectAlignMenu
-   ObjectsAlignTopSmart [button] < ObjectAlignMenu
-   ObjectsAlignMiddleVerticalSmart [button] < ObjectAlignMenu
-   ObjectsAlignBottomSmart [button] < ObjectAlignMenu
-   AlignDistributeHorizontally [button] < ObjectAlignMenu
-   AlignDistributeVertically [button] < ObjectAlignMenu
-   ObjectsAlignRelativeToContainerSmart [toggleButton] < ObjectAlignMenu
-   ObjectsAlignSelectedSmart [toggleButton] < ObjectAlignMenu
- ObjectsGroupMenu [menu]
-   ObjectsGroup [button] < ObjectsGroupMenu
-   ObjectsRegroup [button] < ObjectsGroupMenu
-   ObjectsUngroup [button] < ObjectsGroupMenu
- ObjectRotateGallery [gallery]
-   ObjectRotationOptionsDialog [button] < ObjectRotateGallery
- ObjectsArrangeMenu [menu]
-   ObjectBringToFront [button] < ObjectsArrangeMenu
-   ObjectSendToBack [button] < ObjectsArrangeMenu
-   ObjectBringForward [button] < ObjectsArrangeMenu
-   ObjectSendBackward [button] < ObjectsArrangeMenu
-   ObjectsGroup [button] < ObjectsArrangeMenu
-   ObjectsUngroup [button] < ObjectsArrangeMenu
-   ObjectsRegroup [button] < ObjectsArrangeMenu
-   ObjectAlignMenu [menu] < ObjectsArrangeMenu
-     ObjectsAlignLeftSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignCenterHorizontalSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignRightSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignTopSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignMiddleVerticalSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignBottomSmart [button] < ObjectsArrangeMenu
-     AlignDistributeHorizontally [button] < ObjectsArrangeMenu
-     AlignDistributeVertically [button] < ObjectsArrangeMenu
-     ObjectsAlignRelativeToContainerSmart [toggleButton] < ObjectsArrangeMenu
-     ObjectsAlignSelectedSmart [toggleButton] < ObjectsArrangeMenu
-   ObjectRotateGallery [gallery] < ObjectsArrangeMenu
-     ObjectRotationOptionsDialog [button] < ObjectsArrangeMenu
-   SelectionPane [toggleButton] < ObjectsArrangeMenu
### GroupSize
- ShapeHeight [control]
- ShapeWidth [control]
- ObjectSizeAndPositionDialog [button (dialogBoxLauncher)]

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
- TextBoxInsert [toggleButton]
- TextBoxInsertMenu [splitButton]
-   TextBoxInsertHorizontal [toggleButton] < TextBoxInsertMenu
-   TextBoxInsertVertical [toggleButton] < TextBoxInsertMenu
- TextboxFlyoutMenuAnchor [menu]
-   TextBoxInsertHorizontal [toggleButton] < TextboxFlyoutMenuAnchor
-   TextBoxInsertVertical [toggleButton] < TextboxFlyoutMenuAnchor
- CombineShapesGallery [gallery]
### GroupShapeStyles
- ShapeStylesGallery [gallery]
-   ShapeStylesOtherThemeFillsGallery [gallery] < ShapeStylesGallery
- ShapeFillColorPicker [gallery]
-   ObjectFillMoreColorsDialog [button] < ShapeFillColorPicker
-   EyedropperFill [button] < ShapeFillColorPicker
-   ObjectPictureFill [button] < ShapeFillColorPicker
-   GradientGallery [gallery] < ShapeFillColorPicker
-     ShapeFillMoreGradientsDialog [button] < ShapeFillColorPicker
-   ShapeFillTextureGallery [gallery] < ShapeFillColorPicker
-     MoreTextureOptions [button] < ShapeFillColorPicker
- ShapeOutlineColorPicker [gallery]
-   ObjectBorderOutlineColorMoreColorsDialog [button] < ShapeOutlineColorPicker
-   EyedropperOutline [button] < ShapeOutlineColorPicker
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
-     EyedropperGlow [button] < ShapeEffectsMenu
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
-   EyedropperFillText [button] < TextFillColorPicker
-   TextPictureFill [button] < TextFillColorPicker
-   TextFillGradientGallery [gallery] < TextFillColorPicker
-     TextFillMoreGradientsDialog [button] < TextFillColorPicker
-   TextFillTextureGallery [gallery] < TextFillColorPicker
-     TextFillMoreTextures [button] < TextFillColorPicker
- TextOutlineColorPicker [gallery]
-   TextOutlineColorMoreColorsDialog [button] < TextOutlineColorPicker
-   EyedropperOutlineText [button] < TextOutlineColorPicker
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
-     EyedropperGlowText [button] < TextEffectsMenu
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
-   ObjectsAlignLeftSmart [button] < ObjectAlignMenu
-   ObjectsAlignCenterHorizontalSmart [button] < ObjectAlignMenu
-   ObjectsAlignRightSmart [button] < ObjectAlignMenu
-   ObjectsAlignTopSmart [button] < ObjectAlignMenu
-   ObjectsAlignMiddleVerticalSmart [button] < ObjectAlignMenu
-   ObjectsAlignBottomSmart [button] < ObjectAlignMenu
-   AlignDistributeHorizontally [button] < ObjectAlignMenu
-   AlignDistributeVertically [button] < ObjectAlignMenu
-   ObjectsAlignRelativeToContainerSmart [toggleButton] < ObjectAlignMenu
-   ObjectsAlignSelectedSmart [toggleButton] < ObjectAlignMenu
- ObjectsGroupMenu [menu]
-   ObjectsGroup [button] < ObjectsGroupMenu
-   ObjectsRegroup [button] < ObjectsGroupMenu
-   ObjectsUngroup [button] < ObjectsGroupMenu
- ObjectRotateGallery [gallery]
-   ObjectRotationOptionsDialog [button] < ObjectRotateGallery
### GroupSize
- ShapeHeight [control]
- ShapeWidth [control]
- ObjectSizeAndPositionDialog [button (dialogBoxLauncher)]

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
- PictureBackgroundRemoval [toggleButton]
- AutoBackgroundRemoval [toggleButton]
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
-     EyedropperRecolor [button] < PictureColorMenu
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
-   EyedropperOutline [button] < OutlineColorPicker
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
-     EyedropperGlow [button] < PictureEffectsMenu
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
-   ObjectsAlignLeftSmart [button] < ObjectAlignMenu
-   ObjectsAlignCenterHorizontalSmart [button] < ObjectAlignMenu
-   ObjectsAlignRightSmart [button] < ObjectAlignMenu
-   ObjectsAlignTopSmart [button] < ObjectAlignMenu
-   ObjectsAlignMiddleVerticalSmart [button] < ObjectAlignMenu
-   ObjectsAlignBottomSmart [button] < ObjectAlignMenu
-   AlignDistributeHorizontally [button] < ObjectAlignMenu
-   AlignDistributeVertically [button] < ObjectAlignMenu
-   ObjectsAlignRelativeToContainerSmart [toggleButton] < ObjectAlignMenu
-   ObjectsAlignSelectedSmart [toggleButton] < ObjectAlignMenu
- ObjectsGroupMenu [menu]
-   ObjectsGroup [button] < ObjectsGroupMenu
-   ObjectsRegroup [button] < ObjectsGroupMenu
-   ObjectsUngroup [button] < ObjectsGroupMenu
- ObjectRotateGallery [gallery]
-   ObjectRotationOptionsDialog [button] < ObjectRotateGallery
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
- ObjectSizeAndPositionDialog [button (dialogBoxLauncher)]

##   (TabSetInteractiveIndexTools)
### 
- TabInteractiveIndexToolsFormat [tab]

## TabInteractiveIndexToolsFormat  (TabSetInteractiveIndexTools)
### 
- GroupIndexOptions [group]
- GroupIndexStyles [group]
- GroupAltText [group]
- GroupArrange [group]
- GroupSize [group]
### GroupIndexOptions
- UpdateInteractiveIndex [button]
- ResetLayout [button]
- SectionLinkResetPictureMenu [splitButton]
-   ChangeImage [button] < SectionLinkResetPictureMenu
-   SlideLinkReset [button] < SectionLinkResetPictureMenu
- ReturnBehavior [checkBox]
- EnableTransition [checkBox]
- SectionLinkTransitionDuration [control]
### GroupIndexStyles
- PictureStylesGallery [gallery]
- InteractiveIndexUxColorPickerLine [gallery]
-   ObjectBorderOutlineColorMoreColorsDialog [button] < InteractiveIndexUxColorPickerLine
-   EyedropperOutline [button] < InteractiveIndexUxColorPickerLine
-   OutlineWeightGallery [gallery] < InteractiveIndexUxColorPickerLine
-     LineStylesDialog [button] < InteractiveIndexUxColorPickerLine
-   OutlineDashesGallery [gallery] < InteractiveIndexUxColorPickerLine
-     LineStylesDialog [button] < InteractiveIndexUxColorPickerLine
- InteractiveIndexEffectsMenu [menu]
-   PictureEffectsPresetGallery [gallery] < InteractiveIndexEffectsMenu
-     _3DBevelOptionsDialog [button] < InteractiveIndexEffectsMenu
-   PictureEffectsShadowGallery [gallery] < InteractiveIndexEffectsMenu
-     ShadowOptionsDialog [button] < InteractiveIndexEffectsMenu
-   PictureReflectionGallery [gallery] < InteractiveIndexEffectsMenu
-     ReflectionsMoreOptions [button] < InteractiveIndexEffectsMenu
-   PictureEffectsGlowGallery [gallery] < InteractiveIndexEffectsMenu
-     GlowColorPicker [gallery] < InteractiveIndexEffectsMenu
-     GlowColorMoreColorsDialog [button] < InteractiveIndexEffectsMenu
-     EyedropperGlow [button] < InteractiveIndexEffectsMenu
-     GlowsMoreOptions [button] < InteractiveIndexEffectsMenu
-   PictureEffectsSoftEdgesGallery [gallery] < InteractiveIndexEffectsMenu
-     SoftEdgesMoreOptions [button] < InteractiveIndexEffectsMenu
-   _3DBevelPictureTopGallery [gallery] < InteractiveIndexEffectsMenu
-     _3DBevelOptionsDialog [button] < InteractiveIndexEffectsMenu
-   PictureRotationGallery [gallery] < InteractiveIndexEffectsMenu
-     _3DRotationOptionsDialog [button] < InteractiveIndexEffectsMenu
- SlideLinkHideBackgroundBooleanChoice [toggleButton]
- ObjectFormatDialog [button (dialogBoxLauncher)]
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
-   ObjectsAlignRelativeToContainerSmart [toggleButton] < ObjectAlignMenu
-   ObjectsAlignSelectedSmart [toggleButton] < ObjectAlignMenu
- ObjectsGroupMenu [menu]
-   ObjectsGroup [button] < ObjectsGroupMenu
-   ObjectsRegroup [button] < ObjectsGroupMenu
-   ObjectsUngroup [button] < ObjectsGroupMenu
- ObjectRotateGallery [gallery]
-   ObjectRotationOptionsDialog [button] < ObjectRotateGallery
- ObjectsArrangeMenu [menu]
-   ObjectBringToFront [button] < ObjectsArrangeMenu
-   ObjectSendToBack [button] < ObjectsArrangeMenu
-   ObjectBringForward [button] < ObjectsArrangeMenu
-   ObjectSendBackward [button] < ObjectsArrangeMenu
-   ObjectsGroup [button] < ObjectsArrangeMenu
-   ObjectsUngroup [button] < ObjectsArrangeMenu
-   ObjectsRegroup [button] < ObjectsArrangeMenu
-   ObjectAlignMenu [menu] < ObjectsArrangeMenu
-     ObjectsAlignLeftSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignCenterHorizontalSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignRightSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignTopSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignMiddleVerticalSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignBottomSmart [button] < ObjectsArrangeMenu
-     AlignDistributeHorizontally [button] < ObjectsArrangeMenu
-     AlignDistributeVertically [button] < ObjectsArrangeMenu
-     ObjectsAlignRelativeToContainerSmart [toggleButton] < ObjectsArrangeMenu
-     ObjectsAlignSelectedSmart [toggleButton] < ObjectsArrangeMenu
-   ObjectRotateGallery [gallery] < ObjectsArrangeMenu
-     ObjectRotationOptionsDialog [button] < ObjectsArrangeMenu
-   SelectionPane [toggleButton] < ObjectsArrangeMenu
### GroupSize
- ShapeHeight [control]
- ShapeWidth [control]
- ObjectSizeAndPositionDialog [button (dialogBoxLauncher)]

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
-   EyedropperFill [button] < GraphicsColorPicker
- GraphicsOutlineColorPicker [gallery]
-   ObjectBorderOutlineColorMoreColorsDialog [button] < GraphicsOutlineColorPicker
-   EyedropperOutline [button] < GraphicsOutlineColorPicker
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
-     EyedropperGlow [button] < GraphicsEffectsMenu
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
-   ObjectsAlignLeftSmart [button] < ObjectAlignMenu
-   ObjectsAlignCenterHorizontalSmart [button] < ObjectAlignMenu
-   ObjectsAlignRightSmart [button] < ObjectAlignMenu
-   ObjectsAlignTopSmart [button] < ObjectAlignMenu
-   ObjectsAlignMiddleVerticalSmart [button] < ObjectAlignMenu
-   ObjectsAlignBottomSmart [button] < ObjectAlignMenu
-   AlignDistributeHorizontally [button] < ObjectAlignMenu
-   AlignDistributeVertically [button] < ObjectAlignMenu
-   ObjectsAlignRelativeToContainerSmart [toggleButton] < ObjectAlignMenu
-   ObjectsAlignSelectedSmart [toggleButton] < ObjectAlignMenu
- ObjectsGroupMenu [menu]
-   ObjectsGroup [button] < ObjectsGroupMenu
-   ObjectsRegroup [button] < ObjectsGroupMenu
-   ObjectsUngroup [button] < ObjectsGroupMenu
- ObjectRotateGallery [gallery]
-   ObjectRotationOptionsDialog [button] < ObjectRotateGallery
- ObjectsArrangeMenu [menu]
-   ObjectBringToFront [button] < ObjectsArrangeMenu
-   ObjectSendToBack [button] < ObjectsArrangeMenu
-   ObjectBringForward [button] < ObjectsArrangeMenu
-   ObjectSendBackward [button] < ObjectsArrangeMenu
-   ObjectsGroup [button] < ObjectsArrangeMenu
-   ObjectsUngroup [button] < ObjectsArrangeMenu
-   ObjectsRegroup [button] < ObjectsArrangeMenu
-   ObjectAlignMenu [menu] < ObjectsArrangeMenu
-     ObjectsAlignLeftSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignCenterHorizontalSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignRightSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignTopSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignMiddleVerticalSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignBottomSmart [button] < ObjectsArrangeMenu
-     AlignDistributeHorizontally [button] < ObjectsArrangeMenu
-     AlignDistributeVertically [button] < ObjectsArrangeMenu
-     ObjectsAlignRelativeToContainerSmart [toggleButton] < ObjectsArrangeMenu
-     ObjectsAlignSelectedSmart [toggleButton] < ObjectsArrangeMenu
-   ObjectRotateGallery [gallery] < ObjectsArrangeMenu
-     ObjectRotationOptionsDialog [button] < ObjectsArrangeMenu
-   SelectionPane [toggleButton] < ObjectsArrangeMenu
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
- ObjectSizeAndPositionDialog [button (dialogBoxLauncher)]

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

##   (TabSetTableTools)
### 
- TabTableToolsDesign [tab]
- TabTableToolsLayout [tab]

## TabTableToolsDesign  (TabSetTableTools)
### 
- GroupTableStyleOptionsPowerPoint [group]
- GroupTableStylesPowerPoint [group]
- GroupTextStylesTable [group]
- GroupDrawBorders [group]
### GroupTableStyleOptionsPowerPoint
- TableStyleFirstRowPowerPoint [checkBox]
- TableStyleTotalRowPowerPoint [checkBox]
- TableStyleBandedRowsPowerPoint [checkBox]
- TableStyleFirstColumnPowerPoint [checkBox]
- TableStyleLastColumnPowerPoint [checkBox]
- TableStyleBandedColumnsPowerPoint [checkBox]
### GroupTableStylesPowerPoint
- TableStylesGallery [gallery]
-   TableStyleClearTable [button] < TableStylesGallery
- ShadingColorPicker [gallery]
-   ObjectFillMoreColorsDialog [button] < ShadingColorPicker
-   EyedropperFill [button] < ShadingColorPicker
-   ObjectPictureFill [button] < ShadingColorPicker
-   GradientGallery [gallery] < ShadingColorPicker
-     ShapeFillMoreGradientsDialog [button] < ShadingColorPicker
-   ShapeFillTextureGallery [gallery] < ShadingColorPicker
-     MoreTextureOptions [button] < ShadingColorPicker
-   TableBackgroundGallery [gallery] < ShadingColorPicker
-     TableFillColorMoreColorsDialog [button] < ShadingColorPicker
-     EyedropperFill [button] < ShadingColorPicker
-     TableBackgroundPictureFill [button] < ShadingColorPicker
- TableBordersMenu [splitButton]
-   BorderNone [toggleButton] < TableBordersMenu
-   BordersAll [toggleButton] < TableBordersMenu
-   BorderOutside [toggleButton] < TableBordersMenu
-   BorderInside [toggleButton] < TableBordersMenu
-   BorderTop [toggleButton] < TableBordersMenu
-   BorderBottom [toggleButton] < TableBordersMenu
-   BorderLeft [toggleButton] < TableBordersMenu
-   BorderRight [toggleButton] < TableBordersMenu
-   BorderInsideHorizontal [toggleButton] < TableBordersMenu
-   BorderInsideVertical [toggleButton] < TableBordersMenu
-   BorderDiagonalDown [toggleButton] < TableBordersMenu
-   BorderDiagonalUp [toggleButton] < TableBordersMenu
- TableEffectsMenu [menu]
-   TableEffectsCellBevelGallery [gallery] < TableEffectsMenu
-   ObjectEffectShadowGallery [gallery] < TableEffectsMenu
-     ShadowOptionsDialog [button] < TableEffectsMenu
-   ReflectionGallery [gallery] < TableEffectsMenu
-     ReflectionsMoreOptions [button] < TableEffectsMenu
### GroupTextStylesTable
- TableTextStylesGallery [gallery]
-   WordArtClear [button] < TableTextStylesGallery
- TextFillColorPicker [gallery]
-   TextFillColorMoreColorsDialog [button] < TextFillColorPicker
-   EyedropperFillText [button] < TextFillColorPicker
-   TextPictureFill [button] < TextFillColorPicker
-   TextFillGradientGallery [gallery] < TextFillColorPicker
-     TextFillMoreGradientsDialog [button] < TextFillColorPicker
-   TextFillTextureGallery [gallery] < TextFillColorPicker
-     TextFillMoreTextures [button] < TextFillColorPicker
- TextOutlineColorPicker [gallery]
-   TextOutlineColorMoreColorsDialog [button] < TextOutlineColorPicker
-   EyedropperOutlineText [button] < TextOutlineColorPicker
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
-     EyedropperGlowText [button] < TextEffectsMenu
-     TextEffectsGlowsMoreOptions [button] < TextEffectsMenu
-   BevelTextGallery [gallery] < TextEffectsMenu
-     TextEffectsBevelMore3DOptionsDialog [button] < TextEffectsMenu
-   TextEffects3DRotationGallery [gallery] < TextEffectsMenu
-     TextEffects3DRotationOptionsDialog [button] < TextEffectsMenu
-   TextEffectTransformGallery [gallery] < TextEffectsMenu
- WordArtFormatDialog [button (dialogBoxLauncher)]
### GroupDrawBorders
- TableDrawBorderPenStyle [dropDown]
- TableDrawBorderPenWeight [dropDown]
- TableBorderPenColorPicker [gallery]
-   TableBorderColorMoreColorsDialog [button] < TableBorderPenColorPicker
- TableDrawTable [toggleButton]
- TableEraser [toggleButton]

## TabTableToolsLayout  (TabSetTableTools)
### 
- GroupTable [group]
- GroupTableRowsAndColumns [group]
- GroupMerge [group]
- GroupTableCellSize [group]
- GroupAlignment [group]
- GroupTableSize [group]
- GroupArrange [group]
### GroupTable
- TableSelectMenuPowerPoint [menu]
-   TableSelect [button] < TableSelectMenuPowerPoint
-   TableColumnSelect [button] < TableSelectMenuPowerPoint
-   TableRowSelect [button] < TableSelectMenuPowerPoint
- TableShowGridlines [toggleButton]
### GroupTableRowsAndColumns
- TableDeleteRowsAndColumnsMenuWord [menu]
-   TableDeleteColumns [button] < TableDeleteRowsAndColumnsMenuWord
-   TableDeleteRows [button] < TableDeleteRowsAndColumnsMenuWord
-   TableDelete [button] < TableDeleteRowsAndColumnsMenuWord
- TableInsertRowsAbove [button]
- TableInsertRowsBelow [button]
- TableInsertColumnsLeft [button]
- TableInsertColumnsRight [button]
### GroupMerge
- MergeCells [button]
- SplitCells [button]
### GroupTableCellSize
- TableRowHeight [control]
- TableColumnWidth [control]
- TableRowsDistribute [button]
- TableColumnsDistribute [button]
### GroupAlignment
- AlignLeft [toggleButton]
- AlignCenter [toggleButton]
- AlignRight [toggleButton]
- TableCellAlignTop [toggleButton]
- TableCellAlignCenterVertically [toggleButton]
- TableCellAlignBottom [toggleButton]
- TableSetLeftToRight [toggleButton]
- TableSetRightToLeft [toggleButton]
- TextDirectionGallery [gallery]
-   TextDirectionMoreOptionsDialog [button] < TextDirectionGallery
- TableCellMarginsGallery [gallery]
-   TableCellCustomMarginsDialog [button] < TableCellMarginsGallery
### GroupTableSize
- TableHeight [control]
- TableWidth [control]
- TableLockAspectRatio [checkBox]
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
-   ObjectsAlignRelativeToContainerSmart [toggleButton] < ObjectAlignMenu
-   ObjectsAlignSelectedSmart [toggleButton] < ObjectAlignMenu
- ObjectsGroupMenu [menu]
-   ObjectsGroup [button] < ObjectsGroupMenu
-   ObjectsRegroup [button] < ObjectsGroupMenu
-   ObjectsUngroup [button] < ObjectsGroupMenu
- ObjectRotateGallery [gallery]
-   ObjectRotationOptionsDialog [button] < ObjectRotateGallery
- ObjectsArrangeMenu [menu]
-   ObjectBringToFront [button] < ObjectsArrangeMenu
-   ObjectSendToBack [button] < ObjectsArrangeMenu
-   ObjectBringForward [button] < ObjectsArrangeMenu
-   ObjectSendBackward [button] < ObjectsArrangeMenu
-   ObjectsGroup [button] < ObjectsArrangeMenu
-   ObjectsUngroup [button] < ObjectsArrangeMenu
-   ObjectsRegroup [button] < ObjectsArrangeMenu
-   ObjectAlignMenu [menu] < ObjectsArrangeMenu
-     ObjectsAlignLeftSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignCenterHorizontalSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignRightSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignTopSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignMiddleVerticalSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignBottomSmart [button] < ObjectsArrangeMenu
-     AlignDistributeHorizontally [button] < ObjectsArrangeMenu
-     AlignDistributeVertically [button] < ObjectsArrangeMenu
-     ObjectsAlignRelativeToContainerSmart [toggleButton] < ObjectsArrangeMenu
-     ObjectsAlignSelectedSmart [toggleButton] < ObjectsArrangeMenu
-   ObjectRotateGallery [gallery] < ObjectsArrangeMenu
-     ObjectRotationOptionsDialog [button] < ObjectsArrangeMenu
-   SelectionPane [toggleButton] < ObjectsArrangeMenu

##   (TabSetSoundTools)
### 
- TabSoundToolsOptions [tab]

## TabSoundToolsOptions  (TabSetSoundTools)
### 
- GroupPlay [group]
- GroupSoundOptions [group]
- GroupArrange [group]
- GroupSize [group]
### GroupPlay
- MoviePlay [button]
### GroupSoundOptions
- SlideShowVolume [gallery]
- MediaClipToolsHideDuringShow [checkBox]
- MediaClipLoopUntilStopped [checkBox]
- SoundPlaySoundGallery [dropDown]
- ObjectEditDialog [button (dialogBoxLauncher)]
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
-   ObjectsAlignRelativeToContainerSmart [toggleButton] < ObjectAlignMenu
-   ObjectsAlignSelectedSmart [toggleButton] < ObjectAlignMenu
- ObjectsGroupMenu [menu]
-   ObjectsGroup [button] < ObjectsGroupMenu
-   ObjectsRegroup [button] < ObjectsGroupMenu
-   ObjectsUngroup [button] < ObjectsGroupMenu
- ObjectRotateGallery [gallery]
-   ObjectRotationOptionsDialog [button] < ObjectRotateGallery
- ObjectsArrangeMenu [menu]
-   ObjectBringToFront [button] < ObjectsArrangeMenu
-   ObjectSendToBack [button] < ObjectsArrangeMenu
-   ObjectBringForward [button] < ObjectsArrangeMenu
-   ObjectSendBackward [button] < ObjectsArrangeMenu
-   ObjectsGroup [button] < ObjectsArrangeMenu
-   ObjectsUngroup [button] < ObjectsArrangeMenu
-   ObjectsRegroup [button] < ObjectsArrangeMenu
-   ObjectAlignMenu [menu] < ObjectsArrangeMenu
-     ObjectsAlignLeftSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignCenterHorizontalSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignRightSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignTopSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignMiddleVerticalSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignBottomSmart [button] < ObjectsArrangeMenu
-     AlignDistributeHorizontally [button] < ObjectsArrangeMenu
-     AlignDistributeVertically [button] < ObjectsArrangeMenu
-     ObjectsAlignRelativeToContainerSmart [toggleButton] < ObjectsArrangeMenu
-     ObjectsAlignSelectedSmart [toggleButton] < ObjectsArrangeMenu
-   ObjectRotateGallery [gallery] < ObjectsArrangeMenu
-     ObjectRotationOptionsDialog [button] < ObjectsArrangeMenu
-   SelectionPane [toggleButton] < ObjectsArrangeMenu
### GroupSize
- ShapeHeight [control]
- ShapeWidth [control]
- ObjectSizeAndPositionDialog [button (dialogBoxLauncher)]

##   (TabSetAudioTools)
### 
- TabAudioToolsFormat [tab]
- TabAudioToolsEdit [tab]

## TabAudioToolsFormat  (TabSetAudioTools)
### 
- GroupAudioToolsAdjust [group]
- GroupPictureStyles [group]
- GroupArrange [group]
- GroupPictureSize [group]
### GroupAudioToolsAdjust
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
-     EyedropperRecolor [button] < PictureColorMenu
-   PictureSetTransparentColor [toggleButton] < PictureColorMenu
-   PictureColorDialog [button] < PictureColorMenu
- PictureArtisticEffectsGallery [gallery]
-   ArtisticEffectsDialog [button] < PictureArtisticEffectsGallery
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
-   EyedropperOutline [button] < OutlineColorPicker
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
-     EyedropperGlow [button] < PictureEffectsMenu
-     GlowsMoreOptions [button] < PictureEffectsMenu
-   PictureEffectsSoftEdgesGallery [gallery] < PictureEffectsMenu
-     SoftEdgesMoreOptions [button] < PictureEffectsMenu
-   _3DBevelPictureTopGallery [gallery] < PictureEffectsMenu
-     _3DBevelOptionsDialog [button] < PictureEffectsMenu
-   PictureRotationGallery [gallery] < PictureEffectsMenu
-     _3DRotationOptionsDialog [button] < PictureEffectsMenu
- PicturesConvertToSmartArt [gallery]
- PictureFormatDialog [button (dialogBoxLauncher)]
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
-   ObjectsAlignRelativeToContainerSmart [toggleButton] < ObjectAlignMenu
-   ObjectsAlignSelectedSmart [toggleButton] < ObjectAlignMenu
- ObjectsGroupMenu [menu]
-   ObjectsGroup [button] < ObjectsGroupMenu
-   ObjectsRegroup [button] < ObjectsGroupMenu
-   ObjectsUngroup [button] < ObjectsGroupMenu
- ObjectRotateGallery [gallery]
-   ObjectRotationOptionsDialog [button] < ObjectRotateGallery
- ObjectsArrangeMenu [menu]
-   ObjectBringToFront [button] < ObjectsArrangeMenu
-   ObjectSendToBack [button] < ObjectsArrangeMenu
-   ObjectBringForward [button] < ObjectsArrangeMenu
-   ObjectSendBackward [button] < ObjectsArrangeMenu
-   ObjectsGroup [button] < ObjectsArrangeMenu
-   ObjectsUngroup [button] < ObjectsArrangeMenu
-   ObjectsRegroup [button] < ObjectsArrangeMenu
-   ObjectAlignMenu [menu] < ObjectsArrangeMenu
-     ObjectsAlignLeftSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignCenterHorizontalSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignRightSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignTopSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignMiddleVerticalSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignBottomSmart [button] < ObjectsArrangeMenu
-     AlignDistributeHorizontally [button] < ObjectsArrangeMenu
-     AlignDistributeVertically [button] < ObjectsArrangeMenu
-     ObjectsAlignRelativeToContainerSmart [toggleButton] < ObjectsArrangeMenu
-     ObjectsAlignSelectedSmart [toggleButton] < ObjectsArrangeMenu
-   ObjectRotateGallery [gallery] < ObjectsArrangeMenu
-     ObjectRotationOptionsDialog [button] < ObjectsArrangeMenu
-   SelectionPane [toggleButton] < ObjectsArrangeMenu
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
- ObjectSizeAndPositionDialog [button (dialogBoxLauncher)]

## TabAudioToolsEdit  (TabSetAudioTools)
### 
- GroupMediaPreview [group]
- GroupAudioBookmark [group]
- GroupAudioEditing [group]
- GroupAudioOptions [group]
- GroupAudioStyles [group]
- GroupClosedCaptions [group]
- GroupSaveMedia [group]
### GroupMediaPreview
- MediaPlayPreview [button]
### GroupAudioBookmark
- AudioBookmarkAdd [button]
- AudioBookmarkRemove [button]
### GroupAudioEditing
- AudioToolsTrim [button]
- FadeDurationLabelAudio [labelControl]
- AudioFadeInTime [control]
- AudioFadeOutTime [control]
### GroupAudioOptions
- AudioVolumeGallery [gallery]
- AudioStartGallery [dropDown]
- MediaClipPlayAcrossSlides [checkBox]
- MediaClipLoopUntilStopped [checkBox]
- MediaClipToolsHideDuringShow [checkBox]
- MovieRewindAfterPlaying [checkBox]
### GroupAudioStyles
- NoAudioStyle [toggleButton]
- PlayInBackground [toggleButton]
### GroupClosedCaptions
- ClosedCaptionSplitDropdown [splitButton]
-   InsertClosedCaption [button] < ClosedCaptionSplitDropdown
-   RemoveClosedCaption [button] < ClosedCaptionSplitDropdown
### GroupSaveMedia
- ObjectSaveMedia [button]

##   (TabSetVideoTools)
### 
- TabVideoToolsDesign [tab]
- TabVideoToolsEdit [tab]

## TabVideoToolsDesign  (TabSetVideoTools)
### 
- GroupMediaPreview [group]
- GroupVideoToolsAdjust [group]
- GroupVideoStyles [group]
- GroupAltText [group]
- GroupArrange [group]
- GroupVideoSize [group]
### GroupMediaPreview
- MediaPlayPreview [button]
### GroupVideoToolsAdjust
- VideoCorrectionsMenu [menu]
-   PictureBrightnessAndContrastGallery [gallery] < VideoCorrectionsMenu
-   VideoCorrectionsDialog [button] < VideoCorrectionsMenu
- VideoColorMenu [menu]
-   PictureRecolorGallery [gallery] < VideoColorMenu
-   RecolorColorPicker [gallery] < VideoColorMenu
-     PictureRecolorMoreColorsDialog [button] < VideoColorMenu
-     EyedropperRecolor [button] < VideoColorMenu
-   VideoColorDialog [button] < VideoColorMenu
- VideoPosterFrameGallery [menu]
-   VideoPosterFrameCurrent [button] < VideoPosterFrameGallery
-   VideoPosterFrameFromFile [button] < VideoPosterFrameGallery
-   VideoPosterFrameReset [button] < VideoPosterFrameGallery
- VideoDesignResetSplitButton [splitButton]
-   VideoToolsDesignReset [button] < VideoDesignResetSplitButton
-   VideoDesignResetAndSize [button] < VideoDesignResetSplitButton
### GroupVideoStyles
- VideoStylesGallery [gallery]
- VideoShapeGallery [gallery]
- VideoBorderGallery [gallery]
-   ObjectBorderOutlineColorMoreColorsDialog [button] < VideoBorderGallery
-   EyedropperOutline [button] < VideoBorderGallery
-   OutlineWeightGallery [gallery] < VideoBorderGallery
-     LineStylesDialog [button] < VideoBorderGallery
-   OutlineDashesGallery [gallery] < VideoBorderGallery
-     LineStylesDialog [button] < VideoBorderGallery
- VideoEffectsGallery [menu]
-   PictureEffectsPresetGallery [gallery] < VideoEffectsGallery
-     _3DBevelOptionsDialog [button] < VideoEffectsGallery
-   PictureEffectsShadowGallery [gallery] < VideoEffectsGallery
-     ShadowOptionsDialog [button] < VideoEffectsGallery
-   PictureReflectionGallery [gallery] < VideoEffectsGallery
-     ReflectionsMoreOptions [button] < VideoEffectsGallery
-   PictureEffectsGlowGallery [gallery] < VideoEffectsGallery
-     GlowColorPicker [gallery] < VideoEffectsGallery
-     GlowColorMoreColorsDialog [button] < VideoEffectsGallery
-     EyedropperGlow [button] < VideoEffectsGallery
-     GlowsMoreOptions [button] < VideoEffectsGallery
-   PictureEffectsSoftEdgesGallery [gallery] < VideoEffectsGallery
-     SoftEdgesMoreOptions [button] < VideoEffectsGallery
-   _3DBevelPictureTopGallery [gallery] < VideoEffectsGallery
-     _3DBevelOptionsDialog [button] < VideoEffectsGallery
-   PictureRotationGallery [gallery] < VideoEffectsGallery
-     _3DRotationOptionsDialog [button] < VideoEffectsGallery
- ObjectFormatDialog [button (dialogBoxLauncher)]
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
-   ObjectsAlignRelativeToContainerSmart [toggleButton] < ObjectAlignMenu
-   ObjectsAlignSelectedSmart [toggleButton] < ObjectAlignMenu
- ObjectsGroupMenu [menu]
-   ObjectsGroup [button] < ObjectsGroupMenu
-   ObjectsRegroup [button] < ObjectsGroupMenu
-   ObjectsUngroup [button] < ObjectsGroupMenu
- ObjectRotateGallery [gallery]
-   ObjectRotationOptionsDialog [button] < ObjectRotateGallery
- ObjectsArrangeMenu [menu]
-   ObjectBringToFront [button] < ObjectsArrangeMenu
-   ObjectSendToBack [button] < ObjectsArrangeMenu
-   ObjectBringForward [button] < ObjectsArrangeMenu
-   ObjectSendBackward [button] < ObjectsArrangeMenu
-   ObjectsGroup [button] < ObjectsArrangeMenu
-   ObjectsUngroup [button] < ObjectsArrangeMenu
-   ObjectsRegroup [button] < ObjectsArrangeMenu
-   ObjectAlignMenu [menu] < ObjectsArrangeMenu
-     ObjectsAlignLeftSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignCenterHorizontalSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignRightSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignTopSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignMiddleVerticalSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignBottomSmart [button] < ObjectsArrangeMenu
-     AlignDistributeHorizontally [button] < ObjectsArrangeMenu
-     AlignDistributeVertically [button] < ObjectsArrangeMenu
-     ObjectsAlignRelativeToContainerSmart [toggleButton] < ObjectsArrangeMenu
-     ObjectsAlignSelectedSmart [toggleButton] < ObjectsArrangeMenu
-   ObjectRotateGallery [gallery] < ObjectsArrangeMenu
-     ObjectRotationOptionsDialog [button] < ObjectsArrangeMenu
-   SelectionPane [toggleButton] < ObjectsArrangeMenu
### GroupVideoSize
- VideoCrop [toggleButton]
- VideoHeight [control]
- VideoWidth [control]
- ObjectSizeAndPositionDialog [button (dialogBoxLauncher)]

## TabVideoToolsEdit  (TabSetVideoTools)
### 
- GroupMediaPreview [group]
- GroupVideoBookmark [group]
- GroupVideoEditing [group]
- GroupVideoOptions [group]
- GroupClosedCaptions [group]
- GroupSaveMedia [group]
### GroupMediaPreview
- MediaPlayPreview [button]
### GroupVideoBookmark
- VideoBookmarkAdd [button]
- VideoBookmarkRemove [button]
### GroupVideoEditing
- VideoToolsTrim [button]
- FadeDurationLabelVideo [labelControl]
- VideoFadeInTime [control]
- VideoFadeOutTime [control]
### GroupVideoOptions
- VideoVolumeGallery [gallery]
- VideoStartGallery [dropDown]
- MoviePlayFullScreen [checkBox]
- MediaClipToolsHideDuringShowVideo [checkBox]
- MediaClipLoopUntilStopped [checkBox]
- MovieRewindAfterPlaying [checkBox]
### GroupClosedCaptions
- ClosedCaptionSplitDropdown [splitButton]
-   InsertClosedCaption [button] < ClosedCaptionSplitDropdown
-   RemoveClosedCaption [button] < ClosedCaptionSplitDropdown
### GroupSaveMedia
- ObjectSaveMedia [button]

##   (TabSetCameoTools)
### 
- TabDesignCameo [tab]

## TabDesignCameo  (TabSetCameoTools)
### 
- GroupChunkCameoCamera [group]
- GroupCameoStyles [group]
- GroupAltText [group]
- GroupArrange [group]
- GroupCameoSize [group]
### GroupChunkCameoCamera
- DropdownCameoTogglePreviewState [splitButton]
-   CameoEnablePreview [button] < DropdownCameoTogglePreviewState
-   CameoSourceGallery [gallery] < DropdownCameoTogglePreviewState
### GroupCameoStyles
- CameoStylesGallery [gallery]
- CopyCameoAllSlides [button]
- CameoShapeGallery [gallery]
- CameoBorderGallery [gallery]
-   ObjectBorderOutlineColorMoreColorsDialog [button] < CameoBorderGallery
-   EyedropperOutline [button] < CameoBorderGallery
-   OutlineWeightGallery [gallery] < CameoBorderGallery
-     LineStylesDialog [button] < CameoBorderGallery
-   OutlineDashesGallery [gallery] < CameoBorderGallery
-     LineStylesDialog [button] < CameoBorderGallery
- CameoEdgeEffectsGallery [menu]
-   PictureEffectsPresetGallery [gallery] < CameoEdgeEffectsGallery
-     _3DBevelOptionsDialog [button] < CameoEdgeEffectsGallery
-   PictureEffectsShadowGallery [gallery] < CameoEdgeEffectsGallery
-     ShadowOptionsDialog [button] < CameoEdgeEffectsGallery
-   PictureReflectionGallery [gallery] < CameoEdgeEffectsGallery
-     ReflectionsMoreOptions [button] < CameoEdgeEffectsGallery
-   PictureEffectsGlowGallery [gallery] < CameoEdgeEffectsGallery
-     GlowColorPicker [gallery] < CameoEdgeEffectsGallery
-     GlowColorMoreColorsDialog [button] < CameoEdgeEffectsGallery
-     EyedropperGlow [button] < CameoEdgeEffectsGallery
-     GlowsMoreOptions [button] < CameoEdgeEffectsGallery
-   PictureEffectsSoftEdgesGallery [gallery] < CameoEdgeEffectsGallery
-     SoftEdgesMoreOptions [button] < CameoEdgeEffectsGallery
-   _3DBevelPictureTopGallery [gallery] < CameoEdgeEffectsGallery
-     _3DBevelOptionsDialog [button] < CameoEdgeEffectsGallery
-   PictureRotationGallery [gallery] < CameoEdgeEffectsGallery
-     _3DRotationOptionsDialog [button] < CameoEdgeEffectsGallery
- ObjectFormatDialog [button (dialogBoxLauncher)]
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
-   ObjectsAlignRelativeToContainerSmart [toggleButton] < ObjectAlignMenu
-   ObjectsAlignSelectedSmart [toggleButton] < ObjectAlignMenu
- ObjectsGroupMenu [menu]
-   ObjectsGroup [button] < ObjectsGroupMenu
-   ObjectsRegroup [button] < ObjectsGroupMenu
-   ObjectsUngroup [button] < ObjectsGroupMenu
- ObjectRotateGallery [gallery]
-   ObjectRotationOptionsDialog [button] < ObjectRotateGallery
- ObjectsArrangeMenu [menu]
-   ObjectBringToFront [button] < ObjectsArrangeMenu
-   ObjectSendToBack [button] < ObjectsArrangeMenu
-   ObjectBringForward [button] < ObjectsArrangeMenu
-   ObjectSendBackward [button] < ObjectsArrangeMenu
-   ObjectsGroup [button] < ObjectsArrangeMenu
-   ObjectsUngroup [button] < ObjectsArrangeMenu
-   ObjectsRegroup [button] < ObjectsArrangeMenu
-   ObjectAlignMenu [menu] < ObjectsArrangeMenu
-     ObjectsAlignLeftSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignCenterHorizontalSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignRightSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignTopSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignMiddleVerticalSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignBottomSmart [button] < ObjectsArrangeMenu
-     AlignDistributeHorizontally [button] < ObjectsArrangeMenu
-     AlignDistributeVertically [button] < ObjectsArrangeMenu
-     ObjectsAlignRelativeToContainerSmart [toggleButton] < ObjectsArrangeMenu
-     ObjectsAlignSelectedSmart [toggleButton] < ObjectsArrangeMenu
-   ObjectRotateGallery [gallery] < ObjectsArrangeMenu
-     ObjectRotationOptionsDialog [button] < ObjectsArrangeMenu
-   SelectionPane [toggleButton] < ObjectsArrangeMenu
### GroupCameoSize
- CameoCrop [toggleButton]
- SpinnerCameoHeight [control]
- SpinnerCameoWidth [control]
- ObjectSizeAndPositionDialog [button (dialogBoxLauncher)]

##   (TabSetMovieTools)
### 
- TabMovieToolsOptions [tab]

## TabMovieToolsOptions  (TabSetMovieTools)
### 
- GroupPlay [group]
- GroupMovieOptions [group]
- GroupArrange [group]
- GroupSize [group]
### GroupPlay
- MoviePlay [button]
### GroupMovieOptions
- SlideShowVolume [gallery]
- MoviePlayAutomatically [dropDown]
- MediaClipToolsHideDuringShow [checkBox]
- MoviePlayFullScreen [checkBox]
- MediaClipLoopUntilStopped [checkBox]
- MovieRewindAfterPlaying [checkBox]
- ObjectEditDialog [button (dialogBoxLauncher)]
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
-   ObjectsAlignRelativeToContainerSmart [toggleButton] < ObjectAlignMenu
-   ObjectsAlignSelectedSmart [toggleButton] < ObjectAlignMenu
- ObjectsGroupMenu [menu]
-   ObjectsGroup [button] < ObjectsGroupMenu
-   ObjectsRegroup [button] < ObjectsGroupMenu
-   ObjectsUngroup [button] < ObjectsGroupMenu
- ObjectRotateGallery [gallery]
-   ObjectRotationOptionsDialog [button] < ObjectRotateGallery
- ObjectsArrangeMenu [menu]
-   ObjectBringToFront [button] < ObjectsArrangeMenu
-   ObjectSendToBack [button] < ObjectsArrangeMenu
-   ObjectBringForward [button] < ObjectsArrangeMenu
-   ObjectSendBackward [button] < ObjectsArrangeMenu
-   ObjectsGroup [button] < ObjectsArrangeMenu
-   ObjectsUngroup [button] < ObjectsArrangeMenu
-   ObjectsRegroup [button] < ObjectsArrangeMenu
-   ObjectAlignMenu [menu] < ObjectsArrangeMenu
-     ObjectsAlignLeftSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignCenterHorizontalSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignRightSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignTopSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignMiddleVerticalSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignBottomSmart [button] < ObjectsArrangeMenu
-     AlignDistributeHorizontally [button] < ObjectsArrangeMenu
-     AlignDistributeVertically [button] < ObjectsArrangeMenu
-     ObjectsAlignRelativeToContainerSmart [toggleButton] < ObjectsArrangeMenu
-     ObjectsAlignSelectedSmart [toggleButton] < ObjectsArrangeMenu
-   ObjectRotateGallery [gallery] < ObjectsArrangeMenu
-     ObjectRotationOptionsDialog [button] < ObjectsArrangeMenu
-   SelectionPane [toggleButton] < ObjectsArrangeMenu
### GroupSize
- ShapeHeight [control]
- ShapeWidth [control]
- ObjectSizeAndPositionDialog [button (dialogBoxLauncher)]

##   (TabSetCDAudioTools)
### 
- TabCDAudioToolsOptions [tab]

## TabCDAudioToolsOptions  (TabSetCDAudioTools)
### 
- GroupPlay [group]
- GroupCDAudioSetup [group]
- GroupArrange [group]
- GroupSize [group]
### GroupPlay
- MoviePlay [button]
### GroupCDAudioSetup
- SlideShowVolume [gallery]
- CDAudioStartPlayingLabel [labelControl]
- CDAudioStartTrack [control]
- CDAudioStartTime [control]
- CDAudioStopPlayingLabel [labelControl]
- CDAudioStopTrack [control]
- CDAudioStopTime [control]
- CDAudioPlayTrackAutomatically [dropDown]
- MediaClipLoopUntilStoppedAudio [checkBox]
- MediaClipToolsHideDuringShowAudio [checkBox]
- ObjectEditDialog [button (dialogBoxLauncher)]
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
-   ObjectsAlignRelativeToContainerSmart [toggleButton] < ObjectAlignMenu
-   ObjectsAlignSelectedSmart [toggleButton] < ObjectAlignMenu
- ObjectsGroupMenu [menu]
-   ObjectsGroup [button] < ObjectsGroupMenu
-   ObjectsRegroup [button] < ObjectsGroupMenu
-   ObjectsUngroup [button] < ObjectsGroupMenu
- ObjectRotateGallery [gallery]
-   ObjectRotationOptionsDialog [button] < ObjectRotateGallery
- ObjectsArrangeMenu [menu]
-   ObjectBringToFront [button] < ObjectsArrangeMenu
-   ObjectSendToBack [button] < ObjectsArrangeMenu
-   ObjectBringForward [button] < ObjectsArrangeMenu
-   ObjectSendBackward [button] < ObjectsArrangeMenu
-   ObjectsGroup [button] < ObjectsArrangeMenu
-   ObjectsUngroup [button] < ObjectsArrangeMenu
-   ObjectsRegroup [button] < ObjectsArrangeMenu
-   ObjectAlignMenu [menu] < ObjectsArrangeMenu
-     ObjectsAlignLeftSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignCenterHorizontalSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignRightSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignTopSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignMiddleVerticalSmart [button] < ObjectsArrangeMenu
-     ObjectsAlignBottomSmart [button] < ObjectsArrangeMenu
-     AlignDistributeHorizontally [button] < ObjectsArrangeMenu
-     AlignDistributeVertically [button] < ObjectsArrangeMenu
-     ObjectsAlignRelativeToContainerSmart [toggleButton] < ObjectsArrangeMenu
-     ObjectsAlignSelectedSmart [toggleButton] < ObjectsArrangeMenu
-   ObjectRotateGallery [gallery] < ObjectsArrangeMenu
-     ObjectRotationOptionsDialog [button] < ObjectsArrangeMenu
-   SelectionPane [toggleButton] < ObjectsArrangeMenu
### GroupSize
- ShapeHeight [control]
- ShapeWidth [control]
- ObjectSizeAndPositionDialog [button (dialogBoxLauncher)]

##   (TabSetInkTools)
### 
- TabInkToolsPens [tab]

## TabInkToolsPens  (TabSetInkTools)
### 
- GroupPensWrite [group]
- GroupPens [group]
- GroupDiagramming [group]
- GroupInkSelect [group]
- GroupInkFormat [group]
- GroupInkPens [group]
- GroupInkClose [group]
### GroupPensWrite
- InkBallpointPen [toggleButton]
- InkHighlighter [toggleButton]
- InkEraserMenu [splitButton]
-   CurrentEraser [toggleButton] < InkEraserMenu
-   InkEraser [toggleButton] < InkEraserMenu
-   PointEraserSmall [toggleButton] < InkEraserMenu
-   PointEraserMedium [toggleButton] < InkEraserMenu
-   SegmentEraser [toggleButton] < InkEraserMenu
- LassoSelect [toggleButton]
- ObjectsSelect [toggleButton]
### GroupPens
- PensGallery [gallery]
-   PenAddFavorite [button] < PensGallery
- InkColorPicker [gallery]
-   InkColorMoreColorsDialog [button] < InkColorPicker
-   EyedropperOutline [button] < InkColorPicker
- LineThickness [gallery]
-   LineStylesDialog [button] < LineThickness
### GroupDiagramming
- DiagrammingOnline [toggleButton]
### GroupInkSelect
- ObjectsSelect [toggleButton]
### GroupInkFormat
- ObjectsSelect [toggleButton]
### GroupInkPens
- InkFeltTipPen [toggleButton]
### GroupInkClose
- InkStopInkingReadingView [button]

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

##   (TabSetAccessibleAuthoring)
### 
- TabAccessibleAuthoring [tab]

## TabAccessibleAuthoring  (TabSetAccessibleAuthoring)
### 
- GroupAccessibleAuthoringReview [group]
- GroupColorAndContrast [group]
- GroupScreenReader [group]
- GroupCaptionsAndSubtitles [group]
- GroupAccessibleAuthoringResources [group]
### GroupAccessibleAuthoringReview
- AccessibilityChecker [button]
### GroupColorAndContrast
- FontColorPicker [gallery]
-   FontColorMoreColorsDialogPowerPoint [button] < FontColorPicker
-   EyedropperFillText [button] < FontColorPicker
- TextHighlightColorPickerLicensed [gallery]
-   HighlightingStop [button] < TextHighlightColorPickerLicensed
- ShapeFillColorPicker [gallery]
-   ObjectFillMoreColorsDialog [button] < ShapeFillColorPicker
-   EyedropperFill [button] < ShapeFillColorPicker
-   ObjectPictureFill [button] < ShapeFillColorPicker
-   GradientGallery [gallery] < ShapeFillColorPicker
-     ShapeFillMoreGradientsDialog [button] < ShapeFillColorPicker
-   ShapeFillTextureGallery [gallery] < ShapeFillColorPicker
-     MoreTextureOptions [button] < ShapeFillColorPicker
- ShapeOutlineColorPicker [gallery]
-   ObjectBorderOutlineColorMoreColorsDialog [button] < ShapeOutlineColorPicker
-   EyedropperOutline [button] < ShapeOutlineColorPicker
-   OutlineWeightGallery [gallery] < ShapeOutlineColorPicker
-     LineStylesDialog [button] < ShapeOutlineColorPicker
-   GallerySketchyShapes [gallery] < ShapeOutlineColorPicker
-     MoreLines2 [button] < ShapeOutlineColorPicker
-   OutlineDashesGallery [gallery] < ShapeOutlineColorPicker
-     LineStylesDialog [button] < ShapeOutlineColorPicker
-   ArrowStyleGallery [gallery] < ShapeOutlineColorPicker
-     ArrowsMore [button] < ShapeOutlineColorPicker
- InspectWithoutColor [toggleButton]
- LauncherAccessibleAuthoringColorAndContrast [button (dialogBoxLauncher)]
### GroupScreenReader
- Spelling [button]
- DropdownSlideTitle [splitButton]
-   SlideTitleDefault [button] < DropdownSlideTitle
-   EditSlideTitle [button] < DropdownSlideTitle
-   SetAsSlideTitle [button] < DropdownSlideTitle
-   AddSlideTitle [button] < DropdownSlideTitle
-   AddOffSlideTitle [button] < DropdownSlideTitle
- ReadingOrderPane [toggleButton]
- AltTextPaneRibbon [toggleButton]
- DecorativeToggleButton [toggleButton]
- ObjectsGroupMenu [menu]
-   ObjectsGroup [button] < ObjectsGroupMenu
-   ObjectsRegroup [button] < ObjectsGroupMenu
-   ObjectsUngroup [button] < ObjectsGroupMenu
- InsertHeaderRowSplitDropdown [splitButton]
-   InsertHeaderRow [button] < InsertHeaderRowSplitDropdown
-   TableStructureFirstRowToggleButton [toggleButton] < InsertHeaderRowSplitDropdown
-   InsertHeaderColumn [button] < InsertHeaderRowSplitDropdown
-   TableStructureFirstColToggleButton [toggleButton] < InsertHeaderRowSplitDropdown
- EditHyperlinkText [button]
- LauncherAccessibleAuthoringScreenReader [button (dialogBoxLauncher)]
### GroupCaptionsAndSubtitles
- UseSubtitles [checkBox]
- SubtitleSettingsMenu [menu]
-   SubtitlesSpokenLanguage [gallery] < SubtitleSettingsMenu
-   SubtitlesDisplayLanguage [gallery] < SubtitleSettingsMenu
-   SubtitlesMicrophoneSettings [gallery] < SubtitleSettingsMenu
-   SubtitlePositionBottomOverlayed [toggleButton] < SubtitleSettingsMenu
-   SubtitlePositionTopOverlayed [toggleButton] < SubtitleSettingsMenu
-   SubtitlePositionBelowSlide [toggleButton] < SubtitleSettingsMenu
-   SubtitlePositionAboveSlide [toggleButton] < SubtitleSettingsMenu
-   SubtitleMoreSettings [button] < SubtitleSettingsMenu
- ClosedCaptionSplitDropdown [splitButton]
-   InsertClosedCaption [button] < ClosedCaptionSplitDropdown
-   RemoveClosedCaption [button] < ClosedCaptionSplitDropdown
- LauncherAccessibleAuthoringCaptionsAndSubtitles [button (dialogBoxLauncher)]
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
- GroupUpdatesAvailable [group]
- None [group]
- GroupBroadcastResumeWarning [group]
- GroupBroadcastEndWarning [group]
- GroupOptimizeCompatibility [group]
- GroupCompressMedia [group]
- GroupARMPreview [group]
- GroupWorkflows [group]
- GrupPolicyWarning [group]
- GroupSecurityInformation [group]
- GroupSecurityWarning [group]
- None [group]
- AppGuardTrustPromotion [button]
- GroupProtectedView [group]
- GroupReadOnly [group]
- GroupDownloading [group]
- GroupReadOnlyServer [group]
- GroupReadOnlyOffline [group]
- GroupReadOnlyCheckOutRequired [group]
- GroupCachedOfflineMergeable [group]
- GroupReadOnlyCachedMergeable [group]
- GroupReadOnlyCached [group]
- GroupCheckIn [group]
- GroupFileConvert [group]
- OfficePolicyTips [group]
- GroupAllAuthors [group]
- GroupSignatures [group]
- GroupPermissions [group]
- GroupPrepareForSharing [group]
- DocuemntVersionHistory [group]
- GroupVersions [group]
- GroupThumbnail [group]
- GroupDocumentProperties [group]
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
- GroupDownloading [group]
- GroupPrintSettings [group]
- GroupPrintPreview [group]
### GroupPrintSettings
- FilePrintCustomRange [control]

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
-   BroadcastSlideShowBackstage [task] < Share
-     GroupBroadcast [group] < Share
-   FilePublishSlides [task] < Share
-     GroupPublishSlides [group] < Share
-     FilePublishSlides [button] < Share
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
-   ExportToVideo [task] < FileTypes
-     GroupExportToVideo [group] < FileTypes
-   ExportToAnimatedGif [task] < FileTypes
-     GroupExportToAnimatedGif [group] < FileTypes
-   FilePackageForCD [task] < FileTypes
-     GroupPackageForCD [group] < FileTypes
-   CreateHandoutsInWord [task] < FileTypes
-     GroupCreateHandoutsInWord [group] < FileTypes
-   FileTypeChange [task] < FileTypes
-     GroupFileSave [group] < FileTypes
-   UploadToMSStream [task] < FileTypes
-     UploadToMSStreamExportGroup [group] < FileTypes

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
- ContextMenuCommentMarker [contextMenu]
- ContextMenuCurve [contextMenu]
- ContextMenuDrawnObject [contextMenu]
- ContextMenuFrame [contextMenu]
- None [contextMenu]
- ContextMenuMotionPath [contextMenu]
- ContextMenuPathPoint [contextMenu]
- ContextMenuPathSegment [contextMenu]
- ContextMenuOfficePreviewHandlerPowerPoint [contextMenu]
- ContextMenuRevision [contextMenu]
- ContextMenuPodiumMonitorSettings [contextMenu]
- ContextMenuRightDragDrop [contextMenu]
- ContextMenuSlideShow [contextMenu]
- ContextMenuShowBrowse [contextMenu]
- ContextMenuShowHyperlink [contextMenu]
- ContextMenuShowPointerOptions [contextMenu]
- ContextMenuEndShow [contextMenu]
- ContextMenuSlideSync [contextMenu]
- ContextMenuSpell [contextMenu]
- ContextMenuTextAlternates [contextMenu]
- ContextMenuSlideGap [contextMenu]
- ContextMenuRPVSettings [contextMenu]
- ContextMenuRPVDeviceSelection [contextMenu]
- ContextMenuOverflow [contextMenu]
- ContextMenuRSCameraOptions [contextMenu]
- ContextMenuRSMicOptions [contextMenu]
- ContextMenuCameraMode [contextMenu]
- ContextMenuClearRecordings [contextMenu]
- ContextMenuRetake [contextMenu]
- ContextMenuHelp [contextMenu]
- ContextMenuLabelSection [contextMenu]
- ContextMenuSlideSorter [contextMenu]
- ContextMenuThumbnail [contextMenu]
- ContextMenuPicture [contextMenu]
- ContextMenuGraphic [contextMenu]
- ContextMenu3DModel [contextMenu]
- ContextMenuSectionLink [contextMenu]
- ContextMenuInteractiveIndex [contextMenu]
- ContextMenuAudio [contextMenu]
- ContextMenuVideo [contextMenu]
- ContextMenuOnlineVideo [contextMenu]
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
- ContextMenuNotesEditText [contextMenu]
- ContextMenuTextEditOutline [contextMenu]
- ContextMenuGraphicsCompatibility [contextMenu]
- ContextMenuGraphicOleClassic [contextMenu]
- ContextMenuActiveXControl [contextMenu]
- ContextMenuTable [contextMenu]
- ContextMenuTableWhole [contextMenu]
- ContextMenuSmartArtContentPane [contextMenu]
- ContextMenuSmartArtBackground [contextMenu]
- ContextMenuSmartArtEditSmartArt [contextMenu]
- ContextMenuSmartArtEdit1DShape [contextMenu]
- ContextMenuSmartArtEditText [contextMenu]
- ContextMenuSmartArtEdit1DShapeText [contextMenu]
- ContextMenuGuide [contextMenu]
- ContextMenuOfficeWebExtension [contextMenu]
- CtxPodiumNotesModeSelection [contextMenu]
- ContextMenuPresenterViewNotesEditText [contextMenu]
- None [contextMenu]
- None [gallery]
- None [contextMenu]
- None [gallery]
### ContextMenuPodiumMonitorSettings
- PodiumSwitchMonitors [button]
- PodiumCloneSlideShow [button]
### ContextMenuTextAlternates
- ShowTextAlternates [gallery]
### ContextMenuSlideGap
- PasteGalleryMini [gallery]
- SlideNew [button]
- ReuseSlides [toggleButton]
- PasteAsSlide [toggleButton]
- SectionAdd [button]
- SlideMasterInsertLayout [button]
### ContextMenuRPVSettings
- AudioSourceMenu [menu]
-   None [toggleButton] < AudioSourceMenu
- RecordAudio [toggleButton]
### ContextMenuRPVDeviceSelection
- AudioSourceMenu [menu]
-   None [toggleButton] < AudioSourceMenu
- VideoSourceMenu [menu]
-   None [toggleButton] < VideoSourceMenu
- VideoEffectSourceMenu [menu]
-   None [toggleButton] < VideoEffectSourceMenu
### ContextMenuOverflow
- AudioSourceMenu [menu]
-   None [toggleButton] < AudioSourceMenu
- VideoSourceMenu [menu]
-   None [toggleButton] < VideoSourceMenu
- ClearCurrentSlideRecording [button]
- ClearAllSlidesRecordings [button]
- ResetToCameoOnCurrentSlide [button]
- ResetToCameoOnAllSlides [button]
### ContextMenuRSCameraOptions
- None [toggleButton]
- None [toggleButton]
### ContextMenuRSMicOptions
- None [toggleButton]
### ContextMenuCameraMode
- None [toggleButton]
### ContextMenuClearRecordings
- ClearCurrentSlideRecording [button]
- ClearAllSlidesRecordings [button]
### ContextMenuRetake
- PodiumRetakeOnCurrentSlide [button]
- PodiumRetakeOnAllSlides [button]
- ClearCurrentSlideRecording [button]
- ClearAllSlidesRecordings [button]
- ResetToCameoOnCurrentSlide [button]
- ResetToCameoOnAllSlides [button]
### ContextMenuHelp
- PodiumRecordSlideshowTeachingTour [button]
- PodiumRecordSlideshowShortcutGuide [button]
### ContextMenuLabelSection
- SectionRename [button]
- SectionMergeWithPrevious [button]
- SectionDelete [button]
- SectionRemoveAll [button]
- SectionMoveUp [button]
- SectionMoveDown [button]
- SectionCollapseAll [button]
- SectionExpandAll [button]
### ContextMenuSlideSorter
- Cut [button]
- Copy [button]
- PasteGalleryMini [gallery]
- SelectAll [button]
- SlideNew [button]
- ReuseSlides [toggleButton]
- PasteAsSlide [toggleButton]
- DuplicateSelectedSlides [button]
- DeleteSlideContextual [button]
- SectionAdd [button]
- FilePublishSlides [button]
- SlideSyncMenu [menu]
- SlideBackgroundFormatDialog [button]
- SlideHide [toggleButton]
- HideSlide [button]
- LinkToSlide [button]
- PeopleNearMe [gallery]
### ContextMenuThumbnail
- Cut [button]
- Copy [button]
- PasteGalleryMini [gallery]
- SelectAll [button]
- SlideNew [button]
- ReuseSlides [toggleButton]
- PasteAsSlide [toggleButton]
- SlideMasterInsertLayout [button]
- DuplicateSelectedSlides [button]
- DeleteSlideContextual [button]
- SectionAdd [button]
- SlideMasterPreserveMaster [toggleButton]
- SlideMasterRenameMaster [button]
- FilePublishSlides [button]
- SlideSyncMenu [menu]
- SlideLayoutGallery [gallery]
- SlideReset [button]
- SlideMasterMasterLayout [button]
- SlideBackgroundFormatDialog [button]
- PhotoAlbumEdit [button]
- SlideHide [toggleButton]
- HideSlide [button]
- LinkToSlide [button]
- ReviewNewComment [button]
- PeopleNearMe [gallery]
### ContextMenuPicture
- Cut [button]
- Copy [button]
- PasteGalleryMini [gallery]
- PictureEdit [button]
- VideoToolsTrim [button]
- AudioToolsTrim [button]
- MoviePlay [button]
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
- LockObject [button]
- UnlockObject [button]
- InsertLinkGallery [gallery]
-   None [button] < InsertLinkGallery
- HyperlinkOpenInMenu [menu]
-   HyperlinkOpenInWeb [button] < HyperlinkOpenInMenu
-   HyperlinkOpenInNative [button] < HyperlinkOpenInMenu
- HyperlinkOpenDefaultSettingMenu [menu]
-   OpenDefaultWeb [button] < HyperlinkOpenDefaultSettingMenu
-   OpenDefaultNative [button] < HyperlinkOpenDefaultSettingMenu
- HyperlinkInsertPowerPoint [button]
- BookshelfDefine [button]
- HyperlinkEdit [button]
- HyperlinkOpen [button]
- HyperlinkCopy [button]
- HyperlinkRemove [button]
- PictureAttributionNavigateToSource [button]
- ObjectSaveAsPicture [button]
- BlackWhiteSettingMenu [menu]
- AutoAlternativeText [button]
- ObjectSizeAndPositionDialog [button]
- ObjectFormatDialog [button]
- CopyCameoAllSlides [button]
- ReviewNewComment [button]
- PeopleNearMe [gallery]
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
- LockObject [button]
- UnlockObject [button]
- InsertLinkGallery [gallery]
-   None [button] < InsertLinkGallery
- HyperlinkOpenInMenu [menu]
-   HyperlinkOpenInWeb [button] < HyperlinkOpenInMenu
-   HyperlinkOpenInNative [button] < HyperlinkOpenInMenu
- HyperlinkOpenDefaultSettingMenu [menu]
-   OpenDefaultWeb [button] < HyperlinkOpenDefaultSettingMenu
-   OpenDefaultNative [button] < HyperlinkOpenDefaultSettingMenu
- HyperlinkInsertPowerPoint [button]
- BookshelfDefine [button]
- HyperlinkEdit [button]
- HyperlinkOpen [button]
- HyperlinkCopy [button]
- HyperlinkRemove [button]
- ObjectSaveAsPicture [button]
- BlackWhiteSettingMenu [menu]
- AutoAlternativeText [button]
- ObjectSizeAndPositionDialog [button]
- GraphicsFormatDialog [button]
- ReviewNewComment [button]
- PeopleNearMe [gallery]
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
- LockObject [button]
- UnlockObject [button]
- InsertLinkGallery [gallery]
-   None [button] < InsertLinkGallery
- HyperlinkOpenInMenu [menu]
-   HyperlinkOpenInWeb [button] < HyperlinkOpenInMenu
-   HyperlinkOpenInNative [button] < HyperlinkOpenInMenu
- HyperlinkOpenDefaultSettingMenu [menu]
-   OpenDefaultWeb [button] < HyperlinkOpenDefaultSettingMenu
-   OpenDefaultNative [button] < HyperlinkOpenDefaultSettingMenu
- HyperlinkInsertPowerPoint [button]
- BookshelfDefine [button]
- HyperlinkEdit [button]
- HyperlinkOpen [button]
- HyperlinkCopy [button]
- HyperlinkRemove [button]
- ObjectSaveAsPicture [button]
- PictureAttributionNavigateToSource [button]
- AutoAlternativeText [button]
- ObjectSizeAndPositionDialog [button]
- FormatDialog3DModel [button]
- ReviewNewComment [button]
- PeopleNearMe [gallery]
### ContextMenuSectionLink
- Cut [button]
- Copy [button]
- PasteGalleryMini [gallery]
- SectionLinkResetPictureMenu [splitButton]
-   ChangeImage [button] < SectionLinkResetPictureMenu
-   SlideLinkReset [button] < SectionLinkResetPictureMenu
- MSPPTOpenSectionLink [button]
- SlideLinkHideBackground [checkBox]
- LockObject [button]
- UnlockObject [button]
- ObjectSaveAsPicture [button]
- ObjectSizeAndPositionDialog [button]
- ObjectFormatDialog [button]
- PictureAttributionNavigateToSource [button]
- ReviewNewComment [button]
- PeopleNearMe [gallery]
### ContextMenuInteractiveIndex
- Cut [button]
- Copy [button]
- PasteGalleryMini [gallery]
- SectionLinkResetPictureMenu [splitButton]
-   ChangeImage [button] < SectionLinkResetPictureMenu
-   SlideLinkReset [button] < SectionLinkResetPictureMenu
- MSPPTOpenSectionLink [button]
- SlideLinkHideBackground [checkBox]
- UpdateInteractiveIndex [button]
- ResetLayout [button]
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
- LockObject [button]
- UnlockObject [button]
- ObjectSaveAsPicture [button]
- AutoAlternativeText [button]
- ObjectSizeAndPositionDialog [button]
- ObjectFormatDialog [button]
- ReviewNewComment [button]
- PeopleNearMe [gallery]
### ContextMenuAudio
- Cut [button]
- Copy [button]
- PasteGalleryMini [gallery]
- MoviePlay [button]
- PictureChange [button]
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
- LockObject [button]
- UnlockObject [button]
- InsertLinkGallery [gallery]
-   None [button] < InsertLinkGallery
- HyperlinkOpenInMenu [menu]
-   HyperlinkOpenInWeb [button] < HyperlinkOpenInMenu
-   HyperlinkOpenInNative [button] < HyperlinkOpenInMenu
- HyperlinkOpenDefaultSettingMenu [menu]
-   OpenDefaultWeb [button] < HyperlinkOpenDefaultSettingMenu
-   OpenDefaultNative [button] < HyperlinkOpenDefaultSettingMenu
- HyperlinkInsertPowerPoint [button]
- BookshelfDefine [button]
- HyperlinkEdit [button]
- HyperlinkOpen [button]
- HyperlinkCopy [button]
- HyperlinkRemove [button]
- MediaUpgrade [button]
- ObjectSaveMedia [button]
- ObjectSaveAsPicture [button]
- BlackWhiteSettingMenu [menu]
- AutoAlternativeText [button]
- ObjectSizeAndPositionDialog [button]
- ObjectFormatDialog [button]
- InsertClosedCaption [button]
- RemoveClosedCaption [button]
- ReviewNewComment [button]
- PeopleNearMe [gallery]
### ContextMenuVideo
- Cut [button]
- Copy [button]
- PasteGalleryMini [gallery]
- MoviePlay [button]
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
- LockObject [button]
- UnlockObject [button]
- MediaUpgrade [button]
- ObjectSaveMedia [button]
- ObjectSaveAsPicture [button]
- BlackWhiteSettingMenu [menu]
- AutoAlternativeText [button]
- ObjectSizeAndPositionDialog [button]
- ObjectFormatDialog [button]
- PictureAttributionNavigateToSource [button]
- InsertClosedCaption [button]
- RemoveClosedCaption [button]
- ReviewNewComment [button]
- PeopleNearMe [gallery]
### ContextMenuOnlineVideo
- Cut [button]
- Copy [button]
- PasteGalleryMini [gallery]
- MoviePlay [button]
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
- LockObject [button]
- UnlockObject [button]
- ObjectSaveAsPicture [button]
- BlackWhiteSettingMenu [menu]
- AutoAlternativeText [button]
- ObjectSizeAndPositionDialog [button]
- ObjectFormatDialog [button]
- PictureAttributionNavigateToSource [button]
- ReviewNewComment [button]
- PeopleNearMe [gallery]
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
- LockObject [button]
- UnlockObject [button]
- InsertLinkGallery [gallery]
-   None [button] < InsertLinkGallery
- HyperlinkOpenInMenu [menu]
-   HyperlinkOpenInWeb [button] < HyperlinkOpenInMenu
-   HyperlinkOpenInNative [button] < HyperlinkOpenInMenu
- HyperlinkOpenDefaultSettingMenu [menu]
-   OpenDefaultWeb [button] < HyperlinkOpenDefaultSettingMenu
-   OpenDefaultNative [button] < HyperlinkOpenDefaultSettingMenu
- HyperlinkInsertPowerPoint [button]
- BookshelfDefine [button]
- HyperlinkEdit [button]
- HyperlinkOpen [button]
- HyperlinkCopy [button]
- HyperlinkRemove [button]
- Insights [button]
- PasteAsSlide [button]
- CollectorFromSelection [control]
- ReuseSlidesFromSelection [control]
- ObjectSaveAsPicture [button]
- BlackWhiteSettingMenu [menu]
- Translator [button]
- AutoAlternativeText [button]
- ObjectSetShapeDefaults [button]
- ObjectSizeAndPositionDialog [button]
- ObjectFormatDialog [button]
- PictureAttributionNavigateToSource [button]
- PlaceholderOrientation [menu]
-   PlaceholderOrientationHorizontal [toggleButton] < PlaceholderOrientation
-   PlaceholderOrientationVertical [toggleButton] < PlaceholderOrientation
- ReviewNewComment [button]
- PeopleNearMe [gallery]
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
- LockObject [button]
- UnlockObject [button]
- ObjectSaveAsPicture [button]
- BlackWhiteSettingMenu [menu]
- AutoAlternativeText [button]
- ObjectSetShapeDefaults [button]
- ObjectSizeAndPositionDialog [button]
- ObjectFormatDialog [button]
- PlaceholderOrientation [menu]
-   PlaceholderOrientationHorizontal [toggleButton] < PlaceholderOrientation
-   PlaceholderOrientationVertical [toggleButton] < PlaceholderOrientation
- ReviewNewComment [button]
- PeopleNearMe [gallery]
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
- LockObject [button]
- UnlockObject [button]
- InsertLinkGallery [gallery]
-   None [button] < InsertLinkGallery
- HyperlinkOpenInMenu [menu]
-   HyperlinkOpenInWeb [button] < HyperlinkOpenInMenu
-   HyperlinkOpenInNative [button] < HyperlinkOpenInMenu
- HyperlinkOpenDefaultSettingMenu [menu]
-   OpenDefaultWeb [button] < HyperlinkOpenDefaultSettingMenu
-   OpenDefaultNative [button] < HyperlinkOpenDefaultSettingMenu
- HyperlinkInsertPowerPoint [button]
- BookshelfDefine [button]
- HyperlinkEdit [button]
- HyperlinkOpen [button]
- HyperlinkCopy [button]
- HyperlinkRemove [button]
- ObjectSaveAsPicture [button]
- BlackWhiteSettingMenu [menu]
- AutoAlternativeText [button]
- ObjectSetShapeDefaults [button]
- ObjectSizeAndPositionDialog [button]
- ObjectFormatDialog [button]
- ReviewNewComment [button]
- PeopleNearMe [gallery]
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
- CritiqueViewMenu [gallery]
- FirstCorrectionAlternative [button]
- SecondCorrectionAlternative [button]
- ThirdCorrectionAlternative [button]
- FourthCorrectionAlternative [button]
- FifthCorrectionAlternative [button]
- MoreCorrectionAlternatives [button]
- None [button]
- Cut [button]
- Copy [button]
- PasteGalleryMini [gallery]
- EquationOptionsGallery [gallery]
- CopilotInsertPowerpoint [button]
- TextEditModeExit [button]
- FontDialog [button]
- ParagraphDialog [button]
- BulletsGallery [gallery]
-   BulletsAndNumberingBulletsDialog [button] < BulletsGallery
- NumberingGallery [gallery]
-   BulletsAndNumberingNumberingDialog [button] < NumberingGallery
- ConvertToSmartArt [gallery]
-   ConvertToSmartArtMoreSmartArtGraphicsDialog [button] < ConvertToSmartArt
- ImeReconvert [button]
- LockObject [button]
- UnlockObject [button]
- InsertLinkGallery [gallery]
-   None [button] < InsertLinkGallery
- HyperlinkOpenInMenu [menu]
-   HyperlinkOpenInWeb [button] < HyperlinkOpenInMenu
-   HyperlinkOpenInNative [button] < HyperlinkOpenInMenu
- HyperlinkOpenDefaultSettingMenu [menu]
-   OpenDefaultWeb [button] < HyperlinkOpenDefaultSettingMenu
-   OpenDefaultNative [button] < HyperlinkOpenDefaultSettingMenu
- HyperlinkInsertPowerPoint [button]
- BookshelfDefine [button]
- HyperlinkEdit [button]
- HyperlinkOpen [button]
- HyperlinkCopy [button]
- HyperlinkRemove [button]
- Insights [button]
- PasteAsSlide [button]
- CollectorFromSelection [control]
- ReuseSlidesFromSelection [control]
- FixSynonymChangeMenu [menu]
- Translate [button]
- Translator [button]
- WordArtFormatDialog [button]
- ObjectFormatDialog [button]
- AdditionalActions [gallery]
-   SmartTagOptions [button] < AdditionalActions
- ReviewNewComment [button]
- PeopleNearMe [gallery]
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
- LockObject [button]
- UnlockObject [button]
- InsertLinkGallery [gallery]
-   None [button] < InsertLinkGallery
- HyperlinkOpenInMenu [menu]
-   HyperlinkOpenInWeb [button] < HyperlinkOpenInMenu
-   HyperlinkOpenInNative [button] < HyperlinkOpenInMenu
- HyperlinkOpenDefaultSettingMenu [menu]
-   OpenDefaultWeb [button] < HyperlinkOpenDefaultSettingMenu
-   OpenDefaultNative [button] < HyperlinkOpenDefaultSettingMenu
- HyperlinkInsertPowerPoint [button]
- BookshelfDefine [button]
- HyperlinkEdit [button]
- HyperlinkOpen [button]
- HyperlinkCopy [button]
- HyperlinkRemove [button]
- ObjectSaveAsPicture [button]
- BlackWhiteSettingMenu [menu]
- AutoAlternativeText [button]
- ObjectSetShapeDefaults [button]
- ObjectSizeAndPositionDialog [button]
- ObjectFormatDialog [button]
- ReviewNewComment [button]
- PeopleNearMe [gallery]
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
- LockObject [button]
- UnlockObject [button]
- InsertLinkGallery [gallery]
-   None [button] < InsertLinkGallery
- HyperlinkOpenInMenu [menu]
-   HyperlinkOpenInWeb [button] < HyperlinkOpenInMenu
-   HyperlinkOpenInNative [button] < HyperlinkOpenInMenu
- HyperlinkOpenDefaultSettingMenu [menu]
-   OpenDefaultWeb [button] < HyperlinkOpenDefaultSettingMenu
-   OpenDefaultNative [button] < HyperlinkOpenDefaultSettingMenu
- HyperlinkInsertPowerPoint [button]
- BookshelfDefine [button]
- HyperlinkEdit [button]
- HyperlinkOpen [button]
- HyperlinkCopy [button]
- HyperlinkRemove [button]
- ObjectSaveAsPicture [button]
- BlackWhiteSettingMenu [menu]
- AutoAlternativeText [button]
- ObjectSetShapeDefaults [button]
- ObjectSizeAndPositionDialog [button]
- ObjectFormatDialog [button]
- PictureAttributionNavigateToSource [button]
- ReviewNewComment [button]
- PeopleNearMe [gallery]
### ContextMenuChartArea
- Cut [button]
- Copy [button]
- PasteGalleryMini [gallery]
- ChartResetToMatchStyle [button]
- FontDialog [button]
- ChartChangeType [button]
- ChartSaveTemplates [button]
- ChartEditDataMenu [splitButton]
-   ChartShowDataGrid [button] < ChartEditDataMenu
-   ChartShowDataContextualMenu [button] < ChartEditDataMenu
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
- LockObject [button]
- UnlockObject [button]
- ObjectSaveAsPicture [button]
- AutoAlternativeText [button]
- ChartFormatChartArea [button]
- PictureAttributionNavigateToSource [button]
- ReviewNewComment [button]
- PeopleNearMe [gallery]
### ContextMenuChartAxis
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
- ReviewNewComment [button]
- None [button]
- PeopleNearMe [gallery]
### ContextMenuChartAxisTitle
- FirstCorrectionAlternative [button]
- SecondCorrectionAlternative [button]
- ThirdCorrectionAlternative [button]
- FourthCorrectionAlternative [button]
- FifthCorrectionAlternative [button]
- MoreCorrectionAlternatives [button]
- None [button]
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
- ReviewNewComment [button]
- None [button]
- PeopleNearMe [gallery]
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
- ReviewNewComment [button]
- None [button]
- PeopleNearMe [gallery]
### ContextMenuChartTitle
- FirstCorrectionAlternative [button]
- SecondCorrectionAlternative [button]
- ThirdCorrectionAlternative [button]
- FourthCorrectionAlternative [button]
- FifthCorrectionAlternative [button]
- MoreCorrectionAlternatives [button]
- None [button]
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
- ReviewNewComment [button]
- None [button]
- PeopleNearMe [gallery]
### ContextMenuChartDataLabel
- FirstCorrectionAlternative [button]
- SecondCorrectionAlternative [button]
- ThirdCorrectionAlternative [button]
- FourthCorrectionAlternative [button]
- FifthCorrectionAlternative [button]
- MoreCorrectionAlternatives [button]
- None [button]
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
- ReviewNewComment [button]
- None [button]
- PeopleNearMe [gallery]
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
- ReviewNewComment [button]
- None [button]
- PeopleNearMe [gallery]
### ContextMenuChartDataPoint
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
- ReviewNewComment [button]
- None [button]
- PeopleNearMe [gallery]
### ContextMenuChartDataSeries
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
- ReviewNewComment [button]
- None [button]
- PeopleNearMe [gallery]
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
- ReviewNewComment [button]
- None [button]
- PeopleNearMe [gallery]
### ContextMenuChartDisplayUnit
- FirstCorrectionAlternative [button]
- SecondCorrectionAlternative [button]
- ThirdCorrectionAlternative [button]
- FourthCorrectionAlternative [button]
- FifthCorrectionAlternative [button]
- MoreCorrectionAlternatives [button]
- None [button]
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
- ReviewNewComment [button]
- None [button]
- PeopleNearMe [gallery]
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
- ReviewNewComment [button]
- None [button]
- PeopleNearMe [gallery]
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
- ReviewNewComment [button]
- None [button]
- PeopleNearMe [gallery]
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
- ReviewNewComment [button]
- None [button]
- PeopleNearMe [gallery]
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
- ReviewNewComment [button]
- None [button]
- PeopleNearMe [gallery]
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
- ReviewNewComment [button]
- None [button]
- PeopleNearMe [gallery]
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
- ReviewNewComment [button]
- None [button]
- PeopleNearMe [gallery]
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
- ReviewNewComment [button]
- None [button]
- PeopleNearMe [gallery]
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
- ReviewNewComment [button]
- None [button]
- PeopleNearMe [gallery]
### ContextMenuChartLegendEntry
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
- ReviewNewComment [button]
- None [button]
- PeopleNearMe [gallery]
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
- ReviewNewComment [button]
- None [button]
- PeopleNearMe [gallery]
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
- ReviewNewComment [button]
- None [button]
- PeopleNearMe [gallery]
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
- ReviewNewComment [button]
- None [button]
- PeopleNearMe [gallery]
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
- ReviewNewComment [button]
- None [button]
- PeopleNearMe [gallery]
### ContextMenuChartTrendlineLabel
- FirstCorrectionAlternative [button]
- SecondCorrectionAlternative [button]
- ThirdCorrectionAlternative [button]
- FourthCorrectionAlternative [button]
- FifthCorrectionAlternative [button]
- MoreCorrectionAlternatives [button]
- None [button]
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
- ReviewNewComment [button]
- None [button]
- PeopleNearMe [gallery]
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
- ReviewNewComment [button]
- None [button]
- PeopleNearMe [gallery]
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
- ReviewNewComment [button]
- None [button]
- PeopleNearMe [gallery]
### ContextMenuNotesEditText
- CritiqueViewMenu [gallery]
- FirstCorrectionAlternative [button]
- SecondCorrectionAlternative [button]
- ThirdCorrectionAlternative [button]
- FourthCorrectionAlternative [button]
- FifthCorrectionAlternative [button]
- MoreCorrectionAlternatives [button]
- None [button]
- Cut [button]
- Copy [button]
- PasteGalleryMini [gallery]
- EquationOptionsGallery [gallery]
- FontDialog [button]
- ParagraphDialog [button]
- ImeReconvert [button]
- Insights [button]
- PasteAsSlide [button]
- CollectorFromSelection [control]
- ReuseSlidesFromSelection [control]
- FixSynonymChangeMenu [menu]
- Translate [button]
- Translator [button]
- ReviewNewComment [button]
### ContextMenuTextEditOutline
- CritiqueViewMenu [gallery]
- FirstCorrectionAlternative [button]
- SecondCorrectionAlternative [button]
- ThirdCorrectionAlternative [button]
- FourthCorrectionAlternative [button]
- FifthCorrectionAlternative [button]
- MoreCorrectionAlternatives [button]
- None [button]
- Cut [button]
- Copy [button]
- PasteGalleryMini [gallery]
- EquationOptionsGallery [gallery]
- OutlineCollapseMenu [splitButton]
-   OutlineCollapse [button] < OutlineCollapseMenu
-   CollapseAll [button] < OutlineCollapseMenu
- OutlineExpandMenu [splitButton]
-   OutlineExpand [button] < OutlineExpandMenu
-   OutlineExpandAll [button] < OutlineExpandMenu
- ImeReconvert [button]
- SlideNew [button]
- DeleteSlideContextual [button]
- OutlinePromote [button]
- OutlineDemote [button]
- OutlineMoveUp [button]
- OutlineMoveDown [button]
- InsertLinkGallery [gallery]
-   None [button] < InsertLinkGallery
- HyperlinkOpenInMenu [menu]
-   HyperlinkOpenInWeb [button] < HyperlinkOpenInMenu
-   HyperlinkOpenInNative [button] < HyperlinkOpenInMenu
- HyperlinkOpenDefaultSettingMenu [menu]
-   OpenDefaultWeb [button] < HyperlinkOpenDefaultSettingMenu
-   OpenDefaultNative [button] < HyperlinkOpenDefaultSettingMenu
- HyperlinkInsertPowerPoint [button]
- BookshelfDefine [button]
- HyperlinkEdit [button]
- HyperlinkOpen [button]
- HyperlinkCopy [button]
- HyperlinkRemove [button]
- Insights [button]
- PasteAsSlide [button]
- CollectorFromSelection [control]
- ReuseSlidesFromSelection [control]
- OutlineShowTextFormatting [toggleButton]
- FixSynonymChangeMenu [menu]
- Translate [button]
- Translator [button]
- AdditionalActions [gallery]
-   SmartTagOptions [button] < AdditionalActions
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
- LockObject [button]
- UnlockObject [button]
- ObjectSaveAsPicture [button]
- ReviewNewComment [button]
- PeopleNearMe [gallery]
### ContextMenuGraphicOleClassic
- Cut [button]
- Copy [button]
- PasteGalleryMini [gallery]
- EditOleObjectMenu [menu]
- LinksUpdate [button]
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
- LockObject [button]
- UnlockObject [button]
- InsertLinkGallery [gallery]
-   None [button] < InsertLinkGallery
- HyperlinkOpenInMenu [menu]
-   HyperlinkOpenInWeb [button] < HyperlinkOpenInMenu
-   HyperlinkOpenInNative [button] < HyperlinkOpenInMenu
- HyperlinkOpenDefaultSettingMenu [menu]
-   OpenDefaultWeb [button] < HyperlinkOpenDefaultSettingMenu
-   OpenDefaultNative [button] < HyperlinkOpenDefaultSettingMenu
- HyperlinkInsertPowerPoint [button]
- BookshelfDefine [button]
- HyperlinkEdit [button]
- HyperlinkOpen [button]
- HyperlinkCopy [button]
- HyperlinkRemove [button]
- ObjectSaveAsPicture [button]
- BlackWhiteSettingMenu [menu]
- AutoAlternativeText [button]
- ObjectFormatDialog [button]
- ReviewNewComment [button]
- PeopleNearMe [gallery]
### ContextMenuActiveXControl
- Cut [button]
- Copy [button]
- PasteGalleryMini [gallery]
- ObjectSaveAsPicture [button]
- PropertySheet [button]
- ViewVisualBasicCode [button]
- EditOleObjectMenu [menu]
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
- LockObject [button]
- UnlockObject [button]
- AutoAlternativeText [button]
- ActionInsert [button]
- ObjectSizeAndPositionDialog [button]
- ReviewNewComment [button]
- PeopleNearMe [gallery]
### ContextMenuTable
- CritiqueViewMenu [gallery]
- FirstCorrectionAlternative [button]
- SecondCorrectionAlternative [button]
- ThirdCorrectionAlternative [button]
- FourthCorrectionAlternative [button]
- FifthCorrectionAlternative [button]
- MoreCorrectionAlternatives [button]
- None [button]
- Cut [button]
- Copy [button]
- PasteGalleryMini [gallery]
- FontDialog [button]
- BulletsGallery [gallery]
-   BulletsAndNumberingBulletsDialog [button] < BulletsGallery
- NumberingGallery [gallery]
-   BulletsAndNumberingNumberingDialog [button] < NumberingGallery
- LockObject [button]
- UnlockObject [button]
- ImeReconvert [button]
- InsertLinkGallery [gallery]
-   None [button] < InsertLinkGallery
- HyperlinkOpenInMenu [menu]
-   HyperlinkOpenInWeb [button] < HyperlinkOpenInMenu
-   HyperlinkOpenInNative [button] < HyperlinkOpenInMenu
- HyperlinkOpenDefaultSettingMenu [menu]
-   OpenDefaultWeb [button] < HyperlinkOpenDefaultSettingMenu
-   OpenDefaultNative [button] < HyperlinkOpenDefaultSettingMenu
- HyperlinkInsertPowerPoint [button]
- BookshelfDefine [button]
- HyperlinkEdit [button]
- HyperlinkOpen [button]
- HyperlinkCopy [button]
- HyperlinkRemove [button]
- Insights [button]
- PasteAsSlide [button]
- CollectorFromSelection [control]
- ReuseSlidesFromSelection [control]
- FixSynonymChangeMenu [menu]
- Translate [button]
- Translator [button]
- MergeCells [button]
- SplitCells [button]
- TableSelect [button]
- ObjectFormatDialog [button]
- PictureAttributionNavigateToSource [button]
- AdditionalActions [gallery]
-   SmartTagOptions [button] < AdditionalActions
- ReviewNewComment [button]
- PeopleNearMe [gallery]
### ContextMenuTableWhole
- Cut [button]
- Copy [button]
- PasteGalleryMini [gallery]
- FontDialog [button]
- BulletsGallery [gallery]
-   BulletsAndNumberingBulletsDialog [button] < BulletsGallery
- NumberingGallery [gallery]
-   BulletsAndNumberingNumberingDialog [button] < NumberingGallery
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
- LockObject [button]
- UnlockObject [button]
- InsertLinkGallery [gallery]
-   None [button] < InsertLinkGallery
- HyperlinkOpenInMenu [menu]
-   HyperlinkOpenInWeb [button] < HyperlinkOpenInMenu
-   HyperlinkOpenInNative [button] < HyperlinkOpenInMenu
- HyperlinkOpenDefaultSettingMenu [menu]
-   OpenDefaultWeb [button] < HyperlinkOpenDefaultSettingMenu
-   OpenDefaultNative [button] < HyperlinkOpenDefaultSettingMenu
- HyperlinkInsertPowerPoint [button]
- BookshelfDefine [button]
- HyperlinkEdit [button]
- HyperlinkOpen [button]
- HyperlinkCopy [button]
- HyperlinkRemove [button]
- AutoAlternativeText [button]
- ObjectSaveAsPicture [button]
- Translator [button]
- ObjectFormatDialog [button]
- ReviewNewComment [button]
- PeopleNearMe [gallery]
### ContextMenuSmartArtContentPane
- CritiqueViewMenu [gallery]
- FirstCorrectionAlternative [button]
- SecondCorrectionAlternative [button]
- ThirdCorrectionAlternative [button]
- FourthCorrectionAlternative [button]
- FifthCorrectionAlternative [button]
- MoreCorrectionAlternatives [button]
- None [button]
- Cut [button]
- Copy [button]
- PasteGalleryMini [gallery]
- SmartArtPromote [button]
- SmartArtDemote [button]
- SmartArtReorderUp [button]
- SmartArtReorderDown [button]
- FontDialog [button]
- ImeReconvert [button]
- PeopleNearMe [gallery]
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
- LockObject [button]
- UnlockObject [button]
- InsertLinkGallery [gallery]
-   None [button] < InsertLinkGallery
- HyperlinkOpenInMenu [menu]
-   HyperlinkOpenInWeb [button] < HyperlinkOpenInMenu
-   HyperlinkOpenInNative [button] < HyperlinkOpenInMenu
- HyperlinkOpenDefaultSettingMenu [menu]
-   OpenDefaultWeb [button] < HyperlinkOpenDefaultSettingMenu
-   OpenDefaultNative [button] < HyperlinkOpenDefaultSettingMenu
- HyperlinkInsertPowerPoint [button]
- HyperlinkEdit [button]
- HyperlinkOpen [button]
- HyperlinkCopy [button]
- HyperlinkRemove [button]
- ObjectSaveAsPicture [button]
- Translator [button]
- AutoAlternativeText [button]
- SmartArtResetGraphic [button]
- SmartArtConvedrtToText [button]
- SmartArtConvertToShapes [button]
- ObjectSizeAndPositionDialog [button]
- ObjectFormatDialog [button]
- PictureAttributionNavigateToSource [button]
- ReviewNewComment [button]
- PeopleNearMe [gallery]
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
- HyperlinkInsertPowerPoint [button]
- HyperlinkEdit [button]
- HyperlinkOpen [button]
- HyperlinkCopy [button]
- HyperlinkRemove [button]
- ObjectSaveAsPicture [button]
- Translator [button]
- SmartArtResetShape [button]
- ObjectSizeAndPositionDialog [button]
- ObjectFormatDialog [button]
- PictureAttributionNavigateToSource [button]
- ReviewNewComment [button]
- PeopleNearMe [gallery]
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
- ObjectSizeAndPositionDialog [button]
- ObjectFormatDialog [button]
- ReviewNewComment [button]
- PeopleNearMe [gallery]
### ContextMenuSmartArtEditText
- CritiqueViewMenu [gallery]
- FirstCorrectionAlternative [button]
- SecondCorrectionAlternative [button]
- ThirdCorrectionAlternative [button]
- FourthCorrectionAlternative [button]
- FifthCorrectionAlternative [button]
- MoreCorrectionAlternatives [button]
- None [button]
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
- HyperlinkInsertPowerPoint [button]
- HyperlinkEdit [button]
- HyperlinkOpen [button]
- HyperlinkCopy [button]
- HyperlinkRemove [button]
- Insights [button]
- PasteAsSlide [button]
- CollectorFromSelection [control]
- ReuseSlidesFromSelection [control]
- FixSynonymChangeMenu [menu]
- Translate [button]
- Translator [button]
- SmartArtResetShape [button]
- WordArtFormatDialog [button]
- ObjectSizeAndPositionDialog [button]
- ObjectFormatDialog [button]
- AdditionalActions [gallery]
-   SmartTagOptions [button] < AdditionalActions
- ReviewNewComment [button]
- PeopleNearMe [gallery]
### ContextMenuSmartArtEdit1DShapeText
- CritiqueViewMenu [gallery]
- FirstCorrectionAlternative [button]
- SecondCorrectionAlternative [button]
- ThirdCorrectionAlternative [button]
- FourthCorrectionAlternative [button]
- FifthCorrectionAlternative [button]
- MoreCorrectionAlternatives [button]
- None [button]
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
- InsertLinkGallery [gallery]
-   None [button] < InsertLinkGallery
- HyperlinkOpenInMenu [menu]
-   HyperlinkOpenInWeb [button] < HyperlinkOpenInMenu
-   HyperlinkOpenInNative [button] < HyperlinkOpenInMenu
- HyperlinkOpenDefaultSettingMenu [menu]
-   OpenDefaultWeb [button] < HyperlinkOpenDefaultSettingMenu
-   OpenDefaultNative [button] < HyperlinkOpenDefaultSettingMenu
- HyperlinkInsertPowerPoint [button]
- HyperlinkEdit [button]
- HyperlinkOpen [button]
- HyperlinkCopy [button]
- HyperlinkRemove [button]
- Insights [button]
- PasteAsSlide [button]
- CollectorFromSelection [control]
- ReuseSlidesFromSelection [control]
- FixSynonymChangeMenu [menu]
- Translate [button]
- Translator [button]
- SmartArtResetShape [button]
- WordArtFormatDialog [button]
- ObjectSizeAndPositionDialog [button]
- ObjectFormatDialog [button]
- AdditionalActions [gallery]
-   SmartTagOptions [button] < AdditionalActions
- ReviewNewComment [button]
- PeopleNearMe [gallery]
### ContextMenuGuide
- AddVerticalGuide [button]
- AddHorizontalGuide [button]
- GuideColorGallery [gallery]
- GuideDelete [button]
- GridSettings [button]
- ViewRulerPowerPoint [checkBox]
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
- LockObject [button]
- UnlockObject [button]
- ObjectSetShapeDefaults [button]
- ObjectSizeAndPropertiesDialog [button]
- ObjectFormatDialog [button]
- ReviewNewComment [button]
- PeopleNearMe [gallery]
### CtxPodiumNotesModeSelection
- PodiumShowTeleprompterNotesView [toggleButton]
- PodiumShowSidePaneNotesView [toggleButton]
- PodiumHideNotesView [toggleButton]
### ContextMenuPresenterViewNotesEditText
- CritiqueViewMenu [gallery]
- FirstCorrectionAlternative [button]
- SecondCorrectionAlternative [button]
- ThirdCorrectionAlternative [button]
- FourthCorrectionAlternative [button]
- FifthCorrectionAlternative [button]
- MoreCorrectionAlternatives [button]
- None [button]
- Cut [button]
- Copy [button]
- PasteGalleryMini [gallery]
- FontDialog [button]
- ParagraphDialog [button]
- ImeReconvert [button]
- FixSynonymChangeMenu [menu]
