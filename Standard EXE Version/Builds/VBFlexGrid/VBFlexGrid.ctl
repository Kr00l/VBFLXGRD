VERSION 5.00
Begin VB.UserControl VBFlexGrid 
   Alignable       =   -1  'True
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   DataBindingBehavior=   2  'vbComplexBound
   DrawStyle       =   5  'Transparent
   HasDC           =   0   'False
   PropertyPages   =   "VBFlexGrid.ctx":0000
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   160
   ToolboxBitmap   =   "VBFlexGrid.ctx":005F
End
Attribute VB_Name = "VBFlexGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

#Const ImplementThemedComboButton = True
#Const ImplementDataSource = True ' True = Required: msdatsrc.tlb
#Const ImplementFlexDataSource = True ' True = Required: IVBFlexDataSource.cls
#Const ImplementPreTranslateMsg = (VBFLXGRD_OCX <> 0)

#If False Then
Private FlexOLEDropModeNone, FlexOLEDropModeManual
Private FlexMousePointerDefault, FlexMousePointerArrow, FlexMousePointerCrosshair, FlexMousePointerIbeam, FlexMousePointerHand, FlexMousePointerSizePointer, FlexMousePointerSizeNESW, FlexMousePointerSizeNS, FlexMousePointerSizeNWSE, FlexMousePointerSizeWE, FlexMousePointerUpArrow, FlexMousePointerHourglass, FlexMousePointerNoDrop, FlexMousePointerArrowHourglass, FlexMousePointerArrowQuestion, FlexMousePointerSizeAll, FlexMousePointerArrowCD, FlexMousePointerCustom
Private FlexRightToLeftModeNoControl, FlexRightToLeftModeVBAME, FlexRightToLeftModeSystemLocale, FlexRightToLeftModeUserLocale, FlexRightToLeftModeOSLanguage
Private FlexBorderStyleNone, FlexBorderStyleSingle, FlexBorderStyleThin, FlexBorderStyleSunken, FlexBorderStyleRaised
Private FlexAllowUserResizingNone, FlexAllowUserResizingColumns, FlexAllowUserResizingRows, FlexAllowUserResizingBoth
Private FlexSelectionModeFree, FlexSelectionModeByRow, FlexSelectionModeByColumn, FlexSelectionModeFreeByRow, FlexSelectionModeFreeByColumn
Private FlexFillStyleSingle, FlexFillStyleRepeat
Private FlexHighLightNever, FlexHighLightAlways, FlexHighLightWithFocus
Private FlexFocusRectNone, FlexFocusRectLight, FlexFocusRectHeavy
Private FlexGridLineNone, FlexGridLineFlat, FlexGridLineInset, FlexGridLineRaised, FlexGridLineDashes, FlexGridLineDots
Private FlexTextStyleFlat, FlexTextStyleRaised, FlexTextStyleInset, FlexTextStyleRaisedLight, FlexTextStyleInsetLight
Private FlexHitResultNoWhere, FlexHitResultCell, FlexHitResultDividerRowTop, FlexHitResultDividerRowBottom, FlexHitResultDividerColumnLeft, FlexHitResultDividerColumnRight
Private FlexAlignmentLeftTop, FlexAlignmentLeftCenter, FlexAlignmentLeftBottom, FlexAlignmentCenterTop, FlexAlignmentCenterCenter, FlexAlignmentCenterBottom, FlexAlignmentRightTop, FlexAlignmentRightCenter, FlexAlignmentRightBottom, FlexAlignmentGeneral
Private FlexPictureAlignmentLeftTop, FlexPictureAlignmentLeftCenter, FlexPictureAlignmentLeftBottom, FlexPictureAlignmentCenterTop, FlexPictureAlignmentCenterCenter, FlexPictureAlignmentCenterBottom, FlexPictureAlignmentRightTop, FlexPictureAlignmentRightCenter, FlexPictureAlignmentRightBottom, FlexPictureAlignmentStretch, FlexPictureAlignmentTile, FlexPictureAlignmentLeftTopNoOverlap, FlexPictureAlignmentLeftCenterNoOverlap, FlexPictureAlignmentLeftBottomNoOverlap, FlexPictureAlignmentRightTopNoOverlap, FlexPictureAlignmentRightCenterNoOverlap, FlexPictureAlignmentRightBottomNoOverlap
Private FlexRowSizingModeIndividual, FlexRowSizingModeAll
Private FlexMergeCellsNever, FlexMergeCellsFree, FlexMergeCellsRestrictRows, FlexMergeCellsRestrictColumns, FlexMergeCellsRestrictAll, FlexMergeCellsFixedOnly
Private FlexSortNone, FlexSortGenericAscending, FlexSortGenericDescending, FlexSortNumericAscending, FlexSortNumericDescending, FlexSortStringNoCaseAscending, FlexSortStringNoCaseDescending, FlexSortStringAscending, FlexSortStringDescending, FlexSortCustom, FlexSortUseColSort, FlexSortCurrencyAscending, FlexSortCurrencyDescending, FlexSortDateAscending, FlexSortDateDescending
Private FlexVisibilityPartialOK, FlexVisibilityCompleteOnly
Private FlexPictureTypeColor, FlexPictureTypeMonochrome
Private FlexEllipsisFormatNone, FlexEllipsisFormatEnd, FlexEllipsisFormatPath, FlexEllipsisFormatWord
Private FlexClearEverywhere, FlexClearFixed, FlexClearScrollable, FlexClearMovable, FlexClearFrozen, FlexClearSelection
Private FlexClearEverything, FlexClearText, FlexClearFormatting
Private FlexTabControls, FlexTabCells, FlexTabNext
Private FlexDirectionAfterReturnNone, FlexDirectionAfterReturnUp, FlexDirectionAfterReturnDown, FlexDirectionAfterReturnLeft, FlexDirectionAfterReturnRight
Private FlexWrapNone, FlexWrapRow, FlexWrapGrid
Private FlexCellText, FlexCellClip, FlexCellTextStyle, FlexCellAlignment, FlexCellPicture, FlexCellPictureAlignment, FlexCellBackColor, FlexCellForeColor, FlexCellToolTipText, FlexCellFontName, FlexCellFontSize, FlexCellFontBold, FlexCellFontItalic, FlexCellFontStrikeThrough, FlexCellFontUnderline, FlexCellFontCharset, FlexCellLeft, FlexCellTop, FlexCellWidth, FlexCellHeight, FlexCellSort
Private FlexAutoSizeModeColWidth, FlexAutoSizeModeRowHeight
Private FlexAutoSizeScopeAll, FlexAutoSizeScopeFixed, FlexAutoSizeScopeScrollable, FlexAutoSizeScopeMovable, FlexAutoSizeScopeFrozen
Private FlexClipModeNormal, FlexClipModeExcludeHidden
Private FlexFindDirectionDown, FlexFindDirectionUp
Private FlexIMEModeNoControl, FlexIMEModeOn, FlexIMEModeOff, FlexIMEModeDisable, FlexIMEModeHiragana, FlexIMEModeKatakana, FlexIMEModeKatakanaHalf, FlexIMEModeAlphaFull, FlexIMEModeAlpha, FlexIMEModeHangulFull, FlexIMEModeHangul
Private FlexEditReasonCode, FlexEditReasonF2, FlexEditReasonSpace, FlexEditReasonKeyPress, FlexEditReasonDblClick, FlexEditReasonBackSpace
Private FlexEditCloseModeCode, FlexEditCloseModeLostFocus, FlexEditCloseModeEscape, FlexEditCloseModeReturn, FlexEditCloseModeTab, FlexEditCloseModeShiftTab, FlexEditCloseModeNavigationKey
Private FlexComboModeNone, FlexComboModeDropDown, FlexComboModeEditable, FlexComboModeButton
Private FlexComboButtonValueUnpressed, FlexComboButtonValuePressed, FlexComboButtonValueDisabled
Private FlexComboButtonDrawModeNormal, FlexComboButtonDrawModeOwnerDraw
#End If
Public Enum FlexOLEDropModeConstants
FlexOLEDropModeNone = vbOLEDropNone
FlexOLEDropModeManual = vbOLEDropManual
End Enum
Public Enum FlexMousePointerConstants
FlexMousePointerDefault = 0
FlexMousePointerArrow = 1
FlexMousePointerCrosshair = 2
FlexMousePointerIbeam = 3
FlexMousePointerHand = 4
FlexMousePointerSizePointer = 5
FlexMousePointerSizeNESW = 6
FlexMousePointerSizeNS = 7
FlexMousePointerSizeNWSE = 8
FlexMousePointerSizeWE = 9
FlexMousePointerUpArrow = 10
FlexMousePointerHourglass = 11
FlexMousePointerNoDrop = 12
FlexMousePointerArrowHourglass = 13
FlexMousePointerArrowQuestion = 14
FlexMousePointerSizeAll = 15
FlexMousePointerArrowCD = 16
FlexMousePointerCustom = 99
End Enum
Public Enum FlexRightToLeftModeConstants
FlexRightToLeftModeNoControl = 0
FlexRightToLeftModeVBAME = 1
FlexRightToLeftModeSystemLocale = 2
FlexRightToLeftModeUserLocale = 3
FlexRightToLeftModeOSLanguage = 4
End Enum
Public Enum FlexBorderStyleConstants
FlexBorderStyleNone = 0
FlexBorderStyleSingle = 1
FlexBorderStyleThin = 2
FlexBorderStyleSunken = 3
FlexBorderStyleRaised = 4
End Enum
Public Enum FlexAllowUserResizingConstants
FlexAllowUserResizingNone = 0
FlexAllowUserResizingColumns = 1
FlexAllowUserResizingRows = 2
FlexAllowUserResizingBoth = 3
End Enum
Public Enum FlexSelectionModeConstants
FlexSelectionModeFree = 0
FlexSelectionModeByRow = 1
FlexSelectionModeByColumn = 2
FlexSelectionModeFreeByRow = 3
FlexSelectionModeFreeByColumn = 4
End Enum
Public Enum FlexFillStyleConstants
FlexFillStyleSingle = 0
FlexFillStyleRepeat = 1
End Enum
Public Enum FlexHighLightConstants
FlexHighLightNever = 0
FlexHighLightAlways = 1
FlexHighLightWithFocus = 2
End Enum
Public Enum FlexFocusRectConstants
FlexFocusRectNone = 0
FlexFocusRectLight = 1
FlexFocusRectHeavy = 2
End Enum
Public Enum FlexGridLineConstants
FlexGridLineNone = 0
FlexGridLineFlat = 1
FlexGridLineInset = 2
FlexGridLineRaised = 3
FlexGridLineDashes = 4
FlexGridLineDots = 5
End Enum
Public Enum FlexTextStyleConstants
FlexTextStyleFlat = 0
FlexTextStyleRaised = 1
FlexTextStyleInset = 2
FlexTextStyleRaisedLight = 3
FlexTextStyleInsetLight = 4
End Enum
Public Enum FlexHitResultConstants
FlexHitResultNoWhere = 0
FlexHitResultCell = 1
FlexHitResultDividerRowTop = 2
FlexHitResultDividerRowBottom = 3
FlexHitResultDividerColumnLeft = 4
FlexHitResultDividerColumnRight = 5
End Enum
Public Enum FlexAlignmentConstants
FlexAlignmentLeftTop = 0
FlexAlignmentLeftCenter = 1
FlexAlignmentLeftBottom = 2
FlexAlignmentCenterTop = 3
FlexAlignmentCenterCenter = 4
FlexAlignmentCenterBottom = 5
FlexAlignmentRightTop = 6
FlexAlignmentRightCenter = 7
FlexAlignmentRightBottom = 8
FlexAlignmentGeneral = 9
End Enum
Public Enum FlexPictureAlignmentConstants
FlexPictureAlignmentLeftTop = 0
FlexPictureAlignmentLeftCenter = 1
FlexPictureAlignmentLeftBottom = 2
FlexPictureAlignmentCenterTop = 3
FlexPictureAlignmentCenterCenter = 4
FlexPictureAlignmentCenterBottom = 5
FlexPictureAlignmentRightTop = 6
FlexPictureAlignmentRightCenter = 7
FlexPictureAlignmentRightBottom = 8
FlexPictureAlignmentStretch = 9
FlexPictureAlignmentTile = 10
FlexPictureAlignmentLeftTopNoOverlap = 20
FlexPictureAlignmentLeftCenterNoOverlap = 21
FlexPictureAlignmentLeftBottomNoOverlap = 22
FlexPictureAlignmentRightTopNoOverlap = 26
FlexPictureAlignmentRightCenterNoOverlap = 27
FlexPictureAlignmentRightBottomNoOverlap = 28
End Enum
Public Enum FlexRowSizingModeConstants
FlexRowSizingModeIndividual = 0
FlexRowSizingModeAll = 1
End Enum
Public Enum FlexMergeCellsConstants
FlexMergeCellsNever = 0
FlexMergeCellsFree = 1
FlexMergeCellsRestrictRows = 2
FlexMergeCellsRestrictColumns = 3
FlexMergeCellsRestrictAll = 4
FlexMergeCellsFixedOnly = 5
End Enum
Public Enum FlexSortConstants
FlexSortNone = 0
FlexSortGenericAscending = 1
FlexSortGenericDescending = 2
FlexSortNumericAscending = 3
FlexSortNumericDescending = 4
FlexSortStringNoCaseAscending = 5
FlexSortStringNoCaseDescending = 6
FlexSortStringAscending = 7
FlexSortStringDescending = 8
FlexSortCustom = 9
FlexSortUseColSort = 10
FlexSortCurrencyAscending = 11
FlexSortCurrencyDescending = 12
FlexSortDateAscending = 13
FlexSortDateDescending = 14
End Enum
Public Enum FlexVisibilityConstants
FlexVisibilityPartialOK = 0
FlexVisibilityCompleteOnly = 1
End Enum
Public Enum FlexPictureTypeConstants
FlexPictureTypeColor = 0
FlexPictureTypeMonochrome = 1
End Enum
Public Enum FlexEllipsisFormatConstants
FlexEllipsisFormatNone = 0
FlexEllipsisFormatEnd = 1
FlexEllipsisFormatPath = 2
FlexEllipsisFormatWord = 3
End Enum
Public Enum FlexClearWhereConstants
FlexClearEverywhere = 0
FlexClearFixed = 1
FlexClearScrollable = 2
FlexClearMovable = 3
FlexClearFrozen = 4
FlexClearSelection = 5
End Enum
Public Enum FlexClearWhatConstants
FlexClearEverything = 0
FlexClearText = 1
FlexClearFormatting = 2
End Enum
Public Enum FlexTabBehaviorConstants
FlexTabControls = 0
FlexTabCells = 1
FlexTabNext = 2
End Enum
Public Enum FlexDirectionAfterReturnConstants
FlexDirectionAfterReturnNone = 0
FlexDirectionAfterReturnUp = 1
FlexDirectionAfterReturnDown = 2
FlexDirectionAfterReturnLeft = 3
FlexDirectionAfterReturnRight = 4
End Enum
Public Enum FlexWrapCellBehaviorConstants
FlexWrapNone = 0
FlexWrapRow = 1
FlexWrapGrid = 2
End Enum
Public Enum FlexCellSettings
FlexCellText = 0
FlexCellClip = 1
FlexCellTextStyle = 2
FlexCellAlignment = 3
FlexCellPicture = 4
FlexCellPictureAlignment = 5
FlexCellBackColor = 6
FlexCellForeColor = 7
FlexCellToolTipText = 8
FlexCellFontName = 13
FlexCellFontSize = 14
FlexCellFontBold = 15
FlexCellFontItalic = 16
FlexCellFontStrikeThrough = 17
FlexCellFontUnderline = 18
FlexCellFontCharset = 19
FlexCellLeft = 20
FlexCellTop = 21
FlexCellWidth = 22
FlexCellHeight = 23
FlexCellSort = 24
End Enum
Public Enum FlexAutoSizeModeConstants
FlexAutoSizeModeColWidth = 0
FlexAutoSizeModeRowHeight = 1
End Enum
Public Enum FlexAutoSizeScopeConstants
FlexAutoSizeScopeAll = 0
FlexAutoSizeScopeFixed = 1
FlexAutoSizeScopeScrollable = 2
FlexAutoSizeScopeMovable = 3
FlexAutoSizeScopeFrozen = 4
End Enum
Public Enum FlexClipModeConstants
FlexClipModeNormal = 0
FlexClipModeExcludeHidden = 1
End Enum
Public Enum FlexFindDirectionConstants
FlexFindDirectionDown = 0
FlexFindDirectionUp = 1
End Enum
Public Enum FlexIMEModeConstants
FlexIMEModeNoControl = 0
FlexIMEModeOn = 1
FlexIMEModeOff = 2
FlexIMEModeDisable = 3
FlexIMEModeHiragana = 4
FlexIMEModeKatakana = 5
FlexIMEModeKatakanaHalf = 6
FlexIMEModeAlphaFull = 7
FlexIMEModeAlpha = 8
FlexIMEModeHangulFull = 9
FlexIMEModeHangul = 10
End Enum
Public Enum FlexEditReasonConstants
FlexEditReasonCode = 0
FlexEditReasonF2 = 1
FlexEditReasonSpace = 2
FlexEditReasonKeyPress = 3
FlexEditReasonDblClick = 4
FlexEditReasonBackSpace = 5
End Enum
Public Enum FlexEditCloseModeConstants
FlexEditCloseModeCode = 0
FlexEditCloseModeLostFocus = 1
FlexEditCloseModeEscape = 2
FlexEditCloseModeReturn = 3
FlexEditCloseModeTab = 4
FlexEditCloseModeShiftTab = 5
FlexEditCloseModeNavigationKey = 6
End Enum
Public Enum FlexComboModeConstants
FlexComboModeNone = 0
FlexComboModeDropDown = 1
FlexComboModeEditable = 2
FlexComboModeButton = 3
End Enum
Public Enum FlexComboButtonValueConstants
FlexComboButtonValueUnpressed = 0
FlexComboButtonValuePressed = 1
FlexComboButtonValueDisabled = 2
End Enum
Public Enum FlexComboButtonDrawModeConstants
FlexComboButtonDrawModeNormal = 0
FlexComboButtonDrawModeOwnerDraw = 1
End Enum
Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type
Private Type POINTAPI
X As Long
Y As Long
End Type
Private Type SIZEAPI
CX As Long
CY As Long
End Type
Private Type TRACKMOUSEEVENTSTRUCT
cbSize As Long
dwFlags As Long
hWndTrack As Long
dwHoverTime As Long
End Type
Private Type TMSG
hWnd As Long
Message As Long
wParam As Long
lParam As Long
Time As Long
PT As POINTAPI
End Type
Private Type TEXTMETRIC
TMHeight As Long
TMAscent As Long
TMDescent As Long
TMInternalLeading As Long
TMExternalLeading As Long
TMAveCharWidth As Long
TMMaxCharWidth As Long
TMWeight As Long
TMOverhang As Long
TMDigitizedAspectX As Long
TMDigitizedAspectY As Long
TMFirstChar As Integer
TMLastChar As Integer
TMDefaultChar As Integer
TMBreakChar As Integer
TMItalic As Byte
TMUnderlined As Byte
TMStruckOut As Byte
TMPitchAndFamily As Byte
TMCharset As Byte
End Type
Private Type PAINTSTRUCT
hDC As Long
fErase As Long
RCPaint As RECT
fRestore As Long
fIncUpdate As Long
RGBReserved(0 To 31) As Byte
End Type
Private Type DRAWITEMSTRUCT
CtlType As Long
CtlID As Long
ItemID As Long
ItemAction As Long
ItemState As Long
hWndItem As Long
hDC As Long
RCItem As RECT
ItemData As Long
End Type
Private Type SCROLLINFO
cbSize As Long
fMask As Long
nMin As Long
nMax As Long
nPage As Long
nPos As Long
nTrackPos As Long
End Type
Private Type MONITORINFO
cbSize As Long
RCMonitor As RECT
RCWork As RECT
dwFlags As Long
End Type
Private Type TLOCALESIGNATURE
lsUsb(0 To 15) As Byte
lsCsbDefault(0 To 1) As Long
lsCsbSupported(0 To 1) As Long
End Type
Private Type TOOLINFO
cbSize As Long
uFlags As Long
hWnd As Long
uId As Long
RC As RECT
hInst As Long
lpszText As Long
lParam As Long
End Type
Private Type NMHDR
hWndFrom As Long
IDFrom As Long
Code As Long
End Type
Private Const CDDS_PREPAINT As Long = &H1
Private Type NMCUSTOMDRAW
hdr As NMHDR
dwDrawStage As Long
hDC As Long
RC As RECT
dwItemSpec As Long
uItemState As Long
lItemlParam As Long
End Type
Private Type NMTTCUSTOMDRAW
NMCD As NMCUSTOMDRAW
uDrawFlags As Long
End Type
Private Type NMTTDISPINFO
hdr As NMHDR
lpszText As Long
szText(0 To ((80 * 2) - 1)) As Byte
hInst As Long
uFlags As Long
lParam As Long
End Type
Private Const RCPM_ROW As Long = &H1, RCPM_COL As Long = &H2
Private Const RCPM_ROWSEL As Long = &H4, RCPM_COLSEL As Long = &H8
Private Const RCPM_TOPROW As Long = &H10, RCPM_LEFTCOL As Long = &H20
Private Const RCPF_CHECKTOPROW As Long = &H10, RCPF_CHECKLEFTCOL As Long = &H20
Private Const RCPF_FORCETOPROWMASK As Long = &H40, RCPF_FORCELEFTCOLMASK As Long = &H80
Private Const RCPF_SETSCROLLBARS As Long = &H100
Private Const RCPF_FORCEREDRAW As Long = &H200
Private Type TROWCOLPARAMS
Mask As Long
Flags As Long
Row As Long
Col As Long
RowSel As Long
ColSel As Long
TopRow As Long
LeftCol As Long
End Type
Private Type TINDIRECTCELLREF
InProc As Boolean
SetRCP As Boolean
RCP As TROWCOLPARAMS
End Type
Private Type TCELLRANGE
LeftCol As Long
TopRow As Long
RightCol As Long
BottomRow As Long
End Type
Private Const DIVIDER_SPACING_DIP As Long = 2
Private Type THITTESTINFO
PT As POINTAPI
HitRow As Long
HitCol As Long
HitRowDivider As Long
HitColDivider As Long
HitResult As FlexHitResultConstants
MouseRow As Long
MouseCol As Long
End Type
Private Const LBLI_VALID As Long = &H1
Private Const LBLI_UNFOLDED As Long = &H2
Private Type TLABELINFO
Flags As Long
RC As RECT
DrawFlags As Long
End Type
Private Type TDRAWINFO
SelRange As TCELLRANGE
CellTextWidthPadding As Long
CellTextHeightPadding As Long
GridLinePoints(0 To 5) As POINTAPI
End Type
Private Type TMERGEDRAWCOLINFO
RowOffset As Long
Height As Long
End Type
Private Type TMERGEDRAWROWINFO
ColOffset As Long
Width As Long
Cols() As TMERGEDRAWCOLINFO
End Type
Private Type TMERGEDRAWINFO
Row As TMERGEDRAWROWINFO
End Type
Private Const CELL_TEXT_WIDTH_PADDING_DIP As Long = 3
Private Const CELL_TEXT_HEIGHT_PADDING_DIP As Long = 1
Private Type TCELL
Text As String
TextStyle As FlexTextStyleConstants
Alignment As FlexAlignmentConstants
Picture As IPictureDisp
PictureAlignment As FlexPictureAlignmentConstants
BackColor As Long
ForeColor As Long
ToolTipText As String
FontName As String
FontSize As Single
FontBold As Boolean
FontItalic As Boolean
FontStrikeThrough As Boolean
FontUnderline As Boolean
FontCharset As Integer
PictureRenderFlag As Integer
End Type
Private Const RATIO_OF_ROWINFO_HEIGHT_TO_COLINFO_WIDTH As Long = 4
Private Const ROWINFO_HEIGHT_SPACING_DIP As Long = 3
Private Const RWIS_HIDDEN As Long = &H1
Private Const RWIS_MERGE As Long = &H2
Private Type TROWINFO
Height As Long
Data As Long
State As Long
ID As Long
End Type
Private Const COLINFO_WIDTH_SPACING_DIP As Long = 6
Private Const CLIS_HIDDEN As Long = &H1
Private Const CLIS_MERGE As Long = &H2
Private Type TCOLINFO
Width As Long
Data As Long
State As Long
Key As String
Alignment As FlexAlignmentConstants
FixedAlignment As FlexAlignmentConstants
Sort As FlexSortConstants
ComboMode As FlexComboModeConstants
ComboItems As String
Format As String
End Type
Private Type TCOLS
Cols() As TCELL
RowInfo As TROWINFO
End Type
Private Type TROWS
Rows() As TCOLS
End Type
Public Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600
Public Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Attribute DblClick.VB_UserMemId = -601
Public Event Scroll()
Attribute Scroll.VB_Description = "Occurs when you reposition the scroll box on a control."
Public Event ContextMenu(ByVal X As Single, ByVal Y As Single)
Attribute ContextMenu.VB_Description = "Occurs when the user clicked the right mouse button or types SHIFT + F10."
Public Event BeforeUserResize(ByVal Row As Long, ByVal Col As Long, ByRef Cancel As Boolean)
Attribute BeforeUserResize.VB_Description = "Occurs when the user has begun dragging a divider on a row or column."
Public Event AfterUserResize(ByVal Row As Long, ByVal Col As Long, ByRef NewSize As Long)
Attribute AfterUserResize.VB_Description = "Occurs when the user has finished dragging a divider on a row or column."
Public Event LeaveCell()
Attribute LeaveCell.VB_Description = "Occurs after the cursor leaves a cell."
Public Event EnterCell()
Attribute EnterCell.VB_Description = "Occurs before the cursor enters a cell."
Public Event BeforeRowColChange(ByVal NewRow As Long, ByVal NewCol As Long, ByRef Cancel As Boolean)
Attribute BeforeRowColChange.VB_Description = "Occurs before the current cell changes."
Public Event RowColChange()
Attribute RowColChange.VB_Description = "Occurs when the current cell changes."
Public Event BeforeSelChange(ByVal NewRowSel As Long, ByVal NewColSel As Long, ByRef Cancel As Boolean)
Attribute BeforeSelChange.VB_Description = "Occurs before the selected range of cells changes."
Public Event SelChange()
Attribute SelChange.VB_Description = "Occurs when the selected range of cells changes."
Public Event Compare(ByVal Row1 As Long, ByVal Row2 As Long, ByVal Col As Long, ByRef Cmp As Long)
Attribute Compare.VB_Description = "Occurs during custom sorts to compare two rows."
Public Event BeforeEdit(ByRef Row As Long, ByRef Col As Long, ByVal Reason As FlexEditReasonConstants, ByRef Cancel As Boolean)
Attribute BeforeEdit.VB_Description = "Occurs when a user attempts to edit the text of a cell."
Public Event AfterEdit(ByVal Row As Long, ByVal Col As Long, ByVal Changed As Boolean)
Attribute AfterEdit.VB_Description = "Occurs after a user edits the text of a cell."
Public Event LeaveEdit()
Attribute LeaveEdit.VB_Description = "Occurs when a user leaves edit mode."
Public Event EnterEdit()
Attribute EnterEdit.VB_Description = "Occurs when a user enters edit mode."
Public Event ValidateEdit(ByRef Cancel As Boolean)
Attribute ValidateEdit.VB_Description = "Occurs before any changes made by a user are committed to a cell. If the validation fails the changes will be discarded and the control will remain in edit mode."
Public Event EditSetupStyle(ByRef dwStyle As Long, ByRef dwExStyle As Long)
Attribute EditSetupStyle.VB_Description = "Occurs before the edit control is created. This is a request to perform additional customizations."
Public Event EditSetupWindow(ByRef BackColor As OLE_COLOR, ByRef ForeColor As OLE_COLOR)
Attribute EditSetupWindow.VB_Description = "Occurs after the edit control has been created and before it is displayed. This is a request to perform additional customizations."
Public Event EditQueryClose(ByVal CloseMode As FlexEditCloseModeConstants, ByRef Cancel As Boolean)
Attribute EditQueryClose.VB_Description = "Occurs whenever the edit mode is about to be closed, except when the edit control loses the focus."
Public Event EditChange()
Attribute EditChange.VB_Description = "Occurs when the contents of a control have changed."
Public Event EditContextMenu(ByRef Handled As Boolean, ByVal X As Single, ByVal Y As Single)
Attribute EditContextMenu.VB_Description = "Occurs when the user clicked the right mouse button or types SHIFT + F10."
Public Event EditKeyDown(KeyCode As Integer, Shift As Integer)
Attribute EditKeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Public Event EditKeyUp(KeyCode As Integer, Shift As Integer)
Attribute EditKeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Public Event EditKeyPress(KeyChar As Integer)
Attribute EditKeyPress.VB_Description = "Occurs when the user presses and releases an character key."
Public Event ComboDropDown()
Attribute ComboDropDown.VB_Description = "Occurs when the drop-down list is about to drop down."
Public Event ComboCloseUp()
Attribute ComboCloseUp.VB_Description = "Occurs when the drop-down list has been closed."
Public Event ComboButtonClick()
Attribute ComboButtonClick.VB_Description = "Occurs when the user clicks on a combo button. Only applicable if the combo mode property is set to button."
Public Event ComboButtonOwnerDraw(ByVal Action As Long, ByVal State As Long, ByVal hDC As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
Attribute ComboButtonOwnerDraw.VB_Description = "Occurs when a visual aspect of an owner-drawn combo button has changed."
Public Event DividerDblClick(ByVal Row As Long, ByVal Col As Long)
Attribute DividerDblClick.VB_Description = "Occurs when the user double-clicked the divider on a row or column."
Public Event CellClick(ByVal Row As Long, ByVal Col As Long, ByVal Button As Integer)
Attribute CellClick.VB_Description = "Occurs when a cell is clicked."
Public Event CellDblClick(ByVal Row As Long, ByVal Col As Long, ByVal Button As Integer)
Attribute CellDblClick.VB_Description = "Occurs when a cell is double clicked."
Public Event PreviewKeyDown(ByVal KeyCode As Integer, ByRef IsInputKey As Boolean)
Attribute PreviewKeyDown.VB_Description = "Occurs before the KeyDown event."
Public Event PreviewKeyUp(ByVal KeyCode As Integer, ByRef IsInputKey As Boolean)
Attribute PreviewKeyUp.VB_Description = "Occurs before the KeyUp event."
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Attribute KeyDown.VB_UserMemId = -602
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Attribute KeyUp.VB_UserMemId = -604
Public Event KeyPress(KeyChar As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an character key."
Attribute KeyPress.VB_UserMemId = -603
Public Event BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByRef Cancel As Boolean)
Attribute BeforeMouseDown.VB_Description = "Occurs before the control processes the MouseDown event."
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Attribute MouseDown.VB_UserMemId = -605
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Attribute MouseMove.VB_UserMemId = -606
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Attribute MouseUp.VB_UserMemId = -607
Public Event MouseEnter()
Attribute MouseEnter.VB_Description = "Occurs when the user moves the mouse into the control."
Public Event MouseLeave()
Attribute MouseLeave.VB_Description = "Occurs when the user moves the mouse out of the control."
Public Event OLECompleteDrag(Effect As Long)
Attribute OLECompleteDrag.VB_Description = "Occurs at the OLE drag/drop source control after a manual or automatic drag/drop has been completed or canceled."
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute OLEDragDrop.VB_Description = "Occurs when data is dropped onto the control via an OLE drag/drop operation, and OLEDropMode is set to manual."
Public Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Attribute OLEDragOver.VB_Description = "Occurs when the mouse is moved over the control during an OLE drag/drop operation, if its OLEDropMode property is set to manual."
Public Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Attribute OLEGiveFeedback.VB_Description = "Occurs at the source control of an OLE drag/drop operation when the mouse cursor needs to be changed."
Public Event OLESetData(Data As DataObject, DataFormat As Integer)
Attribute OLESetData.VB_Description = "Occurs at the OLE drag/drop source control when the drop target requests data that was not provided to the DataObject during the OLEDragStart event."
Public Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Attribute OLEStartDrag.VB_Description = "Occurs when an OLE drag/drop operation is initiated either manually or automatically."
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function lstrcmp Lib "kernel32" Alias "lstrcmpW" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
Private Declare Function lstrcmpi Lib "kernel32" Alias "lstrcmpiW" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageW" (ByRef lpMsg As TMSG, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function MonitorFromWindow Lib "user32" (ByVal hWnd As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetMonitorInfo Lib "user32" Alias "GetMonitorInfoW" (ByVal hMonitor As Long, ByRef lpMI As MONITORINFO) As Long
Private Declare Function LBItemFromPt Lib "comctl32" (ByVal hLB As Long, ByVal PX As Long, ByVal PY As Long, ByVal bAutoScroll As Long) As Long
Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function BeginPaint Lib "user32" (ByVal hWnd As Long, ByRef lpPaint As PAINTSTRUCT) As Long
Private Declare Function EndPaint Lib "user32" (ByVal hWnd As Long, ByRef lpPaint As PAINTSTRUCT) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function ExtSelectClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal fnMode As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As Any, ByVal bErase As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function TrackMouseEvent Lib "user32" (ByRef lpEventTrack As TRACKMOUSEEVENTSTRUCT) As Long
Private Declare Function GetKeyboardLayout Lib "user32" (ByVal dwThreadID As Long) As Long
Private Declare Function ImmIsIME Lib "imm32" (ByVal hKL As Long) As Long
Private Declare Function ImmCreateContext Lib "imm32" () As Long
Private Declare Function ImmDestroyContext Lib "imm32" (ByVal hIMC As Long) As Long
Private Declare Function ImmGetContext Lib "imm32" (ByVal hWnd As Long) As Long
Private Declare Function ImmReleaseContext Lib "imm32" (ByVal hWnd As Long, ByVal hIMC As Long) As Long
Private Declare Function ImmGetOpenStatus Lib "imm32" (ByVal hIMC As Long) As Long
Private Declare Function ImmSetOpenStatus Lib "imm32" (ByVal hIMC As Long, ByVal fOpen As Long) As Long
Private Declare Function ImmAssociateContext Lib "imm32" (ByVal hWnd As Long, ByVal hIMC As Long) As Long
Private Declare Function ImmGetConversionStatus Lib "imm32" (ByVal hIMC As Long, ByRef lpfdwConversion As Long, ByRef lpfdwSentence As Long) As Long
Private Declare Function ImmSetConversionStatus Lib "imm32" (ByVal hIMC As Long, ByVal lpfdwConversion As Long, ByVal lpfdwSentence As Long) As Long
Private Declare Function GetSystemDefaultLangID Lib "kernel32" () As Integer
Private Declare Function GetUserDefaultLangID Lib "kernel32" () As Integer
Private Declare Function GetUserDefaultUILanguage Lib "kernel32" () As Integer
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoW" (ByVal LCID As Long, ByVal LCType As Long, ByVal lpLCData As Long, ByVal cchData As Long) As Long
Private Declare Function GetMessagePos Lib "user32" () As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function MapWindowPoints Lib "user32" (ByVal hWndFrom As Long, ByVal hWndTo As Long, ByRef lppt As Any, ByVal cPoints As Long) As Long
Private Declare Function SetRect Lib "user32" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function Polyline Lib "gdi32" (ByVal hDC As Long, ByRef lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, ByRef lpRect As RECT) As Long
Private Declare Function DrawFrameControl Lib "user32" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal nCtlType As Long, ByVal nFlags As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextW" (ByVal hDC As Long, ByVal lpchText As Long, ByVal nCount As Long, ByRef lpRect As RECT, ByVal uFormat As Long) As Long
Private Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutW" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal wOptions As Long, ByRef lpRect As Any, ByVal lpString As Long, ByVal nCount As Long, ByVal lpDX As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32W" (ByVal hDC As Long, ByVal lpsz As Long, ByVal cbString As Long, ByRef lpSize As SIZEAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDCEx Lib "user32" (ByVal hWnd As Long, ByVal hRgnClip As Long, ByVal fdwOptions As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsW" (ByVal hDC As Long, ByRef lpMetrics As TEXTMETRIC) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoW" (ByVal uAction As Long, ByVal uiParam As Long, ByRef lpvParam As Long, ByVal fWinIni As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextAlign Lib "gdi32" (ByVal hDC As Long, ByVal fMode As Long) As Long
Private Declare Function SetLayout Lib "gdi32" (ByVal hDC As Long, ByVal dwLayout As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function SetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal wBar As Long, ByRef lpScrollInfo As SCROLLINFO, ByVal fRedraw As Long) As Long
Private Declare Function GetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal wBar As Long, ByRef lpScrollInfo As SCROLLINFO) As Long
Private Declare Function ChildWindowFromPoint Lib "user32" (ByVal hWndParent As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function PtInRect Lib "user32" (ByRef lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function GetCursor Lib "user32" () As Long
Private Declare Function ClipCursor Lib "user32" (ByRef lpRect As Any) As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

#If ImplementThemedComboButton = True Then

Private Enum UxThemeComboBoxParts
CP_DROPDOWNBUTTON = 1
End Enum
Private Enum UxThemeButtonParts
BP_PUSHBUTTON = 1
End Enum
Private Enum UxThemeComboBoxStates
CBXS_NORMAL = 1
CBXS_HOT = 2
CBXS_PRESSED = 3
CBXS_DISABLED = 4
End Enum
Private Enum UxThemeButtonStates
PBS_NORMAL = 1
PBS_HOT = 2
PBS_PRESSED = 3
PBS_DISABLED = 4
End Enum
Private Declare Function IsThemeBackgroundPartiallyTransparent Lib "uxtheme" (ByVal Theme As Long, iPartId As Long, iStateId As Long) As Long
Private Declare Function DrawThemeParentBackground Lib "uxtheme" (ByVal hWnd As Long, ByVal hDC As Long, ByRef pRect As RECT) As Long
Private Declare Function DrawThemeBackground Lib "uxtheme" (ByVal Theme As Long, ByVal hDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByRef pRect As RECT, ByRef pClipRect As RECT) As Long
Private Declare Function GetThemeBackgroundContentRect Lib "uxtheme" (ByVal Theme As Long, ByVal hDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByRef pBoundingRect As RECT, ByRef pContentRect As RECT) As Long
Private Declare Function OpenThemeData Lib "uxtheme" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme" (ByVal Theme As Long) As Long

#End If

Private Const ICC_STANDARD_CLASSES As Long = &H4000
Private Const ICC_TAB_CLASSES As Long = &H8
Private Const ID_EDITCHILD As Long = 100, ID_COMBOBUTTONCHILD As Long = 101
Private Const RDW_UPDATENOW As Long = &H100, RDW_INVALIDATE As Long = &H1, RDW_ERASE As Long = &H4, RDW_ALLCHILDREN As Long = &H80, RDW_FRAME As Long = &H400
Private Const SWP_FRAMECHANGED As Long = &H20
Private Const SWP_DRAWFRAME As Long = SWP_FRAMECHANGED
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOOWNERZORDER As Long = &H200
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_SHOWWINDOW As Long = &H40
Private Const HWND_DESKTOP As Long = &H0
Private Const COLOR_WINDOW As Long = 5
Private Const COLOR_HOTLIGHT As Long = 26
Private Const SYSTEM_FONT As Long = 13
Private Const MONITOR_DEFAULTTOPRIMARY As Long = &H1
Private Const DCX_WINDOW As Long = &H1
Private Const DCX_INTERSECTRGN As Long = &H80
Private Const DCX_USESTYLE As Long = &H10000
Private Const MK_SHIFT As Long = &H4
Private Const MK_CONTROL As Long = &H8
Private Const TME_LEAVE As Long = &H2
Private Const TME_NONCLIENT As Long = &H10
Private Const ODS_SELECTED As Long = &H1
Private Const ODS_DISABLED As Long = &H4
Private Const ODS_FOCUS As Long = &H10
Private Const ODS_HOTLIGHT As Long = &H40
Private Const ODS_NOFOCUSRECT As Long = &H200
Private Const PS_SOLID As Long = 0
Private Const PS_DASH As Long = 1
Private Const PS_DOT As Long = 2
Private Const SB_HORZ As Long = 0
Private Const SB_VERT As Long = 1
Private Const SB_LINELEFT As Long = 0
Private Const SB_LINEUP As Long = 0
Private Const SB_LINERIGHT As Long = 1
Private Const SB_LINEDOWN As Long = 1
Private Const SB_PAGELEFT As Long = 2
Private Const SB_PAGEUP As Long = 2
Private Const SB_PAGERIGHT As Long = 3
Private Const SB_PAGEDOWN As Long = 3
Private Const SB_THUMBPOSITION As Long = 4
Private Const SB_THUMBTRACK As Long = 5
Private Const SB_TOP As Long = 6
Private Const SB_BOTTOM As Long = 7
Private Const SM_CXVSCROLL As Long = 2
Private Const SM_CYHSCROLL As Long = 3
Private Const SM_CYVSCROLL As Long = 20
Private Const SM_CXHSCROLL As Long = 21
Private Const SM_CXBORDER As Long = 5
Private Const SM_CYBORDER As Long = 6
Private Const SM_CXEDGE As Long = 45
Private Const SM_CYEDGE As Long = 46
Private Const SIF_RANGE As Long = &H1
Private Const SIF_PAGE As Long = &H2
Private Const SIF_POS As Long = &H4
Private Const SIF_DISABLENOSCROLL As Long = &H8
Private Const SIF_TRACKPOS As Long = &H10
Private Const SIF_ALL As Long = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)
Private Const SPI_GETWHEELSCROLLLINES As Long = &H68
Private Const SPI_GETFOCUSBORDERHEIGHT As Long = &H2010
Private Const SPI_GETFOCUSBORDERWIDTH As Long = &H200E
Private Const RGN_DIFF As Long = 4
Private Const RGN_COPY As Long = 5
Private Const DT_NOPREFIX As Long = &H800
Private Const DT_RTLREADING As Long = &H20000
Private Const DT_LEFT As Long = &H0
Private Const DT_CENTER As Long = &H1
Private Const DT_RIGHT As Long = &H2
Private Const DT_VCENTER As Long = &H4
Private Const DT_BOTTOM As Long = &H8
Private Const DT_WORDBREAK As Long = &H10
Private Const DT_SINGLELINE As Long = &H20
Private Const DT_PATH_ELLIPSIS As Long = &H4000
Private Const DT_END_ELLIPSIS As Long = &H8000&
Private Const DT_WORD_ELLIPSIS As Long = &H40000
Private Const DT_CALCRECT As Long = &H400
Private Const ETO_OPAQUE As Long = 2
Private Const ETO_CLIPPED As Long = 4
Private Const TA_CENTER As Long = 6
Private Const TA_BASELINE As Long = 24
Private Const GWL_STYLE As Long = (-16)
Private Const GWL_EXSTYLE As Long = (-20)
Private Const GWL_USERDATA As Long = (-21)
Private Const LAYOUT_RTL As Long = &H1
Private Const EM_GETSEL As Long = &HB0
Private Const EM_SETSEL As Long = &HB1
Private Const EM_REPLACESEL As Long = &HC2
Private Const EM_GETLIMITTEXT As Long = &HD5
Private Const EM_LIMITTEXT As Long = &HC5
Private Const EM_SETLIMITTEXT As Long = EM_LIMITTEXT
Private Const EM_LINEINDEX As Long = &HBB
Private Const EM_LINELENGTH As Long = &HC1
Private Const EM_LINEFROMCHAR As Long = &HC9
Private Const EM_GETLINECOUNT As Long = &HBA
Private Const EM_GETMARGINS As Long = &HD4
Private Const EM_SETMARGINS As Long = &HD3
Private Const EN_CHANGE As Long = &H300
Private Const ES_LEFT As Long = &H0
Private Const ES_CENTER As Long = &H1
Private Const ES_RIGHT As Long = &H2
Private Const ES_MULTILINE As Long = &H4
Private Const ES_AUTOVSCROLL As Long = &H40
Private Const ES_AUTOHSCROLL As Long = &H80
Private Const ES_READONLY As Long = &H800
Private Const EC_LEFTMARGIN As Long = &H1
Private Const EC_RIGHTMARGIN As Long = &H2
Private Const SS_OWNERDRAW As Long = &HD
Private Const SS_NOTIFY As Long = &H100
Private Const STN_CLICKED As Long = &H0
Private Const STN_DBLCLK As Long = &H1
Private Const STN_ENABLE As Long = &H2
Private Const STN_DISABLE As Long = &H3
Private Const LB_ERR As Long = (-1)
Private Const LB_ADDSTRING As Long = &H180
Private Const LB_INSERTSTRING As Long = &H181
Private Const LB_SETCURSEL As Long = &H186
Private Const LB_GETCURSEL As Long = &H188
Private Const LB_GETTEXT As Long = &H189
Private Const LB_GETTEXTLEN As Long = &H18A
Private Const LB_GETCOUNT As Long = &H18B
Private Const LB_GETITEMHEIGHT As Long = &H1A1
Private Const LB_FINDSTRINGEXACT As Long = &H1A2
Private Const LBS_NOTIFY As Long = &H1
Private Const LBS_SORT As Long = &H2
Private Const LBN_SELCHANGE As Long = 1
Private Const DFC_SCROLL As Long = &H3, DFCS_SCROLLCOMBOBOX As Long = &H5
Private Const DFC_BUTTON As Long = &H4, DFCS_BUTTONPUSH As Long = &H10, DFCS_ADJUSTRECT As Long = &H2000
Private Const DFCS_INACTIVE As Long = &H100
Private Const DFCS_PUSHED As Long = &H200
Private Const DFCS_HOT As Long = &H1000
Private Const DFCS_FLAT As Long = &H4000
Private Const WS_BORDER As Long = &H800000
Private Const WS_DLGFRAME As Long = &H400000
Private Const WS_EX_TRANSPARENT As Long = &H20
Private Const WS_EX_CLIENTEDGE As Long = &H200
Private Const WS_EX_STATICEDGE As Long = &H20000
Private Const WS_EX_WINDOWEDGE As Long = &H100
Private Const WS_EX_NOINHERITLAYOUT As Long = &H100000
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_CHILD As Long = &H40000000
Private Const WS_CLIPCHILDREN As Long = &H2000000
Private Const WS_CLIPSIBLINGS As Long = &H4000000
Private Const WS_POPUP As Long = &H80000000
Private Const WS_EX_TOOLWINDOW As Long = &H80
Private Const WS_EX_TOPMOST As Long = &H8
Private Const WS_EX_LAYOUTRTL As Long = &H400000, WS_EX_RTLREADING As Long = &H2000, WS_EX_RIGHT As Long = &H1000, WS_EX_LEFTSCROLLBAR As Long = &H4000
Private Const WS_HSCROLL As Long = &H100000
Private Const WS_VSCROLL As Long = &H200000
Private Const WM_NOTIFY As Long = &H4E
Private Const WM_NOTIFYFORMAT As Long = &H55
Private Const WM_MOUSEACTIVATE As Long = &H21, MA_ACTIVATE As Long = &H1, MA_ACTIVATEANDEAT As Long = &H2, MA_NOACTIVATE As Long = &H3, MA_NOACTIVATEANDEAT As Long = &H4
Private Const WM_SETTINGCHANGE As Long = &H1A
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const SW_HIDE As Long = &H0
Private Const SW_SHOW As Long = &H5
Private Const SW_SHOWNA As Long = &H8
Private Const WM_SETFOCUS As Long = &H7
Private Const WM_KILLFOCUS As Long = &H8
Private Const WM_SHOWWINDOW As Long = &H18
Private Const WM_COMMAND As Long = &H111
Private Const WM_THEMECHANGED As Long = &H31A
Private Const WM_KEYDOWN As Long = &H100
Private Const WM_KEYUP As Long = &H101
Private Const WM_CHAR As Long = &H102
Private Const WM_SYSKEYDOWN As Long = &H104
Private Const WM_SYSKEYUP As Long = &H105
Private Const WM_UNICHAR As Long = &H109, UNICODE_NOCHAR As Long = &HFFFF&
Private Const WM_INPUTLANGCHANGE As Long = &H51
Private Const WM_IME_SETCONTEXT As Long = &H281
Private Const WM_IME_CHAR As Long = &H286
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_MBUTTONDOWN As Long = &H207
Private Const WM_MBUTTONUP As Long = &H208
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_RBUTTONUP As Long = &H205
Private Const WM_LBUTTONDBLCLK As Long = &H203
Private Const WM_MBUTTONDBLCLK As Long = &H209
Private Const WM_RBUTTONDBLCLK As Long = &H206
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_MOUSELEAVE As Long = &H2A3
Private Const WM_CAPTURECHANGED As Long = &H215
Private Const WM_HSCROLL As Long = &H114
Private Const WM_VSCROLL As Long = &H115
Private Const WM_CONTEXTMENU As Long = &H7B
Private Const WM_STYLECHANGED As Long = &H7D
Private Const WM_SETFONT As Long = &H30
Private Const WM_GETFONT As Long = &H31
Private Const WM_SETREDRAW As Long = &HB
Private Const WM_SIZE As Long = &H5
Private Const WM_SETCURSOR As Long = &H20
Private Const WM_CTLCOLOREDIT As Long = &H133
Private Const WM_CTLCOLORSTATIC As Long = &H138
Private Const WM_GETTEXTLENGTH As Long = &HE
Private Const WM_GETTEXT As Long = &HD
Private Const WM_SETTEXT As Long = &HC
Private Const WM_ERASEBKGND As Long = &H14
Private Const WM_PAINT As Long = &HF
Private Const WM_PRINTCLIENT As Long = &H318
Private Const WM_NCCALCSIZE As Long = &H83
Private Const WM_NCHITTEST As Long = &H84, HTCLIENT As Long = 1, HTVSCROLL As Long = 7, HTBORDER As Long = 18
Private Const WM_NCPAINT As Long = &H85
Private Const WM_NCMOUSEMOVE As Long = &HA0
Private Const WM_NCMOUSELEAVE As Long = &H2A2
Private Const WM_DRAWITEM As Long = &H2B, ODT_STATIC As Long = &H5
Private Const WM_USER As Long = &H400
Private Const TTM_ADDTOOLA As Long = (WM_USER + 4)
Private Const TTM_ADDTOOLW As Long = (WM_USER + 50)
Private Const TTM_ADDTOOL As Long = TTM_ADDTOOLW
Private Const TTM_NEWTOOLRECTA As Long = (WM_USER + 6)
Private Const TTM_NEWTOOLRECTW As Long = (WM_USER + 52)
Private Const TTM_NEWTOOLRECT As Long = TTM_NEWTOOLRECTW
Private Const TTM_SETTOOLINFOA As Long = (WM_USER + 9)
Private Const TTM_SETTOOLINFOW As Long = (WM_USER + 54)
Private Const TTM_SETTOOLINFO As Long = TTM_SETTOOLINFOW
Private Const TTM_GETTOOLCOUNT As Long = (WM_USER + 13)
Private Const TTM_ENUMTOOLSA As Long = (WM_USER + 14)
Private Const TTM_ENUMTOOLSW As Long = (WM_USER + 58)
Private Const TTM_ENUMTOOLS As Long = TTM_ENUMTOOLSW
Private Const TTM_SETMAXTIPWIDTH As Long = (WM_USER + 24)
Private Const TTM_POP As Long = (WM_USER + 28)
Private Const TTM_UPDATE As Long = (WM_USER + 29)
Private Const TTM_ADJUSTRECT As Long = (WM_USER + 31)
Private Const LPSTR_TEXTCALLBACK As Long = (-1)
Private Const H_MAX As Long = (&HFFFF + 1)
Private Const NM_FIRST As Long = H_MAX
Private Const NM_CUSTOMDRAW As Long = (NM_FIRST - 12)
Private Const TTF_SUBCLASS As Long = &H10
Private Const TTF_TRANSPARENT As Long = &H100
Private Const TTF_PARSELINKS As Long = &H1000
Private Const TTF_RTLREADING As Long = &H4
Private Const TTS_ALWAYSTIP As Long = &H1
Private Const TTS_NOPREFIX As Long = &H2
Private Const TTN_FIRST As Long = (-520)
Private Const TTN_GETDISPINFOA As Long = (TTN_FIRST - 0)
Private Const TTN_GETDISPINFOW As Long = (TTN_FIRST - 10)
Private Const TTN_GETDISPINFO As Long = TTN_GETDISPINFOW
Private Const TTN_SHOW As Long = (TTN_FIRST - 1)
Implements OLEGuids.IObjectSafety
Implements OLEGuids.IOleInPlaceActiveObjectVB
Implements OLEGuids.IOleControlVB
Private VBFlexGridHandle As Long, VBFlexGridEditHandle As Long, VBFlexGridComboButtonHandle As Long, VBFlexGridComboListHandle As Long, VBFlexGridToolTipHandle As Long
Private VBFlexGridDoubleBufferDC As Long, VBFlexGridDoubleBufferBmp As Long, VBFlexGridDoubleBufferBmpOld As Long
Private VBFlexGridFontHandle As Long, VBFlexGridFontFixedHandle As Long
Private VBFlexGridClientRect As RECT
Private VBFlexGridIMCHandle As Long
Private VBFlexGridBackColorBrush As Long
Private VBFlexGridBackColorAltBrush As Long
Private VBFlexGridBackColorBkgBrush As Long
Private VBFlexGridBackColorFixedBrush As Long
Private VBFlexGridBackColorSelBrush As Long
Private VBFlexGridGridLinePen As Long, VBFlexGridPenStyle As Long
Private VBFlexGridGridLineFixedPen As Long, VBFlexGridFixedPenStyle As Long
Private VBFlexGridGridLineWhitePen As Long, VBFlexGridGridLineBlackPen As Long
Private VBFlexGridIndirectCellRef As TINDIRECTCELLREF
Private VBFlexGridCells As TROWS, VBFlexGridCellsInit As Boolean
Private VBFlexGridColsInfo() As TCOLINFO
Private VBFlexGridDrawInfo As TDRAWINFO
Private VBFlexGridMergeDrawInfo As TMERGEDRAWINFO
Private VBFlexGridDefaultCell As TCELL
Private VBFlexGridDefaultRowInfo As TROWINFO
Private VBFlexGridDefaultColInfo As TCOLINFO
Private VBFlexGridDefaultCols As TCOLS
Private VBFlexGridDefaultRowHeight As Long
Private VBFlexGridDefaultColWidth As Long
Private VBFlexGridDefaultFixedRowHeight As Long
Private VBFlexGridDefaultFixedColWidth As Long
Private VBFlexGridRow As Long, VBFlexGridCol As Long
Private VBFlexGridRowSel As Long, VBFlexGridColSel As Long
Private VBFlexGridTopRow As Long, VBFlexGridLeftCol As Long
Private VBFlexGridCaptureRow As Long, VBFlexGridCaptureCol As Long
Private VBFlexGridCaptureHitResult As FlexHitResultConstants
Private VBFlexGridCaptureDividerRow As Long, VBFlexGridCaptureDividerCol As Long
Private VBFlexGridCaptureDividerDrag As Boolean
Private VBFlexGridToolTipRow As Long, VBFlexGridToolTipCol As Long
Private VBFlexGridMouseMoveRow As Long, VBFlexGridMouseMoveCol As Long
Private VBFlexGridMouseMoveChanged As Boolean
Private VBFlexGridDividerDragSplitterRect As RECT
Private VBFlexGridDividerDragOffset As POINTAPI
Private VBFlexGridHitRow As Long, VBFlexGridHitCol As Long
Private VBFlexGridHitRowDivider As Long, VBFlexGridHitColDivider As Long
Private VBFlexGridHitResult As FlexHitResultConstants
Private VBFlexGridCellClickRow As Long, VBFlexGridCellClickCol As Long
Private VBFlexGridEditRow As Long, VBFlexGridEditCol As Long
Private VBFlexGridEditMergedRange As TCELLRANGE
Private VBFlexGridEditReason As FlexEditReasonConstants
Private VBFlexGridEditCloseMode As FlexEditCloseModeConstants
Private VBFlexGridEditChangeFrozen As Boolean
Private VBFlexGridEditOnValidate As Boolean
Private VBFlexGridEditTextChanged As Boolean
Private VBFlexGridEditAlreadyValidated As Boolean
Private VBFlexGridEditRectChanged As Boolean
Private VBFlexGridEditRectChangedFrozen As Boolean
Private VBFlexGridEditTempFontHandle As Long
Private VBFlexGridEditBackColor As OLE_COLOR, VBFlexGridEditForeColor As OLE_COLOR
Private VBFlexGridEditBackColorBrush As Long
Private VBFlexGridComboMode As FlexComboModeConstants
Private VBFlexGridComboActiveMode As FlexComboModeConstants
Private VBFlexGridComboButtonDrawMode As FlexComboButtonDrawModeConstants
Private VBFlexGridComboItems As String
Private VBFlexGridComboListRect As RECT
Private VBFlexGridComboButtonClick As Boolean
Private VBFlexGridWheelScrollLines As Long
Private VBFlexGridFocusBorder As SIZEAPI
Private VBFlexGridFocused As Boolean
Private VBFlexGridNoRedraw As Boolean
Private VBFlexGridCharCodeCache As Long
Private VBFlexGridIsClick As Boolean
Private VBFlexGridMouseOver As Boolean
Private VBFlexGridDesignMode As Boolean
Private VBFlexGridRTLLayout As Boolean, VBFlexGridRTLReading As Boolean
Private VBFlexGridAlignable As Boolean
Private VBFlexGridEnabledVisualStyles As Boolean
Private VBFlexGridSort As FlexSortConstants

#If ImplementFlexDataSource = True Then

Private VBFlexGridFlexDataSource As IVBFlexDataSource

#End If

Private UCNoSetFocusFwd As Boolean

#If ImplementDataSource = True Then

Private PropDataSource As MSDATASRC.DataSource, PropDataMember As MSDATASRC.DataMember, PropRecordset As Object

#End If

#If ImplementPreTranslateMsg = True Then

Private Const UM_PRETRANSLATEMSG As Long = (WM_USER + 333)
Private VBFlexGridUsePreTranslateMsg As Boolean

#End If

Private WithEvents PropFont As StdFont
Attribute PropFont.VB_VarHelpID = -1
Private WithEvents PropFontFixed As StdFont
Attribute PropFontFixed.VB_VarHelpID = -1
Private PropVisualStyles As Boolean
Private PropBackColor As OLE_COLOR
Private PropBackColorAlt As OLE_COLOR
Private PropBackColorBkg As OLE_COLOR
Private PropBackColorFixed As OLE_COLOR
Private PropBackColorSel As OLE_COLOR
Private PropForeColor As OLE_COLOR
Private PropForeColorFixed As OLE_COLOR
Private PropForeColorSel As OLE_COLOR
Private PropGridColor As OLE_COLOR
Private PropGridColorFixed As OLE_COLOR
Private PropMousePointer As Integer, PropMouseIcon As IPictureDisp
Private PropMouseTrack As Boolean
Private PropRightToLeft As Boolean
Private PropRightToLeftLayout As Boolean
Private PropRightToLeftMode As FlexRightToLeftModeConstants
Private PropBorderStyle As FlexBorderStyleConstants
Private PropFixedRows As Long, PropFixedCols As Long
Private PropFrozenRows As Long, PropFrozenCols As Long
Private PropRows As Long, PropCols As Long
Private PropAllowBigSelection As Boolean
Private PropAllowSelection As Boolean
Private PropAllowUserEditing As Boolean
Private PropAllowUserResizing As FlexAllowUserResizingConstants
Private PropRowSizingMode As FlexRowSizingModeConstants
Private PropMergeCells As FlexMergeCellsConstants
Private PropFillStyle As FlexFillStyleConstants
Private PropSelectionMode As FlexSelectionModeConstants
Private PropScrollBars As VBRUN.ScrollBarConstants
Private PropScrollTrack As Boolean
Private PropDisableNoScroll As Boolean
Private PropHighLight As FlexHighLightConstants
Private PropFocusRect As FlexFocusRectConstants
Private PropRowHeightMin As Long
Private PropRowHeightMax As Long
Private PropColWidthMin As Long
Private PropColWidthMax As Long
Private PropGridLines As FlexGridLineConstants
Private PropGridLinesFixed As FlexGridLineConstants
Private PropGridLineWidth As Integer
Private PropTextStyle As FlexTextStyleConstants
Private PropTextStyleFixed As FlexTextStyleConstants
Private PropPictureType As FlexPictureTypeConstants
Private PropWordWrap As Boolean
Private PropSingleLine As Boolean
Private PropEllipsisFormat As FlexEllipsisFormatConstants
Private PropEllipsisFormatFixed As FlexEllipsisFormatConstants
Private PropRedraw As Boolean
Private PropDoubleBuffer As Boolean
Private PropTabBehavior As FlexTabBehaviorConstants
Private PropDirectionAfterReturn As FlexDirectionAfterReturnConstants
Private PropWrapCellBehavior As FlexWrapCellBehaviorConstants
Private PropShowInfoTips As Boolean
Private PropShowLabelTips As Boolean
Private PropClipSeparators As String
Private PropClipMode As FlexClipModeConstants
Private PropFormatString As String
Private PropIMEMode As FlexIMEModeConstants
Private PropWantReturn As Boolean

Private Sub IObjectSafety_GetInterfaceSafetyOptions(ByRef riid As OLEGuids.OLECLSID, ByRef pdwSupportedOptions As Long, ByRef pdwEnabledOptions As Long)
Const INTERFACESAFE_FOR_UNTRUSTED_CALLER As Long = &H1, INTERFACESAFE_FOR_UNTRUSTED_DATA As Long = &H2
pdwSupportedOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER Or INTERFACESAFE_FOR_UNTRUSTED_DATA
pdwEnabledOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER Or INTERFACESAFE_FOR_UNTRUSTED_DATA
End Sub

Private Sub IObjectSafety_SetInterfaceSafetyOptions(ByRef riid As OLEGuids.OLECLSID, ByVal dwOptionsSetMask As Long, ByVal dwEnabledOptions As Long)
End Sub

Private Sub IOleInPlaceActiveObjectVB_TranslateAccelerator(ByRef Handled As Boolean, ByRef RetVal As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal Shift As Long)
If wMsg = WM_KEYDOWN Or wMsg = WM_KEYUP Then
    Dim KeyCode As Integer, IsInputKey As Boolean
    KeyCode = wParam And &HFF&
    If wMsg = WM_KEYDOWN Then
        RaiseEvent PreviewKeyDown(KeyCode, IsInputKey)
    ElseIf wMsg = WM_KEYUP Then
        RaiseEvent PreviewKeyUp(KeyCode, IsInputKey)
    End If
    Select Case KeyCode
        Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyPageDown, vbKeyPageUp, vbKeyHome, vbKeyEnd
            SendMessage hWnd, wMsg, wParam, ByVal lParam
            Handled = True
        Case vbKeyTab
            Select Case PropTabBehavior
                Case FlexTabCells
                    IsInputKey = True
                Case FlexTabNext
                    Select Case PropWrapCellBehavior
                        Case FlexWrapNone
                            Select Case PropSelectionMode
                                Case FlexSelectionModeFree, FlexSelectionModeFreeByRow, FlexSelectionModeFreeByColumn
                                    If (Shift And vbShiftMask) = 0 Then
                                        If VBFlexGridCol < GetLastMovableCol() Then IsInputKey = True
                                    Else
                                        If VBFlexGridCol > GetFirstMovableCol() Then IsInputKey = True
                                    End If
                            End Select
                        Case FlexWrapRow
                            Select Case PropSelectionMode
                                Case FlexSelectionModeFree, FlexSelectionModeFreeByRow, FlexSelectionModeFreeByColumn
                                    If (Shift And vbShiftMask) = 0 Then
                                        If VBFlexGridRow < GetLastMovableRow() Or VBFlexGridCol < GetLastMovableCol() Then IsInputKey = True
                                    Else
                                        If VBFlexGridRow > GetFirstMovableRow() Or VBFlexGridCol > GetFirstMovableCol() Then IsInputKey = True
                                    End If
                                Case FlexSelectionModeByRow
                                    If (Shift And vbShiftMask) = 0 Then
                                        If VBFlexGridRow < GetLastMovableRow() Then IsInputKey = True
                                    Else
                                        If VBFlexGridRow > GetFirstMovableRow() Then IsInputKey = True
                                    End If
                                Case FlexSelectionModeByColumn
                                    If (Shift And vbShiftMask) = 0 Then
                                        If VBFlexGridCol < GetLastMovableCol() Then IsInputKey = True
                                    Else
                                        If VBFlexGridCol > GetFirstMovableCol() Then IsInputKey = True
                                    End If
                            End Select
                        Case FlexWrapGrid
                            IsInputKey = True
                    End Select
            End Select
            If IsInputKey = True Then
                SendMessage hWnd, wMsg, wParam, ByVal lParam
                Handled = True
            End If
        Case vbKeyReturn, vbKeyEscape
            If VBFlexGridEditHandle = 0 Then
                If IsInputKey = True Then
                    SendMessage hWnd, wMsg, wParam, ByVal lParam
                    Handled = True
                End If
            Else
                SendMessage hWnd, wMsg, wParam, ByVal lParam
                Handled = True
            End If
    End Select
End If
End Sub

Private Sub IOleControlVB_GetControlInfo(ByRef Handled As Boolean, ByRef AccelCount As Integer, ByRef AccelTable As Long, ByRef Flags As Long)
If PropWantReturn = True Then
    Flags = CTRLINFO_EATS_RETURN
    Handled = True
End If
End Sub

Private Sub IOleControlVB_OnMnemonic(ByRef Handled As Boolean, ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal Shift As Long)
End Sub

Private Sub UserControl_Initialize()
Call FlexLoadShellMod
Call FlexInitCC(ICC_STANDARD_CLASSES)
Call FlexWndRegisterClass

#If ImplementPreTranslateMsg = True Then

If SetVTableHandling(Me, VTableInterfaceInPlaceActiveObject) = False Then VBFlexGridUsePreTranslateMsg = True
Call SetVTableHandling(Me, VTableInterfaceControl)

#Else

Call SetVTableHandling(Me, VTableInterfaceInPlaceActiveObject)
Call SetVTableHandling(Me, VTableInterfaceControl)

#End If

With VBFlexGridDefaultCell
.TextStyle = -1
.Alignment = -1
.BackColor = -1
.ForeColor = -1
End With
With VBFlexGridDefaultRowInfo
.Height = -1
End With
With VBFlexGridDefaultColInfo
.Width = -1
.Alignment = FlexAlignmentGeneral
.FixedAlignment = -1
End With
VBFlexGridCaptureRow = -1
VBFlexGridCaptureCol = -1
VBFlexGridCaptureHitResult = FlexHitResultNoWhere
VBFlexGridCaptureDividerRow = -1
VBFlexGridCaptureDividerCol = -1
VBFlexGridCaptureDividerDrag = False
VBFlexGridToolTipRow = -1
VBFlexGridToolTipCol = -1
VBFlexGridMouseMoveRow = -1
VBFlexGridMouseMoveCol = -1
VBFlexGridMouseMoveChanged = False
VBFlexGridHitRow = -1
VBFlexGridHitCol = -1
VBFlexGridHitRowDivider = -1
VBFlexGridHitColDivider = -1
VBFlexGridHitResult = FlexHitResultNoWhere
VBFlexGridCellClickRow = -1
VBFlexGridCellClickCol = -1
VBFlexGridEditRow = -1
VBFlexGridEditCol = -1
VBFlexGridEditReason = -1
VBFlexGridEditCloseMode = -1
VBFlexGridComboMode = FlexComboModeNone
VBFlexGridComboActiveMode = FlexComboModeNone
VBFlexGridComboButtonDrawMode = FlexComboButtonDrawModeNormal
SystemParametersInfo SPI_GETWHEELSCROLLLINES, 0, VBFlexGridWheelScrollLines, 0
If SystemParametersInfo(SPI_GETFOCUSBORDERWIDTH, 0, VBFlexGridFocusBorder.CX, 0) = 0 Then VBFlexGridFocusBorder.CX = 1
If SystemParametersInfo(SPI_GETFOCUSBORDERHEIGHT, 0, VBFlexGridFocusBorder.CY, 0) = 0 Then VBFlexGridFocusBorder.CY = 1
End Sub

Private Sub UserControl_InitProperties()
On Error Resume Next
If UserControl.ParentControls.Count = 0 Then VBFlexGridAlignable = False Else VBFlexGridAlignable = True
VBFlexGridDesignMode = Not Ambient.UserMode
On Error GoTo 0

#If ImplementDataSource = True Then

PropDataMember = vbNullString

#End If

Set PropFont = Ambient.Font
Set PropFontFixed = Nothing
PropVisualStyles = True
PropBackColor = vbWindowBackground
PropBackColorAlt = vbWindowBackground
PropBackColorBkg = &H808080
PropBackColorFixed = vbButtonFace
PropBackColorSel = vbHighlight
PropForeColor = vbWindowText
PropForeColorFixed = vbButtonText
PropForeColorSel = vbHighlightText
PropGridColor = &HC0C0C0
PropGridColorFixed = vbBlack
PropMousePointer = 0: Set PropMouseIcon = Nothing
PropMouseTrack = False
PropRightToLeft = Ambient.RightToLeft
PropRightToLeftLayout = False
PropRightToLeftMode = FlexRightToLeftModeVBAME
If PropRightToLeft = True Then Me.RightToLeft = True
PropBorderStyle = FlexBorderStyleSunken
PropFixedRows = 1
PropFixedCols = 1
PropFrozenRows = 0
PropFrozenCols = 0
PropRows = 2
PropCols = 2
PropAllowBigSelection = True
PropAllowSelection = True
PropAllowUserEditing = False
PropAllowUserResizing = FlexAllowUserResizingNone
PropRowSizingMode = FlexRowSizingModeIndividual
PropMergeCells = FlexMergeCellsNever
PropFillStyle = FlexFillStyleSingle
PropSelectionMode = FlexSelectionModeFree
PropScrollBars = vbBoth
PropScrollTrack = False
PropDisableNoScroll = False
PropHighLight = FlexHighLightAlways
PropFocusRect = FlexFocusRectLight
PropRowHeightMin = 0
PropRowHeightMax = 0
PropColWidthMin = 0
PropColWidthMax = 0
PropGridLines = FlexGridLineFlat
PropGridLinesFixed = FlexGridLineInset
PropGridLineWidth = 1
PropTextStyle = FlexTextStyleFlat
PropTextStyleFixed = FlexTextStyleFlat
PropPictureType = FlexPictureTypeColor
PropWordWrap = False
PropSingleLine = False
PropEllipsisFormat = FlexEllipsisFormatNone
PropEllipsisFormatFixed = FlexEllipsisFormatNone
PropRedraw = True
PropDoubleBuffer = True
PropTabBehavior = FlexTabControls
PropDirectionAfterReturn = FlexDirectionAfterReturnNone
PropWrapCellBehavior = FlexWrapNone
PropShowInfoTips = False
PropShowLabelTips = False
PropClipSeparators = vbNullString
PropClipMode = FlexClipModeNormal
PropFormatString = vbNullString
PropIMEMode = FlexIMEModeNoControl
PropWantReturn = False
Call CreateVBFlexGrid
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
If UserControl.ParentControls.Count = 0 Then VBFlexGridAlignable = False Else VBFlexGridAlignable = True
VBFlexGridDesignMode = Not Ambient.UserMode
On Error GoTo 0
With PropBag

#If ImplementDataSource = True Then

PropDataMember = .ReadProperty("DataMember", vbNullString)

#End If

Set PropFont = .ReadProperty("Font", Nothing)
Set PropFontFixed = .ReadProperty("FontFixed", Nothing)
PropVisualStyles = .ReadProperty("VisualStyles", True)
PropBackColor = .ReadProperty("BackColor", vbWindowBackground)
PropBackColorAlt = .ReadProperty("BackColorAlt", vbWindowBackground)
PropBackColorBkg = .ReadProperty("BackColorBkg", &H808080)
PropBackColorFixed = .ReadProperty("BackColorFixed", vbButtonFace)
PropBackColorSel = .ReadProperty("BackColorSel", vbHighlight)
PropForeColor = .ReadProperty("ForeColor", vbWindowText)
PropForeColorFixed = .ReadProperty("ForeColorFixed", vbButtonText)
PropForeColorSel = .ReadProperty("ForeColorSel", vbHighlightText)
PropGridColor = .ReadProperty("GridColor", &HC0C0C0)
PropGridColorFixed = .ReadProperty("GridColorFixed", vbBlack)
Me.Enabled = .ReadProperty("Enabled", True)
Me.OLEDropMode = .ReadProperty("OLEDropMode", vbOLEDropNone)
PropMousePointer = .ReadProperty("MousePointer", 0)
Set PropMouseIcon = .ReadProperty("MouseIcon", Nothing)
PropMouseTrack = .ReadProperty("MouseTrack", False)
PropRightToLeft = .ReadProperty("RightToLeft", False)
PropRightToLeftLayout = .ReadProperty("RightToLeftLayout", False)
PropRightToLeftMode = .ReadProperty("RightToLeftMode", FlexRightToLeftModeVBAME)
If PropRightToLeft = True Then Me.RightToLeft = True
PropBorderStyle = .ReadProperty("BorderStyle", FlexBorderStyleSunken)
PropFixedRows = .ReadProperty("FixedRows", 1)
PropFixedCols = .ReadProperty("FixedCols", 1)
PropFrozenRows = .ReadProperty("FrozenRows", 0)
PropFrozenCols = .ReadProperty("FrozenCols", 0)
PropRows = .ReadProperty("Rows", 2)
PropCols = .ReadProperty("Cols", 2)
PropAllowBigSelection = .ReadProperty("AllowBigSelection", True)
PropAllowSelection = .ReadProperty("AllowSelection", True)
PropAllowUserEditing = .ReadProperty("AllowUserEditing", False)
PropAllowUserResizing = .ReadProperty("AllowUserResizing", FlexAllowUserResizingNone)
PropRowSizingMode = .ReadProperty("RowSizingMode", FlexRowSizingModeIndividual)
PropMergeCells = .ReadProperty("MergeCells", FlexMergeCellsNever)
PropFillStyle = .ReadProperty("FillStyle", FlexFillStyleSingle)
PropSelectionMode = .ReadProperty("SelectionMode", FlexSelectionModeFree)
PropScrollBars = .ReadProperty("ScrollBars", vbBoth)
PropScrollTrack = .ReadProperty("ScrollTrack", False)
PropDisableNoScroll = .ReadProperty("DisableNoScroll", False)
PropHighLight = .ReadProperty("HighLight", FlexHighLightAlways)
PropFocusRect = .ReadProperty("FocusRect", FlexFocusRectLight)
PropRowHeightMin = (.ReadProperty("RowHeightMin", 0) * PixelsPerDIP_Y())
PropRowHeightMax = (.ReadProperty("RowHeightMax", 0) * PixelsPerDIP_Y())
PropColWidthMin = (.ReadProperty("ColWidthMin", 0) * PixelsPerDIP_X())
PropColWidthMax = (.ReadProperty("ColWidthMax", 0) * PixelsPerDIP_X())
PropGridLines = .ReadProperty("GridLines", FlexGridLineFlat)
PropGridLinesFixed = .ReadProperty("GridLinesFixed", FlexGridLineInset)
PropGridLineWidth = .ReadProperty("GridLineWidth", 1)
PropTextStyle = .ReadProperty("TextStyle", FlexTextStyleFlat)
PropTextStyleFixed = .ReadProperty("TextStyleFixed", FlexTextStyleFlat)
PropPictureType = .ReadProperty("PictureType", FlexPictureTypeColor)
PropWordWrap = .ReadProperty("WordWrap", False)
PropSingleLine = .ReadProperty("SingleLine", False)
PropEllipsisFormat = .ReadProperty("EllipsisFormat", FlexEllipsisFormatNone)
PropEllipsisFormatFixed = .ReadProperty("EllipsisFormatFixed", FlexEllipsisFormatNone)
PropRedraw = .ReadProperty("Redraw", True)
PropDoubleBuffer = .ReadProperty("DoubleBuffer", True)
PropTabBehavior = .ReadProperty("TabBehavior", FlexTabControls)
PropDirectionAfterReturn = .ReadProperty("DirectionAfterReturn", FlexDirectionAfterReturnNone)
PropWrapCellBehavior = .ReadProperty("WrapCellBehavior", FlexWrapNone)
PropShowInfoTips = .ReadProperty("ShowInfoTips", False)
PropShowLabelTips = .ReadProperty("ShowLabelTips", False)
PropClipSeparators = VarToStr(.ReadProperty("ClipSeparators", vbNullString))
PropClipMode = .ReadProperty("ClipMode", FlexClipModeNormal)
PropFormatString = VarToStr(.ReadProperty("FormatString", vbNullString))
PropIMEMode = .ReadProperty("IMEMode", FlexIMEModeNoControl)
PropWantReturn = .ReadProperty("WantReturn", False)
End With
Call CreateVBFlexGrid
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag

#If ImplementDataSource = True Then

.WriteProperty "DataMember", PropDataMember, vbNullString

#End If

.WriteProperty "Font", IIf(OLEFontIsEqual(PropFont, Ambient.Font) = False, PropFont, Nothing), Nothing
.WriteProperty "FontFixed", PropFontFixed, Nothing
.WriteProperty "VisualStyles", PropVisualStyles, True
.WriteProperty "BackColor", PropBackColor, vbWindowBackground
.WriteProperty "BackColorAlt", PropBackColorAlt, vbWindowBackground
.WriteProperty "BackColorBkg", PropBackColorBkg, &H808080
.WriteProperty "BackColorFixed", PropBackColorFixed, vbButtonFace
.WriteProperty "BackColorSel", PropBackColorSel, vbHighlight
.WriteProperty "ForeColor", PropForeColor, vbWindowText
.WriteProperty "ForeColorFixed", PropForeColorFixed, vbButtonText
.WriteProperty "ForeColorSel", PropForeColorSel, vbHighlightText
.WriteProperty "GridColor", PropGridColor, &HC0C0C0
.WriteProperty "GridColorFixed", PropGridColorFixed, vbBlack
.WriteProperty "Enabled", Me.Enabled, True
.WriteProperty "OLEDropMode", Me.OLEDropMode, vbOLEDropNone
.WriteProperty "MousePointer", PropMousePointer, 0
.WriteProperty "MouseIcon", PropMouseIcon, Nothing
.WriteProperty "MouseTrack", PropMouseTrack, False
.WriteProperty "RightToLeft", PropRightToLeft, False
.WriteProperty "RightToLeftLayout", PropRightToLeftLayout, False
.WriteProperty "RightToLeftMode", PropRightToLeftMode, FlexRightToLeftModeVBAME
.WriteProperty "BorderStyle", PropBorderStyle, FlexBorderStyleSunken
.WriteProperty "FixedRows", PropFixedRows, 1
.WriteProperty "FixedCols", PropFixedCols, 1
.WriteProperty "FrozenRows", PropFrozenRows, 0
.WriteProperty "FrozenCols", PropFrozenCols, 0
.WriteProperty "Rows", PropRows, 2
.WriteProperty "Cols", PropCols, 2
.WriteProperty "AllowBigSelection", PropAllowBigSelection, True
.WriteProperty "AllowSelection", PropAllowSelection, True
.WriteProperty "AllowUserEditing", PropAllowUserEditing, False
.WriteProperty "AllowUserResizing", PropAllowUserResizing, FlexAllowUserResizingNone
.WriteProperty "RowSizingMode", PropRowSizingMode, FlexRowSizingModeIndividual
.WriteProperty "MergeCells", PropMergeCells, FlexMergeCellsNever
.WriteProperty "FillStyle", PropFillStyle, FlexFillStyleSingle
.WriteProperty "SelectionMode", PropSelectionMode, FlexSelectionModeFree
.WriteProperty "ScrollBars", PropScrollBars, vbBoth
.WriteProperty "ScrollTrack", PropScrollTrack, False
.WriteProperty "DisableNoScroll", PropDisableNoScroll, False
.WriteProperty "HighLight", PropHighLight, FlexHighLightAlways
.WriteProperty "FocusRect", PropFocusRect, FlexFocusRectLight
.WriteProperty "RowHeightMin", (PropRowHeightMin / PixelsPerDIP_Y()), 0
.WriteProperty "RowHeightMax", (PropRowHeightMax / PixelsPerDIP_Y()), 0
.WriteProperty "ColWidthMin", (PropColWidthMin / PixelsPerDIP_X()), 0
.WriteProperty "ColWidthMax", (PropColWidthMax / PixelsPerDIP_X()), 0
.WriteProperty "GridLines", PropGridLines, FlexGridLineFlat
.WriteProperty "GridLinesFixed", PropGridLinesFixed, FlexGridLineInset
.WriteProperty "GridLineWidth", PropGridLineWidth, 1
.WriteProperty "TextStyle", PropTextStyle, FlexTextStyleFlat
.WriteProperty "TextStyleFixed", PropTextStyleFixed, FlexTextStyleFlat
.WriteProperty "PictureType", PropPictureType, FlexPictureTypeColor
.WriteProperty "WordWrap", PropWordWrap, False
.WriteProperty "SingleLine", PropSingleLine, False
.WriteProperty "EllipsisFormat", PropEllipsisFormat, FlexEllipsisFormatNone
.WriteProperty "EllipsisFormatFixed", PropEllipsisFormatFixed, FlexEllipsisFormatNone
.WriteProperty "Redraw", PropRedraw, True
.WriteProperty "DoubleBuffer", PropDoubleBuffer, True
.WriteProperty "TabBehavior", PropTabBehavior, FlexTabControls
.WriteProperty "DirectionAfterReturn", PropDirectionAfterReturn, FlexDirectionAfterReturnNone
.WriteProperty "WrapCellBehavior", PropWrapCellBehavior, FlexWrapNone
.WriteProperty "ShowInfoTips", PropShowInfoTips, False
.WriteProperty "ShowLabelTips", PropShowLabelTips, False
.WriteProperty "ClipSeparators", StrToVar(PropClipSeparators), vbNullString
.WriteProperty "ClipMode", PropClipMode, FlexClipModeNormal
.WriteProperty "FormatString", StrToVar(PropFormatString), vbNullString
.WriteProperty "IMEMode", PropIMEMode, FlexIMEModeNoControl
.WriteProperty "WantReturn", PropWantReturn, False
End With
End Sub

Private Sub UserControl_Paint()
If VBFlexGridHandle = 0 Or VBFlexGridDesignMode = False Then Exit Sub
Dim OldLayout As Long, hRgn As Long
If PropRightToLeft = True And PropRightToLeftLayout = True Then OldLayout = SetLayout(UserControl.hDC, LAYOUT_RTL)
If PropDoubleBuffer = True Then
    If VBFlexGridDoubleBufferDC = 0 Then
        VBFlexGridDoubleBufferDC = CreateCompatibleDC(UserControl.hDC)
        If VBFlexGridDoubleBufferDC <> 0 Then
            VBFlexGridDoubleBufferBmp = CreateCompatibleBitmap(UserControl.hDC, UserControl.ScaleWidth, UserControl.ScaleHeight)
            If VBFlexGridDoubleBufferBmp <> 0 Then VBFlexGridDoubleBufferBmpOld = SelectObject(VBFlexGridDoubleBufferDC, VBFlexGridDoubleBufferBmp)
        End If
    End If
    If VBFlexGridDoubleBufferDC <> 0 And VBFlexGridDoubleBufferBmp <> 0 Then
        If VBFlexGridCellsInit = False Then
            If VBFlexGridBackColorBkgBrush <> 0 Then
                Dim Brush As Long
                Brush = SelectObject(VBFlexGridDoubleBufferDC, VBFlexGridBackColorBkgBrush)
                PatBlt VBFlexGridDoubleBufferDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, vbPatCopy
                SelectObject VBFlexGridDoubleBufferDC, Brush
            End If
            Call DrawGrid(VBFlexGridDoubleBufferDC, -1)
        Else
            Call DrawGrid(VBFlexGridDoubleBufferDC, hRgn)
            If hRgn <> 0 Then ExtSelectClipRgn UserControl.hDC, hRgn, RGN_COPY
        End If
        BitBlt UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, VBFlexGridDoubleBufferDC, 0, 0, vbSrcCopy
        If hRgn <> 0 Then
            ExtSelectClipRgn UserControl.hDC, 0, RGN_COPY
            DeleteObject hRgn
        End If
    End If
Else
    Call DrawGrid(UserControl.hDC, -1)
End If
If PropRightToLeft = True And PropRightToLeftLayout = True Then SetLayout UserControl.hDC, OldLayout
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, UserControl.ScaleX(X, vbPixels, vbContainerPosition), UserControl.ScaleY(Y, vbPixels, vbContainerPosition))
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
RaiseEvent OLEDragOver(Data, Effect, Button, Shift, UserControl.ScaleX(X, vbPixels, vbContainerPosition), UserControl.ScaleY(Y, vbPixels, vbContainerPosition), State)
End Sub

Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)
RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Public Sub OLEDrag()
Attribute OLEDrag.VB_Description = "Starts an OLE drag/drop event with the given control as the source."
UserControl.OLEDrag
End Sub

Private Sub UserControl_Resize()
Static LastHeight As Single, LastWidth As Single, LastAlign As Integer
Static InProc As Boolean
If InProc = True Then Exit Sub
InProc = True
With UserControl.Extender
Dim Align As Integer
If VBFlexGridAlignable = True Then Align = .Align Else Align = vbAlignNone
Select Case Align
    Case LastAlign
    Case vbAlignNone
    Case vbAlignTop, vbAlignBottom
        Select Case LastAlign
            Case vbAlignLeft, vbAlignRight
                .Height = LastWidth
        End Select
    Case vbAlignLeft, vbAlignRight
        Select Case LastAlign
            Case vbAlignTop, vbAlignBottom
                .Width = LastHeight
        End Select
End Select
LastHeight = .Height
LastWidth = .Width
LastAlign = Align
End With
With UserControl
If DPICorrectionFactor() <> 1 Then Call SyncObjectRectsToContainer(Me)
If VBFlexGridDesignMode = False Then
    If VBFlexGridHandle <> 0 Then MoveWindow VBFlexGridHandle, 0, 0, .ScaleWidth, .ScaleHeight, 1
Else
    If VBFlexGridDoubleBufferDC <> 0 Then
        If VBFlexGridDoubleBufferBmpOld <> 0 Then
            SelectObject VBFlexGridDoubleBufferDC, VBFlexGridDoubleBufferBmpOld
            VBFlexGridDoubleBufferBmpOld = 0
        End If
        If VBFlexGridDoubleBufferBmp <> 0 Then
            DeleteObject VBFlexGridDoubleBufferBmp
            VBFlexGridDoubleBufferBmp = 0
        End If
        DeleteDC VBFlexGridDoubleBufferDC
        VBFlexGridDoubleBufferDC = 0
    End If
    SetRect VBFlexGridClientRect, 0, 0, .ScaleWidth, .ScaleHeight
    Call SetScrollBars
End If
End With
InProc = False
End Sub

Private Sub UserControl_Terminate()

#If ImplementPreTranslateMsg = True Then

If VBFlexGridUsePreTranslateMsg = False Then Call RemoveVTableHandling(Me, VTableInterfaceInPlaceActiveObject)
Call RemoveVTableHandling(Me, VTableInterfaceControl)
Call RemoveVTableHandling(Me, VTableInterfacePerPropertyBrowsing)

#Else

Call RemoveVTableHandling(Me, VTableInterfaceInPlaceActiveObject)
Call RemoveVTableHandling(Me, VTableInterfaceControl)
Call RemoveVTableHandling(Me, VTableInterfacePerPropertyBrowsing)

#End If

Call DestroyVBFlexGrid
Call FlexWndReleaseClass
Call FlexReleaseShellMod
End Sub

Public Property Get Name() As String
Attribute Name.VB_Description = "Returns the name used in code to identify an object."
Name = Ambient.DisplayName
End Property

Public Property Get Tag() As String
Attribute Tag.VB_Description = "Stores any extra data needed for your program."
Tag = Extender.Tag
End Property

Public Property Let Tag(ByVal Value As String)
Extender.Tag = Value
End Property

Public Property Get Parent() As Object
Attribute Parent.VB_Description = "Returns the object on which this object is located."
Set Parent = UserControl.Parent
End Property

Public Property Get Container() As Object
Attribute Container.VB_Description = "Returns the container of an object."
Set Container = Extender.Container
End Property

Public Property Set Container(ByVal Value As Object)
Set Extender.Container = Value
End Property

Public Property Get Left() As Single
Attribute Left.VB_Description = "Returns/sets the distance between the internal left edge of an object and the left edge of its container."
Left = Extender.Left
End Property

Public Property Let Left(ByVal Value As Single)
Extender.Left = Value
End Property

Public Property Get Top() As Single
Attribute Top.VB_Description = "Returns/sets the distance between the internal top edge of an object and the top edge of its container."
Top = Extender.Top
End Property

Public Property Let Top(ByVal Value As Single)
Extender.Top = Value
End Property

Public Property Get Width() As Single
Attribute Width.VB_Description = "Returns/sets the width of an object."
Width = Extender.Width
End Property

Public Property Let Width(ByVal Value As Single)
Extender.Width = Value
End Property

Public Property Get Height() As Single
Attribute Height.VB_Description = "Returns/sets the height of an object."
Height = Extender.Height
End Property

Public Property Let Height(ByVal Value As Single)
Extender.Height = Value
End Property

Public Property Get Visible() As Boolean
Attribute Visible.VB_Description = "Returns/sets a value that determines whether an object is visible or hidden."
Visible = Extender.Visible
End Property

Public Property Let Visible(ByVal Value As Boolean)
Extender.Visible = Value
End Property

Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
Attribute ToolTipText.VB_MemberFlags = "400"
ToolTipText = Extender.ToolTipText
End Property

Public Property Let ToolTipText(ByVal Value As String)
Extender.ToolTipText = Value
End Property

Public Property Get HelpContextID() As Long
Attribute HelpContextID.VB_Description = "Specifies the default Help file context ID for an object."
HelpContextID = Extender.HelpContextID
End Property

Public Property Let HelpContextID(ByVal Value As Long)
Extender.HelpContextID = Value
End Property

Public Property Get WhatsThisHelpID() As Long
Attribute WhatsThisHelpID.VB_Description = "Returns/sets an associated context number for an object."
Attribute WhatsThisHelpID.VB_MemberFlags = "400"
WhatsThisHelpID = Extender.WhatsThisHelpID
End Property

Public Property Let WhatsThisHelpID(ByVal Value As Long)
Extender.WhatsThisHelpID = Value
End Property

Public Property Get Align() As Integer
Attribute Align.VB_Description = "Returns/sets a value that determines where an object is displayed on a form."
Attribute Align.VB_MemberFlags = "400"
Align = Extender.Align
End Property

Public Property Let Align(ByVal Value As Integer)
Extender.Align = Value
End Property

Public Property Get DragIcon() As IPictureDisp
Attribute DragIcon.VB_Description = "Returns/sets the icon to be displayed as the pointer in a drag-and-drop operation."
Attribute DragIcon.VB_MemberFlags = "400"
Set DragIcon = Extender.DragIcon
End Property

Public Property Let DragIcon(ByVal Value As IPictureDisp)
Extender.DragIcon = Value
End Property

Public Property Set DragIcon(ByVal Value As IPictureDisp)
Set Extender.DragIcon = Value
End Property

Public Property Get DragMode() As Integer
Attribute DragMode.VB_Description = "Returns/sets a value that determines whether manual or automatic drag mode is used."
Attribute DragMode.VB_MemberFlags = "400"
DragMode = Extender.DragMode
End Property

Public Property Let DragMode(ByVal Value As Integer)
Extender.DragMode = Value
End Property

Public Sub Drag(Optional ByRef Action As Variant)
Attribute Drag.VB_Description = "Begins, ends, or cancels a drag operation of any object except Line, Menu, Shape, and Timer."
If IsMissing(Action) Then Extender.Drag Else Extender.Drag Action
End Sub

Public Sub SetFocus()
Attribute SetFocus.VB_Description = "Moves the focus to the specified object."
Extender.SetFocus
End Sub

Public Sub ZOrder(Optional ByRef Position As Variant)
Attribute ZOrder.VB_Description = "Places a specified object at the front or back of the z-order within its graphical level."
If IsMissing(Position) Then Extender.ZOrder Else Extender.ZOrder Position
End Sub

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle to a control."
Attribute hWnd.VB_UserMemId = -515
hWnd = VBFlexGridHandle
End Property

Public Property Get hWndUserControl() As Long
Attribute hWndUserControl.VB_Description = "Returns a handle to a control."
hWndUserControl = UserControl.hWnd
End Property

Public Property Get hWndEdit() As Long
Attribute hWndEdit.VB_Description = "Returns a handle to a control."
hWndEdit = VBFlexGridEditHandle
End Property

Public Property Get hWndComboButton() As Long
Attribute hWndComboButton.VB_Description = "Returns a handle to a control."
hWndComboButton = VBFlexGridComboButtonHandle
End Property

Public Property Get hWndComboList() As Long
Attribute hWndComboList.VB_Description = "Returns a handle to a control."
hWndComboList = VBFlexGridComboListHandle
End Property

#If ImplementDataSource = True Then

Public Property Get DataSource() As MSDATASRC.DataSource
Attribute DataSource.VB_Description = "Returns/sets the data source for the control."
Attribute DataSource.VB_MemberFlags = "4"
Set DataSource = PropDataSource
End Property

Public Property Let DataSource(ByVal Value As MSDATASRC.DataSource)
Set Me.DataSource = Value
End Property

Public Property Set DataSource(ByVal Value As MSDATASRC.DataSource)
Set PropDataSource = Value
If VBFlexGridDesignMode = False Then
    If Not PropDataSource Is Nothing Then
        If PropRecordset Is Nothing Then Set PropRecordset = CreateObject("ADODB.Recordset")
        With PropRecordset
        If .State <> 0 Then .Close
        If StrPtr(PropDataMember) = 0 Then .DataMember = "" Else .DataMember = PropDataMember
        Set .DataSource = PropDataSource
        If .State <> 0 Then
            If .RecordCount > -1 Then ' The cursor type of the Recordset affects whether the number of records can be determined.
                Me.Rows = PropFixedRows + .RecordCount
                Me.Cols = PropFixedCols + .Fields.Count
                Dim iRow As Long, iCol As Long
                If PropFixedRows > 0 Then
                    For iCol = 0 To (PropFixedCols - 1)
                        VBFlexGridColsInfo(iCol).Key = vbNullString
                    Next iCol
                    For iCol = 0 To (.Fields.Count - 1)
                        Me.TextMatrix(0, iCol + PropFixedCols) = .Fields(iCol).Name
                        VBFlexGridColsInfo(iCol + PropFixedCols).Key = .Fields(iCol).Name
                    Next iCol
                End If
                If .RecordCount > 0 Then
                    Dim ArrRows As Variant
                    ArrRows = .GetRows(, 1) ' adBookmarkFirst
                    Dim LBoundCols As Long, UBoundCols As Long
                    LBoundCols = LBound(ArrRows, 1)
                    UBoundCols = UBound(ArrRows, 1)
                    Dim LBoundRows As Long, UBoundRows As Long
                    LBoundRows = LBound(ArrRows, 2)
                    UBoundRows = UBound(ArrRows, 2)
                    For iRow = LBoundRows To UBoundRows
                        For iCol = LBoundCols To UBoundCols
                            If Not IsNull(ArrRows(iCol, iRow)) Then
                                Me.TextMatrix((iRow + (0 - LBoundRows)) + PropFixedRows, (iCol + (0 - LBoundCols)) + PropFixedCols) = ArrRows(iCol, iRow)
                            Else
                                Me.TextMatrix((iRow + (0 - LBoundRows)) + PropFixedRows, (iCol + (0 - LBoundCols)) + PropFixedCols) = vbNullString
                            End If
                        Next iCol
                    Next iRow
                End If
            End If
        End If
        End With
    Else
        Set PropRecordset = Nothing
    End If
End If
UserControl.PropertyChanged "DataSource"
End Property

Public Property Get DataMember() As MSDATASRC.DataMember
Attribute DataMember.VB_Description = "Returns/sets the data member for the control."
Attribute DataMember.VB_MemberFlags = "4"
DataMember = PropDataMember
End Property

Public Property Let DataMember(ByVal Value As MSDATASRC.DataMember)
PropDataMember = Value
Set Me.DataSource = PropDataSource
UserControl.PropertyChanged "DataMember"
End Property

#End If

Public Property Get Font() As StdFont
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
Set Font = PropFont
End Property

Public Property Let Font(ByVal NewFont As StdFont)
Set Me.Font = NewFont
End Property

Public Property Set Font(ByVal NewFont As StdFont)
If NewFont Is Nothing Then Set NewFont = Ambient.Font
Dim OldFontHandle As Long
Set PropFont = NewFont
OldFontHandle = VBFlexGridFontHandle
VBFlexGridFontHandle = CreateGDIFontFromOLEFont(PropFont)
Dim hDCScreen As Long
hDCScreen = GetDC(0)
If hDCScreen <> 0 Then
    Dim TM As TEXTMETRIC, hFontOld As Long
    If VBFlexGridFontHandle <> 0 Then hFontOld = SelectObject(hDCScreen, VBFlexGridFontHandle)
    If GetTextMetrics(hDCScreen, TM) <> 0 Then
        VBFlexGridDefaultRowHeight = TM.TMHeight + (ROWINFO_HEIGHT_SPACING_DIP * PixelsPerDIP_Y())
        VBFlexGridDefaultColWidth = VBFlexGridDefaultRowHeight * RATIO_OF_ROWINFO_HEIGHT_TO_COLINFO_WIDTH
    End If
    If hFontOld <> 0 Then SelectObject hDCScreen, hFontOld
    ReleaseDC 0, hDCScreen
End If
Me.Refresh
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
UserControl.PropertyChanged "Font"
End Property

Private Sub PropFont_FontChanged(ByVal PropertyName As String)
Dim OldFontHandle As Long
OldFontHandle = VBFlexGridFontHandle
VBFlexGridFontHandle = CreateGDIFontFromOLEFont(PropFont)
Dim hDCScreen As Long
hDCScreen = GetDC(0)
If hDCScreen <> 0 Then
    Dim TM As TEXTMETRIC, hFontOld As Long
    If VBFlexGridFontHandle <> 0 Then hFontOld = SelectObject(hDCScreen, VBFlexGridFontHandle)
    If GetTextMetrics(hDCScreen, TM) <> 0 Then
        VBFlexGridDefaultRowHeight = TM.TMHeight + (ROWINFO_HEIGHT_SPACING_DIP * PixelsPerDIP_Y())
        VBFlexGridDefaultColWidth = VBFlexGridDefaultRowHeight * RATIO_OF_ROWINFO_HEIGHT_TO_COLINFO_WIDTH
    End If
    If hFontOld <> 0 Then SelectObject hDCScreen, hFontOld
    ReleaseDC 0, hDCScreen
End If
Me.Refresh
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
UserControl.PropertyChanged "Font"
End Sub

Public Property Get FontFixed() As StdFont
Attribute FontFixed.VB_Description = "Returns a Font object."
If PropFontFixed Is Nothing Then
    Set FontFixed = PropFont
Else
    Set FontFixed = PropFontFixed
End If
End Property

Public Property Let FontFixed(ByVal NewFont As StdFont)
Set Me.FontFixed = NewFont
End Property

Public Property Set FontFixed(ByVal NewFont As StdFont)
Dim OldFontHandle As Long
Set PropFontFixed = NewFont
OldFontHandle = VBFlexGridFontFixedHandle
If PropFontFixed Is Nothing Then
    VBFlexGridFontFixedHandle = 0
    VBFlexGridDefaultFixedRowHeight = -1
    VBFlexGridDefaultFixedColWidth = -1
Else
    VBFlexGridFontFixedHandle = CreateGDIFontFromOLEFont(PropFontFixed)
    Dim hDCScreen As Long
    hDCScreen = GetDC(0)
    If hDCScreen <> 0 Then
        Dim TM As TEXTMETRIC, hFontOld As Long
        If VBFlexGridFontFixedHandle <> 0 Then hFontOld = SelectObject(hDCScreen, VBFlexGridFontFixedHandle)
        If GetTextMetrics(hDCScreen, TM) <> 0 Then
            VBFlexGridDefaultFixedRowHeight = TM.TMHeight + (ROWINFO_HEIGHT_SPACING_DIP * PixelsPerDIP_Y())
            VBFlexGridDefaultFixedColWidth = VBFlexGridDefaultFixedRowHeight * RATIO_OF_ROWINFO_HEIGHT_TO_COLINFO_WIDTH
        End If
        If hFontOld <> 0 Then SelectObject hDCScreen, hFontOld
        ReleaseDC 0, hDCScreen
    End If
End If
Me.Refresh
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
UserControl.PropertyChanged "FontFixed"
End Property

Private Sub PropFontFixed_FontChanged(ByVal PropertyName As String)
Dim OldFontHandle As Long
OldFontHandle = VBFlexGridFontFixedHandle
VBFlexGridFontFixedHandle = CreateGDIFontFromOLEFont(PropFontFixed)
Dim hDCScreen As Long
hDCScreen = GetDC(0)
If hDCScreen <> 0 Then
    Dim TM As TEXTMETRIC, hFontOld As Long
    If VBFlexGridFontFixedHandle <> 0 Then hFontOld = SelectObject(hDCScreen, VBFlexGridFontFixedHandle)
    If GetTextMetrics(hDCScreen, TM) <> 0 Then
        VBFlexGridDefaultFixedRowHeight = TM.TMHeight + (ROWINFO_HEIGHT_SPACING_DIP * PixelsPerDIP_Y())
        VBFlexGridDefaultFixedColWidth = VBFlexGridDefaultFixedRowHeight * RATIO_OF_ROWINFO_HEIGHT_TO_COLINFO_WIDTH
    End If
    If hFontOld <> 0 Then SelectObject hDCScreen, hFontOld
    ReleaseDC 0, hDCScreen
End If
Me.Refresh
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
UserControl.PropertyChanged "FontFixed"
End Sub

Public Property Get VisualStyles() As Boolean
Attribute VisualStyles.VB_Description = "Returns/sets a value that determines whether the visual styles are enabled or not. Requires comctl32.dll version 6.0 or higher."
VisualStyles = PropVisualStyles
End Property

Public Property Let VisualStyles(ByVal Value As Boolean)
PropVisualStyles = Value
VBFlexGridEnabledVisualStyles = EnabledVisualStyles()
If VBFlexGridHandle <> 0 And VBFlexGridEnabledVisualStyles = True Then
    If PropVisualStyles = True Then
        ActivateVisualStyles VBFlexGridHandle
    Else
        RemoveVisualStyles VBFlexGridHandle
    End If
    Call SetVisualStylesToolTip
    Me.Refresh
    If VBFlexGridDesignMode = True Then SetWindowPos UserControl.hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
End If
UserControl.PropertyChanged "VisualStyles"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
BackColor = PropBackColor
End Property

Public Property Let BackColor(ByVal Value As OLE_COLOR)
PropBackColor = Value
PropBackColorAlt = Value
If VBFlexGridHandle <> 0 Then
    If VBFlexGridBackColorBrush <> 0 Then DeleteObject VBFlexGridBackColorBrush
    VBFlexGridBackColorBrush = CreateSolidBrush(WinColor(PropBackColor))
    If VBFlexGridBackColorAltBrush <> 0 Then DeleteObject VBFlexGridBackColorAltBrush
    VBFlexGridBackColorAltBrush = CreateSolidBrush(WinColor(PropBackColorAlt))
End If
Me.Refresh
UserControl.PropertyChanged "BackColor"
End Property

Public Property Get BackColorAlt() As OLE_COLOR
Attribute BackColorAlt.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
BackColorAlt = PropBackColorAlt
End Property

Public Property Let BackColorAlt(ByVal Value As OLE_COLOR)
PropBackColorAlt = Value
If VBFlexGridHandle <> 0 Then
    If VBFlexGridBackColorAltBrush <> 0 Then DeleteObject VBFlexGridBackColorAltBrush
    VBFlexGridBackColorAltBrush = CreateSolidBrush(WinColor(PropBackColorAlt))
End If
Me.Refresh
UserControl.PropertyChanged "BackColorAlt"
End Property

Public Property Get BackColorBkg() As OLE_COLOR
Attribute BackColorBkg.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute BackColorBkg.VB_UserMemId = -501
BackColorBkg = PropBackColorBkg
End Property

Public Property Let BackColorBkg(ByVal Value As OLE_COLOR)
PropBackColorBkg = Value
If VBFlexGridHandle <> 0 Then
    If VBFlexGridBackColorBkgBrush <> 0 Then DeleteObject VBFlexGridBackColorBkgBrush
    VBFlexGridBackColorBkgBrush = CreateSolidBrush(WinColor(PropBackColorBkg))
End If
UserControl.BackColor = PropBackColorBkg
Me.Refresh
UserControl.PropertyChanged "BackColorBkg"
End Property

Public Property Get BackColorFixed() As OLE_COLOR
Attribute BackColorFixed.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
BackColorFixed = PropBackColorFixed
End Property

Public Property Let BackColorFixed(ByVal Value As OLE_COLOR)
PropBackColorFixed = Value
If VBFlexGridHandle <> 0 Then
    If VBFlexGridBackColorFixedBrush <> 0 Then DeleteObject VBFlexGridBackColorFixedBrush
    VBFlexGridBackColorFixedBrush = CreateSolidBrush(WinColor(PropBackColorFixed))
End If
Me.Refresh
UserControl.PropertyChanged "BackColorFixed"
End Property

Public Property Get BackColorSel() As OLE_COLOR
Attribute BackColorSel.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
BackColorSel = PropBackColorSel
End Property

Public Property Let BackColorSel(ByVal Value As OLE_COLOR)
PropBackColorSel = Value
If VBFlexGridHandle <> 0 Then
    If VBFlexGridBackColorSelBrush <> 0 Then DeleteObject VBFlexGridBackColorSelBrush
    VBFlexGridBackColorSelBrush = CreateSolidBrush(WinColor(PropBackColorSel))
End If
Me.Refresh
UserControl.PropertyChanged "BackColorSel"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
Attribute ForeColor.VB_UserMemId = -513
ForeColor = PropForeColor
End Property

Public Property Let ForeColor(ByVal Value As OLE_COLOR)
PropForeColor = Value
Me.Refresh
UserControl.PropertyChanged "ForeColor"
End Property

Public Property Get ForeColorFixed() As OLE_COLOR
Attribute ForeColorFixed.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
ForeColorFixed = PropForeColorFixed
End Property

Public Property Let ForeColorFixed(ByVal Value As OLE_COLOR)
PropForeColorFixed = Value
Me.Refresh
UserControl.PropertyChanged "ForeColorFixed"
End Property

Public Property Get ForeColorSel() As OLE_COLOR
Attribute ForeColorSel.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
ForeColorSel = PropForeColorSel
End Property

Public Property Let ForeColorSel(ByVal Value As OLE_COLOR)
PropForeColorSel = Value
Me.Refresh
UserControl.PropertyChanged "ForeColorSel"
End Property

Public Property Get GridColor() As OLE_COLOR
Attribute GridColor.VB_Description = "Returns/sets the color used to draw the lines between flex grid cells."
GridColor = PropGridColor
End Property

Public Property Let GridColor(ByVal Value As OLE_COLOR)
PropGridColor = Value
If VBFlexGridHandle <> 0 Then
    If VBFlexGridGridLinePen <> 0 Then DeleteObject VBFlexGridGridLinePen
    VBFlexGridGridLinePen = CreatePen(VBFlexGridPenStyle, PropGridLineWidth, WinColor(PropGridColor))
End If
Me.Refresh
UserControl.PropertyChanged "GridColor"
End Property

Public Property Get GridColorFixed() As OLE_COLOR
Attribute GridColorFixed.VB_Description = "Returns/sets the color used to draw the lines between flex grid cells."
GridColorFixed = PropGridColorFixed
End Property

Public Property Let GridColorFixed(ByVal Value As OLE_COLOR)
PropGridColorFixed = Value
If VBFlexGridHandle <> 0 Then
    If VBFlexGridGridLineFixedPen <> 0 Then DeleteObject VBFlexGridGridLineFixedPen
    VBFlexGridGridLineFixedPen = CreatePen(VBFlexGridFixedPenStyle, PropGridLineWidth, WinColor(PropGridColorFixed))
End If
Me.Refresh
UserControl.PropertyChanged "GridColorFixed"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_UserMemId = -514
Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal Value As Boolean)
UserControl.Enabled = Value
If VBFlexGridHandle <> 0 And VBFlexGridDesignMode = False Then EnableWindow VBFlexGridHandle, IIf(Value = True, 1, 0)
UserControl.PropertyChanged "Enabled"
End Property

Public Property Get OLEDropMode() As FlexOLEDropModeConstants
Attribute OLEDropMode.VB_Description = "Returns/Sets whether this object can act as an OLE drop target."
OLEDropMode = UserControl.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal Value As FlexOLEDropModeConstants)
Select Case Value
    Case FlexOLEDropModeNone, FlexOLEDropModeManual
        UserControl.OLEDropMode = Value
    Case Else
        Err.Raise 380
End Select
UserControl.PropertyChanged "OLEDropMode"
End Property

Public Property Get MousePointer() As FlexMousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
MousePointer = PropMousePointer
End Property

Public Property Let MousePointer(ByVal Value As FlexMousePointerConstants)
Select Case Value
    Case 0 To 16, 99
        PropMousePointer = Value
    Case Else
        Err.Raise 380
End Select
If VBFlexGridDesignMode = False Then Call RefreshMousePointer
UserControl.PropertyChanged "MousePointer"
End Property

Public Property Get MouseIcon() As IPictureDisp
Attribute MouseIcon.VB_Description = "Returns/sets a custom mouse icon."
Set MouseIcon = PropMouseIcon
End Property

Public Property Let MouseIcon(ByVal Value As IPictureDisp)
Set Me.MouseIcon = Value
End Property

Public Property Set MouseIcon(ByVal Value As IPictureDisp)
If Value Is Nothing Then
    Set PropMouseIcon = Nothing
Else
    If Value.Type = vbPicTypeIcon Or Value.Handle = 0 Then
        Set PropMouseIcon = Value
    Else
        If VBFlexGridDesignMode = True Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
    End If
End If
If VBFlexGridDesignMode = False Then Call RefreshMousePointer
UserControl.PropertyChanged "MouseIcon"
End Property

Public Property Get MouseTrack() As Boolean
Attribute MouseTrack.VB_Description = "Returns/sets whether mouse events occurs when the mouse pointer enters or leaves the control."
MouseTrack = PropMouseTrack
End Property

Public Property Let MouseTrack(ByVal Value As Boolean)
PropMouseTrack = Value
UserControl.PropertyChanged "MouseTrack"
End Property

Public Property Get RightToLeft() As Boolean
Attribute RightToLeft.VB_Description = "Determines text display direction and control visual appearance on a bidirectional system."
Attribute RightToLeft.VB_UserMemId = -611
RightToLeft = PropRightToLeft
End Property

Public Property Let RightToLeft(ByVal Value As Boolean)
PropRightToLeft = Value
UserControl.RightToLeft = PropRightToLeft
If PropRightToLeft = True Then
    Select Case PropRightToLeftMode
        Case FlexRightToLeftModeNoControl
        Case FlexRightToLeftModeVBAME
            PropRightToLeft = UserControl.RightToLeft
        Case FlexRightToLeftModeSystemLocale, FlexRightToLeftModeUserLocale, FlexRightToLeftModeOSLanguage
            Const LOCALE_FONTSIGNATURE As Long = &H58, SORT_DEFAULT As Long = &H0
            Dim LangID As Integer, LCID As Long, LocaleSig As TLOCALESIGNATURE
            Select Case PropRightToLeftMode
                Case FlexRightToLeftModeSystemLocale
                    LangID = GetSystemDefaultLangID()
                Case FlexRightToLeftModeUserLocale
                    LangID = GetUserDefaultLangID()
                Case FlexRightToLeftModeOSLanguage
                    LangID = GetUserDefaultUILanguage()
            End Select
            LCID = (SORT_DEFAULT * &H10000) Or LangID
            If GetLocaleInfo(LCID, LOCALE_FONTSIGNATURE, VarPtr(LocaleSig), (LenB(LocaleSig) / 2)) <> 0 Then
                ' Unicode subset bitfield 0 to 127. Bit 123 = Layout progress, horizontal from right to left
                PropRightToLeft = CBool((LocaleSig.lsUsb(15) And (2 ^ (4 - 1))) <> 0)
            End If
    End Select
End If
Dim dwMask As Long, dwExStyle As Long
If VBFlexGridDesignMode = False Then
    ' Only on run-time the UserControl gets the mirror placement with WS_EX_LAYOUTRTL.
    ' On design-time the mirror effect will be simulated by setting WS_EX_LEFTSCROLLBAR and SetLayout API.
    ' This way the design-time dragging of the control on a form will not be reversed and works as expected.
    If PropRightToLeft = True And PropRightToLeftLayout = True Then dwMask = WS_EX_LAYOUTRTL
    dwExStyle = GetWindowLong(UserControl.hWnd, GWL_EXSTYLE)
    If (dwExStyle And WS_EX_LAYOUTRTL) = WS_EX_LAYOUTRTL Then dwExStyle = dwExStyle And Not WS_EX_LAYOUTRTL
    If (dwExStyle And WS_EX_RTLREADING) = WS_EX_RTLREADING Then dwExStyle = dwExStyle And Not WS_EX_RTLREADING
    If (dwExStyle And WS_EX_RIGHT) = WS_EX_RIGHT Then dwExStyle = dwExStyle And Not WS_EX_RIGHT
    If (dwExStyle And WS_EX_LEFTSCROLLBAR) = WS_EX_LEFTSCROLLBAR Then dwExStyle = dwExStyle And Not WS_EX_LEFTSCROLLBAR
    If (dwMask And WS_EX_LAYOUTRTL) = WS_EX_LAYOUTRTL Then dwExStyle = dwExStyle Or WS_EX_LAYOUTRTL
    If (dwMask And WS_EX_RTLREADING) = WS_EX_RTLREADING Then dwExStyle = dwExStyle Or WS_EX_RTLREADING
    If (dwMask And WS_EX_RIGHT) = WS_EX_RIGHT Then dwExStyle = dwExStyle Or WS_EX_RIGHT
    If (dwMask And WS_EX_LEFTSCROLLBAR) = WS_EX_LEFTSCROLLBAR Then dwExStyle = dwExStyle Or WS_EX_LEFTSCROLLBAR
    SetWindowLong UserControl.hWnd, GWL_EXSTYLE, dwExStyle
    InvalidateRect UserControl.hWnd, ByVal 0&, 1
    SetWindowPos UserControl.hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
    dwMask = 0
    dwExStyle = 0
Else
    ' On design-time the right-to-left layout and reading flag are set manually on this property.
    ' Whereas on run-time the flags are set when receiving the WM_STYLECHANGED message.
    ' This enables an application to change the bidirectional appearance on his own,
    ' independently from this property. (by setting either WS_EX_LAYOUTRTL or WS_EX_RTLREADING)
    VBFlexGridRTLLayout = CBool(PropRightToLeft = True And PropRightToLeftLayout = True)
    VBFlexGridRTLReading = CBool(PropRightToLeft = True And PropRightToLeftLayout = False)
End If
If VBFlexGridHandle <> 0 Then
    If PropRightToLeft = True Then
        If VBFlexGridDesignMode = False Then
            If PropRightToLeftLayout = True Then dwMask = WS_EX_LAYOUTRTL Else dwMask = WS_EX_RTLREADING
        Else
            If PropRightToLeftLayout = True Then dwMask = WS_EX_LEFTSCROLLBAR
        End If
    End If
    dwExStyle = GetWindowLong(VBFlexGridHandle, GWL_EXSTYLE)
    If (dwExStyle And WS_EX_LAYOUTRTL) = WS_EX_LAYOUTRTL Then dwExStyle = dwExStyle And Not WS_EX_LAYOUTRTL
    If (dwExStyle And WS_EX_RTLREADING) = WS_EX_RTLREADING Then dwExStyle = dwExStyle And Not WS_EX_RTLREADING
    If (dwExStyle And WS_EX_RIGHT) = WS_EX_RIGHT Then dwExStyle = dwExStyle And Not WS_EX_RIGHT
    If (dwExStyle And WS_EX_LEFTSCROLLBAR) = WS_EX_LEFTSCROLLBAR Then dwExStyle = dwExStyle And Not WS_EX_LEFTSCROLLBAR
    If (dwMask And WS_EX_LAYOUTRTL) = WS_EX_LAYOUTRTL Then dwExStyle = dwExStyle Or WS_EX_LAYOUTRTL
    If (dwMask And WS_EX_RTLREADING) = WS_EX_RTLREADING Then dwExStyle = dwExStyle Or WS_EX_RTLREADING
    If (dwMask And WS_EX_RIGHT) = WS_EX_RIGHT Then dwExStyle = dwExStyle Or WS_EX_RIGHT
    If (dwMask And WS_EX_LEFTSCROLLBAR) = WS_EX_LEFTSCROLLBAR Then dwExStyle = dwExStyle Or WS_EX_LEFTSCROLLBAR
    SetWindowLong VBFlexGridHandle, GWL_EXSTYLE, dwExStyle
    InvalidateRect VBFlexGridHandle, ByVal 0&, 1
    SetWindowPos VBFlexGridHandle, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
    dwMask = 0
    dwExStyle = 0
End If
If VBFlexGridToolTipHandle <> 0 Then
    If PropRightToLeft = True Then
        If PropRightToLeftLayout = True Then dwMask = WS_EX_LAYOUTRTL Else dwMask = WS_EX_RTLREADING
    Else
        dwMask = 0
    End If
    dwExStyle = GetWindowLong(VBFlexGridToolTipHandle, GWL_EXSTYLE)
    If (dwExStyle And WS_EX_LAYOUTRTL) = WS_EX_LAYOUTRTL Then dwExStyle = dwExStyle And Not WS_EX_LAYOUTRTL
    If (dwExStyle And WS_EX_RTLREADING) = WS_EX_RTLREADING Then dwExStyle = dwExStyle And Not WS_EX_RTLREADING
    If (dwExStyle And WS_EX_RIGHT) = WS_EX_RIGHT Then dwExStyle = dwExStyle And Not WS_EX_RIGHT
    If (dwExStyle And WS_EX_LEFTSCROLLBAR) = WS_EX_LEFTSCROLLBAR Then dwExStyle = dwExStyle And Not WS_EX_LEFTSCROLLBAR
    If (dwMask And WS_EX_LAYOUTRTL) = WS_EX_LAYOUTRTL Then dwExStyle = dwExStyle Or WS_EX_LAYOUTRTL
    ' ToolTip control supports only the WS_EX_LAYOUTRTL flag.
    ' Set TTF_RTLREADING flag when dwMask contains WS_EX_RTLREADING, though WS_EX_RTLREADING will not be actually set.
    SetWindowLong VBFlexGridToolTipHandle, GWL_EXSTYLE, dwExStyle
    Dim i As Long, TI As TOOLINFO, Buffer As String
    With TI
    .cbSize = LenB(TI)
    Buffer = String(80, vbNullChar)
    .lpszText = StrPtr(Buffer)
    For i = 1 To SendMessage(VBFlexGridToolTipHandle, TTM_GETTOOLCOUNT, 0, ByVal 0&)
        If SendMessage(VBFlexGridToolTipHandle, TTM_ENUMTOOLS, i - 1, ByVal VarPtr(TI)) <> 0 Then
            If (dwMask And WS_EX_LAYOUTRTL) = WS_EX_LAYOUTRTL Or (dwMask And WS_EX_RTLREADING) = 0 Then
                If (.uFlags And TTF_RTLREADING) = TTF_RTLREADING Then .uFlags = .uFlags And Not TTF_RTLREADING
            Else
                If (.uFlags And TTF_RTLREADING) = 0 Then .uFlags = .uFlags Or TTF_RTLREADING
            End If
            SendMessage VBFlexGridToolTipHandle, TTM_SETTOOLINFO, 0, ByVal VarPtr(TI)
            SendMessage VBFlexGridToolTipHandle, TTM_UPDATE, 0, ByVal 0&
        End If
    Next i
    End With
End If
UserControl.PropertyChanged "RightToLeft"
End Property

Public Property Get RightToLeftLayout() As Boolean
Attribute RightToLeftLayout.VB_Description = "Returns/sets a value indicating if right-to-left mirror placement is turned on."
RightToLeftLayout = PropRightToLeftLayout
End Property

Public Property Let RightToLeftLayout(ByVal Value As Boolean)
PropRightToLeftLayout = Value
Me.RightToLeft = PropRightToLeft
UserControl.PropertyChanged "RightToLeftLayout"
End Property

Public Property Get RightToLeftMode() As FlexRightToLeftModeConstants
Attribute RightToLeftMode.VB_Description = "Returns/sets the right-to-left mode."
RightToLeftMode = PropRightToLeftMode
End Property

Public Property Let RightToLeftMode(ByVal Value As FlexRightToLeftModeConstants)
Select Case Value
    Case FlexRightToLeftModeNoControl, FlexRightToLeftModeVBAME, FlexRightToLeftModeSystemLocale, FlexRightToLeftModeUserLocale, FlexRightToLeftModeOSLanguage
        PropRightToLeftMode = Value
    Case Else
        Err.Raise 380
End Select
Me.RightToLeft = PropRightToLeft
UserControl.PropertyChanged "RightToLeftMode"
End Property

Public Property Get BorderStyle() As FlexBorderStyleConstants
Attribute BorderStyle.VB_Description = "Returns/sets the border style."
Attribute BorderStyle.VB_UserMemId = -504
BorderStyle = PropBorderStyle
End Property

Public Property Let BorderStyle(ByVal Value As FlexBorderStyleConstants)
Select Case Value
    Case FlexBorderStyleNone, FlexBorderStyleSingle, FlexBorderStyleThin, FlexBorderStyleSunken, FlexBorderStyleRaised
        PropBorderStyle = Value
    Case Else
        Err.Raise 380
End Select
If VBFlexGridHandle <> 0 Then
    Dim dwStyle As Long, dwExStyle As Long
    dwStyle = GetWindowLong(VBFlexGridHandle, GWL_STYLE)
    dwExStyle = GetWindowLong(VBFlexGridHandle, GWL_EXSTYLE)
    If (dwStyle And WS_BORDER) = WS_BORDER Then dwStyle = dwStyle And Not WS_BORDER
    If (dwStyle And WS_DLGFRAME) = WS_DLGFRAME Then dwStyle = dwStyle And Not WS_DLGFRAME
    If (dwExStyle And WS_EX_STATICEDGE) = WS_EX_STATICEDGE Then dwExStyle = dwExStyle And Not WS_EX_STATICEDGE
    If (dwExStyle And WS_EX_CLIENTEDGE) = WS_EX_CLIENTEDGE Then dwExStyle = dwExStyle And Not WS_EX_CLIENTEDGE
    If (dwExStyle And WS_EX_WINDOWEDGE) = WS_EX_WINDOWEDGE Then dwExStyle = dwExStyle And Not WS_EX_WINDOWEDGE
    Select Case PropBorderStyle
        Case FlexBorderStyleSingle
            dwStyle = dwStyle Or WS_BORDER
        Case FlexBorderStyleThin
            dwExStyle = dwExStyle Or WS_EX_STATICEDGE
        Case FlexBorderStyleSunken
            dwExStyle = dwExStyle Or WS_EX_CLIENTEDGE
        Case FlexBorderStyleRaised
            dwExStyle = dwExStyle Or WS_EX_WINDOWEDGE
            dwStyle = dwStyle Or WS_DLGFRAME
    End Select
    SetWindowLong VBFlexGridHandle, GWL_STYLE, dwStyle
    SetWindowLong VBFlexGridHandle, GWL_EXSTYLE, dwExStyle
    SetWindowPos VBFlexGridHandle, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
End If
UserControl.PropertyChanged "BorderStyle"
End Property

Public Property Get FixedRows() As Long
Attribute FixedRows.VB_Description = "Returns/sets the total number of fixed (non-scrollable) columns or rows for the flex grid."
FixedRows = PropFixedRows
End Property

Public Property Let FixedRows(ByVal Value As Long)
If Value < 0 Then
    If VBFlexGridDesignMode = True Then
        MsgBox "Invalid Row Value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise Number:=30009, Description:="Invalid Row value"
    End If
ElseIf Value >= (PropRows - PropFrozenRows) Then
    If VBFlexGridDesignMode = True Then
        MsgBox "FixedRows must be at least one less than Rows minus FrozenRows value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise Number:=30016, Description:="FixedRows must be at least one less than Rows minus FrozenRows value"
    End If
End If
PropFixedRows = Value
Dim RCP As TROWCOLPARAMS
With RCP
.Mask = RCPM_ROW Or RCPM_TOPROW
.Flags = RCPF_FORCETOPROWMASK Or RCPF_SETSCROLLBARS Or RCPF_FORCEREDRAW
.Row = PropFixedRows + PropFrozenRows
.TopRow = PropFixedRows + PropFrozenRows
Select Case PropSelectionMode
    Case FlexSelectionModeFree, FlexSelectionModeFreeByRow, FlexSelectionModeFreeByColumn
        .Mask = .Mask Or RCPM_ROWSEL
        .RowSel = .Row
    Case FlexSelectionModeByRow
        .Mask = .Mask Or RCPM_ROWSEL Or RCPM_COLSEL
        .RowSel = .Row
        .ColSel = (PropCols - 1)
End Select
Call SetRowColParams(RCP)
End With
UserControl.PropertyChanged "FixedRows"
End Property

Public Property Get FixedCols() As Long
Attribute FixedCols.VB_Description = "Returns/sets the total number of fixed (non-scrollable) columns or rows for the flex grid."
FixedCols = PropFixedCols
End Property

Public Property Let FixedCols(ByVal Value As Long)
If Value < 0 Then
    If VBFlexGridDesignMode = True Then
        MsgBox "Invalid Col value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise Number:=30010, Description:="Invalid Col value"
    End If
ElseIf Value >= (PropCols - PropFrozenCols) Then
    If VBFlexGridDesignMode = True Then
        MsgBox "FixedCols must be at least one less than Cols minus FrozenCols value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise Number:=30017, Description:="FixedCols must be at least one less than Cols minus FrozenCols value"
    End If
End If
PropFixedCols = Value
Dim RCP As TROWCOLPARAMS
With RCP
.Mask = RCPM_COL Or RCPM_LEFTCOL
.Flags = RCPF_FORCELEFTCOLMASK Or RCPF_SETSCROLLBARS Or RCPF_FORCEREDRAW
.Col = PropFixedCols + PropFrozenCols
.LeftCol = PropFixedCols + PropFrozenCols
Select Case PropSelectionMode
    Case FlexSelectionModeFree, FlexSelectionModeFreeByRow, FlexSelectionModeFreeByColumn
        .Mask = .Mask Or RCPM_COLSEL
        .ColSel = .Col
    Case FlexSelectionModeByColumn
        .Mask = .Mask Or RCPM_ROWSEL Or RCPM_COLSEL
        .RowSel = (PropRows - 1)
        .ColSel = .Col
End Select
Call SetRowColParams(RCP)
End With
UserControl.PropertyChanged "FixedCols"
End Property

Public Property Get FrozenRows() As Long
Attribute FrozenRows.VB_Description = "Returns/sets the total number of frozen (movable but non-scrollable) columns or rows for the flex grid."
FrozenRows = PropFrozenRows
End Property

Public Property Let FrozenRows(ByVal Value As Long)
If Value < 0 Then
    If VBFlexGridDesignMode = True Then
        MsgBox "Invalid Row Value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise Number:=30009, Description:="Invalid Row value"
    End If
ElseIf Value >= (PropRows - PropFixedRows) Then
    If VBFlexGridDesignMode = True Then
        MsgBox "FrozenRows must be at least one less than Rows minus FixedRows value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise Number:=30016, Description:="FrozenRows must be at least one less than Rows minus FixedRows value"
    End If
End If
PropFrozenRows = Value
Dim RCP As TROWCOLPARAMS
With RCP
.Mask = RCPM_ROW Or RCPM_TOPROW
.Flags = RCPF_FORCETOPROWMASK Or RCPF_SETSCROLLBARS Or RCPF_FORCEREDRAW
.Row = PropFixedRows + PropFrozenRows
.TopRow = PropFixedRows + PropFrozenRows
Select Case PropSelectionMode
    Case FlexSelectionModeFree, FlexSelectionModeFreeByRow, FlexSelectionModeFreeByColumn
        .Mask = .Mask Or RCPM_ROWSEL
        .RowSel = .Row
    Case FlexSelectionModeByRow
        .Mask = .Mask Or RCPM_ROWSEL Or RCPM_COLSEL
        .RowSel = .Row
        .ColSel = (PropCols - 1)
End Select
Call SetRowColParams(RCP)
End With
UserControl.PropertyChanged "FrozenRows"
End Property

Public Property Get FrozenCols() As Long
Attribute FrozenCols.VB_Description = "Returns/sets the total number of frozen (movable but non-scrollable) columns or rows for the flex grid."
FrozenCols = PropFrozenCols
End Property

Public Property Let FrozenCols(ByVal Value As Long)
If Value < 0 Then
    If VBFlexGridDesignMode = True Then
        MsgBox "Invalid Col value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise Number:=30010, Description:="Invalid Col value"
    End If
ElseIf Value >= (PropCols - PropFixedCols) Then
    If VBFlexGridDesignMode = True Then
        MsgBox "FrozenCols must be at least one less than Cols minus FixedCols value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise Number:=30017, Description:="FrozenCols must be at least one less than Cols minus FixedCols value"
    End If
End If
PropFrozenCols = Value
Dim RCP As TROWCOLPARAMS
With RCP
.Mask = RCPM_COL Or RCPM_LEFTCOL
.Flags = RCPF_FORCELEFTCOLMASK Or RCPF_SETSCROLLBARS Or RCPF_FORCEREDRAW
.Col = PropFixedCols + PropFrozenCols
.LeftCol = PropFixedCols + PropFrozenCols
Select Case PropSelectionMode
    Case FlexSelectionModeFree, FlexSelectionModeFreeByRow, FlexSelectionModeFreeByColumn
        .Mask = .Mask Or RCPM_COLSEL
        .ColSel = .Col
    Case FlexSelectionModeByColumn
        .Mask = .Mask Or RCPM_ROWSEL Or RCPM_COLSEL
        .RowSel = (PropRows - 1)
        .ColSel = .Col
End Select
Call SetRowColParams(RCP)
End With
UserControl.PropertyChanged "FrozenCols"
End Property

Public Property Get Rows() As Long
Attribute Rows.VB_Description = "Returns/sets the total number of columns or rows in the flex grid."
Attribute Rows.VB_MemberFlags = "200"
Rows = PropRows
End Property

Public Property Let Rows(ByVal Value As Long)
If Value < 0 Then
    If VBFlexGridDesignMode = True Then
        MsgBox "Invalid Row value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise Number:=30009, Description:="Invalid Row value"
    End If
Else
    If Value < PropFixedRows And Value > 0 Then PropFixedRows = Value
    If Value <= (PropFixedRows + PropFrozenRows) And Value > 0 Then
        If (Value - PropFixedRows - 1) > 0 Then
            PropFrozenRows = Value - PropFixedRows - 1
        Else
            PropFrozenRows = 0
        End If
    End If
End If
If Value > 0 And PropRows < 1 Then
    PropRows = Value
    If PropCols > 0 Then Call InitFlexGridCells
ElseIf Value < 1 And PropRows > 0 Then
    PropRows = Value
    PropFixedRows = 0
    PropFrozenRows = 0
    Call EraseFlexGridCells
ElseIf Value <> PropRows And PropCols > 0 Then
    ReDim Preserve VBFlexGridCells.Rows(0 To (Value - 1)) As TCOLS
    If Value > PropRows Then
        Dim i As Long
        ReDim Preserve VBFlexGridCells.Rows(0 To (Value - 1)) As TCOLS
        PropRows = PropRows + 1 ' First new row.
        For i = (PropRows - 1) To (Value - 1)
            LSet VBFlexGridCells.Rows(i) = VBFlexGridDefaultCols
        Next i
    End If
    PropRows = Value
Else
    PropRows = Value
End If
Dim RCP As TROWCOLPARAMS
With RCP
.Flags = RCPF_SETSCROLLBARS Or RCPF_FORCEREDRAW
.Row = VBFlexGridRow
If .Row > (PropRows - 1) Then
    .Mask = .Mask Or RCPM_ROW
    .Row = (PropRows - 1)
End If
Select Case PropSelectionMode
    Case FlexSelectionModeFree, FlexSelectionModeFreeByRow, FlexSelectionModeFreeByColumn, FlexSelectionModeByRow
        If VBFlexGridRowSel > (PropRows - 1) Then
            .Mask = .Mask Or RCPM_ROWSEL
            .RowSel = (PropRows - 1)
        End If
    Case FlexSelectionModeByColumn
        If VBFlexGridRowSel <> (PropRows - 1) Then
            .Mask = .Mask Or RCPM_ROWSEL
            .RowSel = (PropRows - 1)
        End If
End Select
If .Row < PropFixedRows And PropRows > PropFixedRows Then
    ' In case there were no movable rows before and are now again available.
    ' Then it is necessary that the active row gets adjusted to the first movable row.
    If Not (.Mask And RCPM_ROW) = RCPM_ROW Then .Mask = .Mask Or RCPM_ROW
    .Row = PropFixedRows
    If Not (.Mask And RCPM_ROWSEL) = RCPM_ROWSEL Then .Mask = .Mask Or RCPM_ROWSEL
    If PropSelectionMode <> FlexSelectionModeByColumn Then .RowSel = PropFixedRows Else .RowSel = (PropRows - 1)
End If
If VBFlexGridTopRow > (PropRows - 1) Then
    .Mask = .Mask Or RCPM_TOPROW
    Select Case PropScrollBars
        Case vbVertical, vbBoth
            .Flags = .Flags Or RCPF_CHECKTOPROW
        Case Else
            .Flags = .Flags Or RCPF_FORCETOPROWMASK
    End Select
    .TopRow = (PropRows - 1)
End If
Call SetRowColParams(RCP)
End With
UserControl.PropertyChanged "Rows"
End Property

Public Property Get Cols() As Long
Attribute Cols.VB_Description = "Returns/sets the total number of columns or rows in the flex grid."
Cols = PropCols
End Property

Public Property Let Cols(ByVal Value As Long)
If Value < 0 Then
    If VBFlexGridDesignMode = True Then
        MsgBox "Invalid Col value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise Number:=30010, Description:="Invalid Col value"
    End If
Else
    If Value < PropFixedCols And Value > 0 Then PropFixedCols = Value
    If Value <= (PropFixedCols + PropFrozenCols) And Value > 0 Then
        If (Value - PropFixedCols - 1) > 0 Then
            PropFrozenCols = Value - PropFixedCols - 1
        Else
            PropFrozenCols = 0
        End If
    End If
End If
If Value > 0 And PropCols < 1 Then
    PropCols = Value
    If PropRows > 0 Then Call InitFlexGridCells
ElseIf Value < 1 And PropCols > 0 Then
    PropCols = Value
    PropFixedCols = 0
    PropFrozenCols = 0
    Call EraseFlexGridCells
ElseIf Value <> PropCols And PropRows > 0 Then
    Dim i As Long, j As Long
    If Value > PropCols Then
        PropCols = PropCols + 1 ' First new column.
        For i = 0 To (PropRows - 1)
            With VBFlexGridCells.Rows(i)
            ReDim Preserve .Cols(0 To (Value - 1)) As TCELL
            For j = (PropCols - 1) To (Value - 1)
                LSet .Cols(j) = VBFlexGridDefaultCell
            Next j
            End With
        Next i
        ReDim Preserve VBFlexGridColsInfo(0 To (Value - 1)) As TCOLINFO
        ReDim Preserve VBFlexGridDefaultCols.Cols(0 To (Value - 1)) As TCELL
        For j = (PropCols - 1) To (Value - 1)
            LSet VBFlexGridColsInfo(j) = VBFlexGridDefaultColInfo
            LSet VBFlexGridDefaultCols.Cols(j) = VBFlexGridDefaultCell
        Next j
    Else
        For i = 0 To (PropRows - 1)
            ReDim Preserve VBFlexGridCells.Rows(i).Cols(0 To (Value - 1)) As TCELL
        Next i
        ReDim Preserve VBFlexGridColsInfo(0 To (Value - 1)) As TCOLINFO
        ReDim Preserve VBFlexGridDefaultCols.Cols(0 To (Value - 1)) As TCELL
    End If
    PropCols = Value
Else
    PropCols = Value
End If
Dim RCP As TROWCOLPARAMS
With RCP
.Flags = RCPF_SETSCROLLBARS Or RCPF_FORCEREDRAW
.Col = VBFlexGridCol
If .Col > (PropCols - 1) Then
    .Mask = .Mask Or RCPM_COL
    .Col = (PropCols - 1)
End If
Select Case PropSelectionMode
    Case FlexSelectionModeFree, FlexSelectionModeFreeByRow, FlexSelectionModeFreeByColumn, FlexSelectionModeByColumn
        If VBFlexGridColSel > (PropCols - 1) Then
            .Mask = .Mask Or RCPM_COLSEL
            .ColSel = (PropCols - 1)
        End If
    Case FlexSelectionModeByRow
        If VBFlexGridColSel <> (PropCols - 1) Then
            .Mask = .Mask Or RCPM_COLSEL
            .ColSel = (PropCols - 1)
        End If
End Select
If .Col < PropFixedCols And PropCols > PropFixedCols Then
    ' In case there were no movable columns before and are now again available.
    ' Then it is necessary that the active col gets adjusted to the first movable column.
    If Not (.Mask And RCPM_COL) = RCPM_COL Then .Mask = .Mask Or RCPM_COL
    .Col = PropFixedCols
    If Not (.Mask And RCPM_COLSEL) = RCPM_COLSEL Then .Mask = .Mask Or RCPM_COLSEL
    If PropSelectionMode <> FlexSelectionModeByRow Then .ColSel = PropFixedCols Else .ColSel = (PropCols - 1)
End If
If VBFlexGridLeftCol > (PropCols - 1) Then
    .Mask = .Mask Or RCPM_LEFTCOL
    Select Case PropScrollBars
        Case vbHorizontal, vbBoth
            .Flags = .Flags Or RCPF_CHECKLEFTCOL
        Case Else
            .Flags = .Flags Or RCPF_FORCELEFTCOLMASK
    End Select
    .LeftCol = (PropCols - 1)
End If
Call SetRowColParams(RCP)
End With
UserControl.PropertyChanged "Cols"
End Property

Public Property Get AllowBigSelection() As Boolean
Attribute AllowBigSelection.VB_Description = "Returns/sets whether clicking on a column or row header should cause the entire column or row to be selected."
AllowBigSelection = PropAllowBigSelection
End Property

Public Property Let AllowBigSelection(ByVal Value As Boolean)
PropAllowBigSelection = Value
UserControl.PropertyChanged "AllowBigSelection"
End Property

Public Property Get AllowSelection() As Boolean
Attribute AllowSelection.VB_Description = "Returns/sets a value indicating if the flex grid enables selection of cells."
AllowSelection = PropAllowSelection
End Property

Public Property Let AllowSelection(ByVal Value As Boolean)
PropAllowSelection = Value
Dim RCP As TROWCOLPARAMS
With RCP
.Mask = RCPM_ROWSEL Or RCPM_COLSEL
If PropAllowSelection = True Then
    Select Case PropSelectionMode
        Case FlexSelectionModeFree, FlexSelectionModeFreeByRow, FlexSelectionModeFreeByColumn
            .RowSel = VBFlexGridRow
            .ColSel = VBFlexGridCol
        Case FlexSelectionModeByRow
            .RowSel = VBFlexGridRow
            .ColSel = (PropCols - 1)
        Case FlexSelectionModeByColumn
            .RowSel = (PropRows - 1)
            .ColSel = VBFlexGridCol
    End Select
Else
    .RowSel = VBFlexGridRow
    .ColSel = VBFlexGridCol
End If
Call SetRowColParams(RCP)
End With
UserControl.PropertyChanged "AllowSelection"
End Property

Public Property Get AllowUserEditing() As Boolean
Attribute AllowUserEditing.VB_Description = "Returns/sets a value that determines if a user can edit the text of a cell. The control can be forced to go into editing mode using the 'start edit' method."
AllowUserEditing = PropAllowUserEditing
End Property

Public Property Let AllowUserEditing(ByVal Value As Boolean)
PropAllowUserEditing = Value
UserControl.PropertyChanged "AllowUserEditing"
End Property

Public Property Get AllowUserResizing() As FlexAllowUserResizingConstants
Attribute AllowUserResizing.VB_Description = "Returns/sets whether the user should be allowed to resize rows and columns with the mouse."
AllowUserResizing = PropAllowUserResizing
End Property

Public Property Let AllowUserResizing(ByVal Value As FlexAllowUserResizingConstants)
Select Case Value
    Case FlexAllowUserResizingNone, FlexAllowUserResizingColumns, FlexAllowUserResizingRows, FlexAllowUserResizingBoth
        PropAllowUserResizing = Value
    Case Else
        Err.Raise 380
End Select
UserControl.PropertyChanged "AllowUserResizing"
End Property

Public Property Get RowSizingMode() As FlexRowSizingModeConstants
Attribute RowSizingMode.VB_Description = "Returns/sets the row sizing mode."
RowSizingMode = PropRowSizingMode
End Property

Public Property Let RowSizingMode(ByVal Value As FlexRowSizingModeConstants)
Select Case Value
    Case FlexRowSizingModeIndividual, FlexRowSizingModeAll
        PropRowSizingMode = Value
    Case Else
        Err.Raise 380
End Select
UserControl.PropertyChanged "RowSizingMode"
End Property

Public Property Get MergeCells() As FlexMergeCellsConstants
Attribute MergeCells.VB_Description = "Returns/sets whether cells with the same contents should be grouped in a single cell spanning multiple rows or columns."
MergeCells = PropMergeCells
End Property

Public Property Let MergeCells(ByVal Value As FlexMergeCellsConstants)
Select Case Value
    Case FlexMergeCellsNever, FlexMergeCellsFree, FlexMergeCellsRestrictRows, FlexMergeCellsRestrictColumns, FlexMergeCellsRestrictAll, FlexMergeCellsFixedOnly
        PropMergeCells = Value
    Case Else
        Err.Raise 380
End Select
Call RedrawGrid
UserControl.PropertyChanged "MergeCells"
End Property

Public Property Get SelectionMode() As FlexSelectionModeConstants
Attribute SelectionMode.VB_Description = "Returns/sets whether the flex grid should allow regular cell selection, selection by rows, or selection by columns."
SelectionMode = PropSelectionMode
End Property

Public Property Let SelectionMode(ByVal Value As FlexSelectionModeConstants)
Select Case Value
    Case FlexSelectionModeFree, FlexSelectionModeByRow, FlexSelectionModeByColumn, FlexSelectionModeFreeByRow, FlexSelectionModeFreeByColumn
        PropSelectionMode = Value
    Case Else
        Err.Raise 380
End Select
Dim RCP As TROWCOLPARAMS
With RCP
.Mask = RCPM_ROW Or RCPM_COL Or RCPM_ROWSEL Or RCPM_COLSEL
.Row = PropFixedRows
.Col = PropFixedCols
Select Case PropSelectionMode
    Case FlexSelectionModeFree, FlexSelectionModeFreeByRow, FlexSelectionModeFreeByColumn
        .RowSel = .Row
        .ColSel = .Col
    Case FlexSelectionModeByRow
        .RowSel = .Row
        .ColSel = (PropCols - 1)
    Case FlexSelectionModeByColumn
        .RowSel = (PropRows - 1)
        .ColSel = .Col
End Select
Call SetRowColParams(RCP)
End With
UserControl.PropertyChanged "SelectionMode"
End Property

Public Property Get FillStyle() As FlexFillStyleConstants
Attribute FillStyle.VB_Description = "Returns/sets whether setting the Text property or one of the cell formatting properties of the flex grid applies the change to all selected cells."
FillStyle = PropFillStyle
End Property

Public Property Let FillStyle(ByVal Value As FlexFillStyleConstants)
Select Case Value
    Case FlexFillStyleSingle, FlexFillStyleRepeat
        PropFillStyle = Value
    Case Else
        Err.Raise 380
End Select
UserControl.PropertyChanged "FillStyle"
End Property

Public Property Get ScrollBars() As VBRUN.ScrollBarConstants
Attribute ScrollBars.VB_Description = "Returns/sets a value indicating whether an object has vertical or horizontal scroll bars."
ScrollBars = PropScrollBars
End Property

Public Property Let ScrollBars(ByVal Value As VBRUN.ScrollBarConstants)
Select Case Value
    Case vbSBNone, vbHorizontal, vbVertical, vbBoth
        PropScrollBars = Value
    Case Else
        Err.Raise 380
End Select
Call SetScrollBars
UserControl.PropertyChanged "ScrollBars"
End Property

Public Property Get ScrollTrack() As Boolean
Attribute ScrollTrack.VB_Description = "Returns/sets whether the control should scroll its contents while the user moves the scroll box along the scroll bars."
ScrollTrack = PropScrollTrack
End Property

Public Property Let ScrollTrack(ByVal Value As Boolean)
PropScrollTrack = Value
UserControl.PropertyChanged "ScrollTrack"
End Property

Public Property Get DisableNoScroll() As Boolean
Attribute DisableNoScroll.VB_Description = "Returns/sets a value that determines whether scroll bars are disabled instead of hided when they are not needed."
DisableNoScroll = PropDisableNoScroll
End Property

Public Property Let DisableNoScroll(ByVal Value As Boolean)
PropDisableNoScroll = Value
Call SetScrollBars
UserControl.PropertyChanged "DisableNoScroll"
End Property

Public Property Get HighLight() As FlexHighLightConstants
Attribute HighLight.VB_Description = "Returns/sets whether selected cells appear highlighted."
HighLight = PropHighLight
End Property

Public Property Let HighLight(ByVal Value As FlexHighLightConstants)
Select Case Value
    Case FlexHighLightNever, FlexHighLightAlways, FlexHighLightWithFocus
        PropHighLight = Value
    Case Else
        Err.Raise 380
End Select
Call RedrawGrid
UserControl.PropertyChanged "HighLight"
End Property

Public Property Get FocusRect() As FlexFocusRectConstants
Attribute FocusRect.VB_Description = "Returns/sets whether the flex grid control should draw a focus rectangle around the current cell."
FocusRect = PropFocusRect
End Property

Public Property Let FocusRect(ByVal Value As FlexFocusRectConstants)
Select Case Value
    Case FlexFocusRectNone, FlexFocusRectLight, FlexFocusRectHeavy
        PropFocusRect = Value
    Case Else
        Err.Raise 380
End Select
Call RedrawGrid
UserControl.PropertyChanged "FocusRect"
End Property

Public Property Get RowHeightMin() As Long
Attribute RowHeightMin.VB_Description = "Returns/sets a minimum row height in twips for the entire control."
RowHeightMin = UserControl.ScaleY(PropRowHeightMin, vbPixels, vbTwips)
End Property

Public Property Let RowHeightMin(ByVal Value As Long)
If Value < 0 Then
    If VBFlexGridDesignMode = True Then
        MsgBox "Invalid Row Height value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise Number:=30013, Description:="Invalid Row Height value"
    End If
End If
PropRowHeightMin = UserControl.ScaleY(Value, vbTwips, vbPixels)
Call RedrawGrid
Call SetScrollBars
UserControl.PropertyChanged "RowHeightMin"
End Property

Public Property Get RowHeightMax() As Long
Attribute RowHeightMax.VB_Description = "Returns/sets a maximum row height in twips for the entire control."
RowHeightMax = UserControl.ScaleY(PropRowHeightMax, vbPixels, vbTwips)
End Property

Public Property Let RowHeightMax(ByVal Value As Long)
If Value < 0 Then
    If VBFlexGridDesignMode = True Then
        MsgBox "Invalid Row Height value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise Number:=30013, Description:="Invalid Row Height value"
    End If
End If
PropRowHeightMax = UserControl.ScaleY(Value, vbTwips, vbPixels)
Call RedrawGrid
Call SetScrollBars
UserControl.PropertyChanged "RowHeightMax"
End Property

Public Property Get ColWidthMin() As Long
Attribute ColWidthMin.VB_Description = "Returns/sets a minimum column width in twips for the entire control."
ColWidthMin = UserControl.ScaleX(PropColWidthMin, vbPixels, vbTwips)
End Property

Public Property Let ColWidthMin(ByVal Value As Long)
If Value < 0 Then
    If VBFlexGridDesignMode = True Then
        MsgBox "Invalid Col Width value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise Number:=30014, Description:="Invalid Col Width value"
    End If
End If
PropColWidthMin = UserControl.ScaleX(Value, vbTwips, vbPixels)
Call RedrawGrid
Call SetScrollBars
UserControl.PropertyChanged "ColWidthMin"
End Property

Public Property Get ColWidthMax() As Long
Attribute ColWidthMax.VB_Description = "Returns/sets a maximum column width in twips for the entire control."
ColWidthMax = UserControl.ScaleX(PropColWidthMax, vbPixels, vbTwips)
End Property

Public Property Let ColWidthMax(ByVal Value As Long)
If Value < 0 Then
    If VBFlexGridDesignMode = True Then
        MsgBox "Invalid Col Width value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise Number:=30014, Description:="Invalid Col Width value"
    End If
End If
PropColWidthMax = UserControl.ScaleX(Value, vbTwips, vbPixels)
Call RedrawGrid
Call SetScrollBars
UserControl.PropertyChanged "ColWidthMax"
End Property

Public Property Get GridLines() As FlexGridLineConstants
Attribute GridLines.VB_Description = "Returns/sets the type of lines that should be drawn between cells."
GridLines = PropGridLines
End Property

Public Property Let GridLines(ByVal Value As FlexGridLineConstants)
Select Case Value
    Case FlexGridLineNone, FlexGridLineFlat, FlexGridLineInset, FlexGridLineRaised, FlexGridLineDashes, FlexGridLineDots
        PropGridLines = Value
        Select Case Value
            Case FlexGridLineDashes
                VBFlexGridPenStyle = PS_DASH
            Case FlexGridLineDots
                VBFlexGridPenStyle = PS_DOT
            Case Else
                VBFlexGridPenStyle = PS_SOLID
        End Select
    Case Else
        Err.Raise 380
End Select
If VBFlexGridHandle <> 0 Then
    If VBFlexGridGridLinePen <> 0 Then DeleteObject VBFlexGridGridLinePen
    VBFlexGridGridLinePen = CreatePen(VBFlexGridPenStyle, PropGridLineWidth, WinColor(PropGridColor))
End If
Call RedrawGrid
UserControl.PropertyChanged "GridLines"
End Property

Public Property Get GridLinesFixed() As FlexGridLineConstants
Attribute GridLinesFixed.VB_Description = "Returns/sets the type of lines that should be drawn between cells."
GridLinesFixed = PropGridLinesFixed
End Property

Public Property Let GridLinesFixed(ByVal Value As FlexGridLineConstants)
Select Case Value
    Case FlexGridLineNone, FlexGridLineFlat, FlexGridLineInset, FlexGridLineRaised, FlexGridLineDashes, FlexGridLineDots
        PropGridLinesFixed = Value
        Select Case Value
            Case FlexGridLineDashes
                VBFlexGridFixedPenStyle = PS_DASH
            Case FlexGridLineDots
                VBFlexGridFixedPenStyle = PS_DOT
            Case Else
                VBFlexGridFixedPenStyle = PS_SOLID
        End Select
    Case Else
        Err.Raise 380
End Select
If VBFlexGridHandle <> 0 Then
    If VBFlexGridGridLineFixedPen <> 0 Then DeleteObject VBFlexGridGridLineFixedPen
    VBFlexGridGridLineFixedPen = CreatePen(VBFlexGridFixedPenStyle, PropGridLineWidth, WinColor(PropGridColorFixed))
End If
Call RedrawGrid
UserControl.PropertyChanged "GridLinesFixed"
End Property

Public Property Get GridLineWidth() As Integer
Attribute GridLineWidth.VB_Description = "Returns/sets the width in pixels of the gridlines."
GridLineWidth = PropGridLineWidth
End Property

Public Property Let GridLineWidth(ByVal Value As Integer)
If Value < 1 Then
    If VBFlexGridDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
PropGridLineWidth = Value
If VBFlexGridHandle <> 0 Then
    If VBFlexGridGridLinePen <> 0 Then DeleteObject VBFlexGridGridLinePen
    VBFlexGridGridLinePen = CreatePen(VBFlexGridPenStyle, PropGridLineWidth, WinColor(PropGridColor))
    If VBFlexGridGridLineFixedPen <> 0 Then DeleteObject VBFlexGridGridLineFixedPen
    VBFlexGridGridLineFixedPen = CreatePen(VBFlexGridFixedPenStyle, PropGridLineWidth, WinColor(PropGridColorFixed))
End If
Call RedrawGrid
UserControl.PropertyChanged "GridLineWidth"
End Property

Public Property Get TextStyle() As FlexTextStyleConstants
Attribute TextStyle.VB_Description = "Returns/sets 3D effects for displaying text."
TextStyle = PropTextStyle
End Property

Public Property Let TextStyle(ByVal Value As FlexTextStyleConstants)
Select Case Value
    Case FlexTextStyleFlat, FlexTextStyleRaised, FlexTextStyleInset, FlexTextStyleRaisedLight, FlexTextStyleInsetLight
        PropTextStyle = Value
    Case Else
        Err.Raise 380
End Select
Call RedrawGrid
UserControl.PropertyChanged "TextStyle"
End Property

Public Property Get TextStyleFixed() As FlexTextStyleConstants
Attribute TextStyleFixed.VB_Description = "Returns/sets 3D effects for displaying text."
TextStyleFixed = PropTextStyleFixed
End Property

Public Property Let TextStyleFixed(ByVal Value As FlexTextStyleConstants)
Select Case Value
    Case FlexTextStyleFlat, FlexTextStyleRaised, FlexTextStyleInset, FlexTextStyleRaisedLight, FlexTextStyleInsetLight
        PropTextStyleFixed = Value
    Case Else
        Err.Raise 380
End Select
Call RedrawGrid
UserControl.PropertyChanged "TextStyleFixed"
End Property

Public Property Get PictureType() As FlexPictureTypeConstants
Attribute PictureType.VB_Description = "Returns/sets the type of picture that should be generated by the picture property."
PictureType = PropPictureType
End Property

Public Property Let PictureType(ByVal Value As FlexPictureTypeConstants)
Select Case Value
    Case FlexPictureTypeColor, FlexPictureTypeMonochrome
        PropPictureType = Value
    Case Else
        Err.Raise 380
End Select
UserControl.PropertyChanged "PictureType"
End Property

Public Property Get WordWrap() As Boolean
Attribute WordWrap.VB_Description = "Returns/sets whether text within a cell should be allowed to wrap."
WordWrap = PropWordWrap
End Property

Public Property Let WordWrap(ByVal Value As Boolean)
If PropSingleLine = True And Value = True Then
    If VBFlexGridDesignMode = True Then
        MsgBox "WordWrap must be False when SingleLine is True", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise Number:=383, Description:="WordWrap must be False when SingleLine is True"
    End If
End If
PropWordWrap = Value
Call RedrawGrid
UserControl.PropertyChanged "WordWrap"
End Property

Public Property Get SingleLine() As Boolean
Attribute SingleLine.VB_Description = "Returns/sets whether text within a cell is displayed on a single line only."
SingleLine = PropSingleLine
End Property

Public Property Let SingleLine(ByVal Value As Boolean)
PropSingleLine = Value
If PropSingleLine = True Then PropWordWrap = False
Call RedrawGrid
UserControl.PropertyChanged "SingleLine"
End Property

Public Property Get EllipsisFormat() As FlexEllipsisFormatConstants
Attribute EllipsisFormat.VB_Description = "Returns/sets a value indicating if and where the ellipsis character is appended, denoting that the text extends beyond the length of the cell. The word wrap property may be set to false to see the ellipsis character."
EllipsisFormat = PropEllipsisFormat
End Property

Public Property Let EllipsisFormat(ByVal Value As FlexEllipsisFormatConstants)
Select Case Value
    Case FlexEllipsisFormatNone, FlexEllipsisFormatEnd, FlexEllipsisFormatPath, FlexEllipsisFormatWord
        PropEllipsisFormat = Value
    Case Else
        Err.Raise 380
End Select
Call RedrawGrid
UserControl.PropertyChanged "EllipsisFormat"
End Property

Public Property Get EllipsisFormatFixed() As FlexEllipsisFormatConstants
Attribute EllipsisFormatFixed.VB_Description = "Returns/sets a value indicating if and where the ellipsis character is appended, denoting that the text extends beyond the length of the cell. The word wrap property may be set to false to see the ellipsis character."
EllipsisFormatFixed = PropEllipsisFormatFixed
End Property

Public Property Let EllipsisFormatFixed(ByVal Value As FlexEllipsisFormatConstants)
Select Case Value
    Case FlexEllipsisFormatNone, FlexEllipsisFormatEnd, FlexEllipsisFormatPath, FlexEllipsisFormatWord
        PropEllipsisFormatFixed = Value
    Case Else
        Err.Raise 380
End Select
Call RedrawGrid
UserControl.PropertyChanged "EllipsisFormatFixed"
End Property

Public Property Get Redraw() As Boolean
Attribute Redraw.VB_Description = "Enables or disables redrawing of the flex grid control."
Redraw = PropRedraw
End Property

Public Property Let Redraw(ByVal Value As Boolean)
PropRedraw = Value
If VBFlexGridHandle <> 0 And VBFlexGridDesignMode = False Then
    SendMessage VBFlexGridHandle, WM_SETREDRAW, IIf(PropRedraw = True, 1, 0), ByVal 0&
    If PropRedraw = True Then
        Me.Refresh
        Call SetScrollBars
    End If
End If
UserControl.PropertyChanged "Redraw"
End Property

Public Property Get DoubleBuffer() As Boolean
Attribute DoubleBuffer.VB_Description = "Returns/sets a value that determines whether the control paints via double-buffering, which reduces flicker."
DoubleBuffer = PropDoubleBuffer
End Property

Public Property Let DoubleBuffer(ByVal Value As Boolean)
PropDoubleBuffer = Value
UserControl.PropertyChanged "DoubleBuffer"
End Property

Public Property Get Sort() As FlexSortConstants
Attribute Sort.VB_Description = "Action-type property that sorts selected rows according to selected criteria."
Attribute Sort.VB_MemberFlags = "400"
Err.Raise Number:=394, Description:="Property is write-only"
End Property

Public Property Let Sort(ByVal Value As FlexSortConstants)

#If ImplementFlexDataSource Then

If Not VBFlexGridFlexDataSource Is Nothing Then Err.Raise Number:=5, Description:="This functionality is disabled when custom data source is set."

#End If

Select Case Value
    Case FlexSortNone, FlexSortGenericAscending, FlexSortGenericDescending, FlexSortNumericAscending, FlexSortNumericDescending, FlexSortStringNoCaseAscending, FlexSortStringNoCaseDescending, FlexSortStringAscending, FlexSortStringDescending, FlexSortCustom, FlexSortUseColSort, FlexSortCurrencyAscending, FlexSortCurrencyDescending, FlexSortDateAscending, FlexSortDateDescending
        VBFlexGridSort = Value
        If VBFlexGridSort = FlexSortNone Then Exit Property
        If (VBFlexGridRow < 0 Or VBFlexGridRowSel < 0) Or (VBFlexGridCol < 0 Or VBFlexGridColSel < 0) Then
            ' Error shall not be raised. Do nothing in this case.
            Exit Property
        End If
        Dim SelRange As TCELLRANGE, iCol As Long, Sort As FlexSortConstants
        Call GetSelRangeStruct(SelRange)
        ' The keys used for sorting are determined by the Col and ColSel properties.
        ' To specify the range to be sorted, set the Row and RowSel properties.
        ' Sorting is always done in a left-to-right direction. (Technically the sorting is performed from right-to-left)
        For iCol = SelRange.RightCol To SelRange.LeftCol Step -1
            If VBFlexGridSort <> FlexSortUseColSort Then Sort = VBFlexGridSort Else Sort = VBFlexGridColsInfo(iCol).Sort
            ' MergeSort/BubbleSort are used as they are 'stable sort' algorithms.
            If Sort <> FlexSortCustom Then
                ' MergeSort is used for automatic sorting as it is fast and reliable.
                If VBFlexGridRow = VBFlexGridRowSel Then
                    Call MergeSortRec(PropFixedRows, PropRows - 1, iCol, VBFlexGridCells.Rows(), Sort)
                Else
                    Call MergeSortRec(SelRange.TopRow, SelRange.BottomRow, iCol, VBFlexGridCells.Rows(), Sort)
                End If
            Else
                ' BubbleSort is used for custom sorting as row1/row2 for text matrix must be meaningful in the 'Compare' event.
                If VBFlexGridRow = VBFlexGridRowSel Then
                    Call BubbleSortIter(PropFixedRows, PropRows - 1, iCol, VBFlexGridCells.Rows())
                Else
                    Call BubbleSortIter(SelRange.TopRow, SelRange.BottomRow, iCol, VBFlexGridCells.Rows())
                End If
            End If
        Next iCol
        Dim RCP As TROWCOLPARAMS
        With RCP
        .Mask = RCPM_TOPROW
        .Flags = RCPF_CHECKTOPROW Or RCPF_SETSCROLLBARS Or RCPF_FORCEREDRAW
        .TopRow = VBFlexGridTopRow
        End With
        If VBFlexGridIndirectCellRef.InProc = False Then
            Call SetRowColParams(RCP)
        Else
            LSet VBFlexGridIndirectCellRef.RCP = RCP
            VBFlexGridIndirectCellRef.SetRCP = True
        End If
    Case Else
        Err.Raise 380
End Select
' Action-type property. Not real property.
End Property

Public Property Get TabBehavior() As FlexTabBehaviorConstants
Attribute TabBehavior.VB_Description = "Returns/sets a value that defines the behavior of the tab key."
TabBehavior = PropTabBehavior
End Property

Public Property Let TabBehavior(ByVal Value As FlexTabBehaviorConstants)
Select Case Value
    Case FlexTabControls, FlexTabCells, FlexTabNext
        PropTabBehavior = Value
    Case Else
        Err.Raise 380
End Select
UserControl.PropertyChanged "TabBehavior"
End Property

Public Property Get DirectionAfterReturn() As FlexDirectionAfterReturnConstants
Attribute DirectionAfterReturn.VB_Description = "Returns/sets a value that determines the relative position of the next cell when the user presses the return (Enter) key."
DirectionAfterReturn = PropDirectionAfterReturn
End Property

Public Property Let DirectionAfterReturn(ByVal Value As FlexDirectionAfterReturnConstants)
Select Case Value
    Case FlexDirectionAfterReturnNone, FlexDirectionAfterReturnUp, FlexDirectionAfterReturnDown, FlexDirectionAfterReturnLeft, FlexDirectionAfterReturnRight
        PropDirectionAfterReturn = Value
    Case Else
        Err.Raise 380
End Select
UserControl.PropertyChanged "DirectionAfterReturn"
End Property

Public Property Get WrapCellBehavior() As FlexWrapCellBehaviorConstants
Attribute WrapCellBehavior.VB_Description = "Returns/sets a value that determines what the flex grid does when at last or first column in a row."
WrapCellBehavior = PropWrapCellBehavior
End Property

Public Property Let WrapCellBehavior(ByVal Value As FlexWrapCellBehaviorConstants)
Select Case Value
    Case FlexWrapNone, FlexWrapRow, FlexWrapGrid
        PropWrapCellBehavior = Value
    Case Else
        Err.Raise 380
End Select
UserControl.PropertyChanged "WrapCellBehavior"
End Property

Public Property Get ShowInfoTips() As Boolean
Attribute ShowInfoTips.VB_Description = "Returns/sets a value that determines whether the tool tip text properties will be displayed or not."
ShowInfoTips = PropShowInfoTips
End Property

Public Property Let ShowInfoTips(ByVal Value As Boolean)
PropShowInfoTips = Value
If VBFlexGridHandle <> 0 And VBFlexGridDesignMode = False Then
    If PropShowInfoTips = False And PropShowLabelTips = False Then
        Call DestroyToolTip
    Else
        Call CreateToolTip
    End If
End If
UserControl.PropertyChanged "ShowInfoTips"
End Property

Public Property Get ShowLabelTips() As Boolean
Attribute ShowLabelTips.VB_Description = "Returns/sets a value indicating that if a partially hidden label lacks tool tip text, the flex grid will unfold the label or not."
ShowLabelTips = PropShowLabelTips
End Property

Public Property Let ShowLabelTips(ByVal Value As Boolean)
PropShowLabelTips = Value
If VBFlexGridHandle <> 0 And VBFlexGridDesignMode = False Then
    If PropShowInfoTips = False And PropShowLabelTips = False Then
        Call DestroyToolTip
    Else
        Call CreateToolTip
    End If
End If
UserControl.PropertyChanged "ShowLabelTips"
End Property

Public Property Get ClipSeparators() As String
Attribute ClipSeparators.VB_Description = "Returns/sets two distinct characters to be used as column (first) and row (second) separators in clip strings. If it is empty, the defaults vbTab and vbCr are used."
ClipSeparators = PropClipSeparators
End Property

Public Property Let ClipSeparators(ByVal Value As String)
Select Case Len(Value)
    Case Is > 2, 1
        If VBFlexGridDesignMode = True Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
    Case 2
        If StrComp(Left$(Value, 1), Right$(Value, 1)) = 0 Then
            If VBFlexGridDesignMode = True Then
                MsgBox "Invalid property value", vbCritical + vbOKOnly
                Exit Property
            Else
                Err.Raise 380
            End If
        End If
End Select
PropClipSeparators = Value
UserControl.PropertyChanged "ClipSeparators"
End Property

Public Property Get ClipMode() As FlexClipModeConstants
Attribute ClipMode.VB_Description = "Returns/sets a value that determines whether to include or exclude hidden cells when doing a clip command."
ClipMode = PropClipMode
End Property

Public Property Let ClipMode(ByVal Value As FlexClipModeConstants)
Select Case Value
    Case FlexClipModeNormal, FlexClipModeExcludeHidden
        PropClipMode = Value
    Case Else
        Err.Raise 380
End Select
UserControl.PropertyChanged "ClipMode"
End Property

Public Property Get FormatString() As String
Attribute FormatString.VB_Description = "Allows you to set up column widths, alignments, and fixed row and column text in the flex grid at design time."
FormatString = PropFormatString
End Property

Public Property Let FormatString(ByVal Value As String)
PropFormatString = Value
If VBFlexGridDesignMode = True Then
    If PropRows > 0 And PropCols > 0 Then
        Call EraseFlexGridCells
        Call InitFlexGridCells
        Call RedrawGrid
    End If
End If
If Not PropFormatString = vbNullString Then
    If PropFixedRows > 0 Then
        Dim PosRemainder As Long
        Dim FormatCol As String, FormatRow As String
        PosRemainder = InStr(1, PropFormatString, ";")
        If PosRemainder > 0 Then
            FormatCol = Mid$(PropFormatString, 1, PosRemainder - 1)
            FormatRow = Mid$(PropFormatString, PosRemainder + 1)
        Else
            FormatCol = PropFormatString
        End If
        Dim Pos1 As Long, Pos2 As Long, Temp As String, Spacing As Long
        Spacing = (COLINFO_WIDTH_SPACING_DIP * PixelsPerDIP_X())
        If Not FormatCol = vbNullString Then
            Dim iCol As Long
            Do
                Pos1 = InStr(Pos1 + 1, FormatCol, "|")
                Pos2 = Pos1
                iCol = iCol + 1
            Loop Until Pos1 = 0
            If iCol > PropCols Then Me.Cols = iCol
            Pos1 = 0
            Pos2 = 0
            iCol = 0
            Do
                Pos1 = InStr(Pos1 + 1, FormatCol, "|")
                If Pos1 > 0 Then
                    Temp = Mid$(FormatCol, Pos2 + 1, Pos1 - Pos2 - 1)
                Else
                    Temp = Mid$(FormatCol, Pos2 + 1)
                End If
                With VBFlexGridColsInfo(iCol)
                Select Case Left$(Temp, 1)
                    Case "<"
                        .Alignment = FlexAlignmentLeftCenter
                        Temp = Mid$(Temp, 2)
                    Case "^"
                        .Alignment = FlexAlignmentCenterCenter
                        Temp = Mid$(Temp, 2)
                    Case ">"
                        .Alignment = FlexAlignmentRightCenter
                        Temp = Mid$(Temp, 2)
                    Case Else
                        .Alignment = FlexAlignmentGeneral
                End Select
                .Width = GetTextSize(0, iCol, Temp).CX + Spacing
                End With
                Call SetCellText(0, iCol, Trim$(Temp))
                Pos2 = Pos1
                iCol = iCol + 1
            Loop Until Pos1 = 0
            Pos1 = 0
            Pos2 = 0
        End If
        If (Not FormatRow = vbNullString Or PosRemainder > 0) And PropFixedCols > 0 Then
            Dim iRow As Long
            Do
                Pos1 = InStr(Pos1 + 1, FormatRow, "|")
                Pos2 = Pos1
                iRow = iRow + 1
            Loop Until Pos1 = 0
            If iRow > PropRows Then Me.Rows = iRow
            Pos1 = 0
            Pos2 = 0
            iRow = 0
            Do
                Pos1 = InStr(Pos1 + 1, FormatRow, "|")
                If Pos1 > 0 Then
                    Temp = Mid$(FormatRow, Pos2 + 1, Pos1 - Pos2 - 1)
                Else
                    Temp = Mid$(FormatRow, Pos2 + 1)
                End If
                With GetTextSize(iRow, 0, Temp)
                If (.CX + Spacing) > VBFlexGridColsInfo(0).Width Then VBFlexGridColsInfo(0).Width = .CX + Spacing
                End With
                Call SetCellText(iRow, 0, Trim$(Temp))
                Pos2 = Pos1
                iRow = iRow + 1
            Loop Until Pos1 = 0
            Pos1 = 0
            Pos2 = 0
        End If
        Dim RCP As TROWCOLPARAMS
        With RCP
        .Mask = RCPM_LEFTCOL
        .Flags = RCPF_CHECKLEFTCOL Or RCPF_SETSCROLLBARS Or RCPF_FORCEREDRAW
        .LeftCol = VBFlexGridLeftCol
        Call SetRowColParams(RCP)
        End With
    End If
End If
UserControl.PropertyChanged "FormatString"
End Property

Public Property Get IMEMode() As FlexIMEModeConstants
Attribute IMEMode.VB_Description = "Returns/sets the Input Method Editor (IME) mode."
IMEMode = PropIMEMode
End Property

Public Property Let IMEMode(ByVal Value As FlexIMEModeConstants)
Select Case Value
    Case FlexIMEModeNoControl, FlexIMEModeOn, FlexIMEModeOff, FlexIMEModeDisable, FlexIMEModeHiragana, FlexIMEModeKatakana, FlexIMEModeKatakanaHalf, FlexIMEModeAlphaFull, FlexIMEModeAlpha, FlexIMEModeHangulFull, FlexIMEModeHangul
        PropIMEMode = Value
    Case Else
        Err.Raise 380
End Select
If VBFlexGridHandle <> 0 And VBFlexGridDesignMode = False Then
    If GetFocus() = VBFlexGridHandle Then
        Call SetIMEMode(VBFlexGridHandle, VBFlexGridIMCHandle, PropIMEMode)
    ElseIf VBFlexGridEditHandle <> 0 Then
        If GetFocus() = VBFlexGridEditHandle Then Call SetIMEMode(VBFlexGridEditHandle, VBFlexGridIMCHandle, PropIMEMode)
    End If
End If
UserControl.PropertyChanged "IMEMode"
End Property

Public Property Get WantReturn() As Boolean
Attribute WantReturn.VB_Description = "Returns/sets a value that determines when the user presses RETURN to perform the default button or to allow the flex grid to handle the return key. This property applies only when there is any default button on the form."
WantReturn = PropWantReturn
End Property

Public Property Let WantReturn(ByVal Value As Boolean)
If PropWantReturn = Value Then Exit Property
PropWantReturn = Value
If VBFlexGridHandle <> 0 And VBFlexGridDesignMode = False Then
    Dim PropOleObject As OLEGuids.IOleObject
    Dim PropClientSite As OLEGuids.IOleClientSite
    Dim PropUnknown As IUnknown
    Dim PropControlSite As OLEGuids.IOleControlSite
    On Error Resume Next
    Set PropOleObject = Me
    Set PropClientSite = PropOleObject.GetClientSite
    Set PropUnknown = PropClientSite
    Set PropControlSite = PropUnknown
    PropControlSite.OnControlInfoChanged
    If GetFocus() = VBFlexGridHandle Then
        ' If focus is on the control then force the change immediately.
        PropControlSite.OnFocus 1
    End If
    On Error GoTo 0
End If
UserControl.PropertyChanged "WantReturn"
End Property

Private Sub CreateVBFlexGrid()
If VBFlexGridHandle <> 0 Then Exit Sub
Call InitFlexGridCells
If VBFlexGridDesignMode = False Then
    Dim dwStyle As Long, dwExStyle As Long
    dwStyle = WS_CHILD Or WS_VISIBLE Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS
    dwExStyle = WS_EX_NOINHERITLAYOUT
    If PropRightToLeft = True Then
        If PropRightToLeftLayout = True Then
            dwExStyle = dwExStyle Or WS_EX_LAYOUTRTL
        Else
            dwExStyle = dwExStyle Or WS_EX_RTLREADING
        End If
    End If
    Select Case PropBorderStyle
        Case FlexBorderStyleSingle
            dwStyle = dwStyle Or WS_BORDER
        Case FlexBorderStyleThin
            dwExStyle = dwExStyle Or WS_EX_STATICEDGE
        Case FlexBorderStyleSunken
            dwExStyle = dwExStyle Or WS_EX_CLIENTEDGE
        Case FlexBorderStyleRaised
            dwExStyle = dwExStyle Or WS_EX_WINDOWEDGE
            dwStyle = dwStyle Or WS_DLGFRAME
    End Select
    VBFlexGridRTLLayout = CBool((dwExStyle And WS_EX_LAYOUTRTL) = WS_EX_LAYOUTRTL)
    VBFlexGridRTLReading = CBool((dwExStyle And WS_EX_RTLREADING) = WS_EX_RTLREADING)
    VBFlexGridHandle = CreateWindowEx(dwExStyle, StrPtr("VBFlexGridWndClass"), 0, dwStyle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hWnd, 0, App.hInstance, ByVal ObjPtr(Me))
    If VBFlexGridHandle <> 0 Then
        SetWindowLong VBFlexGridHandle, 0, ObjPtr(Me)
        If VBFlexGridIMCHandle = 0 Then
            VBFlexGridIMCHandle = ImmCreateContext()
            If VBFlexGridIMCHandle <> 0 Then ImmAssociateContext VBFlexGridHandle, VBFlexGridIMCHandle
        End If
    End If
    If PropShowInfoTips = True Or PropShowLabelTips = True Then Call CreateToolTip
Else
    VBFlexGridHandle = UserControl.hWnd
    SetRect VBFlexGridClientRect, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    If PropRightToLeft = True Then Me.RightToLeft = True
    Me.BorderStyle = PropBorderStyle
End If
If VBFlexGridHandle <> 0 Then
    VBFlexGridBackColorBrush = CreateSolidBrush(WinColor(PropBackColor))
    VBFlexGridBackColorAltBrush = CreateSolidBrush(WinColor(PropBackColorAlt))
    VBFlexGridBackColorBkgBrush = CreateSolidBrush(WinColor(PropBackColorBkg))
    VBFlexGridBackColorFixedBrush = CreateSolidBrush(WinColor(PropBackColorFixed))
    VBFlexGridBackColorSelBrush = CreateSolidBrush(WinColor(PropBackColorSel))
    Select Case PropGridLines
        Case FlexGridLineDashes
            VBFlexGridPenStyle = PS_DASH
        Case FlexGridLineDots
            VBFlexGridPenStyle = PS_DOT
        Case Else
            VBFlexGridPenStyle = PS_SOLID
    End Select
    VBFlexGridGridLinePen = CreatePen(VBFlexGridPenStyle, PropGridLineWidth, WinColor(PropGridColor))
    Select Case PropGridLinesFixed
        Case FlexGridLineDashes
            VBFlexGridFixedPenStyle = PS_DASH
        Case FlexGridLineDots
            VBFlexGridFixedPenStyle = PS_DOT
        Case Else
            VBFlexGridFixedPenStyle = PS_SOLID
    End Select
    VBFlexGridGridLineFixedPen = CreatePen(VBFlexGridFixedPenStyle, PropGridLineWidth, WinColor(PropGridColorFixed))
    VBFlexGridGridLineWhitePen = CreatePen(PS_SOLID, 0, vbWhite)
    VBFlexGridGridLineBlackPen = CreatePen(PS_SOLID, 0, vbBlack)
End If
Set Me.Font = PropFont
Set Me.FontFixed = PropFontFixed
Me.VisualStyles = PropVisualStyles
Me.Enabled = UserControl.Enabled
If PropRedraw = False Then Me.Redraw = False
Me.FormatString = PropFormatString
Call SetScrollBars
If VBFlexGridDesignMode = False Then
    Call FlexSetSubclass(UserControl.hWnd, Me, 5)
    
    #If ImplementPreTranslateMsg = True Then
    
    If VBFlexGridUsePreTranslateMsg = True Then
        If VBFlexGridHandle <> 0 Then Call FlexPreTranslateMsgAddHook(VBFlexGridHandle)
    End If
    
    #End If
    
End If
UserControl.BackColor = PropBackColorBkg
End Sub

Private Sub CreateToolTip()
Static Done As Boolean
If VBFlexGridToolTipHandle <> 0 Then Exit Sub
If Done = False Then
    Call FlexInitCC(ICC_TAB_CLASSES)
    Done = True
End If
Dim dwExStyle As Long
dwExStyle = WS_EX_TOOLWINDOW Or WS_EX_TOPMOST Or WS_EX_TRANSPARENT
If VBFlexGridRTLLayout = True Then dwExStyle = dwExStyle Or WS_EX_LAYOUTRTL
VBFlexGridToolTipHandle = CreateWindowEx(dwExStyle, StrPtr("tooltips_class32"), StrPtr("Tool Tip"), WS_POPUP Or TTS_ALWAYSTIP Or TTS_NOPREFIX, 0, 0, 0, 0, UserControl.hWnd, 0, App.hInstance, ByVal 0&)
If VBFlexGridToolTipHandle <> 0 Then
    SendMessage VBFlexGridToolTipHandle, TTM_SETMAXTIPWIDTH, 0, ByVal &H7FFF&
    Dim TI As TOOLINFO
    With TI
    .cbSize = LenB(TI)
    .hWnd = VBFlexGridHandle
    .uId = 0
    .uFlags = TTF_SUBCLASS Or TTF_TRANSPARENT Or TTF_PARSELINKS
    If VBFlexGridRTLReading = True Then .uFlags = .uFlags Or TTF_RTLREADING
    .lpszText = LPSTR_TEXTCALLBACK
    LSet .RC = VBFlexGridClientRect
    End With
    SendMessage VBFlexGridToolTipHandle, TTM_ADDTOOL, 0, ByVal VarPtr(TI)
End If
Call SetVisualStylesToolTip
End Sub

Private Sub DestroyVBFlexGrid()
If VBFlexGridHandle = 0 Then Exit Sub
Call FlexRemoveSubclass(UserControl.hWnd)
Call DestroyToolTip
If VBFlexGridDesignMode = False Then
    
    #If ImplementPreTranslateMsg = True Then
    
    If VBFlexGridUsePreTranslateMsg = True Then Call FlexPreTranslateMsgReleaseHook(VBFlexGridHandle)
    
    #End If
    
    SetWindowLong VBFlexGridHandle, 0, 0
    If VBFlexGridIMCHandle <> 0 Then
        ImmAssociateContext VBFlexGridHandle, 0
        ImmDestroyContext VBFlexGridIMCHandle
        VBFlexGridIMCHandle = 0
    End If
    ShowWindow VBFlexGridHandle, SW_HIDE
    SetParent VBFlexGridHandle, 0
    DestroyWindow VBFlexGridHandle
End If
VBFlexGridHandle = 0
If VBFlexGridDoubleBufferDC <> 0 Then
    If VBFlexGridDoubleBufferBmpOld <> 0 Then
        SelectObject VBFlexGridDoubleBufferDC, VBFlexGridDoubleBufferBmpOld
        VBFlexGridDoubleBufferBmpOld = 0
    End If
    If VBFlexGridDoubleBufferBmp <> 0 Then
        DeleteObject VBFlexGridDoubleBufferBmp
        VBFlexGridDoubleBufferBmp = 0
    End If
    DeleteDC VBFlexGridDoubleBufferDC
    VBFlexGridDoubleBufferDC = 0
End If
Call EraseFlexGridCells
If VBFlexGridFontHandle <> 0 Then
    DeleteObject VBFlexGridFontHandle
    VBFlexGridFontHandle = 0
End If
If VBFlexGridFontFixedHandle <> 0 Then
    DeleteObject VBFlexGridFontFixedHandle
    VBFlexGridFontFixedHandle = 0
End If
If VBFlexGridBackColorBrush <> 0 Then
    DeleteObject VBFlexGridBackColorBrush
    VBFlexGridBackColorBrush = 0
End If
If VBFlexGridBackColorAltBrush <> 0 Then
    DeleteObject VBFlexGridBackColorAltBrush
    VBFlexGridBackColorAltBrush = 0
End If
If VBFlexGridBackColorBkgBrush <> 0 Then
    DeleteObject VBFlexGridBackColorBkgBrush
    VBFlexGridBackColorBkgBrush = 0
End If
If VBFlexGridBackColorFixedBrush <> 0 Then
    DeleteObject VBFlexGridBackColorFixedBrush
    VBFlexGridBackColorFixedBrush = 0
End If
If VBFlexGridBackColorSelBrush <> 0 Then
    DeleteObject VBFlexGridBackColorSelBrush
    VBFlexGridBackColorSelBrush = 0
End If
If VBFlexGridGridLinePen <> 0 Then
    DeleteObject VBFlexGridGridLinePen
    VBFlexGridGridLinePen = 0
End If
If VBFlexGridGridLineFixedPen <> 0 Then
    DeleteObject VBFlexGridGridLineFixedPen
    VBFlexGridGridLineFixedPen = 0
End If
If VBFlexGridGridLineWhitePen <> 0 Then
    DeleteObject VBFlexGridGridLineWhitePen
    VBFlexGridGridLineWhitePen = 0
End If
If VBFlexGridGridLineBlackPen <> 0 Then
    DeleteObject VBFlexGridGridLineBlackPen
    VBFlexGridGridLineBlackPen = 0
End If
End Sub

Private Sub DestroyToolTip()
If VBFlexGridToolTipHandle = 0 Then Exit Sub
DestroyWindow VBFlexGridToolTipHandle
VBFlexGridToolTipHandle = 0
VBFlexGridToolTipRow = -1
VBFlexGridToolTipCol = -1
End Sub

Private Function CreateEdit(ByVal Reason As FlexEditReasonConstants, Optional ByVal Row As Long = -1, Optional ByVal Col As Long = -1) As Boolean
Static InProc As Boolean
If VBFlexGridHandle = 0 Or VBFlexGridEditHandle <> 0 Or InProc = True Then Exit Function
If VBFlexGridEditRow > -1 And VBFlexGridEditCol > -1 Then Exit Function
If VBFlexGridDesignMode = True Then Exit Function
If Row = -1 Then Row = VBFlexGridRow
If Col = -1 Then Col = VBFlexGridCol
If (Row < 0 Or Row > (PropRows - 1)) Or (Col < 0 Or Col > (PropCols - 1)) Then Exit Function
If IsWindowEnabled(VBFlexGridHandle) = 0 Then Exit Function
InProc = True
Dim Cancel As Boolean
RaiseEvent BeforeEdit(Row, Col, Reason, Cancel)
If Cancel = True Then
    VBFlexGridEditReason = -1
    InProc = False
    Exit Function
Else
    If (Row >= 0 And Row <= (PropRows - 1)) Then VBFlexGridEditRow = Row Else VBFlexGridEditRow = VBFlexGridRow
    If (Col >= 0 And Col <= (PropCols - 1)) Then VBFlexGridEditCol = Col Else VBFlexGridEditCol = VBFlexGridCol
    VBFlexGridEditReason = Reason
    VBFlexGridEditCloseMode = -1
    VBFlexGridComboActiveMode = FlexComboModeNone
End If
If VBFlexGridFocused = False Then SetFocusAPI UserControl.hWnd
Dim IsFixedCell As Boolean, Text As String
IsFixedCell = CBool(VBFlexGridEditRow < PropFixedRows Or VBFlexGridEditCol < PropFixedCols)
Call GetCellText(VBFlexGridEditRow, VBFlexGridEditCol, Text)
Dim dwStyle As Long, dwExStyle As Long
dwStyle = WS_CHILD
dwExStyle = 0
If VBFlexGridRTLReading = True Or VBFlexGridRTLLayout = True Then dwExStyle = dwExStyle Or WS_EX_RTLREADING Or WS_EX_LEFTSCROLLBAR
Dim Alignment As FlexAlignmentConstants
With VBFlexGridCells.Rows(VBFlexGridEditRow).Cols(VBFlexGridEditCol)
If .Alignment = -1 Then
    If IsFixedCell = False Then
        Alignment = VBFlexGridColsInfo(VBFlexGridEditCol).Alignment
    Else
        Alignment = VBFlexGridColsInfo(VBFlexGridEditCol).FixedAlignment
    End If
Else
    Alignment = .Alignment
End If
End With
Select Case Alignment
    Case FlexAlignmentLeftTop, FlexAlignmentLeftCenter, FlexAlignmentLeftBottom
        If VBFlexGridRTLLayout = False Then dwStyle = dwStyle Or ES_LEFT Else dwStyle = dwStyle Or ES_RIGHT
    Case FlexAlignmentCenterTop, FlexAlignmentCenterCenter, FlexAlignmentCenterBottom
        dwStyle = dwStyle Or ES_CENTER
    Case FlexAlignmentRightTop, FlexAlignmentRightCenter, FlexAlignmentRightBottom
        If VBFlexGridRTLLayout = False Then dwStyle = dwStyle Or ES_RIGHT Else dwStyle = dwStyle Or ES_LEFT
    Case FlexAlignmentGeneral
        If Not IsNumeric(Text) Then
            If VBFlexGridRTLLayout = False Then dwStyle = dwStyle Or ES_LEFT Else dwStyle = dwStyle Or ES_RIGHT
        Else
            If VBFlexGridRTLLayout = False Then dwStyle = dwStyle Or ES_RIGHT Else dwStyle = dwStyle Or ES_LEFT
        End If
End Select
If PropWordWrap = True Or PropSingleLine = False Then
    dwStyle = dwStyle Or ES_MULTILINE Or ES_AUTOVSCROLL
Else
    dwStyle = dwStyle Or ES_AUTOHSCROLL
End If
' Ellipsis format will be ignored.
RaiseEvent EditSetupStyle(dwStyle, dwExStyle)
If (dwStyle And WS_BORDER) = WS_BORDER Then dwStyle = dwStyle And Not WS_BORDER
If (dwStyle And WS_DLGFRAME) = WS_DLGFRAME Then dwStyle = dwStyle And Not WS_DLGFRAME
If (dwExStyle And WS_EX_STATICEDGE) = WS_EX_STATICEDGE Then dwExStyle = dwExStyle And Not WS_EX_STATICEDGE
If (dwExStyle And WS_EX_CLIENTEDGE) = WS_EX_CLIENTEDGE Then dwExStyle = dwExStyle And Not WS_EX_CLIENTEDGE
If (dwExStyle And WS_EX_WINDOWEDGE) = WS_EX_WINDOWEDGE Then dwExStyle = dwExStyle And Not WS_EX_WINDOWEDGE
Dim CellRangeRect As RECT, EditRect As RECT, ComboItems As String, ComboButtonWidth As Long
Call GetMergedRangeStruct(VBFlexGridEditRow, VBFlexGridEditCol, VBFlexGridEditMergedRange)
Me.CellEnsureVisible , VBFlexGridEditMergedRange.TopRow, VBFlexGridEditMergedRange.LeftCol
Call GetCellRangeRect(VBFlexGridEditMergedRange, False, CellRangeRect)
LSet EditRect = CellRangeRect
If EditRect.Bottom > VBFlexGridClientRect.Bottom Then EditRect.Bottom = VBFlexGridClientRect.Bottom
If EditRect.Right > VBFlexGridClientRect.Right Then EditRect.Right = VBFlexGridClientRect.Right
If VBFlexGridComboMode <> FlexComboModeNone Then
    VBFlexGridComboActiveMode = VBFlexGridComboMode
    ComboItems = VBFlexGridComboItems
ElseIf VBFlexGridColsInfo(VBFlexGridEditCol).ComboMode <> FlexComboModeNone Then
    VBFlexGridComboActiveMode = VBFlexGridColsInfo(VBFlexGridEditCol).ComboMode
    ComboItems = VBFlexGridColsInfo(VBFlexGridEditCol).ComboItems
End If
If VBFlexGridComboActiveMode <> FlexComboModeNone Then
    ComboButtonWidth = GetSystemMetrics(SM_CXVSCROLL)
    EditRect.Right = EditRect.Right - ComboButtonWidth
    If (EditRect.Right - 1) < EditRect.Left Then EditRect.Right = (EditRect.Left + 1)
    If (((CellRangeRect.Right - CellRangeRect.Left) - 1) - ComboButtonWidth) < 0 Then ComboButtonWidth = ((CellRangeRect.Right - CellRangeRect.Left) - 1)
    Select Case VBFlexGridComboActiveMode
        Case FlexComboModeDropDown
            If Not (dwStyle And ES_READONLY) = ES_READONLY Then dwStyle = dwStyle Or ES_READONLY
        Case FlexComboModeEditable
            If (dwStyle And ES_READONLY) = ES_READONLY Then dwStyle = dwStyle And Not ES_READONLY
    End Select
End If
VBFlexGridEditHandle = CreateWindowEx(dwExStyle, StrPtr("Edit"), 0, dwStyle, EditRect.Left, EditRect.Top, (EditRect.Right - EditRect.Left) - 1, (EditRect.Bottom - EditRect.Top) - 1, VBFlexGridHandle, ID_EDITCHILD, App.hInstance, ByVal 0&)
If VBFlexGridEditHandle <> 0 Then
    With VBFlexGridCells.Rows(VBFlexGridEditRow).Cols(VBFlexGridEditCol)
    Dim hFont As Long
    If .FontName = vbNullString Then
        If IsFixedCell = False Then
            hFont = VBFlexGridFontHandle
        Else
            If VBFlexGridFontFixedHandle = 0 Then
                hFont = VBFlexGridFontHandle
            Else
                hFont = VBFlexGridFontFixedHandle
            End If
        End If
    Else
        Dim TempFont As StdFont
        Set TempFont = New StdFont
        TempFont.Name = .FontName
        TempFont.Size = .FontSize
        TempFont.Bold = .FontBold
        TempFont.Italic = .FontItalic
        TempFont.Strikethrough = .FontStrikeThrough
        TempFont.Underline = .FontUnderline
        TempFont.Charset = .FontCharset
        VBFlexGridEditTempFontHandle = CreateGDIFontFromOLEFont(TempFont)
        hFont = VBFlexGridEditTempFontHandle
        Set TempFont = Nothing
    End If
    If .BackColor = -1 Then
        If IsFixedCell = False Then
            VBFlexGridEditBackColor = PropBackColor
        Else
            VBFlexGridEditBackColor = PropBackColorFixed
        End If
    Else
        VBFlexGridEditBackColor = .BackColor
    End If
    If .ForeColor = -1 Then
        If IsFixedCell = False Then
            VBFlexGridEditForeColor = PropForeColor
        Else
            VBFlexGridEditForeColor = PropForeColorFixed
        End If
    Else
        VBFlexGridEditForeColor = .ForeColor
    End If
    End With
    SendMessage VBFlexGridEditHandle, WM_SETFONT, hFont, ByVal 0&
    If VBFlexGridRTLLayout = False Then
        SendMessage VBFlexGridEditHandle, EM_SETMARGINS, EC_LEFTMARGIN Or EC_RIGHTMARGIN, ByVal MakeDWord(CELL_TEXT_WIDTH_PADDING_DIP * PixelsPerDIP_X(), (CELL_TEXT_WIDTH_PADDING_DIP * PixelsPerDIP_X()) - 1)
    Else
        SendMessage VBFlexGridEditHandle, EM_SETMARGINS, EC_LEFTMARGIN Or EC_RIGHTMARGIN, ByVal MakeDWord(CELL_TEXT_WIDTH_PADDING_DIP * PixelsPerDIP_X() - 1, (CELL_TEXT_WIDTH_PADDING_DIP * PixelsPerDIP_X()))
    End If
    SendMessage VBFlexGridEditHandle, WM_SETTEXT, 0, ByVal StrPtr(Text)
    VBFlexGridEditTextChanged = False
    VBFlexGridEditAlreadyValidated = False
    SendMessage VBFlexGridEditHandle, EM_SETSEL, 0, ByVal -1&
    If Not (dwStyle And ES_READONLY) = ES_READONLY Then
        If Reason = FlexEditReasonSpace Then
            SendMessage VBFlexGridEditHandle, EM_SETSEL, -1, ByVal -1&
        ElseIf Reason = FlexEditReasonBackSpace Then
            SendMessage VBFlexGridEditHandle, EM_REPLACESEL, 1, ByVal 0&
        End If
    End If
    If ComboButtonWidth > 0 Then
        dwStyle = WS_CHILD Or SS_OWNERDRAW Or SS_NOTIFY
        dwExStyle = 0
        If VBFlexGridRTLReading = True Then dwExStyle = dwExStyle Or WS_EX_RTLREADING
        If VBFlexGridRTLLayout = True Then dwExStyle = dwExStyle Or WS_EX_LAYOUTRTL
        VBFlexGridComboButtonHandle = CreateWindowEx(dwExStyle, StrPtr("Static"), 0, dwStyle, EditRect.Right - 1, EditRect.Top, ComboButtonWidth, (EditRect.Bottom - EditRect.Top) - 1, VBFlexGridHandle, ID_COMBOBUTTONCHILD, App.hInstance, ByVal 0&)
        If VBFlexGridComboButtonHandle <> 0 And VBFlexGridComboActiveMode <> FlexComboModeButton Then
            dwStyle = WS_POPUP Or WS_BORDER Or WS_VSCROLL Or LBS_NOTIFY Or LBS_SORT
            dwExStyle = WS_EX_TOOLWINDOW Or WS_EX_TOPMOST
            If VBFlexGridRTLReading = True Then dwExStyle = dwExStyle Or WS_EX_RTLREADING
            If VBFlexGridRTLLayout = True Then dwExStyle = dwExStyle Or WS_EX_LAYOUTRTL
            SetRect VBFlexGridComboListRect, CellRangeRect.Left, EditRect.Top, CellRangeRect.Right, EditRect.Bottom
            Dim WndRect As RECT
            LSet WndRect = VBFlexGridComboListRect
            MapWindowPoints VBFlexGridHandle, HWND_DESKTOP, WndRect, 2
            VBFlexGridComboListHandle = CreateWindowEx(dwExStyle, StrPtr("ComboLBox"), 0, dwStyle, WndRect.Left, WndRect.Top, WndRect.Right - WndRect.Left, WndRect.Bottom - WndRect.Top, VBFlexGridHandle, 0, App.hInstance, ByVal 0&)
            If VBFlexGridComboListHandle <> 0 Then
                SendMessage VBFlexGridComboListHandle, WM_SETFONT, hFont, ByVal 0&
                Dim Pos1 As Long, Pos2 As Long, Temp As String, i As Long
                Do
                    Pos1 = InStr(Pos1 + 1, ComboItems, "|")
                    If Pos1 > 0 Then
                        Temp = Mid$(ComboItems, Pos2 + 1, Pos1 - Pos2 - 1)
                    Else
                        Temp = Mid$(ComboItems, Pos2 + 1)
                    End If
                    SendMessage VBFlexGridComboListHandle, LB_INSERTSTRING, i, ByVal StrPtr(Temp)
                    Pos2 = Pos1
                    i = i + 1
                Loop Until Pos1 = 0
                Const EDIT_MAXDROPDOWNITEMS As Integer = 9
                Dim Count As Long, Height As Long
                Count = SendMessage(VBFlexGridComboListHandle, LB_GETCOUNT, 0, ByVal 0&)
                Select Case Count
                    Case 0
                        Count = 1
                    Case Is > EDIT_MAXDROPDOWNITEMS
                        Count = EDIT_MAXDROPDOWNITEMS
                End Select
                Height = SendMessage(VBFlexGridComboListHandle, LB_GETITEMHEIGHT, 0, ByVal 0&) * Count
                MoveWindow VBFlexGridComboListHandle, WndRect.Left, WndRect.Top, WndRect.Right - WndRect.Left, Height + 2, 0
                SendMessage VBFlexGridComboListHandle, LB_SETCURSEL, SendMessage(VBFlexGridComboListHandle, LB_FINDSTRINGEXACT, -1, ByVal StrPtr(Text)), ByVal 0&
            End If
        End If
    End If
    If VBFlexGridEnabledVisualStyles = True Then
        If PropVisualStyles = True Then
            ActivateVisualStyles VBFlexGridEditHandle
            If VBFlexGridComboButtonHandle <> 0 Then ActivateVisualStyles VBFlexGridComboButtonHandle
            If VBFlexGridComboListHandle <> 0 Then ActivateVisualStyles VBFlexGridComboListHandle
        Else
            RemoveVisualStyles VBFlexGridEditHandle
            If VBFlexGridComboButtonHandle <> 0 Then RemoveVisualStyles VBFlexGridComboButtonHandle
            If VBFlexGridComboListHandle <> 0 Then RemoveVisualStyles VBFlexGridComboListHandle
        End If
    End If
    Call FlexSetSubclass(VBFlexGridEditHandle, Me, 2)
    If VBFlexGridComboButtonHandle <> 0 Then Call FlexSetSubclass(VBFlexGridComboButtonHandle, Me, 3)
    If VBFlexGridComboListHandle <> 0 Then Call FlexSetSubclass(VBFlexGridComboListHandle, Me, 4)
    SetWindowPos VBFlexGridEditHandle, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
    RaiseEvent EditSetupWindow(VBFlexGridEditBackColor, VBFlexGridEditForeColor)
    VBFlexGridEditBackColorBrush = CreateSolidBrush(WinColor(VBFlexGridEditBackColor))
    ShowWindow VBFlexGridEditHandle, SW_SHOW
    SetFocusAPI VBFlexGridEditHandle
    If VBFlexGridComboButtonHandle <> 0 Then
        ShowWindow VBFlexGridComboButtonHandle, SW_SHOW
        If VBFlexGridComboListHandle <> 0 Then
            If VBFlexGridComboActiveMode = FlexComboModeDropDown Then Call ComboShowDropDown(True)
        End If
    End If
    
    #If ImplementPreTranslateMsg = True Then
    
    If VBFlexGridUsePreTranslateMsg = True Then Call FlexPreTranslateMsgAddHook(VBFlexGridEditHandle)
    
    #End If
    
    RaiseEvent EnterEdit
    CreateEdit = True
End If
InProc = False
End Function

Private Function DestroyEdit(ByVal Discard As Boolean, ByVal CloseMode As FlexEditCloseModeConstants) As Boolean
Static InProc As Boolean
If VBFlexGridEditHandle = 0 Or InProc = True Then Exit Function
Dim Cancel As Boolean
If CloseMode <> FlexEditCloseModeLostFocus Then
    RaiseEvent EditQueryClose(CloseMode, Cancel)
    If Cancel = True Then Exit Function
Else
    If VBFlexGridEditOnValidate = True Or VBFlexGridComboButtonClick = True Then Exit Function
End If
If Discard = False And VBFlexGridEditTextChanged = True Then
    If VBFlexGridEditAlreadyValidated = False Then
        VBFlexGridEditOnValidate = True
        RaiseEvent ValidateEdit(Cancel)
        VBFlexGridEditOnValidate = False
        If VBFlexGridEditHandle = 0 Then
            DestroyEdit = True
            Exit Function
        End If
    Else
        VBFlexGridEditAlreadyValidated = False
    End If
    If Cancel = False Then
        Dim Text As String, iRow As Long, iCol As Long
        Text = Me.EditText
        With VBFlexGridEditMergedRange
        For iRow = .TopRow To .BottomRow
            For iCol = .LeftCol To .RightCol
                Call SetCellText(iRow, iCol, Text)
            Next iCol
        Next iRow
        End With
        Call RedrawGrid
    Else
        InProc = False
        Exit Function
    End If
Else
    VBFlexGridEditAlreadyValidated = False
End If
InProc = True
VBFlexGridEditCloseMode = CloseMode
RaiseEvent LeaveEdit

#If ImplementPreTranslateMsg = True Then

If VBFlexGridUsePreTranslateMsg = True Then Call FlexPreTranslateMsgReleaseHook(VBFlexGridEditHandle)

#End If

Dim Row As Long, Col As Long
' It is necessary to preserve the edit row and col from here on.
' When the edit control has been destroyed it could be started again resulting that the edit row and col will be overwritten.
Row = VBFlexGridEditRow
Col = VBFlexGridEditCol
If VBFlexGridComboButtonHandle <> 0 Then
    Call FlexRemoveSubclass(VBFlexGridComboButtonHandle)
    ShowWindow VBFlexGridComboButtonHandle, SW_HIDE
    SetParent VBFlexGridComboButtonHandle, 0
    DestroyWindow VBFlexGridComboButtonHandle
End If
If VBFlexGridComboListHandle <> 0 Then
    Call FlexRemoveSubclass(VBFlexGridComboListHandle)
    DestroyWindow VBFlexGridComboListHandle
End If
Call FlexRemoveSubclass(VBFlexGridEditHandle)
Dim hWndTemp As Long
' Temporary cache is necessary as the variable needs to be cleared for internal control before the edit window is destroyed.
hWndTemp = VBFlexGridEditHandle
VBFlexGridEditHandle = 0
ShowWindow hWndTemp, SW_HIDE
SetParent hWndTemp, 0
DestroyWindow hWndTemp
hWndTemp = 0
VBFlexGridComboButtonHandle = 0
VBFlexGridComboListHandle = 0
VBFlexGridEditRectChanged = False
If VBFlexGridEditTempFontHandle <> 0 Then
    DeleteObject VBFlexGridEditTempFontHandle
    VBFlexGridEditTempFontHandle = 0
End If
If VBFlexGridEditBackColorBrush <> 0 Then
    DeleteObject VBFlexGridEditBackColorBrush
    VBFlexGridEditBackColorBrush = 0
End If
VBFlexGridEditRow = -1
VBFlexGridEditCol = -1
VBFlexGridComboActiveMode = FlexComboModeNone
If Discard = False And VBFlexGridEditTextChanged = True Then
    RaiseEvent AfterEdit(Row, Col, True)
Else
    RaiseEvent AfterEdit(Row, Col, False)
End If
DestroyEdit = True
InProc = False
End Function

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
UserControl.Refresh
If VBFlexGridNoRedraw = False And VBFlexGridDesignMode = False Then RedrawWindow UserControl.hWnd, 0, 0, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
End Sub

#If ImplementDataSource = True Or ImplementFlexDataSource = True Then

Public Sub DataRefresh()
Attribute DataRefresh.VB_Description = "Forces the control to re-fetch all data from its data source."

#If ImplementDataSource = True Then

If Not PropDataSource Is Nothing Then Set Me.DataSource = PropDataSource

#End If

#If ImplementFlexDataSource = True Then

If Not VBFlexGridFlexDataSource Is Nothing Then Set Me.FlexDataSource = VBFlexGridFlexDataSource

#End If

End Sub

#End If

#If ImplementFlexDataSource = True Then

Public Property Get FlexDataSource() As IVBFlexDataSource
Attribute FlexDataSource.VB_Description = "Returns/sets a custom data source for the control."
Attribute FlexDataSource.VB_MemberFlags = "400"
Set FlexDataSource = VBFlexGridFlexDataSource
End Property

Public Property Let FlexDataSource(ByVal Value As IVBFlexDataSource)
Set Me.FlexDataSource = Value
End Property

Public Property Set FlexDataSource(ByVal Value As IVBFlexDataSource)
Set VBFlexGridFlexDataSource = Value
If Not VBFlexGridFlexDataSource Is Nothing Then
    With VBFlexGridFlexDataSource
    Dim FieldCount As Long, RecordCount As Long, iRow As Long, iCol As Long
    FieldCount = .GetFieldCount
    If FieldCount > 0 Then
        Me.Cols = FieldCount
        If PropFixedRows > 0 Then
            For iCol = 0 To (FieldCount - 1)
                Me.TextMatrix(0, iCol) = .GetFieldName(iCol)
            Next iCol
        End If
        RecordCount = .GetRecordCount
        If RecordCount > 0 Then
            Me.Rows = PropFixedRows + RecordCount
        Else
            Me.Rows = PropFixedRows + 1
        End If
    End If
    End With
Else
    Me.Refresh
End If
End Property

#End If

Public Sub AddItem(ByVal Item As String, Optional ByVal Index As Variant)
Attribute AddItem.VB_Description = "Adds an item to the flex grid."
Dim IndexLong As Long
If IsMissing(Index) = True Then
    IndexLong = PropRows
Else
    Select Case VarType(Index)
        Case vbLong, vbInteger, vbByte
            IndexLong = Index
        Case vbDouble, vbSingle, vbString
            IndexLong = CLng(Index)
        Case Else
            Err.Raise 13
    End Select
End If
If IndexLong > -1 And IndexLong < PropFixedRows Then
    Err.Raise Number:=30001, Description:="Cannot use AddItem on a fixed row"
ElseIf IndexLong < 0 Or IndexLong > PropRows Then
    Err.Raise Number:=30002, Description:="Grid does not contain that row"
Else
    PropRows = PropRows + 1
    ReDim Preserve VBFlexGridCells.Rows(0 To (PropRows - 1)) As TCOLS
    Dim iRow As Long
    If IndexLong < (PropRows - 1) Then
        For iRow = ((PropRows - 1) - 1) To IndexLong Step -1
            LSet VBFlexGridCells.Rows(iRow + 1) = VBFlexGridCells.Rows(iRow)
        Next iRow
    End If
    LSet VBFlexGridCells.Rows(IndexLong) = VBFlexGridDefaultCols
    Dim Pos1 As Long, Pos2 As Long, iCol As Long, ColSeparator As String
    ColSeparator = GetColSeparator()
    If PropClipMode = FlexClipModeNormal Then
        Do
            Pos1 = InStr(Pos1 + 1, Item, ColSeparator)
            If Pos1 > 0 Then
                If iCol < PropCols Then Call SetCellText(IndexLong, iCol, Mid$(Item, Pos2 + 1, Pos1 - Pos2 - 1))
            Else
                If iCol < PropCols Then Call SetCellText(IndexLong, iCol, Mid$(Item, Pos2 + 1))
            End If
            Pos2 = Pos1
            iCol = iCol + 1
        Loop Until Pos1 = 0
    ElseIf PropClipMode = FlexClipModeExcludeHidden Then
        Dim ColLoop As Boolean
        Do
            If (VBFlexGridColsInfo(iCol).State And CLIS_HIDDEN) = 0 Then
                Pos1 = InStr(Pos1 + 1, Item, ColSeparator)
                If Pos1 > 0 Then
                    If iCol < PropCols Then Call SetCellText(IndexLong, iCol, Mid$(Item, Pos2 + 1, Pos1 - Pos2 - 1))
                Else
                    If iCol < PropCols Then Call SetCellText(IndexLong, iCol, Mid$(Item, Pos2 + 1))
                End If
                Pos2 = Pos1
                iCol = iCol + 1
                ColLoop = CBool(Pos1 <> 0 And iCol < PropCols)
            Else
                iCol = iCol + 1
                ColLoop = CBool(iCol < PropCols)
            End If
        Loop Until ColLoop = False
    End If
    Dim RCP As TROWCOLPARAMS
    With RCP
    .Flags = RCPF_SETSCROLLBARS
    Select Case PropSelectionMode
        Case FlexSelectionModeByColumn
            .Mask = .Mask Or RCPM_ROWSEL
            .RowSel = (PropRows - 1)
    End Select
    Call SetRowColParams(RCP)
    End With
End If
End Sub

Public Sub RemoveItem(ByVal Index As Long)
Attribute RemoveItem.VB_Description = "Removes an item from the flex grid."
If Index > -1 And Index < PropFixedRows Then
    Err.Raise Number:=30000, Description:="Cannot do a RemoveItem on a fixed row"
ElseIf Index < 0 Or Index > PropRows Then
    Err.Raise Number:=30002, Description:="Grid does not contain that row"
ElseIf Index = (PropRows - 1) And (PropRows - PropFixedRows) = 1 Then
    Err.Raise Number:=30015, Description:="Can not remove last non-fixed row"
Else
    PropRows = PropRows - 1
    Dim iRow As Long
    If Index < ((PropRows - 1) + 1) Then
        For iRow = Index To (PropRows - 1)
            LSet VBFlexGridCells.Rows(iRow) = VBFlexGridCells.Rows(iRow + 1)
        Next iRow
    End If
    ReDim Preserve VBFlexGridCells.Rows(0 To (PropRows - 1)) As TCOLS
    Dim RCP As TROWCOLPARAMS
    With RCP
    .Flags = RCPF_SETSCROLLBARS
    If VBFlexGridRow > (PropRows - 1) Then
        .Mask = .Mask Or RCPM_ROW
        .Row = (PropRows - 1)
    End If
    Select Case PropSelectionMode
        Case FlexSelectionModeFree, FlexSelectionModeFreeByRow, FlexSelectionModeFreeByColumn, FlexSelectionModeByRow
            If VBFlexGridRowSel > (PropRows - 1) Then
                .Mask = .Mask Or RCPM_ROWSEL
                .RowSel = (PropRows - 1)
            End If
        Case FlexSelectionModeByColumn
            .Mask = .Mask Or RCPM_ROWSEL
            .RowSel = (PropRows - 1)
    End Select
    If VBFlexGridTopRow > (PropRows - 1) Then
        .Mask = .Mask Or RCPM_TOPROW
        .Flags = .Flags Or RCPF_CHECKTOPROW
        .TopRow = (PropRows - 1)
    End If
    Call SetRowColParams(RCP)
    End With
End If
End Sub

Public Sub Clear(Optional ByVal Where As FlexClearWhereConstants, Optional ByVal What As FlexClearWhatConstants)
Attribute Clear.VB_Description = "Clears the contents of the flex grid."
Select Case Where
    Case FlexClearEverywhere, FlexClearFixed, FlexClearScrollable, FlexClearMovable, FlexClearFrozen, FlexClearSelection
    Case Else
        Err.Raise 380
End Select
Select Case What
    Case FlexClearEverything, FlexClearText
        
        #If ImplementFlexDataSource Then
        
        If Not VBFlexGridFlexDataSource Is Nothing Then Err.Raise Number:=5, Description:="This function cannot be used to clear text (only to clear formatting) when custom data source is set."
        
        #End If
        
    Case FlexClearFormatting
    Case Else
        Err.Raise 380
End Select
Dim iRow As Long, iCol As Long, Temp As String
Select Case Where
    Case FlexClearEverywhere
        Select Case What
            Case FlexClearEverything
                ' This is the default setting. It is also the fastest and most efficient approach.
                For iRow = 0 To (PropRows - 1)
                    VBFlexGridCells.Rows(iRow).Cols() = VBFlexGridDefaultCols.Cols()
                Next iRow
            Case FlexClearText
                For iRow = 0 To (PropRows - 1)
                    For iCol = 0 To (PropCols - 1)
                        VBFlexGridCells.Rows(iRow).Cols(iCol).Text = vbNullString
                    Next iCol
                Next iRow
            Case FlexClearFormatting
                For iRow = 0 To (PropRows - 1)
                    For iCol = 0 To (PropCols - 1)
                        Temp = VBFlexGridCells.Rows(iRow).Cols(iCol).Text
                        LSet VBFlexGridCells.Rows(iRow).Cols(iCol) = VBFlexGridDefaultCell
                        VBFlexGridCells.Rows(iRow).Cols(iCol).Text = Temp
                    Next iCol
                Next iRow
        End Select
    Case FlexClearFixed
        Select Case What
            Case FlexClearEverything
                For iRow = 0 To (PropFixedRows - 1)
                    VBFlexGridCells.Rows(iRow).Cols() = VBFlexGridDefaultCols.Cols()
                Next iRow
                For iCol = 0 To (PropFixedCols - 1)
                    For iRow = PropFixedRows To (PropRows - 1)
                        LSet VBFlexGridCells.Rows(iRow).Cols(iCol) = VBFlexGridDefaultCell
                    Next iRow
                Next iCol
            Case FlexClearText
                For iRow = 0 To (PropFixedRows - 1)
                    For iCol = 0 To (PropCols - 1)
                        VBFlexGridCells.Rows(iRow).Cols(iCol).Text = vbNullString
                    Next iCol
                Next iRow
                For iCol = 0 To (PropFixedCols - 1)
                    For iRow = PropFixedRows To (PropRows - 1)
                        VBFlexGridCells.Rows(iRow).Cols(iCol).Text = vbNullString
                    Next iRow
                Next iCol
            Case FlexClearFormatting
                For iRow = 0 To (PropFixedRows - 1)
                    For iCol = 0 To (PropCols - 1)
                        Temp = VBFlexGridCells.Rows(iRow).Cols(iCol).Text
                        LSet VBFlexGridCells.Rows(iRow).Cols(iCol) = VBFlexGridDefaultCell
                        VBFlexGridCells.Rows(iRow).Cols(iCol).Text = Temp
                    Next iCol
                Next iRow
                For iCol = 0 To (PropFixedCols - 1)
                    For iRow = PropFixedRows To (PropRows - 1)
                        Temp = VBFlexGridCells.Rows(iRow).Cols(iCol).Text
                        LSet VBFlexGridCells.Rows(iRow).Cols(iCol) = VBFlexGridDefaultCell
                        VBFlexGridCells.Rows(iRow).Cols(iCol).Text = Temp
                    Next iRow
                Next iCol
        End Select
    Case FlexClearScrollable
        Select Case What
            Case FlexClearEverything
                For iRow = (PropFixedRows + PropFrozenRows) To (PropRows - 1)
                    For iCol = (PropFixedCols + PropFrozenCols) To (PropCols - 1)
                        LSet VBFlexGridCells.Rows(iRow).Cols(iCol) = VBFlexGridDefaultCell
                    Next iCol
                Next iRow
            Case FlexClearText
                For iRow = (PropFixedRows + PropFrozenRows) To (PropRows - 1)
                    For iCol = (PropFixedCols + PropFrozenCols) To (PropCols - 1)
                        VBFlexGridCells.Rows(iRow).Cols(iCol).Text = vbNullString
                    Next iCol
                Next iRow
            Case FlexClearFormatting
                For iRow = (PropFixedRows + PropFrozenRows) To (PropRows - 1)
                    For iCol = (PropFixedCols + PropFrozenCols) To (PropCols - 1)
                        Temp = VBFlexGridCells.Rows(iRow).Cols(iCol).Text
                        LSet VBFlexGridCells.Rows(iRow).Cols(iCol) = VBFlexGridDefaultCell
                        VBFlexGridCells.Rows(iRow).Cols(iCol).Text = Temp
                    Next iCol
                Next iRow
        End Select
    Case FlexClearMovable
        Select Case What
            Case FlexClearEverything
                For iRow = PropFixedRows To (PropRows - 1)
                    For iCol = PropFixedCols To (PropCols - 1)
                        LSet VBFlexGridCells.Rows(iRow).Cols(iCol) = VBFlexGridDefaultCell
                    Next iCol
                Next iRow
            Case FlexClearText
                For iRow = PropFixedRows To (PropRows - 1)
                    For iCol = PropFixedCols To (PropCols - 1)
                        VBFlexGridCells.Rows(iRow).Cols(iCol).Text = vbNullString
                    Next iCol
                Next iRow
            Case FlexClearFormatting
                For iRow = PropFixedRows To (PropRows - 1)
                    For iCol = PropFixedCols To (PropCols - 1)
                        Temp = VBFlexGridCells.Rows(iRow).Cols(iCol).Text
                        LSet VBFlexGridCells.Rows(iRow).Cols(iCol) = VBFlexGridDefaultCell
                        VBFlexGridCells.Rows(iRow).Cols(iCol).Text = Temp
                    Next iCol
                Next iRow
        End Select
    Case FlexClearFrozen
        Select Case What
            Case FlexClearEverything
                For iRow = PropFixedRows To ((PropFixedRows + PropFrozenRows) - 1)
                    For iCol = PropFixedCols To (PropCols - 1)
                        LSet VBFlexGridCells.Rows(iRow).Cols(iCol) = VBFlexGridDefaultCell
                    Next iCol
                Next iRow
                For iCol = PropFixedCols To ((PropFixedCols + PropFrozenCols) - 1)
                    For iRow = (PropFixedRows + PropFrozenRows) To (PropRows - 1)
                        LSet VBFlexGridCells.Rows(iRow).Cols(iCol) = VBFlexGridDefaultCell
                    Next iRow
                Next iCol
            Case FlexClearText
                For iRow = PropFixedRows To ((PropFixedRows + PropFrozenRows) - 1)
                    For iCol = PropFixedCols To (PropCols - 1)
                        VBFlexGridCells.Rows(iRow).Cols(iCol).Text = vbNullString
                    Next iCol
                Next iRow
                For iCol = PropFixedCols To ((PropFixedCols + PropFrozenCols) - 1)
                    For iRow = (PropFixedRows + PropFrozenRows) To (PropRows - 1)
                        VBFlexGridCells.Rows(iRow).Cols(iCol).Text = vbNullString
                    Next iRow
                Next iCol
            Case FlexClearFormatting
                For iRow = PropFixedRows To ((PropFixedRows + PropFrozenRows) - 1)
                    For iCol = PropFixedCols To (PropCols - 1)
                        Temp = VBFlexGridCells.Rows(iRow).Cols(iCol).Text
                        LSet VBFlexGridCells.Rows(iRow).Cols(iCol) = VBFlexGridDefaultCell
                        VBFlexGridCells.Rows(iRow).Cols(iCol).Text = Temp
                    Next iCol
                Next iRow
                For iCol = PropFixedCols To ((PropFixedCols + PropFrozenCols) - 1)
                    For iRow = (PropFixedRows + PropFrozenRows) To (PropRows - 1)
                        Temp = VBFlexGridCells.Rows(iRow).Cols(iCol).Text
                        LSet VBFlexGridCells.Rows(iRow).Cols(iCol) = VBFlexGridDefaultCell
                        VBFlexGridCells.Rows(iRow).Cols(iCol).Text = Temp
                    Next iRow
                Next iCol
        End Select
    Case FlexClearSelection
        Dim SelRange As TCELLRANGE
        Call GetSelRangeStruct(SelRange)
        Select Case What
            Case FlexClearEverything
                For iRow = SelRange.TopRow To SelRange.BottomRow
                    For iCol = SelRange.LeftCol To SelRange.RightCol
                        LSet VBFlexGridCells.Rows(iRow).Cols(iCol) = VBFlexGridDefaultCell
                    Next iCol
                Next iRow
            Case FlexClearText
                For iRow = SelRange.TopRow To SelRange.BottomRow
                    For iCol = SelRange.LeftCol To SelRange.RightCol
                        VBFlexGridCells.Rows(iRow).Cols(iCol).Text = vbNullString
                    Next iCol
                Next iRow
            Case FlexClearFormatting
                For iRow = SelRange.TopRow To SelRange.BottomRow
                    For iCol = SelRange.LeftCol To SelRange.RightCol
                        Temp = VBFlexGridCells.Rows(iRow).Cols(iCol).Text
                        LSet VBFlexGridCells.Rows(iRow).Cols(iCol) = VBFlexGridDefaultCell
                        VBFlexGridCells.Rows(iRow).Cols(iCol).Text = Temp
                    Next iCol
                Next iRow
        End Select
End Select
Call RedrawGrid
End Sub

Public Property Get Row() As Long
Attribute Row.VB_Description = "Returns/sets the active cell in the flex grid."
Attribute Row.VB_MemberFlags = "400"
Row = VBFlexGridRow
End Property

Public Property Let Row(ByVal Value As Long)
If Value < 0 Or Value > (PropRows - 1) Then Err.Raise Number:=30009, Description:="Invalid Row value"
If VBFlexGridCol < 0 Or VBFlexGridCol > (PropCols - 1) Then Err.Raise Number:=30010, Description:="Invalid Col value"
Dim RCP As TROWCOLPARAMS
With RCP
.Mask = RCPM_ROW Or RCPM_ROWSEL Or RCPM_COLSEL
.Row = Value
.RowSel = VBFlexGridRowSel
.ColSel = VBFlexGridColSel
Select Case PropSelectionMode
    Case FlexSelectionModeFree, FlexSelectionModeFreeByRow, FlexSelectionModeFreeByColumn
        .RowSel = .Row
        .ColSel = VBFlexGridCol
    Case FlexSelectionModeByRow
        .RowSel = .Row
        .ColSel = (PropCols - 1)
    Case FlexSelectionModeByColumn
        .RowSel = (PropRows - 1)
        .ColSel = .Col
End Select
Call SetRowColParams(RCP)
End With
End Property

Public Property Get Col() As Long
Attribute Col.VB_Description = "Returns/sets the active cell in the flex grid."
Attribute Col.VB_MemberFlags = "400"
Col = VBFlexGridCol
End Property

Public Property Let Col(ByVal Value As Long)
If VBFlexGridRow < 0 Or VBFlexGridRow > (PropRows - 1) Then Err.Raise Number:=30009, Description:="Invalid Row value"
If Value < 0 Or Value > (PropCols - 1) Then Err.Raise Number:=30010, Description:="Invalid Col value"
Dim RCP As TROWCOLPARAMS
With RCP
.Mask = RCPM_COL Or RCPM_ROWSEL Or RCPM_COLSEL
.Col = Value
.RowSel = VBFlexGridRowSel
.ColSel = VBFlexGridColSel
Select Case PropSelectionMode
    Case FlexSelectionModeFree, FlexSelectionModeFreeByRow, FlexSelectionModeFreeByColumn
        .RowSel = VBFlexGridRow
        .ColSel = .Col
    Case FlexSelectionModeByRow
        .RowSel = VBFlexGridRow
        .ColSel = (PropCols - 1)
    Case FlexSelectionModeByColumn
        .RowSel = (PropRows - 1)
        .ColSel = .Col
End Select
Call SetRowColParams(RCP)
End With
End Property

Public Property Get RowSel() As Long
Attribute RowSel.VB_Description = "Returns/sets the starting or ending row or column for a range of cells."
Attribute RowSel.VB_MemberFlags = "400"
RowSel = VBFlexGridRowSel
End Property

Public Property Let RowSel(ByVal Value As Long)
If Value < 0 Or Value > (PropRows - 1) Then Err.Raise Number:=30009, Description:="Invalid Row value"
Dim RCP As TROWCOLPARAMS
With RCP
.Mask = RCPM_ROWSEL
.RowSel = Value
Call SetRowColParams(RCP)
End With
End Property

Public Property Get ColSel() As Long
Attribute ColSel.VB_Description = "Returns/sets the starting or ending row or column for a range of cells."
Attribute ColSel.VB_MemberFlags = "400"
ColSel = VBFlexGridColSel
End Property

Public Property Let ColSel(ByVal Value As Long)
If Value < 0 Or Value > (PropCols - 1) Then Err.Raise Number:=30010, Description:="Invalid Col value"
Dim RCP As TROWCOLPARAMS
With RCP
.Mask = RCPM_COLSEL
.ColSel = Value
Call SetRowColParams(RCP)
End With
End Property

Public Sub GetSelRange(ByRef Row1 As Long, ByRef Col1 As Long, ByRef Row2 As Long, ByRef Col2 As Long)
Attribute GetSelRange.VB_Description = "Retrieves the starting and ending row and column for a range of cells ordered so that Row1 <= Row2 and Col1 <= Col2."
Dim SelRange As TCELLRANGE
Call GetSelRangeStruct(SelRange)
With SelRange
Row1 = .TopRow
Col1 = .LeftCol
Row2 = .BottomRow
Col2 = .RightCol
End With
End Sub

Public Sub GetMergedRange(ByVal Row As Long, ByVal Col As Long, ByRef Row1 As Long, ByRef Col1 As Long, ByRef Row2 As Long, ByRef Col2 As Long)
Attribute GetMergedRange.VB_Description = "Retrieves the range of merged cells that includes a given cell."
If (Row < 0 Or Row > (PropRows - 1)) Or (Col < 0 Or Col > (PropCols - 1)) Then Err.Raise Number:=381, Description:="Subscript out of range"
Dim MergedRange As TCELLRANGE
Call GetMergedRangeStruct(Row, Col, MergedRange)
With MergedRange
Row1 = .TopRow
Col1 = .LeftCol
Row2 = .BottomRow
Col2 = .RightCol
End With
End Sub

Public Sub SelectRange(ByVal Row As Long, ByVal Col As Long, Optional ByVal RowSel As Long = -1, Optional ByVal ColSel As Long = -1)
Attribute SelectRange.VB_Description = "Selects a range of cells or a cell (by omitting the last two parameters) with a single command."
If RowSel = -1 Then
    If PropSelectionMode <> FlexSelectionModeByColumn Then
        RowSel = Row
    Else
        RowSel = (PropRows - 1)
    End If
End If
If ColSel = -1 Then
    If PropSelectionMode <> FlexSelectionModeByRow Then
        ColSel = Col
    Else
        ColSel = (PropCols - 1)
    End If
End If
If (Row < 0 Or Row > (PropRows - 1)) Or (Col < 0 Or Col > (PropCols - 1)) Or (RowSel < 0 Or RowSel > (PropRows - 1)) Or (ColSel < 0 Or ColSel > (PropCols - 1)) Then Err.Raise Number:=381, Description:="Subscript out of range"
Dim RCP As TROWCOLPARAMS
With RCP
.Mask = RCPM_ROW Or RCPM_COL Or RCPM_ROWSEL Or RCPM_COLSEL
.Row = Row
.Col = Col
.RowSel = RowSel
.ColSel = ColSel
Call SetRowColParams(RCP)
End With
End Sub

Public Property Get TopRow() As Long
Attribute TopRow.VB_Description = "Returns/sets the uppermost row displayed in the flex grid."
Attribute TopRow.VB_MemberFlags = "400"
TopRow = VBFlexGridTopRow
End Property

Public Property Let TopRow(ByVal Value As Long)
If Value < 0 Or Value > (PropRows - 1) Then Err.Raise Number:=30009, Description:="Invalid Row value"
Dim RCP As TROWCOLPARAMS
With RCP
.Mask = RCPM_TOPROW
.Flags = RCPF_CHECKTOPROW
.TopRow = Value
Call SetRowColParams(RCP)
End With
End Property

Public Property Get BottomRow(Optional ByVal Visibility As FlexVisibilityConstants) As Long
Attribute BottomRow.VB_Description = "Returns the bottommost row displayed in the flex grid."
Attribute BottomRow.VB_MemberFlags = "400"
Select Case Visibility
    Case FlexVisibilityPartialOK, FlexVisibilityCompleteOnly
    Case Else
        Err.Raise 380
End Select
If VBFlexGridHandle = 0 Or (PropRows < 1 Or PropCols < 1) Then Exit Property
Dim GridRect As RECT, iRow As Long
With GridRect
For iRow = 0 To (PropFixedRows - 1)
    .Bottom = .Bottom + GetRowHeight(iRow)
Next iRow
BottomRow = VBFlexGridTopRow
For iRow = VBFlexGridTopRow To (PropRows - 1)
    If Visibility = FlexVisibilityPartialOK Then If .Bottom > VBFlexGridClientRect.Bottom Then Exit For
    .Bottom = .Bottom + GetRowHeight(iRow)
    If Visibility = FlexVisibilityCompleteOnly Then If .Bottom > VBFlexGridClientRect.Bottom Then Exit For
    BottomRow = iRow
Next iRow
End With
End Property

Public Property Get LeftCol() As Long
Attribute LeftCol.VB_Description = "Returns/sets the leftmost column displayed in the flex grid."
Attribute LeftCol.VB_MemberFlags = "400"
LeftCol = VBFlexGridLeftCol
End Property

Public Property Let LeftCol(ByVal Value As Long)
If Value < 0 Or Value > (PropCols - 1) Then Err.Raise Number:=30010, Description:="Invalid Col value"
Dim RCP As TROWCOLPARAMS
With RCP
.Mask = RCPM_LEFTCOL
.Flags = RCPF_CHECKLEFTCOL
.LeftCol = Value
Call SetRowColParams(RCP)
End With
End Property

Public Property Get RightCol(Optional ByVal Visibility As FlexVisibilityConstants) As Long
Attribute RightCol.VB_Description = "Returns the rightmost column displayed in the flex grid."
Attribute RightCol.VB_MemberFlags = "400"
Select Case Visibility
    Case FlexVisibilityPartialOK, FlexVisibilityCompleteOnly
    Case Else
        Err.Raise 380
End Select
If VBFlexGridHandle = 0 Or (PropRows < 1 Or PropCols < 1) Then Exit Property
Dim GridRect As RECT, iCol As Long
With GridRect
For iCol = 0 To (PropFixedCols - 1)
    .Right = .Right + GetColWidth(iCol)
Next iCol
RightCol = VBFlexGridLeftCol
For iCol = VBFlexGridLeftCol To (PropCols - 1)
    If Visibility = FlexVisibilityPartialOK Then If .Right > VBFlexGridClientRect.Right Then Exit For
    .Right = .Right + GetColWidth(iCol)
    If Visibility = FlexVisibilityCompleteOnly Then If .Right > VBFlexGridClientRect.Right Then Exit For
    RightCol = iCol
Next iCol
End With
End Property

Public Property Get MouseRow() As Long
Attribute MouseRow.VB_Description = "Returns the row over which the mouse pointer is."
Attribute MouseRow.VB_MemberFlags = "400"
If VBFlexGridHandle <> 0 Then
    Dim P As POINTAPI, HTI As THITTESTINFO
    GetCursorPos P
    ScreenToClient VBFlexGridHandle, P
    HTI.PT.X = P.X
    HTI.PT.Y = P.Y
    Call GetHitTestInfo(HTI)
    MouseRow = HTI.MouseRow
End If
End Property

Public Property Get MouseCol() As Long
Attribute MouseCol.VB_Description = "Returns the column over which the mouse pointer is."
Attribute MouseCol.VB_MemberFlags = "400"
If VBFlexGridHandle <> 0 Then
    Dim P As POINTAPI, HTI As THITTESTINFO
    GetCursorPos P
    ScreenToClient VBFlexGridHandle, P
    HTI.PT.X = P.X
    HTI.PT.Y = P.Y
    Call GetHitTestInfo(HTI)
    MouseCol = HTI.MouseCol
End If
End Property

Public Property Get HitRow() As Long
Attribute HitRow.VB_Description = "Returns the row returned from the last hit test."
Attribute HitRow.VB_MemberFlags = "400"
HitRow = VBFlexGridHitRow
End Property

Public Property Get HitCol() As Long
Attribute HitCol.VB_Description = "Returns the column returned from the last hit test."
Attribute HitCol.VB_MemberFlags = "400"
HitCol = VBFlexGridHitCol
End Property

Public Property Get HitRowDivider() As Long
Attribute HitRowDivider.VB_Description = "Returns the divider row returned from the last hit test."
Attribute HitRowDivider.VB_MemberFlags = "400"
HitRowDivider = VBFlexGridHitRowDivider
End Property

Public Property Get HitColDivider() As Long
Attribute HitColDivider.VB_Description = "Returns the divider column returned from the last hit test."
Attribute HitColDivider.VB_MemberFlags = "400"
HitColDivider = VBFlexGridHitColDivider
End Property

Public Property Get HitResult() As FlexHitResultConstants
Attribute HitResult.VB_Description = "Returns the result returned from the last hit test."
Attribute HitResult.VB_MemberFlags = "400"
HitResult = VBFlexGridHitResult
End Property

Public Property Get RowPos(ByVal Index As Long) As Long
Attribute RowPos.VB_Description = "Returns the distance in twips between the upper-left corner of the control and the upper-left corner of a specified row."
Attribute RowPos.VB_MemberFlags = "400"
If Index < 0 Or Index > (PropRows - 1) Then Err.Raise Number:=30009, Description:="Invalid Row value"
Dim i As Long, Value As Long
If Index > ((PropFixedRows + PropFrozenRows) - 1) Then
    For i = 0 To ((PropFixedRows + PropFrozenRows) - 1)
        If i < Index Then Value = Value + GetRowHeight(i)
    Next i
    For i = VBFlexGridTopRow To (Index - 1)
        Value = Value + GetRowHeight(i)
    Next i
    If Index < VBFlexGridTopRow Then
        For i = (PropFixedRows + PropFrozenRows) To (VBFlexGridTopRow - 1)
            Value = Value - GetRowHeight(i)
        Next i
    End If
Else
    For i = 0 To (Index - 1)
        Value = Value + GetRowHeight(i)
    Next i
End If
RowPos = UserControl.ScaleY(Value, vbPixels, vbTwips)
End Property

Public Property Get RowPosition(ByVal Index As Long) As Long
Attribute RowPosition.VB_Description = "Sets the position of an row, allowing you to move rows to specific positions."
Attribute RowPosition.VB_MemberFlags = "400"
Err.Raise Number:=394, Description:="Property is write-only"
End Property

Public Property Let RowPosition(ByVal Index As Long, ByVal Value As Long)
If (Index < 0 Or Index > (PropRows - 1)) Or (Value < 0 Or Value > (PropRows - 1)) Then Err.Raise Number:=381, Description:="Subscript out of range"
If Index = Value Then Exit Property
Dim Swap As TCOLS
With VBFlexGridCells
LSet Swap = .Rows(Index)
LSet .Rows(Index) = .Rows(Value)
LSet .Rows(Value) = Swap
End With
Dim RCP As TROWCOLPARAMS
With RCP
.Mask = RCPM_TOPROW
.Flags = RCPF_CHECKTOPROW Or RCPF_SETSCROLLBARS Or RCPF_FORCEREDRAW
.TopRow = VBFlexGridTopRow
Call SetRowColParams(RCP)
End With
End Property

Public Property Get RowHeight(ByVal Index As Long) As Long
Attribute RowHeight.VB_Description = "Returns/sets the height in twips of the specified row."
Attribute RowHeight.VB_MemberFlags = "400"
If Index < 0 Or Index > (PropRows - 1) Then Err.Raise Number:=30009, Description:="Invalid Row value"
RowHeight = UserControl.ScaleY(GetRowHeight(Index), vbPixels, vbTwips)
End Property

Public Property Let RowHeight(ByVal Index As Long, ByVal Value As Long)
If Index <> -1 And (Index < 0 Or Index > (PropRows - 1)) Then Err.Raise Number:=30009, Description:="Invalid Row value"
If Value < -1 Then Err.Raise Number:=30013, Description:="Invalid Row Height value"
If Index > -1 Then
    If Value > -1 Then
        VBFlexGridCells.Rows(Index).RowInfo.Height = UserControl.ScaleY(Value, vbTwips, vbPixels)
    Else
        VBFlexGridCells.Rows(Index).RowInfo.Height = -1
    End If
Else
    Dim i As Long
    For i = 0 To (PropRows - 1)
        If Value > -1 Then
            VBFlexGridCells.Rows(i).RowInfo.Height = UserControl.ScaleY(Value, vbTwips, vbPixels)
        Else
            VBFlexGridCells.Rows(i).RowInfo.Height = -1
        End If
    Next i
End If
Dim RCP As TROWCOLPARAMS
With RCP
.Mask = RCPM_TOPROW
.Flags = RCPF_CHECKTOPROW Or RCPF_SETSCROLLBARS Or RCPF_FORCEREDRAW
.TopRow = VBFlexGridTopRow
Call SetRowColParams(RCP)
End With
End Property

Public Property Get RowData(ByVal Index As Long) As Long
Attribute RowData.VB_Description = "Array of long integer values with one item for each row (RowData) and for each column (ColData) of the flex grid."
Attribute RowData.VB_MemberFlags = "400"
If Index < 0 Or Index > (PropRows - 1) Then Err.Raise Number:=30009, Description:="Invalid Row value"
RowData = VBFlexGridCells.Rows(Index).RowInfo.Data
End Property

Public Property Let RowData(ByVal Index As Long, ByVal Value As Long)
If Index < 0 Or Index > (PropRows - 1) Then Err.Raise Number:=30009, Description:="Invalid Row value"
VBFlexGridCells.Rows(Index).RowInfo.Data = Value
End Property

Public Property Get RowHidden(ByVal Index As Long) As Boolean
Attribute RowHidden.VB_Description = "Returns/sets a value indicating if the specified row is hidden."
Attribute RowHidden.VB_MemberFlags = "400"
If Index < 0 Or Index > (PropRows - 1) Then Err.Raise Number:=30009, Description:="Invalid Row value"
RowHidden = CBool((VBFlexGridCells.Rows(Index).RowInfo.State And RWIS_HIDDEN) = RWIS_HIDDEN)
End Property

Public Property Let RowHidden(ByVal Index As Long, ByVal Value As Boolean)
If Index <> -1 And (Index < 0 Or Index > (PropRows - 1)) Then Err.Raise Number:=30009, Description:="Invalid Row value"
If Index > -1 Then
    With VBFlexGridCells.Rows(Index).RowInfo
    If Value = True Then
        If (.State And RWIS_HIDDEN) = 0 Then .State = .State Or RWIS_HIDDEN
    Else
        If (.State And RWIS_HIDDEN) = RWIS_HIDDEN Then .State = .State And Not RWIS_HIDDEN
    End If
    End With
Else
    Dim i As Long
    If Value = True Then
        For i = 0 To (PropRows - 1)
            With VBFlexGridCells.Rows(i).RowInfo
            If (.State And RWIS_HIDDEN) = 0 Then .State = .State Or RWIS_HIDDEN
            End With
        Next i
    Else
        For i = 0 To (PropRows - 1)
            With VBFlexGridCells.Rows(i).RowInfo
            If (.State And RWIS_HIDDEN) = RWIS_HIDDEN Then .State = .State And Not RWIS_HIDDEN
            End With
        Next i
    End If
End If
Dim RCP As TROWCOLPARAMS
With RCP
.Mask = RCPM_TOPROW
.Flags = RCPF_CHECKTOPROW Or RCPF_SETSCROLLBARS Or RCPF_FORCEREDRAW
.TopRow = VBFlexGridTopRow
Call SetRowColParams(RCP)
End With
End Property

Public Property Get RowID(ByVal Index As Long) As Long
Attribute RowID.VB_Description = "Returns/sets an identification used to identify the specified row."
Attribute RowID.VB_MemberFlags = "400"
If Index < 0 Or Index > (PropRows - 1) Then Err.Raise Number:=30009, Description:="Invalid Row value"
RowID = VBFlexGridCells.Rows(Index).RowInfo.ID
End Property

Public Property Let RowID(ByVal Index As Long, ByVal Value As Long)
If Index < 0 Or Index > (PropRows - 1) Then Err.Raise Number:=30009, Description:="Invalid Row value"
VBFlexGridCells.Rows(Index).RowInfo.ID = Value
End Property

Public Property Get RowIndex(ByVal ID As Long) As Long
Attribute RowIndex.VB_Description = "Returns a row index given its identification."
Attribute RowIndex.VB_MemberFlags = "400"
RowIndex = -1
Dim i As Long
With VBFlexGridCells
For i = 0 To (PropRows - 1)
    With .Rows(i).RowInfo
    If .ID = ID And .ID <> 0 Then
        RowIndex = i
        Exit For
    End If
    End With
Next i
End With
End Property

Public Property Let RowIndex(ByVal ID As Long, ByVal Value As Long)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get RowIsVisible(ByVal Index As Long, Optional ByVal Visibility As FlexVisibilityConstants) As Boolean
Attribute RowIsVisible.VB_Description = "Returns a value indicating if the specified row is visible."
Attribute RowIsVisible.VB_MemberFlags = "400"
If Index < 0 Or Index > (PropRows - 1) Then Err.Raise Number:=30009, Description:="Invalid Row value"
Select Case Visibility
    Case FlexVisibilityPartialOK, FlexVisibilityCompleteOnly
    Case Else
        Err.Raise 380
End Select
If VBFlexGridHandle <> 0 Then
    Dim GridRect As RECT, iRow As Long
    With GridRect
    If Index <= ((PropFixedRows + PropFrozenRows) - 1) Then
        RowIsVisible = True
        For iRow = 0 To ((PropFixedRows + PropFrozenRows) - 1)
            If Visibility = FlexVisibilityCompleteOnly Then .Bottom = .Bottom + GetRowHeight(iRow)
            If .Bottom > VBFlexGridClientRect.Bottom Then
                RowIsVisible = False
                Exit For
            End If
            If Visibility = FlexVisibilityPartialOK Then .Bottom = .Bottom + GetRowHeight(iRow)
            If iRow >= Index Then Exit For
        Next iRow
    ElseIf Index >= VBFlexGridTopRow Then
        RowIsVisible = True
        For iRow = 0 To ((PropFixedRows + PropFrozenRows) - 1)
            .Bottom = .Bottom + GetRowHeight(iRow)
        Next iRow
        For iRow = VBFlexGridTopRow To (PropRows - 1)
            If Visibility = FlexVisibilityCompleteOnly Then .Bottom = .Bottom + GetRowHeight(iRow)
            If .Bottom > VBFlexGridClientRect.Bottom Then
                RowIsVisible = False
                Exit For
            End If
            If Visibility = FlexVisibilityPartialOK Then .Bottom = .Bottom + GetRowHeight(iRow)
            If iRow >= Index Then Exit For
        Next iRow
    End If
    End With
End If
End Property

Public Property Get RowsVisible(Optional ByVal Visibility As FlexVisibilityConstants = FlexVisibilityCompleteOnly) As Long
Attribute RowsVisible.VB_Description = "Returns the total number of columns or rows visible in the flex grid."
Attribute RowsVisible.VB_MemberFlags = "400"
Select Case Visibility
    Case FlexVisibilityPartialOK, FlexVisibilityCompleteOnly
    Case Else
        Err.Raise 380
End Select
If VBFlexGridHandle = 0 Or (PropRows < 1 Or PropCols < 1) Then Exit Property
Dim GridRect As RECT, iRow As Long, Count As Long
With GridRect
For iRow = 0 To ((PropFixedRows + PropFrozenRows) - 1)
    If Visibility = FlexVisibilityPartialOK Then If .Bottom > VBFlexGridClientRect.Bottom Then Exit For
    .Bottom = .Bottom + GetRowHeight(iRow)
    If Visibility = FlexVisibilityCompleteOnly Then If .Bottom > VBFlexGridClientRect.Bottom Then Exit For
    Count = Count + 1
Next iRow
For iRow = VBFlexGridTopRow To (PropRows - 1)
    If Visibility = FlexVisibilityPartialOK Then If .Bottom > VBFlexGridClientRect.Bottom Then Exit For
    .Bottom = .Bottom + GetRowHeight(iRow)
    If Visibility = FlexVisibilityCompleteOnly Then If .Bottom > VBFlexGridClientRect.Bottom Then Exit For
    Count = Count + 1
Next iRow
RowsVisible = Count
End With
End Property

Public Property Get FixedRowsVisible(Optional ByVal Visibility As FlexVisibilityConstants = FlexVisibilityCompleteOnly) As Long
Attribute FixedRowsVisible.VB_Description = "Returns the total number of fixed (non-scrollable) columns or rows visible in the flex grid."
Attribute FixedRowsVisible.VB_MemberFlags = "400"
Select Case Visibility
    Case FlexVisibilityPartialOK, FlexVisibilityCompleteOnly
    Case Else
        Err.Raise 380
End Select
If VBFlexGridHandle = 0 Or (PropRows < 1 Or PropCols < 1) Then Exit Property
Dim GridRect As RECT, iRow As Long, Count As Long
With GridRect
For iRow = 0 To (PropFixedRows - 1)
    If Visibility = FlexVisibilityPartialOK Then If .Bottom > VBFlexGridClientRect.Bottom Then Exit For
    .Bottom = .Bottom + GetRowHeight(iRow)
    If Visibility = FlexVisibilityCompleteOnly Then If .Bottom > VBFlexGridClientRect.Bottom Then Exit For
    Count = Count + 1
Next iRow
FixedRowsVisible = Count
End With
End Property

Public Property Get FrozenRowsVisible(Optional ByVal Visibility As FlexVisibilityConstants = FlexVisibilityCompleteOnly) As Long
Attribute FrozenRowsVisible.VB_Description = "Returns the total number of frozen (movable but non-scrollable) columns or rows visible in the flex grid."
Attribute FrozenRowsVisible.VB_MemberFlags = "400"
Select Case Visibility
    Case FlexVisibilityPartialOK, FlexVisibilityCompleteOnly
    Case Else
        Err.Raise 380
End Select
If VBFlexGridHandle = 0 Or (PropRows < 1 Or PropCols < 1) Then Exit Property
Dim GridRect As RECT, iRow As Long, Count As Long
With GridRect
For iRow = PropFixedRows To ((PropFixedRows + PropFrozenRows) - 1)
    If Visibility = FlexVisibilityPartialOK Then If .Bottom > VBFlexGridClientRect.Bottom Then Exit For
    .Bottom = .Bottom + GetRowHeight(iRow)
    If Visibility = FlexVisibilityCompleteOnly Then If .Bottom > VBFlexGridClientRect.Bottom Then Exit For
    Count = Count + 1
Next iRow
FrozenRowsVisible = Count
End With
End Property

Public Property Get RowsPerPage() As Long
Attribute RowsPerPage.VB_Description = "Returns the total number of non-fixed (scrollable) columns or rows displayed on the current page to scroll through in the flex grid."
Attribute RowsPerPage.VB_MemberFlags = "400"
RowsPerPage = GetRowsPerPage(VBFlexGridTopRow)
End Property

Public Property Get ColPos(ByVal Index As Long) As Long
Attribute ColPos.VB_Description = "Returns the distance in twips between the upper-left corner of the control and the upper-left corner of a specified column."
Attribute ColPos.VB_MemberFlags = "400"
If Index < 0 Or Index > (PropCols - 1) Then Err.Raise Number:=30010, Description:="Invalid Col value"
Dim i As Long, Value As Long
If Index > ((PropFixedCols + PropFrozenCols) - 1) Then
    For i = 0 To ((PropFixedCols + PropFrozenCols) - 1)
        If i < Index Then Value = Value + GetColWidth(i)
    Next i
    For i = VBFlexGridLeftCol To (Index - 1)
        Value = Value + GetColWidth(i)
    Next i
    If Index < VBFlexGridLeftCol Then
        For i = (PropFixedCols + PropFrozenCols) To (VBFlexGridLeftCol - 1)
            Value = Value - GetColWidth(i)
        Next i
    End If
Else
    For i = 0 To (Index - 1)
        Value = Value + GetColWidth(i)
    Next i
End If
ColPos = UserControl.ScaleX(Value, vbPixels, vbTwips)
End Property

Public Property Get ColPosition(ByVal Index As Long) As Long
Attribute ColPosition.VB_Description = "Sets the position of an column, allowing you to move columns to specific positions."
Attribute ColPosition.VB_MemberFlags = "400"
Err.Raise Number:=394, Description:="Property is write-only"
End Property

Public Property Let ColPosition(ByVal Index As Long, ByVal Value As Long)
If (Index < 0 Or Index > (PropCols - 1)) Or (Value < 0 Or Value > (PropCols - 1)) Then Err.Raise Number:=381, Description:="Subscript out of range"
If Index = Value Then Exit Property
Dim i As Long, Swap1 As TCELL, Swap2 As TCOLINFO
For i = 0 To (PropRows - 1)
    With VBFlexGridCells.Rows(i)
    LSet Swap1 = .Cols(Index)
    LSet .Cols(Index) = .Cols(Value)
    LSet .Cols(Value) = Swap1
    End With
Next i
LSet Swap2 = VBFlexGridColsInfo(Index)
LSet VBFlexGridColsInfo(Index) = VBFlexGridColsInfo(Value)
LSet VBFlexGridColsInfo(Value) = Swap2
Dim RCP As TROWCOLPARAMS
With RCP
.Mask = RCPM_LEFTCOL
.Flags = RCPF_CHECKLEFTCOL Or RCPF_SETSCROLLBARS Or RCPF_FORCEREDRAW
.LeftCol = VBFlexGridLeftCol
Call SetRowColParams(RCP)
End With
End Property

Public Property Get ColWidth(ByVal Index As Long) As Long
Attribute ColWidth.VB_Description = "Returns/sets the width in twips of the specified column."
Attribute ColWidth.VB_MemberFlags = "400"
If Index < 0 Or Index > (PropCols - 1) Then Err.Raise Number:=30010, Description:="Invalid Col value"
ColWidth = UserControl.ScaleX(GetColWidth(Index), vbPixels, vbTwips)
End Property

Public Property Let ColWidth(ByVal Index As Long, ByVal Value As Long)
If Index <> -1 And (Index < 0 Or Index > (PropCols - 1)) Then Err.Raise Number:=30010, Description:="Invalid Col value"
If Value < -1 Then Err.Raise Number:=30014, Description:="Invalid Col Width value"
If Index > -1 Then
    If Value > -1 Then
        VBFlexGridColsInfo(Index).Width = UserControl.ScaleX(Value, vbTwips, vbPixels)
    Else
        VBFlexGridColsInfo(Index).Width = -1
    End If
Else
    Dim i As Long
    For i = 0 To (PropCols - 1)
        If Value > -1 Then
            VBFlexGridColsInfo(i).Width = UserControl.ScaleX(Value, vbTwips, vbPixels)
        Else
            VBFlexGridColsInfo(i).Width = -1
        End If
    Next i
End If
Dim RCP As TROWCOLPARAMS
With RCP
.Mask = RCPM_LEFTCOL
.Flags = RCPF_CHECKLEFTCOL Or RCPF_SETSCROLLBARS Or RCPF_FORCEREDRAW
.LeftCol = VBFlexGridLeftCol
Call SetRowColParams(RCP)
End With
End Property

Public Property Get ColData(ByVal Index As Long) As Long
Attribute ColData.VB_Description = "Array of long integer values with one item for each row (RowData) and for each column (ColData) of the flex grid."
Attribute ColData.VB_MemberFlags = "400"
If Index < 0 Or Index > (PropCols - 1) Then Err.Raise Number:=30010, Description:="Invalid Col value"
ColData = VBFlexGridColsInfo(Index).Data
End Property

Public Property Let ColData(ByVal Index As Long, ByVal Value As Long)
If Index < 0 Or Index > (PropCols - 1) Then Err.Raise Number:=30010, Description:="Invalid Col value"
VBFlexGridColsInfo(Index).Data = Value
End Property

Public Property Get ColHidden(ByVal Index As Long) As Boolean
Attribute ColHidden.VB_Description = "Returns/sets a value indicating if the specified column is hidden."
Attribute ColHidden.VB_MemberFlags = "400"
If Index < 0 Or Index > (PropCols - 1) Then Err.Raise Number:=30010, Description:="Invalid Col value"
ColHidden = CBool((VBFlexGridColsInfo(Index).State And CLIS_HIDDEN) = CLIS_HIDDEN)
End Property

Public Property Let ColHidden(ByVal Index As Long, ByVal Value As Boolean)
If Index <> -1 And (Index < 0 Or Index > (PropCols - 1)) Then Err.Raise Number:=30010, Description:="Invalid Col value"
If Index > -1 Then
    With VBFlexGridColsInfo(Index)
    If Value = True Then
        If (.State And CLIS_HIDDEN) = 0 Then .State = .State Or CLIS_HIDDEN
    Else
        If (.State And CLIS_HIDDEN) = CLIS_HIDDEN Then .State = .State And Not CLIS_HIDDEN
    End If
    End With
Else
    Dim i As Long
    If Value = True Then
        For i = 0 To (PropCols - 1)
            With VBFlexGridColsInfo(i)
            If (.State And CLIS_HIDDEN) = 0 Then .State = .State Or CLIS_HIDDEN
            End With
        Next i
    Else
        For i = 0 To (PropCols - 1)
            With VBFlexGridColsInfo(i)
            If (.State And CLIS_HIDDEN) = CLIS_HIDDEN Then .State = .State And Not CLIS_HIDDEN
            End With
        Next i
    End If
End If
Dim RCP As TROWCOLPARAMS
With RCP
.Mask = RCPM_LEFTCOL
.Flags = RCPF_CHECKLEFTCOL Or RCPF_SETSCROLLBARS Or RCPF_FORCEREDRAW
.LeftCol = VBFlexGridLeftCol
Call SetRowColParams(RCP)
End With
End Property

Public Property Get ColKey(ByVal Index As Long) As String
Attribute ColKey.VB_Description = "Returns/sets a key used to identify the specified column."
Attribute ColKey.VB_MemberFlags = "400"
If Index < 0 Or Index > (PropCols - 1) Then Err.Raise Number:=30010, Description:="Invalid Col value"
ColKey = VBFlexGridColsInfo(Index).Key
End Property

Public Property Let ColKey(ByVal Index As Long, ByVal Value As String)
If Index < 0 Or Index > (PropCols - 1) Then Err.Raise Number:=30010, Description:="Invalid Col value"
VBFlexGridColsInfo(Index).Key = Value
End Property

Public Property Get ColIndex(ByVal Key As String) As Long
Attribute ColIndex.VB_Description = "Returns a column index given its key."
Attribute ColIndex.VB_MemberFlags = "400"
ColIndex = -1
Dim i As Long
For i = 0 To (PropCols - 1)
    If Not VBFlexGridColsInfo(i).Key = vbNullString Then
        If StrComp(VBFlexGridColsInfo(i).Key, Key, vbTextCompare) = 0 Then
            ColIndex = i
            Exit For
        End If
    End If
Next i
End Property

Public Property Let ColIndex(ByVal Key As String, ByVal Value As Long)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get ColIsVisible(ByVal Index As Long, Optional ByVal Visibility As FlexVisibilityConstants) As Boolean
Attribute ColIsVisible.VB_Description = "Returns a value indicating if the specified column is visible."
Attribute ColIsVisible.VB_MemberFlags = "400"
If Index < 0 Or Index > (PropCols - 1) Then Err.Raise Number:=30010, Description:="Invalid Col value"
Select Case Visibility
    Case FlexVisibilityPartialOK, FlexVisibilityCompleteOnly
    Case Else
        Err.Raise 380
End Select
If VBFlexGridHandle <> 0 Then
    Dim GridRect As RECT, iCol As Long
    With GridRect
    If Index <= ((PropFixedCols + PropFrozenCols) - 1) Then
        ColIsVisible = True
        For iCol = 0 To ((PropFixedCols + PropFrozenCols) - 1)
            If Visibility = FlexVisibilityCompleteOnly Then .Right = .Right + GetColWidth(iCol)
            If .Right > VBFlexGridClientRect.Right Then
                ColIsVisible = False
                Exit For
            End If
            If Visibility = FlexVisibilityPartialOK Then .Right = .Right + GetColWidth(iCol)
            If iCol >= Index Then Exit For
        Next iCol
    ElseIf Index >= VBFlexGridLeftCol Then
        ColIsVisible = True
        For iCol = 0 To ((PropFixedCols + PropFrozenCols) - 1)
            .Right = .Right + GetColWidth(iCol)
        Next iCol
        For iCol = VBFlexGridLeftCol To (PropCols - 1)
            If Visibility = FlexVisibilityCompleteOnly Then .Right = .Right + GetColWidth(iCol)
            If .Right > VBFlexGridClientRect.Right Then
                ColIsVisible = False
                Exit For
            End If
            If Visibility = FlexVisibilityPartialOK Then .Right = .Right + GetColWidth(iCol)
            If iCol >= Index Then Exit For
        Next iCol
    End If
    End With
End If
End Property

Public Property Get ColsVisible(Optional ByVal Visibility As FlexVisibilityConstants = FlexVisibilityCompleteOnly) As Long
Attribute ColsVisible.VB_Description = "Returns the total number of columns or rows visible in the flex grid."
Attribute ColsVisible.VB_MemberFlags = "400"
Select Case Visibility
    Case FlexVisibilityPartialOK, FlexVisibilityCompleteOnly
    Case Else
        Err.Raise 380
End Select
If VBFlexGridHandle = 0 Or (PropRows < 1 Or PropCols < 1) Then Exit Property
Dim GridRect As RECT, iCol As Long, Count As Long
With GridRect
For iCol = 0 To ((PropFixedCols + PropFrozenCols) - 1)
    If Visibility = FlexVisibilityPartialOK Then If .Right > VBFlexGridClientRect.Right Then Exit For
    .Right = .Right + GetColWidth(iCol)
    If Visibility = FlexVisibilityCompleteOnly Then If .Right > VBFlexGridClientRect.Right Then Exit For
    Count = Count + 1
Next iCol
For iCol = VBFlexGridLeftCol To (PropCols - 1)
    If Visibility = FlexVisibilityPartialOK Then If .Right > VBFlexGridClientRect.Right Then Exit For
    .Right = .Right + GetColWidth(iCol)
    If Visibility = FlexVisibilityCompleteOnly Then If .Right > VBFlexGridClientRect.Right Then Exit For
    Count = Count + 1
Next iCol
ColsVisible = Count
End With
End Property

Public Property Get FixedColsVisible(Optional ByVal Visibility As FlexVisibilityConstants = FlexVisibilityCompleteOnly) As Long
Attribute FixedColsVisible.VB_Description = "Returns the total number of fixed (non-scrollable) columns or rows visible in the flex grid."
Attribute FixedColsVisible.VB_MemberFlags = "400"
Select Case Visibility
    Case FlexVisibilityPartialOK, FlexVisibilityCompleteOnly
    Case Else
        Err.Raise 380
End Select
If VBFlexGridHandle = 0 Or (PropRows < 1 Or PropCols < 1) Then Exit Property
Dim GridRect As RECT, iCol As Long, Count As Long
With GridRect
For iCol = 0 To (PropFixedCols - 1)
    If Visibility = FlexVisibilityPartialOK Then If .Right > VBFlexGridClientRect.Right Then Exit For
    .Right = .Right + GetColWidth(iCol)
    If Visibility = FlexVisibilityCompleteOnly Then If .Right > VBFlexGridClientRect.Right Then Exit For
    Count = Count + 1
Next iCol
FixedColsVisible = Count
End With
End Property

Public Property Get FrozenColsVisible(Optional ByVal Visibility As FlexVisibilityConstants = FlexVisibilityCompleteOnly) As Long
Attribute FrozenColsVisible.VB_Description = "Returns the total number of frozen (movable but non-scrollable) columns or rows visible in the flex grid."
Attribute FrozenColsVisible.VB_MemberFlags = "400"
Select Case Visibility
    Case FlexVisibilityPartialOK, FlexVisibilityCompleteOnly
    Case Else
        Err.Raise 380
End Select
If VBFlexGridHandle = 0 Or (PropRows < 1 Or PropCols < 1) Then Exit Property
Dim GridRect As RECT, iCol As Long, Count As Long
With GridRect
For iCol = PropFixedCols To ((PropFixedCols + PropFrozenCols) - 1)
    If Visibility = FlexVisibilityPartialOK Then If .Right > VBFlexGridClientRect.Right Then Exit For
    .Right = .Right + GetColWidth(iCol)
    If Visibility = FlexVisibilityCompleteOnly Then If .Right > VBFlexGridClientRect.Right Then Exit For
    Count = Count + 1
Next iCol
FrozenColsVisible = Count
End With
End Property

Public Property Get ColsPerPage() As Long
Attribute ColsPerPage.VB_Description = "Returns the total number of non-fixed (scrollable) columns or rows displayed on the current page to scroll through in the flex grid."
Attribute ColsPerPage.VB_MemberFlags = "400"
ColsPerPage = GetColsPerPage(VBFlexGridLeftCol)
End Property

Public Property Get ColAlignment(ByVal Index As Long) As FlexAlignmentConstants
Attribute ColAlignment.VB_Description = "Returns/sets the alignment of data in a column. Indirectly available at design time through the format string property."
Attribute ColAlignment.VB_MemberFlags = "400"
If Index < 0 Or Index > (PropCols - 1) Then Err.Raise Number:=30004, Description:="Invalid Col value for alignment"
ColAlignment = VBFlexGridColsInfo(Index).Alignment
End Property

Public Property Let ColAlignment(ByVal Index As Long, ByVal Value As FlexAlignmentConstants)
If Index <> -1 And (Index < 0 Or Index > (PropCols - 1)) Then Err.Raise Number:=30004, Description:="Invalid Col value for alignment"
Select Case Value
    Case FlexAlignmentLeftTop, FlexAlignmentLeftCenter, FlexAlignmentLeftBottom, FlexAlignmentCenterTop, FlexAlignmentCenterCenter, FlexAlignmentCenterBottom, FlexAlignmentRightTop, FlexAlignmentRightCenter, FlexAlignmentRightBottom, FlexAlignmentGeneral
    Case Else
        Err.Raise Number:=30005, Description:="Invalid Alignment value"
End Select
If Index > -1 Then
    VBFlexGridColsInfo(Index).Alignment = Value
Else
    Dim i As Long
    For i = 0 To (PropCols - 1)
        VBFlexGridColsInfo(i).Alignment = Value
    Next i
End If
Call RedrawGrid
End Property

Public Property Get FixedAlignment(ByVal Index As Long) As FlexAlignmentConstants
Attribute FixedAlignment.VB_Description = "Returns/sets the alignment of data in the fixed cells of a column."
Attribute FixedAlignment.VB_MemberFlags = "400"
If Index < 0 Or Index > (PropCols - 1) Then Err.Raise Number:=30004, Description:="Invalid Col value for alignment"
If VBFlexGridColsInfo(Index).FixedAlignment = -1 Then
    FixedAlignment = VBFlexGridColsInfo(Index).Alignment
Else
    FixedAlignment = VBFlexGridColsInfo(Index).FixedAlignment
End If
End Property

Public Property Let FixedAlignment(ByVal Index As Long, ByVal Value As FlexAlignmentConstants)
If Index <> -1 And (Index < 0 Or Index > (PropCols - 1)) Then Err.Raise Number:=30004, Description:="Invalid Col value for alignment"
Select Case Value
    Case -1, FlexAlignmentLeftTop, FlexAlignmentLeftCenter, FlexAlignmentLeftBottom, FlexAlignmentCenterTop, FlexAlignmentCenterCenter, FlexAlignmentCenterBottom, FlexAlignmentRightTop, FlexAlignmentRightCenter, FlexAlignmentRightBottom, FlexAlignmentGeneral
    Case Else
        Err.Raise Number:=30005, Description:="Invalid Alignment value"
End Select
If Index > -1 Then
    VBFlexGridColsInfo(Index).FixedAlignment = Value
Else
    Dim i As Long
    For i = 0 To (PropCols - 1)
        VBFlexGridColsInfo(i).FixedAlignment = Value
    Next i
End If
Call RedrawGrid
End Property

Public Property Get ColSort(ByVal Index As Long) As FlexSortConstants
Attribute ColSort.VB_Description = "Returns/sets the sorting order for the specified column. In order to perform the sort using the different sorting orders for each column, set the sort property to 'UseColSort'."
Attribute ColSort.VB_MemberFlags = "400"
If Index < 0 Or Index > (PropCols - 1) Then Err.Raise Number:=30010, Description:="Invalid Col value"
ColSort = VBFlexGridColsInfo(Index).Sort
End Property

Public Property Let ColSort(ByVal Index As Long, ByVal Value As FlexSortConstants)
If Index <> -1 And (Index < 0 Or Index > (PropCols - 1)) Then Err.Raise Number:=30010, Description:="Invalid Col value"
Select Case Value
    Case FlexSortNone, FlexSortGenericAscending, FlexSortGenericDescending, FlexSortNumericAscending, FlexSortNumericDescending, FlexSortStringNoCaseAscending, FlexSortStringNoCaseDescending, FlexSortStringAscending, FlexSortStringDescending, FlexSortCustom, FlexSortCurrencyAscending, FlexSortCurrencyDescending, FlexSortDateAscending, FlexSortDateDescending
    Case Else
        Err.Raise 380
End Select
If Index > -1 Then
    VBFlexGridColsInfo(Index).Sort = Value
Else
    Dim i As Long
    For i = 0 To (PropCols - 1)
        VBFlexGridColsInfo(i).Sort = Value
    Next i
End If
End Property

Public Property Get ColComboMode(ByVal Index As Long) As FlexComboModeConstants
Attribute ColComboMode.VB_Description = "Returns/sets the combo functionality mode when editing a cell for the specified column."
Attribute ColComboMode.VB_MemberFlags = "400"
If Index < 0 Or Index > (PropCols - 1) Then Err.Raise Number:=30010, Description:="Invalid Col value"
ColComboMode = VBFlexGridColsInfo(Index).ComboMode
End Property

Public Property Let ColComboMode(ByVal Index As Long, ByVal Value As FlexComboModeConstants)
If Index <> -1 And (Index < 0 Or Index > (PropCols - 1)) Then Err.Raise Number:=30010, Description:="Invalid Col value"
Select Case Value
    Case FlexComboModeNone, FlexComboModeDropDown, FlexComboModeEditable, FlexComboModeButton
    Case Else
        Err.Raise 380
End Select
If Index > -1 Then
    VBFlexGridColsInfo(Index).ComboMode = Value
Else
    Dim i As Long
    For i = 0 To (PropCols - 1)
        VBFlexGridColsInfo(i).ComboMode = Value
    Next i
End If
End Property

Public Property Get ColComboItems(ByVal Index As Long) As String
Attribute ColComboItems.VB_Description = "Returns/sets the items to be used for the drop-down list when editing a cell for the specified column."
Attribute ColComboItems.VB_MemberFlags = "400"
If Index < 0 Or Index > (PropCols - 1) Then Err.Raise Number:=30010, Description:="Invalid Col value"
ColComboItems = VBFlexGridColsInfo(Index).ComboItems
End Property

Public Property Let ColComboItems(ByVal Index As Long, ByVal Value As String)
If Index <> -1 And (Index < 0 Or Index > (PropCols - 1)) Then Err.Raise Number:=30010, Description:="Invalid Col value"
If Index > -1 Then
    VBFlexGridColsInfo(Index).ComboItems = Value
Else
    Dim i As Long
    For i = 0 To (PropCols - 1)
        VBFlexGridColsInfo(i).ComboItems = Value
    Next i
End If
End Property

Public Property Get ColFormat(ByVal Index As Long) As String
Attribute ColFormat.VB_Description = "Returns/sets the format used to display numeric values."
Attribute ColFormat.VB_MemberFlags = "400"
If Index < 0 Or Index > (PropCols - 1) Then Err.Raise Number:=30010, Description:="Invalid Col value"
ColFormat = VBFlexGridColsInfo(Index).Format
End Property

Public Property Let ColFormat(ByVal Index As Long, ByVal Value As String)
If Index <> -1 And (Index < 0 Or Index > (PropCols - 1)) Then Err.Raise Number:=30010, Description:="Invalid Col value"
If Index > -1 Then
    VBFlexGridColsInfo(Index).Format = Value
Else
    Dim i As Long
    For i = 0 To (PropCols - 1)
        VBFlexGridColsInfo(i).Format = Value
    Next i
End If
Call RedrawGrid
End Property

Public Property Get MergeRow(ByVal Index As Long) As Boolean
Attribute MergeRow.VB_Description = "Returns/sets which columns or rows should have their contents merged when the merge cells property is set to a value other than 0 - Never."
Attribute MergeRow.VB_MemberFlags = "400"
If Index < 0 Or Index > (PropRows - 1) Then Err.Raise Number:=30009, Description:="Invalid Row value"
MergeRow = CBool((VBFlexGridCells.Rows(Index).RowInfo.State And RWIS_MERGE) = RWIS_MERGE)
End Property

Public Property Let MergeRow(ByVal Index As Long, ByVal Value As Boolean)
If Index < 0 Or Index > (PropRows - 1) Then Err.Raise Number:=30009, Description:="Invalid Row value"
With VBFlexGridCells.Rows(Index).RowInfo
If Value = True Then
    If (.State And RWIS_MERGE) = 0 Then .State = .State Or RWIS_MERGE
Else
    If (.State And RWIS_MERGE) = RWIS_MERGE Then .State = .State And Not RWIS_MERGE
End If
End With
Call RedrawGrid
End Property

Public Property Get MergeCol(ByVal Index As Long) As Boolean
Attribute MergeCol.VB_Description = "Returns/sets which columns or rows should have their contents merged when the merge cells property is set to a value other than 0 - Never."
Attribute MergeCol.VB_MemberFlags = "400"
If Index < 0 Or Index > (PropCols - 1) Then Err.Raise Number:=30010, Description:="Invalid Col value"
MergeCol = CBool((VBFlexGridColsInfo(Index).State And CLIS_MERGE) = CLIS_MERGE)
End Property

Public Property Let MergeCol(ByVal Index As Long, ByVal Value As Boolean)
If Index < 0 Or Index > (PropCols - 1) Then Err.Raise Number:=30010, Description:="Invalid Col value"
With VBFlexGridColsInfo(Index)
If Value = True Then
    If (.State And CLIS_MERGE) = 0 Then .State = .State Or CLIS_MERGE
Else
    If (.State And CLIS_MERGE) = CLIS_MERGE Then .State = .State And Not CLIS_MERGE
End If
End With
Call RedrawGrid
End Property

Public Property Get Cell(ByVal Setting As FlexCellSettings, Optional ByVal Row As Long = -1, Optional ByVal Col As Long = -1, Optional ByVal RowSel As Long = -1, Optional ByVal ColSel As Long = -1) As Variant
Attribute Cell.VB_Description = "Returns/sets cell settings for an arbitrary cell or range of cells."
Attribute Cell.VB_MemberFlags = "400"
If (Row < -1 Or Row > (PropRows - 1)) Or (Col < -1 Or Col > (PropCols - 1)) Or (RowSel < -1 Or RowSel > (PropRows - 1)) Or (ColSel < -1 Or ColSel > (PropCols - 1)) Then Err.Raise Number:=381, Description:="Subscript out of range"
Dim OldRow As Long, OldCol As Long, OldRowSel As Long, OldColSel As Long
OldRow = VBFlexGridRow
OldCol = VBFlexGridCol
OldRowSel = VBFlexGridRowSel
OldColSel = VBFlexGridColSel
If Row > -1 Then VBFlexGridRow = Row
If Col > -1 Then VBFlexGridCol = Col
If RowSel > -1 Then VBFlexGridRowSel = RowSel Else VBFlexGridRowSel = VBFlexGridRow
If ColSel > -1 Then VBFlexGridColSel = ColSel Else VBFlexGridColSel = VBFlexGridCol
On Error GoTo Cancel
Select Case Setting
    Case FlexCellText
        Cell = Me.Text
    Case FlexCellClip
        Cell = Me.Clip
    Case FlexCellTextStyle
        Cell = Me.CellTextStyle
    Case FlexCellAlignment
        Cell = Me.CellAlignment
    Case FlexCellPicture
        Set Cell = Me.CellPicture
    Case FlexCellPictureAlignment
        Cell = Me.CellPictureAlignment
    Case FlexCellBackColor
        Cell = Me.CellBackColor
    Case FlexCellForeColor
        Cell = Me.CellForeColor
    Case FlexCellToolTipText
        Cell = Me.CellToolTipText
    Case FlexCellFontName
        Cell = Me.CellFontName
    Case FlexCellFontSize
        Cell = Me.CellFontSize
    Case FlexCellFontBold
        Cell = Me.CellFontBold
    Case FlexCellFontItalic
        Cell = Me.CellFontItalic
    Case FlexCellFontStrikeThrough
        Cell = Me.CellFontStrikeThrough
    Case FlexCellFontUnderline
        Cell = Me.CellFontUnderline
    Case FlexCellFontCharset
        Cell = Me.CellFontCharset
    Case FlexCellLeft
        Cell = Me.CellLeft
    Case FlexCellTop
        Cell = Me.CellTop
    Case FlexCellWidth
        Cell = Me.CellWidth
    Case FlexCellHeight
        Cell = Me.CellHeight
    Case FlexCellSort
        Err.Raise Number:=394, Description:="Property is write-only"
    Case Else
        Err.Raise 380
End Select
Cancel:
VBFlexGridRow = OldRow
VBFlexGridCol = OldCol
VBFlexGridRowSel = OldRowSel
VBFlexGridColSel = OldColSel
If Err.Number <> 0 Then Err.Raise Number:=Err.Number, Description:=Err.Description
End Property

Public Property Let Cell(ByVal Setting As FlexCellSettings, Optional ByVal Row As Long = -1, Optional ByVal Col As Long = -1, Optional ByVal RowSel As Long = -1, Optional ByVal ColSel As Long = -1, ByVal Value As Variant)
If (Row < -1 Or Row > (PropRows - 1)) Or (Col < -1 Or Col > (PropCols - 1)) Or (RowSel < -1 Or RowSel > (PropRows - 1)) Or (ColSel < -1 Or ColSel > (PropCols - 1)) Then Err.Raise Number:=381, Description:="Subscript out of range"
Dim OldRow As Long, OldCol As Long, OldRowSel As Long, OldColSel As Long, OldNoRedraw As Boolean
OldRow = VBFlexGridRow
OldCol = VBFlexGridCol
OldRowSel = VBFlexGridRowSel
OldColSel = VBFlexGridColSel
OldNoRedraw = VBFlexGridNoRedraw
If Row > -1 Then VBFlexGridRow = Row
If Col > -1 Then VBFlexGridCol = Col
If RowSel > -1 Then VBFlexGridRowSel = RowSel Else VBFlexGridRowSel = VBFlexGridRow
If ColSel > -1 Then VBFlexGridColSel = ColSel Else VBFlexGridColSel = VBFlexGridCol
VBFlexGridNoRedraw = True
VBFlexGridIndirectCellRef.InProc = True
VBFlexGridIndirectCellRef.SetRCP = False
On Error GoTo Cancel
Select Case Setting
    Case FlexCellText
        Me.Text = Value
    Case FlexCellClip
        Me.Clip = Value
    Case FlexCellTextStyle
        Me.CellTextStyle = Value
    Case FlexCellAlignment
        Me.CellAlignment = Value
    Case FlexCellPicture
        Me.CellPicture = Value
    Case FlexCellPictureAlignment
        Me.CellPictureAlignment = Value
    Case FlexCellBackColor
        Me.CellBackColor = Value
    Case FlexCellForeColor
        Me.CellForeColor = Value
    Case FlexCellToolTipText
        Me.CellToolTipText = Value
    Case FlexCellFontName
        Me.CellFontName = Value
    Case FlexCellFontSize
        Me.CellFontSize = Value
    Case FlexCellFontBold
        Me.CellFontBold = Value
    Case FlexCellFontItalic
        Me.CellFontItalic = Value
    Case FlexCellFontStrikeThrough
        Me.CellFontStrikeThrough = Value
    Case FlexCellFontUnderline
        Me.CellFontUnderline = Value
    Case FlexCellFontCharset
        Me.CellFontCharset = Value
    Case FlexCellLeft
        Err.Raise Number:=383, Description:="Property is read-only"
    Case FlexCellTop
        Err.Raise Number:=383, Description:="Property is read-only"
    Case FlexCellWidth
        Err.Raise Number:=383, Description:="Property is read-only"
    Case FlexCellHeight
        Err.Raise Number:=383, Description:="Property is read-only"
    Case FlexCellSort
        Me.Sort = Value
    Case Else
        Err.Raise 380
End Select
Cancel:
VBFlexGridRow = OldRow
VBFlexGridCol = OldCol
VBFlexGridRowSel = OldRowSel
VBFlexGridColSel = OldColSel
VBFlexGridNoRedraw = OldNoRedraw
VBFlexGridIndirectCellRef.InProc = False
If Err.Number = 0 Then
    If VBFlexGridIndirectCellRef.SetRCP = False Then
        Call RedrawGrid
    Else
        Dim RCP As TROWCOLPARAMS
        LSet RCP = VBFlexGridIndirectCellRef.RCP
        VBFlexGridIndirectCellRef.SetRCP = False
        Call SetRowColParams(RCP)
    End If
Else
    Err.Raise Number:=Err.Number, Description:=Err.Description
End If
End Property

Public Property Set Cell(ByVal Setting As FlexCellSettings, Optional ByVal Row As Long = -1, Optional ByVal Col As Long = -1, Optional ByVal RowSel As Long = -1, Optional ByVal ColSel As Long = -1, ByVal Value As Variant)
If (Row < -1 Or Row > (PropRows - 1)) Or (Col < -1 Or Col > (PropCols - 1)) Or (RowSel < -1 Or RowSel > (PropRows - 1)) Or (ColSel < -1 Or ColSel > (PropCols - 1)) Then Err.Raise Number:=381, Description:="Subscript out of range"
Dim OldRow As Long, OldCol As Long, OldRowSel As Long, OldColSel As Long, OldNoRedraw As Boolean
OldRow = VBFlexGridRow
OldCol = VBFlexGridCol
OldRowSel = VBFlexGridRowSel
OldColSel = VBFlexGridColSel
OldNoRedraw = VBFlexGridNoRedraw
If Row > -1 Then VBFlexGridRow = Row
If Col > -1 Then VBFlexGridCol = Col
If RowSel > -1 Then VBFlexGridRowSel = RowSel Else VBFlexGridRowSel = VBFlexGridRow
If ColSel > -1 Then VBFlexGridColSel = ColSel Else VBFlexGridColSel = VBFlexGridCol
VBFlexGridNoRedraw = True
On Error GoTo Cancel
Select Case Setting
    Case FlexCellPicture
        Set Me.CellPicture = Value
    Case Else
        Err.Raise 380
End Select
Cancel:
VBFlexGridRow = OldRow
VBFlexGridCol = OldCol
VBFlexGridRowSel = OldRowSel
VBFlexGridColSel = OldColSel
VBFlexGridNoRedraw = OldNoRedraw
If Err.Number = 0 Then
    Call RedrawGrid
Else
    Err.Raise Number:=Err.Number, Description:=Err.Description
End If
End Property

Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contents of a cell or range of cells."
Attribute Text.VB_UserMemId = 0
Attribute Text.VB_MemberFlags = "400"
If VBFlexGridRow > -1 And VBFlexGridCol > -1 Then Call GetCellText(VBFlexGridRow, VBFlexGridCol, Text)
End Property

Public Property Let Text(ByVal Value As String)
If VBFlexGridRow < 0 Then
    Err.Raise Number:=30009, Description:="Invalid Row value"
ElseIf VBFlexGridCol < 0 Then
    Err.Raise Number:=30010, Description:="Invalid Col value"
End If
If PropFillStyle = FlexFillStyleSingle Then
    Call SetCellText(VBFlexGridRow, VBFlexGridCol, Value)
ElseIf PropFillStyle = FlexFillStyleRepeat Then
    Dim i As Long, j As Long, SelRange As TCELLRANGE
    Call GetSelRangeStruct(SelRange)
    For i = SelRange.TopRow To SelRange.BottomRow
        For j = SelRange.LeftCol To SelRange.RightCol
            Call SetCellText(i, j, Value)
        Next j
    Next i
End If
Call RedrawGrid
End Property

Public Property Get TextArray(ByVal Index As Long) As String
Attribute TextArray.VB_Description = "Returns/sets the text contents of an arbitrary cell (single subscript)."
Attribute TextArray.VB_MemberFlags = "400"
If (Index < 0 Or Index > ((PropRows * PropCols) - 1)) Then Err.Raise Number:=381, Description:="Subscript out of range"
Dim RetVal As Double
RetVal = Index / PropCols
Call GetCellText(Fix(RetVal), ((RetVal - Fix(RetVal)) * PropCols), TextArray)
End Property

Public Property Let TextArray(ByVal Index As Long, ByVal Value As String)
If (Index < 0 Or Index > ((PropRows * PropCols) - 1)) Then Err.Raise Number:=381, Description:="Subscript out of range"
Dim RetVal As Double
RetVal = Index / PropCols
Call SetCellText(Fix(RetVal), ((RetVal - Fix(RetVal)) * PropCols), Value)
Call RedrawGrid
End Property

Public Property Get TextMatrix(ByVal Row As Long, ByVal Col As Long) As String
Attribute TextMatrix.VB_Description = "Returns/sets the text contents of an arbitrary cell (row/col subscripts)."
Attribute TextMatrix.VB_MemberFlags = "400"
If (Row < 0 Or Row > (PropRows - 1)) Or (Col < 0 Or Col > (PropCols - 1)) Then Err.Raise Number:=381, Description:="Subscript out of range"
Call GetCellText(Row, Col, TextMatrix)
End Property

Public Property Let TextMatrix(ByVal Row As Long, ByVal Col As Long, ByVal Value As String)
If (Row < 0 Or Row > (PropRows - 1)) Or (Col < 0 Or Col > (PropCols - 1)) Then Err.Raise Number:=381, Description:="Subscript out of range"
Call SetCellText(Row, Col, Value)
Call RedrawGrid
End Property

Public Property Get Clip() As String
Attribute Clip.VB_Description = "Returns/sets the contents of the cells in a selected region."
Attribute Clip.VB_MemberFlags = "400"
If VBFlexGridRow < 0 Or VBFlexGridCol < 0 Then Err.Raise 7
Dim iRow As Long, iCol As Long, SelRange As TCELLRANGE, Buffer As String
Dim UBoundRows As Long, UBoundCols As Long
Dim StrArr() As String, StrSize As Long
Dim ColSeparator As String, RowSeparator As String
Call GetSelRangeStruct(SelRange)
If PropClipMode = FlexClipModeNormal Then
    UBoundRows = (SelRange.BottomRow - SelRange.TopRow)
    UBoundCols = (SelRange.RightCol - SelRange.LeftCol)
ElseIf PropClipMode = FlexClipModeExcludeHidden Then
    UBoundRows = -1
    UBoundCols = -1
    For iRow = SelRange.TopRow To SelRange.BottomRow
        If (VBFlexGridCells.Rows(iRow).RowInfo.State And RWIS_HIDDEN) = 0 Then UBoundRows = UBoundRows + 1
    Next iRow
    For iCol = SelRange.LeftCol To SelRange.RightCol
        If (VBFlexGridColsInfo(iCol).State And CLIS_HIDDEN) = 0 Then UBoundCols = UBoundCols + 1
    Next iCol
End If
If UBoundRows > -1 And UBoundCols > -1 Then ReDim StrArr(0 To UBoundRows, 0 To UBoundCols) As String
ColSeparator = GetColSeparator()
RowSeparator = GetRowSeparator()
If PropClipMode = FlexClipModeNormal Then
    For iRow = SelRange.TopRow To SelRange.BottomRow
        For iCol = SelRange.LeftCol To SelRange.RightCol
            Call GetCellText(iRow, iCol, Buffer)
            If iCol < SelRange.RightCol Then
                StrArr(iRow + (0 - SelRange.TopRow), iCol + (0 - SelRange.LeftCol)) = Buffer & ColSeparator
            ElseIf iRow < SelRange.BottomRow Then
                StrArr(iRow + (0 - SelRange.TopRow), iCol + (0 - SelRange.LeftCol)) = Buffer & RowSeparator
            Else
                StrArr(iRow + (0 - SelRange.TopRow), iCol + (0 - SelRange.LeftCol)) = Buffer
            End If
        Next iCol
    Next iRow
ElseIf PropClipMode = FlexClipModeExcludeHidden Then
    ' Adjust bottom row and right col so the separators are placed correctly.
    For iRow = SelRange.BottomRow To SelRange.TopRow Step -1
        If (VBFlexGridCells.Rows(iRow).RowInfo.State And RWIS_HIDDEN) = RWIS_HIDDEN Then SelRange.BottomRow = SelRange.BottomRow - 1 Else Exit For
    Next iRow
    For iCol = SelRange.RightCol To SelRange.LeftCol Step -1
        If (VBFlexGridColsInfo(iCol).State And CLIS_HIDDEN) = CLIS_HIDDEN Then SelRange.RightCol = SelRange.RightCol - 1 Else Exit For
    Next iCol
    Dim ArrRowAdj As Long, ArrColAdj As Long
    For iRow = SelRange.TopRow To SelRange.BottomRow
        If (VBFlexGridCells.Rows(iRow).RowInfo.State And RWIS_HIDDEN) = 0 Then
            For iCol = SelRange.LeftCol To SelRange.RightCol
                If (VBFlexGridColsInfo(iCol).State And CLIS_HIDDEN) = 0 Then
                    Call GetCellText(iRow, iCol, Buffer)
                    If iCol < SelRange.RightCol Then
                        StrArr(iRow + (0 - SelRange.TopRow) - ArrRowAdj, iCol + (0 - SelRange.LeftCol) - ArrColAdj) = Buffer & ColSeparator
                    ElseIf iRow < SelRange.BottomRow Then
                        StrArr(iRow + (0 - SelRange.TopRow) - ArrRowAdj, iCol + (0 - SelRange.LeftCol) - ArrColAdj) = Buffer & RowSeparator
                    Else
                        StrArr(iRow + (0 - SelRange.TopRow) - ArrRowAdj, iCol + (0 - SelRange.LeftCol) - ArrColAdj) = Buffer
                    End If
                Else
                    ArrColAdj = ArrColAdj + 1
                End If
            Next iCol
        Else
            ArrRowAdj = ArrRowAdj + 1
        End If
        ArrColAdj = 0
    Next iRow
End If
For iRow = 0 To UBoundRows
    For iCol = 0 To UBoundCols
        StrSize = StrSize + Len(StrArr(iRow, iCol))
    Next iCol
Next iRow
If StrSize > 0 Then
    Clip = String$(StrSize, vbNullChar)
    StrSize = 1
    For iRow = 0 To UBoundRows
        For iCol = 0 To UBoundCols
            If StrSize <= Len(Clip) Then Mid$(Clip, StrSize, Len(StrArr(iRow, iCol))) = StrArr(iRow, iCol)
            StrSize = StrSize + Len(StrArr(iRow, iCol))
        Next iCol
    Next iRow
End If
End Property

Public Property Let Clip(ByVal Value As String)
If VBFlexGridRow < 0 Then
    Err.Raise Number:=30009, Description:="Invalid Row value"
ElseIf VBFlexGridCol < 0 Then
    Err.Raise Number:=30010, Description:="Invalid Col value"
End If
Dim SelRange As TCELLRANGE, Temp As String, iRow As Long, iCol As Long
Dim Pos1 As Long, Pos2 As Long, Pos3 As Long, Pos4 As Long
Dim ColSeparator As String, RowSeparator As String
Call GetSelRangeStruct(SelRange)
ColSeparator = GetColSeparator()
RowSeparator = GetRowSeparator()
With VBFlexGridCells
If PropClipMode = FlexClipModeNormal Then
    Do
        Pos1 = InStr(Pos1 + 1, Value, RowSeparator)
        If (SelRange.TopRow + iRow) <= SelRange.BottomRow Then
            If Pos1 > 0 Then Temp = Mid$(Value, Pos2 + 1, Pos1 - Pos2 - 1) Else Temp = Mid$(Value, Pos2 + 1)
            Do
                Pos3 = InStr(Pos3 + 1, Temp, ColSeparator)
                If Pos3 > 0 Then
                    If (SelRange.LeftCol + iCol) <= SelRange.RightCol Then Call SetCellText(SelRange.TopRow + iRow, SelRange.LeftCol + iCol, Mid$(Temp, Pos4 + 1, Pos3 - Pos4 - 1))
                Else
                    If (SelRange.LeftCol + iCol) <= SelRange.RightCol Then Call SetCellText(SelRange.TopRow + iRow, SelRange.LeftCol + iCol, Mid$(Temp, Pos4 + 1))
                End If
                Pos4 = Pos3
                iCol = iCol + 1
            Loop Until Pos3 = 0
        End If
        Pos2 = Pos1
        Pos4 = 0
        iRow = iRow + 1
        iCol = 0
    Loop Until Pos1 = 0
ElseIf PropClipMode = FlexClipModeExcludeHidden Then
    Dim RowLoop As Boolean, ColLoop As Boolean
    Do
        If (.Rows(SelRange.TopRow + iRow).RowInfo.State And RWIS_HIDDEN) = 0 Then
            Pos1 = InStr(Pos1 + 1, Value, RowSeparator)
            If (SelRange.TopRow + iRow) <= SelRange.BottomRow Then
                If Pos1 > 0 Then Temp = Mid$(Value, Pos2 + 1, Pos1 - Pos2 - 1) Else Temp = Mid$(Value, Pos2 + 1)
                Do
                    If (VBFlexGridColsInfo(SelRange.LeftCol + iCol).State And CLIS_HIDDEN) = 0 Then
                        Pos3 = InStr(Pos3 + 1, Temp, ColSeparator)
                        If Pos3 > 0 Then
                            If (SelRange.LeftCol + iCol) <= SelRange.RightCol Then Call SetCellText(SelRange.TopRow + iRow, SelRange.LeftCol + iCol, Mid$(Temp, Pos4 + 1, Pos3 - Pos4 - 1))
                        Else
                            If (SelRange.LeftCol + iCol) <= SelRange.RightCol Then Call SetCellText(SelRange.TopRow + iRow, SelRange.LeftCol + iCol, Mid$(Temp, Pos4 + 1))
                        End If
                        Pos4 = Pos3
                        iCol = iCol + 1
                        ColLoop = CBool(Pos3 <> 0 And (SelRange.LeftCol + iCol) <= SelRange.RightCol)
                    Else
                        iCol = iCol + 1
                        ColLoop = CBool((SelRange.LeftCol + iCol) <= SelRange.RightCol)
                    End If
                Loop Until ColLoop = False
            End If
            Pos2 = Pos1
            Pos4 = 0
            iRow = iRow + 1
            iCol = 0
            RowLoop = CBool(Pos1 <> 0 And (SelRange.TopRow + iRow) <= SelRange.BottomRow)
        Else
            iRow = iRow + 1
            iCol = 0
            RowLoop = CBool((SelRange.TopRow + iRow) <= SelRange.BottomRow)
        End If
    Loop Until RowLoop = False
End If
End With
Call RedrawGrid
End Property

Public Property Get CellTextStyle() As FlexTextStyleConstants
Attribute CellTextStyle.VB_Description = "Returns/sets 3D effects for text on a specific cell or range of cells."
Attribute CellTextStyle.VB_MemberFlags = "400"
If VBFlexGridRow < 0 Then
    Err.Raise Number:=30009, Description:="Invalid Row value"
ElseIf VBFlexGridCol < 0 Then
    Err.Raise Number:=30010, Description:="Invalid Col value"
End If
If VBFlexGridCells.Rows(VBFlexGridRow).Cols(VBFlexGridCol).TextStyle = -1 Then
    If VBFlexGridRow > (PropFixedRows - 1) And VBFlexGridCol > (PropFixedCols - 1) Then
        CellTextStyle = PropTextStyle
    Else
        CellTextStyle = PropTextStyleFixed
    End If
Else
    CellTextStyle = VBFlexGridCells.Rows(VBFlexGridRow).Cols(VBFlexGridCol).TextStyle
End If
End Property

Public Property Let CellTextStyle(ByVal Value As FlexTextStyleConstants)
If VBFlexGridRow < 0 Then
    Err.Raise Number:=30009, Description:="Invalid Row value"
ElseIf VBFlexGridCol < 0 Then
    Err.Raise Number:=30010, Description:="Invalid Col value"
End If
Select Case Value
    Case -1, FlexTextStyleFlat, FlexTextStyleRaised, FlexTextStyleInset, FlexTextStyleRaisedLight, FlexTextStyleInsetLight
    Case Else
        Err.Raise 380
End Select
If PropFillStyle = FlexFillStyleSingle Then
    VBFlexGridCells.Rows(VBFlexGridRow).Cols(VBFlexGridCol).TextStyle = Value
ElseIf PropFillStyle = FlexFillStyleRepeat Then
    Dim i As Long, j As Long, SelRange As TCELLRANGE
    Call GetSelRangeStruct(SelRange)
    For i = SelRange.TopRow To SelRange.BottomRow
        With VBFlexGridCells.Rows(i)
        For j = SelRange.LeftCol To SelRange.RightCol
            .Cols(j).TextStyle = Value
        Next j
        End With
    Next i
End If
Call RedrawGrid
End Property

Public Property Get CellAlignment() As FlexAlignmentConstants
Attribute CellAlignment.VB_Description = "Returns/sets the alignment of data in a cell or range of selected cells."
Attribute CellAlignment.VB_MemberFlags = "400"
If VBFlexGridRow < 0 Then
    Err.Raise Number:=30009, Description:="Invalid Row value"
ElseIf VBFlexGridCol < 0 Then
    Err.Raise Number:=30010, Description:="Invalid Col value"
End If
If VBFlexGridCells.Rows(VBFlexGridRow).Cols(VBFlexGridCol).Alignment = -1 Then
    If VBFlexGridRow > (PropFixedRows - 1) And VBFlexGridCol > (PropFixedCols - 1) Then
        CellAlignment = VBFlexGridColsInfo(VBFlexGridCol).Alignment
    Else
        If VBFlexGridColsInfo(VBFlexGridCol).FixedAlignment = -1 Then
            CellAlignment = VBFlexGridColsInfo(VBFlexGridCol).Alignment
        Else
            CellAlignment = VBFlexGridColsInfo(VBFlexGridCol).FixedAlignment
        End If
    End If
Else
    CellAlignment = VBFlexGridCells.Rows(VBFlexGridRow).Cols(VBFlexGridCol).Alignment
End If
End Property

Public Property Let CellAlignment(ByVal Value As FlexAlignmentConstants)
If VBFlexGridRow < 0 Then
    Err.Raise Number:=30009, Description:="Invalid Row value"
ElseIf VBFlexGridCol < 0 Then
    Err.Raise Number:=30010, Description:="Invalid Col value"
End If
Select Case Value
    Case -1, FlexAlignmentLeftTop, FlexAlignmentLeftCenter, FlexAlignmentLeftBottom, FlexAlignmentCenterTop, FlexAlignmentCenterCenter, FlexAlignmentCenterBottom, FlexAlignmentRightTop, FlexAlignmentRightCenter, FlexAlignmentRightBottom, FlexAlignmentGeneral
    Case Else
        Err.Raise Number:=30005, Description:="Invalid Alignment value"
End Select
If PropFillStyle = FlexFillStyleSingle Then
    VBFlexGridCells.Rows(VBFlexGridRow).Cols(VBFlexGridCol).Alignment = Value
ElseIf PropFillStyle = FlexFillStyleRepeat Then
    Dim i As Long, j As Long, SelRange As TCELLRANGE
    Call GetSelRangeStruct(SelRange)
    For i = SelRange.TopRow To SelRange.BottomRow
        With VBFlexGridCells.Rows(i)
        For j = SelRange.LeftCol To SelRange.RightCol
            .Cols(j).Alignment = Value
        Next j
        End With
    Next i
End If
Call RedrawGrid
End Property

Public Property Get CellPicture() As IPictureDisp
Attribute CellPicture.VB_Description = "Returns/sets an picture to be displayed in the current cell or in a range of cells."
Attribute CellPicture.VB_MemberFlags = "400"
If VBFlexGridRow < 0 Then
    Err.Raise Number:=30009, Description:="Invalid Row value"
ElseIf VBFlexGridCol < 0 Then
    Err.Raise Number:=30010, Description:="Invalid Col value"
End If
Set CellPicture = VBFlexGridCells.Rows(VBFlexGridRow).Cols(VBFlexGridCol).Picture
End Property

Public Property Let CellPicture(ByVal Value As IPictureDisp)
Set Me.CellPicture = Value
End Property

Public Property Set CellPicture(ByVal Value As IPictureDisp)
If VBFlexGridRow < 0 Then
    Err.Raise Number:=30009, Description:="Invalid Row value"
ElseIf VBFlexGridCol < 0 Then
    Err.Raise Number:=30010, Description:="Invalid Col value"
End If
Set UserControl.Picture = Value
If PropFillStyle = FlexFillStyleSingle Then
    With VBFlexGridCells.Rows(VBFlexGridRow).Cols(VBFlexGridCol)
    Set .Picture = UserControl.Picture
    .PictureRenderFlag = 0
    End With
ElseIf PropFillStyle = FlexFillStyleRepeat Then
    Dim i As Long, j As Long, SelRange As TCELLRANGE
    Call GetSelRangeStruct(SelRange)
    For i = SelRange.TopRow To SelRange.BottomRow
        With VBFlexGridCells.Rows(i)
        For j = SelRange.LeftCol To SelRange.RightCol
            Set .Cols(j).Picture = UserControl.Picture
            .Cols(j).PictureRenderFlag = 0
        Next j
        End With
    Next i
End If
Set UserControl.Picture = Nothing
Call RedrawGrid
End Property

Public Property Get CellPictureAlignment() As FlexPictureAlignmentConstants
Attribute CellPictureAlignment.VB_Description = "Returns/sets the alignment of pictures in a cell or range of selected cells."
Attribute CellPictureAlignment.VB_MemberFlags = "400"
If VBFlexGridRow < 0 Then
    Err.Raise Number:=30009, Description:="Invalid Row value"
ElseIf VBFlexGridCol < 0 Then
    Err.Raise Number:=30010, Description:="Invalid Col value"
End If
CellPictureAlignment = VBFlexGridCells.Rows(VBFlexGridRow).Cols(VBFlexGridCol).PictureAlignment
End Property

Public Property Let CellPictureAlignment(ByVal Value As FlexPictureAlignmentConstants)
If VBFlexGridRow < 0 Then
    Err.Raise Number:=30009, Description:="Invalid Row value"
ElseIf VBFlexGridCol < 0 Then
    Err.Raise Number:=30010, Description:="Invalid Col value"
End If
Select Case Value
    Case FlexPictureAlignmentLeftTop, FlexPictureAlignmentLeftCenter, FlexPictureAlignmentLeftBottom, FlexPictureAlignmentCenterTop, FlexPictureAlignmentCenterCenter, FlexPictureAlignmentCenterBottom, FlexPictureAlignmentRightTop, FlexPictureAlignmentRightCenter, FlexPictureAlignmentRightBottom, FlexPictureAlignmentStretch, FlexPictureAlignmentTile, FlexPictureAlignmentLeftTopNoOverlap, FlexPictureAlignmentLeftCenterNoOverlap, FlexPictureAlignmentLeftBottomNoOverlap, FlexPictureAlignmentRightTopNoOverlap, FlexPictureAlignmentRightCenterNoOverlap, FlexPictureAlignmentRightBottomNoOverlap
    Case Else
        Err.Raise Number:=30005, Description:="Invalid Alignment value"
End Select
If PropFillStyle = FlexFillStyleSingle Then
    VBFlexGridCells.Rows(VBFlexGridRow).Cols(VBFlexGridCol).PictureAlignment = Value
ElseIf PropFillStyle = FlexFillStyleRepeat Then
    Dim i As Long, j As Long, SelRange As TCELLRANGE
    Call GetSelRangeStruct(SelRange)
    For i = SelRange.TopRow To SelRange.BottomRow
        With VBFlexGridCells.Rows(i)
        For j = SelRange.LeftCol To SelRange.RightCol
            .Cols(j).PictureAlignment = Value
        Next j
        End With
    Next i
End If
Call RedrawGrid
End Property

Public Property Get CellBackColor() As Long
Attribute CellBackColor.VB_Description = "Returns/sets the background and foreground colors of individual cells or ranges of cells."
Attribute CellBackColor.VB_MemberFlags = "400"
If VBFlexGridRow < 0 Then
    Err.Raise Number:=30009, Description:="Invalid Row value"
ElseIf VBFlexGridCol < 0 Then
    Err.Raise Number:=30010, Description:="Invalid Col value"
End If
If VBFlexGridCells.Rows(VBFlexGridRow).Cols(VBFlexGridCol).BackColor = -1 Then
    If VBFlexGridRow > (PropFixedRows - 1) And VBFlexGridCol > (PropFixedCols - 1) Then
        CellBackColor = PropBackColor
    Else
        CellBackColor = PropBackColorFixed
    End If
Else
    CellBackColor = VBFlexGridCells.Rows(VBFlexGridRow).Cols(VBFlexGridCol).BackColor
End If
End Property

Public Property Let CellBackColor(ByVal Value As Long)
If VBFlexGridRow < 0 Then
    Err.Raise Number:=30009, Description:="Invalid Row value"
ElseIf VBFlexGridCol < 0 Then
    Err.Raise Number:=30010, Description:="Invalid Col value"
End If
If PropFillStyle = FlexFillStyleSingle Then
    VBFlexGridCells.Rows(VBFlexGridRow).Cols(VBFlexGridCol).BackColor = Value
ElseIf PropFillStyle = FlexFillStyleRepeat Then
    Dim i As Long, j As Long, SelRange As TCELLRANGE
    Call GetSelRangeStruct(SelRange)
    For i = SelRange.TopRow To SelRange.BottomRow
        With VBFlexGridCells.Rows(i)
        For j = SelRange.LeftCol To SelRange.RightCol
            .Cols(j).BackColor = Value
        Next j
        End With
    Next i
End If
Call RedrawGrid
End Property

Public Property Get CellForeColor() As Long
Attribute CellForeColor.VB_Description = "Returns/sets the background and foreground colors of individual cells or ranges of cells."
Attribute CellForeColor.VB_MemberFlags = "400"
If VBFlexGridRow < 0 Then
    Err.Raise Number:=30009, Description:="Invalid Row value"
ElseIf VBFlexGridCol < 0 Then
    Err.Raise Number:=30010, Description:="Invalid Col value"
End If
If VBFlexGridCells.Rows(VBFlexGridRow).Cols(VBFlexGridCol).ForeColor = -1 Then
    If VBFlexGridRow > (PropFixedRows - 1) And VBFlexGridCol > (PropFixedCols - 1) Then
        CellForeColor = PropForeColor
    Else
        CellForeColor = PropForeColorFixed
    End If
Else
    CellForeColor = VBFlexGridCells.Rows(VBFlexGridRow).Cols(VBFlexGridCol).ForeColor
End If
End Property

Public Property Let CellForeColor(ByVal Value As Long)
If VBFlexGridRow < 0 Then
    Err.Raise Number:=30009, Description:="Invalid Row value"
ElseIf VBFlexGridCol < 0 Then
    Err.Raise Number:=30010, Description:="Invalid Col value"
End If
If PropFillStyle = FlexFillStyleSingle Then
    VBFlexGridCells.Rows(VBFlexGridRow).Cols(VBFlexGridCol).ForeColor = Value
ElseIf PropFillStyle = FlexFillStyleRepeat Then
    Dim i As Long, j As Long, SelRange As TCELLRANGE
    Call GetSelRangeStruct(SelRange)
    For i = SelRange.TopRow To SelRange.BottomRow
        With VBFlexGridCells.Rows(i)
        For j = SelRange.LeftCol To SelRange.RightCol
            .Cols(j).ForeColor = Value
        Next j
        End With
    Next i
End If
Call RedrawGrid
End Property

Public Property Get CellToolTipText() As String
Attribute CellToolTipText.VB_Description = "Returns/sets the tool tip text in a cell or in a range of selected cells. Requires that the show tips property is set to true."
Attribute CellToolTipText.VB_MemberFlags = "400"
If VBFlexGridRow > -1 And VBFlexGridCol > -1 Then CellToolTipText = VBFlexGridCells.Rows(VBFlexGridRow).Cols(VBFlexGridCol).ToolTipText
End Property

Public Property Let CellToolTipText(ByVal Value As String)
If VBFlexGridRow < 0 Then
    Err.Raise Number:=30009, Description:="Invalid Row value"
ElseIf VBFlexGridCol < 0 Then
    Err.Raise Number:=30010, Description:="Invalid Col value"
End If
If PropFillStyle = FlexFillStyleSingle Then
    VBFlexGridCells.Rows(VBFlexGridRow).Cols(VBFlexGridCol).ToolTipText = Value
ElseIf PropFillStyle = FlexFillStyleRepeat Then
    Dim i As Long, j As Long, SelRange As TCELLRANGE
    Call GetSelRangeStruct(SelRange)
    For i = SelRange.TopRow To SelRange.BottomRow
        With VBFlexGridCells.Rows(i)
        For j = SelRange.LeftCol To SelRange.RightCol
            .Cols(j).ToolTipText = Value
        Next j
        End With
    Next i
End If
End Property

Public Property Get CellFontName() As String
Attribute CellFontName.VB_Description = "Returns/sets the font name to be used for individual cells or ranges of cells."
Attribute CellFontName.VB_MemberFlags = "400"
If VBFlexGridRow < 0 Then
    Err.Raise Number:=30009, Description:="Invalid Row value"
ElseIf VBFlexGridCol < 0 Then
    Err.Raise Number:=30010, Description:="Invalid Col value"
End If
If VBFlexGridCells.Rows(VBFlexGridRow).Cols(VBFlexGridCol).FontName = vbNullString Then
    If PropFontFixed Is Nothing Or (VBFlexGridRow > (PropFixedRows - 1) And VBFlexGridCol > (PropFixedCols - 1)) Then
        CellFontName = PropFont.Name
    Else
        CellFontName = PropFontFixed.Name
    End If
Else
    CellFontName = VBFlexGridCells.Rows(VBFlexGridRow).Cols(VBFlexGridCol).FontName
End If
End Property

Public Property Let CellFontName(ByVal Value As String)
If VBFlexGridRow < 0 Then
    Err.Raise Number:=30009, Description:="Invalid Row value"
ElseIf VBFlexGridCol < 0 Then
    Err.Raise Number:=30010, Description:="Invalid Col value"
End If
Dim TempFont As StdFont
If PropFillStyle = FlexFillStyleSingle Then
    With VBFlexGridCells.Rows(VBFlexGridRow).Cols(VBFlexGridCol)
    If Not Value = vbNullString Then
        If .FontName = vbNullString Then
            If PropFontFixed Is Nothing Or (VBFlexGridRow > (PropFixedRows - 1) And VBFlexGridCol > (PropFixedCols - 1)) Then
                .FontSize = PropFont.Size
                .FontBold = PropFont.Bold
                .FontItalic = PropFont.Italic
                .FontStrikeThrough = PropFont.Strikethrough
                .FontUnderline = PropFont.Underline
            Else
                .FontSize = PropFontFixed.Size
                .FontBold = PropFontFixed.Bold
                .FontItalic = PropFontFixed.Italic
                .FontStrikeThrough = PropFontFixed.Strikethrough
                .FontUnderline = PropFontFixed.Underline
            End If
        End If
        Set TempFont = New StdFont
        TempFont.Name = Value
        .FontName = TempFont.Name
        .FontCharset = TempFont.Charset
    Else
        .FontName = vbNullString
    End If
    End With
ElseIf PropFillStyle = FlexFillStyleRepeat Then
    Dim i As Long, j As Long, SelRange As TCELLRANGE
    Call GetSelRangeStruct(SelRange)
    For i = SelRange.TopRow To SelRange.BottomRow
        With VBFlexGridCells.Rows(i)
        For j = SelRange.LeftCol To SelRange.RightCol
            With .Cols(j)
            If Not Value = vbNullString Then
                If .FontName = vbNullString Then
                    If PropFontFixed Is Nothing Or (i > (PropFixedRows - 1) And j > (PropFixedCols - 1)) Then
                        .FontSize = PropFont.Size
                        .FontBold = PropFont.Bold
                        .FontItalic = PropFont.Italic
                        .FontStrikeThrough = PropFont.Strikethrough
                        .FontUnderline = PropFont.Underline
                    Else
                        .FontSize = PropFontFixed.Size
                        .FontBold = PropFontFixed.Bold
                        .FontItalic = PropFontFixed.Italic
                        .FontStrikeThrough = PropFontFixed.Strikethrough
                        .FontUnderline = PropFontFixed.Underline
                    End If
                End If
                Set TempFont = New StdFont
                TempFont.Name = Value
                .FontName = TempFont.Name
                .FontCharset = TempFont.Charset
            Else
                .FontName = vbNullString
            End If
            End With
        Next j
        End With
    Next i
End If
Call RedrawGrid
End Property

Public Property Get CellFontSize() As Single
Attribute CellFontSize.VB_Description = "Returns/sets the font size (in points) to be used for individual cells or ranges of cells."
Attribute CellFontSize.VB_MemberFlags = "400"
If VBFlexGridRow < 0 Then
    Err.Raise Number:=30009, Description:="Invalid Row value"
ElseIf VBFlexGridCol < 0 Then
    Err.Raise Number:=30010, Description:="Invalid Col value"
End If
If VBFlexGridCells.Rows(VBFlexGridRow).Cols(VBFlexGridCol).FontName = vbNullString Then
    If PropFontFixed Is Nothing Or (VBFlexGridRow > (PropFixedRows - 1) And VBFlexGridCol > (PropFixedCols - 1)) Then
        CellFontSize = PropFont.Size
    Else
        CellFontSize = PropFontFixed.Size
    End If
Else
    CellFontSize = VBFlexGridCells.Rows(VBFlexGridRow).Cols(VBFlexGridCol).FontSize
End If
End Property

Public Property Let CellFontSize(ByVal Value As Single)
If VBFlexGridRow < 0 Then
    Err.Raise Number:=30009, Description:="Invalid Row value"
ElseIf VBFlexGridCol < 0 Then
    Err.Raise Number:=30010, Description:="Invalid Col value"
End If
If PropFillStyle = FlexFillStyleSingle Then
    With VBFlexGridCells.Rows(VBFlexGridRow).Cols(VBFlexGridCol)
    If .FontName = vbNullString Then
        If PropFontFixed Is Nothing Or (VBFlexGridRow > (PropFixedRows - 1) And VBFlexGridCol > (PropFixedCols - 1)) Then
            .FontName = PropFont.Name
            .FontBold = PropFont.Bold
            .FontItalic = PropFont.Italic
            .FontStrikeThrough = PropFont.Strikethrough
            .FontUnderline = PropFont.Underline
            .FontCharset = PropFont.Charset
        Else
            .FontName = PropFontFixed.Name
            .FontBold = PropFontFixed.Bold
            .FontItalic = PropFontFixed.Italic
            .FontStrikeThrough = PropFontFixed.Strikethrough
            .FontUnderline = PropFontFixed.Underline
            .FontCharset = PropFontFixed.Charset
        End If
    End If
    .FontSize = Value
    End With
ElseIf PropFillStyle = FlexFillStyleRepeat Then
    Dim i As Long, j As Long, SelRange As TCELLRANGE
    Call GetSelRangeStruct(SelRange)
    For i = SelRange.TopRow To SelRange.BottomRow
        With VBFlexGridCells.Rows(i)
        For j = SelRange.LeftCol To SelRange.RightCol
            With .Cols(j)
            If .FontName = vbNullString Then
                If PropFontFixed Is Nothing Or (i > (PropFixedRows - 1) And j > (PropFixedCols - 1)) Then
                    .FontName = PropFont.Name
                    .FontBold = PropFont.Bold
                    .FontItalic = PropFont.Italic
                    .FontStrikeThrough = PropFont.Strikethrough
                    .FontUnderline = PropFont.Underline
                    .FontCharset = PropFont.Charset
                Else
                    .FontName = PropFontFixed.Name
                    .FontBold = PropFontFixed.Bold
                    .FontItalic = PropFontFixed.Italic
                    .FontStrikeThrough = PropFontFixed.Strikethrough
                    .FontUnderline = PropFontFixed.Underline
                    .FontCharset = PropFontFixed.Charset
                End If
            End If
            .FontSize = Value
            End With
        Next j
        End With
    Next i
End If
Call RedrawGrid
End Property

Public Property Get CellFontBold() As Boolean
Attribute CellFontBold.VB_Description = "Returns/sets the font bold style to be used for individual cells or ranges of cells."
Attribute CellFontBold.VB_MemberFlags = "400"
If VBFlexGridRow < 0 Then
    Err.Raise Number:=30009, Description:="Invalid Row value"
ElseIf VBFlexGridCol < 0 Then
    Err.Raise Number:=30010, Description:="Invalid Col value"
End If
If VBFlexGridCells.Rows(VBFlexGridRow).Cols(VBFlexGridCol).FontName = vbNullString Then
    If PropFontFixed Is Nothing Or (VBFlexGridRow > (PropFixedRows - 1) And VBFlexGridCol > (PropFixedCols - 1)) Then
        CellFontBold = PropFont.Bold
    Else
        CellFontBold = PropFontFixed.Bold
    End If
Else
    CellFontBold = VBFlexGridCells.Rows(VBFlexGridRow).Cols(VBFlexGridCol).FontBold
End If
End Property

Public Property Let CellFontBold(ByVal Value As Boolean)
If VBFlexGridRow < 0 Then
    Err.Raise 30009, Description:="Invalid Row value"
ElseIf VBFlexGridCol < 0 Then
    Err.Raise 30010, Description:="Invalid Col value"
End If
If PropFillStyle = FlexFillStyleSingle Then
    With VBFlexGridCells.Rows(VBFlexGridRow).Cols(VBFlexGridCol)
    If .FontName = vbNullString Then
        If PropFontFixed Is Nothing Or (VBFlexGridRow > (PropFixedRows - 1) And VBFlexGridCol > (PropFixedCols - 1)) Then
            .FontName = PropFont.Name
            .FontSize = PropFont.Size
            .FontItalic = PropFont.Italic
            .FontStrikeThrough = PropFont.Strikethrough
            .FontUnderline = PropFont.Underline
            .FontCharset = PropFont.Charset
        Else
            .FontName = PropFontFixed.Name
            .FontSize = PropFontFixed.Size
            .FontItalic = PropFontFixed.Italic
            .FontStrikeThrough = PropFontFixed.Strikethrough
            .FontUnderline = PropFontFixed.Underline
            .FontCharset = PropFontFixed.Charset
        End If
    End If
    .FontBold = Value
    End With
ElseIf PropFillStyle = FlexFillStyleRepeat Then
    Dim i As Long, j As Long, SelRange As TCELLRANGE
    Call GetSelRangeStruct(SelRange)
    For i = SelRange.TopRow To SelRange.BottomRow
        With VBFlexGridCells.Rows(i)
        For j = SelRange.LeftCol To SelRange.RightCol
            With .Cols(j)
            If .FontName = vbNullString Then
                If PropFontFixed Is Nothing Or (i > (PropFixedRows - 1) And j > (PropFixedCols - 1)) Then
                    .FontName = PropFont.Name
                    .FontSize = PropFont.Size
                    .FontItalic = PropFont.Italic
                    .FontStrikeThrough = PropFont.Strikethrough
                    .FontUnderline = PropFont.Underline
                    .FontCharset = PropFont.Charset
                Else
                    .FontName = PropFontFixed.Name
                    .FontSize = PropFontFixed.Size
                    .FontItalic = PropFontFixed.Italic
                    .FontStrikeThrough = PropFontFixed.Strikethrough
                    .FontUnderline = PropFontFixed.Underline
                    .FontCharset = PropFontFixed.Charset
                End If
            End If
            .FontBold = Value
            End With
        Next j
        End With
    Next i
End If
Call RedrawGrid
End Property

Public Property Get CellFontItalic() As Boolean
Attribute CellFontItalic.VB_Description = "Returns/sets the font italic style to be used for individual cells or ranges of cells."
Attribute CellFontItalic.VB_MemberFlags = "400"
If VBFlexGridRow < 0 Then
    Err.Raise Number:=30009, Description:="Invalid Row value"
ElseIf VBFlexGridCol < 0 Then
    Err.Raise Number:=30010, Description:="Invalid Col value"
End If
If VBFlexGridCells.Rows(VBFlexGridRow).Cols(VBFlexGridCol).FontName = vbNullString Then
    If PropFontFixed Is Nothing Or (VBFlexGridRow > (PropFixedRows - 1) And VBFlexGridCol > (PropFixedCols - 1)) Then
        CellFontItalic = PropFont.Italic
    Else
        CellFontItalic = PropFontFixed.Italic
    End If
Else
    CellFontItalic = VBFlexGridCells.Rows(VBFlexGridRow).Cols(VBFlexGridCol).FontItalic
End If
End Property

Public Property Let CellFontItalic(ByVal Value As Boolean)
If VBFlexGridRow < 0 Then
    Err.Raise 30009, Description:="Invalid Row value"
ElseIf VBFlexGridCol < 0 Then
    Err.Raise 30010, Description:="Invalid Col value"
End If
If PropFillStyle = FlexFillStyleSingle Then
    With VBFlexGridCells.Rows(VBFlexGridRow).Cols(VBFlexGridCol)
    If .FontName = vbNullString Then
        If PropFontFixed Is Nothing Or (VBFlexGridRow > (PropFixedRows - 1) And VBFlexGridCol > (PropFixedCols - 1)) Then
            .FontName = PropFont.Name
            .FontSize = PropFont.Size
            .FontBold = PropFont.Bold
            .FontStrikeThrough = PropFont.Strikethrough
            .FontUnderline = PropFont.Underline
            .FontCharset = PropFont.Charset
        Else
            .FontName = PropFontFixed.Name
            .FontSize = PropFontFixed.Size
            .FontBold = PropFontFixed.Bold
            .FontStrikeThrough = PropFontFixed.Strikethrough
            .FontUnderline = PropFontFixed.Underline
            .FontCharset = PropFontFixed.Charset
        End If
    End If
    .FontItalic = Value
    End With
ElseIf PropFillStyle = FlexFillStyleRepeat Then
    Dim i As Long, j As Long, SelRange As TCELLRANGE
    Call GetSelRangeStruct(SelRange)
    For i = SelRange.TopRow To SelRange.BottomRow
        With VBFlexGridCells.Rows(i)
        For j = SelRange.LeftCol To SelRange.RightCol
            With .Cols(j)
            If .FontName = vbNullString Then
                If PropFontFixed Is Nothing Or (i > (PropFixedRows - 1) And j > (PropFixedCols - 1)) Then
                    .FontName = PropFont.Name
                    .FontSize = PropFont.Size
                    .FontBold = PropFont.Bold
                    .FontStrikeThrough = PropFont.Strikethrough
                    .FontUnderline = PropFont.Underline
                    .FontCharset = PropFont.Charset
                Else
                    .FontName = PropFontFixed.Name
                    .FontSize = PropFontFixed.Size
                    .FontBold = PropFontFixed.Bold
                    .FontStrikeThrough = PropFontFixed.Strikethrough
                    .FontUnderline = PropFontFixed.Underline
                    .FontCharset = PropFontFixed.Charset
                End If
            End If
            .FontItalic = Value
            End With
        Next j
        End With
    Next i
End If
Call RedrawGrid
End Property

Public Property Get CellFontStrikeThrough() As Boolean
Attribute CellFontStrikeThrough.VB_Description = "Returns/sets the font strikethrough style to be used for individual cells or ranges of cells."
Attribute CellFontStrikeThrough.VB_MemberFlags = "400"
If VBFlexGridRow < 0 Then
    Err.Raise Number:=30009, Description:="Invalid Row value"
ElseIf VBFlexGridCol < 0 Then
    Err.Raise Number:=30010, Description:="Invalid Col value"
End If
If VBFlexGridCells.Rows(VBFlexGridRow).Cols(VBFlexGridCol).FontName = vbNullString Then
    If PropFontFixed Is Nothing Or (VBFlexGridRow > (PropFixedRows - 1) And VBFlexGridCol > (PropFixedCols - 1)) Then
        CellFontStrikeThrough = PropFont.Strikethrough
    Else
        CellFontStrikeThrough = PropFontFixed.Strikethrough
    End If
Else
    CellFontStrikeThrough = VBFlexGridCells.Rows(VBFlexGridRow).Cols(VBFlexGridCol).FontStrikeThrough
End If
End Property

Public Property Let CellFontStrikeThrough(ByVal Value As Boolean)
If VBFlexGridRow < 0 Then
    Err.Raise Number:=30009, Description:="Invalid Row value"
ElseIf VBFlexGridCol < 0 Then
    Err.Raise Number:=30010, Description:="Invalid Col value"
End If
If PropFillStyle = FlexFillStyleSingle Then
    With VBFlexGridCells.Rows(VBFlexGridRow).Cols(VBFlexGridCol)
    If .FontName = vbNullString Then
        If PropFontFixed Is Nothing Or (VBFlexGridRow > (PropFixedRows - 1) And VBFlexGridCol > (PropFixedCols - 1)) Then
            .FontName = PropFont.Name
            .FontSize = PropFont.Size
            .FontBold = PropFont.Bold
            .FontItalic = PropFont.Italic
            .FontUnderline = PropFont.Underline
            .FontCharset = PropFont.Charset
        Else
            .FontName = PropFontFixed.Name
            .FontSize = PropFontFixed.Size
            .FontBold = PropFontFixed.Bold
            .FontItalic = PropFontFixed.Italic
            .FontUnderline = PropFontFixed.Underline
            .FontCharset = PropFontFixed.Charset
        End If
    End If
    .FontStrikeThrough = Value
    End With
ElseIf PropFillStyle = FlexFillStyleRepeat Then
    Dim i As Long, j As Long, SelRange As TCELLRANGE
    Call GetSelRangeStruct(SelRange)
    For i = SelRange.TopRow To SelRange.BottomRow
        With VBFlexGridCells.Rows(i)
        For j = SelRange.LeftCol To SelRange.RightCol
            With .Cols(j)
            If .FontName = vbNullString Then
                If PropFontFixed Is Nothing Or (i > (PropFixedRows - 1) And j > (PropFixedCols - 1)) Then
                    .FontName = PropFont.Name
                    .FontSize = PropFont.Size
                    .FontBold = PropFont.Bold
                    .FontItalic = PropFont.Italic
                    .FontUnderline = PropFont.Underline
                    .FontCharset = PropFont.Charset
                Else
                    .FontName = PropFontFixed.Name
                    .FontSize = PropFontFixed.Size
                    .FontBold = PropFontFixed.Bold
                    .FontItalic = PropFontFixed.Italic
                    .FontUnderline = PropFontFixed.Underline
                    .FontCharset = PropFontFixed.Charset
                End If
            End If
            .FontStrikeThrough = Value
            End With
        Next j
        End With
    Next i
End If
Call RedrawGrid
End Property

Public Property Get CellFontUnderline() As Boolean
Attribute CellFontUnderline.VB_Description = "Returns/sets the font underline style to be used for individual cells or ranges of cells."
Attribute CellFontUnderline.VB_MemberFlags = "400"
If VBFlexGridRow < 0 Then
    Err.Raise Number:=30009, Description:="Invalid Row value"
ElseIf VBFlexGridCol < 0 Then
    Err.Raise Number:=30010, Description:="Invalid Col value"
End If
If VBFlexGridCells.Rows(VBFlexGridRow).Cols(VBFlexGridCol).FontName = vbNullString Then
    If PropFontFixed Is Nothing Or (VBFlexGridRow > (PropFixedRows - 1) And VBFlexGridCol > (PropFixedCols - 1)) Then
        CellFontUnderline = PropFont.Underline
    Else
        CellFontUnderline = PropFontFixed.Underline
    End If
Else
    CellFontUnderline = VBFlexGridCells.Rows(VBFlexGridRow).Cols(VBFlexGridCol).FontUnderline
End If
End Property

Public Property Let CellFontUnderline(ByVal Value As Boolean)
If VBFlexGridRow < 0 Then
    Err.Raise 30009, Description:="Invalid Row value"
ElseIf VBFlexGridCol < 0 Then
    Err.Raise 30010, Description:="Invalid Col value"
End If
If PropFillStyle = FlexFillStyleSingle Then
    With VBFlexGridCells.Rows(VBFlexGridRow).Cols(VBFlexGridCol)
    If .FontName = vbNullString Then
        If PropFontFixed Is Nothing Or (VBFlexGridRow > (PropFixedRows - 1) And VBFlexGridCol > (PropFixedCols - 1)) Then
            .FontName = PropFont.Name
            .FontSize = PropFont.Size
            .FontBold = PropFont.Bold
            .FontItalic = PropFont.Italic
            .FontStrikeThrough = PropFont.Strikethrough
            .FontCharset = PropFont.Charset
        Else
            .FontName = PropFontFixed.Name
            .FontSize = PropFontFixed.Size
            .FontBold = PropFontFixed.Bold
            .FontItalic = PropFontFixed.Italic
            .FontStrikeThrough = PropFontFixed.Strikethrough
            .FontCharset = PropFontFixed.Charset
        End If
    End If
    .FontUnderline = Value
    End With
ElseIf PropFillStyle = FlexFillStyleRepeat Then
    Dim i As Long, j As Long, SelRange As TCELLRANGE
    Call GetSelRangeStruct(SelRange)
    For i = SelRange.TopRow To SelRange.BottomRow
        With VBFlexGridCells.Rows(i)
        For j = SelRange.LeftCol To SelRange.RightCol
            With .Cols(j)
            If .FontName = vbNullString Then
                If PropFontFixed Is Nothing Or (i > (PropFixedRows - 1) And j > (PropFixedCols - 1)) Then
                    .FontName = PropFont.Name
                    .FontSize = PropFont.Size
                    .FontBold = PropFont.Bold
                    .FontItalic = PropFont.Italic
                    .FontStrikeThrough = PropFont.Strikethrough
                    .FontCharset = PropFont.Charset
                Else
                    .FontName = PropFontFixed.Name
                    .FontSize = PropFontFixed.Size
                    .FontBold = PropFontFixed.Bold
                    .FontItalic = PropFontFixed.Italic
                    .FontStrikeThrough = PropFontFixed.Strikethrough
                    .FontCharset = PropFontFixed.Charset
                End If
            End If
            .FontUnderline = Value
            End With
        Next j
        End With
    Next i
End If
Call RedrawGrid
End Property

Public Property Get CellFontCharset() As Integer
Attribute CellFontCharset.VB_Description = "Returns/sets the font charset to be used for individual cells or ranges of cells."
Attribute CellFontCharset.VB_MemberFlags = "400"
If VBFlexGridRow < 0 Then
    Err.Raise Number:=30009, Description:="Invalid Row value"
ElseIf VBFlexGridCol < 0 Then
    Err.Raise Number:=30010, Description:="Invalid Col value"
End If
If VBFlexGridCells.Rows(VBFlexGridRow).Cols(VBFlexGridCol).FontName = vbNullString Then
    If PropFontFixed Is Nothing Or (VBFlexGridRow > (PropFixedRows - 1) And VBFlexGridCol > (PropFixedCols - 1)) Then
        CellFontCharset = PropFont.Charset
    Else
        CellFontCharset = PropFontFixed.Charset
    End If
Else
    CellFontCharset = VBFlexGridCells.Rows(VBFlexGridRow).Cols(VBFlexGridCol).FontCharset
End If
End Property

Public Property Let CellFontCharset(ByVal Value As Integer)
If VBFlexGridRow < 0 Then
    Err.Raise Number:=30009, Description:="Invalid Row value"
ElseIf VBFlexGridCol < 0 Then
    Err.Raise Number:=30010, Description:="Invalid Col value"
End If
If PropFillStyle = FlexFillStyleSingle Then
    With VBFlexGridCells.Rows(VBFlexGridRow).Cols(VBFlexGridCol)
    If .FontName = vbNullString Then
        If PropFontFixed Is Nothing Or (VBFlexGridRow > (PropFixedRows - 1) And VBFlexGridCol > (PropFixedCols - 1)) Then
            .FontName = PropFont.Name
            .FontSize = PropFont.Size
            .FontBold = PropFont.Bold
            .FontItalic = PropFont.Italic
            .FontStrikeThrough = PropFont.Strikethrough
            .FontUnderline = PropFont.Underline
        Else
            .FontName = PropFontFixed.Name
            .FontSize = PropFontFixed.Size
            .FontBold = PropFontFixed.Bold
            .FontItalic = PropFontFixed.Italic
            .FontStrikeThrough = PropFontFixed.Strikethrough
            .FontUnderline = PropFontFixed.Underline
        End If
    End If
    .FontCharset = Value
    End With
ElseIf PropFillStyle = FlexFillStyleRepeat Then
    Dim i As Long, j As Long, SelRange As TCELLRANGE
    Call GetSelRangeStruct(SelRange)
    For i = SelRange.TopRow To SelRange.BottomRow
        With VBFlexGridCells.Rows(i)
        For j = SelRange.LeftCol To SelRange.RightCol
            With .Cols(j)
            If .FontName = vbNullString Then
                If PropFontFixed Is Nothing Or (i > (PropFixedRows - 1) And j > (PropFixedCols - 1)) Then
                    .FontName = PropFont.Name
                    .FontSize = PropFont.Size
                    .FontBold = PropFont.Bold
                    .FontItalic = PropFont.Italic
                    .FontStrikeThrough = PropFont.Strikethrough
                    .FontUnderline = PropFont.Underline
                Else
                    .FontName = PropFontFixed.Name
                    .FontSize = PropFontFixed.Size
                    .FontBold = PropFontFixed.Bold
                    .FontItalic = PropFontFixed.Italic
                    .FontStrikeThrough = PropFontFixed.Strikethrough
                    .FontUnderline = PropFontFixed.Underline
                End If
            End If
            .FontCharset = Value
            End With
        Next j
        End With
    Next i
End If
Call RedrawGrid
End Property

Public Sub CellEnsureVisible(Optional ByVal Visibility As FlexVisibilityConstants = FlexVisibilityCompleteOnly, Optional ByVal Row As Long = -1, Optional ByVal Col As Long = -1)
Attribute CellEnsureVisible.VB_Description = "Ensures that the current or an arbitrary cell (row/col subscripts) is visible, scrolling the control if necessary."
Select Case Visibility
    Case FlexVisibilityPartialOK, FlexVisibilityCompleteOnly
    Case Else
        Err.Raise 380
End Select
If Row < -1 Then Err.Raise 380
If Col < -1 Then Err.Raise 380
If PropRows < 1 Or PropCols < 1 Then Exit Sub
If Row = -1 Then Row = VBFlexGridRow
If Col = -1 Then Col = VBFlexGridCol
If (Row < 0 Or Row > (PropRows - 1)) Or (Col < 0 Or Col > (PropCols - 1)) Then Err.Raise Number:=381, Description:="Subscript out of range"
If Visibility = FlexVisibilityPartialOK Then
    If Me.RowIsVisible(Row, FlexVisibilityPartialOK) = True And Me.ColIsVisible(Col, FlexVisibilityPartialOK) = True Then Exit Sub
End If
Dim RCP As TROWCOLPARAMS
With RCP
.Mask = RCPM_TOPROW Or RCPM_LEFTCOL
.TopRow = VBFlexGridTopRow
.LeftCol = VBFlexGridLeftCol
If .TopRow > Row Then
    If Row >= (PropFixedRows + PropFrozenRows) Then .TopRow = Row
ElseIf Row > (.TopRow + GetRowsPerPage(.TopRow) - 1) Then
    .TopRow = Row - GetRowsPerPageRev(Row) + 1
End If
If .LeftCol > Col Then
    If Col >= (PropFixedCols + PropFrozenCols) Then .LeftCol = Col
ElseIf Col > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
    .LeftCol = Col - GetColsPerPageRev(Col) + 1
End If
Call SetRowColParams(RCP)
End With
End Sub

Public Property Get CellLeft() As Long
Attribute CellLeft.VB_Description = "Returns the left coordinate in twips of the current cell."
Attribute CellLeft.VB_MemberFlags = "400"
If VBFlexGridRow < 0 Then
    Err.Raise Number:=30009, Description:="Invalid Row value"
ElseIf VBFlexGridCol < 0 Then
    Err.Raise Number:=30010, Description:="Invalid Col value"
End If
Me.CellEnsureVisible
Dim CellRect As RECT
Call GetCellRect(VBFlexGridRow, VBFlexGridCol, True, CellRect)
CellLeft = UserControl.ScaleX(CellRect.Left, vbPixels, vbTwips)
End Property

Public Property Get CellTop() As Long
Attribute CellTop.VB_Description = "Returns the top coordinate in twips of the current cell."
Attribute CellTop.VB_MemberFlags = "400"
If VBFlexGridRow < 0 Then
    Err.Raise Number:=30009, Description:="Invalid Row value"
ElseIf VBFlexGridCol < 0 Then
    Err.Raise Number:=30010, Description:="Invalid Col value"
End If
Me.CellEnsureVisible
Dim CellRect As RECT
Call GetCellRect(VBFlexGridRow, VBFlexGridCol, True, CellRect)
CellTop = UserControl.ScaleY(CellRect.Top, vbPixels, vbTwips)
End Property

Public Property Get CellWidth() As Long
Attribute CellWidth.VB_Description = "Returns the width in twips of the current cell."
Attribute CellWidth.VB_MemberFlags = "400"
If VBFlexGridRow < 0 Then
    Err.Raise Number:=30009, Description:="Invalid Row value"
ElseIf VBFlexGridCol < 0 Then
    Err.Raise Number:=30010, Description:="Invalid Col value"
End If
Me.CellEnsureVisible
Dim CellRect As RECT
Call GetCellRect(VBFlexGridRow, VBFlexGridCol, True, CellRect)
CellWidth = UserControl.ScaleX((CellRect.Right - CellRect.Left) - 1, vbPixels, vbTwips)
End Property

Public Property Get CellHeight() As Long
Attribute CellHeight.VB_Description = "Returns the height in twips of the current cell."
Attribute CellHeight.VB_MemberFlags = "400"
If VBFlexGridRow < 0 Then
    Err.Raise Number:=30009, Description:="Invalid Row value"
ElseIf VBFlexGridCol < 0 Then
    Err.Raise Number:=30010, Description:="Invalid Col value"
End If
Me.CellEnsureVisible
Dim CellRect As RECT
Call GetCellRect(VBFlexGridRow, VBFlexGridCol, True, CellRect)
CellHeight = UserControl.ScaleY((CellRect.Bottom - CellRect.Top) - 1, vbPixels, vbTwips)
End Property

Public Sub HitTest(ByVal X As Single, ByVal Y As Single)
Attribute HitTest.VB_Description = "A method that returns a value which indicates the element located at the specified X and Y coordinates."
Dim HTI As THITTESTINFO
With HTI
.PT.X = UserControl.ScaleX(X, vbContainerPosition, vbPixels)
.PT.Y = UserControl.ScaleY(Y, vbContainerPosition, vbPixels)
Call GetHitTestInfo(HTI)
VBFlexGridHitRow = .HitRow
VBFlexGridHitCol = .HitCol
VBFlexGridHitRowDivider = .HitRowDivider
VBFlexGridHitColDivider = .HitColDivider
VBFlexGridHitResult = .HitResult
End With
End Sub

Public Function FindItem(ByVal Text As String, Optional ByVal Row As Long = -1, Optional ByVal Col As Long = -1, Optional ByVal Partial As Boolean, Optional ByVal CaseSensitive As Boolean, Optional ByVal ExcludeHidden As Boolean, Optional ByVal Wrap As Boolean, Optional ByVal Direction As FlexFindDirectionConstants) As Long
Attribute FindItem.VB_Description = "Finds an item in the flex grid and returns the index of that item."
If Row < -1 Then Err.Raise 380
If Col < -1 Then Err.Raise 380
Select Case Direction
    Case FlexFindDirectionDown, FlexFindDirectionUp
    Case Else
        Err.Raise 380
End Select
If Row = -1 Then Row = IIf(Direction = FlexFindDirectionDown, PropFixedRows, (PropRows - 1))
If Col = -1 Then Col = PropFixedCols
If (Row < 0 Or Row > (PropRows - 1)) Or (Col < 0 Or Col > (PropCols - 1)) Then Err.Raise Number:=381, Description:="Subscript out of range"
If Row < PropFixedRows Then Err.Raise Number:=30003, Description:="Cannot use FindItem on a fixed row"
Dim iRow As Long, iRowTo As Long, Compare As VbCompareMethod, Buffer As String
FindItem = -1
If Direction = FlexFindDirectionDown Then iRowTo = (PropRows - 1) Else iRowTo = PropFixedRows
If CaseSensitive = False Then Compare = vbTextCompare Else Compare = vbBinaryCompare
If Partial = False Then
    For iRow = Row To iRowTo Step IIf(Direction = FlexFindDirectionDown, 1, -1)
        If (CBool((VBFlexGridCells.Rows(iRow).RowInfo.State And RWIS_HIDDEN) = RWIS_HIDDEN) Xor ExcludeHidden) Or ExcludeHidden = False Then
            Call GetCellText(iRow, Col, Buffer)
            If StrComp(Buffer, Text, Compare) = 0 Then
                FindItem = iRow
                Exit For
            End If
        End If
    Next iRow
Else
    For iRow = Row To iRowTo Step IIf(Direction = FlexFindDirectionDown, 1, -1)
        If (CBool((VBFlexGridCells.Rows(iRow).RowInfo.State And RWIS_HIDDEN) = RWIS_HIDDEN) Xor ExcludeHidden) Or ExcludeHidden = False Then
            Call GetCellText(iRow, Col, Buffer)
            If InStr(1, Buffer, Text, Compare) > 0 Then
                FindItem = iRow
                Exit For
            End If
        End If
    Next iRow
End If
If Wrap = True And FindItem = -1 Then
    If Direction = FlexFindDirectionDown Then iRowTo = PropFixedRows Else iRowTo = (PropRows - 1)
    If Partial = False Then
        For iRow = iRowTo To (Row - IIf(Direction = FlexFindDirectionDown, 1, -1)) Step IIf(Direction = FlexFindDirectionDown, 1, -1)
            If (CBool((VBFlexGridCells.Rows(iRow).RowInfo.State And RWIS_HIDDEN) = RWIS_HIDDEN) Xor ExcludeHidden) Or ExcludeHidden = False Then
                Call GetCellText(iRow, Col, Buffer)
                If StrComp(Buffer, Text, Compare) = 0 Then
                    FindItem = iRow
                    Exit For
                End If
            End If
        Next iRow
    Else
        For iRow = iRowTo To (Row - IIf(Direction = FlexFindDirectionDown, 1, -1)) Step IIf(Direction = FlexFindDirectionDown, 1, -1)
            If (CBool((VBFlexGridCells.Rows(iRow).RowInfo.State And RWIS_HIDDEN) = RWIS_HIDDEN) Xor ExcludeHidden) Or ExcludeHidden = False Then
                Call GetCellText(iRow, Col, Buffer)
                If InStr(1, Buffer, Text, Compare) > 0 Then
                    FindItem = iRow
                    Exit For
                End If
            End If
        Next iRow
    End If
End If
End Function

Public Sub AutoSize(ByVal RowOrCol1 As Long, Optional ByVal RowOrCol2 As Long = -1, Optional ByVal Mode As FlexAutoSizeModeConstants, Optional ByVal Scope As FlexAutoSizeScopeConstants, Optional ByVal Equal As Boolean, Optional ByVal ExtraSpace As Long, Optional ByVal ExcludeHidden As Boolean)
Attribute AutoSize.VB_Description = "Automatically sizes column widths or row heights to fit cell contents."
If RowOrCol2 < -1 Then Err.Raise 380
If RowOrCol2 = -1 Then RowOrCol2 = RowOrCol1
Select Case Mode
    Case FlexAutoSizeModeColWidth, FlexAutoSizeModeRowHeight
    Case Else
        Err.Raise 380
End Select
Select Case Scope
    Case FlexAutoSizeScopeAll, FlexAutoSizeScopeFixed, FlexAutoSizeScopeScrollable, FlexAutoSizeScopeMovable, FlexAutoSizeScopeFrozen
    Case Else
        Err.Raise 380
End Select
If Mode = FlexAutoSizeModeColWidth Then
    If (RowOrCol1 < 0 Or RowOrCol1 > (PropCols - 1)) Or (RowOrCol2 < 0 Or RowOrCol2 > (PropCols - 1)) Then Err.Raise Number:=381, Description:="Subscript out of range"
ElseIf Mode = FlexAutoSizeModeRowHeight Then
    If (RowOrCol1 < 0 Or RowOrCol1 > (PropRows - 1)) Or (RowOrCol2 < 0 Or RowOrCol2 > (PropRows - 1)) Then Err.Raise Number:=381, Description:="Subscript out of range"
End If
Dim iRow As Long, iCol As Long, Text As String, Spacing As Long, Size As SIZEAPI, EqualSize As SIZEAPI
If Mode = FlexAutoSizeModeColWidth Then
    Spacing = (COLINFO_WIDTH_SPACING_DIP * PixelsPerDIP_X()) + CLng(UserControl.ScaleX(ExtraSpace, vbTwips, vbPixels))
    EqualSize.CX = -1
    Select Case Scope
        Case FlexAutoSizeScopeAll
            For iCol = RowOrCol1 To RowOrCol2 Step IIf(RowOrCol2 >= RowOrCol1, 1, -1)
                With VBFlexGridColsInfo(iCol)
                If (CBool((.State And CLIS_HIDDEN) = CLIS_HIDDEN) Xor ExcludeHidden) Or ExcludeHidden = False Then
                    .Width = -1
                    For iRow = 0 To (PropRows - 1)
                        If (CBool((VBFlexGridCells.Rows(iRow).RowInfo.State And RWIS_HIDDEN) = RWIS_HIDDEN) Xor ExcludeHidden) Or ExcludeHidden = False Then
                            Call GetCellText(iRow, iCol, Text)
                            Size.CX = GetTextSize(iRow, iCol, Text).CX
                            If Size.CX > 0 Then
                                Size.CX = Size.CX + Spacing
                                If Size.CX > .Width Then .Width = Size.CX
                                If Size.CX > EqualSize.CX Then EqualSize.CX = Size.CX
                            End If
                        End If
                    Next iRow
                End If
                End With
            Next iCol
        Case FlexAutoSizeScopeFixed
            For iCol = RowOrCol1 To RowOrCol2 Step IIf(RowOrCol2 >= RowOrCol1, 1, -1)
                With VBFlexGridColsInfo(iCol)
                If (CBool((.State And CLIS_HIDDEN) = CLIS_HIDDEN) Xor ExcludeHidden) Or ExcludeHidden = False Then
                    .Width = -1
                    For iRow = 0 To (PropFixedRows - 1)
                        If (CBool((VBFlexGridCells.Rows(iRow).RowInfo.State And RWIS_HIDDEN) = RWIS_HIDDEN) Xor ExcludeHidden) Or ExcludeHidden = False Then
                            Call GetCellText(iRow, iCol, Text)
                            Size.CX = GetTextSize(iRow, iCol, Text).CX
                            If Size.CX > 0 Then
                                Size.CX = Size.CX + Spacing
                                If Size.CX > .Width Then .Width = Size.CX
                                If Size.CX > EqualSize.CX Then EqualSize.CX = Size.CX
                            End If
                        End If
                    Next iRow
                End If
                End With
            Next iCol
        Case FlexAutoSizeScopeScrollable
            For iCol = RowOrCol1 To RowOrCol2 Step IIf(RowOrCol2 >= RowOrCol1, 1, -1)
                With VBFlexGridColsInfo(iCol)
                If (CBool((.State And CLIS_HIDDEN) = CLIS_HIDDEN) Xor ExcludeHidden) Or ExcludeHidden = False Then
                    .Width = -1
                    For iRow = (PropFixedRows + PropFrozenRows) To (PropRows - 1)
                        If (CBool((VBFlexGridCells.Rows(iRow).RowInfo.State And RWIS_HIDDEN) = RWIS_HIDDEN) Xor ExcludeHidden) Or ExcludeHidden = False Then
                            Call GetCellText(iRow, iCol, Text)
                            Size.CX = GetTextSize(iRow, iCol, Text).CX
                            If Size.CX > 0 Then
                                Size.CX = Size.CX + Spacing
                                If Size.CX > .Width Then .Width = Size.CX
                                If Size.CX > EqualSize.CX Then EqualSize.CX = Size.CX
                            End If
                        End If
                    Next iRow
                End If
                End With
            Next iCol
        Case FlexAutoSizeScopeMovable
            For iCol = RowOrCol1 To RowOrCol2 Step IIf(RowOrCol2 >= RowOrCol1, 1, -1)
                With VBFlexGridColsInfo(iCol)
                If (CBool((.State And CLIS_HIDDEN) = CLIS_HIDDEN) Xor ExcludeHidden) Or ExcludeHidden = False Then
                    .Width = -1
                    For iRow = PropFixedRows To (PropRows - 1)
                        If (CBool((VBFlexGridCells.Rows(iRow).RowInfo.State And RWIS_HIDDEN) = RWIS_HIDDEN) Xor ExcludeHidden) Or ExcludeHidden = False Then
                            Call GetCellText(iRow, iCol, Text)
                            Size.CX = GetTextSize(iRow, iCol, Text).CX
                            If Size.CX > 0 Then
                                Size.CX = Size.CX + Spacing
                                If Size.CX > .Width Then .Width = Size.CX
                                If Size.CX > EqualSize.CX Then EqualSize.CX = Size.CX
                            End If
                        End If
                    Next iRow
                End If
                End With
            Next iCol
        Case FlexAutoSizeScopeFrozen
            For iCol = RowOrCol1 To RowOrCol2 Step IIf(RowOrCol2 >= RowOrCol1, 1, -1)
                With VBFlexGridColsInfo(iCol)
                If (CBool((.State And CLIS_HIDDEN) = CLIS_HIDDEN) Xor ExcludeHidden) Or ExcludeHidden = False Then
                    .Width = -1
                    For iRow = PropFixedRows To ((PropFixedRows + PropFrozenRows) - 1)
                        If (CBool((VBFlexGridCells.Rows(iRow).RowInfo.State And RWIS_HIDDEN) = RWIS_HIDDEN) Xor ExcludeHidden) Or ExcludeHidden = False Then
                            Call GetCellText(iRow, iCol, Text)
                            Size.CX = GetTextSize(iRow, iCol, Text).CX
                            If Size.CX > 0 Then
                                Size.CX = Size.CX + Spacing
                                If Size.CX > .Width Then .Width = Size.CX
                                If Size.CX > EqualSize.CX Then EqualSize.CX = Size.CX
                            End If
                        End If
                    Next iRow
                End If
                End With
            Next iCol
    End Select
    If Equal = True Then
        For iCol = RowOrCol1 To RowOrCol2 Step IIf(RowOrCol2 >= RowOrCol1, 1, -1)
            With VBFlexGridColsInfo(iCol)
            If (CBool((.State And CLIS_HIDDEN) = CLIS_HIDDEN) Xor ExcludeHidden) Or ExcludeHidden = False Then .Width = EqualSize.CX
            End With
        Next iCol
    End If
ElseIf Mode = FlexAutoSizeModeRowHeight Then
    Spacing = (ROWINFO_HEIGHT_SPACING_DIP * PixelsPerDIP_Y()) + CLng(UserControl.ScaleY(ExtraSpace, vbTwips, vbPixels))
    EqualSize.CY = -1
    Select Case Scope
        Case FlexAutoSizeScopeAll
            For iRow = RowOrCol1 To RowOrCol2 Step IIf(RowOrCol2 >= RowOrCol1, 1, -1)
                With VBFlexGridCells.Rows(iRow).RowInfo
                If (CBool((.State And RWIS_HIDDEN) = RWIS_HIDDEN) Xor ExcludeHidden) Or ExcludeHidden = False Then
                    .Height = -1
                    For iCol = 0 To (PropCols - 1)
                        If (CBool((VBFlexGridColsInfo(iCol).State And CLIS_HIDDEN) = CLIS_HIDDEN) Xor ExcludeHidden) Or ExcludeHidden = False Then
                            Call GetCellText(iRow, iCol, Text)
                            Size.CY = GetTextHeight(iRow, iCol, Text)
                            If Size.CY > 0 Then
                                Size.CY = Size.CY + Spacing
                                If Size.CY > .Height Then .Height = Size.CY
                                If Size.CY > EqualSize.CY Then EqualSize.CY = Size.CY
                            End If
                        End If
                    Next iCol
                End If
                End With
            Next iRow
        Case FlexAutoSizeScopeFixed
            For iRow = RowOrCol1 To RowOrCol2 Step IIf(RowOrCol2 >= RowOrCol1, 1, -1)
                With VBFlexGridCells.Rows(iRow).RowInfo
                If (CBool((.State And RWIS_HIDDEN) = RWIS_HIDDEN) Xor ExcludeHidden) Or ExcludeHidden = False Then
                    .Height = -1
                    For iCol = 0 To (PropFixedCols - 1)
                        If (CBool((VBFlexGridColsInfo(iCol).State And CLIS_HIDDEN) = CLIS_HIDDEN) Xor ExcludeHidden) Or ExcludeHidden = False Then
                            Call GetCellText(iRow, iCol, Text)
                            Size.CY = GetTextHeight(iRow, iCol, Text)
                            If Size.CY > 0 Then
                                Size.CY = Size.CY + Spacing
                                If Size.CY > .Height Then .Height = Size.CY
                                If Size.CY > EqualSize.CY Then EqualSize.CY = Size.CY
                            End If
                        End If
                    Next iCol
                End If
                End With
            Next iRow
        Case FlexAutoSizeScopeScrollable
            For iRow = RowOrCol1 To RowOrCol2 Step IIf(RowOrCol2 >= RowOrCol1, 1, -1)
                With VBFlexGridCells.Rows(iRow).RowInfo
                If (CBool((.State And RWIS_HIDDEN) = RWIS_HIDDEN) Xor ExcludeHidden) Or ExcludeHidden = False Then
                    .Height = -1
                    For iCol = (PropFixedCols + PropFrozenCols) To (PropCols - 1)
                        If (CBool((VBFlexGridColsInfo(iCol).State And CLIS_HIDDEN) = CLIS_HIDDEN) Xor ExcludeHidden) Or ExcludeHidden = False Then
                            Call GetCellText(iRow, iCol, Text)
                            Size.CY = GetTextHeight(iRow, iCol, Text)
                            If Size.CY > 0 Then
                                Size.CY = Size.CY + Spacing
                                If Size.CY > .Height Then .Height = Size.CY
                                If Size.CY > EqualSize.CY Then EqualSize.CY = Size.CY
                            End If
                        End If
                    Next iCol
                End If
                End With
            Next iRow
        Case FlexAutoSizeScopeMovable
            For iRow = RowOrCol1 To RowOrCol2 Step IIf(RowOrCol2 >= RowOrCol1, 1, -1)
                With VBFlexGridCells.Rows(iRow).RowInfo
                If (CBool((.State And RWIS_HIDDEN) = RWIS_HIDDEN) Xor ExcludeHidden) Or ExcludeHidden = False Then
                    .Height = -1
                    For iCol = PropFixedCols To (PropCols - 1)
                        If (CBool((VBFlexGridColsInfo(iCol).State And CLIS_HIDDEN) = CLIS_HIDDEN) Xor ExcludeHidden) Or ExcludeHidden = False Then
                            Call GetCellText(iRow, iCol, Text)
                            Size.CY = GetTextHeight(iRow, iCol, Text)
                            If Size.CY > 0 Then
                                Size.CY = Size.CY + Spacing
                                If Size.CY > .Height Then .Height = Size.CY
                                If Size.CY > EqualSize.CY Then EqualSize.CY = Size.CY
                            End If
                        End If
                    Next iCol
                End If
                End With
            Next iRow
        Case FlexAutoSizeScopeFrozen
            For iRow = RowOrCol1 To RowOrCol2 Step IIf(RowOrCol2 >= RowOrCol1, 1, -1)
                With VBFlexGridCells.Rows(iRow).RowInfo
                If (CBool((.State And RWIS_HIDDEN) = RWIS_HIDDEN) Xor ExcludeHidden) Or ExcludeHidden = False Then
                    .Height = -1
                    For iCol = PropFixedCols To ((PropFixedCols + PropFrozenCols) - 1)
                        If (CBool((VBFlexGridColsInfo(iCol).State And CLIS_HIDDEN) = CLIS_HIDDEN) Xor ExcludeHidden) Or ExcludeHidden = False Then
                            Call GetCellText(iRow, iCol, Text)
                            Size.CY = GetTextHeight(iRow, iCol, Text)
                            If Size.CY > 0 Then
                                Size.CY = Size.CY + Spacing
                                If Size.CY > .Height Then .Height = Size.CY
                                If Size.CY > EqualSize.CY Then EqualSize.CY = Size.CY
                            End If
                        End If
                    Next iCol
                End If
                End With
            Next iRow
    End Select
    If Equal = True Then
        For iRow = RowOrCol1 To RowOrCol2 Step IIf(RowOrCol2 >= RowOrCol1, 1, -1)
            With VBFlexGridCells.Rows(iRow).RowInfo
            If (CBool((.State And RWIS_HIDDEN) = RWIS_HIDDEN) Xor ExcludeHidden) Or ExcludeHidden = False Then .Height = EqualSize.CY
            End With
        Next iRow
    End If
End If
Dim RCP As TROWCOLPARAMS
With RCP
If Mode = FlexAutoSizeModeColWidth Then
    .Mask = RCPM_LEFTCOL
    .Flags = RCPF_CHECKLEFTCOL
    .LeftCol = VBFlexGridLeftCol
ElseIf Mode = FlexAutoSizeModeRowHeight Then
    .Mask = RCPM_TOPROW
    .Flags = RCPF_CHECKTOPROW
    .TopRow = VBFlexGridTopRow
End If
.Flags = .Flags Or RCPF_SETSCROLLBARS Or RCPF_FORCEREDRAW
Call SetRowColParams(RCP)
End With
End Sub

Public Function TextWidth(ByVal Text As String, Optional ByVal Row As Long = -1, Optional ByVal Col As Long = -1) As Long
Attribute TextWidth.VB_Description = "Returns the text width of the given string using the font of the current or an arbitrary cell (row/col subscripts)."
If Row < -1 Then Err.Raise 380
If Col < -1 Then Err.Raise 380
If Row = -1 Then Row = VBFlexGridRow
If Col = -1 Then Col = VBFlexGridCol
If (Row < 0 Or Row > (PropRows - 1)) Or (Col < 0 Or Col > (PropCols - 1)) Then Err.Raise Number:=381, Description:="Subscript out of range"
Dim Pixels As Long
Pixels = GetTextSize(Row, Col, Text).CX
If Pixels > 0 Then TextWidth = UserControl.ScaleX(Pixels, vbPixels, vbTwips)
End Function

Public Function TextHeight(ByVal Text As String, Optional ByVal Row As Long = -1, Optional ByVal Col As Long = -1) As Long
Attribute TextHeight.VB_Description = "Returns the text height of the given string using the font of the current or an arbitrary cell (row/col subscripts)."
If Row < -1 Then Err.Raise 380
If Col < -1 Then Err.Raise 380
If Row = -1 Then Row = VBFlexGridRow
If Col = -1 Then Col = VBFlexGridCol
If (Row < 0 Or Row > (PropRows - 1)) Or (Col < 0 Or Col > (PropCols - 1)) Then Err.Raise Number:=381, Description:="Subscript out of range"
Dim Pixels As Long
Pixels = GetTextSize(Row, Col, Text).CY
If Pixels > 0 Then TextHeight = UserControl.ScaleY(Pixels, vbPixels, vbTwips)
End Function

Public Property Get Picture() As IPictureDisp
Attribute Picture.VB_Description = "Returns a picture of the flex grid control, suitable for printing, saving to disk, copying to the clipboard, or assigning to a different control."
Attribute Picture.VB_MemberFlags = "400"
If VBFlexGridHandle <> 0 Then
    Dim hDC As Long, hDCBmp As Long
    Dim hBmp As Long, hBmpOld As Long
    hDC = GetDC(VBFlexGridHandle)
    If hDC <> 0 Then
        hDCBmp = CreateCompatibleDC(hDC)
        If hDCBmp <> 0 Then
            Dim RC As RECT, iRow As Long, iCol As Long
            With RC
            .Top = 0
            For iRow = 0 To (PropFixedRows - 1)
                .Bottom = .Bottom + GetRowHeight(iRow)
            Next iRow
            For iRow = VBFlexGridTopRow To (PropRows - 1)
                .Bottom = .Bottom + GetRowHeight(iRow)
            Next iRow
            .Left = 0
            For iCol = 0 To (PropFixedCols - 1)
                .Right = .Right + GetColWidth(iCol)
            Next iCol
            For iCol = VBFlexGridLeftCol To (PropCols - 1)
                .Right = .Right + GetColWidth(iCol)
            Next iCol
            End With
            If PropPictureType = FlexPictureTypeColor Then
                hBmp = CreateCompatibleBitmap(hDC, RC.Right - RC.Left, RC.Bottom - RC.Top)
            ElseIf PropPictureType = FlexPictureTypeMonochrome Then
                hBmp = CreateCompatibleBitmap(hDCBmp, RC.Right - RC.Left, RC.Bottom - RC.Top)
            End If
            If hBmp <> 0 Then
                hBmpOld = SelectObject(hDCBmp, hBmp)
                Call DrawGrid(hDCBmp, -1, True)
                Set Picture = PictureFromHandle(hBmp, vbPicTypeBitmap)
                SelectObject hDCBmp, hBmpOld
            End If
            DeleteDC hDCBmp
        End If
        ReleaseDC VBFlexGridHandle, hDC
    End If
End If
End Property

Public Function StartEdit(Optional ByVal Row As Long = -1, Optional ByVal Col As Long = -1) As Boolean
Attribute StartEdit.VB_Description = "Begins a text editing operation on the current or an arbitrary cell (row/col subscripts)."
If Row < -1 Then Err.Raise 380
If Col < -1 Then Err.Raise 380
If Row = -1 Then Row = VBFlexGridRow
If Col = -1 Then Col = VBFlexGridCol
If (Row < 0 Or Row > (PropRows - 1)) Or (Col < 0 Or Col > (PropCols - 1)) Then Err.Raise Number:=381, Description:="Subscript out of range"
StartEdit = CreateEdit(FlexEditReasonCode, Row, Col)
End Function

Public Function CancelEdit() As Boolean
Attribute CancelEdit.VB_Description = "Ends the text editing operation and discards the changes."
CancelEdit = DestroyEdit(True, FlexEditCloseModeCode)
End Function

Public Function CommitEdit() As Boolean
Attribute CommitEdit.VB_Description = "Ends the text editing operation and saves the changes."
CommitEdit = DestroyEdit(False, FlexEditCloseModeCode)
End Function

Public Property Get EditRow() As Long
Attribute EditRow.VB_Description = "Returns the row bound to a text editing operation."
Attribute EditRow.VB_MemberFlags = "400"
EditRow = VBFlexGridEditRow
End Property

Public Property Get EditCol() As Long
Attribute EditCol.VB_Description = "Returns the column bound to a text editing operation."
Attribute EditCol.VB_MemberFlags = "400"
EditCol = VBFlexGridEditCol
End Property

Public Property Get EditReason() As FlexEditReasonConstants
Attribute EditReason.VB_Description = "Returns a value indicating how the last text editing operation began."
Attribute EditReason.VB_MemberFlags = "400"
EditReason = VBFlexGridEditReason
End Property

Public Property Get EditCloseMode() As FlexEditCloseModeConstants
Attribute EditCloseMode.VB_Description = "Returns a value indicating how the last text editing operation was closed."
Attribute EditCloseMode.VB_MemberFlags = "400"
EditCloseMode = VBFlexGridEditCloseMode
End Property

Public Property Get EditText() As String
Attribute EditText.VB_Description = "Returns/sets the text contained in an object."
Attribute EditText.VB_MemberFlags = "400"
If VBFlexGridEditHandle <> 0 Then
    EditText = String(SendMessage(VBFlexGridEditHandle, WM_GETTEXTLENGTH, 0, ByVal 0&), vbNullChar)
    SendMessage VBFlexGridEditHandle, WM_GETTEXT, Len(EditText) + 1, ByVal StrPtr(EditText)
End If
End Property

Public Property Let EditText(ByVal Value As String)
If VBFlexGridEditHandle <> 0 Then
    Dim MaxLength As Long
    MaxLength = Me.EditMaxLength
    If MaxLength > 0 Then Value = Left$(Value, MaxLength)
    Dim Changed As Boolean
    Changed = CBool(Me.EditText <> Value)
    VBFlexGridEditChangeFrozen = True
    SendMessage VBFlexGridEditHandle, WM_SETTEXT, 0, ByVal StrPtr(Value)
    VBFlexGridEditChangeFrozen = False
    If Changed = True Then
        VBFlexGridEditTextChanged = True
        VBFlexGridEditAlreadyValidated = False
        RaiseEvent EditChange
    End If
End If
End Property

Public Property Get EditMaxLength() As Long
Attribute EditMaxLength.VB_Description = "Returns/sets the maximum number of characters that can be entered in a control."
Attribute EditMaxLength.VB_MemberFlags = "400"
If VBFlexGridEditHandle <> 0 Then EditMaxLength = SendMessage(VBFlexGridEditHandle, EM_GETLIMITTEXT, 0, ByVal 0&)
End Property

Public Property Let EditMaxLength(ByVal Value As Long)
If Value < 0 Then Err.Raise 380
If VBFlexGridEditHandle <> 0 Then SendMessage VBFlexGridEditHandle, EM_SETLIMITTEXT, Value, ByVal 0&
End Property

Public Property Get EditSelStart() As Long
Attribute EditSelStart.VB_Description = "Returns/sets the starting point of text selected; indicates the position of the insertion point if no text is selected."
Attribute EditSelStart.VB_MemberFlags = "400"
If VBFlexGridEditHandle <> 0 Then SendMessage VBFlexGridEditHandle, EM_GETSEL, VarPtr(EditSelStart), ByVal 0&
End Property

Public Property Let EditSelStart(ByVal Value As Long)
If VBFlexGridEditHandle <> 0 Then
    If Value >= 0 Then
        SendMessage VBFlexGridEditHandle, EM_SETSEL, Value, ByVal Value
    Else
        Err.Raise 380
    End If
End If
End Property

Public Property Get EditSelLength() As Long
Attribute EditSelLength.VB_Description = "Returns/sets the number of characters selected."
Attribute EditSelLength.VB_MemberFlags = "400"
If VBFlexGridEditHandle <> 0 Then
    Dim SelStart As Long, SelEnd As Long
    SendMessage VBFlexGridEditHandle, EM_GETSEL, VarPtr(SelStart), ByVal VarPtr(SelEnd)
    EditSelLength = SelEnd - SelStart
End If
End Property

Public Property Let EditSelLength(ByVal Value As Long)
If VBFlexGridEditHandle <> 0 Then
    If Value >= 0 Then
        Dim SelStart As Long
        SendMessage VBFlexGridEditHandle, EM_GETSEL, VarPtr(SelStart), ByVal 0&
        SendMessage VBFlexGridEditHandle, EM_SETSEL, SelStart, ByVal SelStart + Value
    Else
        Err.Raise 380
    End If
End If
End Property

Public Property Get EditSelText() As String
Attribute EditSelText.VB_Description = "Returns/sets the string containing the currently selected text."
Attribute EditSelText.VB_MemberFlags = "400"
If VBFlexGridEditHandle <> 0 Then
    Dim SelStart As Long, SelEnd As Long
    SendMessage VBFlexGridEditHandle, EM_GETSEL, VarPtr(SelStart), ByVal VarPtr(SelEnd)
    On Error Resume Next
    EditSelText = Mid$(Me.EditText, SelStart + 1, (SelEnd - SelStart))
    On Error GoTo 0
End If
End Property

Public Property Let EditSelText(ByVal Value As String)
If VBFlexGridEditHandle <> 0 Then SendMessage VBFlexGridEditHandle, EM_REPLACESEL, 1, ByVal StrPtr(Value)
End Property

Public Property Get ComboMode() As FlexComboModeConstants
Attribute ComboMode.VB_Description = "Returns/sets the combo functionality mode when editing a cell."
Attribute ComboMode.VB_MemberFlags = "400"
ComboMode = VBFlexGridComboMode
End Property

Public Property Let ComboMode(ByVal Value As FlexComboModeConstants)
Select Case Value
    Case FlexComboModeNone, FlexComboModeDropDown, FlexComboModeEditable, FlexComboModeButton
        VBFlexGridComboMode = Value
    Case Else
        Err.Raise 380
End Select
End Property

Public Property Get ComboButtonValue() As FlexComboButtonValueConstants
Attribute ComboButtonValue.VB_Description = "Returns/sets the value of the combo button."
Attribute ComboButtonValue.VB_MemberFlags = "400"
If VBFlexGridComboButtonHandle <> 0 Then
    If IsWindowEnabled(VBFlexGridComboButtonHandle) <> 0 Then
        If ComboButtonGetState(ODS_SELECTED) = False Then
            ComboButtonValue = FlexComboButtonValueUnpressed
        Else
            ComboButtonValue = FlexComboButtonValuePressed
        End If
    Else
        ComboButtonValue = FlexComboButtonValueDisabled
    End If
End If
End Property

Public Property Let ComboButtonValue(ByVal Value As FlexComboButtonValueConstants)
Select Case Value
    Case FlexComboButtonValueUnpressed, FlexComboButtonValuePressed, FlexComboButtonValueDisabled
    Case Else
        Err.Raise 380
End Select
If VBFlexGridComboButtonHandle <> 0 Then
    Select Case Value
        Case FlexComboButtonValueUnpressed
            If IsWindowEnabled(VBFlexGridComboButtonHandle) = 0 Then EnableWindow VBFlexGridComboButtonHandle, 1
            If VBFlexGridComboListHandle <> 0 Then
                Call ComboShowDropDown(False)
            Else
                Call ComboButtonSetState(ODS_SELECTED, False)
            End If
        Case FlexComboButtonValuePressed
            If IsWindowEnabled(VBFlexGridComboButtonHandle) = 0 Then EnableWindow VBFlexGridComboButtonHandle, 1
            If VBFlexGridComboListHandle <> 0 Then
                Call ComboShowDropDown(True)
            Else
                Call ComboButtonPerformClick
            End If
        Case FlexComboButtonValueDisabled
            If VBFlexGridComboListHandle <> 0 Then
                Call ComboShowDropDown(False)
            Else
                Call ComboButtonSetState(ODS_SELECTED, False)
            End If
            EnableWindow VBFlexGridComboButtonHandle, 0
    End Select
End If
End Property

Public Property Get ComboButtonDrawMode() As FlexComboButtonDrawModeConstants
Attribute ComboButtonDrawMode.VB_Description = "Returns/sets a value indicating whether your code or the flex grid will handle drawing of the combo button."
Attribute ComboButtonDrawMode.VB_MemberFlags = "400"
ComboButtonDrawMode = VBFlexGridComboButtonDrawMode
End Property

Public Property Let ComboButtonDrawMode(ByVal Value As FlexComboButtonDrawModeConstants)
Select Case Value
    Case FlexComboButtonDrawModeNormal, FlexComboButtonDrawModeOwnerDraw
        VBFlexGridComboButtonDrawMode = Value
    Case Else
        Err.Raise 380
End Select
End Property

Public Property Get ComboItems() As String
Attribute ComboItems.VB_Description = "Returns/sets the items to be used for the drop-down list when editing a cell."
Attribute ComboItems.VB_MemberFlags = "400"
ComboItems = VBFlexGridComboItems
End Property

Public Property Let ComboItems(ByVal Value As String)
VBFlexGridComboItems = Value
End Property

Public Property Get ComboList(ByVal Index As Long) As String
Attribute ComboList.VB_Description = "Returns the items contained in a drop-down list."
Attribute ComboList.VB_MemberFlags = "400"
If VBFlexGridComboListHandle <> 0 Then
    Dim Length As Long
    Length = SendMessage(VBFlexGridComboListHandle, LB_GETTEXTLEN, Index, ByVal 0&)
    If Not Length = LB_ERR Then
        ComboList = String(Length, vbNullChar)
        SendMessage VBFlexGridComboListHandle, LB_GETTEXT, Index, ByVal StrPtr(ComboList)
    Else
        Err.Raise 5
    End If
End If
End Property

Public Property Get ComboListCount() As Long
Attribute ComboListCount.VB_Description = "Returns the number of items in the drop-down list."
Attribute ComboListCount.VB_MemberFlags = "400"
If VBFlexGridComboListHandle <> 0 Then ComboListCount = SendMessage(VBFlexGridComboListHandle, LB_GETCOUNT, 0, ByVal 0&)
End Property

Public Property Get ComboListIndex() As Long
Attribute ComboListIndex.VB_Description = "Returns/sets the index of the currently selected item in the drop-down list."
Attribute ComboListIndex.VB_MemberFlags = "400"
If VBFlexGridComboListHandle <> 0 Then ComboListIndex = SendMessage(VBFlexGridComboListHandle, LB_GETCURSEL, 0, ByVal 0&)
End Property

Public Property Let ComboListIndex(ByVal Value As Long)
If VBFlexGridComboListHandle <> 0 Then
    If Not Value = -1 Then
        If SendMessage(VBFlexGridComboListHandle, LB_SETCURSEL, Value, ByVal 0&) = LB_ERR Then Err.Raise 380
    Else
        SendMessage VBFlexGridComboListHandle, LB_SETCURSEL, -1, ByVal 0&
    End If
End If
End Property

Public Property Get Version() As Integer
Attribute Version.VB_Description = "Returns the version of the flex grid control currently loaded in memory."
Attribute Version.VB_MemberFlags = "400"
Version = 600
End Property

Private Sub InitFlexGridCells()
If VBFlexGridCellsInit = True Then Exit Sub
If PropRows < 1 Or PropCols < 1 Then
    VBFlexGridRow = -1
    VBFlexGridCol = -1
    VBFlexGridRowSel = -1
    VBFlexGridColSel = -1
    VBFlexGridTopRow = -1
    VBFlexGridLeftCol = -1
    Exit Sub
End If
Dim i As Long, j As Long
ReDim VBFlexGridCells.Rows(0 To (PropRows - 1)) As TCOLS: VBFlexGridCellsInit = True
ReDim VBFlexGridColsInfo(0 To (PropCols - 1)) As TCOLINFO
For i = 0 To (PropRows - 1)
    With VBFlexGridCells.Rows(i)
    ReDim .Cols(0 To (PropCols - 1)) As TCELL
    For j = 0 To (PropCols - 1)
        LSet .Cols(j) = VBFlexGridDefaultCell
    Next j
    LSet .RowInfo = VBFlexGridDefaultRowInfo
    End With
Next i
For i = 0 To (PropCols - 1)
    LSet VBFlexGridColsInfo(i) = VBFlexGridDefaultColInfo
Next i
LSet VBFlexGridDefaultCols = VBFlexGridCells.Rows(0)
VBFlexGridRow = PropFixedRows
VBFlexGridCol = PropFixedCols
If PropAllowSelection = True Then
    Select Case PropSelectionMode
        Case FlexSelectionModeFree, FlexSelectionModeFreeByRow, FlexSelectionModeFreeByColumn
            VBFlexGridRowSel = VBFlexGridRow
            VBFlexGridColSel = VBFlexGridCol
        Case FlexSelectionModeByRow
            VBFlexGridRowSel = VBFlexGridRow
            VBFlexGridColSel = (PropCols - 1)
        Case FlexSelectionModeByColumn
            VBFlexGridRowSel = (PropRows - 1)
            VBFlexGridColSel = VBFlexGridCol
    End Select
Else
    VBFlexGridRowSel = VBFlexGridRow
    VBFlexGridColSel = VBFlexGridCol
End If
VBFlexGridTopRow = PropFixedRows + PropFrozenRows
VBFlexGridLeftCol = PropFixedCols + PropFrozenCols
End Sub

Private Sub EraseFlexGridCells()
If VBFlexGridCellsInit = False Then Exit Sub
Erase VBFlexGridCells.Rows(): VBFlexGridCellsInit = False
Erase VBFlexGridColsInfo()
Erase VBFlexGridDefaultCols.Cols()
VBFlexGridRow = -1
VBFlexGridCol = -1
VBFlexGridRowSel = -1
VBFlexGridColSel = -1
VBFlexGridTopRow = -1
VBFlexGridLeftCol = -1
End Sub

Private Sub RedrawGrid(Optional ByVal UpdateNow As Boolean)
If VBFlexGridHandle <> 0 And VBFlexGridNoRedraw = False Then
    If VBFlexGridDesignMode = False Then
        InvalidateRect VBFlexGridHandle, ByVal 0&, 1
        If UpdateNow = True Then UpdateWindow VBFlexGridHandle
    Else
        UserControl.Refresh
    End If
End If
End Sub

Private Sub DrawGrid(ByVal hDC As Long, ByRef hRgn As Long, Optional ByVal NoClip As Boolean)
If VBFlexGridNoRedraw = True And hDC <> 0 Then
    If hRgn <> -1 Then hRgn = CreateRectRgn(0, 0, 0, 0)
    Exit Sub
ElseIf hDC = 0 Then
    Exit Sub
End If
If VBFlexGridHandle = 0 Or (PropRows < 1 Or PropCols < 1) Then Exit Sub
Dim iRow As Long, iCol As Long, FixedCX As Long, FixedCY As Long, FrozenCX As Long, FrozenCY As Long
Dim CellRect As RECT, GridRect As RECT
Dim OldBkMode As Long, hFontOld As Long, Brush As Long
Call GetSelRangeStruct(VBFlexGridDrawInfo.SelRange)
VBFlexGridDrawInfo.CellTextWidthPadding = CELL_TEXT_WIDTH_PADDING_DIP * PixelsPerDIP_X()
VBFlexGridDrawInfo.CellTextHeightPadding = CELL_TEXT_HEIGHT_PADDING_DIP * PixelsPerDIP_Y()
For iCol = 0 To (PropFixedCols - 1)
    FixedCX = FixedCX + GetColWidth(iCol)
Next iCol
For iRow = 0 To (PropFixedRows - 1)
    FixedCY = FixedCY + GetRowHeight(iRow)
Next iRow
For iCol = PropFixedCols To ((PropFixedCols + PropFrozenCols) - 1)
    FrozenCX = FrozenCX + GetColWidth(iCol)
Next iCol
For iRow = PropFixedRows To ((PropFixedRows + PropFrozenRows) - 1)
    FrozenCY = FrozenCY + GetRowHeight(iRow)
Next iRow
OldBkMode = SetBkMode(hDC, 1)
With CellRect
If PropMergeCells = FlexMergeCellsNever Then
    If VBFlexGridFontFixedHandle = 0 Then
        hFontOld = SelectObject(hDC, VBFlexGridFontHandle)
    Else
        hFontOld = SelectObject(hDC, VBFlexGridFontFixedHandle)
    End If
    Brush = SelectObject(hDC, VBFlexGridBackColorFixedBrush)
    For iRow = 0 To (PropFixedRows - 1)
        .Bottom = .Top + GetRowHeight(iRow)
        If .Bottom > .Top Then
            VBFlexGridDrawInfo.GridLinePoints(0).Y = .Bottom - 1
            VBFlexGridDrawInfo.GridLinePoints(1).Y = .Bottom - 1
            VBFlexGridDrawInfo.GridLinePoints(2).Y = .Top - 1
            VBFlexGridDrawInfo.GridLinePoints(3).Y = .Top
            VBFlexGridDrawInfo.GridLinePoints(4).Y = .Top
            VBFlexGridDrawInfo.GridLinePoints(5).Y = .Bottom
            If PropFrozenCols > 0 Then
                .Left = FixedCX
                For iCol = PropFixedCols To ((PropFixedCols + PropFrozenCols) - 1)
                    .Right = .Left + GetColWidth(iCol)
                    If .Right > .Left Then
                        VBFlexGridDrawInfo.GridLinePoints(0).X = .Left
                        VBFlexGridDrawInfo.GridLinePoints(1).X = .Right - 1
                        VBFlexGridDrawInfo.GridLinePoints(2).X = .Right - 1
                        VBFlexGridDrawInfo.GridLinePoints(3).X = .Right - 2
                        VBFlexGridDrawInfo.GridLinePoints(4).X = .Left
                        VBFlexGridDrawInfo.GridLinePoints(5).X = .Left
                        Call DrawFixedCell(hDC, CellRect, iRow, iCol)
                    End If
                    .Left = .Right
                    If NoClip = False And .Right > VBFlexGridClientRect.Right Then Exit For
                Next iCol
            End If
            .Left = FixedCX + FrozenCX
            For iCol = VBFlexGridLeftCol To (PropCols - 1)
                .Right = .Left + GetColWidth(iCol)
                If .Right > .Left Then
                    VBFlexGridDrawInfo.GridLinePoints(0).X = .Left
                    VBFlexGridDrawInfo.GridLinePoints(1).X = .Right - 1
                    VBFlexGridDrawInfo.GridLinePoints(2).X = .Right - 1
                    VBFlexGridDrawInfo.GridLinePoints(3).X = .Right - 2
                    VBFlexGridDrawInfo.GridLinePoints(4).X = .Left
                    VBFlexGridDrawInfo.GridLinePoints(5).X = .Left
                    Call DrawFixedCell(hDC, CellRect, iRow, iCol)
                End If
                .Left = .Right
                If NoClip = False And .Right > VBFlexGridClientRect.Right Then Exit For
            Next iCol
            If .Right > GridRect.Right Then GridRect.Right = .Right
        End If
        .Top = .Bottom
        If NoClip = False And .Bottom > VBFlexGridClientRect.Bottom Then Exit For
    Next iRow
    If .Bottom > GridRect.Bottom Then GridRect.Bottom = .Bottom
    If PropFixedRows > 0 And PropFixedCols > 0 Then
        .Top = 0
        For iRow = 0 To (PropFixedRows - 1)
            .Bottom = .Top + GetRowHeight(iRow)
            If .Bottom > .Top Then
                VBFlexGridDrawInfo.GridLinePoints(0).Y = .Bottom - 1
                VBFlexGridDrawInfo.GridLinePoints(1).Y = .Bottom - 1
                VBFlexGridDrawInfo.GridLinePoints(2).Y = .Top - 1
                VBFlexGridDrawInfo.GridLinePoints(3).Y = .Top
                VBFlexGridDrawInfo.GridLinePoints(4).Y = .Top
                VBFlexGridDrawInfo.GridLinePoints(5).Y = .Bottom
                .Left = 0
                For iCol = 0 To (PropFixedCols - 1)
                    .Right = .Left + GetColWidth(iCol)
                    If .Right > .Left Then
                        VBFlexGridDrawInfo.GridLinePoints(0).X = .Left
                        VBFlexGridDrawInfo.GridLinePoints(1).X = .Right - 1
                        VBFlexGridDrawInfo.GridLinePoints(2).X = .Right - 1
                        VBFlexGridDrawInfo.GridLinePoints(3).X = .Right - 2
                        VBFlexGridDrawInfo.GridLinePoints(4).X = .Left
                        VBFlexGridDrawInfo.GridLinePoints(5).X = .Left
                        Call DrawFixedCell(hDC, CellRect, iRow, iCol)
                    End If
                    .Left = .Right
                    If NoClip = False And .Right > VBFlexGridClientRect.Right Then Exit For
                Next iCol
            End If
            .Top = .Bottom
            If NoClip = False And .Bottom > VBFlexGridClientRect.Bottom Then Exit For
        Next iRow
    End If
    If PropFixedCols > 0 Then
        .Top = FixedCY
        If PropFrozenRows > 0 Then
            For iRow = PropFixedRows To ((PropFixedRows + PropFrozenRows) - 1)
                .Bottom = .Top + GetRowHeight(iRow)
                If .Bottom > .Top Then
                    VBFlexGridDrawInfo.GridLinePoints(0).Y = .Bottom - 1
                    VBFlexGridDrawInfo.GridLinePoints(1).Y = .Bottom - 1
                    VBFlexGridDrawInfo.GridLinePoints(2).Y = .Top - 1
                    VBFlexGridDrawInfo.GridLinePoints(3).Y = .Top
                    VBFlexGridDrawInfo.GridLinePoints(4).Y = .Top
                    VBFlexGridDrawInfo.GridLinePoints(5).Y = .Bottom
                    .Left = 0
                    For iCol = 0 To (PropFixedCols - 1)
                        .Right = .Left + GetColWidth(iCol)
                        If .Right > .Left Then
                            VBFlexGridDrawInfo.GridLinePoints(0).X = .Left
                            VBFlexGridDrawInfo.GridLinePoints(1).X = .Right - 1
                            VBFlexGridDrawInfo.GridLinePoints(2).X = .Right - 1
                            VBFlexGridDrawInfo.GridLinePoints(3).X = .Right - 2
                            VBFlexGridDrawInfo.GridLinePoints(4).X = .Left
                            VBFlexGridDrawInfo.GridLinePoints(5).X = .Left
                            Call DrawFixedCell(hDC, CellRect, iRow, iCol)
                        End If
                        .Left = .Right
                        If NoClip = False And .Right > VBFlexGridClientRect.Right Then Exit For
                    Next iCol
                End If
                .Top = .Bottom
                If NoClip = False And .Bottom > VBFlexGridClientRect.Bottom Then Exit For
            Next iRow
        End If
        .Top = FixedCY + FrozenCY
        For iRow = VBFlexGridTopRow To (PropRows - 1)
            .Bottom = .Top + GetRowHeight(iRow)
            If .Bottom > .Top Then
                VBFlexGridDrawInfo.GridLinePoints(0).Y = .Bottom - 1
                VBFlexGridDrawInfo.GridLinePoints(1).Y = .Bottom - 1
                VBFlexGridDrawInfo.GridLinePoints(2).Y = .Top - 1
                VBFlexGridDrawInfo.GridLinePoints(3).Y = .Top
                VBFlexGridDrawInfo.GridLinePoints(4).Y = .Top
                VBFlexGridDrawInfo.GridLinePoints(5).Y = .Bottom
                .Left = 0
                For iCol = 0 To (PropFixedCols - 1)
                    .Right = .Left + GetColWidth(iCol)
                    If .Right > .Left Then
                        VBFlexGridDrawInfo.GridLinePoints(0).X = .Left
                        VBFlexGridDrawInfo.GridLinePoints(1).X = .Right - 1
                        VBFlexGridDrawInfo.GridLinePoints(2).X = .Right - 1
                        VBFlexGridDrawInfo.GridLinePoints(3).X = .Right - 2
                        VBFlexGridDrawInfo.GridLinePoints(4).X = .Left
                        VBFlexGridDrawInfo.GridLinePoints(5).X = .Left
                        Call DrawFixedCell(hDC, CellRect, iRow, iCol)
                    End If
                    .Left = .Right
                    If NoClip = False And .Right > VBFlexGridClientRect.Right Then Exit For
                Next iCol
            End If
            .Top = .Bottom
            If NoClip = False And .Bottom > VBFlexGridClientRect.Bottom Then Exit For
        Next iRow
    End If
    If VBFlexGridFontFixedHandle <> 0 Then
        If hFontOld <> 0 Then
            SelectObject hDC, hFontOld
            hFontOld = 0
        End If
        hFontOld = SelectObject(hDC, VBFlexGridFontHandle)
    End If
    SelectObject hDC, VBFlexGridBackColorBrush
    If PropFrozenRows > 0 Then
        .Top = FixedCY
        For iRow = PropFixedRows To ((PropFixedRows + PropFrozenRows) - 1)
            .Bottom = .Top + GetRowHeight(iRow)
            If .Bottom > .Top Then
                VBFlexGridDrawInfo.GridLinePoints(0).Y = .Bottom - 1
                VBFlexGridDrawInfo.GridLinePoints(1).Y = .Bottom - 1
                VBFlexGridDrawInfo.GridLinePoints(2).Y = .Top - 1
                VBFlexGridDrawInfo.GridLinePoints(3).Y = .Top
                VBFlexGridDrawInfo.GridLinePoints(4).Y = .Top
                VBFlexGridDrawInfo.GridLinePoints(5).Y = .Bottom
                If PropFrozenCols > 0 Then
                    .Left = FixedCX
                    For iCol = PropFixedCols To ((PropFixedCols + PropFrozenCols) - 1)
                        .Right = .Left + GetColWidth(iCol)
                        If .Right > .Left Then
                            VBFlexGridDrawInfo.GridLinePoints(0).X = .Left
                            VBFlexGridDrawInfo.GridLinePoints(1).X = .Right - 1
                            VBFlexGridDrawInfo.GridLinePoints(2).X = .Right - 1
                            VBFlexGridDrawInfo.GridLinePoints(3).X = .Right - 2
                            VBFlexGridDrawInfo.GridLinePoints(4).X = .Left
                            VBFlexGridDrawInfo.GridLinePoints(5).X = .Left
                            Call DrawCell(hDC, CellRect, iRow, iCol)
                        End If
                        .Left = .Right
                        If NoClip = False And .Right > VBFlexGridClientRect.Right Then Exit For
                    Next iCol
                End If
                .Left = FixedCX + FrozenCX
                For iCol = VBFlexGridLeftCol To (PropCols - 1)
                    .Right = .Left + GetColWidth(iCol)
                    If .Right > .Left Then
                        VBFlexGridDrawInfo.GridLinePoints(0).X = .Left
                        VBFlexGridDrawInfo.GridLinePoints(1).X = .Right - 1
                        VBFlexGridDrawInfo.GridLinePoints(2).X = .Right - 1
                        VBFlexGridDrawInfo.GridLinePoints(3).X = .Right - 2
                        VBFlexGridDrawInfo.GridLinePoints(4).X = .Left
                        VBFlexGridDrawInfo.GridLinePoints(5).X = .Left
                        Call DrawCell(hDC, CellRect, iRow, iCol)
                    End If
                    .Left = .Right
                    If NoClip = False And .Right > VBFlexGridClientRect.Right Then Exit For
                Next iCol
                If .Right > GridRect.Right Then GridRect.Right = .Right
            End If
            .Top = .Bottom
            If NoClip = False And .Bottom > VBFlexGridClientRect.Bottom Then Exit For
        Next iRow
    End If
    .Top = FixedCY + FrozenCY
    For iRow = VBFlexGridTopRow To (PropRows - 1)
        .Bottom = .Top + GetRowHeight(iRow)
        If .Bottom > .Top Then
            VBFlexGridDrawInfo.GridLinePoints(0).Y = .Bottom - 1
            VBFlexGridDrawInfo.GridLinePoints(1).Y = .Bottom - 1
            VBFlexGridDrawInfo.GridLinePoints(2).Y = .Top - 1
            VBFlexGridDrawInfo.GridLinePoints(3).Y = .Top
            VBFlexGridDrawInfo.GridLinePoints(4).Y = .Top
            VBFlexGridDrawInfo.GridLinePoints(5).Y = .Bottom
            If PropFrozenCols > 0 Then
                .Left = FixedCX
                For iCol = PropFixedCols To ((PropFixedCols + PropFrozenCols) - 1)
                    .Right = .Left + GetColWidth(iCol)
                    If .Right > .Left Then
                        VBFlexGridDrawInfo.GridLinePoints(0).X = .Left
                        VBFlexGridDrawInfo.GridLinePoints(1).X = .Right - 1
                        VBFlexGridDrawInfo.GridLinePoints(2).X = .Right - 1
                        VBFlexGridDrawInfo.GridLinePoints(3).X = .Right - 2
                        VBFlexGridDrawInfo.GridLinePoints(4).X = .Left
                        VBFlexGridDrawInfo.GridLinePoints(5).X = .Left
                        Call DrawCell(hDC, CellRect, iRow, iCol)
                    End If
                    .Left = .Right
                    If NoClip = False And .Right > VBFlexGridClientRect.Right Then Exit For
                Next iCol
            End If
            .Left = FixedCX + FrozenCX
            For iCol = VBFlexGridLeftCol To (PropCols - 1)
                .Right = .Left + GetColWidth(iCol)
                If .Right > .Left Then
                    VBFlexGridDrawInfo.GridLinePoints(0).X = .Left
                    VBFlexGridDrawInfo.GridLinePoints(1).X = .Right - 1
                    VBFlexGridDrawInfo.GridLinePoints(2).X = .Right - 1
                    VBFlexGridDrawInfo.GridLinePoints(3).X = .Right - 2
                    VBFlexGridDrawInfo.GridLinePoints(4).X = .Left
                    VBFlexGridDrawInfo.GridLinePoints(5).X = .Left
                    Call DrawCell(hDC, CellRect, iRow, iCol)
                End If
                .Left = .Right
                If NoClip = False And .Right > VBFlexGridClientRect.Right Then Exit For
            Next iCol
            If .Right > GridRect.Right Then GridRect.Right = .Right
        End If
        .Top = .Bottom
        If NoClip = False And .Bottom > VBFlexGridClientRect.Bottom Then Exit For
    Next iRow
    If .Bottom > GridRect.Bottom Then GridRect.Bottom = .Bottom
    If hFontOld <> 0 Then
        SelectObject hDC, hFontOld
        hFontOld = 0
    End If
    If Brush <> 0 Then
        SelectObject hDC, Brush
        Brush = 0
    End If
Else
    If VBFlexGridFontFixedHandle = 0 Then
        hFontOld = SelectObject(hDC, VBFlexGridFontHandle)
    Else
        hFontOld = SelectObject(hDC, VBFlexGridFontFixedHandle)
    End If
    Brush = SelectObject(hDC, VBFlexGridBackColorFixedBrush)
    ReDim VBFlexGridMergeDrawInfo.Row.Cols(0 To (PropCols - 1)) As TMERGEDRAWCOLINFO
    For iRow = 0 To (PropFixedRows - 1)
        VBFlexGridMergeDrawInfo.Row.ColOffset = 0
        VBFlexGridMergeDrawInfo.Row.Width = 0
        .Bottom = .Top + GetRowHeight(iRow)
        If PropFrozenCols > 0 Then
            .Left = FixedCX
            For iCol = PropFixedCols To ((PropFixedCols + PropFrozenCols) - 1)
                .Right = .Left + GetColWidth(iCol)
                If (VBFlexGridCells.Rows(iRow).RowInfo.State And RWIS_MERGE) = RWIS_MERGE Then
                    If iCol > VBFlexGridLeftCol Then
                        Select Case PropMergeCells
                            Case FlexMergeCellsFree, FlexMergeCellsRestrictRows, FlexMergeCellsFixedOnly
                                If MergeCompareFunction(iRow, iCol, iRow, iCol - 1) = True Then
                                    VBFlexGridMergeDrawInfo.Row.ColOffset = VBFlexGridMergeDrawInfo.Row.ColOffset + 1
                                    VBFlexGridMergeDrawInfo.Row.Width = VBFlexGridMergeDrawInfo.Row.Width + GetColWidth(iCol - 1)
                                Else
                                    VBFlexGridMergeDrawInfo.Row.ColOffset = 0
                                    VBFlexGridMergeDrawInfo.Row.Width = 0
                                End If
                            Case FlexMergeCellsRestrictColumns, FlexMergeCellsRestrictAll
                                If MergeCompareFunction(iRow, iCol, iRow, iCol - 1) = True Then
                                    If iRow > VBFlexGridTopRow Then
                                        If MergeCompareFunction(iRow - 1, iCol, iRow - 1, iCol - 1) = True Then
                                            VBFlexGridMergeDrawInfo.Row.ColOffset = VBFlexGridMergeDrawInfo.Row.ColOffset + 1
                                            VBFlexGridMergeDrawInfo.Row.Width = VBFlexGridMergeDrawInfo.Row.Width + GetColWidth(iCol - 1)
                                        Else
                                            VBFlexGridMergeDrawInfo.Row.ColOffset = 0
                                            VBFlexGridMergeDrawInfo.Row.Width = 0
                                        End If
                                    Else
                                        VBFlexGridMergeDrawInfo.Row.ColOffset = VBFlexGridMergeDrawInfo.Row.ColOffset + 1
                                        VBFlexGridMergeDrawInfo.Row.Width = VBFlexGridMergeDrawInfo.Row.Width + GetColWidth(iCol - 1)
                                    End If
                                Else
                                    VBFlexGridMergeDrawInfo.Row.ColOffset = 0
                                    VBFlexGridMergeDrawInfo.Row.Width = 0
                                End If
                        End Select
                    Else
                        VBFlexGridMergeDrawInfo.Row.ColOffset = 0
                        VBFlexGridMergeDrawInfo.Row.Width = 0
                    End If
                End If
                If (VBFlexGridColsInfo(iCol).State And CLIS_MERGE) = CLIS_MERGE Then
                    If iRow > 0 Then
                        Select Case PropMergeCells
                            Case FlexMergeCellsFree, FlexMergeCellsRestrictColumns, FlexMergeCellsFixedOnly
                                If MergeCompareFunction(iRow, iCol, iRow - 1, iCol) = True Then
                                    VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset + 1
                                    VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height + GetRowHeight(iRow - 1)
                                Else
                                    VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = 0
                                    VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = 0
                                End If
                            Case FlexMergeCellsRestrictRows, FlexMergeCellsRestrictAll
                                If MergeCompareFunction(iRow, iCol, iRow - 1, iCol) = True Then
                                    If iCol > 0 Then
                                        If MergeCompareFunction(iRow, iCol - 1, iRow - 1, iCol - 1) = True Then
                                            VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset + 1
                                            VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height + GetRowHeight(iRow - 1)
                                        Else
                                            VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = 0
                                            VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = 0
                                        End If
                                    Else
                                        VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset + 1
                                        VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height + GetRowHeight(iRow - 1)
                                    End If
                                Else
                                    VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = 0
                                    VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = 0
                                End If
                        End Select
                    Else
                        VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = 0
                        VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = 0
                    End If
                End If
                .Left = .Left - VBFlexGridMergeDrawInfo.Row.Width
                .Top = .Top - VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height
                If .Bottom > .Top And .Right > .Left Then
                    VBFlexGridDrawInfo.GridLinePoints(0).X = .Left
                    VBFlexGridDrawInfo.GridLinePoints(0).Y = .Bottom - 1
                    VBFlexGridDrawInfo.GridLinePoints(1).X = .Right - 1
                    VBFlexGridDrawInfo.GridLinePoints(1).Y = .Bottom - 1
                    VBFlexGridDrawInfo.GridLinePoints(2).X = .Right - 1
                    VBFlexGridDrawInfo.GridLinePoints(2).Y = .Top - 1
                    VBFlexGridDrawInfo.GridLinePoints(3).X = .Right - 2
                    VBFlexGridDrawInfo.GridLinePoints(3).Y = .Top
                    VBFlexGridDrawInfo.GridLinePoints(4).X = .Left
                    VBFlexGridDrawInfo.GridLinePoints(4).Y = .Top
                    VBFlexGridDrawInfo.GridLinePoints(5).X = .Left
                    VBFlexGridDrawInfo.GridLinePoints(5).Y = .Bottom
                    Call DrawFixedCell(hDC, CellRect, iRow, iCol)
                End If
                .Left = .Left + VBFlexGridMergeDrawInfo.Row.Width
                .Top = .Top + VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height
                .Left = .Right
                If NoClip = False And .Right > VBFlexGridClientRect.Right Then Exit For
            Next iCol
        End If
        .Left = FixedCX + FrozenCX
        For iCol = VBFlexGridLeftCol To (PropCols - 1)
            .Right = .Left + GetColWidth(iCol)
            If (VBFlexGridCells.Rows(iRow).RowInfo.State And RWIS_MERGE) = RWIS_MERGE Then
                If iCol > VBFlexGridLeftCol Then
                    Select Case PropMergeCells
                        Case FlexMergeCellsFree, FlexMergeCellsRestrictRows, FlexMergeCellsFixedOnly
                            If MergeCompareFunction(iRow, iCol, iRow, iCol - 1) = True Then
                                VBFlexGridMergeDrawInfo.Row.ColOffset = VBFlexGridMergeDrawInfo.Row.ColOffset + 1
                                VBFlexGridMergeDrawInfo.Row.Width = VBFlexGridMergeDrawInfo.Row.Width + GetColWidth(iCol - 1)
                            Else
                                VBFlexGridMergeDrawInfo.Row.ColOffset = 0
                                VBFlexGridMergeDrawInfo.Row.Width = 0
                            End If
                        Case FlexMergeCellsRestrictColumns, FlexMergeCellsRestrictAll
                            If MergeCompareFunction(iRow, iCol, iRow, iCol - 1) = True Then
                                If iRow > VBFlexGridTopRow Then
                                    If MergeCompareFunction(iRow - 1, iCol, iRow - 1, iCol - 1) = True Then
                                        VBFlexGridMergeDrawInfo.Row.ColOffset = VBFlexGridMergeDrawInfo.Row.ColOffset + 1
                                        VBFlexGridMergeDrawInfo.Row.Width = VBFlexGridMergeDrawInfo.Row.Width + GetColWidth(iCol - 1)
                                    Else
                                        VBFlexGridMergeDrawInfo.Row.ColOffset = 0
                                        VBFlexGridMergeDrawInfo.Row.Width = 0
                                    End If
                                Else
                                    VBFlexGridMergeDrawInfo.Row.ColOffset = VBFlexGridMergeDrawInfo.Row.ColOffset + 1
                                    VBFlexGridMergeDrawInfo.Row.Width = VBFlexGridMergeDrawInfo.Row.Width + GetColWidth(iCol - 1)
                                End If
                            Else
                                VBFlexGridMergeDrawInfo.Row.ColOffset = 0
                                VBFlexGridMergeDrawInfo.Row.Width = 0
                            End If
                    End Select
                Else
                    VBFlexGridMergeDrawInfo.Row.ColOffset = 0
                    VBFlexGridMergeDrawInfo.Row.Width = 0
                End If
            End If
            If (VBFlexGridColsInfo(iCol).State And CLIS_MERGE) = CLIS_MERGE Then
                If iRow > 0 Then
                    Select Case PropMergeCells
                        Case FlexMergeCellsFree, FlexMergeCellsRestrictColumns, FlexMergeCellsFixedOnly
                            If MergeCompareFunction(iRow, iCol, iRow - 1, iCol) = True Then
                                VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset + 1
                                VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height + GetRowHeight(iRow - 1)
                            Else
                                VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = 0
                                VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = 0
                            End If
                        Case FlexMergeCellsRestrictRows, FlexMergeCellsRestrictAll
                            If MergeCompareFunction(iRow, iCol, iRow - 1, iCol) = True Then
                                If iCol > 0 Then
                                    If MergeCompareFunction(iRow, iCol - 1, iRow - 1, iCol - 1) = True Then
                                        VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset + 1
                                        VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height + GetRowHeight(iRow - 1)
                                    Else
                                        VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = 0
                                        VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = 0
                                    End If
                                Else
                                    VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset + 1
                                    VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height + GetRowHeight(iRow - 1)
                                End If
                            Else
                                VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = 0
                                VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = 0
                            End If
                    End Select
                Else
                    VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = 0
                    VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = 0
                End If
            End If
            .Left = .Left - VBFlexGridMergeDrawInfo.Row.Width
            .Top = .Top - VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height
            If .Bottom > .Top And .Right > .Left Then
                VBFlexGridDrawInfo.GridLinePoints(0).X = .Left
                VBFlexGridDrawInfo.GridLinePoints(0).Y = .Bottom - 1
                VBFlexGridDrawInfo.GridLinePoints(1).X = .Right - 1
                VBFlexGridDrawInfo.GridLinePoints(1).Y = .Bottom - 1
                VBFlexGridDrawInfo.GridLinePoints(2).X = .Right - 1
                VBFlexGridDrawInfo.GridLinePoints(2).Y = .Top - 1
                VBFlexGridDrawInfo.GridLinePoints(3).X = .Right - 2
                VBFlexGridDrawInfo.GridLinePoints(3).Y = .Top
                VBFlexGridDrawInfo.GridLinePoints(4).X = .Left
                VBFlexGridDrawInfo.GridLinePoints(4).Y = .Top
                VBFlexGridDrawInfo.GridLinePoints(5).X = .Left
                VBFlexGridDrawInfo.GridLinePoints(5).Y = .Bottom
                Call DrawFixedCell(hDC, CellRect, iRow, iCol)
            End If
            .Left = .Left + VBFlexGridMergeDrawInfo.Row.Width
            .Top = .Top + VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height
            .Left = .Right
            If NoClip = False And .Right > VBFlexGridClientRect.Right Then Exit For
        Next iCol
        If .Right > GridRect.Right Then GridRect.Right = .Right
        .Top = .Bottom
        If NoClip = False And .Bottom > VBFlexGridClientRect.Bottom Then Exit For
    Next iRow
    If .Bottom > GridRect.Bottom Then GridRect.Bottom = .Bottom
    If PropFixedRows > 0 And PropFixedCols > 0 Then
        ReDim VBFlexGridMergeDrawInfo.Row.Cols(0 To (PropFixedCols - 1)) As TMERGEDRAWCOLINFO
        .Top = 0
        For iRow = 0 To (PropFixedRows - 1)
            VBFlexGridMergeDrawInfo.Row.ColOffset = 0
            VBFlexGridMergeDrawInfo.Row.Width = 0
            .Bottom = .Top + GetRowHeight(iRow)
            .Left = 0
            For iCol = 0 To (PropFixedCols - 1)
                .Right = .Left + GetColWidth(iCol)
                If (VBFlexGridCells.Rows(iRow).RowInfo.State And RWIS_MERGE) = RWIS_MERGE Then
                    If iCol > 0 Then
                        Select Case PropMergeCells
                            Case FlexMergeCellsFree, FlexMergeCellsRestrictRows, FlexMergeCellsFixedOnly
                                If MergeCompareFunction(iRow, iCol, iRow, iCol - 1) = True Then
                                    VBFlexGridMergeDrawInfo.Row.ColOffset = VBFlexGridMergeDrawInfo.Row.ColOffset + 1
                                    VBFlexGridMergeDrawInfo.Row.Width = VBFlexGridMergeDrawInfo.Row.Width + GetColWidth(iCol - 1)
                                Else
                                    VBFlexGridMergeDrawInfo.Row.ColOffset = 0
                                    VBFlexGridMergeDrawInfo.Row.Width = 0
                                End If
                            Case FlexMergeCellsRestrictColumns, FlexMergeCellsRestrictAll
                                If MergeCompareFunction(iRow, iCol, iRow, iCol - 1) = True Then
                                    If iRow > VBFlexGridTopRow Then
                                        If MergeCompareFunction(iRow - 1, iCol, iRow - 1, iCol - 1) = True Then
                                            VBFlexGridMergeDrawInfo.Row.ColOffset = VBFlexGridMergeDrawInfo.Row.ColOffset + 1
                                            VBFlexGridMergeDrawInfo.Row.Width = VBFlexGridMergeDrawInfo.Row.Width + GetColWidth(iCol - 1)
                                        Else
                                            VBFlexGridMergeDrawInfo.Row.ColOffset = 0
                                            VBFlexGridMergeDrawInfo.Row.Width = 0
                                        End If
                                    Else
                                        VBFlexGridMergeDrawInfo.Row.ColOffset = VBFlexGridMergeDrawInfo.Row.ColOffset + 1
                                        VBFlexGridMergeDrawInfo.Row.Width = VBFlexGridMergeDrawInfo.Row.Width + GetColWidth(iCol - 1)
                                    End If
                                Else
                                    VBFlexGridMergeDrawInfo.Row.ColOffset = 0
                                    VBFlexGridMergeDrawInfo.Row.Width = 0
                                End If
                        End Select
                    Else
                        VBFlexGridMergeDrawInfo.Row.ColOffset = 0
                        VBFlexGridMergeDrawInfo.Row.Width = 0
                    End If
                End If
                If (VBFlexGridColsInfo(iCol).State And CLIS_MERGE) = CLIS_MERGE Then
                    If iRow > 0 Then
                        Select Case PropMergeCells
                            Case FlexMergeCellsFree, FlexMergeCellsRestrictColumns, FlexMergeCellsFixedOnly
                                If MergeCompareFunction(iRow, iCol, iRow - 1, iCol) = True Then
                                    VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset + 1
                                    VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height + GetRowHeight(iRow - 1)
                                Else
                                    VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = 0
                                    VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = 0
                                End If
                            Case FlexMergeCellsRestrictRows, FlexMergeCellsRestrictAll
                                If MergeCompareFunction(iRow, iCol, iRow - 1, iCol) = True Then
                                    If iCol > 0 Then
                                        If MergeCompareFunction(iRow, iCol - 1, iRow - 1, iCol - 1) = True Then
                                            VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset + 1
                                            VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height + GetRowHeight(iRow - 1)
                                        Else
                                            VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = 0
                                            VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = 0
                                        End If
                                    Else
                                        VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset + 1
                                        VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height + GetRowHeight(iRow - 1)
                                    End If
                                Else
                                    VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = 0
                                    VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = 0
                                End If
                        End Select
                    Else
                        VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = 0
                        VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = 0
                    End If
                End If
                .Left = .Left - VBFlexGridMergeDrawInfo.Row.Width
                .Top = .Top - VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height
                If .Bottom > .Top And .Right > .Left Then
                    VBFlexGridDrawInfo.GridLinePoints(0).X = .Left
                    VBFlexGridDrawInfo.GridLinePoints(0).Y = .Bottom - 1
                    VBFlexGridDrawInfo.GridLinePoints(1).X = .Right - 1
                    VBFlexGridDrawInfo.GridLinePoints(1).Y = .Bottom - 1
                    VBFlexGridDrawInfo.GridLinePoints(2).X = .Right - 1
                    VBFlexGridDrawInfo.GridLinePoints(2).Y = .Top - 1
                    VBFlexGridDrawInfo.GridLinePoints(3).X = .Right - 2
                    VBFlexGridDrawInfo.GridLinePoints(3).Y = .Top
                    VBFlexGridDrawInfo.GridLinePoints(4).X = .Left
                    VBFlexGridDrawInfo.GridLinePoints(4).Y = .Top
                    VBFlexGridDrawInfo.GridLinePoints(5).X = .Left
                    VBFlexGridDrawInfo.GridLinePoints(5).Y = .Bottom
                    Call DrawFixedCell(hDC, CellRect, iRow, iCol)
                End If
                .Left = .Left + VBFlexGridMergeDrawInfo.Row.Width
                .Top = .Top + VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height
                .Left = .Right
                If NoClip = False And .Right > VBFlexGridClientRect.Right Then Exit For
            Next iCol
            .Top = .Bottom
            If NoClip = False And .Bottom > VBFlexGridClientRect.Bottom Then Exit For
        Next iRow
    End If
    If PropFixedCols > 0 Then
        ReDim VBFlexGridMergeDrawInfo.Row.Cols(0 To (PropFixedCols - 1)) As TMERGEDRAWCOLINFO
        .Top = FixedCY ' .Top = FixedCY + FrozenCY
        For iRow = PropFixedRows To (PropRows - 1) ' For iRow = VBFlexGridTopRow To (PropRows - 1)
            ' Below trick will save another loop with just the same code.
            ' For iRow = PropFixedRows To ((PropFixedRows + PropFrozenRows) - 1)
            If iRow >= (PropFixedRows + PropFrozenRows) Then
                If iRow < VBFlexGridTopRow Then iRow = VBFlexGridTopRow
            End If
            VBFlexGridMergeDrawInfo.Row.ColOffset = 0
            VBFlexGridMergeDrawInfo.Row.Width = 0
            .Bottom = .Top + GetRowHeight(iRow)
            .Left = 0
            For iCol = 0 To (PropFixedCols - 1)
                .Right = .Left + GetColWidth(iCol)
                If (VBFlexGridCells.Rows(iRow).RowInfo.State And RWIS_MERGE) = RWIS_MERGE Then
                    If iCol > 0 Then
                        Select Case PropMergeCells
                            Case FlexMergeCellsFree, FlexMergeCellsRestrictRows, FlexMergeCellsFixedOnly
                                If MergeCompareFunction(iRow, iCol, iRow, iCol - 1) = True Then
                                    VBFlexGridMergeDrawInfo.Row.ColOffset = VBFlexGridMergeDrawInfo.Row.ColOffset + 1
                                    VBFlexGridMergeDrawInfo.Row.Width = VBFlexGridMergeDrawInfo.Row.Width + GetColWidth(iCol - 1)
                                Else
                                    VBFlexGridMergeDrawInfo.Row.ColOffset = 0
                                    VBFlexGridMergeDrawInfo.Row.Width = 0
                                End If
                            Case FlexMergeCellsRestrictColumns, FlexMergeCellsRestrictAll
                                If MergeCompareFunction(iRow, iCol, iRow, iCol - 1) = True Then
                                    If iRow > VBFlexGridTopRow Then
                                        If MergeCompareFunction(iRow - 1, iCol, iRow - 1, iCol - 1) = True Then
                                            VBFlexGridMergeDrawInfo.Row.ColOffset = VBFlexGridMergeDrawInfo.Row.ColOffset + 1
                                            VBFlexGridMergeDrawInfo.Row.Width = VBFlexGridMergeDrawInfo.Row.Width + GetColWidth(iCol - 1)
                                        Else
                                            VBFlexGridMergeDrawInfo.Row.ColOffset = 0
                                            VBFlexGridMergeDrawInfo.Row.Width = 0
                                        End If
                                    Else
                                        VBFlexGridMergeDrawInfo.Row.ColOffset = VBFlexGridMergeDrawInfo.Row.ColOffset + 1
                                        VBFlexGridMergeDrawInfo.Row.Width = VBFlexGridMergeDrawInfo.Row.Width + GetColWidth(iCol - 1)
                                    End If
                                Else
                                    VBFlexGridMergeDrawInfo.Row.ColOffset = 0
                                    VBFlexGridMergeDrawInfo.Row.Width = 0
                                End If
                        End Select
                    Else
                        VBFlexGridMergeDrawInfo.Row.ColOffset = 0
                        VBFlexGridMergeDrawInfo.Row.Width = 0
                    End If
                End If
                If (VBFlexGridColsInfo(iCol).State And CLIS_MERGE) = CLIS_MERGE Then
                    If iRow > VBFlexGridTopRow Then
                        Select Case PropMergeCells
                            Case FlexMergeCellsFree, FlexMergeCellsRestrictColumns, FlexMergeCellsFixedOnly
                                If MergeCompareFunction(iRow, iCol, iRow - 1, iCol) = True Then
                                    VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset + 1
                                    VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height + GetRowHeight(iRow - 1)
                                Else
                                    VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = 0
                                    VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = 0
                                End If
                            Case FlexMergeCellsRestrictRows, FlexMergeCellsRestrictAll
                                If MergeCompareFunction(iRow, iCol, iRow - 1, iCol) = True Then
                                    If iCol > 0 Then
                                        If MergeCompareFunction(iRow, iCol - 1, iRow - 1, iCol - 1) = True Then
                                            VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset + 1
                                            VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height + GetRowHeight(iRow - 1)
                                        Else
                                            VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = 0
                                            VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = 0
                                        End If
                                    Else
                                        VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset + 1
                                        VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height + GetRowHeight(iRow - 1)
                                    End If
                                Else
                                    VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = 0
                                    VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = 0
                                End If
                        End Select
                    Else
                        VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = 0
                        VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = 0
                    End If
                End If
                .Left = .Left - VBFlexGridMergeDrawInfo.Row.Width
                .Top = .Top - VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height
                If .Bottom > .Top And .Right > .Left Then
                    VBFlexGridDrawInfo.GridLinePoints(0).X = .Left
                    VBFlexGridDrawInfo.GridLinePoints(0).Y = .Bottom - 1
                    VBFlexGridDrawInfo.GridLinePoints(1).X = .Right - 1
                    VBFlexGridDrawInfo.GridLinePoints(1).Y = .Bottom - 1
                    VBFlexGridDrawInfo.GridLinePoints(2).X = .Right - 1
                    VBFlexGridDrawInfo.GridLinePoints(2).Y = .Top - 1
                    VBFlexGridDrawInfo.GridLinePoints(3).X = .Right - 2
                    VBFlexGridDrawInfo.GridLinePoints(3).Y = .Top
                    VBFlexGridDrawInfo.GridLinePoints(4).X = .Left
                    VBFlexGridDrawInfo.GridLinePoints(4).Y = .Top
                    VBFlexGridDrawInfo.GridLinePoints(5).X = .Left
                    VBFlexGridDrawInfo.GridLinePoints(5).Y = .Bottom
                    Call DrawFixedCell(hDC, CellRect, iRow, iCol)
                End If
                .Left = .Left + VBFlexGridMergeDrawInfo.Row.Width
                .Top = .Top + VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height
                .Left = .Right
                If NoClip = False And .Right > VBFlexGridClientRect.Right Then Exit For
            Next iCol
            .Top = .Bottom
            If NoClip = False And .Bottom > VBFlexGridClientRect.Bottom Then Exit For
        Next iRow
    End If
    If VBFlexGridFontFixedHandle <> 0 Then
        If hFontOld <> 0 Then
            SelectObject hDC, hFontOld
            hFontOld = 0
        End If
        hFontOld = SelectObject(hDC, VBFlexGridFontHandle)
    End If
    SelectObject hDC, VBFlexGridBackColorBrush
    ReDim VBFlexGridMergeDrawInfo.Row.Cols(0 To (PropCols - 1)) As TMERGEDRAWCOLINFO
    .Top = FixedCY ' .Top = FixedCY + FrozenCY
    For iRow = PropFixedRows To (PropRows - 1) ' For iRow = VBFlexGridTopRow To (PropRows - 1)
        ' Below trick will save another loop with just the same code.
        ' For iRow = PropFixedRows To ((PropFixedRows + PropFrozenRows) - 1)
        If iRow >= (PropFixedRows + PropFrozenRows) Then
            If iRow < VBFlexGridTopRow Then iRow = VBFlexGridTopRow
        End If
        VBFlexGridMergeDrawInfo.Row.ColOffset = 0
        VBFlexGridMergeDrawInfo.Row.Width = 0
        .Bottom = .Top + GetRowHeight(iRow)
        If PropFrozenCols > 0 Then
            .Left = FixedCX
            For iCol = PropFixedCols To ((PropFixedCols + PropFrozenCols) - 1)
                .Right = .Left + GetColWidth(iCol)
                If (VBFlexGridCells.Rows(iRow).RowInfo.State And RWIS_MERGE) = RWIS_MERGE Then
                    If iCol > VBFlexGridLeftCol Then
                        Select Case PropMergeCells
                            Case FlexMergeCellsFree, FlexMergeCellsRestrictRows
                                If MergeCompareFunction(iRow, iCol, iRow, iCol - 1) = True Then
                                    VBFlexGridMergeDrawInfo.Row.ColOffset = VBFlexGridMergeDrawInfo.Row.ColOffset + 1
                                    VBFlexGridMergeDrawInfo.Row.Width = VBFlexGridMergeDrawInfo.Row.Width + GetColWidth(iCol - 1)
                                Else
                                    VBFlexGridMergeDrawInfo.Row.ColOffset = 0
                                    VBFlexGridMergeDrawInfo.Row.Width = 0
                                End If
                            Case FlexMergeCellsRestrictColumns, FlexMergeCellsRestrictAll
                                If MergeCompareFunction(iRow, iCol, iRow, iCol - 1) = True Then
                                    If iRow > VBFlexGridTopRow Then
                                        If MergeCompareFunction(iRow - 1, iCol, iRow - 1, iCol - 1) = True Then
                                            VBFlexGridMergeDrawInfo.Row.ColOffset = VBFlexGridMergeDrawInfo.Row.ColOffset + 1
                                            VBFlexGridMergeDrawInfo.Row.Width = VBFlexGridMergeDrawInfo.Row.Width + GetColWidth(iCol - 1)
                                        Else
                                            VBFlexGridMergeDrawInfo.Row.ColOffset = 0
                                            VBFlexGridMergeDrawInfo.Row.Width = 0
                                        End If
                                    Else
                                        VBFlexGridMergeDrawInfo.Row.ColOffset = VBFlexGridMergeDrawInfo.Row.ColOffset + 1
                                        VBFlexGridMergeDrawInfo.Row.Width = VBFlexGridMergeDrawInfo.Row.Width + GetColWidth(iCol - 1)
                                    End If
                                Else
                                    VBFlexGridMergeDrawInfo.Row.ColOffset = 0
                                    VBFlexGridMergeDrawInfo.Row.Width = 0
                                End If
                        End Select
                    Else
                        VBFlexGridMergeDrawInfo.Row.ColOffset = 0
                        VBFlexGridMergeDrawInfo.Row.Width = 0
                    End If
                End If
                If (VBFlexGridColsInfo(iCol).State And CLIS_MERGE) = CLIS_MERGE Then
                    If iRow > VBFlexGridTopRow Then
                        Select Case PropMergeCells
                            Case FlexMergeCellsFree, FlexMergeCellsRestrictColumns
                                If MergeCompareFunction(iRow, iCol, iRow - 1, iCol) = True Then
                                    VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset + 1
                                    VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height + GetRowHeight(iRow - 1)
                                Else
                                    VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = 0
                                    VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = 0
                                End If
                            Case FlexMergeCellsRestrictRows, FlexMergeCellsRestrictAll
                                If MergeCompareFunction(iRow, iCol, iRow - 1, iCol) = True Then
                                    If iCol > VBFlexGridLeftCol Then
                                        If MergeCompareFunction(iRow, iCol - 1, iRow - 1, iCol - 1) = True Then
                                            VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset + 1
                                            VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height + GetRowHeight(iRow - 1)
                                        Else
                                            VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = 0
                                            VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = 0
                                        End If
                                    Else
                                        VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset + 1
                                        VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height + GetRowHeight(iRow - 1)
                                    End If
                                Else
                                    VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = 0
                                    VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = 0
                                End If
                        End Select
                    Else
                        VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = 0
                        VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = 0
                    End If
                End If
                .Left = .Left - VBFlexGridMergeDrawInfo.Row.Width
                .Top = .Top - VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height
                If .Bottom > .Top And .Right > .Left Then
                    VBFlexGridDrawInfo.GridLinePoints(0).X = .Left
                    VBFlexGridDrawInfo.GridLinePoints(0).Y = .Bottom - 1
                    VBFlexGridDrawInfo.GridLinePoints(1).X = .Right - 1
                    VBFlexGridDrawInfo.GridLinePoints(1).Y = .Bottom - 1
                    VBFlexGridDrawInfo.GridLinePoints(2).X = .Right - 1
                    VBFlexGridDrawInfo.GridLinePoints(2).Y = .Top - 1
                    VBFlexGridDrawInfo.GridLinePoints(3).X = .Right - 2
                    VBFlexGridDrawInfo.GridLinePoints(3).Y = .Top
                    VBFlexGridDrawInfo.GridLinePoints(4).X = .Left
                    VBFlexGridDrawInfo.GridLinePoints(4).Y = .Top
                    VBFlexGridDrawInfo.GridLinePoints(5).X = .Left
                    VBFlexGridDrawInfo.GridLinePoints(5).Y = .Bottom
                    Call DrawCell(hDC, CellRect, iRow, iCol)
                End If
                .Left = .Left + VBFlexGridMergeDrawInfo.Row.Width
                .Top = .Top + VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height
                .Left = .Right
                If NoClip = False And .Right > VBFlexGridClientRect.Right Then Exit For
            Next iCol
        End If
        .Left = FixedCX + FrozenCX
        For iCol = VBFlexGridLeftCol To (PropCols - 1)
            .Right = .Left + GetColWidth(iCol)
            If (VBFlexGridCells.Rows(iRow).RowInfo.State And RWIS_MERGE) = RWIS_MERGE Then
                If iCol > VBFlexGridLeftCol Then
                    Select Case PropMergeCells
                        Case FlexMergeCellsFree, FlexMergeCellsRestrictRows
                            If MergeCompareFunction(iRow, iCol, iRow, iCol - 1) = True Then
                                VBFlexGridMergeDrawInfo.Row.ColOffset = VBFlexGridMergeDrawInfo.Row.ColOffset + 1
                                VBFlexGridMergeDrawInfo.Row.Width = VBFlexGridMergeDrawInfo.Row.Width + GetColWidth(iCol - 1)
                            Else
                                VBFlexGridMergeDrawInfo.Row.ColOffset = 0
                                VBFlexGridMergeDrawInfo.Row.Width = 0
                            End If
                        Case FlexMergeCellsRestrictColumns, FlexMergeCellsRestrictAll
                            If MergeCompareFunction(iRow, iCol, iRow, iCol - 1) = True Then
                                If iRow > VBFlexGridTopRow Then
                                    If MergeCompareFunction(iRow - 1, iCol, iRow - 1, iCol - 1) = True Then
                                        VBFlexGridMergeDrawInfo.Row.ColOffset = VBFlexGridMergeDrawInfo.Row.ColOffset + 1
                                        VBFlexGridMergeDrawInfo.Row.Width = VBFlexGridMergeDrawInfo.Row.Width + GetColWidth(iCol - 1)
                                    Else
                                        VBFlexGridMergeDrawInfo.Row.ColOffset = 0
                                        VBFlexGridMergeDrawInfo.Row.Width = 0
                                    End If
                                Else
                                    VBFlexGridMergeDrawInfo.Row.ColOffset = VBFlexGridMergeDrawInfo.Row.ColOffset + 1
                                    VBFlexGridMergeDrawInfo.Row.Width = VBFlexGridMergeDrawInfo.Row.Width + GetColWidth(iCol - 1)
                                End If
                            Else
                                VBFlexGridMergeDrawInfo.Row.ColOffset = 0
                                VBFlexGridMergeDrawInfo.Row.Width = 0
                            End If
                    End Select
                Else
                    VBFlexGridMergeDrawInfo.Row.ColOffset = 0
                    VBFlexGridMergeDrawInfo.Row.Width = 0
                End If
            End If
            If (VBFlexGridColsInfo(iCol).State And CLIS_MERGE) = CLIS_MERGE Then
                If iRow > VBFlexGridTopRow Then
                    Select Case PropMergeCells
                        Case FlexMergeCellsFree, FlexMergeCellsRestrictColumns
                            If MergeCompareFunction(iRow, iCol, iRow - 1, iCol) = True Then
                                VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset + 1
                                VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height + GetRowHeight(iRow - 1)
                            Else
                                VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = 0
                                VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = 0
                            End If
                        Case FlexMergeCellsRestrictRows, FlexMergeCellsRestrictAll
                            If MergeCompareFunction(iRow, iCol, iRow - 1, iCol) = True Then
                                If iCol > VBFlexGridLeftCol Then
                                    If MergeCompareFunction(iRow, iCol - 1, iRow - 1, iCol - 1) = True Then
                                        VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset + 1
                                        VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height + GetRowHeight(iRow - 1)
                                    Else
                                        VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = 0
                                        VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = 0
                                    End If
                                Else
                                    VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset + 1
                                    VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height + GetRowHeight(iRow - 1)
                                End If
                            Else
                                VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = 0
                                VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = 0
                            End If
                    End Select
                Else
                    VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = 0
                    VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = 0
                End If
            End If
            .Left = .Left - VBFlexGridMergeDrawInfo.Row.Width
            .Top = .Top - VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height
            If .Bottom > .Top And .Right > .Left Then
                VBFlexGridDrawInfo.GridLinePoints(0).X = .Left
                VBFlexGridDrawInfo.GridLinePoints(0).Y = .Bottom - 1
                VBFlexGridDrawInfo.GridLinePoints(1).X = .Right - 1
                VBFlexGridDrawInfo.GridLinePoints(1).Y = .Bottom - 1
                VBFlexGridDrawInfo.GridLinePoints(2).X = .Right - 1
                VBFlexGridDrawInfo.GridLinePoints(2).Y = .Top - 1
                VBFlexGridDrawInfo.GridLinePoints(3).X = .Right - 2
                VBFlexGridDrawInfo.GridLinePoints(3).Y = .Top
                VBFlexGridDrawInfo.GridLinePoints(4).X = .Left
                VBFlexGridDrawInfo.GridLinePoints(4).Y = .Top
                VBFlexGridDrawInfo.GridLinePoints(5).X = .Left
                VBFlexGridDrawInfo.GridLinePoints(5).Y = .Bottom
                Call DrawCell(hDC, CellRect, iRow, iCol)
            End If
            .Left = .Left + VBFlexGridMergeDrawInfo.Row.Width
            .Top = .Top + VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height
            .Left = .Right
            If NoClip = False And .Right > VBFlexGridClientRect.Right Then Exit For
        Next iCol
        If .Right > GridRect.Right Then GridRect.Right = .Right
        .Top = .Bottom
        If NoClip = False And .Bottom > VBFlexGridClientRect.Bottom Then Exit For
    Next iRow
    If .Bottom > GridRect.Bottom Then GridRect.Bottom = .Bottom
    If hFontOld <> 0 Then
        SelectObject hDC, hFontOld
        hFontOld = 0
    End If
    If Brush <> 0 Then
        SelectObject hDC, Brush
        Brush = 0
    End If
    Erase VBFlexGridMergeDrawInfo.Row.Cols()
    VBFlexGridMergeDrawInfo.Row.ColOffset = 0
    VBFlexGridMergeDrawInfo.Row.Width = 0
End If
End With
SetBkMode hDC, OldBkMode
With GridRect
Dim hPenOld As Long
hPenOld = SelectObject(hDC, VBFlexGridGridLineFixedPen)
If PropFrozenCols > 0 Then
    VBFlexGridDrawInfo.GridLinePoints(0).X = (FixedCX + FrozenCX) - 1
    VBFlexGridDrawInfo.GridLinePoints(0).Y = .Top
    VBFlexGridDrawInfo.GridLinePoints(1).X = (FixedCX + FrozenCX) - 1
    VBFlexGridDrawInfo.GridLinePoints(1).Y = .Bottom - 1
    Polyline hDC, VBFlexGridDrawInfo.GridLinePoints(0), 2
End If
If PropFrozenRows > 0 Then
    VBFlexGridDrawInfo.GridLinePoints(0).X = .Left
    VBFlexGridDrawInfo.GridLinePoints(0).Y = (FixedCY + FrozenCY) - 1
    VBFlexGridDrawInfo.GridLinePoints(1).X = .Right - 1
    VBFlexGridDrawInfo.GridLinePoints(1).Y = (FixedCY + FrozenCY) - 1
    Polyline hDC, VBFlexGridDrawInfo.GridLinePoints(0), 2
End If
VBFlexGridDrawInfo.GridLinePoints(0).X = .Left
VBFlexGridDrawInfo.GridLinePoints(0).Y = .Bottom - 1
VBFlexGridDrawInfo.GridLinePoints(1).X = .Right - 1
VBFlexGridDrawInfo.GridLinePoints(1).Y = .Bottom - 1
VBFlexGridDrawInfo.GridLinePoints(2).X = .Right - 1
VBFlexGridDrawInfo.GridLinePoints(2).Y = .Top - 1
Polyline hDC, VBFlexGridDrawInfo.GridLinePoints(0), 3
If hPenOld <> 0 Then
    SelectObject hDC, hPenOld
    hPenOld = 0
End If
If hRgn <> -1 Then hRgn = CreateRectRgn(.Left, .Top, .Right, .Bottom)
End With
End Sub

Private Sub DrawFixedCell(ByRef hDC As Long, ByRef CellRect As RECT, ByVal iRow As Long, ByVal iCol As Long)
Dim ItemState As Long
If PropMergeCells <> FlexMergeCellsNever Then
    If (VBFlexGridRow >= (iRow - VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset) And VBFlexGridRow <= iRow) And (VBFlexGridCol >= (iCol - VBFlexGridMergeDrawInfo.Row.ColOffset) And VBFlexGridCol <= iCol) Then
        iRow = VBFlexGridRow
        iCol = VBFlexGridCol
    Else
        iRow = iRow - VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset
        iCol = iCol - VBFlexGridMergeDrawInfo.Row.ColOffset
    End If
End If
With VBFlexGridDrawInfo.SelRange
Select Case PropSelectionMode
    Case FlexSelectionModeFree, FlexSelectionModeByRow, FlexSelectionModeByColumn
        Select Case PropHighLight
            Case FlexHighLightAlways
                If (iCol >= .LeftCol And iCol <= .RightCol) And (iRow >= .TopRow And iRow <= .BottomRow) Then ItemState = ItemState Or ODS_SELECTED
            Case FlexHighLightWithFocus
                If VBFlexGridFocused = True Then
                    If (iCol >= .LeftCol And iCol <= .RightCol) And (iRow >= .TopRow And iRow <= .BottomRow) Then ItemState = ItemState Or ODS_SELECTED
                End If
        End Select
    Case FlexSelectionModeFreeByRow
        Select Case PropHighLight
            Case FlexHighLightAlways
                If (iCol >= .LeftCol And iCol <= .RightCol) And (iRow >= .TopRow And iRow <= .BottomRow) Then ItemState = ItemState Or ODS_SELECTED
            Case FlexHighLightWithFocus
                If VBFlexGridFocused = True Then
                    If (iCol >= .LeftCol And iCol <= .RightCol) And (iRow >= .TopRow And iRow <= .BottomRow) Then ItemState = ItemState Or ODS_SELECTED
                End If
        End Select
    Case FlexSelectionModeFreeByColumn
        Select Case PropHighLight
            Case FlexHighLightAlways
                If (iCol >= .LeftCol And iCol <= .RightCol) And (iRow >= .TopRow And iRow <= .BottomRow) Then ItemState = ItemState Or ODS_SELECTED
            Case FlexHighLightWithFocus
                If VBFlexGridFocused = True Then
                    If (iCol >= .LeftCol And iCol <= .RightCol) And (iRow >= .TopRow And iRow <= .BottomRow) Then ItemState = ItemState Or ODS_SELECTED
                End If
        End Select
End Select
End With
If PropFocusRect <> FlexFocusRectNone Then
    If (iRow = VBFlexGridRow And iCol = VBFlexGridCol) Then ItemState = ItemState Or ODS_FOCUS
End If
If VBFlexGridFocused = False Then ItemState = ItemState Or ODS_NOFOCUSRECT
Dim Text As String, TextRect As RECT
Call GetCellText(iRow, iCol, Text)
With TextRect
.Left = CellRect.Left + VBFlexGridDrawInfo.CellTextWidthPadding
.Top = CellRect.Top + VBFlexGridDrawInfo.CellTextHeightPadding
.Right = CellRect.Right - VBFlexGridDrawInfo.CellTextWidthPadding
.Bottom = CellRect.Bottom - VBFlexGridDrawInfo.CellTextHeightPadding
End With
With VBFlexGridCells.Rows(iRow).Cols(iCol)
Dim hFontTemp As Long, hFontOld As Long
If Not .FontName = vbNullString Then
    Dim TempFont As StdFont
    Set TempFont = New StdFont
    TempFont.Name = .FontName
    TempFont.Size = .FontSize
    TempFont.Bold = .FontBold
    TempFont.Italic = .FontItalic
    TempFont.Strikethrough = .FontStrikeThrough
    TempFont.Underline = .FontUnderline
    TempFont.Charset = .FontCharset
    hFontTemp = CreateGDIFontFromOLEFont(TempFont)
    hFontOld = SelectObject(hDC, hFontTemp)
    Set TempFont = Nothing
End If
Dim Brush As Long
If Not (ItemState And ODS_SELECTED) = ODS_SELECTED Or (ItemState And ODS_FOCUS) = ODS_FOCUS Then
    If .BackColor = -1 Then
        PatBlt hDC, CellRect.Left, CellRect.Top, CellRect.Right - CellRect.Left, CellRect.Bottom - CellRect.Top, vbPatCopy
    Else
        Brush = SetBkColor(hDC, WinColor(.BackColor))
        ExtTextOut hDC, 0, 0, ETO_OPAQUE, CellRect, 0, 0, 0
        SetBkColor hDC, Brush
    End If
Else
    Brush = SelectObject(hDC, VBFlexGridBackColorSelBrush)
    PatBlt hDC, CellRect.Left, CellRect.Top, CellRect.Right - CellRect.Left, CellRect.Bottom - CellRect.Top, vbPatCopy
    SelectObject hDC, Brush
End If
If Not .Picture Is Nothing Then
    If .Picture.Handle <> 0 Then
        Dim PictureWidth As Long, PictureHeight As Long
        Dim PictureLeft As Long, PictureTop As Long, PictureOffsetX As Long, PictureOffsetY As Long
        If .PictureAlignment <> FlexPictureAlignmentStretch Then
            PictureWidth = CHimetricToPixel_X(.Picture.Width)
            PictureHeight = CHimetricToPixel_Y(.Picture.Height)
        Else
            PictureWidth = (CellRect.Right - CellRect.Left)
            PictureHeight = (CellRect.Bottom - CellRect.Top)
        End If
        PictureLeft = CellRect.Left
        PictureTop = CellRect.Top
        Select Case .PictureAlignment
            Case FlexPictureAlignmentLeftCenter, FlexPictureAlignmentLeftCenterNoOverlap
                PictureOffsetY = (((CellRect.Bottom - CellRect.Top) - PictureHeight) / 2)
            Case FlexPictureAlignmentLeftBottom, FlexPictureAlignmentLeftBottomNoOverlap
                PictureOffsetY = ((CellRect.Bottom - CellRect.Top) - PictureHeight)
            Case FlexPictureAlignmentCenterTop
                PictureOffsetX = (((CellRect.Right - CellRect.Left) - PictureWidth) / 2)
            Case FlexPictureAlignmentCenterCenter
                PictureOffsetX = (((CellRect.Right - CellRect.Left) - PictureWidth) / 2)
                PictureOffsetY = (((CellRect.Bottom - CellRect.Top) - PictureHeight) / 2)
            Case FlexPictureAlignmentCenterBottom
                PictureOffsetX = (((CellRect.Right - CellRect.Left) - PictureWidth) / 2)
                PictureOffsetY = ((CellRect.Bottom - CellRect.Top) - PictureHeight)
            Case FlexPictureAlignmentRightTop, FlexPictureAlignmentRightTopNoOverlap
                PictureOffsetX = ((CellRect.Right - CellRect.Left) - PictureWidth)
            Case FlexPictureAlignmentRightCenter, FlexPictureAlignmentRightCenterNoOverlap
                PictureOffsetX = ((CellRect.Right - CellRect.Left) - PictureWidth)
                PictureOffsetY = (((CellRect.Bottom - CellRect.Top) - PictureHeight) / 2)
            Case FlexPictureAlignmentRightBottom, FlexPictureAlignmentRightBottomNoOverlap
                PictureOffsetX = ((CellRect.Right - CellRect.Left) - PictureWidth)
                PictureOffsetY = ((CellRect.Bottom - CellRect.Top) - PictureHeight)
        End Select
        If PictureOffsetX > 0 Then PictureLeft = PictureLeft + PictureOffsetX
        If PictureOffsetY > 0 Then PictureTop = PictureTop + PictureOffsetY
        If .PictureAlignment <> FlexPictureAlignmentTile Then
            Call RenderPicture(.Picture, hDC, PictureLeft, PictureTop, PictureWidth, PictureHeight, .PictureRenderFlag)
        Else
            Do
                Do
                    Call RenderPicture(.Picture, hDC, PictureLeft, PictureTop, PictureWidth, PictureHeight, .PictureRenderFlag)
                    PictureTop = PictureTop + PictureHeight
                Loop While PictureTop < CellRect.Bottom
                PictureLeft = PictureLeft + PictureWidth
                PictureTop = CellRect.Top
            Loop While PictureLeft < CellRect.Right
        End If
        Select Case .PictureAlignment
            Case FlexPictureAlignmentLeftTopNoOverlap, FlexPictureAlignmentLeftCenterNoOverlap, FlexPictureAlignmentLeftBottomNoOverlap
                TextRect.Left = TextRect.Left + PictureWidth
            Case FlexPictureAlignmentRightTopNoOverlap, FlexPictureAlignmentRightCenterNoOverlap, FlexPictureAlignmentRightBottomNoOverlap
                TextRect.Right = TextRect.Right - PictureWidth
        End Select
    End If
End If
Dim OldTextColor As Long
If Not (ItemState And ODS_SELECTED) = ODS_SELECTED Or (ItemState And ODS_FOCUS) = ODS_FOCUS Then
    If Not Text = vbNullString Then
        If .ForeColor = -1 Then
            OldTextColor = SetTextColor(hDC, WinColor(PropForeColorFixed))
        Else
            OldTextColor = SetTextColor(hDC, WinColor(.ForeColor))
        End If
    Else
        OldTextColor = SetTextColor(hDC, WinColor(vbButtonText))
    End If
Else
    OldTextColor = SetTextColor(hDC, WinColor(PropForeColorSel))
End If
Dim hPenOld As Long, P As POINTAPI
Select Case PropGridLinesFixed
    Case FlexGridLineFlat, FlexGridLineDashes, FlexGridLineDots
        hPenOld = SelectObject(hDC, VBFlexGridGridLineFixedPen)
        Polyline hDC, VBFlexGridDrawInfo.GridLinePoints(0), 3
    Case FlexGridLineInset, FlexGridLineRaised
        If PropGridLinesFixed = FlexGridLineInset Then
            hPenOld = SelectObject(hDC, VBFlexGridGridLineBlackPen)
        ElseIf PropGridLinesFixed = FlexGridLineRaised Then
            hPenOld = SelectObject(hDC, VBFlexGridGridLineWhitePen)
        End If
        Polyline hDC, VBFlexGridDrawInfo.GridLinePoints(0), 3
        If PropGridLinesFixed = FlexGridLineInset Then
            SelectObject hDC, VBFlexGridGridLineWhitePen
        ElseIf PropGridLinesFixed = FlexGridLineRaised Then
            SelectObject hDC, VBFlexGridGridLineBlackPen
        End If
        Polyline hDC, VBFlexGridDrawInfo.GridLinePoints(3), 3
End Select
If hPenOld <> 0 Then
    SelectObject hDC, hPenOld
    hPenOld = 0
End If
If (ItemState And ODS_FOCUS) = ODS_FOCUS And Not (ItemState And ODS_NOFOCUSRECT) = ODS_NOFOCUSRECT Then
    Dim FocusRect As RECT
    With FocusRect
    .Left = CellRect.Left
    .Top = CellRect.Top
    .Right = CellRect.Right - 1
    .Bottom = CellRect.Bottom - 1
    If (.Right - VBFlexGridFocusBorder.CX) <= .Left Then .Right = CellRect.Right + 1
    If (.Bottom - VBFlexGridFocusBorder.CY) <= .Top Then .Bottom = CellRect.Bottom + 1
    DrawFocusRect hDC, FocusRect
    If PropFocusRect = FlexFocusRectHeavy Then
        If (.Right - VBFlexGridFocusBorder.CX) > (.Left + VBFlexGridFocusBorder.CX) And (.Bottom - VBFlexGridFocusBorder.CY) > (.Top + VBFlexGridFocusBorder.CY) Then
            .Left = .Left + VBFlexGridFocusBorder.CX
            .Right = .Right - VBFlexGridFocusBorder.CX
            .Top = .Top + VBFlexGridFocusBorder.CY
            .Bottom = .Bottom - VBFlexGridFocusBorder.CY
            DrawFocusRect hDC, FocusRect
        End If
    End If
    End With
End If
If Not Text = vbNullString Then
    Dim TextStyle As FlexTextStyleConstants, Alignment As FlexAlignmentConstants, DrawFlags As Long
    If .TextStyle = -1 Then
        TextStyle = PropTextStyleFixed
    Else
        TextStyle = .TextStyle
    End If
    If .Alignment = -1 Then
        If VBFlexGridColsInfo(iCol).FixedAlignment = -1 Then
            Alignment = VBFlexGridColsInfo(iCol).Alignment
        Else
            Alignment = VBFlexGridColsInfo(iCol).FixedAlignment
        End If
    Else
        Alignment = .Alignment
    End If
    DrawFlags = DT_NOPREFIX
    If VBFlexGridRTLReading = True Then DrawFlags = DrawFlags Or DT_RTLREADING
    Select Case Alignment
        Case FlexAlignmentLeftTop, FlexAlignmentLeftCenter, FlexAlignmentLeftBottom
            DrawFlags = DrawFlags Or DT_LEFT
        Case FlexAlignmentCenterTop, FlexAlignmentCenterCenter, FlexAlignmentCenterBottom
            DrawFlags = DrawFlags Or DT_CENTER
        Case FlexAlignmentRightTop, FlexAlignmentRightCenter, FlexAlignmentRightBottom
            DrawFlags = DrawFlags Or DT_RIGHT
        Case FlexAlignmentGeneral
            If Not IsNumeric(Text) Then
                DrawFlags = DrawFlags Or DT_LEFT
            Else
                DrawFlags = DrawFlags Or DT_RIGHT
            End If
    End Select
    If PropWordWrap = True Then
        DrawFlags = DrawFlags Or DT_WORDBREAK
    ElseIf PropSingleLine = True Then
        DrawFlags = DrawFlags Or DT_SINGLELINE
    End If
    Select Case PropEllipsisFormatFixed
        Case FlexEllipsisFormatEnd
            DrawFlags = DrawFlags Or DT_END_ELLIPSIS
        Case FlexEllipsisFormatPath
            DrawFlags = DrawFlags Or DT_PATH_ELLIPSIS
        Case FlexEllipsisFormatWord
            DrawFlags = DrawFlags Or DT_WORD_ELLIPSIS
    End Select
    If Not VBFlexGridColsInfo(iCol).Format = vbNullString Then Text = Format$(Text, VBFlexGridColsInfo(iCol).Format, vbUseSystemDayOfWeek, vbUseSystem)
    If Not (DrawFlags And DT_SINGLELINE) = DT_SINGLELINE Then
        Dim CalcRect As RECT, Height As Long, Result As Long
        Select Case Alignment
            Case FlexAlignmentLeftCenter, FlexAlignmentCenterCenter, FlexAlignmentRightCenter, FlexAlignmentGeneral
                LSet CalcRect = TextRect
                Height = DrawText(hDC, StrPtr(Text), -1, CalcRect, DrawFlags Or DT_CALCRECT)
                Result = (((TextRect.Bottom - TextRect.Top) - Height) / 2)
            Case FlexAlignmentLeftBottom, FlexAlignmentCenterBottom, FlexAlignmentRightBottom
                LSet CalcRect = TextRect
                Height = DrawText(hDC, StrPtr(Text), -1, CalcRect, DrawFlags Or DT_CALCRECT)
                Result = ((TextRect.Bottom - TextRect.Top) - Height)
        End Select
        If Result > 0 Then TextRect.Top = TextRect.Top + Result
    Else
        Select Case Alignment
            Case FlexAlignmentLeftCenter, FlexAlignmentCenterCenter, FlexAlignmentRightCenter, FlexAlignmentGeneral
                DrawFlags = DrawFlags Or DT_VCENTER
            Case FlexAlignmentLeftBottom, FlexAlignmentCenterBottom, FlexAlignmentRightBottom
                DrawFlags = DrawFlags Or DT_BOTTOM
        End Select
    End If
    Dim Offset As Long, TempTextColor As Long
    Select Case TextStyle
        Case FlexTextStyleRaised
            TempTextColor = SetTextColor(hDC, &H808080)
            Offset = 1
        Case FlexTextStyleRaisedLight
            TempTextColor = SetTextColor(hDC, vbWhite)
            Offset = 1
        Case FlexTextStyleInset
            TempTextColor = SetTextColor(hDC, &H808080)
            Offset = -1
        Case FlexTextStyleInsetLight
            TempTextColor = SetTextColor(hDC, vbWhite)
            Offset = -1
    End Select
    If Offset <> 0 Then
        With TextRect
        .Top = .Top + Offset
        .Left = .Left + Offset
        .Bottom = .Bottom + Offset
        .Right = .Right + Offset
        End With
        DrawText hDC, StrPtr(Text), -1, TextRect, DrawFlags
        SetTextColor hDC, TempTextColor
        With TextRect
        .Top = .Top - Offset
        .Left = .Left - Offset
        .Bottom = .Bottom - Offset
        .Right = .Right - Offset
        End With
    End If
    DrawText hDC, StrPtr(Text), -1, TextRect, DrawFlags
End If
SetTextColor hDC, OldTextColor
If hFontOld <> 0 Then SelectObject hDC, hFontOld
If hFontTemp <> 0 Then DeleteObject hFontTemp
End With
End Sub

Private Sub DrawCell(ByRef hDC As Long, ByRef CellRect As RECT, ByVal iRow As Long, ByVal iCol As Long)
Dim ItemState As Long
If PropMergeCells <> FlexMergeCellsNever Then
    If (VBFlexGridRow >= (iRow - VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset) And VBFlexGridRow <= iRow) And (VBFlexGridCol >= (iCol - VBFlexGridMergeDrawInfo.Row.ColOffset) And VBFlexGridCol <= iCol) Then
        iRow = VBFlexGridRow
        iCol = VBFlexGridCol
    Else
        iRow = iRow - VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset
        iCol = iCol - VBFlexGridMergeDrawInfo.Row.ColOffset
    End If
End If
With VBFlexGridDrawInfo.SelRange
Select Case PropSelectionMode
    Case FlexSelectionModeFree, FlexSelectionModeByRow, FlexSelectionModeByColumn
        Select Case PropHighLight
            Case FlexHighLightAlways
                If (iCol >= .LeftCol And iCol <= .RightCol) And (iRow >= .TopRow And iRow <= .BottomRow) Then ItemState = ItemState Or ODS_SELECTED
            Case FlexHighLightWithFocus
                If VBFlexGridFocused = True Then
                    If (iCol >= .LeftCol And iCol <= .RightCol) And (iRow >= .TopRow And iRow <= .BottomRow) Then ItemState = ItemState Or ODS_SELECTED
                End If
        End Select
    Case FlexSelectionModeFreeByRow
        Select Case PropHighLight
            Case FlexHighLightAlways
                If (iRow >= .TopRow And iRow <= .BottomRow) Then ItemState = ItemState Or ODS_SELECTED
            Case FlexHighLightWithFocus
                If VBFlexGridFocused = True Then
                    If (iRow >= .TopRow And iRow <= .BottomRow) Then ItemState = ItemState Or ODS_SELECTED
                End If
        End Select
    Case FlexSelectionModeFreeByColumn
        Select Case PropHighLight
            Case FlexHighLightAlways
                If (iCol >= .LeftCol And iCol <= .RightCol) Then ItemState = ItemState Or ODS_SELECTED
            Case FlexHighLightWithFocus
                If VBFlexGridFocused = True Then
                    If (iCol >= .LeftCol And iCol <= .RightCol) Then ItemState = ItemState Or ODS_SELECTED
                End If
        End Select
End Select
End With
If PropFocusRect <> FlexFocusRectNone Then
    If (iRow = VBFlexGridRow And iCol = VBFlexGridCol) Then ItemState = ItemState Or ODS_FOCUS
End If
If VBFlexGridFocused = False Then ItemState = ItemState Or ODS_NOFOCUSRECT
Dim Text As String, TextRect As RECT
Call GetCellText(iRow, iCol, Text)
With TextRect
.Left = CellRect.Left + VBFlexGridDrawInfo.CellTextWidthPadding
.Top = CellRect.Top + VBFlexGridDrawInfo.CellTextHeightPadding
.Right = CellRect.Right - VBFlexGridDrawInfo.CellTextWidthPadding
.Bottom = CellRect.Bottom - VBFlexGridDrawInfo.CellTextHeightPadding
End With
With VBFlexGridCells.Rows(iRow).Cols(iCol)
Dim hFontTemp As Long, hFontOld As Long
If Not .FontName = vbNullString Then
    Dim TempFont As StdFont
    Set TempFont = New StdFont
    TempFont.Name = .FontName
    TempFont.Size = .FontSize
    TempFont.Bold = .FontBold
    TempFont.Italic = .FontItalic
    TempFont.Strikethrough = .FontStrikeThrough
    TempFont.Underline = .FontUnderline
    TempFont.Charset = .FontCharset
    hFontTemp = CreateGDIFontFromOLEFont(TempFont)
    hFontOld = SelectObject(hDC, hFontTemp)
    Set TempFont = Nothing
End If
Dim Brush As Long
If Not (ItemState And ODS_SELECTED) = ODS_SELECTED Or (ItemState And ODS_FOCUS) = ODS_FOCUS Then
    If .BackColor = -1 Then
        If PropBackColor = PropBackColorAlt Then
            PatBlt hDC, CellRect.Left, CellRect.Top, CellRect.Right - CellRect.Left, CellRect.Bottom - CellRect.Top, vbPatCopy
        Else
            If (iRow - PropFixedRows) Mod 2 = 0 Then
                PatBlt hDC, CellRect.Left, CellRect.Top, CellRect.Right - CellRect.Left, CellRect.Bottom - CellRect.Top, vbPatCopy
            Else
                Brush = SelectObject(hDC, VBFlexGridBackColorAltBrush)
                PatBlt hDC, CellRect.Left, CellRect.Top, CellRect.Right - CellRect.Left, CellRect.Bottom - CellRect.Top, vbPatCopy
                SelectObject hDC, Brush
            End If
        End If
    Else
        Brush = SetBkColor(hDC, WinColor(.BackColor))
        ExtTextOut hDC, 0, 0, ETO_OPAQUE, CellRect, 0, 0, 0
        SetBkColor hDC, Brush
    End If
Else
    Brush = SelectObject(hDC, VBFlexGridBackColorSelBrush)
    PatBlt hDC, CellRect.Left, CellRect.Top, CellRect.Right - CellRect.Left, CellRect.Bottom - CellRect.Top, vbPatCopy
    SelectObject hDC, Brush
End If
If Not .Picture Is Nothing Then
    If .Picture.Handle <> 0 Then
        Dim PictureWidth As Long, PictureHeight As Long
        Dim PictureLeft As Long, PictureTop As Long, PictureOffsetX As Long, PictureOffsetY As Long
        If .PictureAlignment <> FlexPictureAlignmentStretch Then
            PictureWidth = CHimetricToPixel_X(.Picture.Width)
            PictureHeight = CHimetricToPixel_Y(.Picture.Height)
        Else
            PictureWidth = (CellRect.Right - CellRect.Left)
            PictureHeight = (CellRect.Bottom - CellRect.Top)
        End If
        PictureLeft = CellRect.Left
        PictureTop = CellRect.Top
        Select Case .PictureAlignment
            Case FlexPictureAlignmentLeftCenter, FlexPictureAlignmentLeftCenterNoOverlap
                PictureOffsetY = (((CellRect.Bottom - CellRect.Top) - PictureHeight) / 2)
            Case FlexPictureAlignmentLeftBottom, FlexPictureAlignmentLeftBottomNoOverlap
                PictureOffsetY = ((CellRect.Bottom - CellRect.Top) - PictureHeight)
            Case FlexPictureAlignmentCenterTop
                PictureOffsetX = (((CellRect.Right - CellRect.Left) - PictureWidth) / 2)
            Case FlexPictureAlignmentCenterCenter
                PictureOffsetX = (((CellRect.Right - CellRect.Left) - PictureWidth) / 2)
                PictureOffsetY = (((CellRect.Bottom - CellRect.Top) - PictureHeight) / 2)
            Case FlexPictureAlignmentCenterBottom
                PictureOffsetX = (((CellRect.Right - CellRect.Left) - PictureWidth) / 2)
                PictureOffsetY = ((CellRect.Bottom - CellRect.Top) - PictureHeight)
            Case FlexPictureAlignmentRightTop, FlexPictureAlignmentRightTopNoOverlap
                PictureOffsetX = ((CellRect.Right - CellRect.Left) - PictureWidth)
            Case FlexPictureAlignmentRightCenter, FlexPictureAlignmentRightCenterNoOverlap
                PictureOffsetX = ((CellRect.Right - CellRect.Left) - PictureWidth)
                PictureOffsetY = (((CellRect.Bottom - CellRect.Top) - PictureHeight) / 2)
            Case FlexPictureAlignmentRightBottom, FlexPictureAlignmentRightBottomNoOverlap
                PictureOffsetX = ((CellRect.Right - CellRect.Left) - PictureWidth)
                PictureOffsetY = ((CellRect.Bottom - CellRect.Top) - PictureHeight)
        End Select
        If PictureOffsetX > 0 Then PictureLeft = PictureLeft + PictureOffsetX
        If PictureOffsetY > 0 Then PictureTop = PictureTop + PictureOffsetY
        If .PictureAlignment <> FlexPictureAlignmentTile Then
            Call RenderPicture(.Picture, hDC, PictureLeft, PictureTop, PictureWidth, PictureHeight, .PictureRenderFlag)
        Else
            Do
                Do
                    Call RenderPicture(.Picture, hDC, PictureLeft, PictureTop, PictureWidth, PictureHeight, .PictureRenderFlag)
                    PictureTop = PictureTop + PictureHeight
                Loop While PictureTop < CellRect.Bottom
                PictureLeft = PictureLeft + PictureWidth
                PictureTop = CellRect.Top
            Loop While PictureLeft < CellRect.Right
        End If
        Select Case .PictureAlignment
            Case FlexPictureAlignmentLeftTopNoOverlap, FlexPictureAlignmentLeftCenterNoOverlap, FlexPictureAlignmentLeftBottomNoOverlap
                TextRect.Left = TextRect.Left + PictureWidth
            Case FlexPictureAlignmentRightTopNoOverlap, FlexPictureAlignmentRightCenterNoOverlap, FlexPictureAlignmentRightBottomNoOverlap
                TextRect.Right = TextRect.Right - PictureWidth
        End Select
    End If
End If
Dim OldTextColor As Long
If Not (ItemState And ODS_SELECTED) = ODS_SELECTED Or (ItemState And ODS_FOCUS) = ODS_FOCUS Then
    If Not Text = vbNullString Then
        If .ForeColor = -1 Then
            OldTextColor = SetTextColor(hDC, WinColor(PropForeColor))
        Else
            OldTextColor = SetTextColor(hDC, WinColor(.ForeColor))
        End If
    Else
        OldTextColor = SetTextColor(hDC, WinColor(vbWindowText))
    End If
Else
    OldTextColor = SetTextColor(hDC, WinColor(PropForeColorSel))
End If
Dim hPenOld As Long, P As POINTAPI
Select Case PropGridLines
    Case FlexGridLineFlat, FlexGridLineDashes, FlexGridLineDots
        hPenOld = SelectObject(hDC, VBFlexGridGridLinePen)
        Polyline hDC, VBFlexGridDrawInfo.GridLinePoints(0), 3
    Case FlexGridLineInset, FlexGridLineRaised
        If PropGridLines = FlexGridLineInset Then
            hPenOld = SelectObject(hDC, VBFlexGridGridLineBlackPen)
        ElseIf PropGridLines = FlexGridLineRaised Then
            hPenOld = SelectObject(hDC, VBFlexGridGridLineWhitePen)
        End If
        Polyline hDC, VBFlexGridDrawInfo.GridLinePoints(0), 3
        If PropGridLines = FlexGridLineInset Then
            SelectObject hDC, VBFlexGridGridLineWhitePen
        ElseIf PropGridLines = FlexGridLineRaised Then
            SelectObject hDC, VBFlexGridGridLineBlackPen
        End If
        Polyline hDC, VBFlexGridDrawInfo.GridLinePoints(3), 3
End Select
If hPenOld <> 0 Then
    SelectObject hDC, hPenOld
    hPenOld = 0
End If
If (ItemState And ODS_FOCUS) = ODS_FOCUS And Not (ItemState And ODS_NOFOCUSRECT) = ODS_NOFOCUSRECT Then
    Dim FocusRect As RECT
    With FocusRect
    .Left = CellRect.Left
    .Top = CellRect.Top
    .Right = CellRect.Right - 1
    .Bottom = CellRect.Bottom - 1
    If (.Right - VBFlexGridFocusBorder.CX) <= .Left Then .Right = CellRect.Right + 1
    If (.Bottom - VBFlexGridFocusBorder.CY) <= .Top Then .Bottom = CellRect.Bottom + 1
    DrawFocusRect hDC, FocusRect
    If PropFocusRect = FlexFocusRectHeavy Then
        If (.Right - VBFlexGridFocusBorder.CX) > (.Left + VBFlexGridFocusBorder.CX) And (.Bottom - VBFlexGridFocusBorder.CY) > (.Top + VBFlexGridFocusBorder.CY) Then
            .Left = .Left + VBFlexGridFocusBorder.CX
            .Right = .Right - VBFlexGridFocusBorder.CX
            .Top = .Top + VBFlexGridFocusBorder.CY
            .Bottom = .Bottom - VBFlexGridFocusBorder.CY
            DrawFocusRect hDC, FocusRect
        End If
    End If
    End With
End If
If Not Text = vbNullString Then
    Dim TextStyle As FlexTextStyleConstants, Alignment As FlexAlignmentConstants, DrawFlags As Long
    If .TextStyle = -1 Then
        TextStyle = PropTextStyle
    Else
        TextStyle = .TextStyle
    End If
    If .Alignment = -1 Then
        Alignment = VBFlexGridColsInfo(iCol).Alignment
    Else
        Alignment = .Alignment
    End If
    DrawFlags = DT_NOPREFIX
    If VBFlexGridRTLReading = True Then DrawFlags = DrawFlags Or DT_RTLREADING
    Select Case Alignment
        Case FlexAlignmentLeftTop, FlexAlignmentLeftCenter, FlexAlignmentLeftBottom
            DrawFlags = DrawFlags Or DT_LEFT
        Case FlexAlignmentCenterTop, FlexAlignmentCenterCenter, FlexAlignmentCenterBottom
            DrawFlags = DrawFlags Or DT_CENTER
        Case FlexAlignmentRightTop, FlexAlignmentRightCenter, FlexAlignmentRightBottom
            DrawFlags = DrawFlags Or DT_RIGHT
        Case FlexAlignmentGeneral
            If Not IsNumeric(Text) Then
                DrawFlags = DrawFlags Or DT_LEFT
            Else
                DrawFlags = DrawFlags Or DT_RIGHT
            End If
    End Select
    If PropWordWrap = True Then
        DrawFlags = DrawFlags Or DT_WORDBREAK
    ElseIf PropSingleLine = True Then
        DrawFlags = DrawFlags Or DT_SINGLELINE
    End If
    Select Case PropEllipsisFormat
        Case FlexEllipsisFormatEnd
            DrawFlags = DrawFlags Or DT_END_ELLIPSIS
        Case FlexEllipsisFormatPath
            DrawFlags = DrawFlags Or DT_PATH_ELLIPSIS
        Case FlexEllipsisFormatWord
            DrawFlags = DrawFlags Or DT_WORD_ELLIPSIS
    End Select
    If Not VBFlexGridColsInfo(iCol).Format = vbNullString Then Text = Format$(Text, VBFlexGridColsInfo(iCol).Format, vbUseSystemDayOfWeek, vbUseSystem)
    If Not (DrawFlags And DT_SINGLELINE) = DT_SINGLELINE Then
        Dim CalcRect As RECT, Height As Long, Result As Long
        Select Case Alignment
            Case FlexAlignmentLeftCenter, FlexAlignmentCenterCenter, FlexAlignmentRightCenter, FlexAlignmentGeneral
                LSet CalcRect = TextRect
                Height = DrawText(hDC, StrPtr(Text), -1, CalcRect, DrawFlags Or DT_CALCRECT)
                Result = (((TextRect.Bottom - TextRect.Top) - Height) / 2)
            Case FlexAlignmentLeftBottom, FlexAlignmentCenterBottom, FlexAlignmentRightBottom
                LSet CalcRect = TextRect
                Height = DrawText(hDC, StrPtr(Text), -1, CalcRect, DrawFlags Or DT_CALCRECT)
                Result = ((TextRect.Bottom - TextRect.Top) - Height)
        End Select
        If Result > 0 Then TextRect.Top = TextRect.Top + Result
    Else
        Select Case Alignment
            Case FlexAlignmentLeftCenter, FlexAlignmentCenterCenter, FlexAlignmentRightCenter, FlexAlignmentGeneral
                DrawFlags = DrawFlags Or DT_VCENTER
            Case FlexAlignmentLeftBottom, FlexAlignmentCenterBottom, FlexAlignmentRightBottom
                DrawFlags = DrawFlags Or DT_BOTTOM
        End Select
    End If
    Dim Offset As Long, TempTextColor As Long
    Select Case TextStyle
        Case FlexTextStyleRaised
            TempTextColor = SetTextColor(hDC, &H808080)
            Offset = 1
        Case FlexTextStyleRaisedLight
            TempTextColor = SetTextColor(hDC, vbWhite)
            Offset = 1
        Case FlexTextStyleInset
            TempTextColor = SetTextColor(hDC, &H808080)
            Offset = -1
        Case FlexTextStyleInsetLight
            TempTextColor = SetTextColor(hDC, vbWhite)
            Offset = -1
    End Select
    If Offset <> 0 Then
        With TextRect
        .Top = .Top + Offset
        .Left = .Left + Offset
        .Bottom = .Bottom + Offset
        .Right = .Right + Offset
        End With
        DrawText hDC, StrPtr(Text), -1, TextRect, DrawFlags
        SetTextColor hDC, TempTextColor
        With TextRect
        .Top = .Top - Offset
        .Left = .Left - Offset
        .Bottom = .Bottom - Offset
        .Right = .Right - Offset
        End With
    End If
    DrawText hDC, StrPtr(Text), -1, TextRect, DrawFlags
End If
SetTextColor hDC, OldTextColor
If hFontOld <> 0 Then SelectObject hDC, hFontOld
If hFontTemp <> 0 Then DeleteObject hFontTemp
End With
End Sub

Private Sub GetSelRangeStruct(ByRef SelRange As TCELLRANGE)
With SelRange
If VBFlexGridRow > VBFlexGridRowSel Then .TopRow = VBFlexGridRowSel Else .TopRow = VBFlexGridRow
If VBFlexGridRowSel > VBFlexGridRow Then .BottomRow = VBFlexGridRowSel Else .BottomRow = VBFlexGridRow
If VBFlexGridCol > VBFlexGridColSel Then .LeftCol = VBFlexGridColSel Else .LeftCol = VBFlexGridCol
If VBFlexGridColSel > VBFlexGridCol Then .RightCol = VBFlexGridColSel Else .RightCol = VBFlexGridCol
End With
End Sub

Private Sub GetMergedRangeStruct(ByVal iRow As Long, ByVal iCol As Long, ByRef MergedRange As TCELLRANGE)
With MergedRange
.TopRow = iRow
.BottomRow = iRow
.LeftCol = iCol
.RightCol = iCol
If PropMergeCells <> FlexMergeCellsNever And PropRows > 0 And PropCols > 0 Then
    If iRow > (PropFixedRows - 1) And iCol > (PropFixedCols - 1) Then
        If (VBFlexGridCells.Rows(iRow).RowInfo.State And RWIS_MERGE) = RWIS_MERGE Then
            Select Case PropMergeCells
                Case FlexMergeCellsFree, FlexMergeCellsRestrictRows
                    If iCol > PropFixedCols Then
                        Do
                            If MergeCompareFunction(iRow, iCol, iRow, .LeftCol - 1) = True Then .LeftCol = .LeftCol - 1 Else Exit Do
                        Loop While .LeftCol > PropFixedCols
                    End If
                    If iCol < (PropCols - 1) Then
                        Do
                            If MergeCompareFunction(iRow, iCol, iRow, .RightCol + 1) = True Then .RightCol = .RightCol + 1 Else Exit Do
                        Loop While .RightCol < (PropCols - 1)
                    End If
                Case FlexMergeCellsRestrictColumns, FlexMergeCellsRestrictAll
                    If iCol > PropFixedCols Then
                        Do
                            If MergeCompareFunction(iRow, iCol, iRow, .LeftCol - 1) = True Then
                                If iRow > PropFixedRows Then
                                    If MergeCompareFunction(iRow - 1, iCol, iRow - 1, .LeftCol - 1) = True Then .LeftCol = .LeftCol - 1 Else Exit Do
                                Else
                                    .LeftCol = .LeftCol - 1
                                End If
                            Else
                                Exit Do
                            End If
                        Loop While .LeftCol > PropFixedCols
                    End If
                    If iCol < (PropCols - 1) Then
                        Do
                            If MergeCompareFunction(iRow, iCol, iRow, .RightCol + 1) = True Then
                                If iRow > PropFixedRows Then
                                    If MergeCompareFunction(iRow - 1, iCol, iRow - 1, .RightCol + 1) = True Then .RightCol = .RightCol + 1 Else Exit Do
                                Else
                                    .RightCol = .RightCol + 1
                                End If
                            Else
                                Exit Do
                            End If
                        Loop While .RightCol < (PropCols - 1)
                    End If
            End Select
        End If
        If (VBFlexGridColsInfo(iCol).State And CLIS_MERGE) = CLIS_MERGE Then
            Select Case PropMergeCells
                Case FlexMergeCellsFree, FlexMergeCellsRestrictColumns
                    If iRow > PropFixedRows Then
                        Do
                            If MergeCompareFunction(iRow, iCol, .TopRow - 1, iCol) = True Then .TopRow = .TopRow - 1 Else Exit Do
                        Loop While .TopRow > PropFixedRows
                    End If
                    If iRow < (PropRows - 1) Then
                        Do
                            If MergeCompareFunction(iRow, iCol, .BottomRow + 1, iCol) = True Then .BottomRow = .BottomRow + 1 Else Exit Do
                        Loop While .BottomRow < (PropRows - 1)
                    End If
                Case FlexMergeCellsRestrictRows, FlexMergeCellsRestrictAll
                    If iRow > PropFixedRows Then
                        Do
                            If MergeCompareFunction(iRow, iCol, .TopRow - 1, iCol) = True Then
                                If iCol > PropFixedCols Then
                                    If MergeCompareFunction(iRow, iCol - 1, .TopRow - 1, iCol - 1) = True Then .TopRow = .TopRow - 1 Else Exit Do
                                Else
                                    .TopRow = .TopRow - 1
                                End If
                            Else
                                Exit Do
                            End If
                        Loop While .TopRow > PropFixedRows
                    End If
                    If iRow < (PropRows - 1) Then
                        Do
                            If MergeCompareFunction(iRow, iCol, .BottomRow + 1, iCol) = True Then
                                If iCol > PropFixedCols Then
                                    If MergeCompareFunction(iRow, iCol - 1, .BottomRow + 1, iCol - 1) = True Then .BottomRow = .BottomRow + 1 Else Exit Do
                                Else
                                    .BottomRow = .BottomRow + 1
                                End If
                            Else
                                Exit Do
                            End If
                        Loop While .BottomRow < (PropRows - 1)
                    End If
            End Select
        End If
    Else
        If (VBFlexGridCells.Rows(iRow).RowInfo.State And RWIS_MERGE) = RWIS_MERGE Then
            Select Case PropMergeCells
                Case FlexMergeCellsFree, FlexMergeCellsRestrictRows, FlexMergeCellsFixedOnly
                    If iCol > 0 Then
                        Do
                            If MergeCompareFunction(iRow, iCol, iRow, .LeftCol - 1) = True Then .LeftCol = .LeftCol - 1 Else Exit Do
                        Loop While .LeftCol > 0
                    End If
                    If iCol < (PropCols - 1) Then
                        Do
                            If MergeCompareFunction(iRow, iCol, iRow, .RightCol + 1) = True Then .RightCol = .RightCol + 1 Else Exit Do
                        Loop While .RightCol < (PropCols - 1)
                    End If
                Case FlexMergeCellsRestrictColumns, FlexMergeCellsRestrictAll
                    If iCol > 0 Then
                        Do
                            If MergeCompareFunction(iRow, iCol, iRow, .LeftCol - 1) = True Then
                                If Row > 0 Then
                                    If MergeCompareFunction(iRow - 1, iCol, iRow - 1, .LeftCol - 1) = True Then .LeftCol = .LeftCol - 1 Else Exit Do
                                Else
                                    .LeftCol = .LeftCol - 1
                                End If
                            Else
                                Exit Do
                            End If
                        Loop While .LeftCol > 0
                    End If
                    If iCol < (PropCols - 1) Then
                        Do
                            If MergeCompareFunction(iRow, iCol, iRow, .RightCol + 1) = True Then
                                If Row > 0 Then
                                    If MergeCompareFunction(iRow - 1, iCol, iRow - 1, .RightCol + 1) = True Then .RightCol = .RightCol + 1 Else Exit Do
                                Else
                                    .RightCol = .RightCol + 1
                                End If
                            Else
                                Exit Do
                            End If
                        Loop While .RightCol < (PropCols - 1)
                    End If
            End Select
        End If
        If (VBFlexGridColsInfo(iCol).State And CLIS_MERGE) = CLIS_MERGE Then
            Select Case PropMergeCells
                Case FlexMergeCellsFree, FlexMergeCellsRestrictColumns, FlexMergeCellsFixedOnly
                    If iRow > 0 Then
                        Do
                            If MergeCompareFunction(iRow, iCol, .TopRow - 1, iCol) = True Then .TopRow = .TopRow - 1 Else Exit Do
                        Loop While .TopRow > 0
                    End If
                    If iRow < (PropRows - 1) Then
                        Do
                            If MergeCompareFunction(iRow, iCol, .BottomRow + 1, iCol) = True Then .BottomRow = .BottomRow + 1 Else Exit Do
                        Loop While .BottomRow < (PropRows - 1)
                    End If
                Case FlexMergeCellsRestrictRows, FlexMergeCellsRestrictAll
                    If iRow > 0 Then
                        Do
                            If MergeCompareFunction(iRow, iCol, .TopRow - 1, iCol) = True Then
                                If iCol > 0 Then
                                    If MergeCompareFunction(iRow, iCol - 1, .TopRow - 1, iCol - 1) = True Then .TopRow = .TopRow - 1 Else Exit Do
                                Else
                                    .TopRow = .TopRow - 1
                                End If
                            Else
                                Exit Do
                            End If
                        Loop While .TopRow > 0
                    End If
                    If iRow < (PropRows - 1) Then
                        Do
                            If MergeCompareFunction(iRow, iCol, .BottomRow + 1, iCol) = True Then
                                If iCol > 0 Then
                                    If MergeCompareFunction(iRow, iCol - 1, .BottomRow + 1, iCol - 1) = True Then .BottomRow = .BottomRow + 1 Else Exit Do
                                Else
                                    .BottomRow = .BottomRow + 1
                                End If
                            Else
                                Exit Do
                            End If
                        Loop While .BottomRow < (PropRows - 1)
                    End If
            End Select
        End If
    End If
    ' MergeCol overrules MergeRow.
    For iRow = .TopRow To .BottomRow
        If (VBFlexGridCells.Rows(iRow).RowInfo.State And RWIS_MERGE) = 0 Then
            .RightCol = .LeftCol
            Exit For
        End If
    Next iRow
End If
End With
End Sub

Private Sub GetCellRect(ByVal iRow As Long, ByVal iCol As Long, ByVal BorderOffset As Boolean, ByRef CellRect As RECT)
If PropRows < 1 Or PropCols < 1 Then Exit Sub
Dim i As Long
With CellRect
If BorderOffset = True Then
    Select Case PropBorderStyle
        Case FlexBorderStyleSingle, FlexBorderStyleThin
            .Top = GetSystemMetrics(SM_CYBORDER)
            .Left = GetSystemMetrics(SM_CXBORDER)
        Case FlexBorderStyleSunken, FlexBorderStyleRaised
            .Top = GetSystemMetrics(SM_CYEDGE)
            .Left = GetSystemMetrics(SM_CXEDGE)
    End Select
    .Bottom = .Top
    .Right = .Left
Else
    SetRect CellRect, 0, 0, 0, 0
End If
For i = 0 To iRow
    If i >= VBFlexGridTopRow Or i < (PropFixedRows + PropFrozenRows) Then
        .Top = .Bottom
        .Bottom = .Bottom + GetRowHeight(i)
    End If
Next i
For i = 0 To iCol
    If i >= VBFlexGridLeftCol Or i < (PropFixedCols + PropFrozenCols) Then
        .Left = .Right
        .Right = .Right + GetColWidth(i)
    End If
Next i
End With
End Sub

Private Sub GetCellRangeRect(ByRef CellRange As TCELLRANGE, ByVal BorderOffset As Boolean, ByRef CellRangeRect As RECT)
If PropRows < 1 Or PropCols < 1 Then Exit Sub
Dim i As Long
With CellRangeRect
If BorderOffset = True Then
    Select Case PropBorderStyle
        Case FlexBorderStyleSingle, FlexBorderStyleThin
            .Top = GetSystemMetrics(SM_CYBORDER)
            .Left = GetSystemMetrics(SM_CXBORDER)
        Case FlexBorderStyleSunken, FlexBorderStyleRaised
            .Top = GetSystemMetrics(SM_CYEDGE)
            .Left = GetSystemMetrics(SM_CXEDGE)
    End Select
    .Bottom = .Top
    .Right = .Left
Else
    SetRect CellRangeRect, 0, 0, 0, 0
End If
For i = 0 To CellRange.TopRow
    If i >= VBFlexGridTopRow Or i < (PropFixedRows + PropFrozenRows) Then
        .Top = .Bottom
        .Bottom = .Bottom + GetRowHeight(i)
    End If
Next i
For i = CellRange.TopRow + 1 To CellRange.BottomRow
    If i >= VBFlexGridTopRow Or i < (PropFixedRows + PropFrozenRows) Then .Bottom = .Bottom + GetRowHeight(i)
Next i
For i = 0 To CellRange.LeftCol
    If i >= VBFlexGridLeftCol Or i < (PropFixedCols + PropFrozenCols) Then
        .Left = .Right
        .Right = .Right + GetColWidth(i)
    End If
Next i
For i = CellRange.LeftCol + 1 To CellRange.RightCol
    If i >= VBFlexGridLeftCol Or i < (PropFixedCols + PropFrozenCols) Then .Right = .Right + GetColWidth(i)
Next i
End With
End Sub

Private Sub GetCellText(ByVal iRow As Long, ByVal iCol As Long, ByRef TextOut As String)
If PropRows < 1 Or PropCols < 1 Then Exit Sub

' ByRef parameter is faster than returning the string as the function return value.

#If ImplementFlexDataSource = True Then

If VBFlexGridFlexDataSource Is Nothing Then
    TextOut = VBFlexGridCells.Rows(iRow).Cols(iCol).Text
Else
    If iRow >= PropFixedRows Then
        TextOut = VBFlexGridFlexDataSource.GetData(iCol, iRow - PropFixedRows)
    Else
        TextOut = VBFlexGridCells.Rows(iRow).Cols(iCol).Text
    End If
End If

#Else

TextOut = VBFlexGridCells.Rows(iRow).Cols(iCol).Text

#End If

End Sub

Private Sub SetCellText(ByVal iRow As Long, ByVal iCol As Long, ByRef TextIn As String)
If PropRows < 1 Or PropCols < 1 Then Exit Sub

#If ImplementFlexDataSource = True Then

If VBFlexGridFlexDataSource Is Nothing Then
    VBFlexGridCells.Rows(iRow).Cols(iCol).Text = TextIn
Else
    If iRow >= PropFixedRows Then
        VBFlexGridFlexDataSource.SetData iCol, iRow - PropFixedRows, TextIn
    Else
        VBFlexGridCells.Rows(iRow).Cols(iCol).Text = TextIn
    End If
End If

#Else

VBFlexGridCells.Rows(iRow).Cols(iCol).Text = TextIn

#End If

End Sub

Private Function GetRowHeight(ByVal iRow As Long) As Long
If PropRows < 1 Or PropCols < 1 Then Exit Function
If (VBFlexGridCells.Rows(iRow).RowInfo.State And RWIS_HIDDEN) = 0 Then
    If VBFlexGridCells.Rows(iRow).RowInfo.Height = -1 Then
        If (iRow > (PropFixedRows - 1) And PropFixedCols = 0) Or VBFlexGridDefaultFixedRowHeight = -1 Then
            GetRowHeight = VBFlexGridDefaultRowHeight
        ElseIf iRow > (PropFixedRows - 1) Then
            If VBFlexGridDefaultRowHeight > VBFlexGridDefaultFixedRowHeight Then
                GetRowHeight = VBFlexGridDefaultRowHeight
            Else
                GetRowHeight = VBFlexGridDefaultFixedRowHeight
            End If
        Else
            GetRowHeight = VBFlexGridDefaultFixedRowHeight
        End If
    Else
        GetRowHeight = VBFlexGridCells.Rows(iRow).RowInfo.Height
    End If
    If GetRowHeight > 0 Then
        If PropRowHeightMin > 0 Then If GetRowHeight < PropRowHeightMin Then GetRowHeight = PropRowHeightMin
        If PropRowHeightMax > 0 Then If GetRowHeight > PropRowHeightMax Then GetRowHeight = PropRowHeightMax
    End If
End If
End Function

Private Function GetColWidth(ByVal iCol As Long) As Long
If PropRows < 1 Or PropCols < 1 Then Exit Function
If (VBFlexGridColsInfo(iCol).State And CLIS_HIDDEN) = 0 Then
    If VBFlexGridColsInfo(iCol).Width = -1 Then
        If iCol > (PropFixedCols - 1) Or VBFlexGridDefaultFixedColWidth = -1 Then
            GetColWidth = VBFlexGridDefaultColWidth
        Else
            GetColWidth = VBFlexGridDefaultFixedColWidth
        End If
    Else
        GetColWidth = VBFlexGridColsInfo(iCol).Width
    End If
    If GetColWidth > 0 Then
        If PropColWidthMin > 0 Then If GetColWidth < PropColWidthMin Then GetColWidth = PropColWidthMin
        If PropColWidthMax > 0 Then If GetColWidth > PropColWidthMax Then GetColWidth = PropColWidthMax
    End If
End If
End Function

Private Function GetTextSize(ByVal iRow As Long, ByVal iCol As Long, ByVal Text As String) As SIZEAPI
If VBFlexGridHandle = 0 Or (PropRows < 1 Or PropCols < 1) Then Exit Function
Dim hDC As Long
hDC = GetDC(VBFlexGridHandle)
If hDC <> 0 Then
    Dim hFontTemp As Long, hFontOld As Long
    With VBFlexGridCells.Rows(iRow).Cols(iCol)
    If .FontName = vbNullString Then
        If iRow > (PropFixedRows - 1) And iCol > (PropFixedCols - 1) Then
            hFontOld = SelectObject(hDC, VBFlexGridFontHandle)
        Else
            If VBFlexGridFontFixedHandle = 0 Then
                hFontOld = SelectObject(hDC, VBFlexGridFontHandle)
            Else
                hFontOld = SelectObject(hDC, VBFlexGridFontFixedHandle)
            End If
        End If
    Else
        Dim TempFont As StdFont
        Set TempFont = New StdFont
        TempFont.Name = .FontName
        TempFont.Size = .FontSize
        TempFont.Bold = .FontBold
        TempFont.Italic = .FontItalic
        TempFont.Strikethrough = .FontStrikeThrough
        TempFont.Underline = .FontUnderline
        TempFont.Charset = .FontCharset
        hFontTemp = CreateGDIFontFromOLEFont(TempFont)
        hFontOld = SelectObject(hDC, hFontTemp)
        Set TempFont = Nothing
    End If
    End With
    If Not Text = vbNullString Then
        If PropSingleLine = False Then
            If InStr(Text, vbCrLf) Then Text = Replace$(Text, vbCrLf, vbCr)
            If InStr(Text, vbLf) Then Text = Replace$(Text, vbLf, vbCr)
        Else
            If InStr(Text, vbCr) Then Text = Replace$(Text, vbCr, vbNullString)
            If InStr(Text, vbLf) Then Text = Replace$(Text, vbLf, vbNullString)
        End If
        Dim Pos1 As Long, Pos2 As Long, Temp As String, Size As SIZEAPI
        Do
            Pos1 = InStr(Pos1 + 1, Text, vbCr)
            If Pos1 > 0 Then
                Temp = Mid$(Text, Pos2 + 1, Pos1 - Pos2 - 1)
            Else
                Temp = Mid$(Text, Pos2 + 1)
            End If
            GetTextExtentPoint32 hDC, ByVal StrPtr(Temp), Len(Temp), Size
            With GetTextSize
            .CY = .CY + Size.CY
            If Size.CX > .CX Then .CX = Size.CX
            End With
            Pos2 = Pos1
        Loop Until Pos1 = 0
    Else
        Dim TM As TEXTMETRIC
        If GetTextMetrics(hDC, TM) <> 0 Then GetTextSize.CY = TM.TMHeight
    End If
    If hFontOld <> 0 Then SelectObject hDC, hFontOld
    If hFontTemp <> 0 Then DeleteObject hFontTemp
    ReleaseDC VBFlexGridHandle, hDC
End If
End Function

Private Function GetTextHeight(ByVal iRow As Long, ByVal iCol As Long, ByVal Text As String) As Long
If VBFlexGridHandle = 0 Or (PropRows < 1 Or PropCols < 1) Then Exit Function
Dim hDC As Long
hDC = GetDC(VBFlexGridHandle)
If hDC <> 0 Then
    Dim hFontTemp As Long, hFontOld As Long
    With VBFlexGridCells.Rows(iRow).Cols(iCol)
    If .FontName = vbNullString Then
        If iRow > (PropFixedRows - 1) And iCol > (PropFixedCols - 1) Then
            hFontOld = SelectObject(hDC, VBFlexGridFontHandle)
        Else
            If VBFlexGridFontFixedHandle = 0 Then
                hFontOld = SelectObject(hDC, VBFlexGridFontHandle)
            Else
                hFontOld = SelectObject(hDC, VBFlexGridFontFixedHandle)
            End If
        End If
    Else
        Dim TempFont As StdFont
        Set TempFont = New StdFont
        TempFont.Name = .FontName
        TempFont.Size = .FontSize
        TempFont.Bold = .FontBold
        TempFont.Italic = .FontItalic
        TempFont.Strikethrough = .FontStrikeThrough
        TempFont.Underline = .FontUnderline
        TempFont.Charset = .FontCharset
        hFontTemp = CreateGDIFontFromOLEFont(TempFont)
        hFontOld = SelectObject(hDC, hFontTemp)
        Set TempFont = Nothing
    End If
    End With
    If Not Text = vbNullString Then
        Dim TextRect As RECT, Format As Long
        With TextRect
        .Left = (CELL_TEXT_WIDTH_PADDING_DIP * PixelsPerDIP_X())
        .Top = (CELL_TEXT_HEIGHT_PADDING_DIP * PixelsPerDIP_Y())
        .Right = GetColWidth(iCol) - (CELL_TEXT_WIDTH_PADDING_DIP * PixelsPerDIP_X())
        .Bottom = GetRowHeight(iRow) - (CELL_TEXT_HEIGHT_PADDING_DIP * PixelsPerDIP_Y())
        End With
        Format = DT_NOPREFIX
        If VBFlexGridRTLReading = True Then Format = Format Or DT_RTLREADING
        ' Alignment format will be ignored.
        If PropWordWrap = True Then
            Format = Format Or DT_WORDBREAK
        ElseIf PropSingleLine = True Then
            Format = Format Or DT_SINGLELINE
        End If
        ' Ellipsis format will be ignored.
        GetTextHeight = DrawText(hDC, StrPtr(Text), -1, TextRect, Format Or DT_CALCRECT)
    Else
        Dim TM As TEXTMETRIC
        If GetTextMetrics(hDC, TM) <> 0 Then GetTextHeight = TM.TMHeight
    End If
    If hFontOld <> 0 Then SelectObject hDC, hFontOld
    If hFontTemp <> 0 Then DeleteObject hFontTemp
    ReleaseDC VBFlexGridHandle, hDC
End If
End Function

Private Sub GetHitTestInfo(ByRef HTI As THITTESTINFO)
HTI.HitRow = -1
HTI.HitCol = -1
HTI.HitRowDivider = -1
HTI.HitColDivider = -1
HTI.HitResult = FlexHitResultNoWhere
HTI.MouseRow = 0
HTI.MouseCol = 0
If VBFlexGridHandle = 0 Or (PropRows < 1 Or PropCols < 1) Or (HTI.PT.X < 0 And HTI.PT.Y < 0) Then Exit Sub
Dim iRow As Long, iCol As Long, iRowTo As Long, iColTo As Long
Dim iRowHit As Long, iColHit As Long, iRowDivider As Long, iColDivider As Long
Dim CellRect As RECT, TempRect As RECT
iRowHit = -1
iColHit = -1
iRowDivider = -1
iColDivider = -1
With CellRect
If HTI.PT.Y >= 0 Then iRowTo = (PropRows - 1) Else iRowTo = 0
For iRow = 0 To iRowTo
    If iRow >= VBFlexGridTopRow Or iRow < (PropFixedRows + PropFrozenRows) Then
        .Top = .Bottom
        .Bottom = .Top + GetRowHeight(iRow)
        If HTI.PT.Y >= .Top Then
            HTI.MouseRow = iRow
            If HTI.PT.Y < .Bottom Then
                iRowHit = iRow
                Exit For
            End If
        End If
    Else
        iRow = VBFlexGridTopRow - 1
    End If
Next iRow
If HTI.PT.X >= 0 Then iColTo = (PropCols - 1) Else iColTo = 0
For iCol = 0 To iColTo
    If iCol >= VBFlexGridLeftCol Or iCol < (PropFixedCols + PropFrozenCols) Then
        .Left = .Right
        .Right = .Left + GetColWidth(iCol)
        If HTI.PT.X >= .Left Then
            HTI.MouseCol = iCol
            If HTI.PT.X < .Right Then
                iColHit = iCol
                Exit For
            End If
        End If
    Else
        iCol = VBFlexGridLeftCol - 1
    End If
Next iCol
If iRowHit > -1 And iColHit > -1 Then
    If iRowHit >= VBFlexGridTopRow Then
        If iColHit >= VBFlexGridLeftCol Then
            HTI.HitResult = FlexHitResultCell
        ElseIf iColHit < PropFixedCols Then
            If PropAllowUserResizing = FlexAllowUserResizingRows Or PropAllowUserResizing = FlexAllowUserResizingBoth Then
                SetRect TempRect, .Left, .Top, .Right, .Bottom
                Call AdjustRectRowDividerSpacing(TempRect, iRowHit)
                If PtInRect(TempRect, HTI.PT.X, HTI.PT.Y) <> 0 Then
                    HTI.HitResult = FlexHitResultCell
                Else
                    TempRect.Bottom = .Bottom
                    iRowDivider = iRowHit
                    If PtInRect(TempRect, HTI.PT.X, HTI.PT.Y) = 0 Then
                        HTI.HitResult = FlexHitResultDividerRowTop
                        iRowDivider = iRowDivider - 1
                        Do While (VBFlexGridCells.Rows(iRowDivider).RowInfo.State And RWIS_HIDDEN) = RWIS_HIDDEN
                            iRowDivider = iRowDivider - 1
                            If iRowDivider = -1 Then Exit Do
                        Loop
                        If iRowDivider = -1 Then HTI.HitResult = FlexHitResultCell
                    Else
                        HTI.HitResult = FlexHitResultDividerRowBottom
                    End If
                End If
            Else
                HTI.HitResult = FlexHitResultCell
            End If
        ElseIf iColHit < (PropFixedCols + PropFrozenCols) Then
            HTI.HitResult = FlexHitResultCell
        End If
    ElseIf (iRowHit < PropFixedRows) Then
        If PropAllowUserResizing <> FlexAllowUserResizingNone Then
            SetRect TempRect, .Left, .Top, .Right, .Bottom
            Call AdjustRectColDividerSpacing(TempRect, iColHit)
            If PtInRect(TempRect, HTI.PT.X, HTI.PT.Y) <> 0 Then
                If iColHit < PropFixedCols Then
                    If PropAllowUserResizing <> FlexAllowUserResizingColumns Then
                        SetRect TempRect, .Left, .Top, .Right, .Bottom
                        Call AdjustRectRowDividerSpacing(TempRect, iRowHit)
                        If PtInRect(TempRect, HTI.PT.X, HTI.PT.Y) <> 0 Then
                            HTI.HitResult = FlexHitResultCell
                        Else
                            TempRect.Bottom = .Bottom
                            iRowDivider = iRowHit
                            If PtInRect(TempRect, HTI.PT.X, HTI.PT.Y) = 0 Then
                                HTI.HitResult = FlexHitResultDividerRowTop
                                iRowDivider = iRowDivider - 1
                                Do While (VBFlexGridCells.Rows(iRowDivider).RowInfo.State And RWIS_HIDDEN) = RWIS_HIDDEN
                                    iRowDivider = iRowDivider - 1
                                    If iRowDivider = -1 Then Exit Do
                                Loop
                                If iRowDivider = -1 Then HTI.HitResult = FlexHitResultCell
                            Else
                                HTI.HitResult = FlexHitResultDividerRowBottom
                            End If
                        End If
                    Else
                        HTI.HitResult = FlexHitResultCell
                    End If
                Else
                    HTI.HitResult = FlexHitResultCell
                End If
            ElseIf PropAllowUserResizing <> FlexAllowUserResizingRows Then
                TempRect.Right = .Right
                iColDivider = iColHit
                If PtInRect(TempRect, HTI.PT.X, HTI.PT.Y) = 0 Then
                    HTI.HitResult = FlexHitResultDividerColumnLeft
                    iColDivider = iColDivider - 1
                    Do While (VBFlexGridColsInfo(iColDivider).State And CLIS_HIDDEN) = CLIS_HIDDEN
                        iColDivider = iColDivider - 1
                        If iColDivider = -1 Then Exit Do
                    Loop
                    If iColDivider = -1 Then HTI.HitResult = FlexHitResultCell
                Else
                    HTI.HitResult = FlexHitResultDividerColumnRight
                End If
            Else
                HTI.HitResult = FlexHitResultCell
            End If
        Else
            HTI.HitResult = FlexHitResultCell
        End If
    ElseIf iRowHit < (PropFixedRows + PropFrozenRows) Then
        If iColHit >= VBFlexGridLeftCol Then
            HTI.HitResult = FlexHitResultCell
        ElseIf iColHit < PropFixedCols Then
            If PropAllowUserResizing = FlexAllowUserResizingRows Or PropAllowUserResizing = FlexAllowUserResizingBoth Then
                SetRect TempRect, .Left, .Top, .Right, .Bottom
                Call AdjustRectRowDividerSpacing(TempRect, iRowHit)
                If PtInRect(TempRect, HTI.PT.X, HTI.PT.Y) <> 0 Then
                    HTI.HitResult = FlexHitResultCell
                Else
                    TempRect.Bottom = .Bottom
                    iRowDivider = iRowHit
                    If PtInRect(TempRect, HTI.PT.X, HTI.PT.Y) = 0 Then
                        HTI.HitResult = FlexHitResultDividerRowTop
                        iRowDivider = iRowDivider - 1
                        Do While (VBFlexGridCells.Rows(iRowDivider).RowInfo.State And RWIS_HIDDEN) = RWIS_HIDDEN
                            iRowDivider = iRowDivider - 1
                            If iRowDivider = -1 Then Exit Do
                        Loop
                        If iRowDivider = -1 Then HTI.HitResult = FlexHitResultCell
                    Else
                        HTI.HitResult = FlexHitResultDividerRowBottom
                    End If
                End If
            Else
                HTI.HitResult = FlexHitResultCell
            End If
        ElseIf iColHit < (PropFixedCols + PropFrozenCols) Then
            HTI.HitResult = FlexHitResultCell
        End If
    End If
Else
    If PropAllowUserResizing <> FlexAllowUserResizingNone Then
        If iColHit > -1 And PropAllowUserResizing <> FlexAllowUserResizingColumns Then
            If iRowHit = -1 And iColHit < PropFixedCols Then
                If HTI.PT.Y < (.Bottom + (DIVIDER_SPACING_DIP * PixelsPerDIP_Y())) Then
                    iRowDivider = (PropRows - 1)
                    HTI.HitResult = FlexHitResultDividerRowBottom
                    Do While (VBFlexGridCells.Rows(iRowDivider).RowInfo.State And RWIS_HIDDEN) = RWIS_HIDDEN
                        iRowDivider = iRowDivider - 1
                        If iRowDivider = -1 Then Exit Do
                    Loop
                    If iRowDivider = -1 Then HTI.HitResult = FlexHitResultNoWhere
                End If
            End If
        ElseIf iRowHit > -1 And PropAllowUserResizing <> FlexAllowUserResizingRows Then
            If iColHit = -1 And iRowHit < PropFixedRows Then
                If HTI.PT.X < (.Right + (DIVIDER_SPACING_DIP * PixelsPerDIP_X())) Then
                    iColDivider = (PropCols - 1)
                    HTI.HitResult = FlexHitResultDividerColumnRight
                    Do While (VBFlexGridColsInfo(iColDivider).State And CLIS_HIDDEN) = CLIS_HIDDEN
                        iColDivider = iColDivider - 1
                        If iColDivider = -1 Then Exit Do
                    Loop
                    If iColDivider = -1 Then HTI.HitResult = FlexHitResultNoWhere
                End If
            End If
        End If
    End If
    iRowHit = -1
    iColHit = -1
End If
End With
If HTI.HitResult <> FlexHitResultNoWhere Then
    HTI.HitRow = iRowHit
    HTI.HitCol = iColHit
    Select Case HTI.HitResult
        Case FlexHitResultDividerRowTop, FlexHitResultDividerRowBottom, FlexHitResultDividerColumnLeft, FlexHitResultDividerColumnRight
            HTI.HitRowDivider = iRowDivider
            HTI.HitColDivider = iColDivider
    End Select
End If
End Sub

Private Sub AdjustRectColDividerSpacing(ByRef RC As RECT, ByVal iCol As Long)
Dim Spacing As Long
Spacing = DIVIDER_SPACING_DIP * PixelsPerDIP_X()
If iCol > 0 Then
    If (RC.Right - RC.Left) >= (Spacing * 2) Then
        RC.Left = RC.Left + Spacing
        RC.Right = RC.Right - Spacing
    Else
        ' Rectangle is not wide enough to include the spacing.
        RC.Left = RC.Left + ((RC.Right - RC.Left) / 2)
        RC.Right = RC.Left ' Remainder
    End If
ElseIf iCol > -1 Then
    ' First column need divider spacing to the right only.
    RC.Right = RC.Right - Spacing
    If RC.Right < RC.Left Then RC.Right = RC.Left
End If
End Sub

Private Sub AdjustRectRowDividerSpacing(ByRef RC As RECT, ByVal iRow As Long)
Dim Spacing As Long
Spacing = DIVIDER_SPACING_DIP * PixelsPerDIP_Y()
If iRow > 0 Then
    If (RC.Bottom - RC.Top) >= (Spacing * 2) Then
        RC.Top = RC.Top + Spacing
        RC.Bottom = RC.Bottom - Spacing
    Else
        ' Rectangle is not wide enough to include the spacing.
        RC.Top = RC.Top + ((RC.Bottom - RC.Top) / 2)
        RC.Bottom = RC.Top ' Remainder
    End If
ElseIf iRow > -1 Then
    ' First row need divider spacing to the bottom only.
    RC.Bottom = RC.Bottom - Spacing
    If RC.Bottom < RC.Top Then RC.Bottom = RC.Top
End If
End Sub

Private Sub GetLabelInfo(ByVal iRow As Long, ByVal iCol As Long, ByRef LBLI As TLABELINFO)
LBLI.Flags = 0
If VBFlexGridHandle = 0 Or (PropRows < 1 Or PropCols < 1) Then Exit Sub
Dim CellRect As RECT
Call GetCellRect(iRow, iCol, False, CellRect)
If (CellRect.Bottom - CellRect.Top) <= 0 Or (CellRect.Right - CellRect.Left) <= 0 Then Exit Sub
Dim hDC As Long
hDC = GetDC(VBFlexGridHandle)
If hDC <> 0 Then
    Dim Text As String
    Call GetCellText(iRow, iCol, Text)
    If StrPtr(Text) = 0 Then Text = ""
    Dim hFontTemp As Long, hFontOld As Long
    With VBFlexGridCells.Rows(iRow).Cols(iCol)
    If .FontName = vbNullString Then
        If iRow > (PropFixedRows - 1) And iCol > (PropFixedCols - 1) Then
            hFontOld = SelectObject(hDC, VBFlexGridFontHandle)
        Else
            If VBFlexGridFontFixedHandle = 0 Then
                hFontOld = SelectObject(hDC, VBFlexGridFontHandle)
            Else
                hFontOld = SelectObject(hDC, VBFlexGridFontFixedHandle)
            End If
        End If
    Else
        Dim TempFont As StdFont
        Set TempFont = New StdFont
        TempFont.Name = .FontName
        TempFont.Size = .FontSize
        TempFont.Bold = .FontBold
        TempFont.Italic = .FontItalic
        TempFont.Strikethrough = .FontStrikeThrough
        TempFont.Underline = .FontUnderline
        TempFont.Charset = .FontCharset
        hFontTemp = CreateGDIFontFromOLEFont(TempFont)
        hFontOld = SelectObject(hDC, hFontTemp)
        Set TempFont = Nothing
    End If
    End With
    Dim TextRect As RECT, Alignment As FlexAlignmentConstants, Format As Long
    With TextRect
    .Left = CellRect.Left + (CELL_TEXT_WIDTH_PADDING_DIP * PixelsPerDIP_X())
    .Top = CellRect.Top + (CELL_TEXT_HEIGHT_PADDING_DIP * PixelsPerDIP_Y())
    .Right = CellRect.Right - (CELL_TEXT_WIDTH_PADDING_DIP * PixelsPerDIP_X())
    .Bottom = CellRect.Bottom - (CELL_TEXT_HEIGHT_PADDING_DIP * PixelsPerDIP_Y())
    End With
    If VBFlexGridCells.Rows(iRow).Cols(iCol).Alignment = -1 Then
        If iRow > (PropFixedRows - 1) And iCol > (PropFixedCols - 1) Then
            Alignment = VBFlexGridColsInfo(iCol).Alignment
        Else
            If VBFlexGridColsInfo(iCol).FixedAlignment = -1 Then
                Alignment = VBFlexGridColsInfo(iCol).Alignment
            Else
                Alignment = VBFlexGridColsInfo(iCol).FixedAlignment
            End If
        End If
    Else
        Alignment = VBFlexGridCells.Rows(iRow).Cols(iCol).Alignment
    End If
    Format = DT_NOPREFIX
    If VBFlexGridRTLReading = True Then Format = Format Or DT_RTLREADING
    Select Case Alignment
        Case FlexAlignmentLeftTop, FlexAlignmentLeftCenter, FlexAlignmentLeftBottom
            Format = Format Or DT_LEFT
        Case FlexAlignmentCenterTop, FlexAlignmentCenterCenter, FlexAlignmentCenterBottom
            Format = Format Or DT_CENTER
        Case FlexAlignmentRightTop, FlexAlignmentRightCenter, FlexAlignmentRightBottom
            Format = Format Or DT_RIGHT
        Case FlexAlignmentGeneral
            If Not IsNumeric(Text) Then
                Format = Format Or DT_LEFT
            Else
                Format = Format Or DT_RIGHT
            End If
    End Select
    If PropWordWrap = True Then
        Format = Format Or DT_WORDBREAK
    ElseIf PropSingleLine = True Then
        Format = Format Or DT_SINGLELINE
    End If
    ' Ellipsis format will be ignored.
    Dim CalcRect As RECT, Height As Long, Result As Long
    LSet CalcRect = TextRect
    Select Case Alignment
        Case FlexAlignmentLeftCenter, FlexAlignmentCenterCenter, FlexAlignmentRightCenter, FlexAlignmentGeneral
            Height = DrawText(hDC, StrPtr(Text), -1, CalcRect, Format Or DT_CALCRECT)
            Result = (((TextRect.Bottom - TextRect.Top) - Height) / 2)
            ' DT_VCENTER not applicable to apply here in case of DT_SINGLELINE.
        Case FlexAlignmentLeftBottom, FlexAlignmentCenterBottom, FlexAlignmentRightBottom
            Height = DrawText(hDC, StrPtr(Text), -1, CalcRect, Format Or DT_CALCRECT)
            Result = ((TextRect.Bottom - TextRect.Top) - Height)
            ' DT_BOTTOM not applicable to apply here in case of DT_SINGLELINE.
    End Select
    If Result > 0 Or (Format And DT_SINGLELINE) = DT_SINGLELINE Then
        CalcRect.Top = CalcRect.Top + Result
        CalcRect.Bottom = CalcRect.Bottom + Result
    End If
    With LBLI
    .Flags = LBLI_VALID
    If TextRect.Right <= VBFlexGridClientRect.Right And TextRect.Bottom <= VBFlexGridClientRect.Bottom Then
        If CalcRect.Right <= TextRect.Right And CalcRect.Bottom <= TextRect.Bottom Then .Flags = .Flags Or LBLI_UNFOLDED
    End If
    If (Format And DT_CENTER) = DT_CENTER Then
        Result = (((TextRect.Right - TextRect.Left) - (CalcRect.Right - CalcRect.Left)) / 2)
        CalcRect.Left = CalcRect.Left + Result
        CalcRect.Right = CalcRect.Right + Result
    ElseIf (Format And DT_RIGHT) = DT_RIGHT Then
        Result = ((TextRect.Right - TextRect.Left) - (CalcRect.Right - CalcRect.Left))
        CalcRect.Left = CalcRect.Left + Result
        CalcRect.Right = CalcRect.Right + Result
    End If
    LSet .RC = CalcRect
    .DrawFlags = Format
    End With
    If hFontOld <> 0 Then SelectObject hDC, hFontOld
    If hFontTemp <> 0 Then DeleteObject hFontTemp
    ReleaseDC VBFlexGridHandle, hDC
End If
End Sub

Private Sub SetScrollBars()
Static InProc As Boolean
If VBFlexGridHandle = 0 Or InProc = True Or VBFlexGridNoRedraw = True Then Exit Sub
InProc = True
Dim dwStyleOld As Long, dwStyleNew As Long, dwStyleTemp As Long
dwStyleOld = GetWindowLong(VBFlexGridHandle, GWL_STYLE)
dwStyleNew = dwStyleOld
dwStyleTemp = dwStyleOld
If (dwStyleNew And WS_HSCROLL) = WS_HSCROLL Then dwStyleNew = dwStyleNew And Not WS_HSCROLL
If (dwStyleNew And WS_VSCROLL) = WS_VSCROLL Then dwStyleNew = dwStyleNew And Not WS_VSCROLL
If (PropRows > 0 And PropCols > 0) And PropScrollBars <> vbSBNone Then
    Select Case PropScrollBars
        Case vbHorizontal
            dwStyleNew = dwStyleNew Or WS_HSCROLL
        Case vbVertical
            dwStyleNew = dwStyleNew Or WS_VSCROLL
        Case vbBoth
            dwStyleNew = dwStyleNew Or WS_HSCROLL Or WS_VSCROLL
    End Select
Else
    If dwStyleNew <> dwStyleOld Then
        SetWindowLong VBFlexGridHandle, GWL_STYLE, dwStyleNew
        SetWindowPos VBFlexGridHandle, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
    End If
    InProc = False
    Exit Sub
End If
Dim SCI(0 To 1) As SCROLLINFO, iRow As Long, iCol As Long
Dim ClientRect As RECT, GridRect As RECT, Changed As Boolean
SCI(0).cbSize = LenB(SCI(0))
SCI(0).fMask = SIF_RANGE Or SIF_PAGE
If PropDisableNoScroll = True Then SCI(0).fMask = SCI(0).fMask Or SIF_DISABLENOSCROLL
LSet SCI(1) = SCI(0)
LSet ClientRect = VBFlexGridClientRect
With GridRect
Do
    If PropScrollBars = vbHorizontal Or PropScrollBars = vbBoth Then
        SCI(0).nMin = 0
        SCI(0).nMax = 0
        ' nPage of 0 is appropriate when the columns vary in width.
        ' But then nMax needs be adjusted in a second step.
        SCI(0).nPage = 0
        .Right = 0
        For iCol = 0 To (PropCols - 1)
            .Right = .Right + GetColWidth(iCol)
            If .Right > ClientRect.Right And iCol > (PropFixedCols + PropFrozenCols) Then
                SCI(0).nMax = (PropCols - (PropFixedCols + PropFrozenCols)) - 1
                ' Scroll box is proportional to the scrolling region.
                ' But only appropriate when all columns are equally in width.
                ' SCI(0).nPage = iCol - ((PropFixedCols + PropFrozenCols) - 1) - 1
                Exit For
            End If
        Next iCol
        If SCI(0).nMax > 0 And SCI(0).nPage = 0 Then
            .Right = 0
            For iCol = 0 To ((PropFixedCols + PropFrozenCols) - 1)
                .Right = .Right + GetColWidth(iCol)
            Next iCol
            For iCol = (PropCols - 1) To (PropFixedCols + PropFrozenCols) Step -1
                .Right = .Right + GetColWidth(iCol)
                If .Right > ClientRect.Right And iCol < (PropCols - 1) Then
                    SCI(0).nMax = SCI(0).nMax - ((PropCols - 1) - iCol) + 1
                    Exit For
                End If
            Next iCol
        End If
        If SCI(0).nMax = 0 And PropDisableNoScroll = False Then
            If (dwStyleNew And WS_HSCROLL) = WS_HSCROLL Then dwStyleNew = dwStyleNew And Not WS_HSCROLL
        End If
    End If
    If PropScrollBars = vbVertical Or PropScrollBars = vbBoth Then
        SCI(1).nMin = 0
        SCI(1).nMax = 0
        ' nPage of 0 is appropriate when the rows vary in height.
        ' But then nMax needs be adjusted in a second step.
        SCI(1).nPage = 0
        .Bottom = 0
        For iRow = 0 To (PropRows - 1)
            .Bottom = .Bottom + GetRowHeight(iRow)
            If .Bottom > ClientRect.Bottom And iRow > (PropFixedRows + PropFrozenRows) Then
                SCI(1).nMax = (PropRows - (PropFixedRows + PropFrozenRows)) - 1
                ' Scroll box is proportional to the scrolling region.
                ' But only appropriate when all rows are equally in height.
                ' SCI(1).nPage = iRow - ((PropFixedRows + PropFrozenRows) - 1) - 1
                Exit For
            End If
        Next iRow
        If SCI(1).nMax > 0 And SCI(1).nPage = 0 Then
            .Bottom = 0
            For iRow = 0 To ((PropFixedRows + PropFrozenRows) - 1)
                .Bottom = .Bottom + GetRowHeight(iRow)
            Next iRow
            For iRow = (PropRows - 1) To (PropFixedRows + PropFrozenRows) Step -1
                .Bottom = .Bottom + GetRowHeight(iRow)
                If .Bottom > ClientRect.Bottom And iRow < (PropRows - 1) Then
                    SCI(1).nMax = SCI(1).nMax - ((PropRows - 1) - iRow) + 1
                    Exit For
                End If
            Next iRow
        End If
        If SCI(1).nMax = 0 And PropDisableNoScroll = False Then
            If (dwStyleNew And WS_VSCROLL) = WS_VSCROLL Then dwStyleNew = dwStyleNew And Not WS_VSCROLL
        End If
    End If
    Changed = CBool(dwStyleNew <> dwStyleTemp)
    If (dwStyleNew And WS_VSCROLL) = WS_VSCROLL And (dwStyleTemp And WS_VSCROLL) = 0 Then
        ClientRect.Right = ClientRect.Right - GetSystemMetrics(SM_CXVSCROLL)
    ElseIf (dwStyleNew And WS_VSCROLL) = 0 And (dwStyleTemp And WS_VSCROLL) = WS_VSCROLL Then
        ClientRect.Right = ClientRect.Right + GetSystemMetrics(SM_CXVSCROLL)
    End If
    If (dwStyleNew And WS_HSCROLL) = WS_HSCROLL And (dwStyleTemp And WS_HSCROLL) = 0 Then
        ClientRect.Bottom = ClientRect.Bottom - GetSystemMetrics(SM_CYHSCROLL)
    ElseIf (dwStyleNew And WS_HSCROLL) = 0 And (dwStyleTemp And WS_HSCROLL) = WS_HSCROLL Then
        ClientRect.Bottom = ClientRect.Bottom + GetSystemMetrics(SM_CYHSCROLL)
    End If
    dwStyleTemp = dwStyleNew
Loop Until Changed = False
End With
If dwStyleNew <> dwStyleOld Then
    SetWindowLong VBFlexGridHandle, GWL_STYLE, dwStyleNew
    SetWindowPos VBFlexGridHandle, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
End If
SetScrollInfo VBFlexGridHandle, SB_HORZ, SCI(0), 0
SetScrollInfo VBFlexGridHandle, SB_VERT, SCI(1), 0
SetWindowPos VBFlexGridHandle, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_DRAWFRAME
InProc = False
End Sub

Private Function GetRowsPerPage(ByVal TopRow As Long) As Long
If VBFlexGridHandle = 0 Or (PropRows < 1 Or PropCols < 1) Then Exit Function
Dim GridRect As RECT, iRow As Long, Count As Long
With GridRect
For iRow = 0 To ((PropFixedRows + PropFrozenRows) - 1)
    .Bottom = .Bottom + GetRowHeight(iRow)
Next iRow
For iRow = TopRow To (PropRows - 1)
    .Bottom = .Bottom + GetRowHeight(iRow)
    If iRow > TopRow And .Bottom > VBFlexGridClientRect.Bottom Then Exit For
    Count = Count + 1
Next iRow
GetRowsPerPage = Count
End With
End Function

Private Function GetRowsPerPageRev(ByVal BottomRow As Long) As Long
If VBFlexGridHandle = 0 Or (PropRows < 1 Or PropCols < 1) Then Exit Function
Dim GridRect As RECT, iRow As Long, Count As Long
With GridRect
For iRow = 0 To ((PropFixedRows + PropFrozenRows) - 1)
    .Bottom = .Bottom + GetRowHeight(iRow)
Next iRow
For iRow = BottomRow To (PropFixedRows + PropFrozenRows) Step -1
    .Bottom = .Bottom + GetRowHeight(iRow)
    If iRow < BottomRow And .Bottom > VBFlexGridClientRect.Bottom Then Exit For
    Count = Count + 1
Next iRow
GetRowsPerPageRev = Count
End With
End Function

Private Function GetColsPerPage(ByVal LeftCol As Long) As Long
If VBFlexGridHandle = 0 Or (PropRows < 1 Or PropCols < 1) Then Exit Function
Dim GridRect As RECT, iCol As Long, Count As Long
With GridRect
For iCol = 0 To ((PropFixedCols + PropFrozenCols) - 1)
    .Right = .Right + GetColWidth(iCol)
Next iCol
For iCol = LeftCol To (PropCols - 1)
    .Right = .Right + GetColWidth(iCol)
    If iCol > LeftCol And .Right > VBFlexGridClientRect.Right Then Exit For
    Count = Count + 1
Next iCol
GetColsPerPage = Count
End With
End Function

Private Function GetColsPerPageRev(ByVal RightCol As Long) As Long
If VBFlexGridHandle = 0 Or (PropRows < 1 Or PropCols < 1) Then Exit Function
Dim GridRect As RECT, iCol As Long, Count As Long
With GridRect
For iCol = 0 To ((PropFixedCols + PropFrozenCols) - 1)
    .Right = .Right + GetColWidth(iCol)
Next iCol
For iCol = RightCol To (PropFixedCols + PropFrozenCols) Step -1
    .Right = .Right + GetColWidth(iCol)
    If iCol < RightCol And .Right > VBFlexGridClientRect.Right Then Exit For
    Count = Count + 1
Next iCol
GetColsPerPageRev = Count
End With
End Function

Private Function CheckScrollPos(ByVal wBar As Long) As Boolean
If VBFlexGridHandle = 0 Or (PropRows < 1 Or PropCols < 1) Then Exit Function
Dim dwStyle As Long
dwStyle = GetWindowLong(VBFlexGridHandle, GWL_STYLE)
If Not ((wBar = SB_HORZ And (dwStyle And WS_HSCROLL) = WS_HSCROLL) Or (wBar = SB_VERT And (dwStyle And WS_VSCROLL) = WS_VSCROLL)) Then Exit Function
Dim SCI As SCROLLINFO, PrevPos As Long
SCI.cbSize = LenB(SCI)
SCI.fMask = SIF_POS
GetScrollInfo VBFlexGridHandle, wBar, SCI
If wBar = SB_HORZ Then
    CheckScrollPos = CBool((VBFlexGridLeftCol - (PropFixedCols + PropFrozenCols)) <> SCI.nPos)
ElseIf wBar = SB_VERT Then
    CheckScrollPos = CBool((VBFlexGridTopRow - (PropFixedRows + PropFrozenRows)) <> SCI.nPos)
End If
If CheckScrollPos = False Then Exit Function
PrevPos = SCI.nPos
If wBar = SB_HORZ Then
    SCI.nPos = VBFlexGridLeftCol - (PropFixedCols + PropFrozenCols)
ElseIf wBar = SB_VERT Then
    SCI.nPos = VBFlexGridTopRow - (PropFixedRows + PropFrozenRows)
End If
SetScrollInfo VBFlexGridHandle, wBar, SCI, IIf(VBFlexGridNoRedraw = False, 1, 0)
GetScrollInfo VBFlexGridHandle, wBar, SCI
If PrevPos <> SCI.nPos Then
    Call RedrawGrid
    If PropShowInfoTips = True Or PropShowLabelTips = True Then
        Dim Pos As Long
        Pos = GetMessagePos()
        Call CheckToolTipRowCol(Get_X_lParam(Pos), Get_Y_lParam(Pos))
    End If
    If VBFlexGridEditRow > -1 And VBFlexGridEditCol > -1 Then Call UpdateEditRect
    RaiseEvent Scroll
End If
End Function

Private Sub CheckTopRow(ByRef TopRow As Long)
If VBFlexGridHandle = 0 Or (PropRows < 1 Or PropCols < 1) Then Exit Sub
Dim GridRect As RECT, iRow As Long
With GridRect
For iRow = 0 To ((PropFixedRows + PropFrozenRows) - 1)
    .Bottom = .Bottom + GetRowHeight(iRow)
Next iRow
For iRow = TopRow To (PropRows - 1)
    .Bottom = .Bottom + GetRowHeight(iRow)
    If .Bottom > VBFlexGridClientRect.Bottom Then Exit For
Next iRow
If .Bottom <= VBFlexGridClientRect.Bottom Then
    Do While TopRow > (((PropFixedRows + PropFrozenRows) - 1) + 1)
        .Bottom = .Bottom + GetRowHeight(TopRow - 1)
        If .Bottom > VBFlexGridClientRect.Bottom Then
            Exit Do
        Else
            TopRow = TopRow - 1
        End If
    Loop
End If
End With
End Sub

Private Sub CheckLeftCol(ByRef LeftCol As Long)
If VBFlexGridHandle = 0 Or (PropRows < 1 Or PropCols < 1) Then Exit Sub
Dim GridRect As RECT, iCol As Long
With GridRect
For iCol = 0 To ((PropFixedCols + PropFrozenCols) - 1)
    .Right = .Right + GetColWidth(iCol)
Next iCol
For iCol = LeftCol To (PropCols - 1)
    .Right = .Right + GetColWidth(iCol)
    If .Right > VBFlexGridClientRect.Right Then Exit For
Next iCol
If .Right <= VBFlexGridClientRect.Right Then
    Do While LeftCol > (((PropFixedCols + PropFrozenCols) - 1) + 1)
        .Right = .Right + GetColWidth(LeftCol - 1)
        If .Right > VBFlexGridClientRect.Right Then
            Exit Do
        Else
            LeftCol = LeftCol - 1
        End If
    Loop
End If
End With
End Sub

Private Sub SetRowColParams(ByRef RCP As TROWCOLPARAMS)
Dim RowColChanged As Boolean, SelChanged As Boolean, ScrollChanged As Boolean
Dim NoRedraw As Boolean, Cancel As Boolean
With RCP
Select Case PropScrollBars
    Case vbSBNone
        If Not (.Flags And RCPF_FORCETOPROWMASK) = RCPF_FORCETOPROWMASK Then
            If (.Mask And RCPM_TOPROW) = RCPM_TOPROW Then .Mask = .Mask And Not RCPM_TOPROW
        End If
        If Not (.Flags And RCPF_FORCELEFTCOLMASK) = RCPF_FORCELEFTCOLMASK Then
            If (.Mask And RCPM_LEFTCOL) = RCPM_LEFTCOL Then .Mask = .Mask And Not RCPM_LEFTCOL
        End If
    Case vbHorizontal
        If Not (.Flags And RCPF_FORCETOPROWMASK) = RCPF_FORCETOPROWMASK Then
            If (.Mask And RCPM_TOPROW) = RCPM_TOPROW Then .Mask = .Mask And Not RCPM_TOPROW
        End If
    Case vbVertical
        If Not (.Flags And RCPF_FORCELEFTCOLMASK) = RCPF_FORCELEFTCOLMASK Then
            If (.Mask And RCPM_LEFTCOL) = RCPM_LEFTCOL Then .Mask = .Mask And Not RCPM_LEFTCOL
        End If
End Select
If (.Mask And RCPM_ROW) = RCPM_ROW Then
    If .Row > (PropRows - 1) Then .Row = (PropRows - 1)
    If VBFlexGridRow <> .Row Then RowColChanged = True
End If
If (.Mask And RCPM_COL) = RCPM_COL Then
    If .Col > (PropCols - 1) Then .Col = (PropCols - 1)
    If VBFlexGridCol <> .Col Then RowColChanged = True
End If
If RowColChanged = True Then
    Dim NewRow As Long, NewCol As Long
    NewRow = IIf((.Mask And RCPM_ROW) = RCPM_ROW, .Row, VBFlexGridRow)
    NewCol = IIf((.Mask And RCPM_COL) = RCPM_COL, .Col, VBFlexGridCol)
    RaiseEvent BeforeRowColChange(NewRow, NewCol, Cancel)
    If Cancel = True Then RowColChanged = False
    Cancel = False
End If
If RowColChanged = False Then
    If (.Mask And RCPM_ROWSEL) = RCPM_ROWSEL Then
        If PropAllowSelection = True Then
            If VBFlexGridRowSel <> .RowSel Then SelChanged = True
        Else
            If (.Mask And RCPM_ROW) = RCPM_ROW Then
                If VBFlexGridRowSel <> .Row Then SelChanged = True
            Else
                If VBFlexGridRowSel <> VBFlexGridRow Then SelChanged = True
            End If
        End If
    End If
    If (.Mask And RCPM_COLSEL) = RCPM_COLSEL Then
        If PropAllowSelection = True Then
            If VBFlexGridColSel <> .ColSel Then SelChanged = True
        Else
            If (.Mask And RCPM_COL) = RCPM_COL Then
                If VBFlexGridColSel <> .Col Then SelChanged = True
            Else
                If VBFlexGridColSel <> VBFlexGridCol Then SelChanged = True
            End If
        End If
    End If
Else
    SelChanged = True
End If
If SelChanged = True Then
    Dim NewRowSel As Long, NewColSel As Long
    If PropAllowSelection = True Then
        NewRowSel = IIf((.Mask And RCPM_ROWSEL) = RCPM_ROWSEL, .RowSel, VBFlexGridRowSel)
        NewColSel = IIf((.Mask And RCPM_COLSEL) = RCPM_COLSEL, .ColSel, VBFlexGridColSel)
    Else
        NewRowSel = IIf((.Mask And RCPM_ROW) = RCPM_ROW, .Row, VBFlexGridRow)
        NewColSel = IIf((.Mask And RCPM_COL) = RCPM_COL, .Col, VBFlexGridCol)
    End If
    RaiseEvent BeforeSelChange(NewRowSel, NewColSel, Cancel)
    If Cancel = True Then SelChanged = False
    Cancel = False
End If
If (.Mask And RCPM_TOPROW) = RCPM_TOPROW Then
    If .TopRow < (PropFixedRows + PropFrozenRows) Then .TopRow = PropFixedRows + PropFrozenRows
    If (.Flags And RCPF_CHECKTOPROW) = RCPF_CHECKTOPROW Then Call CheckTopRow(.TopRow)
    If VBFlexGridTopRow <> .TopRow Then ScrollChanged = True
End If
If (.Mask And RCPM_LEFTCOL) = RCPM_LEFTCOL Then
    If .LeftCol < (PropFixedCols + PropFrozenCols) Then .LeftCol = PropFixedCols + PropFrozenCols
    If (.Flags And RCPF_CHECKLEFTCOL) = RCPF_CHECKLEFTCOL Then Call CheckLeftCol(.LeftCol)
    If VBFlexGridLeftCol <> .LeftCol Then ScrollChanged = True
End If
If RowColChanged = True Then
    RaiseEvent LeaveCell
    If (.Mask And RCPM_ROW) = RCPM_ROW Then VBFlexGridRow = .Row
    If (.Mask And RCPM_COL) = RCPM_COL Then VBFlexGridCol = .Col
End If
If SelChanged = True Then
    If PropAllowSelection = True Then
        If (.Mask And RCPM_ROWSEL) = RCPM_ROWSEL Then VBFlexGridRowSel = .RowSel
        If (.Mask And RCPM_COLSEL) = RCPM_COLSEL Then VBFlexGridColSel = .ColSel
    Else
        If (.Mask And RCPM_ROWSEL) = RCPM_ROWSEL Then VBFlexGridRowSel = IIf((.Mask And RCPM_ROW) = RCPM_ROW, .Row, VBFlexGridRow)
        If (.Mask And RCPM_COLSEL) = RCPM_COLSEL Then VBFlexGridColSel = IIf((.Mask And RCPM_COL) = RCPM_COL, .Col, VBFlexGridCol)
    End If
End If
If ScrollChanged = True Then
    If (.Mask And RCPM_TOPROW) = RCPM_TOPROW And (.Mask And RCPM_LEFTCOL) = RCPM_LEFTCOL Then
        VBFlexGridTopRow = .TopRow
        VBFlexGridLeftCol = .LeftCol
        NoRedraw = (CheckScrollPos(SB_HORZ) Or CheckScrollPos(SB_VERT))
    ElseIf (.Mask And RCPM_TOPROW) = RCPM_TOPROW Then
        VBFlexGridTopRow = .TopRow
        NoRedraw = CheckScrollPos(SB_VERT)
    ElseIf (.Mask And RCPM_LEFTCOL) = RCPM_LEFTCOL Then
        VBFlexGridLeftCol = .LeftCol
        NoRedraw = CheckScrollPos(SB_HORZ)
    End If
End If
If NoRedraw = False Then
    If RowColChanged = True Or SelChanged = True Or ScrollChanged = True Or (.Flags And RCPF_FORCEREDRAW) = RCPF_FORCEREDRAW Then Call RedrawGrid
End If
If (.Flags And RCPF_SETSCROLLBARS) = RCPF_SETSCROLLBARS Then Call SetScrollBars
If SelChanged = True Then RaiseEvent SelChange
If RowColChanged = True Then
    RaiseEvent EnterCell
    RaiseEvent RowColChange
End If
End With
End Sub

Private Sub MovePreviousRow(ByRef iRow As Long)
Dim i As Long, Cancel As Boolean
i = iRow
Do
    If i > PropFixedRows Then i = i - 1 Else Cancel = True
Loop Until GetRowHeight(i) > 0 Or Cancel = True
If Cancel = False Then iRow = i
End Sub

Private Sub MoveNextRow(ByRef iRow As Long)
Dim i As Long, Cancel As Boolean
i = iRow
Do
    If i < (PropRows - 1) Then i = i + 1 Else Cancel = True
Loop Until GetRowHeight(i) > 0 Or Cancel = True
If Cancel = False Then iRow = i
End Sub

Private Sub MoveFirstRow(ByRef iRow As Long)
Dim i As Long, Cancel As Boolean
i = PropFixedRows
Do Until GetRowHeight(i) > 0 Or Cancel = True
    If i < iRow Then i = i + 1 Else Cancel = True
Loop
If Cancel = False Then iRow = i
End Sub

Private Sub MoveLastRow(ByRef iRow As Long)
Dim i As Long, Cancel As Boolean
i = PropRows - 1
Do Until GetRowHeight(i) > 0 Or Cancel = True
    If i > iRow Then i = i - 1 Else Cancel = True
Loop
If Cancel = False Then iRow = i
End Sub

Private Sub MovePreviousCol(ByRef iCol As Long)
Dim i As Long, Cancel As Boolean
i = iCol
Do
    If i > PropFixedCols Then i = i - 1 Else Cancel = True
Loop Until GetColWidth(i) > 0 Or Cancel = True
If Cancel = False Then iCol = i
End Sub

Private Sub MoveNextCol(ByRef iCol As Long)
Dim i As Long, Cancel As Boolean
i = iCol
Do
    If i < (PropCols - 1) Then i = i + 1 Else Cancel = True
Loop Until GetColWidth(i) > 0 Or Cancel = True
If Cancel = False Then iCol = i
End Sub

Private Sub MoveFirstCol(ByRef iCol As Long)
Dim i As Long, Cancel As Boolean
i = PropFixedCols
Do Until GetColWidth(i) > 0 Or Cancel = True
    If i < iCol Then i = i + 1 Else Cancel = True
Loop
If Cancel = False Then iCol = i
End Sub

Private Sub MoveLastCol(ByRef iCol As Long)
Dim i As Long, Cancel As Boolean
i = PropCols - 1
Do Until GetColWidth(i) > 0 Or Cancel = True
    If i > iCol Then i = i - 1 Else Cancel = True
Loop
If Cancel = False Then iCol = i
End Sub

Private Function GetFirstMovableRow() As Long
Dim i As Long, Cancel As Boolean
i = PropFixedRows
Do Until GetRowHeight(i) > 0 Or Cancel = True
    If i < (PropRows - 1) Then i = i + 1 Else Cancel = True
Loop
If Cancel = False Then GetFirstMovableRow = i
End Function

Private Function GetLastMovableRow() As Long
Dim i As Long, Cancel As Boolean
i = PropRows - 1
Do Until GetRowHeight(i) > 0 Or Cancel = True
    If i > PropFixedRows Then i = i - 1 Else Cancel = True
Loop
If Cancel = False Then GetLastMovableRow = i
End Function

Private Function GetFirstMovableCol() As Long
Dim i As Long, Cancel As Boolean
i = PropFixedCols
Do Until GetColWidth(i) > 0 Or Cancel = True
    If i < (PropCols - 1) Then i = i + 1 Else Cancel = True
Loop
If Cancel = False Then GetFirstMovableCol = i
End Function

Private Function GetLastMovableCol() As Long
Dim i As Long, Cancel As Boolean
i = PropCols - 1
Do Until GetColWidth(i) > 0 Or Cancel = True
    If i > PropFixedCols Then i = i - 1 Else Cancel = True
Loop
If Cancel = False Then GetLastMovableCol = i
End Function

Private Sub ProcessKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
If PropRows < 1 Or PropCols < 1 Then Exit Sub
If VBFlexGridEditHandle <> 0 Then Exit Sub
Select Case KeyCode
    Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyPageUp, vbKeyPageDown, vbKeyHome, vbKeyEnd
    Case vbKeyTab
        If PropTabBehavior = FlexTabControls Then Exit Sub
    Case vbKeyReturn
        If PropDirectionAfterReturn = FlexDirectionAfterReturnNone Then Exit Sub
    Case Else
        Exit Sub
End Select
If VBFlexGridRTLLayout = True Then
    If KeyCode = vbKeyLeft Then
        KeyCode = vbKeyRight
    ElseIf KeyCode = vbKeyRight Then
        KeyCode = vbKeyLeft
    End If
End If
If PropAllowSelection = False Then
    If (Shift And vbShiftMask) = vbShiftMask And KeyCode <> vbKeyTab And KeyCode <> vbKeyReturn Then Exit Sub
End If
Dim RCP As TROWCOLPARAMS, RowsPerPage As Long, ColsPerPage As Long
With RCP
.Mask = RCPM_ROW Or RCPM_COL Or RCPM_ROWSEL Or RCPM_COLSEL Or RCPM_TOPROW Or RCPM_LEFTCOL
.Row = VBFlexGridRow
.Col = VBFlexGridCol
Select Case KeyCode
    Case vbKeyUp, vbKeyPageUp
        If .Row < PropFixedRows Then .Row = PropFixedRows
    Case vbKeyDown, vbKeyPageDown
        If .Row > (PropRows - 1) Then .Row = (PropRows - 1)
    Case vbKeyLeft, vbKeyHome
        If .Col < PropFixedCols Then .Col = PropFixedCols
    Case vbKeyRight, vbKeyEnd
        If .Col > (PropCols - 1) Then .Col = (PropCols - 1)
End Select
.RowSel = VBFlexGridRowSel
.ColSel = VBFlexGridColSel
.TopRow = VBFlexGridTopRow
.LeftCol = VBFlexGridLeftCol
Select Case PropSelectionMode
    Case FlexSelectionModeFree, FlexSelectionModeFreeByRow, FlexSelectionModeFreeByColumn
        Select Case KeyCode
            Case vbKeyUp
                If (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) = 0 Then
                    Call MovePreviousRow(.Row)
                    .RowSel = .Row
                    .ColSel = .Col
                    If .TopRow > .Row Then
                        If .Row >= (PropFixedRows + PropFrozenRows) Then .TopRow = .Row
                    ElseIf .Row > (.TopRow + GetRowsPerPage(.TopRow) - 1) Then
                        .TopRow = .Row - GetRowsPerPageRev(.Row) + 1
                    End If
                    If .LeftCol > .Col Then
                        If .Col >= (PropFixedCols + PropFrozenCols) Then .LeftCol = .Col
                    ElseIf .Col > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .Col - GetColsPerPageRev(.Col) + 1
                    End If
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    Call MovePreviousRow(.RowSel)
                    If .TopRow > .RowSel Then
                        If .RowSel >= (PropFixedRows + PropFrozenRows) Then .TopRow = .RowSel
                    End If
                ElseIf (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                    Call MoveFirstRow(.Row)
                    .RowSel = .Row
                    .ColSel = .Col
                    If .Row >= (PropFixedRows + PropFrozenRows) Then .TopRow = .Row
                    If .LeftCol > .Col Then
                        If .Col >= (PropFixedCols + PropFrozenCols) Then .LeftCol = .Col
                    ElseIf .Col > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .Col - GetColsPerPageRev(.Col) + 1
                    End If
                Else
                    Call MoveFirstRow(.RowSel)
                    If .RowSel >= (PropFixedRows + PropFrozenRows) Then .TopRow = .RowSel
                End If
            Case vbKeyDown
                If (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) = 0 Then
                    Call MoveNextRow(.Row)
                    .RowSel = .Row
                    .ColSel = .Col
                    If .TopRow > .Row Then
                        If .Row >= (PropFixedRows + PropFrozenRows) Then .TopRow = .Row
                    ElseIf .Row > (.TopRow + GetRowsPerPage(.TopRow) - 1) Then
                        .TopRow = .Row - GetRowsPerPageRev(.Row) + 1
                    End If
                    If .LeftCol > .Col Then
                        If .Col >= (PropFixedCols + PropFrozenCols) Then .LeftCol = .Col
                    ElseIf .Col > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .Col - GetColsPerPageRev(.Col) + 1
                    End If
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    Call MoveNextRow(.RowSel)
                    If .TopRow > .RowSel Then
                        If .RowSel >= (PropFixedRows + PropFrozenRows) Then .TopRow = .RowSel
                    ElseIf .RowSel > (.TopRow + GetRowsPerPage(.TopRow) - 1) Then
                        .TopRow = .RowSel - GetRowsPerPageRev(.RowSel) + 1
                    End If
                ElseIf (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                    Call MoveLastRow(.Row)
                    .RowSel = .Row
                    .ColSel = .Col
                    .TopRow = (PropRows - 1) - GetRowsPerPageRev(PropRows - 1) + 1
                    If .LeftCol > .Col Then
                        If .Col >= (PropFixedCols + PropFrozenCols) Then .LeftCol = .Col
                    ElseIf .Col > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .Col - GetColsPerPageRev(.Col) + 1
                    End If
                Else
                    Call MoveLastRow(.RowSel)
                    .TopRow = (PropRows - 1) - GetRowsPerPageRev(PropRows - 1) + 1
                End If
            Case vbKeyLeft
                If (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) = 0 Then
                    If PropWrapCellBehavior <> FlexWrapNone Then
                        If .Col > GetFirstMovableCol() Then
                            Call MovePreviousCol(.Col)
                        Else
                            If .Row > GetFirstMovableRow() Then
                                Call MoveLastCol(.Col)
                                Call MovePreviousRow(.Row)
                            Else
                                If PropWrapCellBehavior = FlexWrapGrid Then
                                    Call MoveLastCol(.Col)
                                    Call MoveLastRow(.Row)
                                End If
                            End If
                        End If
                    Else
                        Call MovePreviousCol(.Col)
                    End If
                    .RowSel = .Row
                    .ColSel = .Col
                    If .TopRow > .Row Then
                        If .Row >= (PropFixedRows + PropFrozenRows) Then .TopRow = .Row
                    ElseIf .Row > (.TopRow + GetRowsPerPage(.TopRow) - 1) Then
                        .TopRow = .Row - GetRowsPerPageRev(.Row) + 1
                    End If
                    If .LeftCol > .Col Then
                        If .Col >= (PropFixedCols + PropFrozenCols) Then .LeftCol = .Col
                    ElseIf .Col > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .Col - GetColsPerPageRev(.Col) + 1
                    End If
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    Call MovePreviousCol(.ColSel)
                    If .LeftCol > .ColSel Then
                        If .ColSel >= (PropFixedCols + PropFrozenCols) Then .LeftCol = .ColSel
                    End If
                ElseIf (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                    Call MoveFirstCol(.Col)
                    .RowSel = .Row
                    .ColSel = .Col
                    If .TopRow > .Row Then
                        If .Row >= (PropFixedRows + PropFrozenRows) Then .TopRow = .Row
                    ElseIf .Row > (.TopRow + GetRowsPerPage(.TopRow) - 1) Then
                        .TopRow = .Row - GetRowsPerPageRev(.Row) + 1
                    End If
                    If .Col >= (PropFixedCols + PropFrozenCols) Then .LeftCol = .Col
                Else
                    Call MoveFirstCol(.ColSel)
                    If .ColSel >= (PropFixedCols + PropFrozenCols) Then .LeftCol = .ColSel
                End If
            Case vbKeyRight
                If (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) = 0 Then
                    If PropWrapCellBehavior <> FlexWrapNone Then
                        If .Col < GetLastMovableCol() Then
                            Call MoveNextCol(.Col)
                        Else
                            If .Row < GetLastMovableRow() Then
                                Call MoveFirstCol(.Col)
                                Call MoveNextRow(.Row)
                            Else
                                If PropWrapCellBehavior = FlexWrapGrid Then
                                    Call MoveFirstCol(.Col)
                                    Call MoveFirstRow(.Row)
                                End If
                            End If
                        End If
                    Else
                        Call MoveNextCol(.Col)
                    End If
                    .RowSel = .Row
                    .ColSel = .Col
                    If .TopRow > .Row Then
                        If .Row >= (PropFixedRows + PropFrozenRows) Then .TopRow = .Row
                    ElseIf .Row > (.TopRow + GetRowsPerPage(.TopRow) - 1) Then
                        .TopRow = .Row - GetRowsPerPageRev(.Row) + 1
                    End If
                    If .LeftCol > .Col Then
                        If .Col >= (PropFixedCols + PropFrozenCols) Then .LeftCol = .Col
                    ElseIf .Col > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .Col - GetColsPerPageRev(.Col) + 1
                    End If
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    Call MoveNextCol(.ColSel)
                    If .LeftCol > .ColSel Then
                        If .ColSel >= (PropFixedCols + PropFrozenCols) Then .LeftCol = .ColSel
                    ElseIf .ColSel > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .ColSel - GetColsPerPageRev(.ColSel) + 1
                    End If
                ElseIf (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                    Call MoveLastCol(.Col)
                    .RowSel = .Row
                    .ColSel = .Col
                    If .TopRow > .Row Then
                        If .Row >= (PropFixedRows + PropFrozenRows) Then .TopRow = .Row
                    ElseIf .Row > (.TopRow + GetRowsPerPage(.TopRow) - 1) Then
                        .TopRow = .Row - GetRowsPerPageRev(.Row) + 1
                    End If
                    .LeftCol = (PropCols - 1) - GetColsPerPageRev(PropCols - 1) + 1
                Else
                    Call MoveLastCol(.ColSel)
                    .LeftCol = (PropCols - 1) - GetColsPerPageRev(PropCols - 1) + 1
                End If
            Case vbKeyPageUp
                If (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) = 0 Then
                    If .Row >= PropFixedRows And .Row < (PropFixedRows + PropFrozenRows) Then
                        RowsPerPage = PropFrozenRows - .Row + 1
                        If (.Row + RowsPerPage) < (PropRows - 1) Then
                            .Row = .Row + RowsPerPage
                        Else
                            .Row = (PropRows - 1)
                        End If
                        If GetRowHeight(.Row) = 0 Then Call MovePreviousRow(.Row)
                    ElseIf .Row > PropFixedRows Then
                        RowsPerPage = GetRowsPerPageRev(.Row)
                        If (.Row - RowsPerPage) > (PropFixedRows + PropFrozenRows) Then
                            .Row = .Row - RowsPerPage
                        Else
                            .Row = PropFixedRows + PropFrozenRows
                        End If
                        If GetRowHeight(.Row) = 0 Then Call MoveNextRow(.Row)
                    End If
                    .RowSel = .Row
                    .ColSel = .Col
                    If .TopRow > .Row Then
                        If .Row >= (PropFixedRows + PropFrozenRows) Then .TopRow = .Row
                    Else
                        .TopRow = .Row - GetRowsPerPageRev(.Row) + 1
                    End If
                    If .LeftCol > .Col Then
                        If .Col >= (PropFixedCols + PropFrozenCols) Then .LeftCol = .Col
                    ElseIf .ColSel > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .ColSel - GetColsPerPageRev(.ColSel) + 1
                    End If
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    If .RowSel >= PropFixedRows And .RowSel < (PropFixedRows + PropFrozenRows) Then
                        RowsPerPage = PropFrozenRows - .RowSel + 1
                        If (.RowSel + RowsPerPage) < (PropRows - 1) Then
                            .RowSel = .RowSel + RowsPerPage
                        Else
                            .RowSel = (PropRows - 1)
                        End If
                        If GetRowHeight(.Row) = 0 Then Call MovePreviousRow(.Row)
                    ElseIf .RowSel > PropFixedRows Then
                        RowsPerPage = GetRowsPerPageRev(.RowSel)
                        If (.RowSel - RowsPerPage) > (PropFixedRows + PropFrozenRows) Then
                            .RowSel = .RowSel - RowsPerPage
                        Else
                            .RowSel = PropFixedRows + PropFrozenRows
                        End If
                        If GetRowHeight(.RowSel) = 0 Then Call MoveNextRow(.RowSel)
                    End If
                    If .TopRow > .RowSel Then
                        If .RowSel >= (PropFixedRows + PropFrozenRows) Then .TopRow = .RowSel
                    Else
                        .TopRow = .RowSel - GetRowsPerPageRev(.RowSel) + 1
                    End If
                ElseIf (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                    Call MoveFirstRow(.Row)
                    .RowSel = .Row
                    .ColSel = .Col
                    If .Row >= (PropFixedRows + PropFrozenRows) Then .TopRow = .Row
                    If .LeftCol > .Col Then
                        If .Col >= (PropFixedCols + PropFrozenCols) Then .LeftCol = .Col
                    ElseIf .Col > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .Col - GetColsPerPageRev(.Col) + 1
                    End If
                Else
                    Call MoveFirstRow(.RowSel)
                    If .RowSel >= (PropFixedRows + PropFrozenRows) Then .TopRow = .RowSel
                End If
            Case vbKeyPageDown
                If (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) = 0 Then
                    If .Row < (PropRows - 1) Then
                        If .Row >= PropFixedRows And .Row < (PropFixedRows + PropFrozenRows) Then
                            RowsPerPage = PropFrozenRows - .Row + 1
                        Else
                            RowsPerPage = GetRowsPerPage(.Row)
                        End If
                        If (.Row + RowsPerPage) < (PropRows - 1) Then
                            .Row = .Row + RowsPerPage
                        Else
                            .Row = (PropRows - 1)
                        End If
                        If GetRowHeight(.Row) = 0 Then Call MovePreviousRow(.Row)
                    End If
                    .RowSel = .Row
                    .ColSel = .Col
                    If .TopRow > .Row Then
                        If .Row >= (PropFixedRows + PropFrozenRows) Then .TopRow = .Row
                    Else
                        .TopRow = .Row - GetRowsPerPageRev(.Row) + 1
                    End If
                    If .LeftCol > .Col Then
                        If .Col >= (PropFixedCols + PropFrozenCols) Then .LeftCol = .Col
                    ElseIf .ColSel > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .ColSel - GetColsPerPageRev(.ColSel) + 1
                    End If
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    If .RowSel < (PropRows - 1) Then
                        If .RowSel >= PropFixedRows And .RowSel < (PropFixedRows + PropFrozenRows) Then
                            RowsPerPage = PropFrozenRows - .RowSel + 1
                        Else
                            RowsPerPage = GetRowsPerPage(.RowSel)
                        End If
                        If (.RowSel + RowsPerPage) < (PropRows - 1) Then
                            .RowSel = .RowSel + RowsPerPage
                        Else
                            .RowSel = (PropRows - 1)
                        End If
                        If GetRowHeight(.RowSel) = 0 Then Call MovePreviousRow(.RowSel)
                    End If
                    If .TopRow > .RowSel Then
                        If .RowSel >= (PropFixedRows + PropFrozenRows) Then .TopRow = .RowSel
                    Else
                        .TopRow = .RowSel - GetRowsPerPageRev(.RowSel) + 1
                    End If
                ElseIf (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                    Call MoveLastRow(.Row)
                    .RowSel = .Row
                    .ColSel = .Col
                    .TopRow = (PropRows - 1) - GetRowsPerPageRev(PropRows - 1) + 1
                    If .LeftCol > .Col Then
                        If .Col >= (PropFixedCols + PropFrozenCols) Then .LeftCol = .Col
                    ElseIf .Col > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .Col - GetColsPerPageRev(.Col) + 1
                    End If
                Else
                    Call MoveLastRow(.RowSel)
                    .TopRow = (PropRows - 1) - GetRowsPerPageRev(PropRows - 1) + 1
                End If
            Case vbKeyHome
                If (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) = 0 Then
                    Call MoveFirstCol(.Col)
                    .RowSel = .Row
                    .ColSel = .Col
                    If .TopRow > .Row Then
                        If .Row >= (PropFixedRows + PropFrozenRows) Then .TopRow = .Row
                    ElseIf .Row > (.TopRow + GetRowsPerPage(.TopRow) - 1) Then
                        .TopRow = .Row - GetRowsPerPageRev(.Row) + 1
                    End If
                    If .Col >= (PropFixedCols + PropFrozenCols) Then .LeftCol = .Col
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    Call MoveFirstCol(.ColSel)
                    If .ColSel >= (PropFixedCols + PropFrozenCols) Then .LeftCol = .ColSel
                ElseIf (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                    Call MoveFirstRow(.Row)
                    Call MoveFirstCol(.Col)
                    .RowSel = .Row
                    .ColSel = .Col
                    If .Row >= (PropFixedRows + PropFrozenRows) Then .TopRow = .Row
                    If .Col >= (PropFixedCols + PropFrozenCols) Then .LeftCol = .Col
                Else
                    Call MoveFirstRow(.RowSel)
                    Call MoveFirstCol(.ColSel)
                    If .RowSel >= (PropFixedRows + PropFrozenRows) Then .TopRow = .RowSel
                    If .ColSel >= (PropFixedCols + PropFrozenCols) Then .LeftCol = .ColSel
                End If
            Case vbKeyEnd
                If (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) = 0 Then
                    Call MoveLastCol(.Col)
                    .RowSel = .Row
                    .ColSel = .Col
                    If .TopRow > .Row Then
                        If .Row >= (PropFixedRows + PropFrozenRows) Then .TopRow = .Row
                    ElseIf .Row > (.TopRow + GetRowsPerPage(.TopRow) - 1) Then
                        .TopRow = .Row - GetRowsPerPageRev(.Row) + 1
                    End If
                    .LeftCol = (PropCols - 1) - GetColsPerPageRev(PropCols - 1) + 1
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    Call MoveLastCol(.ColSel)
                    .LeftCol = (PropCols - 1) - GetColsPerPageRev(PropCols - 1) + 1
                ElseIf (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                    Call MoveLastRow(.Row)
                    Call MoveLastCol(.Col)
                    .RowSel = .Row
                    .ColSel = .Col
                    .TopRow = (PropRows - 1) - GetRowsPerPageRev(PropRows - 1) + 1
                    .LeftCol = (PropCols - 1) - GetColsPerPageRev(PropCols - 1) + 1
                Else
                    Call MoveLastRow(.RowSel)
                    Call MoveLastCol(.ColSel)
                    .TopRow = (PropRows - 1) - GetRowsPerPageRev(PropRows - 1) + 1
                    .LeftCol = (PropCols - 1) - GetColsPerPageRev(PropCols - 1) + 1
                End If
            Case vbKeyTab
                If (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) = 0 Then
                    If PropWrapCellBehavior <> FlexWrapNone Then
                        If .Col < GetLastMovableCol() Then
                            Call MoveNextCol(.Col)
                        Else
                            If .Row < GetLastMovableRow() Then
                                Call MoveFirstCol(.Col)
                                Call MoveNextRow(.Row)
                            Else
                                If PropWrapCellBehavior = FlexWrapGrid Then
                                    Call MoveFirstCol(.Col)
                                    Call MoveFirstRow(.Row)
                                End If
                            End If
                        End If
                    Else
                        Call MoveNextCol(.Col)
                    End If
                    .RowSel = .Row
                    .ColSel = .Col
                    If .TopRow > .Row Then
                        If .Row >= (PropFixedRows + PropFrozenRows) Then .TopRow = .Row
                    ElseIf .Row > (.TopRow + GetRowsPerPage(.TopRow) - 1) Then
                        .TopRow = .Row - GetRowsPerPageRev(.Row) + 1
                    End If
                    If .LeftCol > .Col Then
                        If .Col >= (PropFixedCols + PropFrozenCols) Then .LeftCol = .Col
                    ElseIf .Col > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .Col - GetColsPerPageRev(.Col) + 1
                    End If
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    If PropWrapCellBehavior <> FlexWrapNone Then
                        If .Col > GetFirstMovableCol() Then
                            Call MovePreviousCol(.Col)
                        Else
                            If .Row > GetFirstMovableRow() Then
                                Call MoveLastCol(.Col)
                                Call MovePreviousRow(.Row)
                            Else
                                If PropWrapCellBehavior = FlexWrapGrid Then
                                    Call MoveLastCol(.Col)
                                    Call MoveLastRow(.Row)
                                End If
                            End If
                        End If
                    Else
                        Call MovePreviousCol(.Col)
                    End If
                    .RowSel = .Row
                    .ColSel = .Col
                    If .TopRow > .Row Then
                        If .Row >= (PropFixedRows + PropFrozenRows) Then .TopRow = .Row
                    ElseIf .Row > (.TopRow + GetRowsPerPage(.TopRow) - 1) Then
                        .TopRow = .Row - GetRowsPerPageRev(.Row) + 1
                    End If
                    If .LeftCol > .Col Then
                        If .Col >= (PropFixedCols + PropFrozenCols) Then .LeftCol = .Col
                    ElseIf .Col > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .Col - GetColsPerPageRev(.Col) + 1
                    End If
                ElseIf (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                    ' Void
                Else
                    ' Void
                End If
            Case vbKeyReturn
                If (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) = 0 Then
                    Select Case PropDirectionAfterReturn
                        Case FlexDirectionAfterReturnUp
                            If .Row > GetFirstMovableRow() Then
                                Call MovePreviousRow(.Row)
                            Else
                                If PropWrapCellBehavior = FlexWrapGrid Then Call MoveLastRow(.Row)
                            End If
                        Case FlexDirectionAfterReturnDown
                            If .Row < GetLastMovableRow() Then
                                Call MoveNextRow(.Row)
                            Else
                                If PropWrapCellBehavior = FlexWrapGrid Then Call MoveFirstRow(.Row)
                            End If
                        Case FlexDirectionAfterReturnLeft
                            If PropWrapCellBehavior <> FlexWrapNone Then
                                If .Col > GetFirstMovableCol() Then
                                    Call MovePreviousCol(.Col)
                                Else
                                    If .Row > GetFirstMovableRow() Then
                                        Call MoveLastCol(.Col)
                                        Call MovePreviousRow(.Row)
                                    Else
                                        If PropWrapCellBehavior = FlexWrapGrid Then
                                            Call MoveLastCol(.Col)
                                            Call MoveLastRow(.Row)
                                        End If
                                    End If
                                End If
                            Else
                                Call MovePreviousCol(.Col)
                            End If
                        Case FlexDirectionAfterReturnRight
                            If PropWrapCellBehavior <> FlexWrapNone Then
                                If .Col < GetLastMovableCol() Then
                                    Call MoveNextCol(.Col)
                                Else
                                    If .Row < GetLastMovableRow() Then
                                        Call MoveFirstCol(.Col)
                                        Call MoveNextRow(.Row)
                                    Else
                                        If PropWrapCellBehavior = FlexWrapGrid Then
                                            Call MoveFirstCol(.Col)
                                            Call MoveFirstRow(.Row)
                                        End If
                                    End If
                                End If
                            Else
                                Call MoveNextCol(.Col)
                            End If
                    End Select
                    .RowSel = .Row
                    .ColSel = .Col
                    If .TopRow > .Row Then
                        If .Row >= (PropFixedRows + PropFrozenRows) Then .TopRow = .Row
                    ElseIf .Row > (.TopRow + GetRowsPerPage(.TopRow) - 1) Then
                        .TopRow = .Row - GetRowsPerPageRev(.Row) + 1
                    End If
                    If .LeftCol > .Col Then
                        If .Col >= (PropFixedCols + PropFrozenCols) Then .LeftCol = .Col
                    ElseIf .Col > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .Col - GetColsPerPageRev(.Col) + 1
                    End If
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    Select Case PropDirectionAfterReturn
                        Case FlexDirectionAfterReturnUp
                            If .Row < GetLastMovableRow() Then
                                Call MoveNextRow(.Row)
                            Else
                                If PropWrapCellBehavior = FlexWrapGrid Then Call MoveFirstRow(.Row)
                            End If
                        Case FlexDirectionAfterReturnDown
                            If .Row > GetFirstMovableRow() Then
                                Call MovePreviousRow(.Row)
                            Else
                                If PropWrapCellBehavior = FlexWrapGrid Then Call MoveLastRow(.Row)
                            End If
                        Case FlexDirectionAfterReturnLeft
                            If PropWrapCellBehavior <> FlexWrapNone Then
                                If .Col < GetLastMovableCol() Then
                                    Call MoveNextCol(.Col)
                                Else
                                    If .Row < GetLastMovableRow() Then
                                        Call MoveFirstCol(.Col)
                                        Call MoveNextRow(.Row)
                                    Else
                                        If PropWrapCellBehavior = FlexWrapGrid Then
                                            Call MoveFirstCol(.Col)
                                            Call MoveFirstRow(.Row)
                                        End If
                                    End If
                                End If
                            Else
                                Call MoveNextCol(.Col)
                            End If
                        Case FlexDirectionAfterReturnRight
                            If PropWrapCellBehavior <> FlexWrapNone Then
                                If .Col > GetFirstMovableCol() Then
                                    Call MovePreviousCol(.Col)
                                Else
                                    If .Row > GetFirstMovableRow() Then
                                        Call MoveLastCol(.Col)
                                        Call MovePreviousRow(.Row)
                                    Else
                                        If PropWrapCellBehavior = FlexWrapGrid Then
                                            Call MoveLastCol(.Col)
                                            Call MoveLastRow(.Row)
                                        End If
                                    End If
                                End If
                            Else
                                Call MovePreviousCol(.Col)
                            End If
                    End Select
                    .RowSel = .Row
                    .ColSel = .Col
                    If .TopRow > .Row Then
                        If .Row >= (PropFixedRows + PropFrozenRows) Then .TopRow = .Row
                    ElseIf .Row > (.TopRow + GetRowsPerPage(.TopRow) - 1) Then
                        .TopRow = .Row - GetRowsPerPageRev(.Row) + 1
                    End If
                    If .LeftCol > .Col Then
                        If .Col >= (PropFixedCols + PropFrozenCols) Then .LeftCol = .Col
                    ElseIf .Col > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .Col - GetColsPerPageRev(.Col) + 1
                    End If
                ElseIf (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                    ' Void
                Else
                    ' Void
                End If
        End Select
    Case FlexSelectionModeByRow
        Select Case KeyCode
            Case vbKeyUp
                If (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) = 0 Then
                    If PropWrapCellBehavior = FlexWrapGrid Then
                        If .Row > GetFirstMovableRow() Then
                            Call MovePreviousRow(.Row)
                        Else
                            Call MoveLastRow(.Row)
                        End If
                    Else
                        Call MovePreviousRow(.Row)
                    End If
                    .RowSel = .Row
                    .ColSel = (PropCols - 1)
                    If .TopRow > .Row Then
                        If .Row >= (PropFixedRows + PropFrozenRows) Then .TopRow = .Row
                    ElseIf .Row > (.TopRow + GetRowsPerPage(.TopRow) - 1) Then
                        .TopRow = .Row - GetRowsPerPageRev(.Row) + 1
                    End If
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    Call MovePreviousRow(.RowSel)
                    If .TopRow > .RowSel Then
                        If .RowSel >= (PropFixedRows + PropFrozenRows) Then .TopRow = .RowSel
                    End If
                ElseIf (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                    Call MoveFirstRow(.Row)
                    .RowSel = .Row
                    .ColSel = (PropCols - 1)
                    If .Row >= (PropFixedRows + PropFrozenRows) Then .TopRow = .Row
                Else
                    Call MoveFirstRow(.RowSel)
                    If .RowSel >= (PropFixedRows + PropFrozenRows) Then .TopRow = .RowSel
                End If
            Case vbKeyDown
                If (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) = 0 Then
                    If PropWrapCellBehavior = FlexWrapGrid Then
                        If .Row < GetLastMovableRow() Then
                            Call MoveNextRow(.Row)
                        Else
                            Call MoveFirstRow(.Row)
                        End If
                    Else
                        Call MoveNextRow(.Row)
                    End If
                    .RowSel = .Row
                    .ColSel = (PropCols - 1)
                    If .TopRow > .Row Then
                        If .Row >= (PropFixedRows + PropFrozenRows) Then .TopRow = .Row
                    ElseIf .Row > (.TopRow + GetRowsPerPage(.TopRow) - 1) Then
                        .TopRow = .Row - GetRowsPerPageRev(.Row) + 1
                    End If
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    Call MoveNextRow(.RowSel)
                    If .TopRow > .RowSel Then
                        If .RowSel >= (PropFixedRows + PropFrozenRows) Then .TopRow = .RowSel
                    ElseIf .RowSel > (.TopRow + GetRowsPerPage(.TopRow) - 1) Then
                        .TopRow = .RowSel - GetRowsPerPageRev(.RowSel) + 1
                    End If
                ElseIf (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                    Call MoveLastRow(.Row)
                    .RowSel = .Row
                    .ColSel = (PropCols - 1)
                    .TopRow = (PropRows - 1) - GetRowsPerPageRev(PropRows - 1) + 1
                Else
                    Call MoveLastRow(.RowSel)
                    .TopRow = (PropRows - 1) - GetRowsPerPageRev(PropRows - 1) + 1
                End If
            Case vbKeyLeft
                If (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) = 0 Then
                    .RowSel = .Row
                    .ColSel = (PropCols - 1)
                    If .TopRow > .Row Then
                        If .Row >= (PropFixedRows + PropFrozenRows) Then .TopRow = .Row
                    ElseIf .Row > (.TopRow + GetRowsPerPage(.TopRow) - 1) Then
                        .TopRow = .Row - GetRowsPerPageRev(.Row) + 1
                    End If
                    Call MovePreviousCol(.LeftCol)
                    If .LeftCol < (PropFixedCols + PropFrozenCols) Then .LeftCol = (PropFixedCols + PropFrozenCols)
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    Call MovePreviousCol(.LeftCol)
                    If .LeftCol < (PropFixedCols + PropFrozenCols) Then .LeftCol = (PropFixedCols + PropFrozenCols)
                ElseIf (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                    .RowSel = .Row
                    .ColSel = (PropCols - 1)
                    If .TopRow > .Row Then
                        If .Row >= (PropFixedRows + PropFrozenRows) Then .TopRow = .Row
                    ElseIf .Row > (.TopRow + GetRowsPerPage(.TopRow) - 1) Then
                        .TopRow = .Row - GetRowsPerPageRev(.Row) + 1
                    End If
                    Call MoveFirstCol(.LeftCol)
                    If .LeftCol < (PropFixedCols + PropFrozenCols) Then .LeftCol = (PropFixedCols + PropFrozenCols)
                Else
                    Call MoveFirstCol(.LeftCol)
                    If .LeftCol < (PropFixedCols + PropFrozenCols) Then .LeftCol = (PropFixedCols + PropFrozenCols)
                End If
            Case vbKeyRight
                If (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) = 0 Then
                    .RowSel = .Row
                    .ColSel = (PropCols - 1)
                    If .TopRow > .Row Then
                        If .Row >= (PropFixedRows + PropFrozenRows) Then .TopRow = .Row
                    ElseIf .Row > (.TopRow + GetRowsPerPage(.TopRow) - 1) Then
                        .TopRow = .Row - GetRowsPerPageRev(.Row) + 1
                    End If
                    If .LeftCol < (PropCols - 1) - GetColsPerPageRev(PropCols - 1) + 1 Then Call MoveNextCol(.LeftCol)
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    If .LeftCol < (PropCols - 1) - GetColsPerPageRev(PropCols - 1) + 1 Then Call MoveNextCol(.LeftCol)
                ElseIf (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                    .RowSel = .Row
                    .ColSel = (PropCols - 1)
                    If .TopRow > .Row Then
                        If .Row >= (PropFixedRows + PropFrozenRows) Then .TopRow = .Row
                    ElseIf .Row > (.TopRow + GetRowsPerPage(.TopRow) - 1) Then
                        .TopRow = .Row - GetRowsPerPageRev(.Row) + 1
                    End If
                    .LeftCol = (PropCols - 1) - GetColsPerPageRev(PropCols - 1) + 1
                Else
                    .LeftCol = (PropCols - 1) - GetColsPerPageRev(PropCols - 1) + 1
                End If
            Case vbKeyPageUp
                If (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) = 0 Then
                    If .Row >= PropFixedRows And .Row < (PropFixedRows + PropFrozenRows) Then
                        RowsPerPage = PropFrozenRows - .Row + 1
                        If (.Row + RowsPerPage) < (PropRows - 1) Then
                            .Row = .Row + RowsPerPage
                        Else
                            .Row = (PropRows - 1)
                        End If
                        If GetRowHeight(.Row) = 0 Then Call MovePreviousRow(.Row)
                    ElseIf .Row > PropFixedRows Then
                        RowsPerPage = GetRowsPerPageRev(.Row)
                        If (.Row - RowsPerPage) > (PropFixedRows + PropFrozenRows) Then
                            .Row = .Row - RowsPerPage
                        Else
                            .Row = PropFixedRows + PropFrozenRows
                        End If
                        If GetRowHeight(.Row) = 0 Then Call MoveNextRow(.Row)
                    End If
                    .RowSel = .Row
                    .ColSel = (PropCols - 1)
                    If .TopRow > .Row Then
                        If .Row >= (PropFixedRows + PropFrozenRows) Then .TopRow = .Row
                    Else
                        .TopRow = .Row - GetRowsPerPageRev(.Row) + 1
                    End If
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    If .RowSel >= PropFixedRows And .RowSel < (PropFixedRows + PropFrozenRows) Then
                        RowsPerPage = PropFrozenRows - .RowSel + 1
                        If (.RowSel + RowsPerPage) < (PropRows - 1) Then
                            .RowSel = .RowSel + RowsPerPage
                        Else
                            .RowSel = (PropRows - 1)
                        End If
                        If GetRowHeight(.Row) = 0 Then Call MovePreviousRow(.Row)
                    ElseIf .RowSel > PropFixedRows Then
                        RowsPerPage = GetRowsPerPageRev(.RowSel)
                        If (.RowSel - RowsPerPage) > (PropFixedRows + PropFrozenRows) Then
                            .RowSel = .RowSel - RowsPerPage
                        Else
                            .RowSel = PropFixedRows + PropFrozenRows
                        End If
                        If GetRowHeight(.RowSel) = 0 Then Call MoveNextRow(.RowSel)
                    End If
                    If .TopRow > .RowSel Then
                        If .RowSel >= (PropFixedRows + PropFrozenRows) Then .TopRow = .RowSel
                    Else
                        .TopRow = .RowSel - GetRowsPerPageRev(.RowSel) + 1
                    End If
                ElseIf (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                    Call MoveFirstRow(.Row)
                    .RowSel = .Row
                    .ColSel = (PropCols - 1)
                    If .Row >= (PropFixedRows + PropFrozenRows) Then .TopRow = .Row
                Else
                    Call MoveFirstRow(.RowSel)
                    If .RowSel >= (PropFixedRows + PropFrozenRows) Then .TopRow = .RowSel
                End If
            Case vbKeyPageDown
                If (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) = 0 Then
                    If .Row < (PropRows - 1) Then
                        If .Row >= PropFixedRows And .Row < (PropFixedRows + PropFrozenRows) Then
                            RowsPerPage = PropFrozenRows - .Row + 1
                        Else
                            RowsPerPage = GetRowsPerPage(.Row)
                        End If
                        If (.Row + RowsPerPage) < (PropRows - 1) Then
                            .Row = .Row + RowsPerPage
                        Else
                            .Row = (PropRows - 1)
                        End If
                        If GetRowHeight(.Row) = 0 Then Call MovePreviousRow(.Row)
                    End If
                    .RowSel = .Row
                    .ColSel = (PropCols - 1)
                    If .TopRow > .Row Then
                        If .Row >= (PropFixedRows + PropFrozenRows) Then .TopRow = .Row
                    Else
                        .TopRow = .Row - GetRowsPerPageRev(.Row) + 1
                    End If
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    If .RowSel < (PropRows - 1) Then
                        If .RowSel >= PropFixedRows And .RowSel < (PropFixedRows + PropFrozenRows) Then
                            RowsPerPage = PropFrozenRows - .RowSel + 1
                        Else
                            RowsPerPage = GetRowsPerPage(.RowSel)
                        End If
                        If (.RowSel + RowsPerPage) < (PropRows - 1) Then
                            .RowSel = .RowSel + RowsPerPage
                        Else
                            .RowSel = (PropRows - 1)
                        End If
                        If GetRowHeight(.RowSel) = 0 Then Call MovePreviousRow(.RowSel)
                    End If
                    If .TopRow > .RowSel Then
                        If .RowSel >= (PropFixedRows + PropFrozenRows) Then .TopRow = .RowSel
                    Else
                        .TopRow = .RowSel - GetRowsPerPageRev(.RowSel) + 1
                    End If
                ElseIf (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                    Call MoveLastRow(.Row)
                    .RowSel = .Row
                    .ColSel = (PropCols - 1)
                    .TopRow = (PropRows - 1) - GetRowsPerPageRev(PropRows - 1) + 1
                Else
                    Call MoveLastRow(.RowSel)
                    .TopRow = (PropRows - 1) - GetRowsPerPageRev(PropRows - 1) + 1
                End If
            Case vbKeyHome
                If (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) = 0 Then
                    .RowSel = .Row
                    .ColSel = (PropCols - 1)
                    If .TopRow > .Row Then
                        If .Row >= (PropFixedRows + PropFrozenRows) Then .TopRow = .Row
                    ElseIf .Row > (.TopRow + GetRowsPerPage(.TopRow) - 1) Then
                        .TopRow = .Row - GetRowsPerPageRev(.Row) + 1
                    End If
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    ' Void
                ElseIf (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                    Call MoveFirstRow(.Row)
                    .Col = PropFixedCols
                    .RowSel = .Row
                    .ColSel = (PropCols - 1)
                    .TopRow = .Row
                Else
                    Call MoveFirstRow(.RowSel)
                    .ColSel = (PropCols - 1)
                    .TopRow = .RowSel
                End If
            Case vbKeyEnd
                If (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) = 0 Then
                    .RowSel = .Row
                    .ColSel = (PropCols - 1)
                    If .TopRow > .Row Then
                        If .Row >= (PropFixedRows + PropFrozenRows) Then .TopRow = .Row
                    ElseIf .Row > (.TopRow + GetRowsPerPage(.TopRow) - 1) Then
                        .TopRow = .Row - GetRowsPerPageRev(.Row) + 1
                    End If
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    ' Void
                ElseIf (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                    Call MoveLastRow(.Row)
                    .Col = PropFixedCols
                    .RowSel = .Row
                    .ColSel = (PropCols - 1)
                    .TopRow = (PropRows - 1) - GetRowsPerPageRev(PropRows - 1) + 1
                Else
                    Call MoveLastRow(.RowSel)
                    .ColSel = (PropCols - 1)
                    .TopRow = (PropRows - 1) - GetRowsPerPageRev(PropRows - 1) + 1
                End If
            Case vbKeyTab
                If (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) = 0 Then
                    If PropWrapCellBehavior <> FlexWrapNone Then
                        If .Row < GetLastMovableRow() Then
                            Call MoveNextRow(.Row)
                        Else
                            If PropWrapCellBehavior = FlexWrapGrid Then Call MoveFirstRow(.Row)
                        End If
                    End If
                    .RowSel = .Row
                    .ColSel = (PropCols - 1)
                    If .TopRow > .Row Then
                        If .Row >= (PropFixedRows + PropFrozenRows) Then .TopRow = .Row
                    ElseIf .Row > (.TopRow + GetRowsPerPage(.TopRow) - 1) Then
                        .TopRow = .Row - GetRowsPerPageRev(.Row) + 1
                    End If
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    If PropWrapCellBehavior <> FlexWrapNone Then
                        If .Row > GetFirstMovableRow() Then
                            Call MovePreviousRow(.Row)
                        Else
                            If PropWrapCellBehavior = FlexWrapGrid Then Call MoveLastRow(.Row)
                        End If
                    End If
                    .RowSel = .Row
                    .ColSel = (PropCols - 1)
                    If .TopRow > .Row Then
                        If .Row >= (PropFixedRows + PropFrozenRows) Then .TopRow = .Row
                    ElseIf .Row > (.TopRow + GetRowsPerPage(.TopRow) - 1) Then
                        .TopRow = .Row - GetRowsPerPageRev(.Row) + 1
                    End If
                ElseIf (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                    ' Void
                Else
                    ' Void
                End If
            Case vbKeyReturn
                If (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) = 0 Then
                    Select Case PropDirectionAfterReturn
                        Case FlexDirectionAfterReturnUp
                            If .Row > GetFirstMovableRow() Then
                                Call MovePreviousRow(.Row)
                            Else
                                If PropWrapCellBehavior = FlexWrapGrid Then Call MoveLastRow(.Row)
                            End If
                        Case FlexDirectionAfterReturnDown
                            If .Row < GetLastMovableRow() Then
                                Call MoveNextRow(.Row)
                            Else
                                If PropWrapCellBehavior = FlexWrapGrid Then Call MoveFirstRow(.Row)
                            End If
                        Case FlexDirectionAfterReturnLeft
                            If PropWrapCellBehavior <> FlexWrapNone Then
                                If .Row > GetFirstMovableRow() Then
                                    Call MovePreviousRow(.Row)
                                Else
                                    If PropWrapCellBehavior = FlexWrapGrid Then Call MoveLastRow(.Row)
                                End If
                            End If
                        Case FlexDirectionAfterReturnRight
                            If PropWrapCellBehavior <> FlexWrapNone Then
                                If .Row < GetLastMovableRow() Then
                                    Call MoveNextRow(.Row)
                                Else
                                    If PropWrapCellBehavior = FlexWrapGrid Then Call MoveFirstRow(.Row)
                                End If
                            End If
                    End Select
                    .RowSel = .Row
                    .ColSel = (PropCols - 1)
                    If .TopRow > .Row Then
                        If .Row >= (PropFixedRows + PropFrozenRows) Then .TopRow = .Row
                    ElseIf .Row > (.TopRow + GetRowsPerPage(.TopRow) - 1) Then
                        .TopRow = .Row - GetRowsPerPageRev(.Row) + 1
                    End If
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    Select Case PropDirectionAfterReturn
                        Case FlexDirectionAfterReturnUp
                            If .Row < GetLastMovableRow() Then
                                Call MoveNextRow(.Row)
                            Else
                                If PropWrapCellBehavior = FlexWrapGrid Then Call MoveFirstRow(.Row)
                            End If
                        Case FlexDirectionAfterReturnDown
                            If .Row > GetFirstMovableRow() Then
                                Call MovePreviousRow(.Row)
                            Else
                                If PropWrapCellBehavior = FlexWrapGrid Then Call MoveLastRow(.Row)
                            End If
                        Case FlexDirectionAfterReturnLeft
                            If PropWrapCellBehavior <> FlexWrapNone Then
                                If .Row < GetLastMovableRow() Then
                                    Call MoveNextRow(.Row)
                                Else
                                    If PropWrapCellBehavior = FlexWrapGrid Then Call MoveFirstRow(.Row)
                                End If
                            End If
                        Case FlexDirectionAfterReturnRight
                            If PropWrapCellBehavior <> FlexWrapNone Then
                                If .Row > GetFirstMovableRow() Then
                                    Call MovePreviousRow(.Row)
                                Else
                                    If PropWrapCellBehavior = FlexWrapGrid Then Call MoveLastRow(.Row)
                                End If
                            End If
                    End Select
                    .RowSel = .Row
                    .ColSel = (PropCols - 1)
                    If .TopRow > .Row Then
                        If .Row >= (PropFixedRows + PropFrozenRows) Then .TopRow = .Row
                    ElseIf .Row > (.TopRow + GetRowsPerPage(.TopRow) - 1) Then
                        .TopRow = .Row - GetRowsPerPageRev(.Row) + 1
                    End If
                ElseIf (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                    ' Void
                Else
                    ' Void
                End If
        End Select
    Case FlexSelectionModeByColumn
        Select Case KeyCode
            Case vbKeyUp
                If (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) = 0 Then
                    .RowSel = (PropRows - 1)
                    .ColSel = .Col
                    Call MovePreviousRow(.TopRow)
                    If .TopRow < (PropFixedRows + PropFrozenRows) Then .TopRow = (PropFixedRows + PropFrozenRows)
                    If .LeftCol > .Col Then
                        If .Col >= (PropFixedCols + PropFrozenCols) Then .LeftCol = .Col
                    ElseIf .Col > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .Col - GetColsPerPageRev(.Col) + 1
                    End If
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    Call MovePreviousRow(.TopRow)
                    If .TopRow < (PropFixedRows + PropFrozenRows) Then .TopRow = (PropFixedRows + PropFrozenRows)
                ElseIf (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                    .RowSel = (PropRows - 1)
                    .ColSel = .Col
                    Call MoveFirstRow(.TopRow)
                    If .TopRow < (PropFixedRows + PropFrozenRows) Then .TopRow = (PropFixedRows + PropFrozenRows)
                    If .LeftCol > .Col Then
                        If .Col >= (PropFixedCols + PropFrozenCols) Then .LeftCol = .Col
                    ElseIf .Col > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .Col - GetColsPerPageRev(.Col) + 1
                    End If
                Else
                    Call MoveFirstRow(.TopRow)
                    If .TopRow < (PropFixedRows + PropFrozenRows) Then .TopRow = (PropFixedRows + PropFrozenRows)
                End If
            Case vbKeyDown
                If (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) = 0 Then
                    .RowSel = (PropRows - 1)
                    .ColSel = .Col
                    If .TopRow < (PropRows - 1) - GetRowsPerPageRev(PropRows - 1) + 1 Then Call MoveNextRow(.TopRow)
                    If .LeftCol > .Col Then
                        If .Col >= (PropFixedCols + PropFrozenCols) Then .LeftCol = .Col
                    ElseIf .Col > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .Col - GetColsPerPageRev(.Col) + 1
                    End If
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    If .TopRow < (PropRows - 1) - GetRowsPerPageRev(PropRows - 1) + 1 Then Call MoveNextRow(.TopRow)
                ElseIf (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                    .RowSel = (PropRows - 1)
                    .ColSel = .Col
                    .TopRow = (PropRows - 1) - GetRowsPerPageRev(PropRows - 1) + 1
                    If .LeftCol > .Col Then
                        If .Col >= (PropFixedCols + PropFrozenCols) Then .LeftCol = .Col
                    ElseIf .Col > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .Col - GetColsPerPageRev(.Col) + 1
                    End If
                Else
                    .TopRow = (PropRows - 1) - GetRowsPerPageRev(PropRows - 1) + 1
                End If
            Case vbKeyLeft
                If (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) = 0 Then
                    If PropWrapCellBehavior = FlexWrapGrid Then
                        If .Col > GetFirstMovableCol() Then
                            Call MovePreviousCol(.Col)
                        Else
                            Call MoveLastCol(.Col)
                        End If
                    Else
                        Call MovePreviousCol(.Col)
                    End If
                    .RowSel = (PropRows - 1)
                    .ColSel = .Col
                    If .LeftCol > .Col Then
                        If .Col >= (PropFixedCols + PropFrozenCols) Then .LeftCol = .Col
                    ElseIf .Col > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .Col - GetColsPerPageRev(.Col) + 1
                    End If
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    Call MovePreviousCol(.ColSel)
                    If .LeftCol > .ColSel Then
                        If .ColSel >= (PropFixedCols + PropFrozenCols) Then .LeftCol = .ColSel
                    End If
                ElseIf (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                    Call MoveFirstCol(.Col)
                    .RowSel = (PropRows - 1)
                    .ColSel = .Col
                    If .Col >= (PropFixedCols + PropFrozenCols) Then .LeftCol = .Col
                Else
                    Call MoveFirstCol(.ColSel)
                    If .ColSel >= (PropFixedCols + PropFrozenCols) Then .LeftCol = .ColSel
                End If
            Case vbKeyRight
                If (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) = 0 Then
                    If PropWrapCellBehavior = FlexWrapGrid Then
                        If .Col < GetLastMovableCol() Then
                            Call MoveNextCol(.Col)
                        Else
                            Call MoveFirstCol(.Col)
                        End If
                    Else
                        Call MoveNextCol(.Col)
                    End If
                    .RowSel = (PropRows - 1)
                    .ColSel = .Col
                    If .LeftCol > .Col Then
                        If .Col >= (PropFixedCols + PropFrozenCols) Then .LeftCol = .Col
                    ElseIf .Col > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .Col - GetColsPerPageRev(.Col) + 1
                    End If
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    Call MoveNextCol(.ColSel)
                    If .LeftCol > .ColSel Then
                        If .ColSel >= (PropFixedCols + PropFrozenCols) Then .LeftCol = .ColSel
                    ElseIf .ColSel > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .ColSel - GetColsPerPageRev(.ColSel) + 1
                    End If
                ElseIf (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                    Call MoveLastCol(.Col)
                    .RowSel = (PropRows - 1)
                    .ColSel = .Col
                    .LeftCol = (PropCols - 1) - GetColsPerPageRev(PropCols - 1) + 1
                Else
                    Call MoveLastCol(.ColSel)
                    .LeftCol = (PropCols - 1) - GetColsPerPageRev(PropCols - 1) + 1
                End If
            Case vbKeyPageUp
                If (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) = 0 Then
                    .RowSel = (PropRows - 1)
                    .ColSel = .Col
                    If .LeftCol > .Col Then
                        If .Col >= (PropFixedCols + PropFrozenCols) Then .LeftCol = .Col
                    ElseIf .ColSel > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .ColSel - GetColsPerPageRev(.ColSel) + 1
                    End If
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    ' Void
                ElseIf (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                    .RowSel = (PropRows - 1)
                    .ColSel = .Col
                    If .LeftCol > .Col Then
                        If .Col >= (PropFixedCols + PropFrozenCols) Then .LeftCol = .Col
                    ElseIf .Col > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .Col - GetColsPerPageRev(.Col) + 1
                    End If
                Else
                    ' Void
                End If
            Case vbKeyPageDown
                If (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) = 0 Then
                    .RowSel = (PropRows - 1)
                    .ColSel = .Col
                    If .LeftCol > .Col Then
                        If .Col >= (PropFixedCols + PropFrozenCols) Then .LeftCol = .Col
                    ElseIf .ColSel > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .ColSel - GetColsPerPageRev(.ColSel) + 1
                    End If
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    ' Void
                ElseIf (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                    .RowSel = (PropRows - 1)
                    .ColSel = .Col
                    If .LeftCol > .Col Then
                        If .Col >= (PropFixedCols + PropFrozenCols) Then .LeftCol = .Col
                    ElseIf .Col > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .Col - GetColsPerPageRev(.Col) + 1
                    End If
                Else
                    ' Void
                End If
            Case vbKeyHome
                If (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) = 0 Then
                    Call MoveFirstCol(.Col)
                    .RowSel = (PropRows - 1)
                    .ColSel = .Col
                    If .Col >= (PropFixedCols + PropFrozenCols) Then .LeftCol = .Col
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    Call MoveFirstCol(.ColSel)
                    If .ColSel >= (PropFixedCols + PropFrozenCols) Then .LeftCol = .ColSel
                ElseIf (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                    .Row = PropFixedRows
                    Call MoveFirstCol(.Col)
                    .RowSel = (PropRows - 1)
                    .ColSel = .Col
                    If .Col >= (PropFixedCols + PropFrozenCols) Then .LeftCol = .Col
                Else
                    .RowSel = (PropRows - 1)
                    Call MoveFirstCol(.ColSel)
                    If .ColSel >= (PropFixedCols + PropFrozenCols) Then .LeftCol = .ColSel
                End If
            Case vbKeyEnd
                If (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) = 0 Then
                    Call MoveLastCol(.Col)
                    .RowSel = (PropRows - 1)
                    .ColSel = .Col
                    .LeftCol = (PropCols - 1) - GetColsPerPageRev(PropCols - 1) + 1
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    Call MoveLastCol(.ColSel)
                    .LeftCol = (PropCols - 1) - GetColsPerPageRev(PropCols - 1) + 1
                ElseIf (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                    .Row = PropFixedRows
                    Call MoveLastCol(.Col)
                    .RowSel = (PropRows - 1)
                    .ColSel = .Col
                    .LeftCol = (PropCols - 1) - GetColsPerPageRev(PropCols - 1) + 1
                Else
                    .RowSel = (PropRows - 1)
                    Call MoveLastCol(.ColSel)
                    .LeftCol = (PropCols - 1) - GetColsPerPageRev(PropCols - 1) + 1
                End If
            Case vbKeyTab
                If (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) = 0 Then
                    If PropWrapCellBehavior <> FlexWrapNone Then
                        If .Col < GetLastMovableCol() Then
                            Call MoveNextCol(.Col)
                        Else
                            If PropWrapCellBehavior = FlexWrapGrid Then Call MoveFirstCol(.Col)
                        End If
                    End If
                    .RowSel = (PropRows - 1)
                    .ColSel = .Col
                    If .LeftCol > .Col Then
                        If .Col >= (PropFixedCols + PropFrozenCols) Then .LeftCol = .Col
                    ElseIf .Col > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .Col - GetColsPerPageRev(.Col) + 1
                    End If
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    If PropWrapCellBehavior <> FlexWrapNone Then
                        If .Col > GetFirstMovableCol() Then
                            Call MovePreviousCol(.Col)
                        Else
                            If PropWrapCellBehavior = FlexWrapGrid Then Call MoveLastCol(.Col)
                        End If
                    End If
                    .RowSel = (PropRows - 1)
                    .ColSel = .Col
                    If .LeftCol > .Col Then
                        If .Col >= (PropFixedCols + PropFrozenCols) Then .LeftCol = .Col
                    ElseIf .Col > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .Col - GetColsPerPageRev(.Col) + 1
                    End If
                ElseIf (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                    ' Void
                Else
                    ' Void
                End If
            Case vbKeyReturn
                If (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) = 0 Then
                    Select Case PropDirectionAfterReturn
                        Case FlexDirectionAfterReturnUp
                            If PropWrapCellBehavior <> FlexWrapNone Then
                                If .Col > GetFirstMovableCol() Then
                                    Call MovePreviousCol(.Col)
                                Else
                                    If PropWrapCellBehavior = FlexWrapGrid Then Call MoveLastCol(.Col)
                                End If
                            End If
                        Case FlexDirectionAfterReturnDown
                            If PropWrapCellBehavior <> FlexWrapNone Then
                                If .Col < GetLastMovableCol() Then
                                    Call MoveNextCol(.Col)
                                Else
                                    If PropWrapCellBehavior = FlexWrapGrid Then Call MoveFirstCol(.Col)
                                End If
                            End If
                        Case FlexDirectionAfterReturnLeft
                            If .Col > GetFirstMovableCol() Then
                                Call MovePreviousCol(.Col)
                            Else
                                If PropWrapCellBehavior = FlexWrapGrid Then Call MoveLastCol(.Col)
                            End If
                        Case FlexDirectionAfterReturnRight
                            If .Col < GetLastMovableCol() Then
                                Call MoveNextRow(.Col)
                            Else
                                If PropWrapCellBehavior = FlexWrapGrid Then Call MoveFirstRow(.Col)
                            End If
                    End Select
                    .RowSel = (PropRows - 1)
                    .ColSel = .Col
                    If .LeftCol > .Col Then
                        If .Col >= (PropFixedCols + PropFrozenCols) Then .LeftCol = .Col
                    ElseIf .Col > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .Col - GetColsPerPageRev(.Col) + 1
                    End If
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    Select Case PropDirectionAfterReturn
                        Case FlexDirectionAfterReturnUp
                            If PropWrapCellBehavior <> FlexWrapNone Then
                                If .Col < GetLastMovableCol() Then
                                    Call MoveNextCol(.Col)
                                Else
                                    If PropWrapCellBehavior = FlexWrapGrid Then Call MoveFirstCol(.Col)
                                End If
                            End If
                        Case FlexDirectionAfterReturnDown
                            If PropWrapCellBehavior <> FlexWrapNone Then
                                If .Col > GetFirstMovableCol() Then
                                    Call MovePreviousCol(.Col)
                                Else
                                    If PropWrapCellBehavior = FlexWrapGrid Then Call MoveLastCol(.Col)
                                End If
                            End If
                        Case FlexDirectionAfterReturnLeft
                            If .Col < GetLastMovableCol() Then
                                Call MoveNextRow(.Col)
                            Else
                                If PropWrapCellBehavior = FlexWrapGrid Then Call MoveFirstRow(.Col)
                            End If
                        Case FlexDirectionAfterReturnRight
                            If .Col > GetFirstMovableCol() Then
                                Call MovePreviousCol(.Col)
                            Else
                                If PropWrapCellBehavior = FlexWrapGrid Then Call MoveLastCol(.Col)
                            End If
                    End Select
                    .RowSel = (PropRows - 1)
                    .ColSel = .Col
                    If .LeftCol > .Col Then
                        If .Col >= (PropFixedCols + PropFrozenCols) Then .LeftCol = .Col
                    ElseIf .Col > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .Col - GetColsPerPageRev(.Col) + 1
                    End If
                ElseIf (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                    ' Void
                Else
                    ' Void
                End If
        End Select
End Select
If VBFlexGridCaptureRow > (PropFixedRows - 1) Or VBFlexGridCaptureCol > (PropFixedCols - 1) Then
    Dim HTI As THITTESTINFO, Pos As Long
    Pos = GetMessagePos()
    HTI.PT.X = Get_X_lParam(Pos)
    HTI.PT.Y = Get_Y_lParam(Pos)
    ScreenToClient VBFlexGridHandle, HTI.PT
    Call GetHitTestInfo(HTI)
    Select Case PropSelectionMode
        Case FlexSelectionModeFree, FlexSelectionModeFreeByRow, FlexSelectionModeFreeByColumn
            If VBFlexGridCaptureRow > (PropFixedRows - 1) Or PropAllowBigSelection = False Then
                If HTI.MouseRow > (PropFixedRows - 1) Then
                    .RowSel = HTI.MouseRow
                Else
                    .RowSel = .TopRow
                End If
            Else
                .RowSel = (PropRows - 1)
            End If
            If VBFlexGridCaptureCol > (PropFixedCols - 1) Or PropAllowBigSelection = False Then
                If HTI.MouseCol > (PropFixedCols - 1) Then
                    .ColSel = HTI.MouseCol
                Else
                    .ColSel = .LeftCol
                End If
            Else
                .ColSel = (PropCols - 1)
            End If
        Case FlexSelectionModeByRow
            If VBFlexGridCaptureRow > (PropFixedRows - 1) Or VBFlexGridCaptureCol > (PropFixedCols - 1) Or PropAllowBigSelection = False Then
                If HTI.MouseRow > (PropFixedRows - 1) Then
                    .RowSel = HTI.MouseRow
                Else
                    .RowSel = .TopRow
                End If
            End If
        Case FlexSelectionModeByColumn
            If VBFlexGridCaptureRow > (PropFixedRows - 1) Or VBFlexGridCaptureCol > (PropFixedCols - 1) Or PropAllowBigSelection = False Then
                If HTI.MouseCol > (PropFixedCols - 1) Then
                    .ColSel = HTI.MouseCol
                Else
                    .ColSel = .LeftCol
                End If
            End If
    End Select
End If
Call SetRowColParams(RCP)
End With
End Sub

Private Function ProcessLButtonDown(ByVal Shift As Integer, ByRef HTI As THITTESTINFO) As Boolean
VBFlexGridCaptureRow = HTI.HitRow
VBFlexGridCaptureCol = HTI.HitCol
VBFlexGridCaptureDividerRow = HTI.HitRowDivider
VBFlexGridCaptureDividerCol = HTI.HitColDivider
VBFlexGridCaptureHitResult = HTI.HitResult
VBFlexGridMouseMoveRow = HTI.HitRow
VBFlexGridMouseMoveCol = HTI.HitCol
VBFlexGridMouseMoveChanged = False
If HTI.HitResult = FlexHitResultNoWhere Then
    Exit Function
ElseIf HTI.HitResult <> FlexHitResultCell Then
    Select Case VBFlexGridCaptureHitResult
        Case FlexHitResultDividerRowTop, FlexHitResultDividerRowBottom, FlexHitResultDividerColumnLeft, FlexHitResultDividerColumnRight
            VBFlexGridCaptureDividerDrag = True
        Case Else
            Exit Function
    End Select
    Dim iRow As Long, iCol As Long, Cancel As Boolean
    iRow = VBFlexGridCaptureDividerRow
    iCol = VBFlexGridCaptureDividerCol
    RaiseEvent BeforeUserResize(iRow, iCol, Cancel)
    If Cancel = False Then
        Dim ClipRect As RECT, i As Long, P As POINTAPI
        LSet ClipRect = VBFlexGridClientRect
        With ClipRect
        If iRow > -1 Then
            For i = 0 To iRow - 1
                If i >= VBFlexGridTopRow Or i < (PropFixedRows + PropFrozenRows) Then
                    .Top = .Top + GetRowHeight(i)
                End If
            Next i
            P.Y = .Top + GetRowHeight(iRow)
            .Top = .Top + (HTI.PT.Y - P.Y) + 1
            .Bottom = .Bottom + (HTI.PT.Y - P.Y) - 1
        End If
        If iCol > -1 Then
            For i = 0 To iCol - 1
                If i >= VBFlexGridLeftCol Or i < (PropFixedCols + PropFrozenCols) Then
                    .Left = .Left + GetColWidth(i)
                End If
            Next i
            P.X = .Left + GetColWidth(iCol)
            .Left = .Left + (HTI.PT.X - P.X) + 1
            .Right = .Right + (HTI.PT.X - P.X) - 1
        End If
        End With
        MapWindowPoints VBFlexGridHandle, HWND_DESKTOP, ClipRect, 2
        ClipCursor ClipRect
        VBFlexGridDividerDragOffset.X = HTI.PT.X - P.X
        VBFlexGridDividerDragOffset.Y = HTI.PT.Y - P.Y
        Call SetDividerDragSplitterRect(P.X, P.Y)
        Call DrawDividerDragSplitter
        ProcessLButtonDown = True
    Else
        ReleaseCapture
    End If
    Exit Function
End If
Dim RCP As TROWCOLPARAMS
With RCP
.Mask = RCPM_ROW Or RCPM_COL Or RCPM_ROWSEL Or RCPM_COLSEL
.Row = VBFlexGridRow
.Col = VBFlexGridCol
.RowSel = VBFlexGridRowSel
.ColSel = VBFlexGridColSel
Select Case PropSelectionMode
    Case FlexSelectionModeFree, FlexSelectionModeFreeByRow, FlexSelectionModeFreeByColumn
        If HTI.HitRow > (PropFixedRows - 1) Then
            If (Shift And vbShiftMask) = 0 Then
                .Row = HTI.HitRow
                .RowSel = .Row
            Else
                .RowSel = HTI.HitRow
            End If
        Else
            If PropAllowBigSelection = True Then
                If HTI.HitCol < PropFixedCols Or PropSelectionMode <> FlexSelectionModeFreeByRow Then
                    .Row = PropFixedRows
                    .RowSel = (PropRows - 1)
                End If
            Else
                If (Shift And vbShiftMask) = 0 Then
                    .Row = VBFlexGridTopRow
                    .RowSel = .Row
                Else
                    .RowSel = VBFlexGridTopRow
                End If
            End If
        End If
        If HTI.HitCol > (PropFixedCols - 1) Then
            If (Shift And vbShiftMask) = 0 Then
                .Col = HTI.HitCol
                .ColSel = .Col
            Else
                .ColSel = HTI.HitCol
            End If
        Else
            If PropAllowBigSelection = True Then
                If HTI.HitRow < PropFixedRows Or PropSelectionMode <> FlexSelectionModeFreeByColumn Then
                    .Col = PropFixedCols
                    .ColSel = (PropCols - 1)
                End If
            Else
                If (Shift And vbShiftMask) = 0 Then
                    .Col = VBFlexGridLeftCol
                    .ColSel = .Col
                Else
                    .ColSel = VBFlexGridLeftCol
                End If
            End If
        End If
    Case FlexSelectionModeByRow
        If HTI.HitRow > (PropFixedRows - 1) Then
            If (Shift And vbShiftMask) = 0 Then
                .Row = HTI.HitRow
                .RowSel = .Row
            Else
                .RowSel = HTI.HitRow
            End If
            .ColSel = (PropCols - 1)
        Else
            If PropAllowBigSelection = True Then
                If HTI.HitCol > (PropFixedCols - 1) Then
                    If (Shift And vbShiftMask) = 0 Then
                        .Row = PropFixedRows
                        .RowSel = .Row
                    Else
                        .RowSel = PropFixedRows
                    End If
                Else
                    .Row = PropFixedRows
                    .RowSel = (PropRows - 1)
                End If
                .ColSel = (PropCols - 1)
            Else
                If (Shift And vbShiftMask) = 0 Then
                    .Row = VBFlexGridTopRow
                    .RowSel = .Row
                Else
                    .RowSel = VBFlexGridTopRow
                End If
                .ColSel = (PropCols - 1)
            End If
        End If
    Case FlexSelectionModeByColumn
        If HTI.HitCol > (PropFixedCols - 1) Then
            If (Shift And vbShiftMask) = 0 Then
                .Col = HTI.HitCol
                .ColSel = .Col
            Else
                .ColSel = HTI.HitCol
            End If
            .RowSel = (PropRows - 1)
        Else
            If PropAllowBigSelection = True Then
                If HTI.HitRow > (PropFixedRows - 1) Then
                    If (Shift And vbShiftMask) = 0 Then
                        .Col = PropFixedCols
                        .ColSel = .Col
                    Else
                        .ColSel = PropFixedCols
                    End If
                Else
                    .Col = PropFixedCols
                    .ColSel = (PropCols - 1)
                End If
                .RowSel = (PropRows - 1)
            Else
                If (Shift And vbShiftMask) = 0 Then
                    .Col = VBFlexGridLeftCol
                    .ColSel = .Col
                Else
                    .ColSel = VBFlexGridLeftCol
                End If
                .RowSel = (PropRows - 1)
            End If
        End If
End Select
If HTI.HitRow <= (PropFixedRows - 1) And HTI.HitCol <= (PropFixedCols - 1) Then
    Select Case PropSelectionMode
        Case FlexSelectionModeFree, FlexSelectionModeFreeByRow, FlexSelectionModeFreeByColumn
            If PropAllowBigSelection = True Or (Shift And vbShiftMask) = 0 Then
                .Mask = .Mask Or RCPM_TOPROW Or RCPM_LEFTCOL
                If .Row >= (PropFixedRows + PropFrozenRows) Then
                    .TopRow = .Row
                Else
                    .TopRow = (PropFixedRows + PropFrozenRows)
                End If
                If .Col >= (PropFixedCols + PropFrozenCols) Then
                    .LeftCol = .Col
                Else
                    .LeftCol = (PropFixedCols + PropFrozenCols)
                End If
            End If
        Case FlexSelectionModeByRow, FlexSelectionModeByColumn
            If PropAllowBigSelection = True And (Shift And vbShiftMask) = 0 Then
                .Mask = .Mask Or RCPM_TOPROW Or RCPM_LEFTCOL
                If .Row >= (PropFixedRows + PropFrozenRows) Then
                    .TopRow = .Row
                Else
                    .TopRow = (PropFixedRows + PropFrozenRows)
                End If
                If .Col >= (PropFixedCols + PropFrozenCols) Then
                    .LeftCol = .Col
                Else
                    .LeftCol = (PropFixedCols + PropFrozenCols)
                End If
            End If
    End Select
End If
Call SetRowColParams(RCP)
End With
End Function

Private Sub ProcessLButtonUp(ByVal X As Long, ByVal Y As Long)
Dim RCP As TROWCOLPARAMS
If VBFlexGridCaptureDividerDrag = True Then
    Dim iRow As Long, iCol As Long
    iRow = VBFlexGridCaptureDividerRow
    iCol = VBFlexGridCaptureDividerCol
    Dim Size As SIZEAPI, NewSize As Long, i As Long
    With Size
    If iRow > -1 Then
        For i = 0 To iRow - 1
            If i >= VBFlexGridTopRow Then
                .CY = .CY + GetRowHeight(i)
            ElseIf i < (PropFixedRows + PropFrozenRows) Then
                .CY = .CY + GetRowHeight(i)
            Else
                i = VBFlexGridTopRow - 1
            End If
        Next i
        If (Y - VBFlexGridDividerDragOffset.Y) < (.CY + 1) Then
            NewSize = UserControl.ScaleY(1, vbPixels, vbTwips)
        ElseIf (Y - VBFlexGridDividerDragOffset.Y) >= (VBFlexGridClientRect.Bottom - 1) Then
            NewSize = UserControl.ScaleY(((VBFlexGridClientRect.Bottom - 1) - .CY), vbPixels, vbTwips)
        Else
            NewSize = UserControl.ScaleY(((Y - VBFlexGridDividerDragOffset.Y) - .CY), vbPixels, vbTwips)
        End If
        RaiseEvent AfterUserResize(iRow, iCol, NewSize)
        If NewSize > 0 Then .CY = UserControl.ScaleY(NewSize, vbTwips, vbPixels) Else .CY = 0
        VBFlexGridCells.Rows(iRow).RowInfo.Height = .CY
        If PropRowSizingMode = FlexRowSizingModeAll Then
            For i = 0 To PropRows - 1
                VBFlexGridCells.Rows(i).RowInfo.Height = .CY
            Next i
        End If
    ElseIf iCol > -1 Then
        For i = 0 To iCol - 1
            If i >= VBFlexGridLeftCol Then
                .CX = .CX + GetColWidth(i)
            ElseIf i < (PropFixedCols + PropFrozenCols) Then
                .CX = .CX + GetColWidth(i)
            Else
                i = VBFlexGridLeftCol - 1
            End If
        Next i
        If (X - VBFlexGridDividerDragOffset.X) < (.CX + 1) Then
            NewSize = UserControl.ScaleX(1, vbPixels, vbTwips)
        ElseIf (X - VBFlexGridDividerDragOffset.X) >= (VBFlexGridClientRect.Right - 1) Then
            NewSize = UserControl.ScaleX(((VBFlexGridClientRect.Right - 1) - .CX), vbPixels, vbTwips)
        Else
            NewSize = UserControl.ScaleX(((X - VBFlexGridDividerDragOffset.X) - .CX), vbPixels, vbTwips)
        End If
        RaiseEvent AfterUserResize(iRow, iCol, NewSize)
        If NewSize > 0 Then .CX = UserControl.ScaleX(NewSize, vbTwips, vbPixels) Else .CX = 0
        VBFlexGridColsInfo(iCol).Width = .CX
    End If
    End With
    ClipCursor ByVal 0&
    SetRect VBFlexGridDividerDragSplitterRect, 0, 0, 0, 0
    VBFlexGridDividerDragOffset.X = 0
    VBFlexGridDividerDragOffset.Y = 0
    With RCP
    .Mask = RCPM_TOPROW Or RCPM_LEFTCOL
    .Flags = RCPF_CHECKTOPROW Or RCPF_CHECKLEFTCOL Or RCPF_SETSCROLLBARS Or RCPF_FORCEREDRAW
    .TopRow = VBFlexGridTopRow
    .LeftCol = VBFlexGridLeftCol
    Call SetRowColParams(RCP)
    End With
    Exit Sub
End If
If VBFlexGridMouseMoveChanged = False Then
    With RCP
    .Mask = RCPM_TOPROW Or RCPM_LEFTCOL
    .TopRow = VBFlexGridTopRow
    .LeftCol = VBFlexGridLeftCol
    If .TopRow <= VBFlexGridRow Then
        If VBFlexGridRow > (.TopRow + GetRowsPerPage(.TopRow) - 1) Then
            .TopRow = VBFlexGridRow - GetRowsPerPageRev(VBFlexGridRow) + 1
        End If
    End If
    If .LeftCol <= VBFlexGridCol Then
        If VBFlexGridCol > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
            .LeftCol = VBFlexGridCol - GetColsPerPageRev(VBFlexGridCol) + 1
        End If
    End If
    Call SetRowColParams(RCP)
    End With
End If
End Sub

Private Sub ProcessMouseMove(ByVal Button As Integer, ByVal X As Long, ByVal Y As Long)
If PropShowInfoTips = True Or PropShowLabelTips = True Then Call CheckToolTipRowCol(X, Y)
If VBFlexGridCaptureDividerDrag = True Then
    Call DrawDividerDragSplitter
    Call SetDividerDragSplitterRect(X - VBFlexGridDividerDragOffset.X, Y - VBFlexGridDividerDragOffset.Y)
    Call DrawDividerDragSplitter
    Exit Sub
End If
If VBFlexGridCaptureRow = -1 Or VBFlexGridCaptureCol = -1 Or VBFlexGridCaptureHitResult = FlexHitResultNoWhere Then Exit Sub
If (Button And vbLeftButton) = 0 Then Exit Sub
If VBFlexGridCaptureRow <= (PropFixedRows - 1) And VBFlexGridCaptureCol <= (PropFixedCols - 1) Then Exit Sub
Dim HTI As THITTESTINFO
HTI.PT.X = X
HTI.PT.Y = Y
Call GetHitTestInfo(HTI)
If HTI.HitRow <> VBFlexGridMouseMoveRow Or HTI.HitCol <> VBFlexGridMouseMoveCol Then
    VBFlexGridMouseMoveRow = HTI.HitRow
    VBFlexGridMouseMoveCol = HTI.HitCol
    VBFlexGridMouseMoveChanged = True
Else
    If VBFlexGridMouseMoveChanged = False Then Exit Sub
End If
Dim RCP As TROWCOLPARAMS, RowsPerPage As Long, ColsPerPage As Long
With RCP
.Mask = RCPM_ROWSEL Or RCPM_COLSEL Or RCPM_TOPROW Or RCPM_LEFTCOL
If PropAllowSelection = False Then
    .Mask = .Mask Or RCPM_ROW Or RCPM_COL
    .Row = VBFlexGridRow
    .Col = VBFlexGridCol
End If
.RowSel = VBFlexGridRowSel
.ColSel = VBFlexGridColSel
.TopRow = VBFlexGridTopRow
.LeftCol = VBFlexGridLeftCol
Select Case PropSelectionMode
    Case FlexSelectionModeFree, FlexSelectionModeFreeByRow, FlexSelectionModeFreeByColumn
        If VBFlexGridCaptureRow > (PropFixedRows - 1) Or PropAllowBigSelection = False Then
            If HTI.MouseRow > ((PropFixedRows + PropFrozenRows) - 1) Then
                .RowSel = HTI.MouseRow
            Else
                If .RowSel > (PropFixedRows + PropFrozenRows) Then
                    .RowSel = .RowSel - 1
                ElseIf HTI.MouseRow > (PropFixedRows - 1) Then
                    .RowSel = HTI.MouseRow
                ElseIf .RowSel > PropFixedRows Then
                    .RowSel = .RowSel - 1
                End If
            End If
            If .TopRow > .RowSel Then
                If .RowSel >= (PropFixedRows + PropFrozenRows) Then .TopRow = .RowSel
            Else
                RowsPerPage = GetRowsPerPage(.TopRow)
                If .RowSel > (.TopRow + RowsPerPage - 1) Then
                    .TopRow = .RowSel - GetRowsPerPageRev(.RowSel) + 1
                End If
            End If
            If PropAllowSelection = False Then .Row = .RowSel
        ElseIf PropSelectionMode <> FlexSelectionModeFreeByRow Then
            .RowSel = (PropRows - 1)
        End If
        If VBFlexGridCaptureCol > (PropFixedCols - 1) Or PropAllowBigSelection = False Then
            If HTI.MouseCol > ((PropFixedCols + PropFrozenCols) - 1) Then
                .ColSel = HTI.MouseCol
            Else
                If .ColSel > (PropFixedCols + PropFrozenCols) Then
                    .ColSel = .ColSel - 1
                ElseIf HTI.MouseCol > (PropFixedCols - 1) Then
                    .ColSel = HTI.MouseCol
                ElseIf .ColSel > PropFixedCols Then
                    .ColSel = .ColSel - 1
                End If
            End If
            If .LeftCol > .ColSel Then
                If .ColSel >= (PropFixedCols + PropFrozenCols) Then .LeftCol = .ColSel
            Else
                ColsPerPage = GetColsPerPage(.LeftCol)
                If .ColSel > (.LeftCol + ColsPerPage - 1) Then
                    .LeftCol = .ColSel - GetColsPerPageRev(.ColSel) + 1
                End If
            End If
            If PropAllowSelection = False Then .Col = .ColSel
        ElseIf PropSelectionMode <> FlexSelectionModeFreeByColumn Then
            .ColSel = (PropCols - 1)
        End If
    Case FlexSelectionModeByRow
        If VBFlexGridCaptureRow > (PropFixedRows - 1) Or VBFlexGridCaptureCol > (PropFixedCols - 1) Or PropAllowBigSelection = False Then
            If HTI.MouseRow > ((PropFixedRows + PropFrozenRows) - 1) Then
                .RowSel = HTI.MouseRow
            Else
                If .RowSel > (PropFixedRows + PropFrozenRows) Then
                    .RowSel = .RowSel - 1
                ElseIf HTI.MouseRow > (PropFixedRows - 1) Then
                    .RowSel = HTI.MouseRow
                ElseIf .RowSel > PropFixedRows Then
                    .RowSel = .RowSel - 1
                End If
            End If
            If .TopRow > .RowSel Then
                If .RowSel >= (PropFixedRows + PropFrozenRows) Then .TopRow = .RowSel
            Else
                RowsPerPage = GetRowsPerPage(.TopRow)
                If .RowSel > (.TopRow + RowsPerPage - 1) Then
                    .TopRow = .RowSel - GetRowsPerPageRev(.RowSel) + 1
                End If
            End If
            If PropAllowSelection = False Then .Row = .RowSel
        End If
    Case FlexSelectionModeByColumn
        If VBFlexGridCaptureRow > (PropFixedRows - 1) Or VBFlexGridCaptureCol > (PropFixedCols - 1) Or PropAllowBigSelection = False Then
            If HTI.MouseCol > ((PropFixedCols + PropFrozenCols) - 1) Then
                .ColSel = HTI.MouseCol
            Else
                If .ColSel > (PropFixedCols + PropFrozenCols) Then
                    .ColSel = .ColSel - 1
                ElseIf HTI.MouseCol > (PropFixedCols - 1) Then
                    .ColSel = HTI.MouseCol
                ElseIf .ColSel > PropFixedCols Then
                    .ColSel = .ColSel - 1
                End If
            End If
            If .LeftCol > .ColSel Then
                If .ColSel >= (PropFixedCols + PropFrozenCols) Then .LeftCol = .ColSel
            Else
                ColsPerPage = GetColsPerPage(.LeftCol)
                If .ColSel > (.LeftCol + ColsPerPage - 1) Then
                    .LeftCol = .ColSel - GetColsPerPageRev(.ColSel) + 1
                End If
            End If
            If PropAllowSelection = False Then .Col = .ColSel
        End If
End Select
Call SetRowColParams(RCP)
End With
End Sub

Private Function MergeCompareFunction(ByVal Row1 As Long, ByVal Col1 As Long, ByVal Row2 As Long, ByVal Col2 As Long) As Boolean
Dim Text1 As String, Text2 As String
Call GetCellText(Row1, Col1, Text1)
Call GetCellText(Row2, Col2, Text2)
If Text1 = vbNullString Or Text2 = vbNullString Then Exit Function
If StrComp(Text1, Text2) = 0 Then MergeCompareFunction = True
End Function

Private Sub DrawDividerDragSplitter()
If VBFlexGridHandle = 0 Or VBFlexGridCaptureDividerDrag = False Or FlexGetSplitterBrush() = 0 Then Exit Sub
Dim hDC As Long, hBmpOld As Long
hDC = GetDC(VBFlexGridHandle)
If hDC <> 0 Then
    hBmpOld = SelectObject(hDC, FlexGetSplitterBrush())
    With VBFlexGridDividerDragSplitterRect
    PatBlt hDC, .Left, .Top, .Right - .Left, .Bottom - .Top, vbPatInvert
    End With
    SelectObject hDC, hBmpOld
    ReleaseDC VBFlexGridHandle, hDC
End If
End Sub

Private Sub SetDividerDragSplitterRect(ByVal X As Long, ByVal Y As Long)
If VBFlexGridHandle = 0 Or VBFlexGridCaptureDividerDrag = False Then Exit Sub
LSet VBFlexGridDividerDragSplitterRect = VBFlexGridClientRect
With VBFlexGridDividerDragSplitterRect
Select Case VBFlexGridCaptureHitResult
    Case FlexHitResultDividerRowTop, FlexHitResultDividerRowBottom
        .Top = Y - (1 * PixelsPerDIP_Y())
        .Bottom = Y + (1 * PixelsPerDIP_Y())
    Case FlexHitResultDividerColumnLeft, FlexHitResultDividerColumnRight
        .Left = X - (1 * PixelsPerDIP_X())
        .Right = X + (1 * PixelsPerDIP_X())
End Select
End With
End Sub

Private Function GetColSeparator() As String
If PropClipSeparators = vbNullString Then
    GetColSeparator = vbTab
Else
    GetColSeparator = Left$(PropClipSeparators, 1)
End If
End Function

Private Function GetRowSeparator() As String
If PropClipSeparators = vbNullString Then
    GetRowSeparator = vbCr
Else
    GetRowSeparator = Right$(PropClipSeparators, 1)
End If
End Function

Private Sub SetVisualStylesToolTip()
If VBFlexGridHandle <> 0 Then
    If VBFlexGridToolTipHandle <> 0 And VBFlexGridEnabledVisualStyles = True Then
        If PropVisualStyles = True Then
            ActivateVisualStyles VBFlexGridToolTipHandle
        Else
            RemoveVisualStyles VBFlexGridToolTipHandle
        End If
    End If
End If
End Sub

Private Sub SetIMEMode(ByVal hWnd As Long, ByVal hIMCOrig As Long, ByVal Value As FlexIMEModeConstants)
Const IME_CMODE_ALPHANUMERIC As Long = &H0, IME_CMODE_NATIVE As Long = &H1, IME_CMODE_KATAKANA As Long = &H2, IME_CMODE_FULLSHAPE As Long = &H8
Dim hKL As Long
hKL = GetKeyboardLayout(0)
If ImmIsIME(hKL) = 0 Or hIMCOrig = 0 Then Exit Sub
Dim hIMC As Long
hIMC = ImmGetContext(hWnd)
If Value = FlexIMEModeDisable Then
    If hIMC <> 0 Then
        ImmReleaseContext hWnd, hIMC
        ImmAssociateContext hWnd, 0
    End If
Else
    If hIMC = 0 Then
        ImmAssociateContext hWnd, hIMCOrig
        hIMC = ImmGetContext(hWnd)
    End If
    If hIMC <> 0 And Value <> FlexIMEModeNoControl Then
        Dim dwConversion As Long, dwSentence As Long
        ImmGetConversionStatus hIMC, dwConversion, dwSentence
        Select Case Value
            Case FlexIMEModeOn
                ImmSetOpenStatus hIMC, 1
            Case FlexIMEModeOff
                ImmSetOpenStatus hIMC, 0
            Case FlexIMEModeHiragana
                ImmSetOpenStatus hIMC, 1
                If Not (dwConversion And IME_CMODE_NATIVE) = IME_CMODE_NATIVE Then dwConversion = dwConversion Or IME_CMODE_NATIVE
                If Not (dwConversion And IME_CMODE_FULLSHAPE) = IME_CMODE_FULLSHAPE Then dwConversion = dwConversion Or IME_CMODE_FULLSHAPE
                If (dwConversion And IME_CMODE_KATAKANA) = IME_CMODE_KATAKANA Then dwConversion = dwConversion And Not IME_CMODE_KATAKANA
            Case FlexIMEModeKatakana
                ImmSetOpenStatus hIMC, 1
                If Not (dwConversion And IME_CMODE_NATIVE) = IME_CMODE_NATIVE Then dwConversion = dwConversion Or IME_CMODE_NATIVE
                If Not (dwConversion And IME_CMODE_KATAKANA) = IME_CMODE_KATAKANA Then dwConversion = dwConversion Or IME_CMODE_KATAKANA
                If Not (dwConversion And IME_CMODE_FULLSHAPE) = IME_CMODE_FULLSHAPE Then dwConversion = dwConversion Or IME_CMODE_FULLSHAPE
            Case FlexIMEModeKatakanaHalf
                ImmSetOpenStatus hIMC, 1
                If Not (dwConversion And IME_CMODE_NATIVE) = IME_CMODE_NATIVE Then dwConversion = dwConversion Or IME_CMODE_NATIVE
                If Not (dwConversion And IME_CMODE_KATAKANA) = IME_CMODE_KATAKANA Then dwConversion = dwConversion Or IME_CMODE_KATAKANA
                If (dwConversion And IME_CMODE_FULLSHAPE) = IME_CMODE_FULLSHAPE Then dwConversion = dwConversion And Not IME_CMODE_FULLSHAPE
            Case FlexIMEModeAlphaFull
                ImmSetOpenStatus hIMC, 1
                If Not (dwConversion And IME_CMODE_FULLSHAPE) = IME_CMODE_FULLSHAPE Then dwConversion = dwConversion Or IME_CMODE_FULLSHAPE
                If (dwConversion And IME_CMODE_NATIVE) = IME_CMODE_NATIVE Then dwConversion = dwConversion And Not IME_CMODE_NATIVE
                If (dwConversion And IME_CMODE_KATAKANA) = IME_CMODE_KATAKANA Then dwConversion = dwConversion And Not IME_CMODE_KATAKANA
            Case FlexIMEModeAlpha
                ImmSetOpenStatus hIMC, 1
                If Not (dwConversion And IME_CMODE_ALPHANUMERIC) = IME_CMODE_ALPHANUMERIC Then dwConversion = dwConversion Or IME_CMODE_ALPHANUMERIC
                If (dwConversion And IME_CMODE_NATIVE) = IME_CMODE_NATIVE Then dwConversion = dwConversion And Not IME_CMODE_NATIVE
                If (dwConversion And IME_CMODE_KATAKANA) = IME_CMODE_KATAKANA Then dwConversion = dwConversion And Not IME_CMODE_KATAKANA
                If (dwConversion And IME_CMODE_FULLSHAPE) = IME_CMODE_FULLSHAPE Then dwConversion = dwConversion And Not IME_CMODE_FULLSHAPE
            Case FlexIMEModeHangulFull
                ImmSetOpenStatus hIMC, 1
                If Not (dwConversion And IME_CMODE_NATIVE) = IME_CMODE_NATIVE Then dwConversion = dwConversion Or IME_CMODE_NATIVE
                If Not (dwConversion And IME_CMODE_FULLSHAPE) = IME_CMODE_FULLSHAPE Then dwConversion = dwConversion Or IME_CMODE_FULLSHAPE
            Case FlexIMEModeHangul
                ImmSetOpenStatus hIMC, 1
                If Not (dwConversion And IME_CMODE_NATIVE) = IME_CMODE_NATIVE Then dwConversion = dwConversion Or IME_CMODE_NATIVE
                If (dwConversion And IME_CMODE_FULLSHAPE) = IME_CMODE_FULLSHAPE Then dwConversion = dwConversion And Not IME_CMODE_FULLSHAPE
        End Select
        ImmSetConversionStatus hIMC, dwConversion, dwSentence
        ImmReleaseContext hWnd, hIMC
    End If
End If
End Sub

Private Sub UpdateToolTipRect()
If VBFlexGridHandle <> 0 And VBFlexGridToolTipHandle <> 0 Then
    Dim TI As TOOLINFO
    With TI
    .cbSize = LenB(TI)
    .hWnd = VBFlexGridHandle
    .uId = 0
    LSet .RC = VBFlexGridClientRect
    SendMessage VBFlexGridToolTipHandle, TTM_NEWTOOLRECT, 0, ByVal VarPtr(TI)
    End With
End If
End Sub

Private Sub CheckToolTipRowCol(ByVal X As Long, ByVal Y As Long)
If VBFlexGridHandle <> 0 And VBFlexGridToolTipHandle <> 0 Then
    Dim HTI As THITTESTINFO
    With HTI
    .PT.X = X
    .PT.Y = Y
    Call GetHitTestInfo(HTI)
    If .HitResult <> FlexHitResultNoWhere Then
        If VBFlexGridToolTipRow <> .HitRow Or VBFlexGridToolTipCol <> .HitCol Then
            VBFlexGridToolTipRow = .HitRow
            VBFlexGridToolTipCol = .HitCol
            SendMessage VBFlexGridToolTipHandle, TTM_POP, 0, ByVal 0&
        End If
    Else
        VBFlexGridToolTipRow = -1
        VBFlexGridToolTipCol = -1
        If VBFlexGridToolTipHandle <> 0 Then SendMessage VBFlexGridToolTipHandle, TTM_POP, 0, ByVal 0&
    End If
    End With
End If
End Sub

Private Sub UpdateEditRect()
If PropRows < 1 Or PropCols < 1 Then Exit Sub
If VBFlexGridHandle <> 0 And VBFlexGridEditHandle <> 0 Then
    With VBFlexGridEditMergedRange
    If .TopRow < (PropFixedRows + PropFrozenRows) And .LeftCol < (PropFixedCols + PropFrozenCols) Then
        ' Void
    Else
        Dim RC As RECT, i As Long
        If .BottomRow >= VBFlexGridTopRow Then
            For i = 0 To ((PropFixedRows + PropFrozenRows) - 1)
                RC.Top = RC.Bottom
                RC.Bottom = RC.Bottom + GetRowHeight(i)
            Next i
            For i = VBFlexGridTopRow To .TopRow
                RC.Top = RC.Bottom
                RC.Bottom = RC.Bottom + GetRowHeight(i)
            Next i
            If .TopRow < VBFlexGridTopRow Then RC.Top = RC.Bottom
            For i = (.TopRow + 1) To .BottomRow
                If i >= VBFlexGridTopRow Then RC.Bottom = RC.Bottom + GetRowHeight(i)
            Next i
        ElseIf .BottomRow < (PropFixedRows + PropFrozenRows) Then
            For i = 0 To .TopRow
                RC.Top = RC.Bottom
                RC.Bottom = RC.Bottom + GetRowHeight(i)
            Next i
            For i = (.TopRow + 1) To .BottomRow
                RC.Bottom = RC.Bottom + GetRowHeight(i)
            Next i
        Else
            For i = .TopRow To .BottomRow
                RC.Bottom = RC.Top
                RC.Top = RC.Top - GetRowHeight(i)
            Next i
        End If
        If .RightCol >= VBFlexGridLeftCol Then
            For i = 0 To ((PropFixedCols + PropFrozenCols) - 1)
                RC.Left = RC.Right
                RC.Right = RC.Right + GetColWidth(i)
            Next i
            For i = VBFlexGridLeftCol To .LeftCol
                RC.Left = RC.Right
                RC.Right = RC.Right + GetColWidth(i)
            Next i
            If .LeftCol < VBFlexGridLeftCol Then RC.Left = RC.Right
            For i = (.LeftCol + 1) To .RightCol
                If i >= VBFlexGridLeftCol Then RC.Right = RC.Right + GetColWidth(i)
            Next i
        ElseIf .RightCol < (PropFixedCols + PropFrozenCols) Then
            For i = 0 To .LeftCol
                RC.Left = RC.Right
                RC.Right = RC.Right + GetColWidth(i)
            Next i
            For i = (.LeftCol + 1) To .RightCol
                RC.Right = RC.Right + GetColWidth(i)
            Next i
        Else
            For i = .LeftCol To .RightCol
                RC.Right = RC.Left
                RC.Left = RC.Left - GetColWidth(i)
            Next i
        End If
        SetWindowPos VBFlexGridEditHandle, 0, RC.Left, RC.Top, 0, 0, SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER
        If VBFlexGridComboButtonHandle <> 0 Then
            Dim EditRect As RECT
            GetClientRect VBFlexGridEditHandle, EditRect
            SetWindowPos VBFlexGridComboButtonHandle, 0, RC.Left + (EditRect.Right - EditRect.Left), RC.Top, 0, 0, SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER
            If VBFlexGridComboListHandle <> 0 Then
                Dim WndRect As RECT
                SetRect VBFlexGridComboListRect, RC.Left, RC.Top, RC.Right, RC.Bottom
                LSet WndRect = VBFlexGridComboListRect
                MapWindowPoints VBFlexGridHandle, HWND_DESKTOP, WndRect, 2
                SetWindowPos VBFlexGridComboListHandle, 0, WndRect.Left, WndRect.Top, 0, 0, SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_NOACTIVATE
            End If
        End If
        If VBFlexGridEditRectChangedFrozen = False Then VBFlexGridEditRectChanged = True
    End If
    End With
End If
End Sub

Private Function ValidateEditOnMouseActivateMsg(ByVal lParam As Long, ByRef RetVal As Long) As Boolean
If VBFlexGridHandle <> 0 And VBFlexGridEditHandle <> 0 Then
    Select Case HiWord(lParam)
        Case WM_LBUTTONDOWN
            If VBFlexGridComboButtonHandle <> 0 Then
                If VBFlexGridComboButtonClick = False Then
                    If IsWindowEnabled(VBFlexGridComboButtonHandle) = 0 Then
                        ' If the combo button window is disabled the mouse message will go trough it and could trigger ending of the editing.
                        ' To avoid this a check is needed and return MA_ACTIVATEANDEAT, if necessary.
                        Dim Pos As Long, P As POINTAPI
                        Pos = GetMessagePos()
                        P.X = Get_X_lParam(Pos)
                        P.Y = Get_Y_lParam(Pos)
                        ScreenToClient VBFlexGridHandle, P
                        If ChildWindowFromPoint(VBFlexGridHandle, P.X, P.Y) = VBFlexGridComboButtonHandle Then
                            RetVal = MA_ACTIVATEANDEAT
                            ValidateEditOnMouseActivateMsg = True
                            Exit Function
                        End If
                    End If
                Else
                    RetVal = MA_ACTIVATEANDEAT
                    ValidateEditOnMouseActivateMsg = True
                    Exit Function
                End If
            End If
            If LoWord(lParam) = HTCLIENT And VBFlexGridEditTextChanged = True Then
                Dim Cancel As Boolean
                VBFlexGridEditOnValidate = True
                RaiseEvent ValidateEdit(Cancel)
                VBFlexGridEditOnValidate = False
                If VBFlexGridEditHandle <> 0 Then
                    If Cancel = True Then
                        ' Edit control remains active and will not be destroyed.
                        RetVal = MA_ACTIVATEANDEAT
                        ValidateEditOnMouseActivateMsg = True
                        Exit Function
                    Else
                        VBFlexGridEditAlreadyValidated = True
                    End If
                End If
            End If
    End Select
End If
End Function

Private Sub ComboShowDropDown(ByVal Value As Boolean)
If VBFlexGridEditHandle <> 0 And VBFlexGridComboButtonHandle <> 0 And VBFlexGridComboListHandle <> 0 Then
    Dim dwLong As Long
    dwLong = GetWindowLong(VBFlexGridComboButtonHandle, GWL_USERDATA)
    If Value = True Then
        If Not (dwLong And ODS_SELECTED) = ODS_SELECTED And Not (dwLong And ODS_DISABLED) = ODS_DISABLED Then
            If GetCursor() = 0 Then
                ' The mouse cursor can be hidden when showing the drop-down list upon a change event.
                ' Reason is that the edit control hides the cursor and a following mouse move will show it again.
                ' However, the drop-down list will set a mouse capture and thus the cursor keeps hidden.
                ' Solution is to refresh the cursor by sending a WM_SETCURSOR.
                Call RefreshMousePointer(VBFlexGridEditHandle)
            End If
            RaiseEvent ComboDropDown
            SetWindowLong VBFlexGridComboButtonHandle, GWL_USERDATA, dwLong Or ODS_SELECTED
            InvalidateRect VBFlexGridComboButtonHandle, ByVal 0&, 0
            If IsWindowVisible(VBFlexGridComboListHandle) = 0 Then
                Dim WndRect(0 To 1) As RECT
                LSet WndRect(0) = VBFlexGridComboListRect
                MapWindowPoints VBFlexGridHandle, HWND_DESKTOP, WndRect(0), 2
                GetWindowRect VBFlexGridComboListHandle, WndRect(1)
                Dim hMonitor As Long, MI As MONITORINFO
                hMonitor = MonitorFromWindow(VBFlexGridEditHandle, MONITOR_DEFAULTTOPRIMARY)
                MI.cbSize = LenB(MI)
                GetMonitorInfo hMonitor, MI
                If (WndRect(0).Bottom + (WndRect(1).Bottom - WndRect(1).Top)) > MI.RCMonitor.Bottom Then
                    SetWindowPos VBFlexGridComboListHandle, 0, WndRect(0).Left, WndRect(0).Top - (WndRect(1).Bottom - WndRect(1).Top), 0, 0, SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_SHOWWINDOW
                Else
                    SetWindowPos VBFlexGridComboListHandle, 0, WndRect(0).Left, WndRect(0).Bottom, 0, 0, SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_SHOWWINDOW
                End If
            End If
            SetCapture VBFlexGridComboListHandle
        End If
    Else
        If (dwLong And ODS_SELECTED) = ODS_SELECTED Then
            SetWindowLong VBFlexGridComboButtonHandle, GWL_USERDATA, dwLong And Not ODS_SELECTED
            InvalidateRect VBFlexGridComboButtonHandle, ByVal 0&, 0
            If GetCapture() = VBFlexGridComboListHandle Then ReleaseCapture
            If IsWindowVisible(VBFlexGridComboListHandle) <> 0 Then ShowWindow VBFlexGridComboListHandle, SW_HIDE
            RaiseEvent ComboCloseUp
        End If
    End If
End If
End Sub

Private Sub ComboButtonPerformClick()
If VBFlexGridEditHandle <> 0 And VBFlexGridComboButtonHandle <> 0 And VBFlexGridComboListHandle = 0 Then
    Dim dwLong As Long
    dwLong = GetWindowLong(VBFlexGridComboButtonHandle, GWL_USERDATA)
    If Not (dwLong And ODS_DISABLED) = ODS_DISABLED Then
        If Not (dwLong And ODS_SELECTED) = ODS_SELECTED Then
            SetWindowLong VBFlexGridComboButtonHandle, GWL_USERDATA, dwLong Or ODS_SELECTED
            InvalidateRect VBFlexGridComboButtonHandle, ByVal 0&, 0
        End If
        VBFlexGridComboButtonClick = True
        RaiseEvent ComboButtonClick
        VBFlexGridComboButtonClick = False
        If VBFlexGridEditHandle <> 0 Then
            If VBFlexGridComboButtonHandle <> 0 Then Call ComboButtonSetState(ODS_SELECTED, False)
            SetFocusAPI VBFlexGridEditHandle
        Else
            SetFocusAPI UserControl.hWnd
        End If
    End If
End If
End Sub

Private Sub ComboButtonDrawEllipsis(ByVal hDC As Long, ByRef ContentRect As RECT)
Dim OldBkMode As Long, OldTextAlign As Long, hFontOld As Long
Dim X As Long, Y As Long, Size As SIZEAPI, Result As Long, DX(0 To 2) As Long
OldBkMode = SetBkMode(hDC, 1)
OldTextAlign = SetTextAlign(hDC, TA_CENTER Or TA_BASELINE)
hFontOld = SelectObject(hDC, GetStockObject(SYSTEM_FONT))
With ContentRect
X = .Left + ((.Right - .Left) / 2)
Y = .Bottom
GetTextExtentPoint32 hDC, ByVal StrPtr("."), 1, Size
Result = (((.Bottom - .Top) - (Size.CY / 2)) / 2)
If Result > 0 Then Y = Y - Result
End With
' The system font is not scaled on higher DPI's.
' For better appearance the dots will be shorten always by 1 unit and then scaled, if necessary.
DX(0) = (Size.CX - 1) * PixelsPerDIP_X()
DX(1) = DX(0)
DX(2) = DX(0)
ExtTextOut hDC, X, Y, ETO_CLIPPED, ContentRect, StrPtr("..."), 3, VarPtr(DX(0))
SetBkMode hDC, OldBkMode
SetTextAlign hDC, OldTextAlign
If hFontOld <> 0 Then SelectObject hDC, hFontOld
End Sub

Private Function ComboButtonGetState(ByVal dwState As Long) As Boolean
If VBFlexGridEditHandle <> 0 And VBFlexGridComboButtonHandle <> 0 Then ComboButtonGetState = CBool((GetWindowLong(VBFlexGridComboButtonHandle, GWL_USERDATA) And dwState) = dwState)
End Function

Private Sub ComboButtonSetState(ByVal dwState As Long, ByVal Value As Boolean)
If VBFlexGridEditHandle <> 0 And VBFlexGridComboButtonHandle <> 0 Then
    Dim dwLong As Long
    dwLong = GetWindowLong(VBFlexGridComboButtonHandle, GWL_USERDATA)
    If Value = True Then
        If Not (dwLong And dwState) = dwState Then
            SetWindowLong VBFlexGridComboButtonHandle, GWL_USERDATA, dwLong Or dwState
            InvalidateRect VBFlexGridComboButtonHandle, ByVal 0&, 0
        End If
    Else
        If (dwLong And dwState) = dwState Then
            SetWindowLong VBFlexGridComboButtonHandle, GWL_USERDATA, dwLong And Not dwState
            InvalidateRect VBFlexGridComboButtonHandle, ByVal 0&, 0
        End If
    End If
End If
End Sub

Private Function ComboListSelFromPt(ByVal X As Long, ByVal Y As Long) As Long
ComboListSelFromPt = LB_ERR
If VBFlexGridComboListHandle <> 0 Then
    Dim P As POINTAPI, Index As Long
    P.X = X
    P.Y = Y
    ClientToScreen VBFlexGridComboListHandle, P
    Index = LBItemFromPt(VBFlexGridComboListHandle, P.X, P.Y, 0)
    If Not Index = LB_ERR Then
        If Index <> SendMessage(VBFlexGridComboListHandle, LB_GETCURSEL, 0, ByVal 0&) Then SendMessage VBFlexGridComboListHandle, LB_SETCURSEL, Index, ByVal 0&
    End If
    ComboListSelFromPt = Index
End If
End Function

Private Sub ComboListCommitSel()
If VBFlexGridEditHandle <> 0 And VBFlexGridComboListHandle <> 0 Then
    Dim Index As Long, Length As Long
    Index = SendMessage(VBFlexGridComboListHandle, LB_GETCURSEL, 0, ByVal 0&)
    Length = SendMessage(VBFlexGridComboListHandle, LB_GETTEXTLEN, Index, ByVal 0&)
    If Not Length = LB_ERR Then
        Dim Text As String
        Text = String(Length, vbNullChar)
        SendMessage VBFlexGridComboListHandle, LB_GETTEXT, Index, ByVal StrPtr(Text)
        Me.EditText = Text
        SendMessage VBFlexGridEditHandle, EM_SETSEL, 0, ByVal -1&
    End If
End If
End Sub

Private Sub InplaceMergeSort(ByVal Left As Long, ByVal Middle As Long, ByVal Right As Long, ByVal Col As Long, ByRef Data() As TCOLS, ByVal Sort As FlexSortConstants)
Dim Temp() As TCOLS, Cmp As Long, Dst As Long
Dim i As Long, j As Long
Dim Dbl1 As Double, Dbl2 As Double
ReDim Temp(Middle - Left) As TCOLS
j = 0
For i = Left To Middle
    LSet Temp(j) = Data(i)
    j = j + 1
Next i
j = 0
Dst = Left
Do While i <= Right And j <= UBound(Temp)
    Cmp = Empty
    Select Case Sort
        Case FlexSortGenericAscending, FlexSortGenericDescending
            If Not IsNumeric(Data(i).Cols(Col).Text) Or Not IsNumeric(Temp(j).Cols(Col).Text) Then
                If Data(i).Cols(Col).Text < Temp(j).Cols(Col).Text Then
                    Cmp = -1
                ElseIf Data(i).Cols(Col).Text > Temp(j).Cols(Col).Text Then
                    Cmp = 1
                End If
            Else
                Dbl1 = Empty: Dbl2 = Empty
                On Error Resume Next
                Dbl1 = CDbl(Data(i).Cols(Col).Text)
                Dbl2 = CDbl(Temp(j).Cols(Col).Text)
                On Error GoTo 0
                Cmp = Sgn(Dbl1 - Dbl2)
            End If
            If Sort = FlexSortGenericDescending Then Cmp = -Cmp
        Case FlexSortNumericAscending, FlexSortNumericDescending
            Dbl1 = Empty: Dbl2 = Empty
            On Error Resume Next
            Dbl1 = CDbl(Data(i).Cols(Col).Text)
            Dbl2 = CDbl(Temp(j).Cols(Col).Text)
            On Error GoTo 0
            Cmp = Sgn(Dbl1 - Dbl2)
            If Sort = FlexSortNumericDescending Then Cmp = -Cmp
        Case FlexSortStringNoCaseAscending, FlexSortStringNoCaseDescending
            Cmp = lstrcmpi(StrPtr(Data(i).Cols(Col).Text), StrPtr(Temp(j).Cols(Col).Text))
            If Sort = FlexSortStringNoCaseDescending Then Cmp = -Cmp
        Case FlexSortStringAscending, FlexSortStringDescending
            Cmp = lstrcmp(StrPtr(Data(i).Cols(Col).Text), StrPtr(Temp(j).Cols(Col).Text))
            If Sort = FlexSortStringDescending Then Cmp = -Cmp
        Case FlexSortCurrencyAscending, FlexSortCurrencyDescending
            Dim Cur1 As Currency, Cur2 As Currency
            Cur1 = Empty: Cur2 = Empty
            On Error Resume Next
            Cur1 = CCur(Data(i).Cols(Col).Text)
            Cur2 = CCur(Temp(j).Cols(Col).Text)
            On Error GoTo 0
            Cmp = Sgn(Cur1 - Cur2)
            If Sort = FlexSortCurrencyDescending Then Cmp = -Cmp
        Case FlexSortDateAscending, FlexSortDateDescending
            Dim Date1 As Date, Date2 As Date
            Date1 = Empty: Date2 = Empty
            On Error Resume Next
            Date1 = CDate(Data(i).Cols(Col).Text)
            Date2 = CDate(Temp(j).Cols(Col).Text)
            On Error GoTo 0
            Cmp = Sgn(Date1 - Date2)
            If Sort = FlexSortDateDescending Then Cmp = -Cmp
    End Select
    If Cmp < 0 Then
        LSet Data(Dst) = Data(i)
        i = i + 1
    Else
        LSet Data(Dst) = Temp(j)
        j = j + 1
    End If
    Dst = Dst + 1
Loop
Do While j <= UBound(Temp)
    LSet Data(Dst) = Temp(j)
    Dst = Dst + 1
    j = j + 1
Loop
End Sub

Private Sub MergeSortRec(ByVal Left As Long, ByVal Right As Long, ByVal Col As Long, ByRef Data() As TCOLS, ByVal Sort As FlexSortConstants)
Dim Middle As Long
Middle = (Left + Right) \ 2
If Left < Right Then
    Call MergeSortRec(Left, Middle, Col, Data(), Sort)
    Call MergeSortRec(Middle + 1, Right, Col, Data(), Sort)
    Call InplaceMergeSort(Left, Middle, Right, Col, Data(), Sort)
End If
End Sub

Private Sub BubbleSortIter(ByVal First As Long, ByVal Last As Long, ByVal Col As Long, ByRef Data() As TCOLS)
Dim Swap As TCOLS, Cmp As Long
Dim i As Long, j As Long
Do While Last > First
    i = First
    For j = First To Last - 1
        Cmp = Empty
        RaiseEvent Compare(j, j + 1, Col, Cmp)
        If Cmp > 0 Then
            LSet Swap = Data(j + 1)
            LSet Data(j + 1) = Data(j)
            LSet Data(j) = Swap
            i = j
        End If
    Next j
    Last = i
Loop
End Sub

#If ImplementPreTranslateMsg = True Then

Private Function PreTranslateMsg(ByVal lParam As Long) As Long
PreTranslateMsg = 0
If lParam <> 0 Then
    Dim Msg As TMSG, Handled As Boolean, RetVal As Long
    CopyMemory Msg, ByVal lParam, LenB(Msg)
    IOleInPlaceActiveObjectVB_TranslateAccelerator Handled, RetVal, Msg.hWnd, Msg.Message, Msg.wParam, Msg.lParam, GetShiftStateFromMsg()
    If Handled = True Then
        PreTranslateMsg = 1
    ElseIf PropWantReturn = True Then
        If Msg.Message = WM_KEYDOWN Or Msg.Message = WM_KEYUP Then
            If (Msg.wParam And &HFF&) = vbKeyReturn Then
                SendMessage Msg.hWnd, Msg.Message, Msg.wParam, ByVal Msg.lParam
                PreTranslateMsg = 1
            End If
        End If
    End If
End If
End Function

#End If

Friend Function FSubclass_Message(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long
Select Case dwRefData
    Case 1
        FSubclass_Message = WindowProcControl(hWnd, wMsg, wParam, lParam)
    Case 2
        FSubclass_Message = WindowProcEdit(hWnd, wMsg, wParam, lParam)
    Case 3
        FSubclass_Message = WindowProcComboButton(hWnd, wMsg, wParam, lParam)
    Case 4
        FSubclass_Message = WindowProcComboList(hWnd, wMsg, wParam, lParam)
    Case 5
        FSubclass_Message = WindowProcUserControl(hWnd, wMsg, wParam, lParam)
End Select
End Function

Private Function WindowProcControl(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim HTI As THITTESTINFO, Pos As Long, Cancel As Boolean
Select Case wMsg
    Case WM_SETFOCUS
        If wParam <> UserControl.hWnd Then SetFocusAPI UserControl.hWnd: Exit Function
        If VBFlexGridEditHandle <> 0 Then SetFocusAPI VBFlexGridEditHandle: Exit Function
        
        #If ImplementPreTranslateMsg = True Then
        
        If VBFlexGridUsePreTranslateMsg = False Then Call ActivateIPAO(Me)
        
        #Else
        
        Call ActivateIPAO(Me)
        
        #End If
        
    Case WM_KILLFOCUS
        
        #If ImplementPreTranslateMsg = True Then
        
        If VBFlexGridUsePreTranslateMsg = False Then Call DeActivateIPAO
        
        #Else
        
        Call DeActivateIPAO
        
        #End If
        
    Case WM_GETFONT
        WindowProcControl = VBFlexGridFontHandle
        Exit Function
    Case WM_SETREDRAW
        VBFlexGridNoRedraw = CBool(wParam = 0)
        WindowProcControl = 0
        Exit Function
    Case WM_SIZE
        If VBFlexGridDoubleBufferDC <> 0 Then
            If VBFlexGridDoubleBufferBmpOld <> 0 Then
                SelectObject VBFlexGridDoubleBufferDC, VBFlexGridDoubleBufferBmpOld
                VBFlexGridDoubleBufferBmpOld = 0
            End If
            If VBFlexGridDoubleBufferBmp <> 0 Then
                DeleteObject VBFlexGridDoubleBufferBmp
                VBFlexGridDoubleBufferBmp = 0
            End If
            DeleteDC VBFlexGridDoubleBufferDC
            VBFlexGridDoubleBufferDC = 0
        End If
        GetClientRect hWnd, VBFlexGridClientRect
        Dim RCP As TROWCOLPARAMS
        With RCP
        .Mask = RCPM_TOPROW Or RCPM_LEFTCOL
        .Flags = RCPF_CHECKTOPROW Or RCPF_CHECKLEFTCOL Or RCPF_SETSCROLLBARS
        .TopRow = VBFlexGridTopRow
        .LeftCol = VBFlexGridLeftCol
        Call SetRowColParams(RCP)
        End With
        If PropShowInfoTips = True Or PropShowLabelTips = True Then Call UpdateToolTipRect
    Case WM_HSCROLL, WM_VSCROLL
        Dim dwStyle As Long
        dwStyle = GetWindowLong(hWnd, GWL_STYLE)
        If lParam = 0 And ((wMsg = WM_HSCROLL And (dwStyle And WS_HSCROLL) = WS_HSCROLL) Or (wMsg = WM_VSCROLL And (dwStyle And WS_VSCROLL) = WS_VSCROLL)) Then
            Dim SCI As SCROLLINFO, wBar As Long, PrevPos As Long
            SCI.cbSize = LenB(SCI)
            SCI.fMask = SIF_ALL
            If wMsg = WM_HSCROLL Then
                wBar = SB_HORZ
            ElseIf wMsg = WM_VSCROLL Then
                wBar = SB_VERT
            End If
            GetScrollInfo hWnd, wBar, SCI
            PrevPos = SCI.nPos
            Select Case LoWord(wParam)
                Case SB_LINELEFT, SB_LINEUP
                    If wMsg = WM_HSCROLL Then
                        SCI.nPos = VBFlexGridLeftCol
                        Call MovePreviousCol(SCI.nPos)
                        If SCI.nPos < VBFlexGridLeftCol Then
                            SCI.nPos = SCI.nPos - (PropFixedCols + PropFrozenCols)
                        Else
                            SCI.nPos = SCI.nMin
                        End If
                    ElseIf wMsg = WM_VSCROLL Then
                        SCI.nPos = VBFlexGridTopRow
                        Call MovePreviousRow(SCI.nPos)
                        If SCI.nPos < VBFlexGridTopRow Then
                            SCI.nPos = SCI.nPos - (PropFixedRows + PropFrozenRows)
                        Else
                            SCI.nPos = SCI.nMin
                        End If
                    End If
                Case SB_LINERIGHT, SB_LINEDOWN
                    If wMsg = WM_HSCROLL Then
                        SCI.nPos = VBFlexGridLeftCol
                        Call MoveNextCol(SCI.nPos)
                        If SCI.nPos > VBFlexGridLeftCol Then
                            SCI.nPos = SCI.nPos - (PropFixedCols + PropFrozenCols)
                        Else
                            SCI.nPos = SCI.nMax
                        End If
                    ElseIf wMsg = WM_VSCROLL Then
                        SCI.nPos = VBFlexGridTopRow
                        Call MoveNextRow(SCI.nPos)
                        If SCI.nPos > VBFlexGridTopRow Then
                            SCI.nPos = SCI.nPos - (PropFixedRows + PropFrozenRows)
                        Else
                            SCI.nPos = SCI.nMax
                        End If
                    End If
                Case SB_PAGELEFT, SB_PAGEUP
                    If SCI.nPage = 0 Then
                        If wBar = SB_HORZ Then
                            SCI.nPos = SCI.nPos - GetColsPerPageRev(VBFlexGridLeftCol)
                        ElseIf wBar = SB_VERT Then
                            SCI.nPos = SCI.nPos - GetRowsPerPageRev(VBFlexGridTopRow)
                        End If
                    Else
                        SCI.nPos = SCI.nPos - SCI.nPage
                    End If
                Case SB_PAGERIGHT, SB_PAGEDOWN
                    If SCI.nPage = 0 Then
                        If wBar = SB_HORZ Then
                            SCI.nPos = SCI.nPos + GetColsPerPage(VBFlexGridLeftCol)
                        ElseIf wBar = SB_VERT Then
                            SCI.nPos = SCI.nPos + GetRowsPerPage(VBFlexGridTopRow)
                        End If
                    Else
                        SCI.nPos = SCI.nPos + SCI.nPage
                    End If
                Case SB_THUMBPOSITION
                    SCI.nPos = SCI.nTrackPos
                Case SB_THUMBTRACK
                    If PropScrollTrack = True Then SCI.nPos = SCI.nTrackPos
                Case SB_TOP
                    SCI.nPos = SCI.nMin
                Case SB_BOTTOM
                    SCI.nPos = SCI.nMax
            End Select
            If SCI.nPos > SCI.nMax Then
                SCI.nPos = SCI.nMax
            ElseIf SCI.nPos < SCI.nMin Then
                SCI.nPos = SCI.nMin
            End If
            If PrevPos <> SCI.nPos Then
                SCI.fMask = SIF_POS
                SetScrollInfo hWnd, wBar, SCI, 1
                If wMsg = WM_HSCROLL Then
                    VBFlexGridLeftCol = (PropFixedCols + PropFrozenCols) + SCI.nPos
                ElseIf wMsg = WM_VSCROLL Then
                    VBFlexGridTopRow = (PropFixedRows + PropFrozenRows) + SCI.nPos
                End If
                Call RedrawGrid
                If PropShowInfoTips = True Or PropShowLabelTips = True Then
                    Pos = GetMessagePos()
                    Call CheckToolTipRowCol(Get_X_lParam(Pos), Get_Y_lParam(Pos))
                End If
                If VBFlexGridEditRow > -1 And VBFlexGridEditCol > -1 Then Call UpdateEditRect
                RaiseEvent Scroll
            End If
            WindowProcControl = 0
            Exit Function
        End If
    Case WM_PAINT
        If wParam = 0 Then
            Dim PS As PAINTSTRUCT, hDC As Long, hRgn As Long
            hDC = BeginPaint(hWnd, PS)
            With PS
            If PropDoubleBuffer = True Then
                If VBFlexGridDoubleBufferDC = 0 Then
                    VBFlexGridDoubleBufferDC = CreateCompatibleDC(hDC)
                    If VBFlexGridDoubleBufferDC <> 0 Then
                        VBFlexGridDoubleBufferBmp = CreateCompatibleBitmap(hDC, VBFlexGridClientRect.Right - VBFlexGridClientRect.Left, VBFlexGridClientRect.Bottom - VBFlexGridClientRect.Top)
                        If VBFlexGridDoubleBufferBmp <> 0 Then VBFlexGridDoubleBufferBmpOld = SelectObject(VBFlexGridDoubleBufferDC, VBFlexGridDoubleBufferBmp)
                    End If
                End If
                If VBFlexGridDoubleBufferDC <> 0 And VBFlexGridDoubleBufferBmp <> 0 Then
                    If .fErase <> 0 Then
                        If VBFlexGridBackColorBkgBrush <> 0 Then FillRect VBFlexGridDoubleBufferDC, VBFlexGridClientRect, VBFlexGridBackColorBkgBrush
                        Call DrawGrid(VBFlexGridDoubleBufferDC, -1)
                    Else
                        Call DrawGrid(VBFlexGridDoubleBufferDC, hRgn)
                        If hRgn <> 0 Then ExtSelectClipRgn hDC, hRgn, RGN_COPY
                    End If
                    With PS.RCPaint
                    BitBlt hDC, .Left, .Top, .Right - .Left, .Bottom - .Top, VBFlexGridDoubleBufferDC, .Left, .Top, vbSrcCopy
                    End With
                    If hRgn <> 0 Then
                        ExtSelectClipRgn hDC, 0, RGN_COPY
                        DeleteObject hRgn
                    End If
                End If
            Else
                If .fErase <> 0 Then
                    Call DrawGrid(hDC, hRgn)
                    If hRgn <> 0 Then ExtSelectClipRgn hDC, hRgn, RGN_DIFF
                    If VBFlexGridBackColorBkgBrush <> 0 Then FillRect hDC, VBFlexGridClientRect, VBFlexGridBackColorBkgBrush
                Else
                    Call DrawGrid(hDC, -1)
                End If
                If hRgn <> 0 Then
                    ExtSelectClipRgn hDC, 0, RGN_COPY
                    DeleteObject hRgn
                End If
            End If
            End With
            EndPaint hWnd, PS
        Else
            Dim hDCBmp As Long, hBmp As Long, hBmpOld As Long
            hDCBmp = CreateCompatibleDC(wParam)
            If hDCBmp <> 0 Then
                hBmp = CreateCompatibleBitmap(wParam, VBFlexGridClientRect.Right - VBFlexGridClientRect.Left, VBFlexGridClientRect.Bottom - VBFlexGridClientRect.Top)
                If hBmp <> 0 Then
                    hBmpOld = SelectObject(hDCBmp, hBmp)
                    If SendMessage(hWnd, WM_ERASEBKGND, hDCBmp, ByVal 0&) = 0 Then
                        If VBFlexGridBackColorBkgBrush <> 0 Then FillRect hDCBmp, VBFlexGridClientRect, VBFlexGridBackColorBkgBrush
                    End If
                    Call DrawGrid(hDCBmp, -1)
                    BitBlt wParam, 0, 0, VBFlexGridClientRect.Right - VBFlexGridClientRect.Left, VBFlexGridClientRect.Bottom - VBFlexGridClientRect.Top, hDCBmp, 0, 0, vbSrcCopy
                    SelectObject hDCBmp, hBmpOld
                    DeleteObject hBmp
                End If
                DeleteDC hDCBmp
            End If
        End If
        WindowProcControl = 0
        Exit Function
    Case WM_PRINTCLIENT
        SendMessage hWnd, WM_PAINT, wParam, ByVal lParam
        WindowProcControl = 0
        Exit Function
    Case WM_MOUSEACTIVATE
        If VBFlexGridEditRow > -1 And VBFlexGridEditCol > -1 Then
            If ValidateEditOnMouseActivateMsg(lParam, WindowProcControl) = True Then
                ' In case the edit window is still active due to failed validation then this ensures that the focus is properly set when clicked from outside.
                If VBFlexGridEditHandle <> 0 Then
                    If GetFocus() <> VBFlexGridEditHandle Then SetFocusAPI UserControl.hWnd
                End If
                Exit Function
            End If
        End If
    Case WM_SETCURSOR
        If LoWord(lParam) = HTCLIENT Then
            With HTI
            Pos = GetMessagePos()
            .PT.X = Get_X_lParam(Pos)
            .PT.Y = Get_Y_lParam(Pos)
            ScreenToClient hWnd, .PT
            Call GetHitTestInfo(HTI)
            Select Case .HitResult
                Case FlexHitResultDividerRowTop, FlexHitResultDividerRowBottom
                    SetCursor LoadCursor(0, MousePointerID(vbSizeNS))
                    WindowProcControl = 1
                    Exit Function
                Case FlexHitResultDividerColumnLeft, FlexHitResultDividerColumnRight
                    SetCursor LoadCursor(0, MousePointerID(vbSizeWE))
                    WindowProcControl = 1
                    Exit Function
            End Select
            End With
            If MousePointerID(PropMousePointer) <> 0 Then
                SetCursor LoadCursor(0, MousePointerID(PropMousePointer))
                WindowProcControl = 1
                Exit Function
            ElseIf PropMousePointer = 99 Then
                If Not PropMouseIcon Is Nothing Then
                    SetCursor PropMouseIcon.Handle
                    WindowProcControl = 1
                    Exit Function
                End If
            End If
        End If
    Case WM_SETTINGCHANGE
        SystemParametersInfo SPI_GETWHEELSCROLLLINES, 0, VBFlexGridWheelScrollLines, 0
        If SystemParametersInfo(SPI_GETFOCUSBORDERWIDTH, 0, VBFlexGridFocusBorder.CX, 0) = 0 Then VBFlexGridFocusBorder.CX = 1
        If SystemParametersInfo(SPI_GETFOCUSBORDERHEIGHT, 0, VBFlexGridFocusBorder.CY, 0) = 0 Then VBFlexGridFocusBorder.CY = 1
    Case WM_STYLECHANGED
        If wParam = GWL_EXSTYLE Then
            Dim dwStyleNew As Long
            CopyMemory dwStyleNew, ByVal UnsignedAdd(lParam, 4), 4
            VBFlexGridRTLLayout = CBool((dwStyleNew And WS_EX_LAYOUTRTL) = WS_EX_LAYOUTRTL)
            VBFlexGridRTLReading = CBool((dwStyleNew And WS_EX_RTLREADING) = WS_EX_RTLREADING)
        End If
    Case WM_MOUSEWHEEL
        If VBFlexGridWheelScrollLines > 0 Then
            Static WheelDelta As Long, LastWheelDelta As Long
            If Sgn(HiWord(wParam)) <> Sgn(LastWheelDelta) Then WheelDelta = 0
            WheelDelta = WheelDelta + HiWord(wParam)
            If Abs(WheelDelta) >= 120 Then
                Dim WheelDeltaPerLine As Long
                WheelDeltaPerLine = (WheelDelta / VBFlexGridWheelScrollLines)
                If Sgn(WheelDelta) = -1 Then
                    While WheelDelta <= WheelDeltaPerLine
                        SendMessage hWnd, WM_VSCROLL, MakeDWord(SB_LINEDOWN, 0), ByVal 0&
                        WheelDelta = WheelDelta - WheelDeltaPerLine
                    Wend
                Else
                    While WheelDelta >= WheelDeltaPerLine
                        SendMessage hWnd, WM_VSCROLL, MakeDWord(SB_LINEUP, 0), ByVal 0&
                        WheelDelta = WheelDelta - WheelDeltaPerLine
                    Wend
                End If
                WheelDelta = 0
            End If
            LastWheelDelta = HiWord(wParam)
            WindowProcControl = 0
            Exit Function
        End If
    Case WM_KEYDOWN, WM_KEYUP, WM_SYSKEYDOWN, WM_SYSKEYUP
        Dim KeyCode As Integer
        KeyCode = wParam And &HFF&
        If wMsg = WM_KEYDOWN Or wMsg = WM_KEYUP Then
            If wMsg = WM_KEYDOWN Then
                RaiseEvent KeyDown(KeyCode, GetShiftStateFromMsg())
                If PropAllowUserEditing = True Then
                    Select Case KeyCode
                        Case vbKeyF2
                            If CreateEdit(FlexEditReasonF2) = True Then Exit Function
                        Case vbKeySpace
                            If CreateEdit(FlexEditReasonSpace) = True Then Exit Function
                        Case vbKeyBack
                            If CreateEdit(FlexEditReasonBackSpace) = True Then Exit Function
                    End Select
                End If
            ElseIf wMsg = WM_KEYUP Then
                RaiseEvent KeyUp(KeyCode, GetShiftStateFromMsg())
            End If
            Dim Msg As TMSG
            Const PM_NOREMOVE As Long = &H0
            If PeekMessage(Msg, hWnd, WM_CHAR, WM_CHAR, PM_NOREMOVE) <> 0 Then VBFlexGridCharCodeCache = Msg.wParam
            If wMsg = WM_KEYDOWN Then Call ProcessKeyDown(KeyCode, GetShiftStateFromMsg())
        ElseIf wMsg = WM_SYSKEYDOWN Then
            RaiseEvent KeyDown(KeyCode, GetShiftStateFromMsg())
        ElseIf wMsg = WM_SYSKEYUP Then
            RaiseEvent KeyUp(KeyCode, GetShiftStateFromMsg())
        End If
        wParam = KeyCode
    Case WM_CHAR
        Dim KeyChar As Integer
        If VBFlexGridCharCodeCache <> 0 Then
            KeyChar = CUIntToInt(VBFlexGridCharCodeCache And &HFFFF&)
            VBFlexGridCharCodeCache = 0
        Else
            KeyChar = CUIntToInt(wParam And &HFFFF&)
        End If
        RaiseEvent KeyPress(KeyChar)
        wParam = CIntToUInt(KeyChar)
        If PropAllowUserEditing = True Then
            If wParam >= 33 Then ' 0 to 31 are non-printable and 32 is space char
                If CreateEdit(FlexEditReasonKeyPress) = True Then
                    If VBFlexGridEditHandle <> 0 Then PostMessage VBFlexGridEditHandle, wMsg, wParam, ByVal 0&
                    Exit Function
                End If
            End If
        End If
    Case WM_UNICHAR
        If wParam = UNICODE_NOCHAR Then
            WindowProcControl = 1
        Else
            Dim UTF16 As String
            UTF16 = UTF32CodePoint_To_UTF16(wParam)
            If Len(UTF16) = 1 Then
                SendMessage hWnd, WM_CHAR, CIntToUInt(AscW(UTF16)), ByVal lParam
            ElseIf Len(UTF16) = 2 Then
                SendMessage hWnd, WM_CHAR, CIntToUInt(AscW(Left$(UTF16, 1))), ByVal lParam
                SendMessage hWnd, WM_CHAR, CIntToUInt(AscW(Right$(UTF16, 1))), ByVal lParam
            End If
            WindowProcControl = 0
        End If
        Exit Function
    Case WM_INPUTLANGCHANGE
        Call SetIMEMode(hWnd, VBFlexGridIMCHandle, PropIMEMode)
    Case WM_IME_SETCONTEXT
        If wParam <> 0 Then Call SetIMEMode(hWnd, VBFlexGridIMCHandle, PropIMEMode)
    Case WM_IME_CHAR
        SendMessage hWnd, WM_CHAR, wParam, ByVal lParam
        Exit Function
    Case WM_LBUTTONDOWN, WM_MBUTTONDOWN, WM_RBUTTONDOWN
        With HTI
        .PT.X = Get_X_lParam(lParam)
        .PT.Y = Get_Y_lParam(lParam)
        Call GetHitTestInfo(HTI)
        Select Case wMsg
            Case WM_LBUTTONDOWN
                RaiseEvent BeforeMouseDown(vbLeftButton, GetShiftStateFromParam(wParam), UserControl.ScaleX(.PT.X, vbPixels, vbTwips), UserControl.ScaleY(.PT.Y, vbPixels, vbTwips), Cancel)
                If Cancel = False Then
                    SetCapture hWnd
                    If GetFocus() <> hWnd Then SetFocusAPI UserControl.hWnd
                    Cancel = ProcessLButtonDown(GetShiftStateFromParam(wParam), HTI)
                End If
            Case WM_MBUTTONDOWN
                RaiseEvent BeforeMouseDown(vbMiddleButton, GetShiftStateFromParam(wParam), UserControl.ScaleX(.PT.X, vbPixels, vbTwips), UserControl.ScaleY(.PT.Y, vbPixels, vbTwips), Cancel)
            Case WM_RBUTTONDOWN
                RaiseEvent BeforeMouseDown(vbRightButton, GetShiftStateFromParam(wParam), UserControl.ScaleX(.PT.X, vbPixels, vbTwips), UserControl.ScaleY(.PT.Y, vbPixels, vbTwips), Cancel)
        End Select
        End With
        If Cancel = True Then
            VBFlexGridCellClickRow = -1
            VBFlexGridCellClickCol = -1
            WindowProcControl = 0
            Exit Function
        Else
            With HTI
            If .HitResult = FlexHitResultCell Then
                VBFlexGridCellClickRow = .HitRow
                VBFlexGridCellClickCol = .HitCol
            Else
                VBFlexGridCellClickRow = -1
                VBFlexGridCellClickCol = -1
            End If
            End With
        End If
    Case WM_MOUSEMOVE
        Call ProcessMouseMove(GetMouseStateFromParam(wParam), Get_X_lParam(lParam), Get_Y_lParam(lParam))
    Case WM_LBUTTONUP
        Call ProcessLButtonUp(Get_X_lParam(lParam), Get_Y_lParam(lParam))
        ReleaseCapture
    Case WM_CAPTURECHANGED
        VBFlexGridCaptureRow = -1
        VBFlexGridCaptureCol = -1
        VBFlexGridCaptureHitResult = FlexHitResultNoWhere
        VBFlexGridCaptureDividerRow = -1
        VBFlexGridCaptureDividerCol = -1
        VBFlexGridCaptureDividerDrag = False
        VBFlexGridMouseMoveRow = -1
        VBFlexGridMouseMoveCol = -1
        VBFlexGridMouseMoveChanged = False
    Case WM_NOTIFY
        Dim NM As NMHDR
        CopyMemory NM, ByVal lParam, LenB(NM)
        If NM.hWndFrom = VBFlexGridToolTipHandle And VBFlexGridToolTipHandle <> 0 Then
            Static ShowInfoTip As Boolean, LBLI As TLABELINFO
            Select Case NM.Code
                Case TTN_GETDISPINFO
                    Dim NMTTDI As NMTTDISPINFO
                    CopyMemory NMTTDI, ByVal lParam, LenB(NMTTDI)
                    ShowInfoTip = False
                    LBLI.Flags = 0
                    With NMTTDI
                    Dim Text As String
                    With HTI
                    Pos = GetMessagePos()
                    .PT.X = Get_X_lParam(Pos)
                    .PT.Y = Get_Y_lParam(Pos)
                    ScreenToClient hWnd, .PT
                    Call GetHitTestInfo(HTI)
                    If .HitRow > -1 And .HitCol > -1 Then
                        If PropShowLabelTips = True Then Call GetLabelInfo(.HitRow, .HitCol, LBLI)
                        If (LBLI.Flags And LBLI_VALID) = LBLI_VALID And Not (LBLI.Flags And LBLI_UNFOLDED) = LBLI_UNFOLDED Then
                            Call GetCellText(.HitRow, .HitCol, Text)
                            If (LBLI.DrawFlags And DT_SINGLELINE) = DT_SINGLELINE Then
                                If InStr(Text, vbCr) Then Text = Replace$(Text, vbCr, vbNullString)
                                If InStr(Text, vbLf) Then Text = Replace$(Text, vbLf, vbNullString)
                            End If
                        ElseIf PropShowInfoTips = True Then
                            Text = VBFlexGridCells.Rows(.HitRow).Cols(.HitCol).ToolTipText
                            ShowInfoTip = True
                        End If
                    End If
                    End With
                    If Not Text = vbNullString Then
                        If Len(Text) <= 80 Then
                            Text = Left$(Text & vbNullChar, 80)
                            CopyMemory .szText(0), ByVal StrPtr(Text), LenB(Text)
                        Else
                            .lpszText = StrPtr(Text)
                        End If
                        .hInst = 0
                        CopyMemory ByVal lParam, NMTTDI, LenB(NMTTDI)
                    Else
                        ShowInfoTip = False
                    End If
                    End With
                Case TTN_SHOW
                    If PropShowLabelTips = True And ShowInfoTip = False Then
                        If (LBLI.Flags And LBLI_VALID) = LBLI_VALID Then
                            Dim RC As RECT
                            LSet RC = LBLI.RC
                            MapWindowPoints VBFlexGridHandle, HWND_DESKTOP, RC, 2
                            SendMessage VBFlexGridToolTipHandle, TTM_ADJUSTRECT, 1, ByVal VarPtr(RC)
                            SetWindowPos VBFlexGridToolTipHandle, 0, RC.Left, RC.Top, 0, 0, SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_NOACTIVATE
                            WindowProcControl = 1
                            Exit Function
                        End If
                    End If
                Case NM_CUSTOMDRAW
                    Dim NMTTCD As NMTTCUSTOMDRAW
                    CopyMemory NMTTCD, ByVal lParam, LenB(NMTTCD)
                    Select Case NMTTCD.NMCD.dwDrawStage
                        Case CDDS_PREPAINT
                            If PropShowLabelTips = True And ShowInfoTip = False Then
                                If (LBLI.Flags And LBLI_VALID) = LBLI_VALID Then
                                    If (NMTTCD.uDrawFlags And DT_CALCRECT) = DT_CALCRECT Then
                                        NMTTCD.uDrawFlags = LBLI.DrawFlags
                                        If Not (NMTTCD.uDrawFlags And DT_CALCRECT) = DT_CALCRECT Then NMTTCD.uDrawFlags = NMTTCD.uDrawFlags Or DT_CALCRECT
                                    Else
                                        NMTTCD.uDrawFlags = LBLI.DrawFlags
                                        If (NMTTCD.uDrawFlags And DT_CALCRECT) = DT_CALCRECT Then NMTTCD.uDrawFlags = NMTTCD.uDrawFlags And Not DT_CALCRECT
                                    End If
                                    CopyMemory ByVal lParam, NMTTCD, LenB(NMTTCD)
                                    Exit Function
                                End If
                            End If
                    End Select
            End Select
        End If
    Case WM_COMMAND
        If lParam <> 0 Then
            Select Case HiWord(wParam)
                Case EN_CHANGE
                    If LoWord(wParam) = ID_EDITCHILD And lParam = VBFlexGridEditHandle And VBFlexGridEditHandle <> 0 Then
                        If VBFlexGridEditChangeFrozen = False Then
                            If VBFlexGridComboActiveMode = FlexComboModeEditable And VBFlexGridComboListHandle <> 0 Then
                                Dim Index As Long
                                Index = SendMessage(VBFlexGridComboListHandle, LB_FINDSTRINGEXACT, -1, ByVal StrPtr(Me.EditText))
                                If Not Index = LB_ERR Then
                                    SendMessage VBFlexGridComboListHandle, LB_SETCURSEL, Index, ByVal 0&
                                    Call ComboListCommitSel
                                End If
                            End If
                            VBFlexGridEditTextChanged = True
                            VBFlexGridEditAlreadyValidated = False
                            RaiseEvent EditChange
                        End If
                    End If
                Case STN_CLICKED
                    If LoWord(wParam) = ID_COMBOBUTTONCHILD And lParam = VBFlexGridComboButtonHandle And VBFlexGridComboButtonHandle <> 0 Then
                        If VBFlexGridComboListHandle <> 0 Then
                            Call ComboShowDropDown(True)
                        Else
                            Call ComboButtonPerformClick
                        End If
                    End If
                Case STN_ENABLE
                    If LoWord(wParam) = ID_COMBOBUTTONCHILD And lParam = VBFlexGridComboButtonHandle And VBFlexGridComboButtonHandle <> 0 Then Call ComboButtonSetState(ODS_DISABLED, False)
                Case STN_DISABLE
                    If LoWord(wParam) = ID_COMBOBUTTONCHILD And lParam = VBFlexGridComboButtonHandle And VBFlexGridComboButtonHandle <> 0 Then Call ComboButtonSetState(ODS_DISABLED, True)
                Case LBN_SELCHANGE
                    If LoWord(wParam) = 0 And lParam = VBFlexGridComboListHandle And VBFlexGridComboListHandle <> 0 Then Call ComboListCommitSel
            End Select
        End If
    Case WM_NOTIFYFORMAT
        Const NF_QUERY As Long = 3
        If wParam = VBFlexGridToolTipHandle And VBFlexGridToolTipHandle <> 0 And lParam = NF_QUERY Then
            Const NFR_ANSI As Long = 1
            Const NFR_UNICODE As Long = 2
            WindowProcControl = NFR_UNICODE
            Exit Function
        End If
    Case WM_CTLCOLOREDIT, WM_CTLCOLORSTATIC
        If lParam = VBFlexGridEditHandle And VBFlexGridEditHandle <> 0 Then
            If VBFlexGridEditBackColorBrush <> 0 Then
                SetBkColor wParam, WinColor(VBFlexGridEditBackColor)
                SetTextColor wParam, WinColor(VBFlexGridEditForeColor)
                WindowProcControl = VBFlexGridEditBackColorBrush
            Else
                SetBkColor wParam, WinColor(vbWindowBackground)
                SetTextColor wParam, WinColor(vbWindowText)
                WindowProcControl = GetSysColorBrush(COLOR_WINDOW)
            End If
            Exit Function
        End If
    Case WM_THEMECHANGED
        VBFlexGridEnabledVisualStyles = EnabledVisualStyles()
    Case WM_DRAWITEM
        Dim DIS As DRAWITEMSTRUCT
        CopyMemory DIS, ByVal lParam, LenB(DIS)
        If DIS.CtlType = ODT_STATIC And DIS.CtlID = ID_COMBOBUTTONCHILD And DIS.hWndItem = VBFlexGridComboButtonHandle And VBFlexGridComboButtonHandle <> 0 Then
            Dim Brush As Long
            If VBFlexGridEditBackColorBrush <> 0 Then
                Brush = VBFlexGridEditBackColorBrush
            Else
                Brush = GetSysColorBrush(COLOR_WINDOW)
            End If
            FillRect DIS.hDC, DIS.RCItem, Brush
            DIS.ItemState = GetWindowLong(DIS.hWndItem, GWL_USERDATA)
            If VBFlexGridComboButtonDrawMode = FlexComboButtonDrawModeNormal Then
                Dim OldTextColor As Long
                
                #If ImplementThemedComboButton = True Then
                
                Dim Theme As Long
                If VBFlexGridEnabledVisualStyles = True And PropVisualStyles = True Then
                    If VBFlexGridComboListHandle <> 0 Then
                        Theme = OpenThemeData(VBFlexGridHandle, StrPtr("ComboBox"))
                    Else
                        Theme = OpenThemeData(VBFlexGridHandle, StrPtr("Button"))
                    End If
                End If
                If Theme <> 0 Then
                    If VBFlexGridComboListHandle <> 0 Then
                        Dim ComboBoxPart As Long, ComboBoxState As Long
                        ComboBoxPart = CP_DROPDOWNBUTTON
                        If Not (DIS.ItemState And ODS_DISABLED) = ODS_DISABLED Then
                            If Not (DIS.ItemState And ODS_SELECTED) = ODS_SELECTED Then
                                If Not (DIS.ItemState And ODS_HOTLIGHT) = ODS_HOTLIGHT Then
                                    ComboBoxState = CBXS_NORMAL
                                Else
                                    ComboBoxState = CBXS_HOT
                                End If
                            Else
                                ComboBoxState = CBXS_PRESSED
                            End If
                        Else
                            ComboBoxState = CBXS_DISABLED
                        End If
                        If IsThemeBackgroundPartiallyTransparent(Theme, ComboBoxPart, ComboBoxState) <> 0 Then DrawThemeParentBackground DIS.hWndItem, DIS.hDC, DIS.RCItem
                        DrawThemeBackground Theme, DIS.hDC, ComboBoxPart, ComboBoxState, DIS.RCItem, DIS.RCItem
                    Else
                        Dim ButtonPart As Long, ButtonState As Long
                        ButtonPart = BP_PUSHBUTTON
                        If Not (DIS.ItemState And ODS_DISABLED) = ODS_DISABLED Then
                            If Not (DIS.ItemState And ODS_SELECTED) = ODS_SELECTED Then
                                If Not (DIS.ItemState And ODS_HOTLIGHT) = ODS_HOTLIGHT Then
                                    ButtonState = PBS_NORMAL
                                Else
                                    ButtonState = PBS_HOT
                                End If
                            Else
                                ButtonState = PBS_PRESSED
                            End If
                        Else
                            ButtonState = PBS_DISABLED
                        End If
                        If IsThemeBackgroundPartiallyTransparent(Theme, ButtonPart, ButtonState) <> 0 Then DrawThemeParentBackground DIS.hWndItem, DIS.hDC, DIS.RCItem
                        DrawThemeBackground Theme, DIS.hDC, ButtonPart, ButtonState, DIS.RCItem, DIS.RCItem
                        GetThemeBackgroundContentRect Theme, DIS.hDC, ButtonPart, ButtonState, DIS.RCItem, DIS.RCItem
                        OldTextColor = SetTextColor(DIS.hDC, WinColor(vbButtonText))
                        Call ComboButtonDrawEllipsis(DIS.hDC, DIS.RCItem)
                        SetTextColor DIS.hDC, OldTextColor
                    End If
                    CloseThemeData Theme
                Else
                    Dim CtlType As Long, Flags As Long
                    If VBFlexGridComboListHandle <> 0 Then
                        CtlType = DFC_SCROLL
                        Flags = DFCS_SCROLLCOMBOBOX
                    Else
                        CtlType = DFC_BUTTON
                        Flags = DFCS_BUTTONPUSH Or DFCS_ADJUSTRECT
                    End If
                    If (DIS.ItemState And ODS_SELECTED) = ODS_SELECTED Then Flags = Flags Or DFCS_PUSHED Or DFCS_FLAT
                    If (DIS.ItemState And ODS_DISABLED) = ODS_DISABLED Then Flags = Flags Or DFCS_INACTIVE
                    If (DIS.ItemState And ODS_HOTLIGHT) = ODS_HOTLIGHT Then Flags = Flags Or DFCS_HOT
                    DrawFrameControl DIS.hDC, DIS.RCItem, CtlType, Flags
                    If CtlType = DFC_BUTTON Then
                        If (Flags And DFCS_HOT) = DFCS_HOT Then
                            OldTextColor = SetTextColor(DIS.hDC, GetSysColor(COLOR_HOTLIGHT))
                        Else
                            OldTextColor = SetTextColor(DIS.hDC, WinColor(vbButtonText))
                        End If
                        Call ComboButtonDrawEllipsis(DIS.hDC, DIS.RCItem)
                        SetTextColor DIS.hDC, OldTextColor
                    End If
                End If
                
                #Else
                
                Dim CtlType As Long, Flags As Long
                If VBFlexGridComboListHandle <> 0 Then
                    CtlType = DFC_SCROLL
                    Flags = DFCS_SCROLLCOMBOBOX
                Else
                    CtlType = DFC_BUTTON
                    Flags = DFCS_BUTTONPUSH Or DFCS_ADJUSTRECT
                End If
                If (DIS.ItemState And ODS_SELECTED) = ODS_SELECTED Then Flags = Flags Or DFCS_PUSHED Or DFCS_FLAT
                If (DIS.ItemState And ODS_DISABLED) = ODS_DISABLED Then Flags = Flags Or DFCS_INACTIVE
                If (DIS.ItemState And ODS_HOTLIGHT) = ODS_HOTLIGHT Then Flags = Flags Or DFCS_HOT
                DrawFrameControl DIS.hDC, DIS.RCItem, CtlType, Flags
                If CtlType = DFC_BUTTON Then
                    If (Flags And DFCS_HOT) = DFCS_HOT Then
                        OldTextColor = SetTextColor(DIS.hDC, GetSysColor(COLOR_HOTLIGHT))
                    Else
                        OldTextColor = SetTextColor(DIS.hDC, WinColor(vbButtonText))
                    End If
                    Call ComboButtonDrawEllipsis(DIS.hDC, DIS.RCItem)
                    SetTextColor DIS.hDC, OldTextColor
                End If
                
                #End If
                
            Else
                With DIS
                RaiseEvent ComboButtonOwnerDraw(.ItemAction, .ItemState, .hDC, .RCItem.Left, .RCItem.Top, .RCItem.Right, .RCItem.Bottom)
                End With
            End If
            WindowProcControl = 1
            Exit Function
        End If
    
    #If ImplementPreTranslateMsg = True Then
    
    Case UM_PRETRANSLATEMSG
        WindowProcControl = PreTranslateMsg(lParam)
        Exit Function
    
    #End If
    
End Select
WindowProcControl = DefWindowProc(hWnd, wMsg, wParam, lParam)
Select Case wMsg
    Case WM_SETFOCUS, WM_KILLFOCUS
        VBFlexGridFocused = CBool(wMsg = WM_SETFOCUS)
        Call RedrawGrid
    Case WM_LBUTTONDBLCLK, WM_MBUTTONDBLCLK, WM_RBUTTONDBLCLK
        With HTI
        Pos = GetMessagePos()
        .PT.X = Get_X_lParam(Pos)
        .PT.Y = Get_Y_lParam(Pos)
        ScreenToClient hWnd, .PT
        Call GetHitTestInfo(HTI)
        If wMsg = WM_LBUTTONDBLCLK Then
            Select Case .HitResult
                Case FlexHitResultDividerRowTop, FlexHitResultDividerRowBottom, FlexHitResultDividerColumnLeft, FlexHitResultDividerColumnRight
                    RaiseEvent DividerDblClick(.HitRowDivider, .HitColDivider)
                Case FlexHitResultCell
                    RaiseEvent CellDblClick(.HitRow, .HitCol, vbLeftButton)
            End Select
        Else
            If .HitResult = FlexHitResultCell Then
                If wMsg = WM_MBUTTONDBLCLK Then
                    RaiseEvent CellDblClick(.HitRow, .HitCol, vbMiddleButton)
                ElseIf wMsg = WM_RBUTTONDBLCLK Then
                    RaiseEvent CellDblClick(.HitRow, .HitCol, vbRightButton)
                End If
            End If
        End If
        RaiseEvent DblClick
        If PropAllowUserEditing = True Then
            If wMsg = WM_LBUTTONDBLCLK And .HitResult = FlexHitResultCell Then
                If .HitRow > (PropFixedRows - 1) And .HitCol > (PropFixedCols - 1) Then
                    Select Case PropSelectionMode
                        Case FlexSelectionModeByRow
                            CreateEdit FlexEditReasonDblClick, .HitRow
                        Case FlexSelectionModeByColumn
                            CreateEdit FlexEditReasonDblClick, , .HitCol
                        Case Else
                            CreateEdit FlexEditReasonDblClick, .HitRow, .HitCol
                    End Select
                End If
            End If
        End If
        End With
    Case WM_LBUTTONDOWN, WM_MBUTTONDOWN, WM_RBUTTONDOWN, WM_MOUSEMOVE, WM_LBUTTONUP, WM_MBUTTONUP, WM_RBUTTONUP
        Dim X As Single
        Dim Y As Single
        X = UserControl.ScaleX(Get_X_lParam(lParam), vbPixels, vbTwips)
        Y = UserControl.ScaleY(Get_Y_lParam(lParam), vbPixels, vbTwips)
        Select Case wMsg
            Case WM_LBUTTONDOWN
                RaiseEvent MouseDown(vbLeftButton, GetShiftStateFromParam(wParam), X, Y)
                VBFlexGridIsClick = True
            Case WM_MBUTTONDOWN
                RaiseEvent MouseDown(vbMiddleButton, GetShiftStateFromParam(wParam), X, Y)
                VBFlexGridIsClick = True
            Case WM_RBUTTONDOWN
                RaiseEvent MouseDown(vbRightButton, GetShiftStateFromParam(wParam), X, Y)
                VBFlexGridIsClick = True
            Case WM_MOUSEMOVE
                If VBFlexGridMouseOver = False And PropMouseTrack = True Then
                    VBFlexGridMouseOver = True
                    RaiseEvent MouseEnter
                    Dim TME As TRACKMOUSEEVENTSTRUCT
                    With TME
                    .cbSize = LenB(TME)
                    .hWndTrack = hWnd
                    .dwFlags = TME_LEAVE
                    End With
                    TrackMouseEvent TME
                End If
                RaiseEvent MouseMove(GetMouseStateFromParam(wParam), GetShiftStateFromParam(wParam), X, Y)
            Case WM_LBUTTONUP, WM_MBUTTONUP, WM_RBUTTONUP
                With HTI
                .PT.X = Get_X_lParam(lParam)
                .PT.Y = Get_Y_lParam(lParam)
                Call GetHitTestInfo(HTI)
                If .HitResult = FlexHitResultCell And VBFlexGridIsClick = True Then
                    If VBFlexGridCellClickRow = .HitRow And VBFlexGridCellClickCol = .HitCol Then
                        Select Case wMsg
                            Case WM_LBUTTONUP
                                RaiseEvent CellClick(.HitRow, .HitCol, vbLeftButton)
                            Case WM_MBUTTONUP
                                RaiseEvent CellClick(.HitRow, .HitCol, vbMiddleButton)
                            Case WM_RBUTTONUP
                                RaiseEvent CellClick(.HitRow, .HitCol, vbRightButton)
                        End Select
                    End If
                End If
                End With
                Select Case wMsg
                    Case WM_LBUTTONUP
                        RaiseEvent MouseUp(vbLeftButton, GetShiftStateFromParam(wParam), X, Y)
                    Case WM_MBUTTONUP
                        RaiseEvent MouseUp(vbMiddleButton, GetShiftStateFromParam(wParam), X, Y)
                    Case WM_RBUTTONUP
                        RaiseEvent MouseUp(vbRightButton, GetShiftStateFromParam(wParam), X, Y)
                End Select
                If VBFlexGridIsClick = True Then
                    VBFlexGridIsClick = False
                    If (X >= 0 And X <= UserControl.Width) And (Y >= 0 And Y <= UserControl.Height) Then RaiseEvent Click
                End If
        End Select
    Case WM_MOUSELEAVE
        If VBFlexGridMouseOver = True Then
            VBFlexGridMouseOver = False
            RaiseEvent MouseLeave
        End If
End Select
End Function

Private Function WindowProcEdit(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_SETFOCUS
        
        #If ImplementPreTranslateMsg = True Then
        
        If VBFlexGridUsePreTranslateMsg = False Then Call ActivateIPAO(Me)
        
        #Else
        
        Call ActivateIPAO(Me)
        
        #End If
        
    Case WM_KILLFOCUS
        
        #If ImplementPreTranslateMsg = True Then
        
        If VBFlexGridUsePreTranslateMsg = False Then Call DeActivateIPAO
        
        #Else
        
        Call DeActivateIPAO
        
        #End If
        
    Case WM_MOUSEACTIVATE
        ' It is necessary to break the chain and return MA_ACTIVATE for this window.
        ' This enables the parent window - when it receives WM_MOUSEACTIVATE - to destroy this child window.
        WindowProcEdit = MA_ACTIVATE
        Exit Function
    Case WM_MOUSEWHEEL
        If VBFlexGridComboListHandle <> 0 Then
            If ComboButtonGetState(ODS_SELECTED) = True Then
                SendMessage VBFlexGridComboListHandle, WM_MOUSEWHEEL, wParam, ByVal lParam
                Exit Function
            End If
        End If
    Case WM_KEYDOWN, WM_KEYUP, WM_SYSKEYDOWN, WM_SYSKEYUP
        Dim KeyCode As Integer
        KeyCode = wParam And &HFF&
        If wMsg = WM_KEYDOWN Or wMsg = WM_KEYUP Then
            If wMsg = WM_KEYDOWN Then
                If VBFlexGridEditRectChanged = True Then
                    VBFlexGridEditRectChanged = False
                    VBFlexGridEditRectChangedFrozen = True
                    Me.CellEnsureVisible , VBFlexGridEditMergedRange.TopRow, VBFlexGridEditMergedRange.LeftCol
                    VBFlexGridEditRectChangedFrozen = False
                End If
                RaiseEvent EditKeyDown(KeyCode, GetShiftStateFromMsg())
                If VBFlexGridEditHandle <> 0 Then
                    Select Case KeyCode
                        Case vbKeyEscape
                            If VBFlexGridComboButtonHandle <> 0 And VBFlexGridComboListHandle <> 0 Then Call ComboShowDropDown(False)
                            If DestroyEdit(True, FlexEditCloseModeEscape) = True Then Exit Function
                        Case vbKeyF4
                            If VBFlexGridComboButtonHandle <> 0 Then
                                If VBFlexGridComboListHandle <> 0 Then
                                    Call ComboShowDropDown(Not ComboButtonGetState(ODS_SELECTED))
                                Else
                                    Call ComboButtonPerformClick
                                End If
                                Exit Function
                            End If
                        Case vbKeyReturn
                            If VBFlexGridComboButtonHandle <> 0 And VBFlexGridComboListHandle <> 0 Then
                                If ComboButtonGetState(ODS_SELECTED) = True Then
                                    Call ComboListCommitSel
                                    Call ComboShowDropDown(False)
                                    DestroyEdit False, FlexEditCloseModeReturn
                                    Exit Function
                                End If
                            End If
                            If GetShiftStateFromMsg() = 0 Then
                                If DestroyEdit(False, FlexEditCloseModeReturn) = True Then Exit Function
                            Else
                                PostMessage hWnd, WM_CHAR, vbKeyReturn, ByVal 0&
                            End If
                        Case vbKeyTab
                            If PropTabBehavior <> FlexTabControls Then
                                Select Case GetShiftStateFromMsg()
                                    Case 0
                                        If DestroyEdit(False, FlexEditCloseModeTab) = True Then PostMessage VBFlexGridHandle, wMsg, wParam, ByVal 0&: Exit Function
                                    Case vbShiftMask
                                        If DestroyEdit(False, FlexEditCloseModeShiftTab) = True Then PostMessage VBFlexGridHandle, wMsg, wParam, ByVal 0&: Exit Function
                                End Select
                            End If
                        Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyPageDown, vbKeyPageUp, vbKeyHome, vbKeyEnd
                            If VBFlexGridComboButtonHandle <> 0 And VBFlexGridComboListHandle <> 0 Then
                                Select Case KeyCode
                                    Case vbKeyUp, vbKeyDown, vbKeyPageDown, vbKeyPageUp, vbKeyHome, vbKeyEnd
                                        SendMessage VBFlexGridComboListHandle, wMsg, wParam, ByVal lParam
                                        Exit Function
                                End Select
                            End If
                            Dim SelStart As Long, SelEnd As Long
                            SendMessage hWnd, EM_GETSEL, VarPtr(SelStart), ByVal VarPtr(SelEnd)
                            If SelStart = SelEnd Then
                                Dim CloseMode As FlexEditCloseModeConstants
                                CloseMode = -1
                                Select Case KeyCode
                                    Case vbKeyLeft, vbKeyRight, vbKeyPageDown, vbKeyPageUp
                                        Select Case KeyCode
                                            Case vbKeyLeft
                                                If SelEnd = 0 Then CloseMode = FlexEditCloseModeNavigationKey
                                            Case vbKeyRight
                                                If SelEnd = SendMessage(hWnd, WM_GETTEXTLENGTH, 0, ByVal 0&) Then CloseMode = FlexEditCloseModeNavigationKey
                                            Case vbKeyPageDown, vbKeyPageUp
                                                If SelStart = SelEnd Then CloseMode = FlexEditCloseModeNavigationKey
                                        End Select
                                    Case vbKeyUp, vbKeyDown, vbKeyHome, vbKeyEnd
                                        Dim FirstCharPos As Long, LineFromChar As Long
                                        FirstCharPos = SendMessage(hWnd, EM_LINEINDEX, -1, ByVal 0&)
                                        LineFromChar = SendMessage(hWnd, EM_LINEFROMCHAR, FirstCharPos, ByVal 0&)
                                        Select Case KeyCode
                                            Case vbKeyUp
                                                If LineFromChar = 0 Then CloseMode = FlexEditCloseModeNavigationKey
                                            Case vbKeyDown
                                                If LineFromChar = (SendMessage(hWnd, EM_GETLINECOUNT, 0, ByVal 0&) - 1) Then CloseMode = FlexEditCloseModeNavigationKey
                                            Case vbKeyHome
                                                If SelEnd = FirstCharPos Then CloseMode = FlexEditCloseModeNavigationKey
                                            Case vbKeyEnd
                                                If SelEnd = (FirstCharPos + SendMessage(hWnd, EM_LINELENGTH, FirstCharPos, ByVal 0&)) Then CloseMode = FlexEditCloseModeNavigationKey
                                        End Select
                                End Select
                                If CloseMode > -1 Then
                                    If DestroyEdit(False, CloseMode) = True Then PostMessage VBFlexGridHandle, wMsg, wParam, ByVal 0&: Exit Function
                                End If
                            End If
                    End Select
                Else
                    Exit Function
                End If
            ElseIf wMsg = WM_KEYUP Then
                RaiseEvent EditKeyUp(KeyCode, GetShiftStateFromMsg())
            End If
            Dim Msg As TMSG
            Const PM_NOREMOVE As Long = &H0
            If PeekMessage(Msg, hWnd, WM_CHAR, WM_CHAR, PM_NOREMOVE) <> 0 Then VBFlexGridCharCodeCache = Msg.wParam
        ElseIf wMsg = WM_SYSKEYDOWN Then
            If VBFlexGridEditRectChanged = True Then
                VBFlexGridEditRectChanged = False
                VBFlexGridEditRectChangedFrozen = True
                Me.CellEnsureVisible , VBFlexGridEditMergedRange.TopRow, VBFlexGridEditMergedRange.LeftCol
                VBFlexGridEditRectChangedFrozen = False
            End If
            RaiseEvent EditKeyDown(KeyCode, GetShiftStateFromMsg())
            If VBFlexGridEditHandle <> 0 Then
                If KeyCode = vbKeyReturn Then
                    PostMessage hWnd, WM_CHAR, vbKeyReturn, ByVal 0&
                ElseIf VBFlexGridComboButtonHandle <> 0 And VBFlexGridComboListHandle <> 0 Then
                    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then Call ComboShowDropDown(Not ComboButtonGetState(ODS_SELECTED))
                End If
            Else
                Exit Function
            End If
        ElseIf wMsg = WM_SYSKEYUP Then
            RaiseEvent EditKeyUp(KeyCode, GetShiftStateFromMsg())
        End If
        wParam = KeyCode
    Case WM_CHAR
        Dim KeyChar As Integer
        If VBFlexGridCharCodeCache <> 0 Then
            KeyChar = CUIntToInt(VBFlexGridCharCodeCache And &HFFFF&)
            VBFlexGridCharCodeCache = 0
        Else
            KeyChar = CUIntToInt(wParam And &HFFFF&)
        End If
        RaiseEvent EditKeyPress(KeyChar)
        If (wParam And &HFFFF&) <> 0 And KeyChar = 0 Then
            Exit Function
        Else
            wParam = CIntToUInt(KeyChar)
        End If
        If VBFlexGridComboActiveMode = FlexComboModeDropDown And VBFlexGridComboListHandle <> 0 Then
            SendMessage VBFlexGridComboListHandle, wMsg, wParam, ByVal lParam
            Exit Function
        End If
    Case WM_UNICHAR
        If wParam = UNICODE_NOCHAR Then
            WindowProcEdit = 1
        Else
            Dim UTF16 As String
            UTF16 = UTF32CodePoint_To_UTF16(wParam)
            If Len(UTF16) = 1 Then
                SendMessage hWnd, WM_CHAR, CIntToUInt(AscW(UTF16)), ByVal lParam
            ElseIf Len(UTF16) = 2 Then
                SendMessage hWnd, WM_CHAR, CIntToUInt(AscW(Left$(UTF16, 1))), ByVal lParam
                SendMessage hWnd, WM_CHAR, CIntToUInt(AscW(Right$(UTF16, 1))), ByVal lParam
            End If
            WindowProcEdit = 0
        End If
        Exit Function
    Case WM_INPUTLANGCHANGE
        Call SetIMEMode(hWnd, VBFlexGridIMCHandle, PropIMEMode)
    Case WM_IME_SETCONTEXT
        If wParam <> 0 Then Call SetIMEMode(hWnd, VBFlexGridIMCHandle, PropIMEMode)
    Case WM_IME_CHAR
        SendMessage hWnd, WM_CHAR, wParam, ByVal lParam
        Exit Function
    Case WM_CONTEXTMENU
        If wParam = hWnd Then
            Dim P As POINTAPI, Handled As Boolean
            P.X = Get_X_lParam(lParam)
            P.Y = Get_Y_lParam(lParam)
            If P.X = -1 And P.Y = -1 Then
                ' If the user types SHIFT + F10 then the X and Y coordinates are -1.
                RaiseEvent EditContextMenu(Handled, -1, -1)
            Else
                ScreenToClient VBFlexGridHandle, P
                RaiseEvent EditContextMenu(Handled, UserControl.ScaleX(P.X, vbPixels, vbContainerPosition), UserControl.ScaleY(P.Y, vbPixels, vbContainerPosition))
            End If
            If Handled = True Then Exit Function
        End If
    Case WM_LBUTTONDOWN
        If GetFocus() <> hWnd Then UCNoSetFocusFwd = True: SetFocusAPI UserControl.hWnd: UCNoSetFocusFwd = False
    Case WM_NCCALCSIZE, WM_NCHITTEST, WM_NCPAINT
        Dim RC As RECT
        Select Case wMsg
            Case WM_NCCALCSIZE
                Dim dwStyle As Long, dwExStyle As Long
                dwStyle = GetWindowLong(hWnd, GWL_STYLE)
                dwExStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
                If (dwStyle And WS_BORDER) = WS_BORDER Then dwStyle = dwStyle And Not WS_BORDER
                If (dwStyle And WS_DLGFRAME) = WS_DLGFRAME Then dwStyle = dwStyle And Not WS_DLGFRAME
                If (dwExStyle And WS_EX_STATICEDGE) = WS_EX_STATICEDGE Then dwExStyle = dwExStyle And Not WS_EX_STATICEDGE
                If (dwExStyle And WS_EX_CLIENTEDGE) = WS_EX_CLIENTEDGE Then dwExStyle = dwExStyle And Not WS_EX_CLIENTEDGE
                If (dwExStyle And WS_EX_WINDOWEDGE) = WS_EX_WINDOWEDGE Then dwExStyle = dwExStyle And Not WS_EX_WINDOWEDGE
                SetWindowLong hWnd, GWL_STYLE, dwStyle
                SetWindowLong hWnd, GWL_EXSTYLE, dwExStyle
                WindowProcEdit = FlexDefaultProc(hWnd, wMsg, wParam, lParam)
                ' The NCCALCSIZE_PARAMS struct is not necessary because only the first rectangle is adjusted.
                ' If wParam is 1 or not, the treatment is the same.
                CopyMemory RC, ByVal lParam, LenB(RC)
                RC.Top = RC.Top + (CELL_TEXT_HEIGHT_PADDING_DIP * PixelsPerDIP_Y())
                RC.Bottom = RC.Bottom - ((CELL_TEXT_HEIGHT_PADDING_DIP * PixelsPerDIP_Y()) - 1)
                CopyMemory ByVal lParam, RC, LenB(RC)
                WindowProcEdit = 0
                Exit Function
            Case WM_NCHITTEST
                GetWindowRect hWnd, RC
                DefWindowProc hWnd, WM_NCCALCSIZE, 0, ByVal VarPtr(RC)
                If PtInRect(RC, Get_X_lParam(lParam), Get_Y_lParam(lParam)) <> 0 Then
                    WindowProcEdit = HTCLIENT
                Else
                    WindowProcEdit = FlexDefaultProc(hWnd, wMsg, wParam, lParam)
                    If WindowProcEdit = 0 Then WindowProcEdit = HTBORDER
                End If
                Exit Function
            Case WM_NCPAINT
                Dim hDC As Long
                If wParam = 1 Then ' Alias for entire window
                    hDC = GetWindowDC(hWnd)
                Else
                    hDC = GetDCEx(hWnd, wParam, DCX_WINDOW Or DCX_INTERSECTRGN Or DCX_USESTYLE)
                End If
                If hDC <> 0 Then
                    Dim Brush As Long
                    If VBFlexGridEditBackColorBrush <> 0 Then
                        Brush = VBFlexGridEditBackColorBrush
                    Else
                        Brush = GetSysColorBrush(COLOR_WINDOW)
                    End If
                    Dim WndRect As RECT
                    GetWindowRect hWnd, WndRect
                    RC.Left = 0
                    RC.Right = (WndRect.Right - WndRect.Left)
                    RC.Top = 0
                    RC.Bottom = RC.Top + (CELL_TEXT_HEIGHT_PADDING_DIP * PixelsPerDIP_Y())
                    FillRect hDC, RC, Brush
                    RC.Bottom = (WndRect.Bottom - WndRect.Top)
                    RC.Top = RC.Bottom - ((CELL_TEXT_HEIGHT_PADDING_DIP * PixelsPerDIP_Y()) - 1)
                    FillRect hDC, RC, Brush
                    ReleaseDC hWnd, hDC
                End If
                WindowProcEdit = FlexDefaultProc(hWnd, wMsg, wParam, lParam)
                Exit Function
        End Select
    
    #If ImplementPreTranslateMsg = True Then
    
    Case UM_PRETRANSLATEMSG
        WindowProcEdit = PreTranslateMsg(lParam)
        Exit Function
    
    #End If
    
End Select
WindowProcEdit = FlexDefaultProc(hWnd, wMsg, wParam, lParam)
If wMsg = WM_KILLFOCUS Then DestroyEdit False, FlexEditCloseModeLostFocus
End Function

Private Function WindowProcComboButton(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_MOUSEACTIVATE
        ' It is necessary to break the chain and return MA_ACTIVATE for this window.
        ' This enables the parent window - when it receives WM_MOUSEACTIVATE - to destroy this child window.
        WindowProcComboButton = MA_ACTIVATE
        Exit Function
    Case WM_LBUTTONDOWN
        ' In case the edit window is still active due to failed validation then this ensures that the focus is properly set when clicked from outside.
        If VBFlexGridEditHandle <> 0 Then
            If GetFocus() <> VBFlexGridEditHandle Then SetFocusAPI UserControl.hWnd
        End If
    Case WM_MOUSEMOVE
        If ComboButtonGetState(ODS_HOTLIGHT) = False Then
            Call ComboButtonSetState(ODS_HOTLIGHT, True)
            Dim TME As TRACKMOUSEEVENTSTRUCT
            With TME
            .cbSize = LenB(TME)
            .hWndTrack = hWnd
            .dwFlags = TME_LEAVE
            End With
            TrackMouseEvent TME
        End If
    Case WM_MOUSELEAVE
        Call ComboButtonSetState(ODS_HOTLIGHT, False)
End Select
WindowProcComboButton = FlexDefaultProc(hWnd, wMsg, wParam, lParam)
End Function

Private Function WindowProcComboList(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Static NonClientMouseOver As Boolean, LastMouseMoveLParam As Long
Select Case wMsg
    Case WM_MOUSEACTIVATE
        ' To prevent the popup window from being activated it is necessary to return MA_NOACTIVATE.
        WindowProcComboList = MA_NOACTIVATE
        Exit Function
    Case WM_SHOWWINDOW
        LastMouseMoveLParam = 0
    Case WM_MOUSEMOVE
        If SendMessage(hWnd, WM_NCHITTEST, 0, ByVal GetMessagePos()) = HTVSCROLL Then ReleaseCapture
        If LastMouseMoveLParam <> lParam Or LastMouseMoveLParam = 0 Then
            ComboListSelFromPt Get_X_lParam(lParam), Get_Y_lParam(lParam)
            LastMouseMoveLParam = lParam
        End If
    Case WM_NCMOUSEMOVE
        If NonClientMouseOver = False Then
            NonClientMouseOver = True
            Dim TME As TRACKMOUSEEVENTSTRUCT
            With TME
            .cbSize = LenB(TME)
            .hWndTrack = hWnd
            .dwFlags = TME_LEAVE Or TME_NONCLIENT
            End With
            TrackMouseEvent TME
        End If
    Case WM_NCMOUSELEAVE
        NonClientMouseOver = False
        SetCapture hWnd
    Case WM_LBUTTONDOWN, WM_LBUTTONDBLCLK
        If Not ComboListSelFromPt(Get_X_lParam(lParam), Get_Y_lParam(lParam)) = LB_ERR Then
            Call ComboListCommitSel
            If VBFlexGridComboActiveMode = FlexComboModeDropDown Then
                Call ComboShowDropDown(False)
                DestroyEdit False, FlexEditCloseModeReturn
                Exit Function
            End If
        End If
        ReleaseCapture
        Exit Function ' Prevents the popup window from being focused.
    Case WM_CAPTURECHANGED
        If SendMessage(hWnd, WM_NCHITTEST, 0, ByVal GetMessagePos()) <> HTVSCROLL Then Call ComboShowDropDown(False)
End Select
WindowProcComboList = FlexDefaultProc(hWnd, wMsg, wParam, lParam)
End Function

Private Function WindowProcUserControl(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_CONTEXTMENU
        If wParam = VBFlexGridHandle Then
            Dim P As POINTAPI
            P.X = Get_X_lParam(lParam)
            P.Y = Get_Y_lParam(lParam)
            If P.X = -1 And P.Y = -1 Then
                ' If the user types SHIFT + F10 then the X and Y coordinates are -1.
                RaiseEvent ContextMenu(-1, -1)
            Else
                ScreenToClient VBFlexGridHandle, P
                RaiseEvent ContextMenu(UserControl.ScaleX(P.X, vbPixels, vbContainerPosition), UserControl.ScaleY(P.Y, vbPixels, vbContainerPosition))
            End If
        End If
End Select
WindowProcUserControl = FlexDefaultProc(hWnd, wMsg, wParam, lParam)
If wMsg = WM_SETFOCUS And UCNoSetFocusFwd = False Then SetFocusAPI VBFlexGridHandle
End Function
