VERSION 5.00
Begin VB.UserControl VBFlexGrid 
   Alignable       =   -1  'True
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   DataBindingBehavior=   2  'vbComplexBound
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

#Const ImplementDataSource = True ' True = Required: msdatsrc.tlb

#If False Then
Private FlexOLEDropModeNone, FlexOLEDropModeManual
Private FlexRightToLeftModeNoControl, FlexRightToLeftModeVBAME, FlexRightToLeftModeSystemLocale, FlexRightToLeftModeUserLocale, FlexRightToLeftModeOSLanguage
Private FlexBorderStyleNone, FlexBorderStyleSingle, FlexBorderStyleThin, FlexBorderStyleSunken, FlexBorderStyleRaised
Private FlexAllowUserResizingNone, FlexAllowUserResizingColumns, FlexAllowUserResizingRows, FlexAllowUserResizingBoth
Private FlexSelectionModeFree, FlexSelectionModeByRow, FlexSelectionModeByColumn
Private FlexFillStyleSingle, FlexFillStyleRepeat
Private FlexHighLightNever, FlexHighLightAlways, FlexHighLightWithFocus
Private FlexFocusRectNone, FlexFocusRectLight, FlexFocusRectHeavy
Private FlexGridLineNone, FlexGridLineFlat, FlexGridLineInset, FlexGridLineRaised, FlexGridLineDashes, FlexGridLineDots
Private FlexTextStyleFlat, FlexTextStyleRaised, FlexTextStyleInset, FlexTextStyleRaisedLight, FlexTextStyleInsetLight
Private FlexHitResultNoWhere, FlexHitResultCell, FlexHitResultDividerRowTop, FlexHitResultDividerRowBottom, FlexHitResultDividerColumnLeft, FlexHitResultDividerColumnRight
Private FlexAlignmentLeftTop, FlexAlignmentLeftCenter, FlexAlignmentLeftBottom, FlexAlignmentCenterTop, FlexAlignmentCenterCenter, FlexAlignmentCenterBottom, FlexAlignmentRightTop, FlexAlignmentRightCenter, FlexAlignmentRightBottom, FlexAlignmentGeneral
Private FlexPictureAlignmentLeftTop, FlexPictureAlignmentLeftCenter, FlexPictureAlignmentLeftBottom, FlexPictureAlignmentCenterTop, FlexPictureAlignmentCenterCenter, FlexPictureAlignmentCenterBottom, FlexPictureAlignmentRightTop, FlexPictureAlignmentRightCenter, FlexPictureAlignmentRightBottom, FlexPictureAlignmentStretch, FlexPictureAlignmentTile
Private FlexRowSizingModeIndividual, FlexRowSizingModeAll
Private FlexMergeCellsNever, FlexMergeCellsFree, FlexMergeCellsRestrictRows, FlexMergeCellsRestrictColumns, FlexMergeCellsRestrictAll, FlexMergeCellsFixedOnly
Private FlexSortNone, FlexSortGenericAscending, FlexSortGenericDescending, FlexSortNumericAscending, FlexSortNumericDescending, FlexSortStringNoCaseAscending, FlexSortStringNoCaseDescending, FlexSortStringAscending, FlexSortStringDescending, FlexSortCustom, FlexSortUseColSort, FlexSortCurrencyAscending, FlexSortCurrencyDescending, FlexSortDateAscending, FlexSortDateDescending
Private FlexVisibilityPartialOK, FlexVisibilityCompleteOnly
Private FlexPictureTypeColor, FlexPictureTypeMonochrome
Private FlexEllipsisFormatNone, FlexEllipsisFormatEnd, FlexEllipsisFormatPath, FlexEllipsisFormatWord
Private FlexClearEverywhere, FlexClearFixed, FlexClearScrollable, FlexClearSelection
Private FlexClearEverything, FlexClearText, FlexClearFormatting
Private FlexTabControls, FlexTabCells, FlexTabNext
Private FlexWrapNone, FlexWrapRow, FlexWrapGrid
Private FlexCellText, FlexCellClip, FlexCellTextStyle, FlexCellAlignment, FlexCellPicture, FlexCellPictureAlignment, FlexCellBackColor, FlexCellForeColor, FlexCellToolTipText, FlexCellFontName, FlexCellFontSize, FlexCellFontBold, FlexCellFontItalic, FlexCellFontStrikeThrough, FlexCellFontUnderline, FlexCellFontCharset, FlexCellLeft, FlexCellTop, FlexCellWidth, FlexCellHeight, FlexCellSort
Private FlexAutoSizeModeColWidth, FlexAutoSizeModeRowHeight
Private FlexAutoSizeScopeAll, FlexAutoSizeScopeFixed, FlexAutoSizeScopeScrollable
#End If
Public Enum FlexOLEDropModeConstants
FlexOLEDropModeNone = vbOLEDropNone
FlexOLEDropModeManual = vbOLEDropManual
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
FlexClearSelection = 3
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
TMFirstChar As Byte
TMLastChar As Byte
TMDefaultChar As Byte
TMBreakChar As Byte
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
Private Type SCROLLINFO
cbSize As Long
fMask As Long
nMin As Long
nMax As Long
nPage As Long
nPos As Long
nTrackPos As Long
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
Private Const RCPF_SETSCROLLBARS As Long = &H100
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
Private Type TSELRANGE
LeftCol As Long
TopRow As Long
RightCol As Long
BottomRow As Long
End Type
Private Type THITTESTINFO
PT As POINTAPI
HitRow As Long
HitCol As Long
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
End Type
Private Const RATIO_OF_ROWINFO_HEIGHT_TO_COLINFO_WIDTH As Long = 4
Private Const ROWINFO_HEIGHT_SPACING_DIP As Long = 3
Private Type TROWINFO
Height As Long
Data As Long
Hidden As Boolean
Merge As Boolean
End Type
Private Const COLINFO_WIDTH_SPACING_DIP As Long = 6
Private Type TCOLINFO
Width As Long
Data As Long
Hidden As Boolean
Key As String
Alignment As FlexAlignmentConstants
FixedAlignment As FlexAlignmentConstants
Merge As Boolean
Sort As FlexSortConstants
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
Public Event DividerDblClick(ByVal Row As Long, ByVal Col As Long)
Attribute DividerDblClick.VB_Description = "Occurs when the user double-clicked the divider on a row or column."
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
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function BeginPaint Lib "user32" (ByVal hWnd As Long, ByRef lpPaint As PAINTSTRUCT) As Long
Private Declare Function EndPaint Lib "user32" (ByVal hWnd As Long, ByRef lpPaint As PAINTSTRUCT) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function ExtSelectClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal fnMode As Long) As Long
Private Declare Function GetRgnBox Lib "gdi32" (ByVal hRgn As Long, ByRef lpRect As RECT) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As Any, ByVal bErase As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function TrackMouseEvent Lib "user32" (ByRef lpEventTrack As TRACKMOUSEEVENTSTRUCT) As Long
Private Declare Function GetSystemDefaultLangID Lib "kernel32" () As Integer
Private Declare Function GetUserDefaultLangID Lib "kernel32" () As Integer
Private Declare Function GetUserDefaultUILanguage Lib "kernel32" () As Integer
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoW" (ByVal LCID As Long, ByVal LCType As Long, ByVal lpLCData As Long, ByVal cchData As Long) As Long
Private Declare Function GetMessagePos Lib "user32" () As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function MapWindowPoints Lib "user32" (ByVal hWndFrom As Long, ByVal hWndTo As Long, ByRef lppt As Any, ByVal cPoints As Long) As Long
Private Declare Function SetViewportOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function SetRect Lib "user32" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByRef lpPoint As Any) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, ByRef lpRect As RECT) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextW" (ByVal hDC As Long, ByVal lpchText As Long, ByVal nCount As Long, ByRef lpRect As RECT, ByVal uFormat As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal XLeft As Long, ByVal YTop As Long, ByVal hIcon As Long, ByVal CXWidth As Long, ByVal CYWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32W" (ByVal hDC As Long, ByVal lpsz As Long, ByVal cbString As Long, ByRef lpSize As SIZEAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsW" (ByVal hDC As Long, ByRef lpMetrics As TEXTMETRIC) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoW" (ByVal uAction As Long, ByVal uiParam As Long, ByRef lpvParam As Long, ByVal fWinIni As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetLayout Lib "gdi32" (ByVal hDC As Long, ByVal dwLayout As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function SetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal wBar As Long, ByRef lpScrollInfo As SCROLLINFO, ByVal fRedraw As Long) As Long
Private Declare Function GetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal wBar As Long, ByRef lpScrollInfo As SCROLLINFO) As Long
Private Declare Function PtInRect Lib "user32" (ByRef lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function ClipCursor Lib "user32" (ByRef lpRect As Any) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Const ICC_STANDARD_CLASSES As Long = &H4000
Private Const ICC_TAB_CLASSES As Long = &H8
Private Const RDW_UPDATENOW As Long = &H100, RDW_INVALIDATE As Long = &H1, RDW_ERASE As Long = &H4, RDW_ALLCHILDREN As Long = &H80, RDW_FRAME As Long = &H400
Private Const SWP_FRAMECHANGED As Long = &H20
Private Const SWP_DRAWFRAME As Long = SWP_FRAMECHANGED
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOOWNERZORDER As Long = &H200
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_NOACTIVATE As Long = &H10
Private Const HWND_DESKTOP As Long = &H0
Private Const MK_SHIFT As Long = &H4
Private Const MK_CONTROL As Long = &H8
Private Const TME_LEAVE As Long = &H2
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
Private Const SM_CXBORDER As Long = &H5
Private Const SM_CYBORDER As Long = &H6
Private Const SM_CXEDGE As Long = 45
Private Const SM_CYEDGE As Long = 46
Private Const SIF_RANGE As Long = &H1
Private Const SIF_PAGE As Long = &H2
Private Const SIF_POS As Long = &H4
Private Const SIF_DISABLENOSCROLL As Long = &H8
Private Const SIF_TRACKPOS As Long = &H10
Private Const SIF_ALL As Long = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)
Private Const SPI_GETWHEELSCROLLLINES As Long = &H68
Private Const RGN_DIFF As Long = 4
Private Const RGN_COPY As Long = 5
Private Const DI_NORMAL As Long = &H3
Private Const DT_NOPREFIX As Long = &H800
Private Const DT_RTLREADING As Long = &H20000
Private Const DT_LEFT As Long = &H0
Private Const DT_CENTER As Long = &H1
Private Const DT_RIGHT As Long = &H2
Private Const DT_WORDBREAK As Long = &H10
Private Const DT_PATH_ELLIPSIS As Long = &H4000
Private Const DT_END_ELLIPSIS As Long = &H8000&
Private Const DT_WORD_ELLIPSIS As Long = &H40000
Private Const DT_CALCRECT As Long = &H400
Private Const GWL_STYLE As Long = (-16)
Private Const GWL_EXSTYLE As Long = (-20)
Private Const LAYOUT_RTL As Long = &H1
Private Const WS_BORDER As Long = &H800000
Private Const WS_DLGFRAME As Long = &H400000
Private Const WS_EX_CLIENTEDGE As Long = &H200
Private Const WS_EX_STATICEDGE As Long = &H20000
Private Const WS_EX_WINDOWEDGE As Long = &H100
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_CHILD As Long = &H40000000
Private Const WS_CLIPSIBLINGS As Long = &H4000000
Private Const WS_POPUP As Long = &H80000000
Private Const WS_EX_TOOLWINDOW As Long = &H80
Private Const WS_EX_TOPMOST As Long = &H8
Private Const WS_EX_LAYOUTRTL As Long = &H400000, WS_EX_RTLREADING As Long = &H2000, WS_EX_RIGHT As Long = &H1000, WS_EX_LEFTSCROLLBAR As Long = &H4000
Private Const WS_HSCROLL As Long = &H100000
Private Const WS_VSCROLL As Long = &H200000
Private Const WM_NOTIFY As Long = &H4E
Private Const WM_NOTIFYFORMAT As Long = &H55
Private Const WM_MOUSEACTIVATE As Long = &H21, MA_NOACTIVATE As Long = &H3, MA_NOACTIVATEANDEAT As Long = &H4, HTBORDER As Long = 18
Private Const WM_SETTINGCHANGE As Long = &H1A
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const SW_HIDE As Long = &H0
Private Const WM_SETFOCUS As Long = &H7
Private Const WM_KILLFOCUS As Long = &H8
Private Const WM_KEYDOWN As Long = &H100
Private Const WM_KEYUP As Long = &H101
Private Const WM_CHAR As Long = &H102
Private Const WM_UNICHAR As Long = &H109, UNICODE_NOCHAR As Long = &HFFFF&
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
Private Const WM_GETFONT As Long = &H31
Private Const WM_SETREDRAW As Long = &HB
Private Const WM_SIZE As Long = &H5
Private Const WM_SETCURSOR As Long = &H20, HTCLIENT As Long = 1
Private Const WM_ERASEBKGND As Long = &H14
Private Const WM_PAINT As Long = &HF
Private Const WM_PRINTCLIENT As Long = &H318
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
Implements OLEGuids.IOleInPlaceActiveObjectVB
Implements OLEGuids.IPerPropertyBrowsingVB
Private VBFlexGridHandle As Long, VBFlexGridToolTipHandle As Long
Private VBFlexGridFontHandle As Long, VBFlexGridFontFixedHandle As Long
Private VBFlexGridBackColorBrush As Long
Private VBFlexGridBackColorAltBrush As Long
Private VBFlexGridBackColorBkgBrush As Long
Private VBFlexGridBackColorFixedBrush As Long
Private VBFlexGridBackColorSelBrush As Long
Private VBFlexGridGridLinePen As Long, VBFlexGridPenStyle As Long
Private VBFlexGridGridLineFixedPen As Long, VBFlexGridFixedPenStyle As Long
Private VBFlexGridGridLineWhitePen As Long, VBFlexGridGridLineBlackPen As Long
Private VBFlexGridCells As TROWS
Private VBFlexGridColsInfo() As TCOLINFO
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
Private VBFlexGridCaptureDividerDrag As Boolean
Private VBFlexGridToolTipRow As Long, VBFlexGridToolTipCol As Long
Private VBFlexGridMouseMoveRow As Long, VBFlexGridMouseMoveCol As Long
Private VBFlexGridMouseMoveChanged As Boolean
Private VBFlexGridDividerDragSplitterRect As RECT
Private VBFlexGridHitRow As Long, VBFlexGridHitCol As Long
Private VBFlexGridHitResult As FlexHitResultConstants
Private VBFlexGridWheelScrollLines As Long
Private VBFlexGridFocused As Boolean
Private VBFlexGridNoRedraw As Boolean
Private VBFlexGridCharCodeCache As Long
Private VBFlexGridIsClick As Boolean
Private VBFlexGridMouseOver As Boolean
Private VBFlexGridDesignMode As Boolean
Private VBFlexGridRTLLayout As Boolean, VBFlexGridRTLReading As Boolean
Private VBFlexGridAlignable As Boolean
Private VBFlexGridSort As FlexSortConstants
Private DispIDMousePointer As Long

#If ImplementDataSource = True Then

Private PropDataSource As MSDATASRC.DataSource, PropDataMember As MSDATASRC.DataMember, PropRecordset As Object

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
Private PropRows As Long, PropCols As Long
Private PropAllowBigSelection As Boolean
Private PropAllowSelection As Boolean
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
Private PropEllipsisFormat As FlexEllipsisFormatConstants
Private PropEllipsisFormatFixed As FlexEllipsisFormatConstants
Private PropRedraw As Boolean
Private PropDoubleBuffer As Boolean
Private PropTabBehavior As FlexTabBehaviorConstants
Private PropWrapCellBehavior As FlexWrapCellBehaviorConstants
Private PropShowInfoTips As Boolean
Private PropShowLabelTips As Boolean
Private PropClipSeparators As String
Private PropFormatString As String

Private Sub IOleInPlaceActiveObjectVB_TranslateAccelerator(ByRef Handled As Boolean, ByRef RetVal As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal Shift As Long)
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
            SendMessage VBFlexGridHandle, wMsg, wParam, ByVal lParam
            Handled = True
        Case vbKeyTab, vbKeyReturn, vbKeyEscape
            If KeyCode = vbKeyTab Then
                Select Case PropTabBehavior
                    Case FlexTabCells
                        IsInputKey = True
                    Case FlexTabNext
                        Select Case PropWrapCellBehavior
                            Case FlexWrapNone
                                If (Shift And vbShiftMask) = 0 Then
                                    If VBFlexGridCol < (PropCols - 1) Then IsInputKey = True
                                Else
                                    If VBFlexGridCol > PropFixedCols Then IsInputKey = True
                                End If
                            Case FlexWrapRow
                                If (Shift And vbShiftMask) = 0 Then
                                    If VBFlexGridRow < (PropRows - 1) Or VBFlexGridCol < (PropCols - 1) Then IsInputKey = True
                                Else
                                    If VBFlexGridRow > PropFixedRows Or VBFlexGridCol > PropFixedCols Then IsInputKey = True
                                End If
                            Case FlexWrapGrid
                                IsInputKey = True
                        End Select
                End Select
            End If
            If IsInputKey = True Then
                SendMessage VBFlexGridHandle, wMsg, wParam, ByVal lParam
                Handled = True
            End If
    End Select
End If
End Sub

Private Sub IPerPropertyBrowsingVB_GetDisplayString(ByRef Handled As Boolean, ByVal DispID As Long, ByRef DisplayName As String)
If DispID = DispIDMousePointer Then
    Select Case PropMousePointer
        Case 0: DisplayName = "0 - Default"
        Case 1: DisplayName = "1 - Arrow"
        Case 2: DisplayName = "2 - Cross"
        Case 3: DisplayName = "3 - I-Beam"
        Case 4: DisplayName = "4 - Hand"
        Case 5: DisplayName = "5 - Size"
        Case 6: DisplayName = "6 - Size NE SW"
        Case 7: DisplayName = "7 - Size N S"
        Case 8: DisplayName = "8 - Size NW SE"
        Case 9: DisplayName = "9 - Size W E"
        Case 10: DisplayName = "10 - Up Arrow"
        Case 11: DisplayName = "11 - Hourglass"
        Case 12: DisplayName = "12 - No Drop"
        Case 13: DisplayName = "13 - Arrow and Hourglass"
        Case 14: DisplayName = "14 - Arrow and Question"
        Case 15: DisplayName = "15 - Size All"
        Case 16: DisplayName = "16 - Arrow and CD"
        Case 99: DisplayName = "99 - Custom"
    End Select
    Handled = True
End If
End Sub

Private Sub IPerPropertyBrowsingVB_GetPredefinedStrings(ByRef Handled As Boolean, ByVal DispID As Long, ByRef StringsOut() As String, ByRef CookiesOut() As Long)
If DispID = DispIDMousePointer Then
    ReDim StringsOut(0 To (17 + 1)) As String
    ReDim CookiesOut(0 To (17 + 1)) As Long
    StringsOut(0) = "0 - Default": CookiesOut(0) = 0
    StringsOut(1) = "1 - Arrow": CookiesOut(1) = 1
    StringsOut(2) = "2 - Cross": CookiesOut(2) = 2
    StringsOut(3) = "3 - I-Beam": CookiesOut(3) = 3
    StringsOut(4) = "4 - Hand": CookiesOut(4) = 4
    StringsOut(5) = "5 - Size": CookiesOut(5) = 5
    StringsOut(6) = "6 - Size NE SW": CookiesOut(6) = 6
    StringsOut(7) = "7 - Size N S": CookiesOut(7) = 7
    StringsOut(8) = "8 - Size NW SE": CookiesOut(8) = 8
    StringsOut(9) = "9 - Size W E": CookiesOut(9) = 9
    StringsOut(10) = "10 - Up Arrow": CookiesOut(10) = 10
    StringsOut(11) = "11 - Hourglass": CookiesOut(11) = 11
    StringsOut(12) = "12 - No Drop": CookiesOut(12) = 12
    StringsOut(13) = "13 - Arrow and Hourglass": CookiesOut(13) = 13
    StringsOut(14) = "14 - Arrow and Question": CookiesOut(14) = 14
    StringsOut(15) = "15 - Size All": CookiesOut(15) = 15
    StringsOut(16) = "16 - Arrow and CD": CookiesOut(16) = 16
    StringsOut(17) = "99 - Custom": CookiesOut(17) = 99
    Handled = True
End If
End Sub

Private Sub IPerPropertyBrowsingVB_GetPredefinedValue(ByRef Handled As Boolean, ByVal DispID As Long, ByVal Cookie As Long, ByRef Value As Variant)
If DispID = DispIDMousePointer Then
    Value = Cookie
    Handled = True
End If
End Sub

Private Sub UserControl_Initialize()
Call FlexLoadShellMod
Call FlexWndRegisterClass
Call FlexInitCC(ICC_STANDARD_CLASSES)
Call SetVTableSubclass(Me, VTableInterfaceInPlaceActiveObject)
Call SetVTableSubclass(Me, VTableInterfacePerPropertyBrowsing)
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
VBFlexGridCaptureDividerDrag = False
VBFlexGridToolTipRow = -1
VBFlexGridToolTipCol = -1
VBFlexGridMouseMoveRow = -1
VBFlexGridMouseMoveCol = -1
VBFlexGridMouseMoveChanged = False
VBFlexGridHitRow = -1
VBFlexGridHitCol = -1
VBFlexGridHitResult = FlexHitResultNoWhere
SystemParametersInfo SPI_GETWHEELSCROLLLINES, 0, VBFlexGridWheelScrollLines, 0
End Sub

Private Sub UserControl_InitProperties()
If DispIDMousePointer = 0 Then DispIDMousePointer = GetDispID(Me, "MousePointer")
On Error Resume Next
If UserControl.ParentControls.Count = 0 Then VBFlexGridAlignable = False Else VBFlexGridAlignable = True
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
PropRows = 2
PropCols = 2
PropAllowBigSelection = True
PropAllowSelection = True
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
PropEllipsisFormat = FlexEllipsisFormatNone
PropEllipsisFormatFixed = FlexEllipsisFormatNone
PropRedraw = True
PropDoubleBuffer = True
PropTabBehavior = FlexTabControls
PropWrapCellBehavior = FlexWrapNone
PropShowInfoTips = False
PropShowLabelTips = False
PropClipSeparators = vbNullString
PropFormatString = vbNullString
VBFlexGridDesignMode = Not Ambient.UserMode
Call CreateVBFlexGrid
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
If DispIDMousePointer = 0 Then DispIDMousePointer = GetDispID(Me, "MousePointer")
On Error Resume Next
If UserControl.ParentControls.Count = 0 Then VBFlexGridAlignable = False Else VBFlexGridAlignable = True
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
PropRows = .ReadProperty("Rows", 2)
PropCols = .ReadProperty("Cols", 2)
PropAllowBigSelection = .ReadProperty("AllowBigSelection", True)
PropAllowSelection = .ReadProperty("AllowSelection", True)
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
PropEllipsisFormat = .ReadProperty("EllipsisFormat", FlexEllipsisFormatNone)
PropEllipsisFormatFixed = .ReadProperty("EllipsisFormatFixed", FlexEllipsisFormatNone)
PropRedraw = .ReadProperty("Redraw", True)
PropDoubleBuffer = .ReadProperty("DoubleBuffer", True)
PropTabBehavior = .ReadProperty("TabBehavior", FlexTabControls)
PropWrapCellBehavior = .ReadProperty("WrapCellBehavior", FlexWrapNone)
PropShowInfoTips = .ReadProperty("ShowInfoTips", False)
PropShowLabelTips = .ReadProperty("ShowLabelTips", False)
PropClipSeparators = VarToStr(.ReadProperty("ClipSeparators", vbNullString))
PropFormatString = VarToStr(.ReadProperty("FormatString", vbNullString))
End With
VBFlexGridDesignMode = Not Ambient.UserMode
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
.WriteProperty "Rows", PropRows, 2
.WriteProperty "Cols", PropCols, 2
.WriteProperty "AllowBigSelection", PropAllowBigSelection, True
.WriteProperty "AllowSelection", PropAllowSelection, True
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
.WriteProperty "EllipsisFormat", PropEllipsisFormat, FlexEllipsisFormatNone
.WriteProperty "EllipsisFormatFixed", PropEllipsisFormatFixed, FlexEllipsisFormatNone
.WriteProperty "Redraw", PropRedraw, True
.WriteProperty "DoubleBuffer", PropDoubleBuffer, True
.WriteProperty "TabBehavior", PropTabBehavior, FlexTabControls
.WriteProperty "WrapCellBehavior", PropWrapCellBehavior, FlexWrapNone
.WriteProperty "ShowInfoTips", PropShowInfoTips, False
.WriteProperty "ShowLabelTips", PropShowLabelTips, False
.WriteProperty "ClipSeparators", StrToVar(PropClipSeparators), vbNullString
.WriteProperty "FormatString", StrToVar(PropFormatString), vbNullString
End With
End Sub

Private Sub UserControl_Paint()
If VBFlexGridHandle = 0 Or VBFlexGridDesignMode = False Then Exit Sub
Dim OldLayout As Long, ClientRect As RECT, hRgn As Long
If PropRightToLeft = True And PropRightToLeftLayout = True Then OldLayout = SetLayout(UserControl.hDC, LAYOUT_RTL)
GetClientRect UserControl.hWnd, ClientRect
If PropDoubleBuffer = True Then
    Dim hDCBmp As Long
    Dim hBmp As Long, hBmpOld As Long
    hDCBmp = CreateCompatibleDC(UserControl.hDC)
    If hDCBmp <> 0 Then
        hBmp = CreateCompatibleBitmap(UserControl.hDC, ClientRect.Right - ClientRect.Left, ClientRect.Bottom - ClientRect.Top)
        If hBmp <> 0 Then
            hBmpOld = SelectObject(hDCBmp, hBmp)
            If (0 / 1) + (Not Not VBFlexGridCells.Rows()) = 0 Then
                If VBFlexGridBackColorBkgBrush <> 0 Then FillRect hDCBmp, ClientRect, VBFlexGridBackColorBkgBrush
                Call DrawGrid(hDCBmp, -1)
            Else
                Call DrawGrid(hDCBmp, hRgn)
                If hRgn <> 0 Then ExtSelectClipRgn UserControl.hDC, hRgn, RGN_COPY
            End If
            BitBlt UserControl.hDC, 0, 0, ClientRect.Right - ClientRect.Left, ClientRect.Bottom - ClientRect.Top, hDCBmp, 0, 0, vbSrcCopy
            If hRgn <> 0 Then
                ExtSelectClipRgn UserControl.hDC, 0, RGN_COPY
                DeleteObject hRgn
            End If
            SelectObject hDCBmp, hBmpOld
            DeleteObject hBmp
        End If
        DeleteDC hDCBmp
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
If DPICorrectionFactor() <> 1 Then
    .Extender.Move .Extender.Left + .ScaleX(1, vbPixels, vbContainerPosition), .Extender.Top + .ScaleY(1, vbPixels, vbContainerPosition)
    .Extender.Move .Extender.Left - .ScaleX(1, vbPixels, vbContainerPosition), .Extender.Top - .ScaleY(1, vbPixels, vbContainerPosition)
End If
If VBFlexGridHandle <> 0 And VBFlexGridDesignMode = False Then MoveWindow VBFlexGridHandle, 0, 0, .ScaleWidth, .ScaleHeight, 1
End With
InProc = False
End Sub

Private Sub UserControl_Terminate()
Call RemoveVTableSubclass(Me, VTableInterfaceInPlaceActiveObject)
Call RemoveVTableSubclass(Me, VTableInterfacePerPropertyBrowsing)
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
If Ambient.UserMode = True Then
    If Not PropDataSource Is Nothing Then
        If PropRecordset Is Nothing Then Set PropRecordset = CreateObject("ADODB.Recordset")
        With PropRecordset
        If .State <> 0 Then .Close
        If StrPtr(PropDataMember) = 0 Then .DataMember = "" Else .DataMember = PropDataMember
        Set .DataSource = PropDataSource
        If .State <> 0 Then
            If .RecordCount > -1 Then ' The cursor type of the Recordset affects whether the number of records can be determined.
                If .RecordCount > 0 Then
                    ' In ADO a .MoveLast and .MoveFirst to fully populate the Recordset is not necessary. (unlike to DAO)
                    .MoveFirst
                    Me.Rows = PropFixedRows + .RecordCount
                Else
                    Me.Rows = PropFixedRows + 1
                End If
                Me.Cols = PropFixedCols + .Fields.Count
                Dim iRow As Long, iCol As Long
                If PropFixedRows > 0 Then
                    For iCol = 0 To (.Fields.Count - 1)
                        Me.TextMatrix(0, iCol + PropFixedCols) = .Fields(iCol).Name
                    Next iCol
                End If
                iRow = PropFixedRows
                If .RecordCount > 0 Then
                    Do Until .EOF
                        For iCol = PropFixedCols To (PropCols - 1)
                            If Not IsNull(.Fields(iCol - PropFixedCols).Value) Then
                                Me.TextMatrix(iRow, iCol) = .Fields(iCol - PropFixedCols).Value
                            Else
                                Me.TextMatrix(iRow, iCol) = vbNullString
                            End If
                        Next iCol
                        .MoveNext
                        iRow = iRow + 1
                    Loop
                Else
                    For iCol = PropFixedCols To (PropCols - 1)
                        Me.TextMatrix(iRow, iCol) = vbNullString
                    Next iCol
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
    Dim TM As TEXTMETRIC
    If VBFlexGridFontHandle <> 0 Then SelectObject hDCScreen, VBFlexGridFontHandle
    If GetTextMetrics(hDCScreen, TM) <> 0 Then
        VBFlexGridDefaultRowHeight = TM.TMHeight + (ROWINFO_HEIGHT_SPACING_DIP * PixelsPerDIP_Y())
        VBFlexGridDefaultColWidth = VBFlexGridDefaultRowHeight * RATIO_OF_ROWINFO_HEIGHT_TO_COLINFO_WIDTH
    End If
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
    Dim TM As TEXTMETRIC
    If VBFlexGridFontHandle <> 0 Then SelectObject hDCScreen, VBFlexGridFontHandle
    If GetTextMetrics(hDCScreen, TM) <> 0 Then
        VBFlexGridDefaultRowHeight = TM.TMHeight + (ROWINFO_HEIGHT_SPACING_DIP * PixelsPerDIP_Y())
        VBFlexGridDefaultColWidth = VBFlexGridDefaultRowHeight * RATIO_OF_ROWINFO_HEIGHT_TO_COLINFO_WIDTH
    End If
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
        Dim TM As TEXTMETRIC
        If VBFlexGridFontFixedHandle <> 0 Then SelectObject hDCScreen, VBFlexGridFontFixedHandle
        If GetTextMetrics(hDCScreen, TM) <> 0 Then
            VBFlexGridDefaultFixedRowHeight = TM.TMHeight + (ROWINFO_HEIGHT_SPACING_DIP * PixelsPerDIP_Y())
            VBFlexGridDefaultFixedColWidth = VBFlexGridDefaultFixedRowHeight * RATIO_OF_ROWINFO_HEIGHT_TO_COLINFO_WIDTH
        End If
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
    Dim TM As TEXTMETRIC
    If VBFlexGridFontFixedHandle <> 0 Then SelectObject hDCScreen, VBFlexGridFontFixedHandle
    If GetTextMetrics(hDCScreen, TM) <> 0 Then
        VBFlexGridDefaultFixedRowHeight = TM.TMHeight + (ROWINFO_HEIGHT_SPACING_DIP * PixelsPerDIP_Y())
        VBFlexGridDefaultFixedColWidth = VBFlexGridDefaultFixedRowHeight * RATIO_OF_ROWINFO_HEIGHT_TO_COLINFO_WIDTH
    End If
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
If VBFlexGridHandle <> 0 And EnabledVisualStyles() = True Then
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

Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
MousePointer = PropMousePointer
End Property

Public Property Let MousePointer(ByVal Value As Integer)
Select Case Value
    Case 0 To 16, 99
        PropMousePointer = Value
    Case Else
        Err.Raise 380
End Select
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
        If Ambient.UserMode = False Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
    End If
End If
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
If Ambient.UserMode = True Then
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
        If Ambient.UserMode = True Then
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
    If Ambient.UserMode = False Then
        MsgBox "Invalid Row Value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise Number:=30009, Description:="Invalid Row value"
    End If
ElseIf Value >= PropRows Then
    If Ambient.UserMode = False Then
        MsgBox "FixedRows must be at least one less than Rows value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise Number:=30016, Description:="FixedRows must be at least one less than Rows value"
    End If
End If
PropFixedRows = Value
Dim RCP As TROWCOLPARAMS
With RCP
.Mask = RCPM_ROW Or RCPM_TOPROW
.Flags = RCPF_SETSCROLLBARS
.Row = PropFixedRows
.TopRow = PropFixedRows
Select Case PropSelectionMode
    Case FlexSelectionModeFree
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
    If Ambient.UserMode = False Then
        MsgBox "Invalid Col value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise Number:=30010, Description:="Invalid Col value"
    End If
ElseIf Value >= PropCols Then
    If Ambient.UserMode = False Then
        MsgBox "FixedCols must be at least one less than Cols value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise Number:=30017, Description:="FixedCols must be at least one less than Cols value"
    End If
End If
PropFixedCols = Value
Dim RCP As TROWCOLPARAMS
With RCP
.Mask = RCPM_COL Or RCPM_LEFTCOL
.Flags = RCPF_SETSCROLLBARS
.Col = PropFixedCols
.LeftCol = PropFixedCols
Select Case PropSelectionMode
    Case FlexSelectionModeFree
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

Public Property Get Rows() As Long
Attribute Rows.VB_Description = "Returns/sets the total number of columns or rows in the flex grid."
Attribute Rows.VB_MemberFlags = "200"
Rows = PropRows
End Property

Public Property Let Rows(ByVal Value As Long)
If Value < 0 Then
    If Ambient.UserMode = False Then
        MsgBox "Invalid Row value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise Number:=30009, Description:="Invalid Row value"
    End If
ElseIf Value <= PropFixedRows And Value > 0 Then
    PropFixedRows = Value - 1
End If
If Value > 0 And PropRows < 1 Then
    PropRows = Value
    If PropCols > 0 Then Call InitFlexGridCells
ElseIf Value < 1 And PropRows > 0 Then
    PropRows = Value
    PropFixedRows = 0
    Call EraseFlexGridCells
ElseIf Value <> PropRows And PropCols > 0 Then
    ReDim Preserve VBFlexGridCells.Rows(0 To (Value - 1)) As TCOLS
    If Value > PropRows Then
        Dim i As Long, j As Long
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
.Flags = RCPF_SETSCROLLBARS
If VBFlexGridRow > (PropRows - 1) Then
    .Mask = .Mask Or RCPM_ROW
    .Row = (PropRows - 1)
End If
Select Case PropSelectionMode
    Case FlexSelectionModeFree, FlexSelectionModeByRow
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
If VBFlexGridTopRow > (PropRows - 1) Then
    .Mask = .Mask Or RCPM_TOPROW
    .Flags = .Flags Or RCPF_CHECKTOPROW
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
    If Ambient.UserMode = False Then
        MsgBox "Invalid Col value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise Number:=30010, Description:="Invalid Col value"
    End If
ElseIf Value <= PropFixedCols And Value > 0 Then
    PropFixedCols = Value - 1
End If
If Value > 0 And PropCols < 1 Then
    PropCols = Value
    If PropRows > 0 Then Call InitFlexGridCells
ElseIf Value < 1 And PropCols > 0 Then
    PropCols = Value
    PropFixedCols = 0
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
.Flags = RCPF_SETSCROLLBARS
If VBFlexGridCol > (PropCols - 1) Then
    .Mask = .Mask Or RCPM_COL
    .Col = (PropCols - 1)
End If
Select Case PropSelectionMode
    Case FlexSelectionModeFree, FlexSelectionModeByColumn
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
If VBFlexGridLeftCol > (PropCols - 1) Then
    .Mask = .Mask Or RCPM_LEFTCOL
    .Flags = .Flags Or RCPF_CHECKLEFTCOL
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
        Case FlexSelectionModeFree
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
    Case FlexSelectionModeFree, FlexSelectionModeByRow, FlexSelectionModeByColumn
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
    Case FlexSelectionModeFree
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
    If Ambient.UserMode = False Then
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
    If Ambient.UserMode = False Then
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
    If Ambient.UserMode = False Then
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
    If Ambient.UserMode = False Then
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
    If Ambient.UserMode = False Then
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
PropWordWrap = Value
Call RedrawGrid
UserControl.PropertyChanged "WordWrap"
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
Select Case Value
    Case FlexSortNone, FlexSortGenericAscending, FlexSortGenericDescending, FlexSortNumericAscending, FlexSortNumericDescending, FlexSortStringNoCaseAscending, FlexSortStringNoCaseDescending, FlexSortStringAscending, FlexSortStringDescending, FlexSortCustom, FlexSortUseColSort, FlexSortCurrencyAscending, FlexSortCurrencyDescending, FlexSortDateAscending, FlexSortDateDescending
        VBFlexGridSort = Value
        If VBFlexGridSort = FlexSortNone Then Exit Property
        If (VBFlexGridRow < 0 Or VBFlexGridRowSel < 0) Or (VBFlexGridCol < 0 Or VBFlexGridColSel < 0) Then
            ' Error shall not be raised. Do nothing in this case.
            Exit Property
        End If
        Dim SelRange As TSELRANGE, iCol As Long, Sort As FlexSortConstants
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
        .Flags = RCPF_CHECKTOPROW Or RCPF_SETSCROLLBARS
        .TopRow = VBFlexGridTopRow
        Call SetRowColParams(RCP)
        End With
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
If VBFlexGridHandle <> 0 And Ambient.UserMode = True Then
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
If VBFlexGridHandle <> 0 And Ambient.UserMode = True Then
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
        If Ambient.UserMode = False Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
    Case 2
        If StrComp(Left$(Value, 1), Right$(Value, 1)) = 0 Then
            If Ambient.UserMode = False Then
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
                        .Alignment = FlexAlignmentLeftCenter
                End Select
                .Width = GetTextSize(0, iCol, Temp).CX + Spacing
                End With
                With VBFlexGridCells.Rows(0).Cols(iCol)
                .Text = Trim$(Temp)
                End With
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
                With VBFlexGridCells.Rows(iRow).Cols(0)
                .Text = Trim$(Temp)
                End With
                Pos2 = Pos1
                iRow = iRow + 1
            Loop Until Pos1 = 0
            Pos1 = 0
            Pos2 = 0
        End If
        Dim RCP As TROWCOLPARAMS
        With RCP
        .Mask = RCPM_LEFTCOL
        .Flags = RCPF_CHECKLEFTCOL Or RCPF_SETSCROLLBARS
        .LeftCol = VBFlexGridLeftCol
        Call SetRowColParams(RCP)
        End With
    End If
End If
UserControl.PropertyChanged "FormatString"
End Property

Private Sub CreateVBFlexGrid()
If VBFlexGridHandle <> 0 Then Exit Sub
Call InitFlexGridCells
If VBFlexGridDesignMode = False Then
    Dim dwStyle As Long, dwExStyle As Long
    dwStyle = WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS
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
    If VBFlexGridHandle <> 0 Then SetWindowLong VBFlexGridHandle, 0, ObjPtr(Me)
    If PropShowInfoTips = True Or PropShowLabelTips = True Then Call CreateToolTip
Else
    VBFlexGridHandle = UserControl.hWnd
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
If VBFlexGridDesignMode = False Then Call FlexSetSubclass(UserControl.hWnd, Me, 2)
UserControl.BackColor = PropBackColorBkg
End Sub

Private Sub CreateToolTip()
Static Done As Boolean
Dim dwExStyle As Long
If VBFlexGridToolTipHandle <> 0 Then Exit Sub
If Done = False Then
    Call FlexInitCC(ICC_TAB_CLASSES)
    Done = True
End If
dwExStyle = WS_EX_TOOLWINDOW Or WS_EX_TOPMOST
If PropRightToLeft = True And PropRightToLeftLayout = True Then dwExStyle = dwExStyle Or WS_EX_LAYOUTRTL
VBFlexGridToolTipHandle = CreateWindowEx(dwExStyle, StrPtr("tooltips_class32"), StrPtr("Tool Tip"), WS_POPUP Or TTS_ALWAYSTIP Or TTS_NOPREFIX, 0, 0, 0, 0, UserControl.hWnd, 0, App.hInstance, ByVal 0&)
If VBFlexGridToolTipHandle <> 0 Then
    SendMessage VBFlexGridToolTipHandle, TTM_SETMAXTIPWIDTH, 0, ByVal &H7FFF&
    Dim TI As TOOLINFO
    With TI
    .cbSize = LenB(TI)
    .hWnd = VBFlexGridHandle
    .uId = 0
    .uFlags = TTF_SUBCLASS Or TTF_TRANSPARENT Or TTF_PARSELINKS
    If PropRightToLeft = True And PropRightToLeftLayout = False Then .uFlags = .uFlags Or TTF_RTLREADING
    .lpszText = LPSTR_TEXTCALLBACK
    GetClientRect VBFlexGridHandle, .RC
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
    SetWindowLong VBFlexGridHandle, 0, 0
    ShowWindow VBFlexGridHandle, SW_HIDE
    SetParent VBFlexGridHandle, 0
    DestroyWindow VBFlexGridHandle
End If
VBFlexGridHandle = 0
Call EraseFlexGridCells
If VBFlexGridFontHandle <> 0 Then
    DeleteObject VBFlexGridFontHandle
    VBFlexGridFontHandle = 0
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
SetParent VBFlexGridToolTipHandle, 0
DestroyWindow VBFlexGridToolTipHandle
VBFlexGridToolTipHandle = 0
VBFlexGridToolTipRow = -1
VBFlexGridToolTipCol = -1
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
UserControl.Refresh
If VBFlexGridNoRedraw = False And VBFlexGridDesignMode = False Then RedrawWindow UserControl.hWnd, 0, 0, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
End Sub

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
    With VBFlexGridCells.Rows(IndexLong)
    Do
        Pos1 = InStr(Pos1 + 1, Item, ColSeparator)
        If Pos1 > 0 Then
            If iCol < PropCols Then .Cols(iCol).Text = Mid$(Item, Pos2 + 1, Pos1 - Pos2 - 1)
        Else
            If iCol < PropCols Then .Cols(iCol).Text = Mid$(Item, Pos2 + 1)
        End If
        Pos2 = Pos1
        iCol = iCol + 1
    Loop Until Pos1 = 0
    End With
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
        Case FlexSelectionModeFree, FlexSelectionModeByRow
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
    Case FlexClearEverywhere, FlexClearFixed, FlexClearScrollable, FlexClearSelection
    Case Else
        Err.Raise 380
End Select
Select Case What
    Case FlexClearEverything, FlexClearText, FlexClearFormatting
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
    Case FlexClearSelection
        Dim SelRange As TSELRANGE
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
    Case FlexSelectionModeFree
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
    Case FlexSelectionModeFree
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
Dim SelRange As TSELRANGE
Call GetSelRangeStruct(SelRange)
With SelRange
Row1 = .TopRow
Col1 = .LeftCol
Row2 = .BottomRow
Col2 = .RightCol
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
Dim ClientRect As RECT, GridRect As RECT, iRow As Long
GetClientRect VBFlexGridHandle, ClientRect
With GridRect
For iRow = 0 To (PropFixedRows - 1)
    .Bottom = .Bottom + GetRowHeight(iRow)
Next iRow
BottomRow = VBFlexGridTopRow
For iRow = VBFlexGridTopRow To (PropRows - 1)
    If Visibility = FlexVisibilityPartialOK Then If .Bottom > ClientRect.Bottom Then Exit For
    .Bottom = .Bottom + GetRowHeight(iRow)
    If Visibility = FlexVisibilityCompleteOnly Then If .Bottom > ClientRect.Bottom Then Exit For
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
Dim ClientRect As RECT, GridRect As RECT, iCol As Long
GetClientRect VBFlexGridHandle, ClientRect
With GridRect
For iCol = 0 To (PropFixedCols - 1)
    .Right = .Right + GetColWidth(iCol)
Next iCol
RightCol = VBFlexGridLeftCol
For iCol = VBFlexGridLeftCol To (PropCols - 1)
    If Visibility = FlexVisibilityPartialOK Then If .Right > ClientRect.Right Then Exit For
    .Right = .Right + GetColWidth(iCol)
    If Visibility = FlexVisibilityCompleteOnly Then If .Right > ClientRect.Right Then Exit For
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
If Index > (PropFixedRows - 1) Then
    For i = 0 To (PropFixedRows - 1)
        If i < Index Then Value = Value + GetRowHeight(i)
    Next i
    For i = VBFlexGridTopRow To (Index - 1)
        Value = Value + GetRowHeight(i)
    Next i
    If Index < VBFlexGridTopRow Then
        For i = PropFixedRows To (VBFlexGridTopRow - 1)
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
.Flags = RCPF_CHECKTOPROW Or RCPF_SETSCROLLBARS
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
.Flags = RCPF_CHECKTOPROW Or RCPF_SETSCROLLBARS
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
RowHidden = VBFlexGridCells.Rows(Index).RowInfo.Hidden
End Property

Public Property Let RowHidden(ByVal Index As Long, ByVal Value As Boolean)
If Index <> -1 And (Index < 0 Or Index > (PropRows - 1)) Then Err.Raise Number:=30009, Description:="Invalid Row value"
If Index > -1 Then
    VBFlexGridCells.Rows(Index).RowInfo.Hidden = Value
Else
    Dim i As Long
    For i = 0 To (PropRows - 1)
        VBFlexGridCells.Rows(i).RowInfo.Hidden = Value
    Next i
End If
Dim RCP As TROWCOLPARAMS
With RCP
.Mask = RCPM_TOPROW
.Flags = RCPF_CHECKTOPROW Or RCPF_SETSCROLLBARS
.TopRow = VBFlexGridTopRow
Call SetRowColParams(RCP)
End With
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
    Dim ClientRect As RECT, GridRect As RECT, iRow As Long
    GetClientRect VBFlexGridHandle, ClientRect
    With GridRect
    If Index <= (PropFixedRows - 1) Then
        RowIsVisible = True
        For iRow = 0 To (PropFixedRows - 1)
            If Visibility = FlexVisibilityCompleteOnly Then .Bottom = .Bottom + GetRowHeight(iRow)
            If .Bottom > ClientRect.Bottom Then
                RowIsVisible = False
                Exit For
            End If
            If Visibility = FlexVisibilityPartialOK Then .Bottom = .Bottom + GetRowHeight(iRow)
            If iRow >= Index Then Exit For
        Next iRow
    ElseIf Index >= VBFlexGridTopRow Then
        RowIsVisible = True
        For iRow = 0 To (PropFixedRows - 1)
            .Bottom = .Bottom + GetRowHeight(iRow)
        Next iRow
        For iRow = VBFlexGridTopRow To (PropRows - 1)
            If Visibility = FlexVisibilityCompleteOnly Then .Bottom = .Bottom + GetRowHeight(iRow)
            If .Bottom > ClientRect.Bottom Then
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
Dim ClientRect As RECT, GridRect As RECT, iRow As Long, Count As Long
GetClientRect VBFlexGridHandle, ClientRect
With GridRect
For iRow = 0 To (PropFixedRows - 1)
    If Visibility = FlexVisibilityPartialOK Then If .Bottom > ClientRect.Bottom Then Exit For
    .Bottom = .Bottom + GetRowHeight(iRow)
    If Visibility = FlexVisibilityCompleteOnly Then If .Bottom > ClientRect.Bottom Then Exit For
    Count = Count + 1
Next iRow
For iRow = VBFlexGridTopRow To (PropRows - 1)
    If Visibility = FlexVisibilityPartialOK Then If .Bottom > ClientRect.Bottom Then Exit For
    .Bottom = .Bottom + GetRowHeight(iRow)
    If Visibility = FlexVisibilityCompleteOnly Then If .Bottom > ClientRect.Bottom Then Exit For
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
Dim ClientRect As RECT, GridRect As RECT, iRow As Long, Count As Long
GetClientRect VBFlexGridHandle, ClientRect
With GridRect
For iRow = 0 To (PropFixedRows - 1)
    If Visibility = FlexVisibilityPartialOK Then If .Bottom > ClientRect.Bottom Then Exit For
    .Bottom = .Bottom + GetRowHeight(iRow)
    If Visibility = FlexVisibilityCompleteOnly Then If .Bottom > ClientRect.Bottom Then Exit For
    Count = Count + 1
Next iRow
FixedRowsVisible = Count
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
If Index > (PropFixedCols - 1) Then
    For i = 0 To (PropFixedCols - 1)
        If i < Index Then Value = Value + GetColWidth(i)
    Next i
    For i = VBFlexGridLeftCol To (Index - 1)
        Value = Value + GetColWidth(i)
    Next i
    If Index < VBFlexGridLeftCol Then
        For i = PropFixedCols To (VBFlexGridLeftCol - 1)
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
.Flags = RCPF_CHECKLEFTCOL Or RCPF_SETSCROLLBARS
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
.Flags = RCPF_CHECKLEFTCOL Or RCPF_SETSCROLLBARS
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
ColHidden = VBFlexGridColsInfo(Index).Hidden
End Property

Public Property Let ColHidden(ByVal Index As Long, ByVal Value As Boolean)
If Index <> -1 And (Index < 0 Or Index > (PropCols - 1)) Then Err.Raise Number:=30010, Description:="Invalid Col value"
If Index > -1 Then
    VBFlexGridColsInfo(Index).Hidden = Value
Else
    Dim i As Long
    For i = 0 To (PropCols - 1)
        VBFlexGridColsInfo(i).Hidden = Value
    Next i
End If
Dim RCP As TROWCOLPARAMS
With RCP
.Mask = RCPM_LEFTCOL
.Flags = RCPF_CHECKLEFTCOL Or RCPF_SETSCROLLBARS
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
    Dim ClientRect As RECT, GridRect As RECT, iCol As Long
    GetClientRect VBFlexGridHandle, ClientRect
    With GridRect
    If Index <= (PropFixedCols - 1) Then
        ColIsVisible = True
        For iCol = 0 To (PropFixedCols - 1)
            If Visibility = FlexVisibilityCompleteOnly Then .Right = .Right + GetColWidth(iCol)
            If .Right > ClientRect.Right Then
                ColIsVisible = False
                Exit For
            End If
            If Visibility = FlexVisibilityPartialOK Then .Right = .Right + GetColWidth(iCol)
            If iCol >= Index Then Exit For
        Next iCol
    ElseIf Index >= VBFlexGridLeftCol Then
        ColIsVisible = True
        For iCol = 0 To (PropFixedCols - 1)
            .Right = .Right + GetColWidth(iCol)
        Next iCol
        For iCol = VBFlexGridLeftCol To (PropCols - 1)
            If Visibility = FlexVisibilityCompleteOnly Then .Right = .Right + GetColWidth(iCol)
            If .Right > ClientRect.Right Then
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
Dim ClientRect As RECT, GridRect As RECT, iCol As Long, Count As Long
GetClientRect VBFlexGridHandle, ClientRect
With GridRect
For iCol = 0 To (PropFixedCols - 1)
    If Visibility = FlexVisibilityPartialOK Then If .Right > ClientRect.Right Then Exit For
    .Right = .Right + GetColWidth(iCol)
    If Visibility = FlexVisibilityCompleteOnly Then If .Right > ClientRect.Right Then Exit For
    Count = Count + 1
Next iCol
For iCol = VBFlexGridLeftCol To (PropCols - 1)
    If Visibility = FlexVisibilityPartialOK Then If .Right > ClientRect.Right Then Exit For
    .Right = .Right + GetColWidth(iCol)
    If Visibility = FlexVisibilityCompleteOnly Then If .Right > ClientRect.Right Then Exit For
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
Dim ClientRect As RECT, GridRect As RECT, iCol As Long, Count As Long
GetClientRect VBFlexGridHandle, ClientRect
With GridRect
For iCol = 0 To (PropFixedCols - 1)
    If Visibility = FlexVisibilityPartialOK Then If .Right > ClientRect.Right Then Exit For
    .Right = .Right + GetColWidth(iCol)
    If Visibility = FlexVisibilityCompleteOnly Then If .Right > ClientRect.Right Then Exit For
    Count = Count + 1
Next iCol
FixedColsVisible = Count
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
If Index < 0 Or Index > (PropCols - 1) Then Err.Raise Number:=30004, Description:="Invalid Col value for alignment"
ColSort = VBFlexGridColsInfo(Index).Sort
End Property

Public Property Let ColSort(ByVal Index As Long, ByVal Value As FlexSortConstants)
If Index <> -1 And (Index < 0 Or Index > (PropCols - 1)) Then Err.Raise Number:=30004, Description:="Invalid Col value for alignment"
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

Public Property Get MergeRow(ByVal Index As Long) As Boolean
Attribute MergeRow.VB_Description = "Returns/sets which columns or rows should have their contents merged when the merge cells property is set to a value other than 0 - Never."
Attribute MergeRow.VB_MemberFlags = "400"
If Index < 0 Or Index > (PropRows - 1) Then Err.Raise Number:=30009, Description:="Invalid Row value"
MergeRow = VBFlexGridCells.Rows(Index).RowInfo.Merge
End Property

Public Property Let MergeRow(ByVal Index As Long, ByVal Value As Boolean)
If Index < 0 Or Index > (PropRows - 1) Then Err.Raise Number:=30009, Description:="Invalid Row value"
VBFlexGridCells.Rows(Index).RowInfo.Merge = Value
Call RedrawGrid
End Property

Public Property Get MergeCol(ByVal Index As Long) As Boolean
Attribute MergeCol.VB_Description = "Returns/sets which columns or rows should have their contents merged when the merge cells property is set to a value other than 0 - Never."
Attribute MergeCol.VB_MemberFlags = "400"
If Index < 0 Or Index > (PropCols - 1) Then Err.Raise Number:=30010, Description:="Invalid Col value"
MergeCol = VBFlexGridColsInfo(Index).Merge
End Property

Public Property Let MergeCol(ByVal Index As Long, ByVal Value As Boolean)
If Index < 0 Or Index > (PropCols - 1) Then Err.Raise Number:=30010, Description:="Invalid Col value"
VBFlexGridColsInfo(Index).Merge = Value
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
If Err.Number = 0 Then
    Call RedrawGrid
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
If VBFlexGridRow > -1 And VBFlexGridCol > -1 Then Text = VBFlexGridCells.Rows(VBFlexGridRow).Cols(VBFlexGridCol).Text
End Property

Public Property Let Text(ByVal Value As String)
If VBFlexGridRow < 0 Then
    Err.Raise Number:=30009, Description:="Invalid Row value"
ElseIf VBFlexGridCol < 0 Then
    Err.Raise Number:=30010, Description:="Invalid Col value"
End If
If PropFillStyle = FlexFillStyleSingle Then
    VBFlexGridCells.Rows(VBFlexGridRow).Cols(VBFlexGridCol).Text = Value
ElseIf PropFillStyle = FlexFillStyleRepeat Then
    Dim i As Long, j As Long, SelRange As TSELRANGE
    Call GetSelRangeStruct(SelRange)
    For i = SelRange.TopRow To SelRange.BottomRow
        With VBFlexGridCells.Rows(i)
        For j = SelRange.LeftCol To SelRange.RightCol
            .Cols(j).Text = Value
        Next j
        End With
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
TextArray = VBFlexGridCells.Rows(Fix(RetVal)).Cols(((RetVal - Fix(RetVal)) * PropCols)).Text
End Property

Public Property Let TextArray(ByVal Index As Long, ByVal Value As String)
If (Index < 0 Or Index > ((PropRows * PropCols) - 1)) Then Err.Raise Number:=381, Description:="Subscript out of range"
Dim RetVal As Double
RetVal = Index / PropCols
VBFlexGridCells.Rows(Fix(RetVal)).Cols(((RetVal - Fix(RetVal)) * PropCols)).Text = Value
Call RedrawGrid
End Property

Public Property Get TextMatrix(ByVal Row As Long, ByVal Col As Long) As String
Attribute TextMatrix.VB_Description = "Returns/sets the text contents of an arbitrary cell (row/col subscripts)."
Attribute TextMatrix.VB_MemberFlags = "400"
If (Row < 0 Or Row > (PropRows - 1)) Or (Col < 0 Or Col > (PropCols - 1)) Then Err.Raise Number:=381, Description:="Subscript out of range"
TextMatrix = VBFlexGridCells.Rows(Row).Cols(Col).Text
End Property

Public Property Let TextMatrix(ByVal Row As Long, ByVal Col As Long, ByVal Value As String)
If (Row < 0 Or Row > (PropRows - 1)) Or (Col < 0 Or Col > (PropCols - 1)) Then Err.Raise Number:=381, Description:="Subscript out of range"
VBFlexGridCells.Rows(Row).Cols(Col).Text = Value
Call RedrawGrid
End Property

Public Property Get Clip() As String
Attribute Clip.VB_Description = "Returns/sets the contents of the cells in a selected region."
Attribute Clip.VB_MemberFlags = "400"
If VBFlexGridRow < 0 Or VBFlexGridCol < 0 Then Err.Raise 7
Dim i As Long, j As Long, SelRange As TSELRANGE, Buffer As String
Dim ColSeparator As String, RowSeparator As String
ColSeparator = GetColSeparator()
RowSeparator = GetRowSeparator()
Call GetSelRangeStruct(SelRange)
For i = SelRange.TopRow To SelRange.BottomRow
    With VBFlexGridCells.Rows(i)
    For j = SelRange.LeftCol To SelRange.RightCol
        Buffer = Buffer & .Cols(j).Text
        If Len(Buffer) > 1000 Then Clip = Clip & Buffer: Buffer = vbNullString
        If j < SelRange.RightCol Then Buffer = Buffer & ColSeparator
    Next j
    If i < SelRange.BottomRow Then Buffer = Buffer & RowSeparator
    End With
Next i
If Len(Buffer) > 0 Then Clip = Clip & Buffer
End Property

Public Property Let Clip(ByVal Value As String)
If VBFlexGridRow < 0 Then
    Err.Raise Number:=30009, Description:="Invalid Row value"
ElseIf VBFlexGridCol < 0 Then
    Err.Raise Number:=30010, Description:="Invalid Col value"
End If
Dim SelRange As TSELRANGE, Temp As String, iRow As Long, iCol As Long
Dim Pos1 As Long, Pos2 As Long, Pos3 As Long, Pos4 As Long
Dim ColSeparator As String, RowSeparator As String
Call GetSelRangeStruct(SelRange)
ColSeparator = GetColSeparator()
RowSeparator = GetRowSeparator()
With VBFlexGridCells
Do
    Pos1 = InStr(Pos1 + 1, Value, RowSeparator)
    If Pos1 > 0 Then
        If (SelRange.TopRow + iRow) <= SelRange.BottomRow Then
            Temp = Mid$(Value, Pos2 + 1, Pos1 - Pos2 - 1)
            With .Rows(SelRange.TopRow + iRow)
            Do
                Pos3 = InStr(Pos3 + 1, Temp, ColSeparator)
                If Pos3 > 0 Then
                    If (SelRange.LeftCol + iCol) <= SelRange.RightCol Then .Cols(SelRange.LeftCol + iCol).Text = Mid$(Temp, Pos4 + 1, Pos3 - Pos4 - 1)
                Else
                    If (SelRange.LeftCol + iCol) <= SelRange.RightCol Then .Cols(SelRange.LeftCol + iCol).Text = Mid$(Temp, Pos4 + 1)
                End If
                Pos4 = Pos3
                iCol = iCol + 1
            Loop Until Pos3 = 0
            End With
        End If
    Else
        If (SelRange.TopRow + iRow) <= SelRange.BottomRow Then
            Temp = Mid$(Value, Pos2 + 1)
            With .Rows(SelRange.TopRow + iRow)
            Do
                Pos3 = InStr(Pos3 + 1, Temp, ColSeparator)
                If Pos3 > 0 Then
                    If (SelRange.LeftCol + iCol) <= SelRange.RightCol Then .Cols(SelRange.LeftCol + iCol).Text = Mid$(Temp, Pos4 + 1, Pos3 - Pos4 - 1)
                Else
                    If (SelRange.LeftCol + iCol) <= SelRange.RightCol Then .Cols(SelRange.LeftCol + iCol).Text = Mid$(Temp, Pos4 + 1)
                End If
                Pos4 = Pos3
                iCol = iCol + 1
            Loop Until Pos3 = 0
            End With
        End If
    End If
    Pos2 = Pos1
    Pos4 = 0
    iRow = iRow + 1
    iCol = 0
Loop Until Pos1 = 0
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
    Dim i As Long, j As Long, SelRange As TSELRANGE
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
    Dim i As Long, j As Long, SelRange As TSELRANGE
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

Public Property Get CellPicture() As IPicture
Attribute CellPicture.VB_Description = "Returns/sets an picture to be displayed in the current cell or in a range of cells."
Attribute CellPicture.VB_MemberFlags = "400"
If VBFlexGridRow < 0 Then
    Err.Raise Number:=30009, Description:="Invalid Row value"
ElseIf VBFlexGridCol < 0 Then
    Err.Raise Number:=30010, Description:="Invalid Col value"
End If
Set CellPicture = VBFlexGridCells.Rows(VBFlexGridRow).Cols(VBFlexGridCol).Picture
End Property

Public Property Let CellPicture(ByVal Value As IPicture)
Set Me.CellPicture = Value
End Property

Public Property Set CellPicture(ByVal Value As IPicture)
If VBFlexGridRow < 0 Then
    Err.Raise Number:=30009, Description:="Invalid Row value"
ElseIf VBFlexGridCol < 0 Then
    Err.Raise Number:=30010, Description:="Invalid Col value"
End If
Set UserControl.Picture = Value
If PropFillStyle = FlexFillStyleSingle Then
    Set VBFlexGridCells.Rows(VBFlexGridRow).Cols(VBFlexGridCol).Picture = UserControl.Picture
ElseIf PropFillStyle = FlexFillStyleRepeat Then
    Dim i As Long, j As Long, SelRange As TSELRANGE
    Call GetSelRangeStruct(SelRange)
    For i = SelRange.TopRow To SelRange.BottomRow
        With VBFlexGridCells.Rows(i)
        For j = SelRange.LeftCol To SelRange.RightCol
            Set .Cols(j).Picture = UserControl.Picture
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
    Case FlexPictureAlignmentLeftTop, FlexPictureAlignmentLeftCenter, FlexPictureAlignmentLeftBottom, FlexPictureAlignmentCenterTop, FlexPictureAlignmentCenterCenter, FlexPictureAlignmentCenterBottom, FlexPictureAlignmentRightTop, FlexPictureAlignmentRightCenter, FlexPictureAlignmentRightBottom, FlexPictureAlignmentStretch, FlexPictureAlignmentTile
    Case Else
        Err.Raise Number:=30005, Description:="Invalid Alignment value"
End Select
If PropFillStyle = FlexFillStyleSingle Then
    VBFlexGridCells.Rows(VBFlexGridRow).Cols(VBFlexGridCol).PictureAlignment = Value
ElseIf PropFillStyle = FlexFillStyleRepeat Then
    Dim i As Long, j As Long, SelRange As TSELRANGE
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
    Dim i As Long, j As Long, SelRange As TSELRANGE
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
    Dim i As Long, j As Long, SelRange As TSELRANGE
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
    Dim i As Long, j As Long, SelRange As TSELRANGE
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
        Dim TempFont As StdFont
        Set TempFont = New StdFont
        TempFont.Name = Value
        .FontName = TempFont.Name
        .FontCharset = TempFont.Charset
    Else
        .FontName = vbNullString
    End If
    End With
ElseIf PropFillStyle = FlexFillStyleRepeat Then
    Dim i As Long, j As Long, SelRange As TSELRANGE
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
                        .FontCharset = PropFont.Charset
                    Else
                        .FontSize = PropFontFixed.Size
                        .FontBold = PropFontFixed.Bold
                        .FontItalic = PropFontFixed.Italic
                        .FontStrikeThrough = PropFontFixed.Strikethrough
                        .FontUnderline = PropFontFixed.Underline
                        .FontCharset = PropFontFixed.Charset
                    End If
                End If
                .FontName = Value
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
    Dim i As Long, j As Long, SelRange As TSELRANGE
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
    Dim i As Long, j As Long, SelRange As TSELRANGE
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
    Dim i As Long, j As Long, SelRange As TSELRANGE
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
    Dim i As Long, j As Long, SelRange As TSELRANGE
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
    Dim i As Long, j As Long, SelRange As TSELRANGE
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
    Dim i As Long, j As Long, SelRange As TSELRANGE
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

Public Sub CellEnsureVisible(Optional ByVal Visibility As FlexVisibilityConstants = FlexVisibilityCompleteOnly)
Attribute CellEnsureVisible.VB_Description = "Ensures the current cell is visible, scrolling the control if necessary."
Select Case Visibility
    Case FlexVisibilityPartialOK, FlexVisibilityCompleteOnly
    Case Else
        Err.Raise 380
End Select
If PropRows < 1 Or PropCols < 1 Then Exit Sub
If Visibility = FlexVisibilityPartialOK Then
    If Me.RowIsVisible(VBFlexGridRow, FlexVisibilityPartialOK) = True And Me.ColIsVisible(VBFlexGridCol, FlexVisibilityPartialOK) = True Then Exit Sub
End If
Dim RCP As TROWCOLPARAMS
With RCP
.Mask = RCPM_TOPROW Or RCPM_LEFTCOL
.TopRow = VBFlexGridTopRow
.LeftCol = VBFlexGridLeftCol
If .TopRow > VBFlexGridRow Then
    .TopRow = VBFlexGridRow
ElseIf VBFlexGridRow > (.TopRow + GetRowsPerPage(.TopRow) - 1) Then
    .TopRow = VBFlexGridRow - GetRowsPerPageRev(VBFlexGridRow) + 1
End If
If .LeftCol > VBFlexGridCol Then
    .LeftCol = VBFlexGridCol
ElseIf VBFlexGridCol > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
    .LeftCol = VBFlexGridCol - GetColsPerPageRev(VBFlexGridCol) + 1
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
CellWidth = UserControl.ScaleX((CellRect.Right - CellRect.Left), vbPixels, vbTwips)
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
CellHeight = UserControl.ScaleY((CellRect.Bottom - CellRect.Top), vbPixels, vbTwips)
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
VBFlexGridHitResult = .HitResult
End With
End Sub

Public Function FindItem(ByVal Text As String, Optional ByVal Row As Long = -1, Optional ByVal Col As Long = -1, Optional ByVal Partial As Boolean, Optional ByVal CaseSensitive As Boolean) As Long
Attribute FindItem.VB_Description = "Finds an item in the flex grid and returns the index of that item."
If Row < -1 Then Err.Raise 380
If Col < -1 Then Err.Raise 380
If Row = -1 Then Row = PropFixedRows
If Col = -1 Then Col = PropFixedCols
If (Row < 0 Or Row > (PropRows - 1)) Or (Col < 0 Or Col > (PropCols - 1)) Then Err.Raise Number:=381, Description:="Subscript out of range"
Dim iRow As Long, Compare As VbCompareMethod
FindItem = -1
If CaseSensitive = False Then Compare = vbTextCompare Else Compare = vbBinaryCompare
With VBFlexGridCells
If Partial = False Then
    For iRow = Row To (PropRows - 1)
        If StrComp(.Rows(iRow).Cols(Col).Text, Text, Compare) = 0 Then
            FindItem = iRow
            Exit For
        End If
    Next iRow
Else
    For iRow = Row To (PropRows - 1)
        If InStr(1, .Rows(iRow).Cols(Col).Text, Text, Compare) > 0 Then
            FindItem = iRow
            Exit For
        End If
    Next iRow
End If
End With
End Function

Public Sub AutoSize(ByVal RowOrCol1 As Long, Optional ByVal RowOrCol2 As Long = -1, Optional ByVal Mode As FlexAutoSizeModeConstants, Optional ByVal Scope As FlexAutoSizeScopeConstants, Optional ByVal Equal As Boolean, Optional ByVal ExtraSpace As Long)
Attribute AutoSize.VB_Description = "Automatically sizes column widths or row heights to fit cell contents."
If RowOrCol2 < -1 Then Err.Raise 380
If RowOrCol2 = -1 Then RowOrCol2 = RowOrCol1
Select Case Mode
    Case FlexAutoSizeModeColWidth, FlexAutoSizeModeRowHeight
    Case Else
        Err.Raise 380
End Select
Select Case Scope
    Case FlexAutoSizeScopeAll, FlexAutoSizeScopeFixed, FlexAutoSizeScopeScrollable
    Case Else
        Err.Raise 380
End Select
If Mode = FlexAutoSizeModeColWidth Then
    If (RowOrCol1 < 0 Or RowOrCol1 > (PropCols - 1)) Or (RowOrCol2 < 0 Or RowOrCol2 > (PropCols - 1)) Then Err.Raise Number:=381, Description:="Subscript out of range"
ElseIf Mode = FlexAutoSizeModeRowHeight Then
    If (RowOrCol1 < 0 Or RowOrCol1 > (PropRows - 1)) Or (RowOrCol2 < 0 Or RowOrCol2 > (PropRows - 1)) Then Err.Raise Number:=381, Description:="Subscript out of range"
End If
Dim iRow As Long, iCol As Long, Spacing As Long, Size As SIZEAPI, EqualSize As SIZEAPI
If Mode = FlexAutoSizeModeColWidth Then
    Spacing = (COLINFO_WIDTH_SPACING_DIP * PixelsPerDIP_X()) + CLng(UserControl.ScaleX(ExtraSpace, vbTwips, vbPixels))
    EqualSize.CX = -1
    Select Case Scope
        Case FlexAutoSizeScopeAll
            For iCol = RowOrCol1 To RowOrCol2 Step IIf(RowOrCol2 >= RowOrCol1, 1, -1)
                With VBFlexGridColsInfo(iCol)
                .Width = -1
                For iRow = 0 To (PropRows - 1)
                    Size.CX = GetTextSize(iRow, iCol, VBFlexGridCells.Rows(iRow).Cols(iCol).Text).CX
                    If Size.CX > 0 Then
                        Size.CX = Size.CX + Spacing
                        If Size.CX > .Width Then .Width = Size.CX
                        If Size.CX > EqualSize.CX Then EqualSize.CX = Size.CX
                    End If
                Next iRow
                End With
            Next iCol
        Case FlexAutoSizeScopeFixed
            For iCol = RowOrCol1 To RowOrCol2 Step IIf(RowOrCol2 >= RowOrCol1, 1, -1)
                With VBFlexGridColsInfo(iCol)
                .Width = -1
                For iRow = 0 To (PropFixedRows - 1)
                    Size.CX = GetTextSize(iRow, iCol, VBFlexGridCells.Rows(iRow).Cols(iCol).Text).CX
                    If Size.CX > 0 Then
                        Size.CX = Size.CX + Spacing
                        If Size.CX > .Width Then .Width = Size.CX
                        If Size.CX > EqualSize.CX Then EqualSize.CX = Size.CX
                    End If
                Next iRow
                End With
            Next iCol
        Case FlexAutoSizeScopeScrollable
            For iCol = RowOrCol1 To RowOrCol2 Step IIf(RowOrCol2 >= RowOrCol1, 1, -1)
                With VBFlexGridColsInfo(iCol)
                .Width = -1
                For iRow = PropFixedRows To (PropRows - 1)
                    Size.CX = GetTextSize(iRow, iCol, VBFlexGridCells.Rows(iRow).Cols(iCol).Text).CX
                    If Size.CX > 0 Then
                        Size.CX = Size.CX + Spacing
                        If Size.CX > .Width Then .Width = Size.CX
                        If Size.CX > EqualSize.CX Then EqualSize.CX = Size.CX
                    End If
                Next iRow
                End With
            Next iCol
    End Select
    If Equal = True Then
        For iCol = RowOrCol1 To RowOrCol2 Step IIf(RowOrCol2 >= RowOrCol1, 1, -1)
            With VBFlexGridColsInfo(iCol)
            .Width = EqualSize.CX
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
                .Height = -1
                For iCol = 0 To (PropCols - 1)
                    Size.CY = GetTextSize(iRow, iCol, VBFlexGridCells.Rows(iRow).Cols(iCol).Text).CY
                    If Size.CY > 0 Then
                        Size.CY = Size.CY + Spacing
                        If Size.CY > .Height Then .Height = Size.CY
                        If Size.CY > EqualSize.CY Then EqualSize.CY = Size.CY
                    End If
                Next iCol
                End With
            Next iRow
        Case FlexAutoSizeScopeFixed
            For iRow = RowOrCol1 To RowOrCol2 Step IIf(RowOrCol2 >= RowOrCol1, 1, -1)
                With VBFlexGridCells.Rows(iRow).RowInfo
                .Height = -1
                For iCol = 0 To (PropFixedCols - 1)
                    Size.CY = GetTextSize(iRow, iCol, VBFlexGridCells.Rows(iRow).Cols(iCol).Text).CY
                    If Size.CY > 0 Then
                        Size.CY = Size.CY + Spacing
                        If Size.CY > .Height Then .Height = Size.CY
                        If Size.CY > EqualSize.CY Then EqualSize.CY = Size.CY
                    End If
                Next iCol
                End With
            Next iRow
        Case FlexAutoSizeScopeScrollable
            For iRow = RowOrCol1 To RowOrCol2 Step IIf(RowOrCol2 >= RowOrCol1, 1, -1)
                With VBFlexGridCells.Rows(iRow).RowInfo
                .Height = -1
                For iCol = PropFixedCols To (PropCols - 1)
                    Size.CY = GetTextSize(iRow, iCol, VBFlexGridCells.Rows(iRow).Cols(iCol).Text).CY
                    If Size.CY > 0 Then
                        Size.CY = Size.CY + Spacing
                        If Size.CY > .Height Then .Height = Size.CY
                        If Size.CY > EqualSize.CY Then EqualSize.CY = Size.CY
                    End If
                Next iCol
                End With
            Next iRow
    End Select
    If Equal = True Then
        For iRow = RowOrCol1 To RowOrCol2 Step IIf(RowOrCol2 >= RowOrCol1, 1, -1)
            With VBFlexGridCells.Rows(iRow).RowInfo
            .Height = EqualSize.CY
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
.Flags = .Flags Or RCPF_SETSCROLLBARS
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
            Dim hRgn As Long, RC As RECT
            Call DrawGrid(0, hRgn, True)
            If hRgn <> 0 Then
                GetRgnBox hRgn, RC
                DeleteObject hRgn
            End If
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

Public Property Get Version() As Integer
Attribute Version.VB_Description = "Returns the version of the flex grid control currently loaded in memory."
Attribute Version.VB_MemberFlags = "400"
Version = 600
End Property

Private Sub InitFlexGridCells()
If (0 / 1) + (Not Not VBFlexGridCells.Rows()) <> 0 Then Exit Sub
If PropRows < 1 Or PropCols < 1 Then
    VBFlexGridRow = -1
    VBFlexGridCol = -1
    VBFlexGridRowSel = VBFlexGridRow
    VBFlexGridColSel = VBFlexGridCol
    VBFlexGridTopRow = VBFlexGridRow
    VBFlexGridLeftCol = VBFlexGridCol
    Exit Sub
End If
Dim i As Long, j As Long
ReDim VBFlexGridCells.Rows(0 To (PropRows - 1)) As TCOLS
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
        Case FlexSelectionModeFree
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
VBFlexGridTopRow = VBFlexGridRow
VBFlexGridLeftCol = VBFlexGridCol
End Sub

Private Sub EraseFlexGridCells()
If (0 / 1) + (Not Not VBFlexGridCells.Rows()) = 0 Then Exit Sub
Erase VBFlexGridCells.Rows()
Erase VBFlexGridColsInfo()
Erase VBFlexGridDefaultCols.Cols()
VBFlexGridRow = -1
VBFlexGridCol = -1
VBFlexGridRowSel = VBFlexGridRow
VBFlexGridColSel = VBFlexGridCol
VBFlexGridTopRow = VBFlexGridRow
VBFlexGridLeftCol = VBFlexGridCol
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
End If
If VBFlexGridHandle = 0 Or (PropRows < 1 Or PropCols < 1) Then Exit Sub
Dim iRow As Long, iCol As Long
Dim ClientRect As RECT, CellRect As RECT, GridRect As RECT
GetClientRect VBFlexGridHandle, ClientRect
With CellRect
If PropMergeCells = FlexMergeCellsNever Then
    For iRow = 0 To (PropRows - 1)
        If iRow >= VBFlexGridTopRow Then
            .Bottom = .Top + GetRowHeight(iRow)
            For iCol = 0 To (PropCols - 1)
                If iCol >= VBFlexGridLeftCol Then
                    .Left = .Right
                    .Right = .Right + GetColWidth(iCol)
                    If hDC <> 0 Then Call DrawCell(hDC, CellRect, iRow, iCol, False)
                ElseIf iCol < PropFixedCols Then
                    .Left = .Right
                    .Right = .Right + GetColWidth(iCol)
                    If hDC <> 0 Then Call DrawCell(hDC, CellRect, iRow, iCol, True)
                Else
                    iCol = VBFlexGridLeftCol - 1
                End If
                If NoClip = False And .Right > ClientRect.Right Then Exit For
            Next iCol
            If .Bottom > GridRect.Bottom Then GridRect.Bottom = .Bottom
            If .Right > GridRect.Right Then GridRect.Right = .Right
            .Left = 0
            .Right = 0
            .Top = .Top + GetRowHeight(iRow)
        ElseIf iRow < PropFixedRows Then
            .Bottom = .Top + GetRowHeight(iRow)
            For iCol = 0 To (PropCols - 1)
                If iCol >= VBFlexGridLeftCol Or iCol < PropFixedCols Then
                    .Left = .Right
                    .Right = .Right + GetColWidth(iCol)
                    If hDC <> 0 Then Call DrawCell(hDC, CellRect, iRow, iCol, True)
                Else
                    iCol = VBFlexGridLeftCol - 1
                End If
                If NoClip = False And .Right > ClientRect.Right Then Exit For
            Next iCol
            If .Bottom > GridRect.Bottom Then GridRect.Bottom = .Bottom
            If .Right > GridRect.Right Then GridRect.Right = .Right
            .Left = 0
            .Right = 0
            .Top = .Top + GetRowHeight(iRow)
        Else
            iRow = VBFlexGridTopRow - 1
        End If
        If NoClip = False And .Bottom > ClientRect.Bottom Then Exit For
    Next iRow
Else
    ReDim VBFlexGridMergeDrawInfo.Row.Cols(0 To (PropCols - 1)) As TMERGEDRAWCOLINFO
    For iRow = 0 To (PropRows - 1)
        VBFlexGridMergeDrawInfo.Row.ColOffset = 0
        VBFlexGridMergeDrawInfo.Row.Width = 0
        If iRow >= VBFlexGridTopRow Then
            .Bottom = .Top + GetRowHeight(iRow)
            For iCol = 0 To (PropCols - 1)
                If iCol >= VBFlexGridLeftCol Then
                    .Left = .Right
                    .Right = .Right + GetColWidth(iCol)
                    If VBFlexGridCells.Rows(iRow).RowInfo.Merge = True Then
                        If iCol > VBFlexGridLeftCol Then
                            Select Case PropMergeCells
                                Case FlexMergeCellsFree, FlexMergeCellsRestrictRows
                                    If MergeCompareFunction(VBFlexGridCells.Rows(iRow).Cols(iCol).Text, VBFlexGridCells.Rows(iRow).Cols(iCol - 1).Text) = True Then
                                        VBFlexGridMergeDrawInfo.Row.ColOffset = VBFlexGridMergeDrawInfo.Row.ColOffset + 1
                                        VBFlexGridMergeDrawInfo.Row.Width = VBFlexGridMergeDrawInfo.Row.Width + GetColWidth(iCol - 1)
                                    Else
                                        VBFlexGridMergeDrawInfo.Row.ColOffset = 0
                                        VBFlexGridMergeDrawInfo.Row.Width = 0
                                    End If
                                Case FlexMergeCellsRestrictColumns, FlexMergeCellsRestrictAll
                                    If MergeCompareFunction(VBFlexGridCells.Rows(iRow).Cols(iCol).Text, VBFlexGridCells.Rows(iRow).Cols(iCol - 1).Text) = True Then
                                        If iRow > VBFlexGridTopRow Then
                                            If MergeCompareFunction(VBFlexGridCells.Rows(iRow - 1).Cols(iCol).Text, VBFlexGridCells.Rows(iRow - 1).Cols(iCol - 1).Text) = True Then
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
                    If VBFlexGridColsInfo(iCol).Merge = True Then
                        If iRow > VBFlexGridTopRow Then
                            Select Case PropMergeCells
                                Case FlexMergeCellsFree, FlexMergeCellsRestrictColumns
                                    If MergeCompareFunction(VBFlexGridCells.Rows(iRow).Cols(iCol).Text, VBFlexGridCells.Rows(iRow - 1).Cols(iCol).Text) = True Then
                                        VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset + 1
                                        VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height + GetRowHeight(iRow - 1)
                                    Else
                                        VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = 0
                                        VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = 0
                                    End If
                                Case FlexMergeCellsRestrictRows, FlexMergeCellsRestrictAll
                                    If MergeCompareFunction(VBFlexGridCells.Rows(iRow).Cols(iCol).Text, VBFlexGridCells.Rows(iRow - 1).Cols(iCol).Text) = True Then
                                        If iCol > VBFlexGridLeftCol Then
                                            If MergeCompareFunction(VBFlexGridCells.Rows(iRow).Cols(iCol - 1).Text, VBFlexGridCells.Rows(iRow - 1).Cols(iCol - 1).Text) = True Then
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
                    If hDC <> 0 Then Call DrawCell(hDC, CellRect, iRow, iCol, False)
                    .Left = .Left + VBFlexGridMergeDrawInfo.Row.Width
                    .Top = .Top + VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height
                ElseIf iCol < PropFixedCols Then
                    .Left = .Right
                    .Right = .Right + GetColWidth(iCol)
                    If VBFlexGridCells.Rows(iRow).RowInfo.Merge = True Then
                        If iCol > 0 Then
                            Select Case PropMergeCells
                                Case FlexMergeCellsFree, FlexMergeCellsRestrictRows, FlexMergeCellsFixedOnly
                                    If MergeCompareFunction(VBFlexGridCells.Rows(iRow).Cols(iCol).Text, VBFlexGridCells.Rows(iRow).Cols(iCol - 1).Text) = True Then
                                        VBFlexGridMergeDrawInfo.Row.ColOffset = VBFlexGridMergeDrawInfo.Row.ColOffset + 1
                                        VBFlexGridMergeDrawInfo.Row.Width = VBFlexGridMergeDrawInfo.Row.Width + GetColWidth(iCol - 1)
                                    Else
                                        VBFlexGridMergeDrawInfo.Row.ColOffset = 0
                                        VBFlexGridMergeDrawInfo.Row.Width = 0
                                    End If
                                Case FlexMergeCellsRestrictColumns, FlexMergeCellsRestrictAll
                                    If MergeCompareFunction(VBFlexGridCells.Rows(iRow).Cols(iCol).Text, VBFlexGridCells.Rows(iRow).Cols(iCol - 1).Text) = True Then
                                        If iRow > VBFlexGridTopRow Then
                                            If MergeCompareFunction(VBFlexGridCells.Rows(iRow - 1).Cols(iCol).Text, VBFlexGridCells.Rows(iRow - 1).Cols(iCol - 1).Text) = True Then
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
                    If VBFlexGridColsInfo(iCol).Merge = True Then
                        If iRow > 0 Then
                            Select Case PropMergeCells
                                Case FlexMergeCellsFree, FlexMergeCellsRestrictColumns, FlexMergeCellsFixedOnly
                                    If MergeCompareFunction(VBFlexGridCells.Rows(iRow).Cols(iCol).Text, VBFlexGridCells.Rows(iRow - 1).Cols(iCol).Text) = True Then
                                        VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset + 1
                                        VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height + GetRowHeight(iRow - 1)
                                    Else
                                        VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = 0
                                        VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = 0
                                    End If
                                Case FlexMergeCellsRestrictRows, FlexMergeCellsRestrictAll
                                    If MergeCompareFunction(VBFlexGridCells.Rows(iRow).Cols(iCol).Text, VBFlexGridCells.Rows(iRow - 1).Cols(iCol).Text) = True Then
                                        If iCol > 0 Then
                                            If MergeCompareFunction(VBFlexGridCells.Rows(iRow).Cols(iCol - 1).Text, VBFlexGridCells.Rows(iRow - 1).Cols(iCol - 1).Text) = True Then
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
                    If hDC <> 0 Then Call DrawCell(hDC, CellRect, iRow, iCol, True)
                    .Left = .Left + VBFlexGridMergeDrawInfo.Row.Width
                    .Top = .Top + VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height
                Else
                    iCol = VBFlexGridLeftCol - 1
                End If
                If NoClip = False And .Right > ClientRect.Right Then Exit For
            Next iCol
            If .Bottom > GridRect.Bottom Then GridRect.Bottom = .Bottom
            If .Right > GridRect.Right Then GridRect.Right = .Right
            .Left = 0
            .Right = 0
            .Top = .Top + GetRowHeight(iRow)
        ElseIf iRow < PropFixedRows Then
            .Bottom = .Top + GetRowHeight(iRow)
            For iCol = 0 To (PropCols - 1)
                If iCol >= VBFlexGridLeftCol Or iCol < PropFixedCols Then
                    .Left = .Right
                    .Right = .Right + GetColWidth(iCol)
                    If VBFlexGridCells.Rows(iRow).RowInfo.Merge = True Then
                        If iCol > VBFlexGridLeftCol Or (iCol > 0 And iCol < PropFixedCols) Then
                            Select Case PropMergeCells
                                Case FlexMergeCellsFree, FlexMergeCellsRestrictRows, FlexMergeCellsFixedOnly
                                    If MergeCompareFunction(VBFlexGridCells.Rows(iRow).Cols(iCol).Text, VBFlexGridCells.Rows(iRow).Cols(iCol - 1).Text) = True Then
                                        VBFlexGridMergeDrawInfo.Row.ColOffset = VBFlexGridMergeDrawInfo.Row.ColOffset + 1
                                        VBFlexGridMergeDrawInfo.Row.Width = VBFlexGridMergeDrawInfo.Row.Width + GetColWidth(iCol - 1)
                                    Else
                                        VBFlexGridMergeDrawInfo.Row.ColOffset = 0
                                        VBFlexGridMergeDrawInfo.Row.Width = 0
                                    End If
                                Case FlexMergeCellsRestrictColumns, FlexMergeCellsRestrictAll
                                    If MergeCompareFunction(VBFlexGridCells.Rows(iRow).Cols(iCol).Text, VBFlexGridCells.Rows(iRow).Cols(iCol - 1).Text) = True Then
                                        If iRow > 0 Then
                                            If MergeCompareFunction(VBFlexGridCells.Rows(iRow - 1).Cols(iCol).Text, VBFlexGridCells.Rows(iRow - 1).Cols(iCol - 1).Text) = True Then
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
                    If VBFlexGridColsInfo(iCol).Merge = True Then
                        If iRow > 0 Then
                            Select Case PropMergeCells
                                Case FlexMergeCellsFree, FlexMergeCellsRestrictColumns, FlexMergeCellsFixedOnly
                                    If MergeCompareFunction(VBFlexGridCells.Rows(iRow).Cols(iCol).Text, VBFlexGridCells.Rows(iRow - 1).Cols(iCol).Text) = True Then
                                        VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset + 1
                                        VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height + GetRowHeight(iRow - 1)
                                    Else
                                        VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset = 0
                                        VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height = 0
                                    End If
                                Case FlexMergeCellsRestrictRows, FlexMergeCellsRestrictAll
                                    If MergeCompareFunction(VBFlexGridCells.Rows(iRow).Cols(iCol).Text, VBFlexGridCells.Rows(iRow - 1).Cols(iCol).Text) = True Then
                                        If iCol > VBFlexGridLeftCol Or (iCol > 0 And iCol < PropFixedCols) Then
                                            If MergeCompareFunction(VBFlexGridCells.Rows(iRow).Cols(iCol - 1).Text, VBFlexGridCells.Rows(iRow - 1).Cols(iCol - 1).Text) = True Then
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
                    If hDC <> 0 Then Call DrawCell(hDC, CellRect, iRow, iCol, True)
                    .Left = .Left + VBFlexGridMergeDrawInfo.Row.Width
                    .Top = .Top + VBFlexGridMergeDrawInfo.Row.Cols(iCol).Height
                Else
                    iCol = VBFlexGridLeftCol - 1
                End If
                If NoClip = False And .Right > ClientRect.Right Then Exit For
            Next iCol
            If .Bottom > GridRect.Bottom Then GridRect.Bottom = .Bottom
            If .Right > GridRect.Right Then GridRect.Right = .Right
            .Left = 0
            .Right = 0
            .Top = .Top + GetRowHeight(iRow)
        Else
            iRow = VBFlexGridTopRow - 1
        End If
        If NoClip = False And .Bottom > ClientRect.Bottom Then Exit For
    Next iRow
    Erase VBFlexGridMergeDrawInfo.Row.Cols()
    VBFlexGridMergeDrawInfo.Row.ColOffset = 0
    VBFlexGridMergeDrawInfo.Row.Width = 0
End If
End With
With GridRect
If hDC <> 0 Then
    Dim hPenOld As Long, P As POINTAPI
    If VBFlexGridGridLineFixedPen <> 0 Then hPenOld = SelectObject(hDC, VBFlexGridGridLineFixedPen)
    MoveToEx hDC, .Left, .Bottom - 1, P
    LineTo hDC, .Right - 1, .Bottom - 1
    LineTo hDC, .Right - 1, .Top - 1
    MoveToEx hDC, P.X, P.Y, ByVal 0&
    If hPenOld <> 0 Then
        SelectObject hDC, hPenOld
        hPenOld = 0
    End If
End If
If hRgn <> -1 Then hRgn = CreateRectRgn(.Left, .Top, .Right, .Bottom)
End With
End Sub

Private Sub DrawCell(ByVal hDC As Long, ByRef CellRect As RECT, ByVal iRow As Long, ByVal iCol As Long, ByVal IsFixedCell As Boolean)
If (CellRect.Bottom - CellRect.Top) = 0 Or (CellRect.Right - CellRect.Left) = 0 Or hDC = 0 Then Exit Sub
Const ODS_SELECTED As Long = &H1, ODS_FOCUS As Long = &H10, ODS_NOFOCUSRECT As Long = &H200
Dim SelRange As TSELRANGE, ItemState As Long
Call GetSelRangeStruct(SelRange)
If PropMergeCells <> FlexMergeCellsNever Then
    If (VBFlexGridRow >= (iRow - VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset) And VBFlexGridRow <= iRow) And (VBFlexGridCol >= (iCol - VBFlexGridMergeDrawInfo.Row.ColOffset) And VBFlexGridCol <= iCol) Then
        iRow = VBFlexGridRow
        iCol = VBFlexGridCol
    Else
        iRow = iRow - VBFlexGridMergeDrawInfo.Row.Cols(iCol).RowOffset
        iCol = iCol - VBFlexGridMergeDrawInfo.Row.ColOffset
    End If
End If
Select Case PropHighLight
    Case FlexHighLightAlways
        If (iCol >= SelRange.LeftCol And iCol <= SelRange.RightCol) And (iRow >= SelRange.TopRow And iRow <= SelRange.BottomRow) Then ItemState = ItemState Or ODS_SELECTED
    Case FlexHighLightWithFocus
        If VBFlexGridFocused = True And (iCol >= SelRange.LeftCol And iCol <= SelRange.RightCol) And (iRow >= SelRange.TopRow And iRow <= SelRange.BottomRow) Then ItemState = ItemState Or ODS_SELECTED
End Select
If PropFocusRect <> FlexFocusRectNone Then
    If (iRow = VBFlexGridRow And iCol = VBFlexGridCol) Then ItemState = ItemState Or ODS_FOCUS
End If
If VBFlexGridFocused = False Then ItemState = ItemState Or ODS_NOFOCUSRECT
With VBFlexGridCells.Rows(iRow).Cols(iCol)
Dim hFontTemp As Long, hFontOld As Long
If .FontName = vbNullString Then
    If IsFixedCell = False Then
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
If Not (ItemState And ODS_SELECTED) = ODS_SELECTED Or (ItemState And ODS_FOCUS) = ODS_FOCUS Then
    If .BackColor = -1 Then
        If IsFixedCell = False Then
            If PropBackColor = PropBackColorAlt Then
                If VBFlexGridBackColorBrush <> 0 Then FillRect hDC, CellRect, VBFlexGridBackColorBrush
            Else
                If (iRow - PropFixedRows) Mod 2 = 0 Then
                    If VBFlexGridBackColorBrush <> 0 Then FillRect hDC, CellRect, VBFlexGridBackColorBrush
                Else
                    If VBFlexGridBackColorAltBrush <> 0 Then FillRect hDC, CellRect, VBFlexGridBackColorAltBrush
                End If
            End If
        Else
            If VBFlexGridBackColorFixedBrush <> 0 Then FillRect hDC, CellRect, VBFlexGridBackColorFixedBrush
        End If
    Else
        Dim Brush As Long
        Brush = CreateSolidBrush(WinColor(.BackColor))
        If Brush <> 0 Then
            FillRect hDC, CellRect, Brush
            DeleteObject Brush
        End If
    End If
Else
    If VBFlexGridBackColorSelBrush <> 0 Then FillRect hDC, CellRect, VBFlexGridBackColorSelBrush
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
            Case FlexPictureAlignmentLeftCenter
                PictureOffsetY = (((CellRect.Bottom - CellRect.Top) - PictureHeight) / 2)
            Case FlexPictureAlignmentLeftBottom
                PictureOffsetY = ((CellRect.Bottom - CellRect.Top) - PictureHeight)
            Case FlexPictureAlignmentCenterTop
                PictureOffsetX = (((CellRect.Right - CellRect.Left) - PictureWidth) / 2)
            Case FlexPictureAlignmentCenterCenter
                PictureOffsetX = (((CellRect.Right - CellRect.Left) - PictureWidth) / 2)
                PictureOffsetY = (((CellRect.Bottom - CellRect.Top) - PictureHeight) / 2)
            Case FlexPictureAlignmentCenterBottom
                PictureOffsetX = (((CellRect.Right - CellRect.Left) - PictureWidth) / 2)
                PictureOffsetY = ((CellRect.Bottom - CellRect.Top) - PictureHeight)
            Case FlexPictureAlignmentRightTop
                PictureOffsetX = ((CellRect.Right - CellRect.Left) - PictureWidth)
            Case FlexPictureAlignmentRightCenter
                PictureOffsetX = ((CellRect.Right - CellRect.Left) - PictureWidth)
                PictureOffsetY = (((CellRect.Bottom - CellRect.Top) - PictureHeight) / 2)
            Case FlexPictureAlignmentRightBottom
                PictureOffsetX = ((CellRect.Right - CellRect.Left) - PictureWidth)
                PictureOffsetY = ((CellRect.Bottom - CellRect.Top) - PictureHeight)
        End Select
        If PictureOffsetX > 0 Then PictureLeft = PictureLeft + PictureOffsetX
        If PictureOffsetY > 0 Then PictureTop = PictureTop + PictureOffsetY
        If .PictureAlignment <> FlexPictureAlignmentTile Then
            With .Picture
            If .Type = vbPicTypeIcon Then
                DrawIconEx hDC, PictureLeft, PictureTop, .Handle, PictureWidth, PictureHeight, 0, 0, DI_NORMAL
            Else
                .Render hDC Or 0&, PictureLeft Or 0&, PictureTop Or 0&, PictureWidth Or 0&, PictureHeight Or 0&, 0&, .Height, .Width, -.Height, ByVal 0&
            End If
            End With
        Else
            With .Picture
            If .Type = vbPicTypeIcon Then
                Do
                    Do
                        DrawIconEx hDC, PictureLeft, PictureTop, .Handle, PictureWidth, PictureHeight, 0, 0, DI_NORMAL
                        PictureTop = PictureTop + PictureHeight
                    Loop While PictureTop < CellRect.Bottom
                    PictureLeft = PictureLeft + PictureWidth
                    PictureTop = CellRect.Top
                Loop While PictureLeft < CellRect.Right
            Else
                Do
                    Do
                        .Render hDC Or 0&, PictureLeft Or 0&, PictureTop Or 0&, PictureWidth Or 0&, PictureHeight Or 0&, 0&, .Height, .Width, -.Height, ByVal 0&
                        PictureTop = PictureTop + PictureHeight
                    Loop While PictureTop < CellRect.Bottom
                    PictureLeft = PictureLeft + PictureWidth
                    PictureTop = CellRect.Top
                Loop While PictureLeft < CellRect.Right
            End If
            End With
        End If
    End If
End If
Dim OldBkMode As Long, OldTextColor As Long
OldBkMode = SetBkMode(hDC, 1)
If Not (ItemState And ODS_SELECTED) = ODS_SELECTED Or (ItemState And ODS_FOCUS) = ODS_FOCUS Then
    If Not .Text = vbNullString Then
        If .ForeColor = -1 Then
            If IsFixedCell = False Then
                OldTextColor = SetTextColor(hDC, WinColor(PropForeColor))
            Else
                OldTextColor = SetTextColor(hDC, WinColor(PropForeColorFixed))
            End If
        Else
            OldTextColor = SetTextColor(hDC, WinColor(.ForeColor))
        End If
    Else
        If IsFixedCell = False Then
            OldTextColor = SetTextColor(hDC, WinColor(vbWindowText))
        Else
            OldTextColor = SetTextColor(hDC, WinColor(vbButtonText))
        End If
    End If
Else
    OldTextColor = SetTextColor(hDC, WinColor(ForeColorSel))
End If
Dim GridLines As FlexGridLineConstants, hPenOld As Long, P As POINTAPI
GridLines = IIf(IsFixedCell = False, PropGridLines, PropGridLinesFixed)
Select Case GridLines
    Case FlexGridLineFlat, FlexGridLineDashes, FlexGridLineDots
        If IsFixedCell = False Then
            If VBFlexGridGridLinePen <> 0 Then hPenOld = SelectObject(hDC, VBFlexGridGridLinePen)
        Else
            If VBFlexGridGridLineFixedPen <> 0 Then hPenOld = SelectObject(hDC, VBFlexGridGridLineFixedPen)
        End If
        MoveToEx hDC, CellRect.Left, CellRect.Bottom - 1, P
        LineTo hDC, CellRect.Right - 1, CellRect.Bottom - 1
        LineTo hDC, CellRect.Right - 1, CellRect.Top - 1
        MoveToEx hDC, P.X, P.Y, ByVal 0&
    Case FlexGridLineInset, FlexGridLineRaised
        If GridLines = FlexGridLineInset Then
            If VBFlexGridGridLineWhitePen <> 0 Then hPenOld = SelectObject(hDC, VBFlexGridGridLineWhitePen)
        ElseIf GridLines = FlexGridLineRaised Then
            If VBFlexGridGridLineBlackPen <> 0 Then hPenOld = SelectObject(hDC, VBFlexGridGridLineBlackPen)
        End If
        MoveToEx hDC, CellRect.Left, CellRect.Bottom - 1, P
        LineTo hDC, CellRect.Left, CellRect.Top
        LineTo hDC, CellRect.Right - 1, CellRect.Top
        If GridLines = FlexGridLineInset Then
            If VBFlexGridGridLineBlackPen <> 0 Then SelectObject hDC, VBFlexGridGridLineBlackPen
        ElseIf GridLines = FlexGridLineRaised Then
            If VBFlexGridGridLineWhitePen <> 0 Then SelectObject hDC, VBFlexGridGridLineWhitePen
        End If
        LineTo hDC, CellRect.Right - 1, CellRect.Bottom - 1
        LineTo hDC, CellRect.Left, CellRect.Bottom - 1
        MoveToEx hDC, P.X, P.Y, ByVal 0&
End Select
If hPenOld <> 0 Then
    SelectObject hDC, hPenOld
    hPenOld = 0
End If
If (ItemState And ODS_FOCUS) = ODS_FOCUS And Not (ItemState And ODS_NOFOCUSRECT) = ODS_NOFOCUSRECT Then
    Dim FocusRect As RECT
    With FocusRect
    .Top = CellRect.Top
    .Left = CellRect.Left
    If (CellRect.Bottom - 1) > .Top Then .Bottom = CellRect.Bottom - 1 Else .Bottom = CellRect.Bottom + 1
    If (CellRect.Right - 1) > .Left Then .Right = CellRect.Right - 1 Else .Right = CellRect.Right + 1
    DrawFocusRect hDC, FocusRect
    If PropFocusRect = FlexFocusRectHeavy Then
        If (.Bottom - 1) > (.Top + 1) And (.Right - 1) > (.Left + 1) Then
            .Top = .Top + 1
            .Bottom = .Bottom - 1
            .Left = .Left + 1
            .Right = .Right - 1
            DrawFocusRect hDC, FocusRect
        End If
    End If
    End With
End If
If Not .Text = vbNullString Then
    Dim TextRect As RECT, TextStyle As FlexTextStyleConstants, Alignment As FlexAlignmentConstants, Format As Long
    With TextRect
    .Top = CellRect.Top + (1 * PixelsPerDIP_Y())
    .Left = CellRect.Left + (3 * PixelsPerDIP_X())
    .Bottom = CellRect.Bottom - (1 * PixelsPerDIP_Y())
    .Right = CellRect.Right - (3 * PixelsPerDIP_X())
    End With
    If .TextStyle = -1 Then
        If IsFixedCell = False Then
            TextStyle = PropTextStyle
        Else
            TextStyle = PropTextStyleFixed
        End If
    Else
        TextStyle = .TextStyle
    End If
    If .Alignment = -1 Then
        If IsFixedCell = False Then
            Alignment = VBFlexGridColsInfo(iCol).Alignment
        Else
            If VBFlexGridColsInfo(iCol).FixedAlignment = -1 Then
                Alignment = VBFlexGridColsInfo(iCol).Alignment
            Else
                Alignment = VBFlexGridColsInfo(iCol).FixedAlignment
            End If
        End If
    Else
        Alignment = .Alignment
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
            If Not IsNumeric(.Text) Then
                Format = Format Or DT_LEFT
            Else
                Format = Format Or DT_RIGHT
            End If
    End Select
    If PropWordWrap = True Then Format = Format Or DT_WORDBREAK
    If IsFixedCell = False Then
        Select Case PropEllipsisFormat
            Case FlexEllipsisFormatEnd
                Format = Format Or DT_END_ELLIPSIS
            Case FlexEllipsisFormatPath
                Format = Format Or DT_PATH_ELLIPSIS
            Case FlexEllipsisFormatWord
                Format = Format Or DT_WORD_ELLIPSIS
        End Select
    Else
        Select Case PropEllipsisFormatFixed
            Case FlexEllipsisFormatEnd
                Format = Format Or DT_END_ELLIPSIS
            Case FlexEllipsisFormatPath
                Format = Format Or DT_PATH_ELLIPSIS
            Case FlexEllipsisFormatWord
                Format = Format Or DT_WORD_ELLIPSIS
        End Select
    End If
    Dim CalcRect As RECT, Height As Long, Result As Long
    Select Case Alignment
        Case FlexAlignmentLeftCenter, FlexAlignmentCenterCenter, FlexAlignmentRightCenter, FlexAlignmentGeneral
            LSet CalcRect = TextRect
            Height = DrawText(hDC, StrPtr(.Text), -1, CalcRect, Format Or DT_CALCRECT)
            Result = (((TextRect.Bottom - TextRect.Top) - Height) / 2)
        Case FlexAlignmentLeftBottom, FlexAlignmentCenterBottom, FlexAlignmentRightBottom
            LSet CalcRect = TextRect
            Height = DrawText(hDC, StrPtr(.Text), -1, CalcRect, Format Or DT_CALCRECT)
            Result = ((TextRect.Bottom - TextRect.Top) - Height)
    End Select
    If Result > 0 Then TextRect.Top = TextRect.Top + Result
    Dim Offset As Long
    Select Case TextStyle
        Case FlexTextStyleRaised
            Result = SetTextColor(hDC, &H808080)
            Offset = 1
        Case FlexTextStyleRaisedLight
            Result = SetTextColor(hDC, vbWhite)
            Offset = 1
        Case FlexTextStyleInset
            Result = SetTextColor(hDC, &H808080)
            Offset = -1
        Case FlexTextStyleInsetLight
            Result = SetTextColor(hDC, vbWhite)
            Offset = -1
    End Select
    If Offset <> 0 Then
        With TextRect
        .Top = .Top + Offset
        .Left = .Left + Offset
        .Bottom = .Bottom + Offset
        .Right = .Right + Offset
        End With
        DrawText hDC, StrPtr(.Text), -1, TextRect, Format
        SetTextColor hDC, Result
        With TextRect
        .Top = .Top - Offset
        .Left = .Left - Offset
        .Bottom = .Bottom - Offset
        .Right = .Right - Offset
        End With
    End If
    DrawText hDC, StrPtr(.Text), -1, TextRect, Format
End If
SetBkMode hDC, OldBkMode
SetTextColor hDC, OldTextColor
If hFontOld <> 0 Then SelectObject hDC, hFontOld
If hFontTemp <> 0 Then DeleteObject hFontTemp
End With
End Sub

Private Sub GetSelRangeStruct(ByRef SelRange As TSELRANGE)
With SelRange
If VBFlexGridRow > VBFlexGridRowSel Then .TopRow = VBFlexGridRowSel Else .TopRow = VBFlexGridRow
If VBFlexGridRowSel > VBFlexGridRow Then .BottomRow = VBFlexGridRowSel Else .BottomRow = VBFlexGridRow
If VBFlexGridCol > VBFlexGridColSel Then .LeftCol = VBFlexGridColSel Else .LeftCol = VBFlexGridCol
If VBFlexGridColSel > VBFlexGridCol Then .RightCol = VBFlexGridColSel Else .RightCol = VBFlexGridCol
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
    If i >= VBFlexGridTopRow Or i < PropFixedRows Then
        .Top = .Bottom
        .Bottom = .Bottom + GetRowHeight(i)
    End If
Next i
For i = 0 To iCol
    If i >= VBFlexGridLeftCol Or i < PropFixedCols Then
        .Left = .Right
        .Right = .Right + GetColWidth(i)
    End If
Next i
End With
End Sub

Private Function GetRowHeight(ByVal iRow As Long) As Long
If PropRows < 1 Or PropCols < 1 Then Exit Function
If VBFlexGridCells.Rows(iRow).RowInfo.Hidden = False Then
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
If VBFlexGridColsInfo(iCol).Hidden = False Then
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
If PropRows < 1 Or PropCols < 1 Then Exit Function
If VBFlexGridHandle <> 0 Then
    Dim hDC As Long
    hDC = GetDC(VBFlexGridHandle)
    If hDC <> 0 Then
        Dim hFontTemp As Long
        With VBFlexGridCells.Rows(iRow).Cols(iCol)
        If .FontName = vbNullString Then
            If iRow >= PropFixedRows Or iCol >= PropFixedCols Then
                SelectObject hDC, VBFlexGridFontHandle
            Else
                If VBFlexGridFontFixedHandle = 0 Then
                    SelectObject hDC, VBFlexGridFontHandle
                Else
                    SelectObject hDC, VBFlexGridFontFixedHandle
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
            SelectObject hDC, hFontTemp
            Set TempFont = Nothing
        End If
        End With
        Dim Pos1 As Long, Pos2 As Long, Temp As String, Size As SIZEAPI
        If InStr(Text, vbCrLf) Then Text = Replace$(Text, vbCrLf, vbCr)
        If InStr(Text, vbLf) Then Text = Replace$(Text, vbLf, vbCr)
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
        ReleaseDC VBFlexGridHandle, hDC
        If hFontTemp <> 0 Then DeleteObject hFontTemp
    End If
End If
End Function

Private Sub GetHitTestInfo(ByRef HTI As THITTESTINFO)
HTI.HitRow = -1
HTI.HitCol = -1
HTI.HitResult = FlexHitResultNoWhere
HTI.MouseRow = 0
HTI.MouseCol = 0
If VBFlexGridHandle = 0 Or (PropRows < 1 Or PropCols < 1) Then Exit Sub
Dim iRow As Long, iCol As Long
Dim ClientRect As RECT, CellRect As RECT, TempRect As RECT
GetClientRect VBFlexGridHandle, ClientRect
With CellRect
For iRow = 0 To (PropRows - 1)
    If iRow >= VBFlexGridTopRow Then
        .Bottom = .Top + GetRowHeight(iRow)
        For iCol = 0 To (PropCols - 1)
            If iCol >= VBFlexGridLeftCol Then
                .Left = .Right
                .Right = .Right + GetColWidth(iCol)
                If PtInRect(CellRect, HTI.PT.X, HTI.PT.Y) <> 0 Then HTI.HitResult = FlexHitResultCell
            ElseIf iCol < PropFixedCols Then
                .Left = .Right
                .Right = .Right + GetColWidth(iCol)
                If PtInRect(CellRect, HTI.PT.X, HTI.PT.Y) <> 0 Then
                    If PropAllowUserResizing = FlexAllowUserResizingRows Or PropAllowUserResizing = FlexAllowUserResizingBoth Then
                        SetRect TempRect, .Left, .Top, .Right, .Bottom
                        If iRow > 0 Then TempRect.Top = TempRect.Top + (2 * PixelsPerDIP_Y())
                        TempRect.Bottom = TempRect.Bottom - (2 * PixelsPerDIP_Y())
                        If PtInRect(TempRect, HTI.PT.X, HTI.PT.Y) <> 0 Then
                            HTI.HitResult = FlexHitResultCell
                        Else
                            TempRect.Bottom = .Bottom
                            If PtInRect(TempRect, HTI.PT.X, HTI.PT.Y) = 0 Then
                                HTI.HitResult = FlexHitResultDividerRowTop
                            Else
                                HTI.HitResult = FlexHitResultDividerRowBottom
                            End If
                        End If
                    Else
                        HTI.HitResult = FlexHitResultCell
                    End If
                End If
            End If
            If HTI.PT.Y >= CellRect.Top Then HTI.MouseRow = iRow
            If HTI.PT.X >= CellRect.Left Then HTI.MouseCol = iCol
            If HTI.HitResult <> FlexHitResultNoWhere Then Exit For
        Next iCol
        .Left = 0
        .Right = 0
        .Top = .Top + GetRowHeight(iRow)
    ElseIf iRow < PropFixedRows Then
        .Bottom = .Top + GetRowHeight(iRow)
        For iCol = 0 To (PropCols - 1)
            If iCol >= VBFlexGridLeftCol Or iCol < PropFixedCols Then
                .Left = .Right
                .Right = .Right + GetColWidth(iCol)
                If PtInRect(CellRect, HTI.PT.X, HTI.PT.Y) <> 0 Then
                    If PropAllowUserResizing <> FlexAllowUserResizingNone Then
                        SetRect TempRect, .Left, .Top, .Right, .Bottom
                        If iCol > 0 Then TempRect.Left = TempRect.Left + (2 * PixelsPerDIP_X())
                        TempRect.Right = TempRect.Right - (2 * PixelsPerDIP_X())
                        If PtInRect(TempRect, HTI.PT.X, HTI.PT.Y) <> 0 Then
                            If iCol < PropFixedCols Then
                                If PropAllowUserResizing <> FlexAllowUserResizingColumns Then
                                    SetRect TempRect, .Left, .Top, .Right, .Bottom
                                    If iRow > 0 Then TempRect.Top = TempRect.Top + (2 * PixelsPerDIP_Y())
                                    TempRect.Bottom = TempRect.Bottom - (2 * PixelsPerDIP_Y())
                                    If PtInRect(TempRect, HTI.PT.X, HTI.PT.Y) <> 0 Then
                                        HTI.HitResult = FlexHitResultCell
                                    Else
                                        TempRect.Bottom = .Bottom
                                        If PtInRect(TempRect, HTI.PT.X, HTI.PT.Y) = 0 Then
                                            HTI.HitResult = FlexHitResultDividerRowTop
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
                            If PtInRect(TempRect, HTI.PT.X, HTI.PT.Y) = 0 Then
                                HTI.HitResult = FlexHitResultDividerColumnLeft
                            Else
                                HTI.HitResult = FlexHitResultDividerColumnRight
                            End If
                        Else
                            HTI.HitResult = FlexHitResultCell
                        End If
                    Else
                        HTI.HitResult = FlexHitResultCell
                    End If
                End If
            End If
            If HTI.PT.Y >= CellRect.Top Then HTI.MouseRow = iRow
            If HTI.PT.X >= CellRect.Left Then HTI.MouseCol = iCol
            If HTI.HitResult <> FlexHitResultNoWhere Then Exit For
        Next iCol
        .Left = 0
        .Right = 0
        .Top = .Top + GetRowHeight(iRow)
    End If
    If HTI.HitResult <> FlexHitResultNoWhere Then Exit For
Next iRow
End With
If HTI.HitResult <> FlexHitResultNoWhere Then
    HTI.HitRow = iRow
    HTI.HitCol = iCol
End If
End Sub

Private Sub GetLabelInfo(ByVal iRow As Long, ByVal iCol As Long, ByRef LBLI As TLABELINFO)
LBLI.Flags = 0
If VBFlexGridHandle = 0 Or (PropRows < 1 Or PropCols < 1) Then Exit Sub
Dim hDC As Long
hDC = GetDC(VBFlexGridHandle)
If hDC <> 0 Then
    Dim CellRect As RECT
    Call GetCellRect(iRow, iCol, False, CellRect)
    If (CellRect.Bottom - CellRect.Top) > 0 And (CellRect.Right - CellRect.Left) > 0 Then
        Dim ClientRect As RECT, IsFixedCell As Boolean
        GetClientRect VBFlexGridHandle, ClientRect
        IsFixedCell = CBool(iRow < PropFixedRows Or iCol < PropFixedCols)
        With VBFlexGridCells.Rows(iRow).Cols(iCol)
        Dim hFontTemp As Long, hFontOld As Long
        If .FontName = vbNullString Then
            If IsFixedCell = False Then
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
        Dim TextRect As RECT, TextStyle As FlexTextStyleConstants, Alignment As FlexAlignmentConstants, Format As Long
        With TextRect
        .Top = CellRect.Top + (1 * PixelsPerDIP_Y())
        .Left = CellRect.Left + (3 * PixelsPerDIP_X())
        .Bottom = CellRect.Bottom - (1 * PixelsPerDIP_Y())
        .Right = CellRect.Right - (3 * PixelsPerDIP_X())
        End With
        If .TextStyle = -1 Then
            If IsFixedCell = False Then
                TextStyle = PropTextStyle
            Else
                TextStyle = PropTextStyleFixed
            End If
        Else
            TextStyle = .TextStyle
        End If
        If .Alignment = -1 Then
            If IsFixedCell = False Then
                Alignment = VBFlexGridColsInfo(iCol).Alignment
            Else
                Alignment = VBFlexGridColsInfo(iCol).FixedAlignment
            End If
        Else
            Alignment = .Alignment
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
                If Not IsNumeric(.Text) Then
                    Format = Format Or DT_LEFT
                Else
                    Format = Format Or DT_RIGHT
                End If
        End Select
        If PropWordWrap = True Then Format = Format Or DT_WORDBREAK
        ' Ellipsis format will be ignored.
        Dim CalcRect As RECT, Height As Long, Result As Long
        LSet CalcRect = TextRect
        Select Case Alignment
            Case FlexAlignmentLeftCenter, FlexAlignmentCenterCenter, FlexAlignmentRightCenter, FlexAlignmentGeneral
                Height = DrawText(hDC, StrPtr(.Text), -1, CalcRect, Format Or DT_CALCRECT)
                Result = (((TextRect.Bottom - TextRect.Top) - Height) / 2)
            Case FlexAlignmentLeftBottom, FlexAlignmentCenterBottom, FlexAlignmentRightBottom
                Height = DrawText(hDC, StrPtr(.Text), -1, CalcRect, Format Or DT_CALCRECT)
                Result = ((TextRect.Bottom - TextRect.Top) - Height)
        End Select
        If Result > 0 Then
            CalcRect.Top = CalcRect.Top + Result
            CalcRect.Bottom = CalcRect.Bottom + Result
        End If
        With LBLI
        .Flags = LBLI_VALID
        If TextRect.Right <= ClientRect.Right And TextRect.Bottom <= ClientRect.Bottom Then
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
        End With
    End If
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
GetClientRect VBFlexGridHandle, ClientRect
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
            If .Right > ClientRect.Right And iCol > PropFixedCols Then
                SCI(0).nMax = (PropCols - PropFixedCols) - 1
                ' Scroll box is proportional to the scrolling region.
                ' But only appropriate when all columns are equally in width.
                ' SCI(0).nPage = iCol - (PropFixedCols - 1) - 1
                Exit For
            End If
        Next iCol
        If SCI(0).nMax > 0 And SCI(0).nPage = 0 Then
            .Right = 0
            For iCol = 0 To (PropFixedCols - 1)
                .Right = .Right + GetColWidth(iCol)
            Next iCol
            For iCol = (PropCols - 1) To PropFixedCols Step -1
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
            If .Bottom > ClientRect.Bottom And iRow > PropFixedRows Then
                SCI(1).nMax = (PropRows - PropFixedRows) - 1
                ' Scroll box is proportional to the scrolling region.
                ' But only appropriate when all rows are equally in height.
                ' SCI(1).nPage = iRow - (PropFixedRows - 1) - 1
                Exit For
            End If
        Next iRow
        If SCI(1).nMax > 0 And SCI(1).nPage = 0 Then
            .Bottom = 0
            For iRow = 0 To (PropFixedRows - 1)
                .Bottom = .Bottom + GetRowHeight(iRow)
            Next iRow
            For iRow = (PropRows - 1) To PropFixedRows Step -1
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
Dim ClientRect As RECT, GridRect As RECT, iRow As Long, Count As Long
GetClientRect VBFlexGridHandle, ClientRect
With GridRect
For iRow = 0 To (PropFixedRows - 1)
    .Bottom = .Bottom + GetRowHeight(iRow)
Next iRow
For iRow = TopRow To (PropRows - 1)
    .Bottom = .Bottom + GetRowHeight(iRow)
    If iRow > TopRow And .Bottom > ClientRect.Bottom Then Exit For
    Count = Count + 1
Next iRow
GetRowsPerPage = Count
End With
End Function

Private Function GetRowsPerPageRev(ByVal BottomRow As Long) As Long
If VBFlexGridHandle = 0 Or (PropRows < 1 Or PropCols < 1) Then Exit Function
Dim ClientRect As RECT, GridRect As RECT, iRow As Long, Count As Long
GetClientRect VBFlexGridHandle, ClientRect
With GridRect
For iRow = 0 To (PropFixedRows - 1)
    .Bottom = .Bottom + GetRowHeight(iRow)
Next iRow
For iRow = BottomRow To PropFixedRows Step -1
    .Bottom = .Bottom + GetRowHeight(iRow)
    If iRow < BottomRow And .Bottom > ClientRect.Bottom Then Exit For
    Count = Count + 1
Next iRow
GetRowsPerPageRev = Count
End With
End Function

Private Function GetColsPerPage(ByVal LeftCol As Long) As Long
If VBFlexGridHandle = 0 Or (PropRows < 1 Or PropCols < 1) Then Exit Function
Dim ClientRect As RECT, GridRect As RECT, iCol As Long, Count As Long
GetClientRect VBFlexGridHandle, ClientRect
With GridRect
For iCol = 0 To (PropFixedCols - 1)
    .Right = .Right + GetColWidth(iCol)
Next iCol
For iCol = LeftCol To (PropCols - 1)
    .Right = .Right + GetColWidth(iCol)
    If iCol > LeftCol And .Right > ClientRect.Right Then Exit For
    Count = Count + 1
Next iCol
GetColsPerPage = Count
End With
End Function

Private Function GetColsPerPageRev(ByVal RightCol As Long) As Long
If VBFlexGridHandle = 0 Or (PropRows < 1 Or PropCols < 1) Then Exit Function
Dim ClientRect As RECT, GridRect As RECT, iCol As Long, Count As Long
GetClientRect VBFlexGridHandle, ClientRect
With GridRect
For iCol = 0 To (PropFixedCols - 1)
    .Right = .Right + GetColWidth(iCol)
Next iCol
For iCol = RightCol To PropFixedCols Step -1
    .Right = .Right + GetColWidth(iCol)
    If iCol < RightCol And .Right > ClientRect.Right Then Exit For
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
    CheckScrollPos = CBool((VBFlexGridLeftCol - PropFixedCols) <> SCI.nPos)
ElseIf wBar = SB_VERT Then
    CheckScrollPos = CBool((VBFlexGridTopRow - PropFixedRows) <> SCI.nPos)
End If
If CheckScrollPos = False Then Exit Function
PrevPos = SCI.nPos
If wBar = SB_HORZ Then
    SCI.nPos = VBFlexGridLeftCol - PropFixedCols
ElseIf wBar = SB_VERT Then
    SCI.nPos = VBFlexGridTopRow - PropFixedRows
End If
SetScrollInfo VBFlexGridHandle, wBar, SCI, 1
GetScrollInfo VBFlexGridHandle, wBar, SCI
If PrevPos <> SCI.nPos Then
    Call RedrawGrid
    If PropShowInfoTips = True Or PropShowLabelTips = True Then
        Dim Pos As Long
        Pos = GetMessagePos()
        Call CheckToolTipRowCol(Get_X_lParam(Pos), Get_Y_lParam(Pos))
    End If
    RaiseEvent Scroll
End If
End Function

Private Sub CheckTopRow(ByRef TopRow As Long)
If VBFlexGridHandle = 0 Or (PropRows < 1 Or PropCols < 1) Then Exit Sub
Dim ClientRect As RECT, GridRect As RECT, iRow As Long, Changed As Boolean
GetClientRect VBFlexGridHandle, ClientRect
With GridRect
Do
    Changed = False
    For iRow = 0 To (PropFixedRows - 1)
        .Bottom = .Bottom + GetRowHeight(iRow)
    Next iRow
    For iRow = TopRow To (PropRows - 1)
        .Bottom = .Bottom + GetRowHeight(iRow)
        If .Bottom > ClientRect.Bottom Then Exit For
    Next iRow
    If .Bottom <= ClientRect.Bottom And TopRow > ((PropFixedRows - 1) + 1) Then
        .Bottom = .Bottom + GetRowHeight(TopRow - 1)
        If .Bottom <= ClientRect.Bottom Then
            .Bottom = 0
            TopRow = TopRow - 1
            Changed = True
        End If
    End If
Loop Until Changed = False
End With
End Sub

Private Sub CheckLeftCol(ByRef LeftCol As Long)
If VBFlexGridHandle = 0 Or (PropRows < 1 Or PropCols < 1) Then Exit Sub
Dim ClientRect As RECT, GridRect As RECT, iCol As Long, Changed As Boolean
GetClientRect VBFlexGridHandle, ClientRect
With GridRect
Do
    Changed = False
    For iCol = 0 To (PropFixedCols - 1)
        .Right = .Right + GetColWidth(iCol)
    Next iCol
    For iCol = LeftCol To (PropCols - 1)
        .Right = .Right + GetColWidth(iCol)
        If .Right > ClientRect.Right Then Exit For
    Next iCol
    If .Right <= ClientRect.Right And LeftCol > ((PropFixedCols - 1) + 1) Then
        .Right = .Right + GetColWidth(LeftCol - 1)
        If .Right <= ClientRect.Right Then
            .Right = 0
            LeftCol = LeftCol - 1
            Changed = True
        End If
    End If
Loop Until Changed = False
End With
End Sub

Private Sub SetRowColParams(ByRef RCP As TROWCOLPARAMS)
Dim RowColChanged As Boolean, SelChanged As Boolean, ScrollChanged As Boolean
Dim NoRedraw As Boolean, Cancel As Boolean
With RCP
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
    If .TopRow < PropFixedRows Then .TopRow = PropFixedRows
    If (.Flags And RCPF_CHECKTOPROW) = RCPF_CHECKTOPROW Then Call CheckTopRow(.TopRow)
    If VBFlexGridTopRow <> .TopRow Then ScrollChanged = True
End If
If (.Mask And RCPM_LEFTCOL) = RCPM_LEFTCOL Then
    If .LeftCol < PropFixedCols Then .LeftCol = PropFixedCols
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
If NoRedraw = False Then Call RedrawGrid
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
    If iRow < (PropRows - 1) Then i = i + 1 Else Cancel = True
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
    If iCol < (PropCols - 1) Then i = i + 1 Else Cancel = True
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

Private Sub ProcessKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
If PropRows < 1 Or PropCols < 1 Then Exit Sub
Select Case KeyCode
    Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyPageUp, vbKeyPageDown, vbKeyHome, vbKeyEnd
    Case vbKeyTab
        If PropTabBehavior = FlexTabControls Then Exit Sub
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
    If (Shift And vbShiftMask) = vbShiftMask Then Exit Sub
End If
Dim RCP As TROWCOLPARAMS, RowsPerPage As Long, ColsPerPage As Long
With RCP
.Mask = RCPM_ROW Or RCPM_COL Or RCPM_ROWSEL Or RCPM_COLSEL Or RCPM_TOPROW Or RCPM_LEFTCOL
.Row = VBFlexGridRow
.Col = VBFlexGridCol
Select Case KeyCode
    Case vbKeyUp, vbKeyDown, vbKeyPageUp, vbKeyPageDown
        If .Row < PropFixedRows Then .Row = PropFixedRows
        If .Row > (PropRows - 1) Then .Row = (PropRows - 1)
    Case vbKeyLeft, vbKeyRight, vbKeyHome, vbKeyEnd
        If .Col < PropFixedCols Then .Col = PropFixedCols
        If .Col > (PropCols - 1) Then .Col = (PropCols - 1)
    Case vbKeyTab
        If .Row < PropFixedRows Then .Row = PropFixedRows
        If .Row > (PropRows - 1) Then .Row = (PropRows - 1)
        If .Col < PropFixedCols Then .Col = PropFixedCols
        If .Col > (PropCols - 1) Then .Col = (PropCols - 1)
End Select
.RowSel = VBFlexGridRowSel
.ColSel = VBFlexGridColSel
.TopRow = VBFlexGridTopRow
.LeftCol = VBFlexGridLeftCol
Select Case PropSelectionMode
    Case FlexSelectionModeFree
        Select Case KeyCode
            Case vbKeyUp
                If (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) = 0 Then
                    Call MovePreviousRow(.Row)
                    .RowSel = .Row
                    .ColSel = .Col
                    If .TopRow > .Row Then
                        .TopRow = .Row
                    ElseIf .Row > (.TopRow + GetRowsPerPage(.TopRow) - 1) Then
                        .TopRow = .Row - GetRowsPerPageRev(.Row) + 1
                    End If
                    If .LeftCol > .Col Then
                        .LeftCol = .Col
                    ElseIf .Col > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .Col - GetColsPerPageRev(.Col) + 1
                    End If
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    Call MovePreviousRow(.RowSel)
                    If .TopRow > .RowSel Then .TopRow = .RowSel
                ElseIf (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                    Call MoveFirstRow(.Row)
                    .RowSel = .Row
                    .ColSel = .Col
                    .TopRow = .Row
                    If .LeftCol > .Col Then
                        .LeftCol = .Col
                    ElseIf .Col > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .Col - GetColsPerPageRev(.Col) + 1
                    End If
                Else
                    Call MoveFirstRow(.RowSel)
                    .TopRow = .RowSel
                End If
            Case vbKeyDown
                If (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) = 0 Then
                    Call MoveNextRow(.Row)
                    .RowSel = .Row
                    .ColSel = .Col
                    If .TopRow > .Row Then
                        .TopRow = .Row
                    ElseIf .Row > (.TopRow + GetRowsPerPage(.TopRow) - 1) Then
                        .TopRow = .Row - GetRowsPerPageRev(.Row) + 1
                    End If
                    If .LeftCol > .Col Then
                        .LeftCol = .Col
                    ElseIf .Col > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .Col - GetColsPerPageRev(.Col) + 1
                    End If
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    Call MoveNextRow(.RowSel)
                    If .TopRow > .RowSel Then
                        .TopRow = .RowSel
                    ElseIf .RowSel > (.TopRow + GetRowsPerPage(.TopRow) - 1) Then
                        .TopRow = .RowSel - GetRowsPerPageRev(.RowSel) + 1
                    End If
                ElseIf (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                    Call MoveLastRow(.Row)
                    .RowSel = .Row
                    .ColSel = .Col
                    .TopRow = (PropRows - 1) - GetRowsPerPageRev(PropRows - 1) + 1
                    If .LeftCol > .Col Then
                        .LeftCol = .Col
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
                        If .Col > PropFixedCols Then
                            Call MovePreviousCol(.Col)
                        Else
                            If .Row > PropFixedRows Then
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
                        .TopRow = .Row
                    ElseIf .Row > (.TopRow + GetRowsPerPage(.TopRow) - 1) Then
                        .TopRow = .Row - GetRowsPerPageRev(.Row) + 1
                    End If
                    If .LeftCol > .Col Then
                        .LeftCol = .Col
                    ElseIf .Col > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .Col - GetColsPerPageRev(.Col) + 1
                    End If
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    Call MovePreviousCol(.ColSel)
                    If .LeftCol > .ColSel Then .LeftCol = .ColSel
                ElseIf (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                    Call MoveFirstCol(.Col)
                    .RowSel = .Row
                    .ColSel = .Col
                    If .TopRow > .Row Then
                        .TopRow = .Row
                    ElseIf .Row > (.TopRow + GetRowsPerPage(.TopRow) - 1) Then
                        .TopRow = .Row - GetRowsPerPageRev(.Row) + 1
                    End If
                    .LeftCol = .Col
                Else
                    Call MoveFirstCol(.ColSel)
                    .LeftCol = .ColSel
                End If
            Case vbKeyRight
                If (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) = 0 Then
                    If PropWrapCellBehavior <> FlexWrapNone Then
                        If .Col < (PropCols - 1) Then
                            Call MoveNextCol(.Col)
                        Else
                            If .Row < (PropRows - 1) Then
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
                        .TopRow = .Row
                    ElseIf .Row > (.TopRow + GetRowsPerPage(.TopRow) - 1) Then
                        .TopRow = .Row - GetRowsPerPageRev(.Row) + 1
                    End If
                    If .LeftCol > .Col Then
                        .LeftCol = .Col
                    ElseIf .Col > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .Col - GetColsPerPageRev(.Col) + 1
                    End If
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    Call MoveNextCol(.ColSel)
                    If .LeftCol > .ColSel Then
                        .LeftCol = .ColSel
                    ElseIf .ColSel > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .ColSel - GetColsPerPageRev(.ColSel) + 1
                    End If
                ElseIf (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                    Call MoveLastCol(.Col)
                    .RowSel = .Row
                    .ColSel = .Col
                    If .TopRow > .Row Then
                        .TopRow = .Row
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
                    If .Row > PropFixedRows Then
                        RowsPerPage = GetRowsPerPageRev(.Row)
                        If (.Row - RowsPerPage) > PropFixedRows Then
                            .Row = .Row - RowsPerPage
                        Else
                            .Row = PropFixedRows
                        End If
                    End If
                    .RowSel = .Row
                    .ColSel = .Col
                    If .TopRow > .Row Then
                        .TopRow = .Row
                    Else
                        .TopRow = .Row - GetRowsPerPageRev(.Row) + 1
                    End If
                    If .LeftCol > .Col Then
                        .LeftCol = .Col
                    ElseIf .ColSel > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .ColSel - GetColsPerPageRev(.ColSel) + 1
                    End If
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    If .RowSel > PropFixedRows Then
                        RowsPerPage = GetRowsPerPageRev(.RowSel)
                        If (.RowSel - RowsPerPage) > PropFixedRows Then
                            .RowSel = .RowSel - RowsPerPage
                        Else
                            .RowSel = PropFixedRows
                        End If
                    End If
                    If .TopRow > .RowSel Then
                        .TopRow = .RowSel
                    Else
                        .TopRow = .RowSel - GetRowsPerPageRev(.RowSel) + 1
                    End If
                ElseIf (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                    Call MoveFirstRow(.Row)
                    .RowSel = .Row
                    .ColSel = .Col
                    .TopRow = .Row
                    If .LeftCol > .Col Then
                        .LeftCol = .Col
                    ElseIf .Col > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .Col - GetColsPerPageRev(.Col) + 1
                    End If
                Else
                    Call MoveFirstRow(.RowSel)
                    .TopRow = .RowSel
                End If
            Case vbKeyPageDown
                If (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) = 0 Then
                    If .Row < (PropRows - 1) Then
                        RowsPerPage = GetRowsPerPage(.Row)
                        If (.Row + RowsPerPage) < (PropRows - 1) Then
                            .Row = .Row + RowsPerPage
                        Else
                            .Row = (PropRows - 1)
                        End If
                    End If
                    .RowSel = .Row
                    .ColSel = .Col
                    If .TopRow > .Row Then
                        .TopRow = .Row
                    Else
                        .TopRow = .Row - GetRowsPerPageRev(.Row) + 1
                    End If
                    If .LeftCol > .Col Then
                        .LeftCol = .Col
                    ElseIf .ColSel > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .ColSel - GetColsPerPageRev(.ColSel) + 1
                    End If
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    If .RowSel < (PropRows - 1) Then
                        RowsPerPage = GetRowsPerPage(.RowSel)
                        If (.RowSel + RowsPerPage) < (PropRows - 1) Then
                            .RowSel = .RowSel + RowsPerPage
                        Else
                            .RowSel = (PropRows - 1)
                        End If
                    End If
                    If .TopRow > .RowSel Then
                        .TopRow = .RowSel
                    Else
                        .TopRow = .RowSel - GetRowsPerPageRev(.RowSel) + 1
                    End If
                ElseIf (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                    Call MoveLastRow(.Row)
                    .RowSel = .Row
                    .ColSel = .Col
                    .TopRow = (PropRows - 1) - GetRowsPerPageRev(PropRows - 1) + 1
                    If .LeftCol > .Col Then
                        .LeftCol = .Col
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
                        .TopRow = .Row
                    ElseIf .Row > (.TopRow + GetRowsPerPage(.TopRow) - 1) Then
                        .TopRow = .Row - GetRowsPerPageRev(.Row) + 1
                    End If
                    .LeftCol = .Col
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    Call MoveFirstCol(.ColSel)
                    .LeftCol = .ColSel
                ElseIf (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                    Call MoveFirstRow(.Row)
                    Call MoveFirstCol(.Col)
                    .RowSel = .Row
                    .ColSel = .Col
                    .TopRow = .Row
                    .LeftCol = .Col
                Else
                    Call MoveFirstRow(.RowSel)
                    Call MoveFirstCol(.ColSel)
                    .TopRow = .RowSel
                    .LeftCol = .ColSel
                End If
            Case vbKeyEnd
                If (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) = 0 Then
                    Call MoveLastCol(.Col)
                    .RowSel = .Row
                    .ColSel = .Col
                    If .TopRow > .Row Then
                        .TopRow = .Row
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
                        If .Col < (PropCols - 1) Then
                            Call MoveNextCol(.Col)
                        Else
                            If .Row < (PropRows - 1) Then
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
                        .TopRow = .Row
                    ElseIf .Row > (.TopRow + GetRowsPerPage(.TopRow) - 1) Then
                        .TopRow = .Row - GetRowsPerPageRev(.Row) + 1
                    End If
                    If .LeftCol > .Col Then
                        .LeftCol = .Col
                    ElseIf .Col > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .Col - GetColsPerPageRev(.Col) + 1
                    End If
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    If PropWrapCellBehavior <> FlexWrapNone Then
                        If .Col > PropFixedCols Then
                            Call MovePreviousCol(.Col)
                        Else
                            If .Row > PropFixedRows Then
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
                        .TopRow = .Row
                    ElseIf .Row > (.TopRow + GetRowsPerPage(.TopRow) - 1) Then
                        .TopRow = .Row - GetRowsPerPageRev(.Row) + 1
                    End If
                    If .LeftCol > .Col Then
                        .LeftCol = .Col
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
                        If .Row > PropFixedRows Then
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
                        .TopRow = .Row
                    ElseIf .Row > (.TopRow + GetRowsPerPage(.TopRow) - 1) Then
                        .TopRow = .Row - GetRowsPerPageRev(.Row) + 1
                    End If
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    Call MovePreviousRow(.RowSel)
                    If .TopRow > .RowSel Then .TopRow = .RowSel
                ElseIf (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                    Call MoveFirstRow(.Row)
                    .RowSel = .Row
                    .ColSel = (PropCols - 1)
                    .TopRow = .Row
                Else
                    Call MoveFirstRow(.RowSel)
                    .TopRow = .RowSel
                End If
            Case vbKeyDown
                If (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) = 0 Then
                    If PropWrapCellBehavior = FlexWrapGrid Then
                        If .Row < (PropRows - 1) Then
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
                        .TopRow = .Row
                    ElseIf .Row > (.TopRow + GetRowsPerPage(.TopRow) - 1) Then
                        .TopRow = .Row - GetRowsPerPageRev(.Row) + 1
                    End If
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    Call MoveNextRow(.RowSel)
                    If .TopRow > .RowSel Then
                        .TopRow = .RowSel
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
                        .TopRow = .Row
                    ElseIf .Row > (.TopRow + GetRowsPerPage(.TopRow) - 1) Then
                        .TopRow = .Row - GetRowsPerPageRev(.Row) + 1
                    End If
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    ' Void
                ElseIf (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                    .RowSel = .Row
                    .ColSel = (PropCols - 1)
                    If .TopRow > .Row Then
                        .TopRow = .Row
                    ElseIf .Row > (.TopRow + GetRowsPerPage(.TopRow) - 1) Then
                        .TopRow = .Row - GetRowsPerPageRev(.Row) + 1
                    End If
                Else
                    ' Void
                End If
            Case vbKeyRight
                If (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) = 0 Then
                    .RowSel = .Row
                    .ColSel = (PropCols - 1)
                    If .TopRow > .Row Then
                        .TopRow = .Row
                    ElseIf .Row > (.TopRow + GetRowsPerPage(.TopRow) - 1) Then
                        .TopRow = .Row - GetRowsPerPageRev(.Row) + 1
                    End If
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    ' Void
                ElseIf (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                    .RowSel = .Row
                    .ColSel = (PropCols - 1)
                    If .TopRow > .Row Then
                        .TopRow = .Row
                    ElseIf .Row > (.TopRow + GetRowsPerPage(.TopRow) - 1) Then
                        .TopRow = .Row - GetRowsPerPageRev(.Row) + 1
                    End If
                Else
                    ' Void
                End If
            Case vbKeyPageUp
                If (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) = 0 Then
                    If .Row > PropFixedRows Then
                        RowsPerPage = GetRowsPerPageRev(.Row)
                        If (.Row - RowsPerPage) > PropFixedRows Then
                            .Row = .Row - RowsPerPage
                        Else
                            .Row = PropFixedRows
                        End If
                    End If
                    .RowSel = .Row
                    .ColSel = (PropCols - 1)
                    If .TopRow > .Row Then
                        .TopRow = .Row
                    Else
                        .TopRow = .Row - GetRowsPerPageRev(.Row) + 1
                    End If
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    If .RowSel > PropFixedRows Then
                        RowsPerPage = GetRowsPerPageRev(.RowSel)
                        If (.RowSel - RowsPerPage) > PropFixedRows Then
                            .RowSel = .RowSel - RowsPerPage
                        Else
                            .RowSel = PropFixedRows
                        End If
                    End If
                    If .TopRow > .RowSel Then
                        .TopRow = .RowSel
                    Else
                        .TopRow = .RowSel - GetRowsPerPageRev(.RowSel) + 1
                    End If
                ElseIf (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                    Call MoveFirstRow(.Row)
                    .RowSel = .Row
                    .ColSel = (PropCols - 1)
                    .TopRow = .Row
                Else
                    Call MoveFirstRow(.RowSel)
                    .TopRow = .RowSel
                End If
            Case vbKeyPageDown
                If (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) = 0 Then
                    If .Row < (PropRows - 1) Then
                        RowsPerPage = GetRowsPerPage(.Row)
                        If (.Row + RowsPerPage) < (PropRows - 1) Then
                            .Row = .Row + RowsPerPage
                        Else
                            .Row = (PropRows - 1)
                        End If
                    End If
                    .RowSel = .Row
                    .ColSel = (PropCols - 1)
                    If .TopRow > .Row Then
                        .TopRow = .Row
                    Else
                        .TopRow = .Row - GetRowsPerPageRev(.Row) + 1
                    End If
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    If .RowSel < (PropRows - 1) Then
                        RowsPerPage = GetRowsPerPage(.RowSel)
                        If (.RowSel + RowsPerPage) < (PropRows - 1) Then
                            .RowSel = .RowSel + RowsPerPage
                        Else
                            .RowSel = (PropRows - 1)
                        End If
                    End If
                    If .TopRow > .RowSel Then
                        .TopRow = .RowSel
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
                        .TopRow = .Row
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
                        .TopRow = .Row
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
                        If .Row < (PropRows - 1) Then
                            Call MoveNextRow(.Row)
                        Else
                            If PropWrapCellBehavior = FlexWrapGrid Then Call MoveFirstRow(.Row)
                        End If
                    End If
                    .RowSel = .Row
                    .ColSel = (PropCols - 1)
                    If .TopRow > .Row Then
                        .TopRow = .Row
                    ElseIf .Row > (.TopRow + GetRowsPerPage(.TopRow) - 1) Then
                        .TopRow = .Row - GetRowsPerPageRev(.Row) + 1
                    End If
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    If PropWrapCellBehavior <> FlexWrapNone Then
                        If .Row > PropFixedRows Then
                            Call MovePreviousRow(.Row)
                        Else
                            If PropWrapCellBehavior = FlexWrapGrid Then Call MoveLastRow(.Row)
                        End If
                    End If
                    .RowSel = .Row
                    .ColSel = (PropCols - 1)
                    If .TopRow > .Row Then
                        .TopRow = .Row
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
                    If .LeftCol > .Col Then
                        .LeftCol = .Col
                    ElseIf .Col > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .Col - GetColsPerPageRev(.Col) + 1
                    End If
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    ' Void
                ElseIf (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                    .RowSel = (PropRows - 1)
                    .ColSel = .Col
                    If .LeftCol > .Col Then
                        .LeftCol = .Col
                    ElseIf .Col > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .Col - GetColsPerPageRev(.Col) + 1
                    End If
                Else
                    ' Void
                End If
            Case vbKeyDown
                If (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) = 0 Then
                    .RowSel = (PropRows - 1)
                    .ColSel = .Col
                    If .LeftCol > .Col Then
                        .LeftCol = .Col
                    ElseIf .Col > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .Col - GetColsPerPageRev(.Col) + 1
                    End If
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    ' Void
                ElseIf (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                    .RowSel = (PropRows - 1)
                    .ColSel = .Col
                    If .LeftCol > .Col Then
                        .LeftCol = .Col
                    ElseIf .Col > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .Col - GetColsPerPageRev(.Col) + 1
                    End If
                Else
                    ' Void
                End If
            Case vbKeyLeft
                If (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) = 0 Then
                    If PropWrapCellBehavior = FlexWrapGrid Then
                        If .Col > PropFixedCols Then
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
                        .LeftCol = .Col
                    ElseIf .Col > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .Col - GetColsPerPageRev(.Col) + 1
                    End If
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    Call MovePreviousCol(.ColSel)
                    If .LeftCol > .ColSel Then .LeftCol = .ColSel
                ElseIf (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                    Call MoveFirstCol(.Col)
                    .RowSel = (PropRows - 1)
                    .ColSel = .Col
                    .LeftCol = .Col
                Else
                    Call MoveFirstCol(.ColSel)
                    .LeftCol = .ColSel
                End If
            Case vbKeyRight
                If (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) = 0 Then
                    If PropWrapCellBehavior = FlexWrapGrid Then
                        If .Col < (PropCols - 1) Then
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
                        .LeftCol = .Col
                    ElseIf .Col > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .Col - GetColsPerPageRev(.Col) + 1
                    End If
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    Call MoveNextCol(.ColSel)
                    If .LeftCol > .ColSel Then
                        .LeftCol = .ColSel
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
                        .LeftCol = .Col
                    ElseIf .ColSel > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .ColSel - GetColsPerPageRev(.ColSel) + 1
                    End If
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    ' Void
                ElseIf (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                    .RowSel = (PropRows - 1)
                    .ColSel = .Col
                    If .LeftCol > .Col Then
                        .LeftCol = .Col
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
                        .LeftCol = .Col
                    ElseIf .ColSel > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .ColSel - GetColsPerPageRev(.ColSel) + 1
                    End If
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    ' Void
                ElseIf (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                    .RowSel = (PropRows - 1)
                    .ColSel = .Col
                    If .LeftCol > .Col Then
                        .LeftCol = .Col
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
                    .LeftCol = .Col
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    Call MoveFirstCol(.ColSel)
                    .LeftCol = .ColSel
                ElseIf (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                    .Row = PropFixedRows
                    Call MoveFirstCol(.Col)
                    .RowSel = (PropRows - 1)
                    .ColSel = .Col
                    .LeftCol = .Col
                Else
                    .RowSel = (PropRows - 1)
                    Call MoveFirstCol(.ColSel)
                    .LeftCol = .ColSel
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
                        If .Col < (PropCols - 1) Then
                            Call MoveNextCol(.Col)
                        Else
                            If PropWrapCellBehavior = FlexWrapGrid Then Call MoveFirstCol(.Col)
                        End If
                    End If
                    .RowSel = (PropRows - 1)
                    .ColSel = .Col
                    If .LeftCol > .Col Then
                        .LeftCol = .Col
                    ElseIf .Col > (.LeftCol + GetColsPerPage(.LeftCol) - 1) Then
                        .LeftCol = .Col - GetColsPerPageRev(.Col) + 1
                    End If
                ElseIf (Shift And vbShiftMask) <> 0 And (Shift And vbCtrlMask) = 0 Then
                    If PropWrapCellBehavior <> FlexWrapNone Then
                        If .Col > PropFixedCols Then
                            Call MovePreviousCol(.Col)
                        Else
                            If PropWrapCellBehavior = FlexWrapGrid Then Call MoveLastCol(.Col)
                        End If
                    End If
                    .RowSel = (PropRows - 1)
                    .ColSel = .Col
                    If .LeftCol > .Col Then
                        .LeftCol = .Col
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
        Case FlexSelectionModeFree
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

Private Sub ProcessLButtonDown(ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)
Dim HTI As THITTESTINFO
HTI.PT.X = X
HTI.PT.Y = Y
Call GetHitTestInfo(HTI)
VBFlexGridCaptureRow = HTI.HitRow
VBFlexGridCaptureCol = HTI.HitCol
VBFlexGridCaptureHitResult = HTI.HitResult
VBFlexGridMouseMoveRow = HTI.HitRow
VBFlexGridMouseMoveCol = HTI.HitCol
VBFlexGridMouseMoveChanged = False
If HTI.HitResult = FlexHitResultNoWhere Then
    Exit Sub
ElseIf HTI.HitResult <> FlexHitResultCell Then
    VBFlexGridCaptureDividerDrag = True
    Dim iRow As Long, iCol As Long, Cancel As Boolean
    Select Case VBFlexGridCaptureHitResult
        Case FlexHitResultDividerRowTop
            iRow = VBFlexGridCaptureRow - 1
            iCol = -1
        Case FlexHitResultDividerRowBottom
            iRow = VBFlexGridCaptureRow
            iCol = -1
        Case FlexHitResultDividerColumnLeft
            iRow = -1
            iCol = VBFlexGridCaptureCol - 1
        Case FlexHitResultDividerColumnRight
            iRow = -1
            iCol = VBFlexGridCaptureCol
    End Select
    RaiseEvent BeforeUserResize(iRow, iCol, Cancel)
    If Cancel = False Then
        Dim ClipRect As RECT, i As Long
        GetClientRect VBFlexGridHandle, ClipRect
        With ClipRect
        If iRow > -1 Then
            For i = 0 To iRow - 1
                If i >= VBFlexGridTopRow Or i < PropFixedRows Then
                    .Top = .Top + GetRowHeight(i)
                End If
            Next i
            .Top = .Top + (1 * PixelsPerDIP_Y())
            .Bottom = .Bottom - (1 * PixelsPerDIP_Y())
        End If
        If iCol > -1 Then
            For i = 0 To iCol - 1
                If i >= VBFlexGridLeftCol Or i < PropFixedCols Then
                    .Left = .Left + GetColWidth(i)
                End If
            Next i
            .Left = .Left + (1 * PixelsPerDIP_X())
            .Right = .Right - (1 * PixelsPerDIP_X())
        End If
        End With
        MapWindowPoints VBFlexGridHandle, HWND_DESKTOP, ClipRect, 2
        ClipCursor ClipRect
        Call SetDividerDragSplitterRect(X, Y)
        Call DrawDividerDragSplitter
        Exit Sub
    Else
        ReleaseCapture
        Exit Sub
    End If
End If
Dim RCP As TROWCOLPARAMS
With RCP
.Mask = RCPM_ROW Or RCPM_COL Or RCPM_ROWSEL Or RCPM_COLSEL
.Row = VBFlexGridRow
.Col = VBFlexGridCol
.RowSel = VBFlexGridRowSel
.ColSel = VBFlexGridColSel
Select Case PropSelectionMode
    Case FlexSelectionModeFree
        If HTI.HitRow > (PropFixedRows - 1) Then
            If (Shift And vbShiftMask) = 0 Then
                .Row = HTI.HitRow
                .RowSel = .Row
            Else
                .RowSel = HTI.HitRow
            End If
        Else
            If PropAllowBigSelection = True Then
                .Row = PropFixedRows
                .RowSel = (PropRows - 1)
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
                .Col = PropFixedCols
                .ColSel = (PropCols - 1)
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
        Case FlexSelectionModeFree
            If PropAllowBigSelection = True Or (Shift And vbShiftMask) = 0 Then
                .Mask = .Mask Or RCPM_TOPROW Or RCPM_LEFTCOL
                .TopRow = .Row
                .LeftCol = .Col
            End If
        Case FlexSelectionModeByRow, FlexSelectionModeByColumn
            If PropAllowBigSelection = True And (Shift And vbShiftMask) = 0 Then
                .Mask = .Mask Or RCPM_TOPROW Or RCPM_LEFTCOL
                .TopRow = .Row
                .LeftCol = .Col
            End If
    End Select
End If
Call SetRowColParams(RCP)
End With
End Sub

Private Sub ProcessLButtonUp(ByVal X As Long, ByVal Y As Long)
Dim RCP As TROWCOLPARAMS
If VBFlexGridCaptureDividerDrag = True Then
    Dim iRow As Long, iCol As Long, NewSize As Long
    Select Case VBFlexGridCaptureHitResult
        Case FlexHitResultDividerRowTop
            iRow = VBFlexGridCaptureRow - 1
            iCol = -1
        Case FlexHitResultDividerRowBottom
            iRow = VBFlexGridCaptureRow
            iCol = -1
        Case FlexHitResultDividerColumnLeft
            iRow = -1
            iCol = VBFlexGridCaptureCol - 1
        Case FlexHitResultDividerColumnRight
            iRow = -1
            iCol = VBFlexGridCaptureCol
    End Select
    Dim ClientRect As RECT, Size As SIZEAPI, i As Long
    GetClientRect VBFlexGridHandle, ClientRect
    With Size
    If iRow > -1 Then
        For i = 0 To iRow - 1
            If i >= VBFlexGridTopRow Or i < PropFixedRows Then
                .CY = .CY + GetRowHeight(i)
            End If
        Next i
        If Y < (.CY + (1 * PixelsPerDIP_Y())) Then
            NewSize = UserControl.ScaleY(1, vbPixels, vbTwips)
        ElseIf Y >= (ClientRect.Bottom - (1 * PixelsPerDIP_Y())) Then
            NewSize = UserControl.ScaleY(((ClientRect.Bottom - 1) - .CY), vbPixels, vbTwips)
        Else
            NewSize = UserControl.ScaleY((Y - .CY), vbPixels, vbTwips)
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
            If i >= VBFlexGridLeftCol Or i < PropFixedCols Then
                .CX = .CX + GetColWidth(i)
            End If
        Next i
        If X < (.CX + (1 * PixelsPerDIP_X())) Then
            NewSize = UserControl.ScaleX(1, vbPixels, vbTwips)
        ElseIf X >= (ClientRect.Right - (1 * PixelsPerDIP_X())) Then
            NewSize = UserControl.ScaleX(((ClientRect.Right - 1) - .CX), vbPixels, vbTwips)
        Else
            NewSize = UserControl.ScaleX((X - .CX), vbPixels, vbTwips)
        End If
        RaiseEvent AfterUserResize(iRow, iCol, NewSize)
        If NewSize > 0 Then .CX = UserControl.ScaleX(NewSize, vbTwips, vbPixels) Else .CX = 0
        VBFlexGridColsInfo(iCol).Width = .CX
    End If
    End With
    ClipCursor ByVal 0&
    SetRect VBFlexGridDividerDragSplitterRect, 0, 0, 0, 0
    With RCP
    .Mask = RCPM_TOPROW Or RCPM_LEFTCOL
    .Flags = RCPF_CHECKTOPROW Or RCPF_CHECKLEFTCOL Or RCPF_SETSCROLLBARS
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
If VBFlexGridCaptureRow = -1 Or VBFlexGridCaptureCol = -1 Or VBFlexGridCaptureHitResult = FlexHitResultNoWhere Then Exit Sub
If VBFlexGridCaptureDividerDrag = True Then
    Call DrawDividerDragSplitter
    Call SetDividerDragSplitterRect(X, Y)
    Call DrawDividerDragSplitter
    Exit Sub
End If
If (Button And vbLeftButton) = 0 Then Exit Sub
If VBFlexGridCaptureRow <= (PropFixedRows - 1) And VBFlexGridCaptureCol <= (PropFixedCols - 1) Then Exit Sub
Dim HTI As THITTESTINFO
HTI.PT.X = X
HTI.PT.Y = Y
Call GetHitTestInfo(HTI)
If HTI.HitRow <> VBFlexGridMouseMoveRow Or HTI.HitCol <> VBFlexGridMouseMoveCol Or HTI.HitRow <= (PropFixedRows - 1) Or HTI.HitCol <= (PropFixedCols - 1) Then
    VBFlexGridMouseMoveRow = HTI.HitRow
    VBFlexGridMouseMoveCol = HTI.HitCol
    VBFlexGridMouseMoveChanged = True
Else
    Exit Sub
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
    Case FlexSelectionModeFree
        If VBFlexGridCaptureRow > (PropFixedRows - 1) Or PropAllowBigSelection = False Then
            If HTI.MouseRow > (PropFixedRows - 1) Then
                .RowSel = HTI.MouseRow
            Else
                If .RowSel > PropFixedRows Then .RowSel = .RowSel - 1
            End If
            If .TopRow > .RowSel Then
                .TopRow = .RowSel
            Else
                RowsPerPage = GetRowsPerPage(.TopRow)
                If .RowSel > (.TopRow + RowsPerPage - 1) Then
                    .TopRow = .RowSel - GetRowsPerPageRev(.RowSel) + 1
                End If
            End If
            If PropAllowSelection = False Then .Row = .RowSel
        Else
            .RowSel = (PropRows - 1)
        End If
        If VBFlexGridCaptureCol > (PropFixedCols - 1) Or PropAllowBigSelection = False Then
            If HTI.MouseCol > (PropFixedCols - 1) Then
                .ColSel = HTI.MouseCol
            Else
                If .ColSel > PropFixedCols Then .ColSel = .ColSel - 1
            End If
            If .LeftCol > .ColSel Then
                .LeftCol = .ColSel
            Else
                ColsPerPage = GetColsPerPage(.LeftCol)
                If .ColSel > (.LeftCol + ColsPerPage - 1) Then
                    .LeftCol = .ColSel - GetColsPerPageRev(.ColSel) + 1
                End If
            End If
            If PropAllowSelection = False Then .Col = .ColSel
        Else
            .ColSel = (PropCols - 1)
        End If
    Case FlexSelectionModeByRow
        If VBFlexGridCaptureRow > (PropFixedRows - 1) Or VBFlexGridCaptureCol > (PropFixedCols - 1) Or PropAllowBigSelection = False Then
            If HTI.MouseRow > (PropFixedRows - 1) Then
                .RowSel = HTI.MouseRow
            Else
                If .RowSel > PropFixedRows Then .RowSel = .RowSel - 1
            End If
            If .TopRow > .RowSel Then
                .TopRow = .RowSel
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
            If HTI.MouseCol > (PropFixedCols - 1) Then
                .ColSel = HTI.MouseCol
            Else
                If .ColSel > PropFixedCols Then .ColSel = .ColSel - 1
            End If
            If .LeftCol > .ColSel Then
                .LeftCol = .ColSel
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

Private Function MergeCompareFunction(ByVal Text1 As String, ByVal Text2 As String) As Boolean
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
GetClientRect VBFlexGridHandle, VBFlexGridDividerDragSplitterRect
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
    If VBFlexGridToolTipHandle <> 0 And EnabledVisualStyles() = True Then
        If PropVisualStyles = True Then
            ActivateVisualStyles VBFlexGridToolTipHandle
        Else
            RemoveVisualStyles VBFlexGridToolTipHandle
        End If
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
    GetClientRect VBFlexGridHandle, .RC
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

Friend Function FSubclass_Message(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long
Select Case dwRefData
    Case 1
        FSubclass_Message = WindowProcControl(hWnd, wMsg, wParam, lParam)
    Case 2
        FSubclass_Message = WindowProcUserControl(hWnd, wMsg, wParam, lParam)
End Select
End Function

Private Function WindowProcControl(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim HTI As THITTESTINFO, Pos As Long
Select Case wMsg
    Case WM_SETFOCUS
        If wParam <> UserControl.hWnd Then SetFocusAPI UserControl.hWnd: Exit Function
        Call ActivateIPAO(Me)
    Case WM_KILLFOCUS
        Call DeActivateIPAO
    Case WM_GETFONT
        WindowProcControl = VBFlexGridFontHandle
        Exit Function
    Case WM_SETREDRAW
        VBFlexGridNoRedraw = CBool(wParam = 0)
        WindowProcControl = 0
        Exit Function
    Case WM_SIZE
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
                            SCI.nPos = SCI.nPos - PropFixedCols
                        Else
                            SCI.nPos = SCI.nMin
                        End If
                    ElseIf wMsg = WM_VSCROLL Then
                        SCI.nPos = VBFlexGridTopRow
                        Call MovePreviousRow(SCI.nPos)
                        If SCI.nPos < VBFlexGridTopRow Then
                            SCI.nPos = SCI.nPos - PropFixedRows
                        Else
                            SCI.nPos = SCI.nMin
                        End If
                    End If
                Case SB_LINERIGHT, SB_LINEDOWN
                    If wMsg = WM_HSCROLL Then
                        SCI.nPos = VBFlexGridLeftCol
                        Call MoveNextCol(SCI.nPos)
                        If SCI.nPos > VBFlexGridLeftCol Then
                            SCI.nPos = SCI.nPos - PropFixedCols
                        Else
                            SCI.nPos = SCI.nMax
                        End If
                    ElseIf wMsg = WM_VSCROLL Then
                        SCI.nPos = VBFlexGridTopRow
                        Call MoveNextRow(SCI.nPos)
                        If SCI.nPos > VBFlexGridTopRow Then
                            SCI.nPos = SCI.nPos - PropFixedRows
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
                    VBFlexGridLeftCol = PropFixedCols + SCI.nPos
                ElseIf wMsg = WM_VSCROLL Then
                    VBFlexGridTopRow = PropFixedRows + SCI.nPos
                End If
                Call RedrawGrid
                If PropShowInfoTips = True Or PropShowLabelTips = True Then
                    Pos = GetMessagePos()
                    Call CheckToolTipRowCol(Get_X_lParam(Pos), Get_Y_lParam(Pos))
                End If
                RaiseEvent Scroll
            End If
            WindowProcControl = 0
            Exit Function
        End If
    Case WM_PAINT, WM_PRINTCLIENT
        Dim ClientRect As RECT, hDC As Long, hRgn As Long
        Dim hDCBmp As Long
        Dim hBmp As Long, hBmpOld As Long
        GetClientRect hWnd, ClientRect
        If wMsg = WM_PRINTCLIENT Then
            hDC = wParam
            hDCBmp = CreateCompatibleDC(hDC)
            If hDCBmp <> 0 Then
                hBmp = CreateCompatibleBitmap(hDC, ClientRect.Right - ClientRect.Left, ClientRect.Bottom - ClientRect.Top)
                If hBmp <> 0 Then
                    hBmpOld = SelectObject(hDCBmp, hBmp)
                    If SendMessage(hWnd, WM_ERASEBKGND, hDCBmp, ByVal 0&) = 0 Then
                        If VBFlexGridBackColorBkgBrush <> 0 Then FillRect hDCBmp, ClientRect, VBFlexGridBackColorBkgBrush
                    End If
                    Call DrawGrid(hDCBmp, -1)
                    BitBlt hDC, 0, 0, ClientRect.Right - ClientRect.Left, ClientRect.Bottom - ClientRect.Top, hDCBmp, 0, 0, vbSrcCopy
                    SelectObject hDCBmp, hBmpOld
                    DeleteObject hBmp
                End If
                DeleteDC hDCBmp
            End If
            WindowProcControl = 0
            Exit Function
        End If
        Dim PS As PAINTSTRUCT
        hDC = BeginPaint(hWnd, PS)
        With PS
        If wParam <> 0 Then hDC = wParam
        If PropDoubleBuffer = True Then
            hDCBmp = CreateCompatibleDC(hDC)
            If hDCBmp <> 0 Then
                hBmp = CreateCompatibleBitmap(hDC, ClientRect.Right - ClientRect.Left, ClientRect.Bottom - ClientRect.Top)
                If hBmp <> 0 Then
                    hBmpOld = SelectObject(hDCBmp, hBmp)
                    If .fErase <> 0 Then
                        If VBFlexGridBackColorBkgBrush <> 0 Then FillRect hDCBmp, ClientRect, VBFlexGridBackColorBkgBrush
                        Call DrawGrid(hDCBmp, -1)
                    Else
                        Call DrawGrid(hDCBmp, hRgn)
                        If hRgn <> 0 Then ExtSelectClipRgn hDC, hRgn, RGN_COPY
                    End If
                    With PS.RCPaint
                    BitBlt hDC, .Left, .Top, .Right - .Left, .Bottom - .Top, hDCBmp, .Left, .Top, vbSrcCopy
                    End With
                    If hRgn <> 0 Then
                        ExtSelectClipRgn hDC, 0, RGN_COPY
                        DeleteObject hRgn
                    End If
                    SelectObject hDCBmp, hBmpOld
                    DeleteObject hBmp
                End If
                DeleteDC hDCBmp
            End If
        Else
            If .fErase <> 0 Then
                Call DrawGrid(hDC, hRgn)
                If hRgn <> 0 Then ExtSelectClipRgn hDC, hRgn, RGN_DIFF
                If VBFlexGridBackColorBkgBrush <> 0 Then FillRect hDC, ClientRect, VBFlexGridBackColorBkgBrush
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
        WindowProcControl = 0
        Exit Function
    Case WM_MOUSEACTIVATE
        Static InProc As Boolean
        If FlexRootIsEditor(hWnd) = False And GetFocus() <> VBFlexGridHandle Then
            If InProc = True Or LoWord(lParam) = HTBORDER Then WindowProcControl = MA_NOACTIVATEANDEAT: Exit Function
            Select Case HiWord(lParam)
                Case WM_LBUTTONDOWN
                    On Error Resume Next
                    With UserControl
                    If .Extender.CausesValidation = True Then
                        InProc = True
                        Call FlexTopParentValidateControls(Me)
                        InProc = False
                        If Err.Number = 380 Then
                            WindowProcControl = MA_NOACTIVATEANDEAT
                        Else
                            SetFocusAPI .hWnd
                            WindowProcControl = MA_NOACTIVATE
                        End If
                    Else
                        SetFocusAPI .hWnd
                        WindowProcControl = MA_NOACTIVATE
                    End If
                    End With
                    On Error GoTo 0
                    Exit Function
            End Select
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
    Case WM_KEYDOWN, WM_KEYUP
        Dim KeyCode As Integer
        KeyCode = wParam And &HFF&
        If wMsg = WM_KEYDOWN Then
            RaiseEvent KeyDown(KeyCode, GetShiftStateFromMsg())
        ElseIf wMsg = WM_KEYUP Then
            RaiseEvent KeyUp(KeyCode, GetShiftStateFromMsg())
        End If
        Dim Msg As TMSG
        Const PM_NOREMOVE As Long = &H0
        If PeekMessage(Msg, hWnd, WM_CHAR, WM_CHAR, PM_NOREMOVE) <> 0 Then VBFlexGridCharCodeCache = Msg.wParam
        If wMsg = WM_KEYDOWN Then Call ProcessKeyDown(KeyCode, GetShiftStateFromMsg())
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
    Case WM_UNICHAR
        If wParam = UNICODE_NOCHAR Then WindowProcControl = 1 Else SendMessage hWnd, WM_CHAR, wParam, ByVal lParam
        Exit Function
    Case WM_IME_CHAR
        SendMessage hWnd, WM_CHAR, wParam, ByVal lParam
        Exit Function
    Case WM_LBUTTONDOWN, WM_MBUTTONDOWN, WM_RBUTTONDOWN
        Dim P As POINTAPI, Cancel As Boolean
        P.X = Get_X_lParam(lParam)
        P.Y = Get_Y_lParam(lParam)
        Select Case wMsg
            Case WM_LBUTTONDOWN
                RaiseEvent BeforeMouseDown(vbLeftButton, GetShiftStateFromParam(wParam), UserControl.ScaleX(P.X, vbPixels, vbTwips), UserControl.ScaleY(P.Y, vbPixels, vbTwips), Cancel)
                If Cancel = False Then
                    SetCapture hWnd
                    Call ProcessLButtonDown(GetShiftStateFromParam(wParam), P.X, P.Y)
                End If
            Case WM_MBUTTONDOWN
                RaiseEvent BeforeMouseDown(vbMiddleButton, GetShiftStateFromParam(wParam), UserControl.ScaleX(P.X, vbPixels, vbTwips), UserControl.ScaleY(P.Y, vbPixels, vbTwips), Cancel)
            Case WM_RBUTTONDOWN
                RaiseEvent BeforeMouseDown(vbRightButton, GetShiftStateFromParam(wParam), UserControl.ScaleX(P.X, vbPixels, vbTwips), UserControl.ScaleY(P.Y, vbPixels, vbTwips), Cancel)
        End Select
        If Cancel = True Then
            WindowProcControl = 0
            Exit Function
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
                        With VBFlexGridCells.Rows(.HitRow).Cols(.HitCol)
                        If (LBLI.Flags And LBLI_VALID) = LBLI_VALID And Not (LBLI.Flags And LBLI_UNFOLDED) = LBLI_UNFOLDED Then
                            Text = .Text
                        ElseIf PropShowInfoTips = True Then
                            Text = .ToolTipText
                            ShowInfoTip = True
                        End If
                        End With
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
                            SetWindowPos VBFlexGridToolTipHandle, 0, RC.Left, RC.Top, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
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
    Case WM_NOTIFYFORMAT
        Const NF_QUERY As Long = 3
        If wParam = VBFlexGridToolTipHandle And VBFlexGridToolTipHandle <> 0 And lParam = NF_QUERY Then
            Const NFR_ANSI As Long = 1
            Const NFR_UNICODE As Long = 2
            WindowProcControl = NFR_UNICODE
            Exit Function
        End If
End Select
WindowProcControl = DefWindowProc(hWnd, wMsg, wParam, lParam)
Select Case wMsg
    Case WM_SETFOCUS, WM_KILLFOCUS
        VBFlexGridFocused = CBool(wMsg = WM_SETFOCUS)
        Call RedrawGrid
    Case WM_LBUTTONDBLCLK, WM_MBUTTONDBLCLK, WM_RBUTTONDBLCLK
        RaiseEvent DblClick
        If wMsg = WM_LBUTTONDBLCLK Then
            With HTI
            Pos = GetMessagePos()
            .PT.X = Get_X_lParam(Pos)
            .PT.Y = Get_Y_lParam(Pos)
            ScreenToClient hWnd, .PT
            Call GetHitTestInfo(HTI)
            Select Case .HitResult
                Case FlexHitResultDividerRowTop
                    RaiseEvent DividerDblClick(.HitRow - 1, -1)
                Case FlexHitResultDividerRowBottom
                    RaiseEvent DividerDblClick(.HitRow, -1)
                Case FlexHitResultDividerColumnLeft
                    RaiseEvent DividerDblClick(-1, .HitCol - 1)
                Case FlexHitResultDividerColumnRight
                    RaiseEvent DividerDblClick(-1, .HitCol)
            End Select
            End With
        End If
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

Private Function WindowProcUserControl(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_CONTEXTMENU
        If wParam = VBFlexGridHandle Then
            Dim P As POINTAPI
            P.X = Get_X_lParam(lParam)
            P.Y = Get_Y_lParam(lParam)
            If P.X > 0 And P.Y > 0 Then
                ScreenToClient VBFlexGridHandle, P
                RaiseEvent ContextMenu(UserControl.ScaleX(P.X, vbPixels, vbContainerPosition), UserControl.ScaleY(P.Y, vbPixels, vbContainerPosition))
            ElseIf P.X = -1 And P.Y = -1 Then
                ' If the user types SHIFT + F10 then the X and Y coordinates are -1.
                RaiseEvent ContextMenu(-1, -1)
            End If
        End If
End Select
WindowProcUserControl = FlexDefaultProc(hWnd, wMsg, wParam, lParam)
If wMsg = WM_SETFOCUS Then SetFocusAPI VBFlexGridHandle
End Function
