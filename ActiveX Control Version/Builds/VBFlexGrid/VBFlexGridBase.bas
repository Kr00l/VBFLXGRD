Attribute VB_Name = "VBFlexGridBase"
Option Explicit
#If (VBA7 = 0) Then
Private Enum LongPtr
[_]
End Enum
#End If
#If Win64 Then
Private Const NULL_PTR As LongPtr = 0
Private Const PTR_SIZE As Long = 8
#Else
Private Const NULL_PTR As Long = 0
Private Const PTR_SIZE As Long = 4
#End If

#Const ImplementPreTranslateMsg = (VBFLXGRD_OCX <> 0)

Private Type TINITCOMMONCONTROLSEX
dwSize As Long
dwICC As Long
End Type
Private Type WNDCLASSEX
cbSize As Long
dwStyle As Long
lpfnWndProc As LongPtr
cbClsExtra As Long
cbWndExtra As Long
hInstance As LongPtr
hIcon As LongPtr
hCursor As LongPtr
hbrBackground As LongPtr
lpszMenuName As LongPtr
lpszClassName As LongPtr
hIconSm As LongPtr
End Type
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
Private Type TMSG
hWnd As LongPtr
Message As Long
wParam As LongPtr
lParam As LongPtr
Time As Long
PT As POINTAPI
End Type
Private Type CURSORINFO
cbSize As Long
Flags As Long
hCursor As LongPtr
PTScreenPos As POINTAPI
End Type
Private Type ICONINFO
fIcon As Long
XHotspot As Long
YHotspot As Long
hBMMask As LongPtr
hBMColor As LongPtr
End Type
Private Type BITMAP
BMType As Long
BMWidth As Long
BMHeight As Long
BMWidthBytes As Long
BMPlanes As Integer
BMBitsPixel As Integer
BMBits As LongPtr
End Type
Private Const RMF_ZEROCURSOR As Long = &H1
Private Const RMF_VERTICALONLY As Long = &H2
Private Const RMF_HORIZONTALONLY As Long = &H4
Private Type READERMODEINFO
cbSize As Long
hWnd As LongPtr
dwFlags As Long
lpRC As LongPtr
lpfnScroll As LongPtr
lpfnDispatch As LongPtr
lParam As LongPtr
End Type
#If VBA7 Then
Public Declare PtrSafe Function FlexObjAddRef Lib "msvbvm60.dll" Alias "__vbaObjAddref" (ByVal lpObject As LongPtr) As Long
Public Declare PtrSafe Function FlexObjSet Lib "msvbvm60.dll" Alias "__vbaObjSet" (ByRef Destination As Any, ByVal lpObject As LongPtr) As Long
Public Declare PtrSafe Function FlexObjSetAddRef Lib "msvbvm60.dll" Alias "__vbaObjSetAddref" (ByRef Destination As Any, ByVal lpObject As LongPtr) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare PtrSafe Sub DoReaderMode Lib "comctl32" Alias "#383" (ByRef lpRMI As READERMODEINFO)
Private Declare PtrSafe Function InitCommonControlsEx Lib "comctl32" (ByRef ICCEX As TINITCOMMONCONTROLSEX) As Long
Private Declare PtrSafe Function GetClassInfoEx Lib "user32" Alias "GetClassInfoExW" (ByVal hInstance As LongPtr, ByVal lpClassName As LongPtr, ByRef lpWndClassEx As WNDCLASSEX) As Long
Private Declare PtrSafe Function RegisterClassEx Lib "user32" Alias "RegisterClassExW" (ByRef lpWndClassEx As WNDCLASSEX) As Integer
Private Declare PtrSafe Function UnregisterClass Lib "user32" Alias "UnregisterClassW" (ByVal lpClassName As LongPtr, ByVal hInstance As LongPtr) As Long
Private Declare PtrSafe Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, ByRef lpParam As Any) As LongPtr
Private Declare PtrSafe Function DefWindowProc Lib "user32" Alias "DefWindowProcW" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long
#If Win64 Then
Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrW" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrW" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
#Else
Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
#End If
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByRef lParam As Any) As LongPtr
Private Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageW" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByRef lParam As Any) As LongPtr
Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As LongPtr, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function DestroyWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hWnd As LongPtr, ByRef lpRect As RECT) As Long
Private Declare PtrSafe Function GetCursorInfo Lib "user32" (ByRef pCI As CURSORINFO) As LongPtr
Private Declare PtrSafe Function GetIconInfo Lib "user32" (ByVal hIcon As LongPtr, ByRef pIconInfo As ICONINFO) As Long
Private Declare PtrSafe Function DrawIconEx Lib "user32" (ByVal hDC As LongPtr, ByVal XLeft As Long, ByVal YTop As Long, ByVal hIcon As LongPtr, ByVal CXWidth As Long, ByVal CYWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As LongPtr, ByVal diFlags As Long) As Long
Private Declare PtrSafe Function CopyIcon Lib "user32" (ByVal hIcon As LongPtr) As LongPtr
Private Declare PtrSafe Function DestroyIcon Lib "user32" (ByVal hIcon As LongPtr) As Long
Private Declare PtrSafe Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As LongPtr, ByVal lpCursorName As Any) As LongPtr
Private Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal hModule As LongPtr, ByVal lpProcName As Any) As LongPtr
Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As LongPtr) As LongPTr
Private Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal hLibModule As LongPtr) As Long
Private Declare PtrSafe Function CreateBitmapIndirect Lib "gdi32" (ByRef lpBitmap As BITMAP) As LongPtr
Private Declare PtrSafe Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As LongPtr) As LongPtr
Private Declare PtrSafe Function GetObjectAPI Lib "gdi32" Alias "GetObjectW" (ByVal hObject As LongPtr, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function SetProp Lib "user32" Alias "SetPropW" (ByVal hWnd As LongPtr, ByVal lpString As LongPtr, ByVal hData As LongPtr) As Long
Private Declare PtrSafe Function GetProp Lib "user32" Alias "GetPropW" (ByVal hWnd As LongPtr, ByVal lpString As LongPtr) As LongPtr
Private Declare PtrSafe Function RemoveProp Lib "user32" Alias "RemovePropW" (ByVal hWnd As LongPtr, ByVal lpString As LongPtr) As LongPtr
Private Declare PtrSafe Function SetWindowSubclass Lib "comctl32" (ByVal hWnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As LongPtr) As Long
Private Declare PtrSafe Function RemoveWindowSubclass Lib "comctl32" (ByVal hWnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr) As Long
Private Declare PtrSafe Function DefSubclassProc Lib "comctl32" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
Public Declare Function FlexObjAddRef Lib "msvbvm60.dll" Alias "__vbaObjAddref" (ByVal lpObject As Long) As Long
Public Declare Function FlexObjSet Lib "msvbvm60.dll" Alias "__vbaObjSet" (ByRef Destination As Any, ByVal lpObject As Long) As Long
Public Declare Function FlexObjSetAddRef Lib "msvbvm60.dll" Alias "__vbaObjSetAddref" (ByRef Destination As Any, ByVal lpObject As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub DoReaderMode Lib "comctl32" Alias "#383" (ByRef lpRMI As READERMODEINFO)
Private Declare Function InitCommonControlsEx Lib "comctl32" (ByRef ICCEX As TINITCOMMONCONTROLSEX) As Long
Private Declare Function GetClassInfoEx Lib "user32" Alias "GetClassInfoExW" (ByVal hInstance As Long, ByVal lpClassName As Long, ByRef lpWndClassEx As WNDCLASSEX) As Long
Private Declare Function RegisterClassEx Lib "user32" Alias "RegisterClassExW" (ByRef lpWndClassEx As WNDCLASSEX) As Integer
Private Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassW" (ByVal lpClassName As Long, ByVal hInstance As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetCursorInfo Lib "user32" (ByRef pCI As CURSORINFO) As Long
Private Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, ByRef pIconInfo As ICONINFO) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal XLeft As Long, ByVal YTop As Long, ByVal hIcon As Long, ByVal CXWidth As Long, ByVal CYWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function CopyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As Any) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function CreateBitmapIndirect Lib "gdi32" (ByRef lpBitmap As BITMAP) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectW" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropW" (ByVal hWnd As Long, ByVal lpString As Long, ByVal hData As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropW" (ByVal hWnd As Long, ByVal lpString As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropW" (ByVal hWnd As Long, ByVal lpString As Long) As Long
Private Declare Function SetWindowSubclass Lib "comctl32" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
Private Declare Function RemoveWindowSubclass Lib "comctl32" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Long
Private Declare Function DefSubclassProc Lib "comctl32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowSubclassW2K Lib "comctl32" Alias "#410" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
Private Declare Function RemoveWindowSubclassW2K Lib "comctl32" Alias "#412" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Long
Private Declare Function DefSubclassProcW2K Lib "comctl32" Alias "#413" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If
Private Const GWL_STYLE As Long = (-16)
Private Const GWL_EXSTYLE As Long = (-20)
Private Const WM_CREATE As Long = &H1
Private Const WM_DESTROY As Long = &H2
Private Const WM_NCDESTROY As Long = &H82
Private Const WM_UAHDESTROYWINDOW As Long = &H90
Private Const CS_DBLCLKS As Long = &H8
Private Const CS_DROPSHADOW As Long = &H20000
Private Const IDC_ARROW As Long = 32512
Private Const WHITE_BRUSH As Long = 0
Private ShellModHandle As LongPtr, ShellModCount As Long
Private FlexSubclassProcPtr As LongPtr
#If (VBA7 = 0) Then
Private FlexSubclassW2K As Integer
#End If
Private FlexClassAtom As Integer, FlexRefCount As Long
Private FlexComboCalendarClassAtom As Integer, FlexComboCalendarRefCount As Long
Private FlexSplitterBrush As LongPtr
Private FlexReaderModeScrolled As Boolean, FlexReaderModeCursorInitialized As Boolean, FlexReaderModeAnchorClassAtom As Integer, FlexReaderModeAnchorRefCount As Long, FlexReaderModeAnchorHandle As LongPtr

#If ImplementPreTranslateMsg = True Then

#If VBA7 Then
Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExW" (ByVal IDHook As Long, ByVal lpfn As LongPtr, ByVal hMod As LongPtr, ByVal dwThreadID As Long) As LongPtr
Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As LongPtr) As Long
Private Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hHook As LongPtr, ByVal nCode As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExW" (ByVal IDHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadID As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If
Private Const WM_USER As Long = &H400
Private Const UM_PRETRANSLATEMSG As Long = (WM_USER + 333)
Private FlexPreTranslateMsgHookHandle As LongPtr
Private FlexPreTranslateMsgHwnd() As LongPtr, FlexPreTranslateMsgCount As Long

#End If

Public Sub FlexLoadShellMod()
If ShellModHandle = NULL_PTR And ShellModCount = 0 Then ShellModHandle = LoadLibrary(StrPtr("Shell32.dll"))
ShellModCount = ShellModCount + 1
End Sub

Public Sub FlexReleaseShellMod()
ShellModCount = ShellModCount - 1
If ShellModHandle <> NULL_PTR And ShellModCount = 0 Then
    FreeLibrary ShellModHandle
    ShellModHandle = NULL_PTR
End If
End Sub

Public Sub FlexInitCC(ByVal ICC As Long)
Dim ICCEX As TINITCOMMONCONTROLSEX
With ICCEX
.dwSize = LenB(ICCEX)
.dwICC = ICC
End With
InitCommonControlsEx ICCEX
End Sub

#If VBA7 Then
Public Sub FlexSetSubclass(ByVal hWnd As LongPtr, ByVal This As VBFlexGrid, ByVal dwRefData As LongPtr, Optional ByVal Name As String)
#Else
Public Sub FlexSetSubclass(ByVal hWnd As Long, ByVal This As VBFlexGrid, ByVal dwRefData As Long, Optional ByVal Name As String)
#End If
If hWnd = NULL_PTR Then Exit Sub
If Name = vbNullString Then Name = "Flex"
If GetProp(hWnd, StrPtr(Name & "SubclassInit")) = 0 Then
    If FlexSubclassProcPtr = NULL_PTR Then FlexSubclassProcPtr = ProcPtr(AddressOf FlexSubclassProc)
    #If VBA7 Then
    SetWindowSubclass hWnd, FlexSubclassProcPtr, ObjPtr(This), dwRefData
    #Else
    If FlexSubclassW2K = 0 Then
        Dim hLib As LongPtr
        hLib = LoadLibrary(StrPtr("comctl32.dll"))
        If hLib <> NULL_PTR Then
            If GetProcAddress(hLib, "SetWindowSubclass") <> NULL_PTR Then
                FlexSubclassW2K = 1
            ElseIf GetProcAddress(hLib, 410&) <> NULL_PTR Then
                FlexSubclassW2K = -1
            End If
            FreeLibrary hLib
        End If
    End If
    If FlexSubclassW2K > -1 Then
        SetWindowSubclass hWnd, FlexSubclassProcPtr, ObjPtr(This), dwRefData
    Else
        SetWindowSubclassW2K hWnd, FlexSubclassProcPtr, ObjPtr(This), dwRefData
    End If
    #End If
    SetProp hWnd, StrPtr(Name & "SubclassID"), ObjPtr(This)
    SetProp hWnd, StrPtr(Name & "SubclassInit"), 1
End If
End Sub

#If VBA7 Then
Public Function FlexDefaultProc(ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
Public Function FlexDefaultProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If
#If VBA7 Then
FlexDefaultProc = DefSubclassProc(hWnd, wMsg, wParam, lParam)
#Else
If FlexSubclassW2K > -1 Then
    FlexDefaultProc = DefSubclassProc(hWnd, wMsg, wParam, lParam)
Else
    FlexDefaultProc = DefSubclassProcW2K(hWnd, wMsg, wParam, lParam)
End If
#End If
End Function

#If VBA7 Then
Public Sub FlexRemoveSubclass(ByVal hWnd As LongPtr, Optional ByVal Name As String)
#Else
Public Sub FlexRemoveSubclass(ByVal hWnd As Long, Optional ByVal Name As String)
#End If
If hWnd = NULL_PTR Then Exit Sub
If Name = vbNullString Then Name = "Flex"
If GetProp(hWnd, StrPtr(Name & "SubclassInit")) = 1 Then
    #If VBA7 Then
    RemoveWindowSubclass hWnd, FlexSubclassProcPtr, GetProp(hWnd, StrPtr(Name & "SubclassID"))
    #Else
    If FlexSubclassW2K > -1 Then
        RemoveWindowSubclass hWnd, FlexSubclassProcPtr, GetProp(hWnd, StrPtr(Name & "SubclassID"))
    Else
        RemoveWindowSubclassW2K hWnd, FlexSubclassProcPtr, GetProp(hWnd, StrPtr(Name & "SubclassID"))
    End If
    #End If
    RemoveProp hWnd, StrPtr(Name & "SubclassID")
    RemoveProp hWnd, StrPtr(Name & "SubclassInit")
End If
End Sub

#If VBA7 Then
Public Function FlexSubclassProc(ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As LongPtr) As LongPtr
#Else
Public Function FlexSubclassProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
#End If
Select Case wMsg
    Case WM_DESTROY
        FlexSubclassProc = FlexDefaultProc(hWnd, wMsg, wParam, lParam)
        Exit Function
    Case WM_NCDESTROY, WM_UAHDESTROYWINDOW
        FlexSubclassProc = FlexDefaultProc(hWnd, wMsg, wParam, lParam)
        #If VBA7 Then
        RemoveWindowSubclass hWnd, FlexSubclassProcPtr, uIdSubclass
        #Else
        If FlexSubclassW2K > -1 Then
            RemoveWindowSubclass hWnd, FlexSubclassProcPtr, uIdSubclass
        Else
            RemoveWindowSubclassW2K hWnd, FlexSubclassProcPtr, uIdSubclass
        End If
        #End If
        Exit Function
End Select
On Error Resume Next
Dim This As VBFlexGrid
FlexObjSetAddRef This, uIdSubclass
If Err.Number = 0 Then
    FlexSubclassProc = This.FSubclass_Message(hWnd, wMsg, wParam, lParam, dwRefData)
Else
    FlexSubclassProc = FlexDefaultProc(hWnd, wMsg, wParam, lParam)
End If
End Function

Public Sub FlexWndRegisterClass()
If FlexClassAtom = 0 And FlexRefCount = 0 Then
    Dim WCEX As WNDCLASSEX, ClassName As String
    ClassName = "VBFlexGridWndClass"
    With WCEX
    .cbSize = LenB(WCEX)
    ' CS_VREDRAW and CS_HREDRAW will not be specified as entire redraw upon resize is not necessary.
    .dwStyle = CS_DBLCLKS
    .lpfnWndProc = ProcPtr(AddressOf FlexWindowProc)
    .cbWndExtra = PTR_SIZE
    .hInstance = App.hInstance
    .hCursor = LoadCursor(NULL_PTR, IDC_ARROW)
    .hbrBackground = NULL_PTR
    .lpszClassName = StrPtr(ClassName)
    End With
    FlexClassAtom = RegisterClassEx(WCEX)
    If FlexSplitterBrush = NULL_PTR Then
        Dim Bmp As BITMAP, Pattern(0 To 3) As Long, hBmp As LongPtr
        Pattern(0) = &HAAAA5555
        Pattern(1) = &HAAAA5555
        Pattern(2) = &HAAAA5555
        Pattern(3) = &HAAAA5555
        With Bmp
        .BMType = 0
        .BMWidth = 16
        .BMHeight = 8
        .BMWidthBytes = 2
        .BMPlanes = 1
        .BMBitsPixel = 1
        .BMBits = VarPtr(Pattern(0))
        End With
        hBmp = CreateBitmapIndirect(Bmp)
        If hBmp <> NULL_PTR Then
            FlexSplitterBrush = CreatePatternBrush(hBmp)
            DeleteObject hBmp
        End If
    End If
End If
FlexRefCount = FlexRefCount + 1
End Sub

Public Sub FlexWndReleaseClass()
FlexRefCount = FlexRefCount - 1
If FlexClassAtom <> 0 And FlexRefCount = 0 Then
    UnregisterClass MakeDWord(FlexClassAtom, 0), App.hInstance
    FlexClassAtom = 0
    If FlexSplitterBrush <> NULL_PTR Then
        DeleteObject FlexSplitterBrush
        FlexSplitterBrush = NULL_PTR
    End If
End If
End Sub

Public Sub FlexComboCalendarRegisterClass()
If FlexComboCalendarClassAtom = 0 And FlexComboCalendarRefCount = 0 Then
    Dim WCEX As WNDCLASSEX, ClassName As String
    GetClassInfoEx App.hInstance, StrPtr("SysMonthCal32"), WCEX
    ClassName = "VBFlexGridComboCalendarClass"
    With WCEX
    .cbSize = LenB(WCEX)
    If Not (.dwStyle And CS_DROPSHADOW) = CS_DROPSHADOW Then .dwStyle = .dwStyle Or CS_DROPSHADOW
    .hInstance = App.hInstance
    .lpszClassName = StrPtr(ClassName)
    End With
    FlexComboCalendarClassAtom = RegisterClassEx(WCEX)
End If
FlexComboCalendarRefCount = FlexComboCalendarRefCount + 1
End Sub

Public Sub FlexComboCalendarReleaseClass()
FlexComboCalendarRefCount = FlexComboCalendarRefCount - 1
If FlexComboCalendarClassAtom <> 0 And FlexComboCalendarRefCount = 0 Then
    UnregisterClass MakeDWord(FlexComboCalendarClassAtom, 0), App.hInstance
    FlexComboCalendarClassAtom = 0
End If
End Sub

#If VBA7 Then
Public Function FlexGetSplitterBrush() As LongPtr
#Else
Public Function FlexGetSplitterBrush() As Long
#End If
FlexGetSplitterBrush = FlexSplitterBrush
End Function

#If VBA7 Then
Public Function FlexWindowProc(ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
Public Function FlexWindowProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If
Dim lCustData As LongPtr
Select Case wMsg
    Case WM_CREATE
        CopyMemory ByVal VarPtr(lCustData), ByVal lParam, PTR_SIZE
        SetWindowLongPtr hWnd, 0, lCustData
        FlexWindowProc = DefWindowProc(hWnd, wMsg, wParam, lParam)
        Exit Function
    Case WM_DESTROY, WM_NCDESTROY, WM_UAHDESTROYWINDOW
        FlexWindowProc = DefWindowProc(hWnd, wMsg, wParam, lParam)
        Exit Function
End Select
lCustData = GetWindowLongPtr(hWnd, 0)
If lCustData <> NULL_PTR Then
    On Error Resume Next
    Dim This As VBFlexGrid
    FlexObjSetAddRef This, lCustData
    If Err.Number = 0 Then
        FlexWindowProc = This.FSubclass_Message(hWnd, wMsg, wParam, lParam, 1)
    Else
        FlexWindowProc = DefWindowProc(hWnd, wMsg, wParam, lParam)
    End If
Else
    FlexWindowProc = DefWindowProc(hWnd, wMsg, wParam, lParam)
End If
End Function

#If VBA7 Then
Public Sub FlexDoReaderMode(ByVal hWnd As LongPtr, ByVal wParam As LongPtr, ByVal lParam As LongPtr)
#Else
Public Sub FlexDoReaderMode(ByVal hWnd As Long, ByVal wParam As Long, ByVal lParam As Long)
#End If
If hWnd = NULL_PTR Or GetShiftStateFromParam(wParam) <> 0 Then Exit Sub
Const WS_HSCROLL As Long = &H100000, WS_VSCROLL As Long = &H200000
Dim dwStyle As Long
dwStyle = GetWindowLong(hWnd, GWL_STYLE)
If Not (dwStyle And WS_HSCROLL) = WS_HSCROLL And Not (dwStyle And WS_VSCROLL) = WS_VSCROLL Then Exit Sub
Dim X As Long, Y As Long, RC As RECT
X = Get_X_lParam(lParam)
Y = Get_Y_lParam(lParam)
RC.Left = X - 8
RC.Top = Y - 8
RC.Right = X + 8
RC.Bottom = Y + 8
Dim RMI As READERMODEINFO
RMI.cbSize = LenB(RMI)
RMI.hWnd = hWnd
RMI.dwFlags = 0
If Not (dwStyle And WS_HSCROLL) = WS_HSCROLL And (dwStyle And WS_VSCROLL) = WS_VSCROLL Then
    RMI.dwFlags = RMF_VERTICALONLY
ElseIf (dwStyle And WS_HSCROLL) = WS_HSCROLL And Not (dwStyle And WS_VSCROLL) = WS_VSCROLL Then
    RMI.dwFlags = RMF_HORIZONTALONLY
End If
RMI.lpRC = VarPtr(RC)
RMI.lpfnScroll = ProcPtr(AddressOf FlexReaderModeScroll)
RMI.lpfnDispatch = ProcPtr(AddressOf FlexReaderModeDispatch)
RMI.lParam = GetWindowLongPtr(hWnd, 0)
FlexReaderModeScrolled = False
FlexReaderModeCursorInitialized = False
' Ensure that the cursor will be set immediately.
Const WM_MOUSEMOVE As Long = &H200
PostMessage hWnd, WM_MOUSEMOVE, wParam, ByVal lParam
DoReaderMode RMI
If FlexReaderModeAnchorHandle <> NULL_PTR Then
    Dim hIcon As LongPtr
    hIcon = GetWindowLongPtr(FlexReaderModeAnchorHandle, 0)
    DestroyWindow FlexReaderModeAnchorHandle
    FlexReaderModeAnchorHandle = NULL_PTR
    If hIcon <> NULL_PTR Then DestroyIcon hIcon
End If
End Sub

Private Function FlexReaderModeScroll(ByVal lpRMI As LongPtr, ByVal DX As Long, ByVal DY As Long) As Long
If lpRMI = NULL_PTR Then
    FlexReaderModeScroll = 0
    Exit Function
End If
If DX <> 0 Or DY <> 0 Then FlexReaderModeScrolled = True
Dim RMI As READERMODEINFO
CopyMemory RMI, ByVal lpRMI, LenB(RMI)
If RMI.lParam <> NULL_PTR Then
    On Error Resume Next
    Dim This As VBFlexGrid
    FlexObjSetAddRef This, RMI.lParam
    If Err.Number = 0 Then This.FReaderModeScroll DX, DY
End If
FlexReaderModeScroll = 1
End Function

Private Function FlexReaderModeDispatch(ByVal lpMsg As LongPtr) As Long
Dim Msg As TMSG
CopyMemory Msg, ByVal lpMsg, LenB(Msg)
Const WM_MOUSEMOVE As Long = &H200, WM_MBUTTONUP As Long = &H208, WM_MOUSEWHEEL As Long = &H20A, WM_MOUSEHWHEEL As Long = &H20E
Select Case Msg.Message
    Case WM_MOUSEMOVE
        If FlexReaderModeCursorInitialized = False Then
            PostMessage Msg.hWnd, WM_MOUSEMOVE, Msg.wParam, ByVal Msg.lParam
            FlexReaderModeCursorInitialized = True
        ElseIf FlexReaderModeAnchorHandle = NULL_PTR Then
            Call FlexReaderModeCreateAnchor
        End If
    Case WM_MBUTTONUP
        ' ReaderMode will be finished at default handler.
        If FlexReaderModeScrolled = False Then
            FlexReaderModeDispatch = 1
            Exit Function
        End If
    Case WM_MOUSEWHEEL, WM_MOUSEHWHEEL
        FlexReaderModeDispatch = 1
        Exit Function
End Select
FlexReaderModeDispatch = 0
End Function

Private Sub FlexReaderModeCreateAnchor()
If FlexReaderModeAnchorHandle <> NULL_PTR Then Exit Sub
Dim CI As CURSORINFO
CI.cbSize = LenB(CI)
GetCursorInfo CI
If CI.hCursor = NULL_PTR Then Exit Sub
Dim pIconInfo As ICONINFO, Bmp As BITMAP, CX As Long, CY As Long
GetIconInfo CI.hCursor, pIconInfo
GetObjectAPI pIconInfo.hBMMask, LenB(Bmp), Bmp
CX = Bmp.BMWidth
If pIconInfo.hBMColor <> NULL_PTR Then
    CY = Bmp.BMHeight
Else
    CY = Bmp.BMHeight / 2
End If
If pIconInfo.hBMColor <> NULL_PTR Then DeleteObject pIconInfo.hBMColor
If pIconInfo.hBMMask <> NULL_PTR Then DeleteObject pIconInfo.hBMMask
Const WS_POPUP As Long = &H80000000
Const WS_EX_TOOLWINDOW As Long = &H80, WS_EX_TOPMOST As Long = &H8, WS_EX_TRANSPARENT As Long = &H20, WS_EX_LAYERED As Long = &H80000
Dim dwStyle As Long, dwExStyle As Long
dwStyle = WS_POPUP
dwExStyle = WS_EX_TOOLWINDOW Or WS_EX_TOPMOST Or WS_EX_TRANSPARENT Or WS_EX_LAYERED
FlexReaderModeAnchorHandle = CreateWindowEx(dwExStyle, StrPtr("VBFlexGridReaderModeAnchorClass"), NULL_PTR, dwStyle, CI.PTScreenPos.X - pIconInfo.XHotspot, CI.PTScreenPos.Y - pIconInfo.YHotspot, CX, CY, NULL_PTR, NULL_PTR, App.hInstance, ByVal NULL_PTR)
If FlexReaderModeAnchorHandle <> NULL_PTR Then
    SetWindowLongPtr FlexReaderModeAnchorHandle, 0, CopyIcon(CI.hCursor)
    Const LWA_COLORKEY As Long = &H1, LWA_ALPHA As Long = &H2
    SetLayeredWindowAttributes FlexReaderModeAnchorHandle, vbWhite, 128, LWA_COLORKEY Or LWA_ALPHA
    Const SW_SHOWNA As Long = 8
    ShowWindow FlexReaderModeAnchorHandle, SW_SHOWNA
End If
End Sub

Public Sub FlexReaderModeAnchorRegisterClass()
If FlexReaderModeAnchorClassAtom = 0 And FlexReaderModeAnchorRefCount = 0 Then
    Dim WCEX As WNDCLASSEX, ClassName As String
    ClassName = "VBFlexGridReaderModeAnchorClass"
    With WCEX
    .cbSize = LenB(WCEX)
    ' CS_VREDRAW and CS_HREDRAW will not be specified as entire redraw upon resize is not necessary.
    .dwStyle = 0
    .lpfnWndProc = ProcPtr(AddressOf FlexReaderModeAnchorWindowProc)
    .cbWndExtra = PTR_SIZE
    .hInstance = App.hInstance
    .hCursor = LoadCursor(NULL_PTR, IDC_ARROW)
    .hbrBackground = GetStockObject(WHITE_BRUSH)
    .lpszClassName = StrPtr(ClassName)
    End With
    FlexReaderModeAnchorClassAtom = RegisterClassEx(WCEX)
End If
FlexReaderModeAnchorRefCount = FlexReaderModeAnchorRefCount + 1
End Sub

Public Sub FlexReaderModeAnchorReleaseClass()
FlexReaderModeAnchorRefCount = FlexReaderModeAnchorRefCount - 1
If FlexReaderModeAnchorClassAtom <> 0 And FlexReaderModeAnchorRefCount = 0 Then
    UnregisterClass MakeDWord(FlexReaderModeAnchorClassAtom, 0), App.hInstance
    FlexReaderModeAnchorClassAtom = 0
End If
End Sub

Private Function FlexReaderModeAnchorWindowProc(ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Const WM_ERASEBKGND As Long = &H14
Select Case wMsg
    Case WM_ERASEBKGND
        Dim hIcon As LongPtr
        hIcon = GetWindowLongPtr(hWnd, 0)
        If hIcon <> NULL_PTR Then
            Dim RC As RECT
            GetClientRect hWnd, RC
            Const DI_NORMAL As Long = &H3
            DrawIconEx wParam, RC.Left, RC.Top, hIcon, RC.Right - RC.Left, RC.Bottom - RC.Top, 0, GetStockObject(WHITE_BRUSH), DI_NORMAL
            FlexReaderModeAnchorWindowProc = 1
            Exit Function
        End If
End Select
FlexReaderModeAnchorWindowProc = DefWindowProc(hWnd, wMsg, wParam, lParam)
End Function

#If ImplementPreTranslateMsg = True Then

#If VBA7 Then
Public Sub FlexPreTranslateMsgAddHook(ByVal hWnd As LongPtr)
#Else
Public Sub FlexPreTranslateMsgAddHook(ByVal hWnd As Long)
#End If
If FlexPreTranslateMsgHookHandle = NULL_PTR And FlexPreTranslateMsgCount = 0 Then
    Const WH_GETMESSAGE As Long = 3
    FlexPreTranslateMsgHookHandle = SetWindowsHookEx(WH_GETMESSAGE, AddressOf FlexPreTranslateMsgHookProc, NULL_PTR, App.ThreadID)
    ReDim FlexPreTranslateMsgHwnd(0) ' As LongPtr
    FlexPreTranslateMsgHwnd(0) = hWnd
Else
    ReDim Preserve FlexPreTranslateMsgHwnd(0 To FlexPreTranslateMsgCount) ' As LongPtr
    FlexPreTranslateMsgHwnd(FlexPreTranslateMsgCount) = hWnd
End If
FlexPreTranslateMsgCount = FlexPreTranslateMsgCount + 1
End Sub

#If VBA7 Then
Public Sub FlexPreTranslateMsgReleaseHook(ByVal hWnd As LongPtr)
#Else
Public Sub FlexPreTranslateMsgReleaseHook(ByVal hWnd As Long)
#End If
FlexPreTranslateMsgCount = FlexPreTranslateMsgCount - 1
If FlexPreTranslateMsgHookHandle <> NULL_PTR And FlexPreTranslateMsgCount = 0 Then
    UnhookWindowsHookEx FlexPreTranslateMsgHookHandle
    FlexPreTranslateMsgHookHandle = NULL_PTR
    Erase FlexPreTranslateMsgHwnd()
Else
    If FlexPreTranslateMsgCount > 0 Then
        Dim i As Long
        For i = 0 To FlexPreTranslateMsgCount
            If FlexPreTranslateMsgHwnd(i) = hWnd And i < FlexPreTranslateMsgCount Then
                FlexPreTranslateMsgHwnd(i) = FlexPreTranslateMsgHwnd(i + 1)
            End If
        Next i
        ReDim Preserve FlexPreTranslateMsgHwnd(0 To FlexPreTranslateMsgCount - 1) ' As LongPtr
    End If
End If
End Sub

Private Function FlexPreTranslateMsgHookProc(ByVal nCode As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Const HC_ACTION As Long = 0, PM_REMOVE As Long = &H1
Const WM_KEYFIRST As Long = &H100, WM_KEYLAST As Long = &H108, WM_NULL As Long = &H0
If nCode >= HC_ACTION And wParam = PM_REMOVE Then
    Dim Msg As TMSG
    CopyMemory Msg, ByVal lParam, LenB(Msg)
    If Msg.Message >= WM_KEYFIRST And Msg.Message <= WM_KEYLAST Then
        If FlexPreTranslateMsgCount > 0 Then
            Dim i As Long
            For i = 0 To FlexPreTranslateMsgCount - 1
                If Msg.hWnd = FlexPreTranslateMsgHwnd(i) Then
                    If SendMessage(Msg.hWnd, UM_PRETRANSLATEMSG, 0, ByVal lParam) <> 0 Then
                        Msg.Message = WM_NULL
                        Msg.wParam = 0
                        Msg.lParam = 0
                        CopyMemory ByVal lParam, Msg, LenB(Msg)
                        Exit For
                    End If
                End If
            Next i
        End If
    End If
End If
FlexPreTranslateMsgHookProc = CallNextHookEx(FlexPreTranslateMsgHookHandle, nCode, wParam, lParam)
End Function

#End If
