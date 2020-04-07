Attribute VB_Name = "VBFlexGridBase"
Option Explicit
Private Type TINITCOMMONCONTROLSEX
dwSize As Long
dwICC As Long
End Type
Private Type DLLVERSIONINFO
cbSize As Long
dwMajor As Long
dwMinor As Long
dwBuildNumber As Long
dwPlatformID As Long
End Type
Private Type OSVERSIONINFO
dwOSVersionInfoSize As Long
dwMajorVersion As Long
dwMinorVersion As Long
dwBuildNumber As Long
dwPlatformID As Long
szCSDVersion(0 To ((128 * 2) - 1)) As Byte
End Type
Private Type WNDCLASSEX
cbSize As Long
dwStyle As Long
lpfnWndProc As Long
cbClsExtra As Long
cbWndExtra As Long
hInstance As Long
hIcon As Long
hCursor As Long
hbrBackground As Long
lpszMenuName As Long
lpszClassName As Long
hIconSm As Long
End Type
Private Type BITMAP
BMType As Long
BMWidth As Long
BMHeight As Long
BMWidthBytes As Long
BMPlanes As Integer
BMBitsPixel As Integer
BMBits As Long
End Type
Public Declare Function FlexPtrToShadowObj Lib "msvbvm60.dll" Alias "__vbaObjSetAddref" (ByRef Destination As Any, ByVal lpObject As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function InitCommonControlsEx Lib "comctl32" (ByRef ICCEX As TINITCOMMONCONTROLSEX) As Long
Private Declare Function RegisterClassEx Lib "user32" Alias "RegisterClassExW" (ByRef lpWndClassEx As WNDCLASSEX) As Integer
Private Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassW" (ByVal lpClassName As Long, ByVal hInstance As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function DllGetVersion Lib "comctl32" (ByRef pdvi As DLLVERSIONINFO) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExW" (ByRef lpVersionInfo As OSVERSIONINFO) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function CreateBitmapIndirect Lib "gdi32" (ByRef lpBitmap As BITMAP) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleW" (ByVal lpModuleName As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropW" (ByVal hWnd As Long, ByVal lpString As Long, ByVal hData As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropW" (ByVal hWnd As Long, ByVal lpString As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropW" (ByVal hWnd As Long, ByVal lpString As Long) As Long
Private Declare Function SetWindowSubclass Lib "comctl32" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
Private Declare Function RemoveWindowSubclass Lib "comctl32" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Long
Private Declare Function DefSubclassProc Lib "comctl32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowSubclass_W2K Lib "comctl32" Alias "#410" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
Private Declare Function RemoveWindowSubclass_W2K Lib "comctl32" Alias "#412" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Long
Private Declare Function DefSubclassProc_W2K Lib "comctl32" Alias "#413" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Private Const WM_CREATE As Long = &H1
Private Const WM_DESTROY As Long = &H2
Private Const WM_NCDESTROY As Long = &H82
Private Const WM_UAHDESTROYWINDOW As Long = &H90
Private Const CS_VREDRAW As Long = &H1, CS_HREDRAW As Long = &H2
Private Const CS_DBLCLKS As Long = &H8
Private Const IDC_ARROW As Long = 32512
Private ShellModHandle As Long, ShellModCount As Long
Private FlexClassAtom As Integer, FlexRefCount As Long
Private FlexSplitterBrush As Long

Public Sub FlexLoadShellMod()
If (ShellModHandle Or ShellModCount) = 0 Then ShellModHandle = LoadLibrary(StrPtr("Shell32.dll"))
ShellModCount = ShellModCount + 1
End Sub

Public Sub FlexReleaseShellMod()
ShellModCount = ShellModCount - 1
If ShellModCount = 0 And ShellModHandle <> 0 Then
    FreeLibrary ShellModHandle
    ShellModHandle = 0
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

Public Function FlexW2KCompatibility() As Boolean
Static Done As Boolean, Value As Boolean
If Done = False Then
    Dim Version As OSVERSIONINFO
    On Error Resume Next
    Version.dwOSVersionInfoSize = LenB(Version)
    If GetVersionEx(Version) <> 0 Then
        With Version
        Const VER_PLATFORM_WIN32_NT As Long = 2
        If .dwPlatformID = VER_PLATFORM_WIN32_NT Then
            If .dwMajorVersion = 5 And .dwMinorVersion = 0 Then Value = True
        End If
        End With
    End If
    Done = True
End If
FlexW2KCompatibility = Value
End Function

Public Sub FlexSetSubclass(ByVal hWnd As Long, ByVal This As VBFlexGrid, ByVal dwRefData As Long, Optional ByVal Name As String)
If hWnd = 0 Then Exit Sub
If Name = vbNullString Then Name = "VBFlexGrid"
If GetProp(hWnd, StrPtr(Name & "SubclassInit")) = 0 Then
    If FlexW2KCompatibility() = False Then
        SetWindowSubclass hWnd, AddressOf FlexSubclassProc, ObjPtr(This), dwRefData
    Else
        SetWindowSubclass_W2K hWnd, AddressOf FlexSubclassProc, ObjPtr(This), dwRefData
    End If
    SetProp hWnd, StrPtr(Name & "SubclassID"), ObjPtr(This)
    SetProp hWnd, StrPtr(Name & "SubclassInit"), 1
End If
End Sub

Public Function FlexDefaultProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
If FlexW2KCompatibility() = False Then
    FlexDefaultProc = DefSubclassProc(hWnd, wMsg, wParam, lParam)
Else
    FlexDefaultProc = DefSubclassProc_W2K(hWnd, wMsg, wParam, lParam)
End If
End Function

Public Sub FlexRemoveSubclass(ByVal hWnd As Long, Optional ByVal Name As String)
If hWnd = 0 Then Exit Sub
If Name = vbNullString Then Name = "VBFlexGrid"
If GetProp(hWnd, StrPtr(Name & "SubclassInit")) = 1 Then
    If FlexW2KCompatibility() = False Then
        RemoveWindowSubclass hWnd, AddressOf FlexSubclassProc, GetProp(hWnd, StrPtr(Name & "SubclassID"))
    Else
        RemoveWindowSubclass_W2K hWnd, AddressOf FlexSubclassProc, GetProp(hWnd, StrPtr(Name & "SubclassID"))
    End If
    RemoveProp hWnd, StrPtr(Name & "SubclassID")
    RemoveProp hWnd, StrPtr(Name & "SubclassInit")
End If
End Sub

Public Function FlexSubclassProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
Select Case wMsg
    Case WM_DESTROY
        FlexSubclassProc = FlexDefaultProc(hWnd, wMsg, wParam, lParam)
        Exit Function
    Case WM_NCDESTROY, WM_UAHDESTROYWINDOW
        FlexSubclassProc = FlexDefaultProc(hWnd, wMsg, wParam, lParam)
        If FlexW2KCompatibility() = False Then
            RemoveWindowSubclass hWnd, AddressOf VBFlexGridBase.FlexSubclassProc, uIdSubclass
        Else
            RemoveWindowSubclass_W2K hWnd, AddressOf VBFlexGridBase.FlexSubclassProc, uIdSubclass
        End If
        Exit Function
End Select
On Error Resume Next
Dim This As VBFlexGrid
FlexPtrToShadowObj This, uIdSubclass
If Err.Number = 0 Then
    FlexSubclassProc = This.FSubclass_Message(hWnd, wMsg, wParam, lParam, dwRefData)
Else
    FlexSubclassProc = FlexDefaultProc(hWnd, wMsg, wParam, lParam)
End If
End Function

Public Sub FlexWndRegisterClass()
If (FlexClassAtom Or FlexRefCount) = 0 Then
    Dim WCEX As WNDCLASSEX, ClassName As String
    ClassName = "VBFlexGridWndClass"
    With WCEX
    .cbSize = LenB(WCEX)
    .dwStyle = CS_VREDRAW Or CS_HREDRAW Or CS_DBLCLKS
    .lpfnWndProc = ProcPtr(AddressOf FlexWindowProc)
    .cbWndExtra = 4
    .hInstance = App.hInstance
    .hCursor = LoadCursor(0, IDC_ARROW)
    .hbrBackground = 0
    .lpszClassName = StrPtr(ClassName)
    End With
    FlexClassAtom = RegisterClassEx(WCEX)
    If FlexSplitterBrush = 0 Then
        Dim Bmp As BITMAP, Pattern(0 To 3) As Long, hBmp As Long
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
        If hBmp <> 0 Then
            FlexSplitterBrush = CreatePatternBrush(hBmp)
            DeleteObject hBmp
        End If
    End If
End If
FlexRefCount = FlexRefCount + 1
End Sub

Public Sub FlexWndReleaseClass()
FlexRefCount = FlexRefCount - 1
If FlexRefCount = 0 And FlexClassAtom <> 0 Then
    UnregisterClass MakeDWord(FlexClassAtom, 0), App.hInstance
    FlexClassAtom = 0
    If FlexSplitterBrush <> 0 Then
        DeleteObject FlexSplitterBrush
        FlexSplitterBrush = 0
    End If
End If
End Sub

Public Function FlexGetSplitterBrush() As Long
FlexGetSplitterBrush = FlexSplitterBrush
End Function

Public Function FlexWindowProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim lCustData As Long
Select Case wMsg
    Case WM_CREATE
        CopyMemory ByVal VarPtr(lCustData), ByVal lParam, 4
        SetWindowLong hWnd, 0, lCustData
        FlexWindowProc = DefWindowProc(hWnd, wMsg, wParam, lParam)
        Exit Function
    Case WM_DESTROY, WM_NCDESTROY, WM_UAHDESTROYWINDOW
        FlexWindowProc = DefWindowProc(hWnd, wMsg, wParam, lParam)
        Exit Function
End Select
lCustData = GetWindowLong(hWnd, 0)
If lCustData <> 0 Then
    On Error Resume Next
    Dim This As VBFlexGrid
    FlexPtrToShadowObj This, lCustData
    If Err.Number = 0 Then
        FlexWindowProc = This.FSubclass_Message(hWnd, wMsg, wParam, lParam, 1)
    Else
        FlexWindowProc = DefWindowProc(hWnd, wMsg, wParam, lParam)
    End If
Else
    FlexWindowProc = DefWindowProc(hWnd, wMsg, wParam, lParam)
End If
End Function
