Attribute VB_Name = "VBFlexGridBase"
Option Explicit

#Const ImplementIDEStopProtection = True

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
Private Declare Function GetAncestor Lib "user32" (ByVal hWnd As Long, ByVal gaFlags As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameW" (ByVal hWnd As Long, ByVal lpClassName As Long, ByVal nMaxCount As Long) As Long
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

#If ImplementIDEStopProtection = True Then

Private Declare Function VirtualAlloc Lib "kernel32" (ByRef lpAddress As Long, ByVal dwSize As Long, ByVal flAllocType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, ByRef lpflOldProtect As Long) As Long
Private Const MEM_COMMIT As Long = &H1000
Private Const PAGE_EXECUTE_READWRITE As Long = &H40
Private Type IMAGE_DATA_DIRECTORY
VirtualAddress As Long
Size As Long
End Type
Private Type IMAGE_OPTIONAL_HEADER32
Magic As Integer
MajorLinkerVersion As Byte
MinorLinkerVersion As Byte
SizeOfCode As Long
SizeOfInitalizedData As Long
SizeOfUninitalizedData As Long
AddressOfEntryPoint As Long
BaseOfCode As Long
BaseOfData As Long
ImageBase As Long
SectionAlignment As Long
FileAlignment As Long
MajorOperatingSystemVer As Integer
MinorOperatingSystemVer As Integer
MajorImageVersion As Integer
MinorImageVersion As Integer
MajorSubsystemVersion As Integer
MinorSubsystemVersion As Integer
Reserved1 As Long
SizeOfImage As Long
SizeOfHeaders As Long
CheckSum As Long
Subsystem As Integer
DllCharacteristics As Integer
SizeOfStackReserve As Long
SizeOfStackCommit As Long
SizeOfHeapReserve As Long
SizeOfHeapCommit As Long
LoaderFlags As Long
NumberOfRvaAndSizes As Long
DataDirectory(15) As IMAGE_DATA_DIRECTORY
End Type
Private Type IMAGE_DOS_HEADER
e_magic As Integer
e_cblp As Integer
e_cp As Integer
e_crlc As Integer
e_cparhdr As Integer
e_minalloc As Integer
e_maxalloc As Integer
e_ss As Integer
e_sp As Integer
e_csum As Integer
e_ip As Integer
e_cs As Integer
e_lfarlc As Integer
e_onvo As Integer
e_res(0 To 3) As Integer
e_oemid As Integer
e_oeminfo As Integer
e_res2(0 To 9) As Integer
e_lfanew As Long
End Type

#End If

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

Public Function FlexRootIsEditor(ByVal hWnd As Long) As Boolean
Static Done As Boolean, Value As Boolean
If Done = False Then
    Const GA_ROOT As Long = 2
    hWnd = GetAncestor(hWnd, GA_ROOT)
    If hWnd <> 0 Then
        Dim Buffer As String, RetVal As Long
        Buffer = String(256, vbNullChar)
        RetVal = GetClassName(hWnd, StrPtr(Buffer), Len(Buffer))
        If RetVal <> 0 Then Value = CBool(Left$(Buffer, RetVal) = "wndclass_desked_gsk")
    End If
    Done = True
End If
FlexRootIsEditor = Value
End Function

Public Sub FlexTopParentValidateControls(ByVal UserControl As Object)
With GetTopUserControl(UserControl)
If TypeOf .Parent Is VB.Form Then
    Dim Form As VB.Form
    Set Form = .Parent
    Form.ValidateControls
Else
    Const IID_IPropertyPage As String = "{B196B28D-BAB4-101A-B69C-00AA00341D07}"
    If VTableInterfaceSupported(.Parent, IID_IPropertyPage) = True Then
        Dim PropertyPage As VB.PropertyPage, TempPropertyPage As VB.PropertyPage
        CopyMemory TempPropertyPage, ObjPtr(.Parent), 4
        Set PropertyPage = TempPropertyPage
        CopyMemory TempPropertyPage, 0&, 4
        PropertyPage.ValidateControls
    End If
End If
End With
End Sub

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

Public Sub FlexInitIDEStopProtection()

#If ImplementIDEStopProtection = True Then

If InIDE() = True Then
    Dim ASMWrapper As Long, RestorePointer As Long, OldAddress As Long
    ASMWrapper = VirtualAlloc(ByVal 0, 20, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    OldAddress = GetProcAddress(GetModuleHandle(StrPtr("vba6.dll")), "EbProjectReset")
    RestorePointer = HookIATEntry("vb6.exe", "vba6.dll", "EbProjectReset", ASMWrapper)
    WriteCall ASMWrapper, AddressOf FlexIDEStopProtectionHandler
    WriteByte ASMWrapper, &HC7 ' MOV
    WriteByte ASMWrapper, &H5
    WriteLong ASMWrapper, RestorePointer ' IAT Entry
    WriteLong ASMWrapper, OldAddress ' Address from EbProjectReset
    WriteJump ASMWrapper, OldAddress
End If

#End If

End Sub

#If ImplementIDEStopProtection = True Then

Private Sub FlexIDEStopProtectionHandler()
On Error Resume Next
Call RemoveAllVTableSubclass(VTableInterfaceInPlaceActiveObject)
Call RemoveAllVTableSubclass(VTableInterfaceControl)
Call RemoveAllVTableSubclass(VTableInterfacePerPropertyBrowsing)
Dim AppForm As Form, CurrControl As Control
For Each AppForm In Forms
    For Each CurrControl In AppForm.Controls
        If TypeOf CurrControl Is VBFlexGrid Then
            Call FlexRemoveSubclass(CurrControl.hWnd)
            Call FlexRemoveSubclass(CurrControl.hWndUserControl)
        End If
    Next CurrControl
Next AppForm
End Sub

Private Function HookIATEntry(ByVal Module As String, ByVal Lib As String, ByVal Fnc As String, ByVal NewAddr As Long) As Long
Dim hMod As Long, OldLibFncAddr As Long
Dim lpIAT As Long, IATLen As Long, IATPos As Long
Dim DOSHdr As IMAGE_DOS_HEADER
Dim PEHdr As IMAGE_OPTIONAL_HEADER32
hMod = GetModuleHandle(StrPtr(Module))
If hMod = 0 Then Exit Function
OldLibFncAddr = GetProcAddress(GetModuleHandle(StrPtr(Lib)), Fnc)
If OldLibFncAddr = 0 Then Exit Function
CopyMemory DOSHdr, ByVal hMod, LenB(DOSHdr)
CopyMemory PEHdr, ByVal UnsignedAdd(hMod, DOSHdr.e_lfanew), LenB(PEHdr)
Const IMAGE_NT_SIGNATURE As Long = &H4550
If PEHdr.Magic = IMAGE_NT_SIGNATURE Then
    lpIAT = UnsignedAdd(PEHdr.DataDirectory(15).VirtualAddress, hMod)
    IATLen = PEHdr.DataDirectory(15).Size
    IATPos = lpIAT
    Do Until CLngToULng(IATPos) >= CLngToULng(UnsignedAdd(lpIAT, IATLen))
        If DeRef(IATPos) = OldLibFncAddr Then
            VirtualProtect IATPos, 4, PAGE_EXECUTE_READWRITE, 0
            CopyMemory ByVal IATPos, NewAddr, 4
            HookIATEntry = IATPos
            Exit Do
        End If
        IATPos = UnsignedAdd(IATPos, 4)
    Loop
End If
End Function

Private Function DeRef(ByVal Addr As Long) As Long
CopyMemory DeRef, ByVal Addr, 4
End Function

Private Sub WriteJump(ByRef ASM As Long, ByRef Addr As Long)
WriteByte ASM, &HE9
WriteLong ASM, Addr - ASM - 4
End Sub

Private Sub WriteCall(ByRef ASM As Long, ByRef Addr As Long)
WriteByte ASM, &HE8
WriteLong ASM, Addr - ASM - 4
End Sub

Private Sub WriteLong(ByRef ASM As Long, ByRef Lng As Long)
CopyMemory ByVal ASM, Lng, 4
ASM = ASM + 4
End Sub

Private Sub WriteByte(ByRef ASM As Long, ByRef B As Byte)
CopyMemory ByVal ASM, B, 1
ASM = ASM + 1
End Sub

#End If
