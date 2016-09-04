Attribute VB_Name = "mPEiDPlugin"
'David Zimmer <dzzie@yahoo.com>
'http://sandsprite.com
Public ccWindows As Collection

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Sub DebugBreak Lib "kernel32" ()
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CallAsm Lib "user32" Alias "CallWindowProcA" (ByRef lpBytes As Any, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CallAsmAddr Lib "user32" Alias "CallWindowProcA" (ByVal lpCode As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)

Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

Function ChildWindows(Optional hwnd As Long = 0) As Collection 'of CWindow
    
    Set ccWindows = New Collection
    X = EnumChildWindows(0, AddressOf mPEiDPlugin.EnumChildProc, ByVal 0&)
    Set ChildWindows = ccWindows

End Function

Public Function EnumChildProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
    Dim c As New Cwindow
    c.hwnd = hwnd
    If Not IsObject(ccWindows) Then Set ccWindows = New Collection
    ccWindows.Add c 'module level collection object...
    EnumChildProc = 1  'continue enum
End Function

Function ColToStr(c As Collection) As String
    Dim tmp() As String
    
    For Each X In c
        push tmp, X
    Next
    
    ColToStr = Join(tmp, vbCrLf)
End Function

Function LaunchPeidPlugin(dll As String, fPath As String)
    
    Dim h As Long
    Dim lpProc As Long
    Dim lib As String
    
    lib = dll
    If Not FileExists(lib) Then lib = App.path & "\" & dll
    If Not FileExists(lib) Then lib = App.path & "\..\" & dll
    
    If Not FileExists(lib) Then
         MsgBox "PEID plugin Dll not found: " & lib, vbInformation
        Exit Function
    End If
    
    h = LoadLibrary(lib)
    
    If h = 0 Then
        MsgBox "LoadLibrary failed: " & lib, vbInformation
        Exit Function
    End If
    
    lpProc = GetProcAddress(h, "DoMyJob")
    
    If lpProc = 0 Then
        MsgBox "GetProcAddress(DoMyJob) failed: " & lib, vbInformation
        Exit Function
    End If
    
    Dim path() As Byte
    path = StrConv(fPath & Chr(0), vbFromUnicode)
    
    Const peid = &H50456944
    CallCdecl lpProc, 0, VarPtr(path(0)), peid, 0
    'DWORD DoMyJob(HWND hMainDlg, char *szFname, DWORD lpReserved, LPVOID lpParam)
    
End Function

'should be dep safe..
Function CallCdecl(lpfn As Long, ParamArray args()) As Long

    Dim asm() As String
    Dim stub() As Byte
    Dim i As Long
    Dim argSize As Byte
    Dim ret As Long
    Const PAGE_RWX      As Long = &H40
    Const MEM_COMMIT    As Long = &H1000
    Dim asmAddr As Long
    Dim sz As Long
    
    Const depSafe = True
    
    If lpfn = 0 Then Exit Function
    
    'push asm(), "CC"  'enable this to debug asm
    
    'we step through args backwards to preserve intutive ordering
    For i = UBound(args) To 0 Step -1
        If Not IsNumeric(args(i)) Then
            MsgBox "CallCdecl Invalid Parameter #" & i & " TypeName=" & TypeName(args(i))
            Exit Function
        End If
        push asm(), "68 " & lng2Hex(CLng(args(i)))  '68 90807000    PUSH 708090
        argSize = argSize + 4
    Next

    push asm(), "B8 " & lng2Hex(lpfn)        'B8 90807000    MOV EAX,708090
    push asm(), "FF D0"                      'FFD0           CALL EAX
    push asm(), "83 C4 " & Hex(argSize)      '83 C4 XX       add esp, XX     'cleanup args
    push asm(), "C2 10 00"                   'C2 10 00       retn 10h        'cleanup our callwindowproc args
    
    stub() = toBytes(Join(asm, " "))
    
    If Not depSafe Then
        CallCdecl = CallAsm(stub(0), 0, 0, 0, 0)
        Exit Function
    End If
    
    sz = UBound(stub) + 1
    asmAddr = VirtualAlloc(ByVal 0&, sz, MEM_COMMIT, PAGE_RWX)
    
    If asmAddr = 0 Then
        MsgBox "Failed to allocate RWE memory size: " & sz, vbInformation
        Exit Function
    End If
    
    RtlMoveMemory asmAddr, VarPtr(stub(0)), sz
    CallCdecl = CallAsmAddr(asmAddr, 0, 0, 0, 0)
    VirtualFree asmAddr, sz, 0
    
    
End Function

'endian swap and return spaced out hex string
Private Function lng2Hex(X As Long) As String
    Dim b(1 To 4) As Byte
    CopyMemory b(1), X, 4
    lng2Hex = Hex(b(1)) & " " & Hex(b(2)) & " " & Hex(b(3)) & " " & Hex(b(4))
End Function

Private Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim X As Long
    X = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub

'not as efficient I know but cleaner to look at in the code
'and easier to view/debug/change as a single string..tradeoffs
Private Function toBytes(X As String) As Byte()
    Dim tmp() As String
    Dim fx() As Byte
    Dim i As Long
    
    tmp = Split(X, " ")
    ReDim fx(UBound(tmp))
    
    For i = 0 To UBound(tmp)
        fx(i) = CInt("&h" & tmp(i))
    Next
    
    toBytes = fx()

End Function

Private Function FileExists(path As String) As Boolean
  On Error GoTo hell
  Dim tmp As String
  tmp = Replace(path, "'", Empty)
  tmp = Replace(tmp, """", Empty)
  If Len(tmp) = 0 Then Exit Function
  If Dir(tmp, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
  Exit Function
hell: FileExists = False
End Function


