Attribute VB_Name = "Module1"
Global fso As New CFileSystem2
Global dlg As New clsCmnDlg
Global reg As New clsRegistry2
Global killbitted As New Collection

Sub LoadKillBittedControlList()
    Dim tmp() As String
    
    reg.hive = HKEY_LOCAL_MACHINE
    Const base = "\SOFTWARE\Microsoft\Internet Explorer\ActiveX Compatibility"
    tmp() = reg.EnumKeys(base)
    
    For Each t In tmp
        v = reg.ReadValue(base & "\" & t, "Compatibility Flags")
        If v = &H400 Then killbitted.Add t, t
    Next
    
End Sub

Function GetProgID(GUID As String) As String
    Dim tmp As String
    Dim f As String
    
    reg.hive = HKEY_CLASSES_ROOT
    If Len(GUID) = 0 Then Exit Function
    
    f = "\CLSID\" & f
        
    If reg.keyExists(f) Then
        f = f & "\ProgID"
        If reg.keyExists(f) Then
            f = reg.ReadValue(f, "")
            GetProgID = f
            Exit Function
        End If
    End If
    
    tmp = Split(GUID, "-")(0)
    GetProgID = Right(tmp, Len(tmp) - 1)

End Function

Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Long
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub



Sub glue(ary, value) 'this modifies parent ary object
    On Error GoTo hell
    ary(UBound(ary)) = ary(UBound(ary)) & " " & value
Exit Sub
hell: push ary, value
      Stop
End Sub



Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    Dim i As Long
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function

Function IsIde() As Boolean
    On Error GoTo hell
    Debug.Print 1 / 0
    IsIde = False
    Exit Function
hell: IsIde = True
End Function

Sub dbg(msg)
    'debug.Print msg
End Sub

Function StripQuotes(ByVal x)
    x = Replace(x, "'", Empty)
    StripQuotes = Replace(x, """", Empty)
End Function

Function ExpandPath(ByVal fPath As String) As String
    Dim x As Long
    Dim tmp As String
    
    On Error Resume Next
    
    fPath = StripQuotes(fPath)
    x = InStrRev(fPath, "%")
    If x > 0 Then
        env = Mid(fPath, 1, x)
        fPath = Replace(fPath, env, Environ(Replace(env, "%", "")))
    End If
        
    If InStr(LCase(fPath), ":\") < 1 Then
        tmp = Environ("WinDIR") & "\" & fPath
        If fso.FileExists(tmp) Then
            fPath = tmp
        Else
            tmp = Environ("WinDIR") & "\System32\" & fPath
            If fso.FileExists(tmp) Then fPath = tmp
        End If
    End If
    
    ExpandPath = fPath
    
End Function

Function KeyExistsInCollection(c As Collection, val As String) As Boolean
    On Error GoTo nope
    Dim t
    If IsObject(c(val)) Then
        Set t = c(val)
    Else
        t = c(val)
    End If
    KeyExistsInCollection = True
 Exit Function
nope: KeyExistsInCollection = False
End Function

Function LikeAnyOfThese(ByVal sIn, ByVal sCmp) As Boolean
    Dim tmp() As String, i As Integer
    On Error GoTo hell
    sIn = LCase(sIn)
    sCmp = LCase(sCmp)
    tmp() = Split(sCmp, ",")
    For i = 0 To UBound(tmp)
        tmp(i) = "*" & Trim(tmp(i)) & "*"
        If Len(tmp(i)) > 0 And sIn Like tmp(i) Then
            LikeAnyOfThese = True
            Exit Function
        End If
    Next
hell:
End Function

Function AnyOfTheseInstr(sIn, sCmp) As Boolean
    Dim tmp() As String, i As Integer
    On Error GoTo hell
    tmp() = Split(sCmp, ",")
    For i = 0 To UBound(tmp)
        tmp(i) = Trim(tmp(i))
        If Len(tmp(i)) > 0 And InStr(1, sIn, tmp(i), vbTextCompare) > 0 Then
            AnyOfTheseInstr = True
            Exit Function
        End If
    Next
hell:
End Function

