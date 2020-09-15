VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStrings 
   Caption         =   "Strings"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   14130
   LinkTopic       =   "Form2"
   ScaleHeight     =   5340
   ScaleWidth      =   14130
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optVa 
      Caption         =   "va"
      Height          =   255
      Left            =   10800
      TabIndex        =   15
      Top             =   90
      Width           =   555
   End
   Begin VB.OptionButton optRaw 
      Caption         =   "raw"
      Height          =   225
      Left            =   10170
      TabIndex        =   14
      Top             =   90
      Value           =   -1  'True
      Width           =   585
   End
   Begin VB.Timer tmrReRun 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   1590
      Top             =   150
   End
   Begin VB.CheckBox chkFilter 
      Caption         =   "Filter Results"
      Height          =   285
      Left            =   11430
      TabIndex        =   11
      Top             =   30
      Width           =   255
   End
   Begin VB.CheckBox chkShowOffsets 
      Caption         =   "Offsets"
      Height          =   285
      Left            =   9300
      TabIndex        =   10
      Top             =   30
      Width           =   825
   End
   Begin VB.CommandButton cmdFindAll 
      Caption         =   "All"
      Height          =   315
      Left            =   4440
      TabIndex        =   9
      Top             =   0
      Width           =   885
   End
   Begin VB.CommandButton cmdRescan 
      Caption         =   "Rescan"
      Height          =   315
      Left            =   7680
      TabIndex        =   8
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox txtMinLen 
      Height          =   285
      Left            =   7170
      TabIndex        =   7
      Text            =   "6"
      Top             =   0
      Width           =   465
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   225
      Left            =   60
      TabIndex        =   5
      Top             =   330
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save As"
      Height          =   315
      Left            =   5370
      TabIndex        =   4
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Find"
      Height          =   315
      Left            =   3480
      TabIndex        =   3
      Top             =   0
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   540
      TabIndex        =   2
      Top             =   0
      Width           =   2895
   End
   Begin RichTextLib.RichTextBox rtf 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   8281
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      TextRTF         =   $"frmPeek.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblMore 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "More"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   12915
      TabIndex        =   16
      Top             =   45
      Width           =   735
   End
   Begin VB.Label chkResetMin 
      Caption         =   "save min"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   8460
      TabIndex        =   13
      Top             =   30
      Width           =   675
   End
   Begin VB.Label Label3 
      Caption         =   "Filter Results"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   11700
      TabIndex        =   12
      Top             =   60
      Width           =   990
   End
   Begin VB.Label Label2 
      Caption         =   "Min Size"
      Height          =   255
      Left            =   6510
      TabIndex        =   6
      Top             =   30
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "Find"
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   435
   End
   Begin VB.Menu mnuMore 
      Caption         =   "mnuMore"
      Begin VB.Menu mnuDelphiFIlter 
         Caption         =   "Delphi Filter"
      End
      Begin VB.Menu mnuStringMatch 
         Caption         =   "Find String Matches"
      End
      Begin VB.Menu mnuStringDiff 
         Caption         =   "String Diff"
      End
      Begin VB.Menu mnuChangeFont 
         Caption         =   "Change Font"
      End
   End
End
Attribute VB_Name = "frmStrings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'License: Copyright (C) 2005 David Zimmer <david@idefense.com, dzzie@yahoo.com>
'
'         This program is free software; you can redistribute it and/or modify it
'         under the terms of the GNU General Public License as published by the Free
'         Software Foundation; either version 2 of the License, or (at your option)
'         any later version.
'
'         This program is distributed in the hope that it will be useful, but WITHOUT
'         ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or
'         FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for
'         more details.
'
'         You should have received a copy of the GNU General Public License along with
'         this program; if not, write to the Free Software Foundation, Inc., 59 Temple
'         Place, Suite 330, Boston, MA 02111-1307 USA

Option Explicit

Dim sSearch
Dim lastFind As Long
Dim lastSize As Long
Dim curFile As String
Dim pe As New CPEEditor

Dim d As New RegExp
Dim mc As MatchCollection
Dim m As match
Dim ret() As String
Dim lines As Long
Dim formLoaded As Boolean
Dim filtered() As String
Dim running As Boolean
Dim ranHidden As Boolean

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long


Option Compare Binary

Sub DisplayList(data As String)
    
    rtf.text = data
    Me.Show 1
    
End Sub


Private Sub chkEntropy_Click()
     If Not formLoaded Then Exit Sub
     ParseFile curFile, True
End Sub


Private Sub chkResetMin_Click()

    If Not IsNumeric(txtMinLen) Then
        MsgBox "Minimum String Length must be numeric", vbInformation
        Exit Sub
    End If
    
    On Error Resume Next
    minStrLen = CLng(txtMinLen)
    If Len(minStrLen) = 0 Then minStrLen = 4
    SaveMySetting "minStrLen", minStrLen
    
End Sub

Private Sub chkShowOffsets_Click()
    If Not formLoaded Then Exit Sub
    ParseFile curFile, True
End Sub

Private Sub cmdFindAll_Click()
    On Error Resume Next
    
    'pretty sure all these like operators hold for vb6 as well.. http://msdn.microsoft.com/en-us/library/8t3khw5f.aspx
    
    Dim tmp, x, ret(), i, f As String
     
    If Len(Text1) = 0 Then Exit Sub
    tmp = Split(rtf.text, vbCrLf)
    
    pb.value = 0
    For Each x In tmp
         i = i + 1
        If InStr(Text1, "*") > 0 Then
            If x Like Text1 Then
                push ret, x
            End If
        Else
            If InStr(1, x, Text1, vbTextCompare) > 0 Then
                push ret, x
            End If
        End If
        If i Mod 5 = 0 Then setpb i, UBound(tmp)
    Next
    pb.value = 0
    
    x = UBound(ret)
    If x < 0 Then
        Me.Caption = "No results found.."
        Exit Sub
    End If
    
    Dim data As String
    data = Join(ret, vbCrLf)
    
    If Len(data) = 0 Then
        Me.Caption = "Search for: " & Text1 & " 0 hits"
        Exit Sub
    Else
        Me.Caption = "Search for: " & Text1 & " " & UBound(ret) + 1 & " hits"
    End If
    
    f = fso.GetFreeFileName(Environ("temp"))
    fso.WriteFile f, data
    Shell "notepad.exe """ & f & """", vbNormalFocus
    
End Sub

Private Sub cmdRescan_Click()
    ParseFile curFile
End Sub

Private Sub Command1_Click()
        
    On Error Resume Next
    
    If sSearch <> Text1 Then
        sSearch = Text1
        lastFind = 0
        lastFind = rtf.find(sSearch)
        lastFind = lastFind + 1
        Me.Caption = "Search for: " & Text1 & " - " & occuranceCount(rtf.text, Text1) & " hits"
    Else
        lastFind = rtf.find(sSearch, lastFind)
        lastFind = lastFind + 1
    End If
    
    If lastFind > 0 Then
        rtf.SelStart = lastFind
        rtf.SelLength = Len(Text1)
    End If
    
End Sub

Public Sub AutoSave()
    Dim f As String
    Dim def As String
    Dim pf As String
    On Error Resume Next
    pf = fso.GetParentFolder(curFile)
    def = fso.GetBaseName(curFile)
    'If Len(def) > 12 Then def = VBA.Left(def, 8)
    f = pf & "\str_" & def & ".txt"
    fso.WriteFile f, rtf.text
End Sub

Private Sub Command3_Click()
    Dim f As String
    Dim def As String
    Dim pf As String
    On Error Resume Next
    pf = fso.GetParentFolder(curFile)
    def = fso.GetBaseName(curFile)
    If Len(def) > 12 Then def = VBA.Left(def, 5)
    def = "str_" & def & ".txt"
    f = dlg.SaveDialog(def, pf, "Save Report as")
    If Len(f) = 0 Then Exit Sub
    fso.WriteFile f, rtf.text
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Me.Icon = myIcon
    sSearch = -1
    txtMinLen = minStrLen 'global
    pb.max = 100
    pb.value = 0
    RestoreFormSizeAnPosition Me
    Me.Visible = True
    chkShowOffsets.value = GetMySetting("offsests", 1)
   ' mnuHiddenStrings.Checked = IIf(GetMySetting("hiddenstrings", 0) = 0, False, True)
    chkFilter.value = GetMySetting("Filter", 0)
    optRaw.value = IIf(GetMySetting("Raw", 1) = 1, True, False)
    If Not optRaw.value Then optVa.value = True
    rtf.Font.Name = GetMySetting("strings.font.name", "Courier New")
    rtf.Font.size = GetMySetting("strings.font.size", 11)
    rtf.Font.Bold = GetMySetting("strings.font.bold", False)
    mnuMore.Visible = False
    formLoaded = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   abort = True
   SaveFormSizeAnPosition Me
   SaveMySetting "offsests", chkShowOffsets.value
   SaveMySetting "Filter", chkFilter.value
   SaveMySetting "Raw", IIf(optRaw.value, 1, 0)
   If running Then End
   'SaveMySetting "hiddenstrings", IIf(mnuHiddenStrings.Checked, 1, 0)
   'End
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    rtf.Move 100, rtf.top, Me.Width - 400, Me.Height - rtf.top - 650
    pb.Width = rtf.Width
End Sub
 
 Sub setpb(cur, max)
    On Error Resume Next
    pb.value = (cur / max) * 100
    Me.Refresh
    DoEvents
 End Sub


Sub ParseFile(fPath As String, Optional force As Boolean = False)
    On Error GoTo hell
    
    Dim f As Long, pointer As Long
    Dim buf()  As Byte
    Dim x As Long
    Dim fs As Long
    Dim tmp As String
    
    If Not formLoaded Then Form_Load
        
    Erase ret
    Erase filtered
    
    If Is64BitProcessor() And InStr(1, fPath, Environ("windir") & "\System32", vbTextCompare) > 0 Then
        'file system redirection is stupid, pe.loadfile dll dependancies fail to load with redir off..tried everything..
        tmp = Environ("temp") & "\" & fso.FileNameFromPath(fPath)
        If fso.FileExists(tmp) Then fso.DeleteFile tmp
        fs = DisableRedir()
        FileCopy fPath, tmp
        RevertRedir
        fPath = tmp
    End If
    
    'MsgBox fPath & " " & fso.FileExists(fPath)
    curFile = fPath
    
    If Not IsNumeric(txtMinLen) Then txtMinLen = 4
    
    'If Not force Then If lastSize = txtMinLen Then Exit Sub
    lastSize = CLng(txtMinLen)
    
    'fs = DisableRedir()
    If Not fso.FileExists(fPath) Then
        MsgBox "File not found: " & fPath, vbExclamation
        GoTo done
    End If
    
    pe.LoadFile fPath
    
    If running Then
        abort = True
        tmrReRun.Enabled = True 'relaunch in 200ms
        'RevertRedir fs
        Exit Sub
    End If
    
    running = True
        
    'd.Pattern = "[a-z,A-Z,0-9 /?.\-_=+$\\@!*\(\)#]{4,}" 'ascii string search
    d.Pattern = "[\w0-9 /?.\-_=+$\\@!*\(\)#%~`\^&\|\{\}\[\]:;'""<>\,]{" & txtMinLen & ",}"
    'd.Pattern = "[\w0-9 /?.\-_=+$\\@!*\(\)#%&\|\[\]:;'""<>]{" & txtMinLen & ",}"
    d.Global = True
    
    Me.Caption = "Scanning for ASCII Strings..."
    push ret, "File: " & fso.FileNameFromPath(fPath)
    push ret, "MD5:  " & LCase(hash.HashFile(fPath))
    push ret, "Size: " & FileLen(fPath) & vbCrLf
    push ret, "Ascii Strings:" & vbCrLf & String(75, "-")
    
    f = FreeFile
    ReDim buf(9000)
    Open fPath For Binary Access Read As f
    
    pb.value = 0
    Do While pointer < LOF(f)
        If abort Then GoTo aborting
        pointer = Seek(f)
        x = LOF(f) - pointer
        If x < 1 Then Exit Do
        If x < 9000 Then ReDim buf(x)
        Get f, , buf()
        search buf, pointer
        setpb pointer, LOF(f)
    Loop
    
    lines = UBound(ret)
    rtf.text = Join(ret, vbCrLf)
    
    Erase ret
    
    Me.Caption = "Scanning for unicode strings.."
    push ret, ""
    push ret, "Unicode Strings:" & vbCrLf & String(75, "-")
    
    'd.Pattern = "([\w0-9 /?.\-=+$\\@!*\(\)#][\x00]){4,}"
    d.Pattern = "([\w0-9 /?.\-=+$\\@!\*\(\)#%~`\^&\|\{\}\[\]:;'""<>\,][\x00]){" & txtMinLen & ",}"
    'd.Pattern = "([\w0-9 /?.\-_=+$\\@!*\(\)#%&\|\[\]:;'""<>][\x00]){" & txtMinLen & ",}"
    
    ReDim buf(9000)
    pointer = 1
    Seek f, 1
    
    pb.value = 0
    Do While pointer < LOF(f)
        If abort Then GoTo aborting
        pointer = Seek(f)
        x = LOF(f) - pointer
        If x < 1 Then Exit Do
        If x < 9000 Then ReDim buf(x)
        Get f, , buf()
        search buf, pointer
        setpb pointer, LOF(f)
    Loop
    pb.value = 0
    
    Close f
     
    On Error Resume Next
    Dim topLine As Integer
    
    lines = lines + UBound(ret)
    LockWindowUpdate rtf.hwnd 'try to make it not jump when we add more...
    topLine = TopLineIndex(rtf)
    rtf.text = rtf.text & vbCrLf & vbCrLf & Join(ret, vbCrLf)
    ScrollToLine rtf, topLine
    LockWindowUpdate 0
    
    Erase ret
    Me.Caption = lines & " matches found..."
    Me.Show 1
   
    If chkFilter.value = 1 Then
        Me.Caption = Me.Caption & "  ( " & UBound(filtered) & " results filtered)"
    End If
    
    Me.Caption = Me.Caption & "  -  " & fPath
    running = False
    'RevertRedir fs
    
    
Exit Sub
hell:
      MsgBox "Error getting strings: " & Err.Description & "Line: " & Erl, vbExclamation
      Close f
done:
      'RevertRedir fs
      'Unload Me
      End
      
aborting:
      RevertRedir fs
      running = False
      abort = False
      pb.value = 0
      
End Sub

Private Sub search(buf() As Byte, offset As Long)
    Dim b As String
    Dim m As match
    
    b = StrConv(buf, vbUnicode)
    Set mc = d.Execute(b)
    
    For Each m In mc
        DoEvents
        If abort Then Exit Sub
        If chkFilter.value = 1 Then
            If Not Filter(m.value) Then AddResult m, offset
        Else
            AddResult m, offset
        End If
    Next
    
End Sub

'todo: this is not x64 safe..
Function AddResult(m As match, offset As Long)
    Dim x As Long, xx, sect As String, o As String
    
    If chkShowOffsets.value = 1 Then
        x = m.FirstIndex + offset - 1
        If optVa.value And pe.isLoaded = True Then
            xx = pe.OffsetToVA(x, sect).toString()
            If xx = "0" Then
                o = pad(x) & "  "
            Else
                o = sect & ":" & xx & "  "
            End If
        Else
            o = pad(x) & "  "
        End If
    End If
    
    push ret(), o & Replace(m.value, Chr(0), Empty)
    
End Function

Function pad(v, Optional leng = 8)
    On Error GoTo hell
    Dim x As String
    x = Hex(v)
    While Len(x) < leng
        x = "0" & x
    Wend
    pad = x
    Exit Function
hell:
    pad = x
End Function

'Function ApplyFilters(r() As String) As String()
'    Dim x, out() As String
'    Dim i, max
'
'    On Error Resume Next
'
'    max = UBound(r)
'    pb.value = 0
'    Me.Caption = "Applying filters..."
'
'    For Each x In r
'        If toManySpecialChars(x) Then
'            push filtered, x
'        ElseIf isJunk(x) Then
'            push filtered, x
'        ElseIf toManyNumbers(x) Then
'            push filtered, x
'        Else
'            push out, x
'        End If
'        i = i + 1
'        setPB i, max
'    Next
'
'    ApplyFilters = out
'    pb.value = 0
'
'End Function

Function Filter(x As String) As Boolean
    
    
    On Error Resume Next
    Dim f As String
    
    If InStr(x, "http://") > 0 Then
        Filter = False
    ElseIf toManySpecialChars(x) Then
        If IsIde() Then f = vbTab & vbTab & "(SpecialCharsFilter)"
        push filtered, x & f
        Filter = True
    ElseIf toManyRepeats(x) Then
        If IsIde() Then f = vbTab & vbTab & "(RepeatFilter)"
        push filtered, x & f
        Filter = True
    ElseIf toManyNumbers(x) Then
        If IsIde() Then f = vbTab & vbTab & "(NumberFilter)"
        push filtered, x & f
        Filter = True
    Else
        Filter = False
    End If
 
End Function

Function IsIde() As Boolean
    On Error GoTo hell
    Debug.Print 1 / 0
    Exit Function
hell: IsIde = True
End Function

Function toManyRepeats(ByVal s As String) As Boolean

    Dim os As String
    Dim hits As Long
    Dim pcent, i As Long, sl As Long, fl As Long
    
    os = s 'for debugging sake
    
    If Len(s) > 20 Then
        toManyRepeats = False
        Exit Function
    End If
    
    Dim b() As Byte
    b() = StrConv(s, vbFromUnicode, LANG_US)
    
    For i = 0 To UBound(b)
        If InStr(1, s, Chr(b(i))) > 0 Then
            s = Replace(s, Chr(b(i)), Empty)
            hits = hits + 1
        End If
        If Len(s) = 0 Then Exit For
    Next
    
    sl = UBound(b) + 1 'original length
    fl = hits
    
    pcent = fl / sl
    
    If pcent < 0.54445 Then toManyRepeats = True
        
End Function

Function toManySpecialChars(ByVal s) As Boolean

    'Const c = "/?.-_=+$@!*()#%~`^&|{}[]:;'""<>\,]"
    Const c = "?-_=+$@!*()#~`^&|{}[]:;'""<>,]" 'javascript fragments will trigger this...
    
    Dim sl As Long, i As Long
    Dim fl As Long
    Dim hits As Long
    Dim pcent As Long
    
    Dim cc
    
    sl = Len(s)
    
    For i = 1 To Len(c)
        cc = Mid(c, i, 1)
        s = Replace(s, cc, Empty)
    Next
       
    fl = Len(s)
    pcent = 100 - ((fl / sl) * 100)
    
    If pcent <= 20 Then
        toManySpecialChars = False
    Else
        toManySpecialChars = True
    End If
    
End Function

Function toManyNumbers(ByVal s) As Boolean
    
    Dim sl As Long, i As Long
    Dim fl As Long
    Dim hits As Long
    Dim pcent As Long
    
    Dim cc
    
    sl = Len(s)
    
    For i = 1 To 9
        s = Replace(s, CStr(i), Empty)
    Next
       
    fl = Len(s)
    pcent = 100 - ((fl / sl) * 100)
    
    If pcent <= 20 Then
        toManyNumbers = False
    Else
        toManyNumbers = True
    End If
    
End Function

Private Sub Label3_Click()
    On Error Resume Next
    Dim f As String
    If AryIsEmpty(filtered) Then
        MsgBox "No results have been filtered", vbInformation
    Else
        f = fso.GetFreeFileName(Environ("temp"))
        fso.WriteFile f, "Results filtered from main display: " & vbCrLf & vbCrLf & Join(filtered, vbCrLf)
        Shell "notepad.exe """ & f & """", vbNormalFocus
    End If
End Sub

Private Sub lblMore_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
   PopupMenu mnuMore
End Sub

Private Sub mnuChangeFont_Click()
    Dim f As CFont
    On Error Resume Next
    Dim dlg As New CCmnDlg
    Set f = dlg.ChooseFont(rtf)
    If Not f.selected Then Exit Sub
    rtf.Font.Name = f.Name
    rtf.Font.size = f.size
    rtf.Font.Bold = f.Bold
    If Err.Number = 0 Then
        SaveMySetting "strings.font.name", f.Name
        SaveMySetting "strings.font.size", f.size
        SaveMySetting "strings.font.bold", f.Bold
    End If
End Sub

Private Sub mnuDelphiFIlter_Click()
    On Error Resume Next
    
    Dim f As String, filt() As String
    Dim txt() As String, i, removed, tmp
    
    f = App.path & IIf(IsIde(), "\..\", "\") & "delphi_filter.txt"
    
    If Not fso.FileExists(f) Then
        MsgBox "Filter file not found: " & f
        Exit Sub
    End If
    
    filt = Split(fso.ReadFile(f), vbCrLf)
    txt = Split(rtf.text, vbCrLf)
    
    Me.Caption = "Filter contains: " & UBound(filt)
    pb.max = 100
    pb.value = 0
    i = 0
    
    tmp = rtf.text
    
    For i = 0 To UBound(filt)
        tmp = Replace(tmp, filt(i), "-[dtd]-")
        If i Mod 100 = 0 Then
            pb.value = i / UBound(txt) * 100
            DoEvents
        End If
        If abort Then Exit Sub
    Next
    
    Erase filt
    txt = Split(tmp, vbCrLf)
    pb.value = 0
    
    For i = 0 To UBound(txt)
        
        If InStr(txt(i), "-[dtd]-") < 1 Then
            push filt, txt(i)
        Else
            removed = removed + 1
        End If
        
        If i Mod 100 = 0 Then
            pb.value = i / UBound(txt) * 100
            DoEvents
        End If
        
        If abort Then Exit Sub
    Next
    
    Me.Caption = UBound(filt) & " results shown  (" & removed & " removed by Delphi Filter)"
    pb.value = 0
    
    rtf.text = Join(filt, vbCrLf)
    
End Sub

Private Sub mnuStringDiff_Click()
    
    Dim f2 As String
    Dim c1 As New CollectionEx
    Dim c2 As New CollectionEx
    Dim dif1 As CollectionEx
    Dim dif2 As CollectionEx
    Dim tmp As String
    Dim dat() As String
    
    On Error Resume Next
    
    f2 = dlg.OpenDialog("", "Load String Dump")
    If Len(f2) = 0 Then Exit Sub
    
    c1.fromArray Split(rtf.text, vbCrLf), , True, True
    c2.fromTextFile f2, , True
    
    Set dif1 = c1.diff(c2)
    Set dif2 = c2.diff(c1)
     
    tmp = fso.GetFreeFileName(Environ("temp"))
    
    push dat, "Strings not found in File1: " & fso.FileNameFromPath(curFile) & " - " & Now
    push dat, dif1.toString()
    push dat, String(50, "-")
    push dat, "Strings not found in File2: " & fso.FileNameFromPath(f2) & " - " & Now
    push dat, dif2.toString()
    
    fso.WriteFile tmp, Join(dat, vbCrLf)
    Shell "notepad.exe """ & tmp & """", vbNormalFocus
    
    
End Sub

Private Sub mnuStringMatch_Click()
    
    Dim f2 As String
    Dim c1 As New CollectionEx
    Dim c2 As New CollectionEx
    Dim matches As CollectionEx
    Dim tmp As String
    
    On Error Resume Next
    
    f2 = dlg.OpenDialog("", "Load String Dump")
    If Len(f2) = 0 Then Exit Sub
    
    c1.fromArray Split(rtf.text, vbCrLf), , True, True
    c2.fromTextFile f2, , True
    
    Set matches = c1.findMatches(c2)
    
    If matches.Count = 0 Then
        MsgBox "No matches found"
    Else
        tmp = fso.GetFreeFileName(Environ("temp"))
        matches.toTextFile tmp, "Matches in both " & fso.FileNameFromPath(curFile) & " - " & fso.FileNameFromPath(f2) & " - " & Now
        Shell "notepad.exe """ & tmp & """", vbNormalFocus
    End If
    
End Sub

Private Sub optRaw_Click()
    If Not formLoaded Then Exit Sub
    If chkShowOffsets.value = 1 Then ParseFile curFile, True
End Sub

Private Sub optVa_Click()
    If Not formLoaded Then Exit Sub
    If chkShowOffsets.value = 1 Then ParseFile curFile, True
End Sub

Private Sub tmrReRun_Timer()
    tmrReRun.Enabled = False
    ParseFile curFile
End Sub
