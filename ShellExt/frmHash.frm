VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHash 
   Caption         =   "Directory File Hasher - Right Click on ListView for Menu Options"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   12060
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   12060
   StartUpPosition =   2  'CenterScreen
   Begin ShellExt.ucFilterList lv 
      Height          =   3660
      Left            =   90
      TabIndex        =   1
      Top             =   315
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   6456
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   225
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Begin VB.Menu mnuSaveTable 
         Caption         =   "Save Table"
      End
      Begin VB.Menu mnuCopyTop 
         Caption         =   "Copy"
         Begin VB.Menu mnuCopyTable 
            Caption         =   "All"
         End
         Begin VB.Menu mnuCopySelected 
            Caption         =   "Selected"
         End
         Begin VB.Menu mnuCopyTableCSV 
            Caption         =   "CSV"
         End
         Begin VB.Menu mnuCopyHashs 
            Caption         =   "Hashs"
         End
         Begin VB.Menu mnuCopyDetailed 
            Caption         =   "Detailed"
         End
      End
      Begin VB.Menu mnuDiv 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRenameToMD5 
         Caption         =   "Rename All to Hash"
      End
      Begin VB.Menu mnuMakeExtSafe 
         Caption         =   "Make All Extensions Safe"
      End
      Begin VB.Menu mnuCustomExtension 
         Caption         =   "Set All Custom Extension "
      End
      Begin VB.Menu mnuMakeSubFolders 
         Caption         =   "Make folders for each"
      End
      Begin VB.Menu mnuRecursiveHash 
         Caption         =   "Hash all files below"
      End
      Begin VB.Menu mnuMoveSelected 
         Caption         =   "Move Selected to SubFolder"
      End
      Begin VB.Menu mnuReports 
         Caption         =   "Reports"
         Begin VB.Menu mnuDisplayUnique 
            Caption         =   "Display unique"
         End
         Begin VB.Menu mnuUniqueImpHash 
            Caption         =   "Unique ImpHash"
         End
         Begin VB.Menu mnuFilePropsReport 
            Caption         =   "File Properties"
         End
         Begin VB.Menu mnuDieReport 
            Caption         =   "DiE Report"
         End
         Begin VB.Menu mnuSectNameReport 
            Caption         =   "Section Names"
         End
         Begin VB.Menu mnuStringsDumpAll 
            Caption         =   "Strings Dump"
         End
         Begin VB.Menu mnuRich 
            Caption         =   "Rich Header"
            Begin VB.Menu mnuRichBasic 
               Caption         =   "Basic"
            End
            Begin VB.Menu mnuRichReport 
               Caption         =   "Detailed"
            End
         End
         Begin VB.Menu mnuImportsTop 
            Caption         =   "Imports"
            Begin VB.Menu mnuDumpImports 
               Caption         =   "Dlls"
               Index           =   0
            End
            Begin VB.Menu mnuDumpImports 
               Caption         =   "All"
               Index           =   1
            End
         End
         Begin VB.Menu mnuDumpExports 
            Caption         =   "Exports"
         End
      End
      Begin VB.Menu mnuSpacer33 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVTTop 
         Caption         =   "VT"
         Begin VB.Menu mnuVTAll 
            Caption         =   "Lookup All"
         End
         Begin VB.Menu mnuVTLookupSelected 
            Caption         =   "Lookup Selected"
         End
         Begin VB.Menu mnuSubmitSelToVT 
            Caption         =   "Submit Selected"
         End
      End
      Begin VB.Menu mnuGoogleSelected 
         Caption         =   "Google Selected"
      End
      Begin VB.Menu mnudivider 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeleteSelected 
         Caption         =   "Deleted Selected Files"
      End
      Begin VB.Menu mnuDeleteDuplicates 
         Caption         =   "Delete All Duplicates"
      End
      Begin VB.Menu mnuHashDiff 
         Caption         =   "Hash Diff against.."
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Options"
         Begin VB.Menu mnuSetFont 
            Caption         =   "Set Font"
         End
         Begin VB.Menu mnuUseSHA256 
            Caption         =   "Use SHA256"
         End
         Begin VB.Menu mnuHumanReadableSizes 
            Caption         =   "KB/MB Sizes"
            Shortcut        =   ^M
         End
      End
   End
End
Attribute VB_Name = "frmHash"
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

'7-6-05  Added Delete All Duplicates option
'4-19-12 moved buttons to right click menu options, integrated VirusTotal.exe options
'5.17.12 added progress bar, fixed integer overflow in vbDevKit.CWinHash
'9.18.12 added x64 file system redirection awareness to main hashing routines (not to all right click options..)
'Implements IExtColSort

Private Declare Function GetWindowLong Lib "User" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "User" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long

Const MF_STRING As Long = &H0
Const IDM_COMPARE As Long = 1010
Const IDM_HASHSEARCH As Long = 1011
Const IDM_STRINGDUMP As Long = 1012

Const WM_SYSCOMMAND = &H112
   
Dim WithEvents sc As CSubclass2
Attribute sc.VB_VarHelpID = -1

Public path As String
Public isComplete As Boolean
Dim abort As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Debug.Print KeyCode & " " & Shift
    If KeyCode = 77 And Shift = 2 Then 'Ctrl-M
        mnuHumanReadableSizes_Click
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    abort = True
    'SaveSetting "shellext", "settings", "mnuIncludeFileName", mnuIncludeFileName.Checked
    sc.DetatchMessage Me.hwnd, WM_SYSCOMMAND
    SaveMySetting "mnuUseSHA256.Checked", mnuUseSHA256.Checked
End Sub



Private Sub lv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then PopupMenu mnuPopup
End Sub

Private Sub mnuDumpExports_Click()
    On Error Resume Next
    
    Dim li As ListItem
    Dim fPath As String
    Dim selOnly As Boolean
    Dim doit As Boolean
    Dim fs As New clsFileStream
    Dim report As String
    Dim tmp As String
    Dim pe2 As New sppe3.CPEEditor

    report = fso.GetFreeFileName(Environ("temp"))
    fs.fOpen report, otwriting
    fs.WriteLine vbCrLf & "This is a temp file saveAs to save"
    fs.WriteDivider
    
    selOnly = (lvSelCount(lv) > 1)
    
    pb.value = 0
    If selOnly Then
        pb.max = lvSelCount(lv)
    Else
        pb.max = lv.ListItems.Count
    End If
    
    For Each li In lv.ListItems
        doit = False
        If selOnly Then
            If li.selected Then doit = True
        Else
            doit = True
        End If
        If doit Then
            fPath = li.Tag
            If fso.FileExists(fPath) Then
                fs.WriteLine "File: " & fPath
                If Not pe2.LoadFile(fPath) Then
                    fs.WriteLine "Error: " & pe2.errMessage
                Else
                    If pe2.Exports.functions.Count = 0 Then
                        fs.WriteLine "No Exports"
                    Else
                        fs.WriteLine c2s(pe2.Exports.dumpNames, ", ")
                    End If
                End If
                fs.WriteLine rh.dump
                fs.WriteDivider
            End If
        End If
        pb.value = pb.value + 1
    Next
    
    pb.value = 0
    fs.fClose
    Shell "notepad.exe """ & report & """", vbNormalFocus

End Sub

Private Sub mnuDumpImports_Click(index As Integer)
    On Error Resume Next
    
    Dim li As ListItem
    Dim fPath As String
    Dim selOnly As Boolean
    Dim doit As Boolean
    Dim fs As New clsFileStream
    Dim report As String
    Dim tmp As String
    Dim pe2 As New sppe3.CPEEditor
    Dim rh As CRichHeader
    
    report = fso.GetFreeFileName(Environ("temp"))
    fs.fOpen report, otwriting
    fs.WriteLine vbCrLf & "This is a temp file saveAs to save"
    fs.WriteDivider
    
    selOnly = (lvSelCount(lv) > 1)
    
    pb.value = 0
    If selOnly Then
        pb.max = lvSelCount(lv)
    Else
        pb.max = lv.ListItems.Count
    End If
    
    For Each li In lv.ListItems
        doit = False
        If selOnly Then
            If li.selected Then doit = True
        Else
            doit = True
        End If
        If doit Then
            fPath = li.Tag
            If fso.FileExists(fPath) Then
                fs.WriteLine "File: " & fPath
                If Not pe2.LoadFile(fPath) Then
                    fs.WriteLine "Error: " & pe2.errMessage
                Else
                    If pe2.Imports.Modules.Count = 0 Then
                        fs.WriteLine "No Imports ?"
                    Else
                        Dim ci As CImport
                        For Each ci In pe2.Imports.Modules
                            If index = 1 Then tmp = ": " & c2s(ci.functions, ", ")
                            fs.WriteLine ci.DllName & tmp
                        Next
                    End If
                End If
                fs.WriteLine rh.dump
                fs.WriteDivider
            End If
        End If
        pb.value = pb.value + 1
    Next
    
    pb.value = 0
    fs.fClose
    Shell "notepad.exe """ & report & """", vbNormalFocus
End Sub

Private Sub mnuHumanReadableSizes_Click()
    mnuHumanReadableSizes.Checked = Not mnuHumanReadableSizes.Checked
    If lv.ListItems.Count > 0 Then
         If Not mnuHumanReadableSizes.Checked Then
            For Each li In lv.ListItems
                li.subItems(1) = pad(FileLen(li.Tag))
            Next
        Else
            For Each li In lv.ListItems
                li.subItems(1) = pad(FileSize(li.Tag, False))
            Next
        End If
    End If
End Sub

Private Sub mnuRichBasic_Click()
On Error Resume Next
    
    Dim li As ListItem
    Dim fPath As String
    Dim selOnly As Boolean
    Dim doit As Boolean
    Dim fs As New clsFileStream
    Dim report As String
    Dim tmp As String
    Dim pe2 As New sppe3.CPEEditor
    Dim rh As CRichHeader
    
    report = fso.GetFreeFileName(Environ("temp"))
    fs.fOpen report, otwriting
    fs.WriteLine vbCrLf & "This is a temp file saveAs to save"
    fs.WriteDivider
    
    selOnly = (lvSelCount(lv) > 1)
    
    pb.value = 0
    If selOnly Then
        pb.max = lvSelCount(lv)
    Else
        pb.max = lv.ListItems.Count
    End If
    
    For Each li In lv.ListItems
        doit = False
        If selOnly Then
            If li.selected Then doit = True
        Else
            doit = True
        End If
        If doit Then
            fPath = li.Tag
            If fso.FileExists(fPath) Then
                fs.WriteLine "File: " & fPath
                Set rh = New CRichHeader
                Call rh.Load(fPath)
                fs.WriteLine "Exists: " & rh.Exists & "       CheckSumOk: " & rh.checkSumOk & "       Entries: " & rh.entries.Count
                fs.WriteDivider
            End If
        End If
        pb.value = pb.value + 1
    Next
    
    pb.value = 0
    fs.fClose
    Shell "notepad.exe """ & report & """", vbNormalFocus
End Sub

Private Sub mnuRichReport_Click()
    On Error Resume Next
    
    Dim li As ListItem
    Dim fPath As String
    Dim selOnly As Boolean
    Dim doit As Boolean
    Dim fs As New clsFileStream
    Dim report As String
    Dim tmp As String
    Dim pe2 As New sppe3.CPEEditor
    Dim rh As CRichHeader
    
    report = fso.GetFreeFileName(Environ("temp"))
    fs.fOpen report, otwriting
    fs.WriteLine vbCrLf & "This is a temp file saveAs to save"
    fs.WriteDivider
    
    selOnly = (lvSelCount(lv) > 1)
    
    pb.value = 0
    If selOnly Then
        pb.max = lvSelCount(lv)
    Else
        pb.max = lv.ListItems.Count
    End If
    
    For Each li In lv.ListItems
        doit = False
        If selOnly Then
            If li.selected Then doit = True
        Else
            doit = True
        End If
        If doit Then
            fPath = li.Tag
            If fso.FileExists(fPath) Then
                fs.WriteLine "File: " & fPath
                Set rh = New CRichHeader
                Call rh.Load(fPath)
                fs.WriteLine rh.dump
                fs.WriteDivider
            End If
        End If
        pb.value = pb.value + 1
    Next
    
    pb.value = 0
    fs.fClose
    Shell "notepad.exe """ & report & """", vbNormalFocus
End Sub

Private Sub mnuSetFont_Click()
    On Error Resume Next
    Dim s() As String, sz As Long, tmp As String, msg As String

    Dim f As CFont, c As New CCmnDlg
    Set f = frmSelectFont.SelectFont(c, lv.currentLV)
    If f Is Nothing Then Exit Sub
    SetFont f.Name & "," & f.Size
    
'    tmp = GetMySetting("lvFont", "MS Sans Serif,8")
'
'tryAgain:
'    tmp = InputBox("Enter FontName,Size to use " & msg, , tmp)
'    If Len(tmp) = 0 Then Exit Sub
'
'    s = Split(tmp, ",")
'    If UBound(s) >= 1 Then sz = CLng(s(1))
'    If sz = 0 Then
'        msg = "Invalid Size! specified!"
'        GoTo tryAgain
'    End If
    
'    SetFont tmp
    
End Sub

Function SetFont(csvNameSize) As Boolean
    On Error Resume Next
    Dim s() As String, sz As Long
    If Len(csvNameSize) = 0 Then Exit Function
    s = Split(csvNameSize, ",")
    If UBound(s) >= 1 Then sz = CLng(s(1))
    If sz = 0 Then sz = 8
    lv.SetFont s(0), sz
    SetFont = (Err.Number = 0)
    If Err.Number = 0 Then SaveMySetting "lvFont", csvNameSize
End Function

Private Sub mnuStringsDumpAll_Click()
    
    On Error Resume Next
    Dim li As ListItem
    Dim f As String
    
    abort = False
    
    For Each li In lv.ListItems
        If abort Then Exit For
        If VBA.Left(li.text, 4) <> "str_" Then
            f = path & "\" & li.text
            If fso.FileExists(f) Then
                frmStrings.ParseFile f
                frmStrings.AutoSave
            End If
        End If
    Next
    
    Unload frmStrings
    
End Sub

Private Sub mnuCopyDetailed_Click()
    
    Dim selOnly As Boolean, li As ListItem, ret() As String, tmp As String, SelCount As Long, org As String
    
    org = Me.Caption
    pb.value = 0
    frmFileHash.Visible = False
   
    SelCount = lv_selCount(lv)
    If SelCount > 1 Then selOnly = True
    
    If selOnly Then
         pb.max = SelCount
    Else
         pb.max = lv.ListItems.Count
    End If
    
    For Each li In lv.ListItems
        If selOnly Then
            If li.selected Then GoSub AddItem
        Else
            GoSub AddItem 'i never use it and its there so... wtf not
        End If
    Next
    
    Unload frmFileHash
    pb.value = 0
    Clipboard.Clear
    Clipboard.SetText Join(ret, vbCrLf)
    Me.Caption = org
    
Exit Sub

AddItem:
    Me.Caption = "Processing: " & li.text
    tmp = Replace(frmFileHash.ShowFileStats(path & "\" & li.text, True), vbCrLf & vbCrLf, vbCrLf)
    If Right(tmp, 2) = vbCrLf Then tmp = Mid(tmp, 1, Len(tmp) - 2)
    If Right(tmp, 2) = vbCrLf Then tmp = Mid(tmp, 1, Len(tmp) - 2)
    push ret(), tmp
    push ret(), String(50, "-")
    pb.value = pb.value + 1
    Return
    
End Sub

Function lvSelCount(lv As ListView) As Long
    Dim i As Long, li As ListItem
    For Each li In lv.ListItems
        If li.selected Then i = i + 1
    Next
    lvSelCount = i
End Function

Private Sub mnuDieReport_Click()
    On Error Resume Next
    
    Dim li As ListItem
    Dim fPath As String
    Dim selOnly As Boolean
    Dim doit As Boolean
    Dim fs As New clsFileStream
    Dim report As String
    Dim tmp As String
    
    report = fso.GetFreeFileName(Environ("temp"))
    fs.fOpen report, otwriting
    fs.WriteLine vbCrLf & "This is a temp file saveAs to save"
    fs.WriteDivider
    
    selOnly = (lvSelCount(lv) > 1)
    
    pb.value = 0
    If selOnly Then
        pb.max = lvSelCount(lv)
    Else
        pb.max = lv.ListItems.Count
    End If
    
    For Each li In lv.ListItems
        doit = False
        If selOnly Then
            If li.selected Then doit = True
        Else
            doit = True
        End If
        If doit Then
            fPath = li.Tag
            If fso.FileExists(fPath) Then
                fs.WriteLine "File: " & fPath
                 If DiEScan(fPath, tmp) Then
                    fs.WriteLine "DiE:      " & tmp
                 Else
                    fs.WriteLine "DiE Error " & tmp
                 End If
                fs.WriteDivider
            End If
        End If
        pb.value = pb.value + 1
    Next
    
    pb.value = 0
    fs.fClose
    Shell "notepad.exe """ & report & """", vbNormalFocus
End Sub

Private Sub mnuFilePropsReport_Click()
    
    On Error Resume Next
    
    Dim li As ListItem
    Dim fPath As String
    Dim selOnly As Boolean
    Dim doit As Boolean
    Dim fs As New clsFileStream
    Dim report As String
    
    report = fso.GetFreeFileName(Environ("temp"))
    fs.fOpen report, otwriting
    fs.WriteLine vbCrLf & "This is a temp file saveAs to save"
    fs.WriteDivider
    
    selOnly = (lvSelCount(lv) > 1)
    
    For Each li In lv.ListItems
        doit = False
        If selOnly Then
            If li.selected Then doit = True
        Else
            doit = True
        End If
        If doit Then
            fPath = li.Tag
            If fso.FileExists(fPath) Then
                fs.WriteLine "File: " & fPath
                fs.WriteLine FileProps.FileInfo(fPath).asStr
                fs.WriteDivider
            End If
        End If
    Next
    
    fs.fClose
    Shell "notepad.exe """ & report & """", vbNormalFocus
    
End Sub

Private Sub mnuMoveSelected_Click()
    On Error Resume Next
    
    Dim li As ListItem
    Dim fPath As String
    Dim newDir As String
    Dim fname As String
    Dim i As Long
    Dim moved As Long, total As Long
    
    newDir = dlg.FolderDialog2(path)
    If Len(newDir) = 0 Then Exit Sub
    
    For i = lv.ListItems.Count To 1 Step -1
        Set li = lv.ListItems(i)
        If li.selected Then
            total = total + 1
            fPath = li.Tag
            fname = "\" & fso.FileNameFromPath(fPath)
            If Not fso.FileExists(newDir & fname) Then
                fso.Move fPath, newDir
                lv.ListItems.Remove i
                moved = moved + 1
            End If
        End If
    Next
    
    MsgBox moved & "/" & total & " files moved", vbInformation
    
End Sub

Private Sub mnuRecursiveHash_Click()
    frmRecursiveHashFiles.RecursiveHashDir path
End Sub

Private Sub mnuSectNameReport_Click()
    
    On Error Resume Next
    
    Dim li As ListItem
    Dim fPath As String
    Dim selOnly As Boolean
    Dim doit As Boolean
    Dim fs As New clsFileStream
    Dim report As String
    Dim tmp As String
    Dim pe2 As New sppe3.CPEEditor
    
    report = fso.GetFreeFileName(Environ("temp"))
    fs.fOpen report, otwriting
    fs.WriteLine vbCrLf & "This is a temp file saveAs to save"
    fs.WriteDivider
    
    selOnly = (lvSelCount(lv) > 1)
    
    pb.value = 0
    If selOnly Then
        pb.max = lvSelCount(lv)
    Else
        pb.max = lv.ListItems.Count
    End If
    
    For Each li In lv.ListItems
        doit = False
        If selOnly Then
            If li.selected Then doit = True
        Else
            doit = True
        End If
        If doit Then
            fPath = li.Tag
            If fso.FileExists(fPath) Then
                fs.WriteLine "File: " & fPath
                 If pe2.LoadFile(fPath) Then
                    fs.WriteLine sectNamesList(pe2.Sections)
                 Else
                    fs.WriteLine "DiE Error " & tmp
                 End If
                fs.WriteDivider
            End If
        End If
        pb.value = pb.value + 1
    Next
    
    pb.value = 0
    fs.fClose
    Shell "notepad.exe """ & report & """", vbNormalFocus
    
End Sub

Private Function sectNamesList(c As Collection) As String
    On Error Resume Next
    Dim cs As sppe3.CSection, tmp As String
    For Each cs In c
        tmp = tmp & "," & cs.nameSec
    Next
    sectNamesList = Mid(tmp, 2)
End Function



Private Sub mnuUniqueImpHash_Click()
    
    On Error Resume Next
    Dim li As ListItem
    Dim c As New CollectionEx 'key = imphash, val = csv md5 list
    Dim pe As New CPEEditor
    Dim ih As String, fPath As String
    Dim i As Long
    Dim r() As String
    Dim k() As String
    Dim x
    
    pb.value = 0
    pb.max = lv.ListItems.Count
    
    push r, "imphash : hits"
    push r, vbTab & "sample md5" & vbCrLf
    
    For Each li In lv.ListItems
        fPath = li.Tag
        If pe.LoadFile(fPath) Then
                ih = LCase(pe.impHash())
                If c.keyExists(ih) Then
                    c(ih, 1) = c(ih) & "," & li.text 'did i mention i love CollectionEx?
                Else
                    c.Add li.text, ih
                End If
        End If
        pb.value = pb.value + 1
    Next
    
    pb.value = 0
    
    For i = 1 To c.Count
        k = Split(c.Item(i), ",")
        push r, c.keyForIndex(i) & " hits: " & (UBound(k) + 1)
        For Each x In k
            push r, vbTab & x
        Next
        push r, vbCrLf
    Next
    
    'Clipboard.Clear
    'Clipboard.SetText Join(r, vbCrLf)
    
    Dim report As String, fs As New clsFileStream
    
    report = fso.GetFreeFileName(Environ("temp"))
    fs.fOpen report, otwriting
    fs.WriteLine vbCrLf & "This is a temp file saveAs to save" & vbCrLf
    fs.WriteLine c.Count & " unique imphashs copied for " & pb.max & " samples"
    fs.WriteDivider
    fs.WriteLine Join(r, vbCrLf)
    fs.fClose
    
    Shell "notepad.exe """ & report & """", vbNormalFocus

End Sub

Private Sub mnuUseSHA256_Click()
    On Error Resume Next
    mnuUseSHA256.Checked = Not mnuUseSHA256.Checked
    SaveMySetting "mnuUseSHA256.Checked", mnuUseSHA256.Checked
    lv.ChangeColHeaderText 3, IIf(mnuUseSHA256.Checked, "SHA256", "MD5")
    HashDir path
End Sub

Private Sub sc_MessageReceived(hwnd As Long, wMsg As Long, wParam As Long, lParam As Long, Cancel As Boolean)
    If wParam = IDM_COMPARE Then frmCompareHashSets.Show
    If wParam = IDM_HASHSEARCH Then
        frmMD5FileSearch.Show
        frmMD5FileSearch.txtBaseDir = Me.path
    End If
    If wParam = IDM_STRINGDUMP Then mnuStringsDumpAll_Click
End Sub

Sub Form_Load()
    On Error Resume Next
    
    Me.Icon = myIcon
    'Me.Icon = frmMain.Icon   'can not do this as frmMain has code in form_load and its already unloaded by this point (if we use the diff feature only) !! I so confuzzzed..
    
    
    mnuPopup.Visible = False
    'Set lv.ExternalColumnSorter = Me
    lv.SetColumnHeaders "File*,ByteSize,MD5,CompileDate (GMT)", "2000,1500,3000,2500"
    SetFont GetMySetting("lvFont", "MS Sans Serif,8")
    
    'lv.ColumnHeaders(1).Width = lv.Width - lv.ColumnHeaders(2).Width - 400 - lv.ColumnHeaders(3).Width - lv.ColumnHeaders(4).Width
    'mnuIncludeFileName.Checked = GetSetting("shellext", "settings", "mnuIncludeFileName", 1)
    
    If sc Is Nothing Then
        Set sc = New CSubclass2
        sc.AttachMessage Me.hwnd, WM_SYSCOMMAND
        AppendMenu GetSystemMenu(Me.hwnd, 0), MF_STRING, IDM_COMPARE, "Compare Hash Sets..."
        AppendMenu GetSystemMenu(Me.hwnd, 0), MF_STRING, IDM_HASHSEARCH, "Hash Search..."
        AppendMenu GetSystemMenu(Me.hwnd, 0), MF_STRING, IDM_STRINGDUMP, "Generate Strings Dump for All"
    End If
    
    mnuUseSHA256.Checked = GetMySetting("mnuUseSHA256.Checked", 0)
    If mnuUseSHA256.Checked Then
        'lv.ColumnHeaders(3).text = "SHA256"
        lv.ChangeColHeaderText 3, "SHA256"
    End If
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    lv.Width = Me.Width - lv.Left - 400
    lv.Height = Me.Height - lv.top - 550
    pb.Width = lv.Width
   'lv.ColumnHeaders(lv.ColumnHeaders.Count).Width = lv.Width - lv.ColumnHeaders(lv.ColumnHeaders.Count).Left - 100
End Sub

'md5 or sha256 depending on mode
Public Function GetFilesForHash(md5, Optional ByRef totalHits As Long, Optional includeCount As Boolean = True) As String    ' returns csv of file names or nothing

    Dim li As ListItem
    Dim ret As String
    Dim cnt As Long
     
    md5 = CStr(md5)
    
    For Each li In lv.ListItems
        If li.subItems(2) = md5 Then
            ret = ret & li.text & " , "
            cnt = cnt + 1
        End If
    Next
    
    ret = trim(ret)
    If Len(ret) > 0 And Right(ret, 1) = "," Then
        ret = Mid(ret, 1, Len(ret) - 1)
    End If
    
    If Len(ret) > 0 Then
        If includeCount Then
            GetFilesForHash = cnt & " hits: " & ret
        Else
            GetFilesForHash = ret
        End If
        totalHits = totalHits + cnt
    End If
       
End Function

Sub setpb(cur, max)
    On Error Resume Next
    pb.value = (cur / max) * 100
    Me.Refresh
    DoEvents
End Sub

Sub HashDir(dPath As String, Optional diffMode As Boolean = False)
   
    On Error GoTo out
    Dim f() As String, i As Long
    Dim pf As String
    Dim fs As Long
    'MsgBox "entering hash dir"
    
    abort = False
    path = dPath
    pf = fso.GetParentFolder(path) & "\"
    pf = Replace(path, pf, Empty)
    lv.ListItems.Clear
    
    Me.Caption = Me.Caption & "    Folder: " & pf
        
    fs = DisableRedir()
    If Not fso.FolderExists(dPath) Then
        MsgBox "Folder not found: " & dPath
        GoTo done
    End If
             
    f() = fso.GetFolderFiles(dPath)
    RevertRedir fs
    
    If AryIsEmpty(f) Then
        If MsgBox("No files in this directory, do you want to hash all files within all subfolders?", vbInformation + vbYesNo) = vbNo Then
            GoTo done
        Else
            frmRecursiveHashFiles.RecursiveHashDir dPath
            RevertRedir fs
            isComplete = True
            Unload Me
            Exit Sub
        End If
    End If
     
    'MsgBox "Going to scan " & UBound(f) & " files"
    pb.value = 0
    Me.Visible = True
    
    For i = 0 To UBound(f)
         If abort Then Exit Sub
         handleFile f(i)
         setpb i, UBound(f)
    Next
    pb.value = 0
    'MsgBox "ready to show"
     
    On Error Resume Next
    Me.Show 1
    
    Me.Caption = Me.Caption & "    Files: " & lv.ListItems.Count
    
    isComplete = True
    
    Exit Sub
out:
    MsgBox "HashFiles Error: " & Err.Description, vbExclamation
done:
    'Unload Me
    RevertRedir fs
    isComplete = True
    If Not diffMode Then End
    
End Sub



Function KeyExistsInCollection(c As Collection, val As String) As Boolean
    On Error GoTo nope
    Dim t
    t = c(val)
    KeyExistsInCollection = True
 Exit Function
nope: KeyExistsInCollection = False
End Function

'Private Sub IExtColSort_ExternalColumnSort(lv As MSComctlLib.IListView, Column As MSComctlLib.IColumnHeader)
''Private Sub lv_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'    Dim li As ListItem
'    On Error Resume Next
'    If GetKeyState(vbKeyShift) Then
'        If humanReadableSizes Then
'            humanReadableSizes = False
'            For Each li In lv.ListItems
'                li.subItems(1) = pad(FileLen(li.Tag))
'            Next
'        Else
'            humanReadableSizes = True
'            For Each li In lv.ListItems
'                li.subItems(1) = pad(FileSize(li.Tag, False))
'            Next
'        End If
'    Else
'        LV_ColumnSort lv, Column
'    End If
'End Sub

Private Sub mnuCopyHashs_Click()
    Dim li As ListItem
    Dim t As String
    
    For Each li In lv.ListItems
        t = t & li.subItems(2) & vbCrLf
    Next
    
    Clipboard.Clear
    Clipboard.SetText t
    MsgBox "Copy Complete", vbInformation
End Sub

Private Sub mnuCopySelected_Click()

    Dim li As ListItem
    Dim t As String
    Dim ln As Long
    
    ln = LongestFileName() + 3
    
    For Each li In lv.ListItems
        If li.selected Then
            t = t & rpad(li.text, ln) & vbTab & li.subItems(1) & vbTab & li.subItems(2) & vbTab & li.subItems(3) & vbCrLf
        End If
    Next
    
    Clipboard.Clear
    Clipboard.SetText t
    'MsgBox "Copy Complete", vbInformation
    
End Sub

Private Sub mnuCopyTableCSV_Click()
    mnuCopyTable_Click
    t = Clipboard.GetText
    Clipboard.SetText Replace(t, vbTab, ",")
End Sub

Private Sub mnuCustomExtension_Click()
    On Error Resume Next
    Dim li As ListItem
    Dim pdir As String
    Dim i As Long
    Dim ext As String
    
    ext = InputBox("Enter custom extension. Can not be blank")
    If Len(ext) = 0 Then Exit Sub
    If VBA.Left(ext, 1) <> "." Then ext = "." & ext
    
    For Each li In lv.ListItems
        i = 1
        fPath = li.Tag
        fname = li.text
        pdir = fso.GetParentFolder(fPath) & "\"
        
        If InStrRev(fname, ".") > 1 Then
            fname = Mid(fname, 1, InStrRev(fname, ".") - 1)
        End If
        
        h = fname & ext
        
        If LCase(VBA.Right(fname, 4)) = ".txt" Then GoTo nextone  'txt files are fine..
        If LCase(VBA.Right(fname, Len(ext))) = LCase(ext) Then GoTo nextone   'already set
        
        While fso.FileExists(pdir & h) 'dont delete dups, but append counter onto end..
            h = fname & "_" & i
            i = i + 1
        Wend
        
        Name fPath As pdir & h
    
        li.text = h
        li.Tag = pdir & h
        li.EnsureVisible
        'lv.Refresh
        DoEvents
        
nextone:
    Next

End Sub

Private Sub mnuDeleteDuplicates_Click()
    
    Dim li As ListItem
    Dim hashs As New Collection
    Dim h As String
    Dim f As String
    Dim cnt As Long
    
    Close
    
    Const msg As String = "Are you sure you want to DELETE all DUPLICATE files?"
    If MsgBox(msg, vbYesNo) = vbNo Then Exit Sub
    
    For Each li In lv.ListItems
        h = li.subItems(2)
        If InStr(h, "Error") < 1 Then
            If KeyExistsInCollection(hashs, h) Then
                li.Tag = "DeleteMe"
                cnt = cnt + 1
            Else
                li.Tag = ""
                hashs.Add h, h
            End If
        End If
    Next
        
    On Error Resume Next
    Dim errs As Long
    
    For i = lv.ListItems.Count To 1 Step -1
        Set li = lv.ListItems(i)
        If li.Tag = "DeleteMe" Then
            f = path & "\" & li.text
            If fso.FileExists(f) Then
                If Not fso.DeleteFile(f) Then
                    errs = errs + 1
                    Err.Clear
                End If
            End If
            lv.ListItems.Remove li.index
        End If
    Next

    MsgBox cnt & " files deleted, " & lv.ListItems.Count & " remain" & IIf(errs > 0, " " & errs & " errors", ""), vbInformation
    
End Sub

Private Sub mnuDisplayUnique_Click()

     Dim li As ListItem
     Dim hashs As New Collection 'to perform unique value lookup and corrolate to ary index
     Dim h() As String 'count per hash    '\_matched arrays
     Dim b() As String 'actual hash value '/
     Dim hash As String
     Dim v As Long
     Dim i As Long
     
     On Error GoTo hell
     
     ReDim h(0) 'we cant use 0 anyway cause collections index start at 1
     ReDim b(0)
     
     For Each li In lv.ListItems
        hash = li.subItems(2)
        If KeyExistsInCollection(hashs, hash) Then
            i = hashs(hash)
            h(i) = h(i) + 1
        Else
            push h, 1
            push b, hash
            i = UBound(h)
            hashs.Add i, hash
        End If
     Next
     
     Dim tmp() As String
         
     For i = 1 To UBound(h)
        push tmp, b(i) & "   -   " & Me.GetFilesForHash(b(i))
     Next
     
     Dim t As String
     t = Environ("TMP")
     If Len(t) = 0 Then t = Environ("TEMP")
     If Len(t) = 0 Or Not fso.FolderExists(t) Then
            MsgBox Join(tmp, vbCrLf)
            Exit Sub
     End If
     
     t = fso.GetFreeFileName(t)
     fso.WriteFile t, Join(tmp, vbCrLf)
     
     Shell "notepad """ & t & """", vbNormalFocus
     fso.DeleteFile t
     
Exit Sub
hell: MsgBox Err.Description
End Sub

Private Sub mnuDeleteSelected_Click()
    Dim li As ListItem
    Dim f As String
    On Error Resume Next
    
    Const msg As String = "Are you sure you want to delete these files?"
    If MsgBox(msg, vbYesNo + vbInformation) = vbNo Then Exit Sub
    
    
nextone:
    For Each li In lv.ListItems
        If li.selected Then
            f = path & "\" & li.text
            If fso.FileExists(f) Then
                Kill f
            End If
            lv.ListItems.Remove li.index
            GoTo nextone
        End If
    Next
    
End Sub

Private Function LongestFileName() As Long
    Dim li As ListItem
    Dim r As Long
    
    For Each li In lv.ListItems
        If Len(li.text) > r Then r = Len(li.text)
    Next
    
    LongestFileName = r + 1
    
End Function

Private Sub mnuCopyTable_Click()

    Dim li As ListItem
    Dim t As String
    Dim ln As Long
    Dim sig As String
    Dim allNamedMD5 As Boolean
    
    ln = LongestFileName() + 2
    
    'if the file name contains the md5 and is not just a .txt file
    'if all files meet this criteria then copy table will not contain the md5 of the files since duplicate info
    'this is common if the rename all to md5 option was used on a directory (which also produces rename_log.txt)
    'this allows some additions to filename without needing exact match..
    'this logic below is reused in copy table csv and just post processed.
    allNamedMD5 = True
    For Each li In lv.ListItems
        If InStr(1, li.text, li.subItems(2), vbTextCompare) < 1 Then
            If LCase(fso.GetExtension(li.text)) <> ".txt" Then
                allNamedMD5 = False
                Exit For
            End If
        End If
    Next
            
    For Each li In lv.ListItems
        sig = Empty
        If li.ForeColor <> &H80000008 And li.ForeColor <> vbBlack Then
            sig = " " & vbTab & IIf(li.ForeColor = vbBlue, " VALID", " INVALID") & " signature"
        End If
        'subitems2 = md5
        If allNamedMD5 Then
            t = t & "  " & rpad(li.text, ln) & vbTab & li.subItems(1) & vbTab & li.subItems(3) & sig & vbCrLf
        Else
            t = t & "  " & rpad(li.text, ln) & vbTab & li.subItems(1) & vbTab & li.subItems(2) & vbTab & li.subItems(3) & sig & vbCrLf
        End If
    Next
    
    Clipboard.Clear
    Clipboard.SetText t
    'MsgBox "Copy Complete", vbInformation
    
End Sub

Sub handleFile(f As String)
    Dim h  As String
    Dim li As ListItem
    Dim e, fs As Long
    Dim sz As String
    Dim v As SigResults
    
    On Error Resume Next
    
    fs = DisableRedir()
    
    If mnuUseSHA256.Checked Then
        h = LCase(hash.HashFile(f, 256))
    Else
        h = LCase(hash.HashFile(f))
    End If
    
    v = VerifyFileSignature(f) 'can be slow on large files but most viruses are small and they are our target workign set. option to disable?
    
    If mnuHumanReadableSizes.Checked Then
        sz = FileSize(f, False)
    Else
        sz = FileLen(f)
    End If
        
    RevertRedir fs
   
    Dim ss As String
    
    If Len(h) = 0 Then
        e = Split(hash.error_message, "-")
        e = Replace(e(UBound(e)), vbCrLf, Empty)
        h = "Error: " & e 'library error...can happen if filesize > maxlong i think? fix me eventually..
    End If
    
    Set li = lv.ListItems.Add(, , fso.FileNameFromPath(f))
    li.subItems(1) = pad(sz)
    li.subItems(2) = h
    li.subItems(3) = GetCompileDateOrType(f)
    li.Tag = f
    
    If isSigned(v) Then
         SetLiColor li, SigToColor(v)
         li.ToolTipText = SigToStr(v)
    End If
    
End Sub



Private Sub mnuGoogleSelected_Click()
    On Error Resume Next
    Dim hashs() As String
    Dim li As ListItem
    Dim h As String
    Dim i As Long
    Dim x
    
    For Each li In lv.ListItems
        If li.selected Then
            h = li.subItems(2)
            If Len(h) > 0 And InStr(h, "Error") < 1 Then
                push hashs, li.subItems(2)
                i = i + 1
            End If
        End If
    Next

    If i = 0 Then
        MsgBox "No items were selected!", vbInformation
        Exit Sub
    End If
    
    For Each x In hashs
        Google CStr(x), Me.hwnd
    Next
    
End Sub

Private Sub mnuHashDiff_Click()
    
    Dim pth2 As String, tmp As String
    Dim f As frmHash
    Dim li As ListItem
    Dim results As String
    Dim r As String, report As String
    Dim totalHits As Long
    Dim unique2 As String
    Dim unique2_hits As Long
    
    Dim unique1 As String
    Dim unique1_hits As Long
    
    pth2 = dlg.FolderDialog2()
    If Len(pth2) = 0 Then Exit Sub
    
    Set f = New frmHash
    f.Visible = True
    f.Left = Me.Left + 300
    f.top = Me.top + 300
    f.Refresh
    DoEvents
    
    f.HashDir pth2, True

    Dim hashs As New Collection
    Dim hashs2 As New Collection
    Dim h
    
    'build a unique list of hashs in base directory set..
    For Each li In lv.ListItems
        h = li.subItems(2)
        If Not KeyExistsInCollection(hashs, CStr(h)) Then
            hashs.Add h, CStr(h)
        End If
     Next
     
     'build a unique list of hashs in compare directory set..
     For Each li In f.lv.ListItems
        h = li.subItems(2)
        If Not KeyExistsInCollection(hashs2, CStr(h)) Then
            hashs2.Add h, CStr(h)
        End If
     Next
     
    For Each h In hashs
        results = f.GetFilesForHash(h, totalHits)
        If Len(results) > 0 Then
            r = r & h & vbTab & results & vbCrLf
        Else
            unique1 = unique1 & h & vbTab & GetFilesForHash(h, , False) & vbCrLf
            unique1_hits = unique1_hits + 1
        End If
    Next
    
    'now we find the files from the second directory not found in main compare dir..
    For Each h In hashs2
        results = GetFilesForHash(h)
        If Len(results) = 0 Then
            unique2 = unique2 & h & vbTab & f.GetFilesForHash(h, , False) & vbCrLf
            unique2_hits = unique2_hits + 1
        End If
    Next
    
    If Len(r) > 0 Then
        
        report = "Base Dir:    " & path & " " & vbTab & lv.ListItems.Count & " files / " & hashs.Count & " unique" & vbCrLf & _
                 "Compare Dir: " & f.path & " " & vbTab & f.lv.ListItems.Count & " files / " & hashs2.Count & " unique" & vbCrLf & vbCrLf & _
                 "Files found in both sets:  " & totalHits & vbCrLf & vbCrLf & _
                 r & vbCrLf & vbCrLf & _
                 "Files unique to base dir: " & unique1_hits & " files" & vbCrLf & vbCrLf & _
                 unique1 & vbCrLf & vbCrLf & _
                 "Files unique to compare dir: " & unique2_hits & " files" & vbCrLf & vbCrLf & _
                 unique2
                
        tmp = fso.GetFreeFileName(Environ("temp"))
        fso.WriteFile tmp, report
        Shell "notepad.exe " & tmp, vbNormalFocus
    Else
        MsgBox "There were no hash matches in these two sample sets.", vbInformation
    End If
    
    On Error Resume Next
    'Unload f
     
End Sub

Private Sub mnuMakeExtSafe_Click()
     On Error Resume Next
    Dim li As ListItem
    Dim pdir As String
    Dim i As Long
    
    For Each li In lv.ListItems
        i = 1
        fPath = li.Tag
        fname = li.text
        pdir = fso.GetParentFolder(fPath) & "\"
        h = fname & "_"
        
        If LCase(VBA.Right(fname, 4)) = ".txt" Then GoTo nextone  'txt files are fine..
        If InStr(fname, ".") < 1 Then GoTo nextone                'no extension
        If VBA.Right(fname, 1) = "_" Then GoTo nextone            'already safe
        
        While fso.FileExists(pdir & h) 'dont delete dups, but append counter onto end..
            h = fname & "_" & i
            i = i + 1
        Wend
        
        Name fPath As pdir & h
    
        li.text = h
        li.Tag = pdir & h
        li.EnsureVisible
        'lv.Refresh
        DoEvents
        
nextone:
    Next
   
End Sub

Private Sub mnuMakeSubFolders_Click()
    
    On Error Resume Next
    Dim li As ListItem
    Dim pdir As String
    Dim baseName As String
    Dim fPath As String
    
    For Each li In lv.ListItems
        fPath = li.Tag
        fname = li.text
        pdir = fso.GetParentFolder(fPath) & "\"
        baseName = fso.GetBaseName(fPath)
        MkDir pdir & baseName
    Next
        
End Sub

Private Sub mnuRenameToMD5_Click()
    
    On Error Resume Next
    Dim li As ListItem
    Dim pdir As String
    Dim i As Long
    Dim rlog As String
    
    If MsgBox("Are you sure you want to rename all of these files to their MD5 hash values?", vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    For Each li In lv.ListItems
        i = 2
        fPath = li.Tag
        fname = li.text
        h = li.subItems(2)
        pdir = fso.GetParentFolder(fPath) & "\"
        
        If InStr(h, "Error") >= 1 Then GoTo nextone
        If LCase(fname) = LCase(h) Then GoTo nextone
        While fso.FileExists(pdir & h) 'dont delete dups, but append counter onto end..
            h = li.subItems(2) & "_" & i
            i = i + 1
        Wend
        
        rlog = rlog & fname & vbTab & "->" & vbTab & h & vbCrLf
        Name fPath As pdir & h
    
        li.text = h
        li.Tag = pdir & h
        li.EnsureVisible
        'lv.Refresh
        DoEvents
        
nextone:
    Next
        
    fso.WriteFile pdir & "\rename_log.txt", rlog
    
End Sub

Private Sub mnuSaveTable_Click()
    
    Dim pdir As String
    Dim ppdir As String
    Dim defName As String
    Dim fname As String
    Dim dat As String
    Dim li As ListItem
    
    On Error Resume Next
    
    defName = fso.GetBaseName(path) & "_hashs.txt"
    pdir = fso.GetParentFolder(path)
    ppdir = fso.GetParentFolder(pdir)
    
    'fname = dlg.SaveDialog(AllFiles, pdir, "Save output as", True, Me.hwnd, defName)
    'If Len(fname) = 0 Then Exit Sub
    
    fname = pdir & "\" & defName
    
    mnuCopyTable_Click
    
    dat = vbCrLf & Now() & " Directory hashs for " & lv.ListItems.Count & " files in: " & _
            Replace(path, ppdir, Empty) & _
            vbCrLf & vbCrLf & Clipboard.GetText
    
    fso.WriteFile fname, dat
    
End Sub

Private Sub mnuSubmitSelToVT_Click()
    On Error Resume Next
    Dim paths() As String
    Dim li As ListItem
    Dim i As Long
    Dim f As String
    
    For Each li In lv.ListItems
        If li.selected Then
            f = path & "\" & li.text
            If fso.FileExists(f) Then
                push paths, f
                i = i + 1
            End If
        End If
    Next

    If i = 0 Then
        MsgBox "No items were selected!", vbInformation
        Exit Sub
    End If
    
    If i = 1 Then
        Shell App.path & "\virustotal.exe ""/submit " & paths(0) & """", vbNormalFocus
    Else
        Clipboard.Clear
        Clipboard.SetText Join(paths, vbCrLf)
        Shell App.path & "\virustotal.exe /submitbulk", vbNormalFocus
    End If
    
End Sub

Private Sub mnuVTAll_Click()

    On Error Resume Next
    Dim li As ListItem
    Dim t As String
    
    For Each li In lv.ListItems
        If InStr(h, "Error") < 1 Then
             't = t & li.SubItems(2) & vbCrLf
             t = t & li.subItems(2) & "," & path & "\" & li.text & vbCrLf 'new format hash,path
        End If
    Next
    
    If Len(t) = 0 Then Exit Sub
    
    Clipboard.Clear
    Clipboard.SetText t
    Shell App.path & "\virustotal.exe /bulk", vbNormalFocus
    
End Sub

Private Sub mnuVTLookupSelected_Click()
    On Error Resume Next
    Dim hashs() As String
    Dim li As ListItem
    Dim h As String
    Dim i As Long
    
    For Each li In lv.ListItems
        If li.selected Then
            h = li.subItems(2)
            If Len(h) > 0 And InStr(h, "Error") < 1 Then
                'push hashs, li.SubItems(2)
                push hashs, li.subItems(2) & "," & path & "\" & li.text & vbCrLf 'new format hash,path
                i = i + 1
            End If
        End If
    Next

    If i = 0 Then
        MsgBox "No items were selected!", vbInformation
        Exit Sub
    End If
    
    If i = 1 Then
        'will this allow submit to work? is it path or hash?
        Shell App.path & "\virustotal.exe """ & lv.SelectedItem.Tag & """", vbNormalFocus
    Else
        Clipboard.Clear
        Clipboard.SetText Join(hashs, vbCrLf)
        Shell App.path & "\virustotal.exe /bulk", vbNormalFocus
    End If
    
End Sub


