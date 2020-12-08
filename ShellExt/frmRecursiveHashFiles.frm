VERSION 5.00
Begin VB.Form frmRecursiveHashFiles 
   Caption         =   "Hash all files below "
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   14460
   LinkTopic       =   "frmRecursiveHashFiles"
   ScaleHeight     =   5895
   ScaleWidth      =   14460
   StartUpPosition =   2  'CenterScreen
   Begin ShellExt.ucFilterList lv 
      Height          =   5280
      Left            =   45
      TabIndex        =   3
      Top             =   450
      Width           =   14235
      _ExtentX        =   25109
      _ExtentY        =   9313
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hash Files Below"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11925
      TabIndex        =   1
      Top             =   0
      Width           =   2490
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   900
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Text            =   "drop here"
      Top             =   45
      Width           =   10995
   End
   Begin VB.Label Label1 
      Caption         =   "Folder"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      TabIndex        =   2
      Top             =   45
      Width           =   1365
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Begin VB.Menu mnuCopyHashs 
         Caption         =   "Copy Hashs"
      End
      Begin VB.Menu mnuCopyCSV 
         Caption         =   "Copy CSV Results"
      End
      Begin VB.Menu mnuCopyReport 
         Caption         =   "Copy for Report"
      End
      Begin VB.Menu mnuStringsForAll 
         Caption         =   "Strings for All"
      End
      Begin VB.Menu mnuSafeExt 
         Caption         =   "Make all Ext Safe"
      End
      Begin VB.Menu mnuFlattenToHash 
         Caption         =   "Flatten to Hash"
      End
   End
End
Attribute VB_Name = "frmRecursiveHashFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fso As New CFileSystem2
Dim hash As New CWinHash

Sub RecursiveHashDir(folder As String)
    Text1 = folder
    Me.Visible = True
    Command1_Click
End Sub

Private Sub Command1_Click()
    
    On Error Resume Next
    Dim useSha256 As Boolean
    
    If Not fso.FolderExists(Text1) Then
        MsgBox "No folder found"
        Exit Sub
    End If
    
    Dim ff() As String, f, li As ListItem
    ff() = fso.GetFolderFiles(Text1, , , True)
    
    lv.ListItems.Clear
    useSha256 = CBool(GetMySetting("mnuUseSHA256.Checked", 0)) 'set in frmHash (hashDir)
    
    For Each f In ff
        Set li = lv.ListItems.Add()
        li.text = hash.HashFile(CStr(f), IIf(useSha256, 256, md5))
        li.subItems(1) = pad(Hex(FileLen(CStr(f))))
        li.subItems(2) = f
        DoEvents
    Next
    
    Me.Caption = "Hash all files below: " & lv.ListItems.Count & " files"
    
End Sub

 

Private Sub Form_Resize()
    On Error Resume Next
    lv.Width = Me.Width - lv.Left - 300
    lv.Height = Me.Height - lv.top - 500
    'sizeLvCol lv
End Sub

Private Sub lv_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = 2 Then PopupMenu mnuPopup
End Sub

Private Sub mnuCopyHashs_Click()
    Dim li As ListItem
    Dim X()
    
    push X, "hashs for: " & Text1 & vbCrLf
    
    For Each li In lv.ListItems
        push X, li.text
    Next
    
    Clipboard.Clear
    Clipboard.SetText Join(X, vbCrLf)
    
End Sub

Private Sub mnuCopyCSV_Click()
    Dim li As ListItem
    Dim X()
    
    push X, "hash,hexSize,path"
    
    For Each li In lv.ListItems
        push X, li.text & "," & li.subItems(1) & "," & li.subItems(2)
    Next
    
    Clipboard.Clear
    Clipboard.SetText Join(X, vbCrLf)
End Sub

Private Sub Form_Load()
    mnuPopup.Visible = False
    lv.SetColumnHeaders "Hash,Size,File*", "4000,1500"
    lv.SetFont "Courier", 12
    'sizeLvCol lv
End Sub

Private Sub mnuCopyReport_Click()
    Dim li As ListItem
    Dim X()
    Dim mbr As VbMsgBoxResult
    
    mbr = MsgBox("One item per line?", vbYesNo)
    
    For Each li In lv.ListItems
        If mbr = vbYes Then
            push X, pad(li.subItems(1), 6) & "   " & li.text & "   " & li.subItems(2)
        Else
            push X, rpad("File: ") & li.subItems(2)
            push X, rpad("Size: ") & li.subItems(1)
            push X, rpad("MD5: ") & li.text & vbCrLf
        End If
    Next
    
    Clipboard.Clear
    Clipboard.SetText Join(X, vbCrLf)
End Sub

Private Sub mnuFlattenToHash_Click()

    On Error Resume Next
    
    Dim li As ListItem
    Dim pdir As String, fPath As String, fname As String, h As String
    Dim i As Long, c As New CollectionEx, hits() As Long, hadDups As Boolean
    Dim tmp() As String, useSHA As Boolean, sz As String
    
    useSHA = GetMySetting("mnuUseSHA256.Checked", 0)
    
    'pdir = fso.GetParentFolder(Text1) & "\flat" 'currently must exist to recommend...
    pdir = dlg.FolderDialog2(Text1)
    If Len(pdir) = 0 Then Exit Sub
    
    pdir = pdir & "\"
    ReDim hits(lv.ListItems.Count)
    
    For Each li In lv.ListItems
        fPath = li.subItems(2)
        
        If useSHA Then
            h = LCase(hash.HashFile(fPath, 256))
        Else
            h = LCase(hash.HashFile(fPath))
        End If
        
        If c.keyExists(h) Then
            i = c.indexForKey(h)
            hits(i) = hits(i) + 1
        Else
            c.Add fPath, h
            hits(c.Count) = 1
            FileCopy fPath, pdir & h
        End If

        DoEvents

    Next
    
    push tmp, "Hits,Size,File,Hash,Type"
    For i = 0 To UBound(hits)
        fPath = c(i)
        'sz = FileLen(fPath)
        sz = FileSize(fPath, False)  'human readable size in mb,kb
        push tmp, pad(hits(i), 4) & ", " & pad(sz, 10) & ",  " & fso.FileNameFromPath(fPath) & ",  " & c.keyForIndex(i) & ",  " & GetCompileDateOrType(fPath)
    Next
    'fPath = fso.GetFreeFileName(Environ("temp"))
    fPath = pdir & "\index.txt"
    fso.WriteFile fPath, Join(tmp, vbCrLf)
    Shell "notepad.exe """ & fPath & """", vbNormalFocus
    
End Sub

Private Sub mnuSafeExt_Click()
    
    On Error Resume Next
    
    Dim li As ListItem
    Dim pdir As String, fPath As String, fname As String, h As String
    Dim i As Long
    
    For Each li In lv.ListItems
        i = 1
        fPath = li.subItems(2)
        fname = fso.FileNameFromPath(fPath)
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
    
        li.subItems(2) = pdir & h
        li.EnsureVisible
        'lv.Refresh
        DoEvents
        
nextone:
    Next
    
End Sub

Private Sub mnuStringsForAll_Click()
    
    On Error Resume Next
    Dim li As ListItem
    Dim f As String
    Dim n As Long
    Dim e As Long
    
    abort = False
    
    For Each li In lv.ListItems
        'If abort Then Exit For
        li.EnsureVisible
        li.selected = True
        Err.Clear
        If VBA.Left(li.text, 4) <> "str_" Then
            f = li.subItems(2)
            If fso.FileExists(f) Then
                frmStrings.ParseFile f
                frmStrings.AutoSave
                n = n + 1
            End If
        End If
        li.selected = False
        If Err.Number <> 0 Then e = e + 1
        DoEvents
    Next
    
    Unload frmStrings
    
    MsgBox n & " string dumps generated" & vbCrLf & "Errors: " & e, vbInformation
    
End Sub

Private Sub Text1_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
    On Error Resume Next
    If Not fso.FolderExists(data.Files(1)) Then Exit Sub
    If Err.Number <> 0 Then Exit Sub
    Text1 = data.Files(1)
End Sub





Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    X = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub
