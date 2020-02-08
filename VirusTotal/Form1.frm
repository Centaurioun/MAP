VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   Caption         =   "Bulk Hash Lookup"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11910
   LinkTopic       =   "Form2"
   ScaleHeight     =   7770
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtLimit 
      Height          =   285
      Left            =   8520
      TabIndex        =   13
      Text            =   "-1"
      Top             =   120
      Width           =   675
   End
   Begin VB.CheckBox chkSearchAll 
      Caption         =   "search all"
      Height          =   195
      Left            =   1440
      TabIndex        =   11
      Top             =   4635
      Width           =   1095
   End
   Begin VB.TextBox txtFilter 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2610
      TabIndex        =   9
      Top             =   4590
      Width           =   8385
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   10485
      Top             =   495
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   285
      Left            =   7200
      TabIndex        =   8
      Top             =   120
      Width           =   645
   End
   Begin VB.TextBox txtCacheDir 
      Height          =   285
      Left            =   3360
      TabIndex        =   7
      Text            =   "C:\VT_Cache\"
      Top             =   90
      Width           =   3675
   End
   Begin VB.CheckBox chkCache 
      Caption         =   "Cache Reports"
      Height          =   285
      Left            =   1710
      TabIndex        =   6
      Top             =   90
      Width           =   1545
   End
   Begin VB.Timer tmrDelay 
      Enabled         =   0   'False
      Interval        =   4500
      Left            =   8700
      Top             =   540
   End
   Begin VB.CommandButton cmdAbort 
      Caption         =   "Abort"
      Height          =   405
      Left            =   9480
      TabIndex        =   5
      Top             =   60
      Width           =   1455
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   285
      Left            =   60
      TabIndex        =   4
      Top             =   1740
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "Begin Query"
      Height          =   405
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   1605
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2715
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   4980
      Width           =   10965
   End
   Begin MSComctlLib.ListView lv 
      Height          =   2445
      Left            =   30
      TabIndex        =   1
      Top             =   2040
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   4313
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "hash"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "detects"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "subs"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "firstSeen"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "description"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   30
      TabIndex        =   0
      Top             =   510
      Width           =   10935
   End
   Begin VB.Label Label1 
      Caption         =   "Limit"
      Height          =   255
      Left            =   8040
      TabIndex        =   12
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "CSV Line Filters:"
      Height          =   240
      Left            =   90
      TabIndex        =   10
      Top             =   4635
      Width           =   1230
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Tools"
      Begin VB.Menu mnuBulkDownload 
         Caption         =   "Bulk Download"
      End
      Begin VB.Menu mnuSearchVT 
         Caption         =   "Search VT"
      End
      Begin VB.Menu mnuSpacer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUsePrivateKey 
         Caption         =   "Use Private API Key"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Begin VB.Menu mnuCopyTable 
         Caption         =   "Copy Table"
      End
      Begin VB.Menu mnuCopyResult 
         Caption         =   "Copy Result"
      End
      Begin VB.Menu mnuCopyAll 
         Caption         =   "Copy All Results"
      End
      Begin VB.Menu mnuCopyAllProps 
         Caption         =   "Copy All w/FileProps"
      End
      Begin VB.Menu mnuMoveSelected 
         Caption         =   "Move Selected"
      End
      Begin VB.Menu mnuSaveReports 
         Caption         =   "Save Reports to Files"
      End
      Begin VB.Menu mnuDivider 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClearList 
         Caption         =   "Remove All"
      End
      Begin VB.Menu mnuRemoveSelected 
         Caption         =   "Remove Selected"
      End
      Begin VB.Menu mnuRemoveUnsel 
         Caption         =   "Remove Unselected"
      End
      Begin VB.Menu mnuPrune 
         Caption         =   "Remove No Detections"
      End
      Begin VB.Menu mnuspacer55 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "Search All"
      End
      Begin VB.Menu mnuRescanSelected 
         Caption         =   "Rescan Selected"
      End
      Begin VB.Menu mnuDivider2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLoadH 
         Caption         =   "Load Hashs from"
         Begin VB.Menu mnuAddHashsFile 
            Caption         =   "Text File"
         End
         Begin VB.Menu mnuAddHashs 
            Caption         =   "Clipboard"
         End
         Begin VB.Menu mnuHashForDir 
            Caption         =   "Directory of Files"
         End
      End
      Begin VB.Menu mnuViewRaw 
         Caption         =   "View raw JSON"
      End
      Begin VB.Menu mnuSpacer4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSubmitSelected 
         Caption         =   "Submit Selected Files"
      End
      Begin VB.Menu mnuSPacer5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClearCache 
         Caption         =   "Clear Cache"
      End
      Begin VB.Menu mnuClearSelectedFromCache 
         Caption         =   "Clear Selected From Cache"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vt As New CVirusTotal
Dim selli As ListItem
Dim scan As CScan
Dim dlg As New CCmnDlg
'Dim fso As New CFileSystem2

Dim humanReadableSizes As Boolean
Dim files As New Collection 'of cfile
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

'current subitems:
    'li.subItems(1) = Right("   " & scan.positives, 3)
    'li.subItems(2) = Right("   " & scan.times_submitted, 3)
    'li.subItems(3) = lpad(scan.CodeSize, 12)
    'li.subItems(4) = scan.scan_date
    'li.subItems(5) = scan.first_seen
    'li.subItems(6) = scan.verbose_msg
            
Private Sub cmdAbort_Click()
    vt.abort = True
End Sub

Private Sub cmdClear_Click()
    List1.Clear
    lv.ListItems.Clear
    Text2 = Empty
    Set selli = Nothing
    Set files = New Collection
End Sub

Private Sub cmdBrowse_Click()
    Dim f As String
    f = dlg.FolderDialog2(txtCacheDir)
    If Len(f) = 0 Then Exit Sub
    txtCacheDir = f
End Sub


Sub cmdQuery_Click()
    Dim limit As Long
    Dim report As String
    Dim detections As Long
    Dim li As ListItem
    Dim scan As CScan
    Dim pth As String
    
    On Error Resume Next
    
    limit = CLng(txtLimit)
    If limit < 1 Then limit = 0
    
    If lv.ListItems.count = 0 Then
        MsgBox "Load hashs first!"
        Exit Sub
    End If
    
    vt.report_cache_dir = Empty
    
    If chkCache.value = 1 Then        'currently sets cache_dir to exist or not once at start of sub, cant change during operation..
        If Len(txtCacheDir) = 0 Then
            chkCache.value = 0
        Else
            If Not fso.FolderExists(txtCacheDir) Then
                If Not fso.buildPath(txtCacheDir) Then
                    chkCache.value = 0
                End If
            End If
        End If
        If fso.FolderExists(txtCacheDir) Then vt.report_cache_dir = txtCacheDir
    End If
        
    
    vt.abort = False
    pb.value = 0
    pb.Max = lv.ListItems.count
    vt.delayInterval = IIf(pb.Max < 5, 2500, 17300) 'cant exceed 4 requests per minute...
    List1.AddItem "Max: " & pb.Max & " Interval: " & vt.delayInterval
    
    If Not vt.usingPrivateKey Then
        lv.ColumnHeaders(2).Width = 0 'times submitted 'codesize(3) avail to all?
        lv.ColumnHeaders(5).Width = 0 'first seen only available w/ priv key
    End If
    
    For Each li In lv.ListItems
    
        If vt.abort Then Exit For
        
        If limit > 0 Then
            If pb.value >= limit Then Exit For
        End If
        
        If Len(Trim(li.Text)) = 0 Then GoTo nextone
        
        Set scan = vt.GetReport(li.Text)
        
        pth = PathForHash(li.Text)
        If Len(pth) > 0 Then scan.LocalFilePath = pth
    
        If Not scan.HadError Then
            li.subItems(1) = Right("   " & scan.positives, 3)
            li.subItems(2) = Right("   " & scan.times_submitted, 3)
            li.subItems(3) = lpad(scan.CodeSize, 12)
            li.subItems(4) = scan.scan_date
            li.subItems(5) = scan.first_seen
            li.subItems(6) = scan.verbose_msg
            Set li.Tag = scan
        Else
            li.subItems(1) = "Failure"
            li.subItems(2) = Empty
            li.subItems(3) = Empty
            Set li.Tag = Nothing
'            Set vt = New CVirusTotal
'            Set vt.Timer1 = tmrDelay
'            Set vt.winInet = Inet1
'            Set vt.debugLog = List1
        End If
        
        li.EnsureVisible
        DoEvents
        Me.Refresh
        pb.value = pb.value + 1
        Me.Caption = pb.value & "/" & pb.Max
        If pb.value = lv.ListItems.count Then GoTo nextone
        
nextone:
    Next
    
    lv.ListItems(1).EnsureVisible
    lv.ListItems(1).Selected = True
    lv_ItemClick lv.ListItems(1)
    pb.value = 0
    'MsgBox "Queries Complete" & vbCrLf & vbcrllf & "Click on an item to view report.", vbInformation
 
 
End Sub

Function lpad(x, Optional cnt = 8) As String
    
    If Len(x) >= cnt Then
        lpad = x
    Else
        lpad = String(cnt - Len(x), " ") & x
    End If
    
End Function


Private Sub Command2_Click()
'    Dim m_json As String
'    Dim b As String
'    Dim c As Long
'    Dim d As Dictionary
'
'    If AddComment(txtHash, Text2, m_json, b, c) Then
'        Set d = JSON.parse(m_json)
'        If d Is Nothing Then
'            List1.AddItem "AddComment JSON parsing error"
'            Exit Sub
'        End If
'
'        If JSON.GetParserErrors <> "" Then
'            List1.AddItem "AddComment Json Parse Error: " & JSON.GetParserErrors
'            Exit Sub
'        End If
'
'        If d.Item("response") <> 1 Then
'            MsgBox d.Item("verbose_msg")
'            Exit Sub
'        End If
'
'        MsgBox "Comment was added successfully", vbInformation
'
'    Else
'        MsgBox "Add Comment Failed Status code: " & c & " " & b
'    End If
        
   
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    abort = True
    SaveSetting "vt", "settings", "cachedir", txtCacheDir.Text
    SaveSetting "vt", "settings", "usecache", chkCache.value
End Sub

Sub mnuAddHashs_Click()
    On Error Resume Next
    Dim f As CFile
    Dim tmp
    Dim useLF As Boolean
    
    x = Clipboard.GetText
    x = Replace(x, " ", Empty)
    x = Replace(x, vbTab, Empty)
    If InStr(Command, "/bulk") < 1 Then x = Replace(x, ",", Empty) '/bulk command line from shellext uses hash,path\r\n format
    x = Replace(x, "'", Empty)
    x = Replace(x, """", Empty)
    x = Replace(x, ";", Empty)
    x = Replace(x, "}", Empty)
    x = Replace(x, ")", Empty)
    
    If Len(x) > 1000 Then
        If InStr(Mid(x, 1, 1000), vbCrLf) < 1 Then useLF = True
    Else
        If InStr(x, vbCrLf) < 1 Then useLF = True
    End If
    
    tmp = Split(x, IIf(useLF, vbLf, vbCrLf))
    pb.Max = UBound(tmp)
    pb.value = 0
    For Each x In tmp
        If InStr(x, ":") > 0 And InStr(x, ",") < 1 Then
            x = Split(x, ":")(1) 'its from yara match output? sigName:hash
        End If
        x = Trim(x)
        If Len(x) > 0 Then
            If InStr(x, ",") > 0 Then 'new "hash,path" format (is hash always md5? probably...)
                y = Split(x, ",")
                Set f = New CFile
                f.hash = y(0)
                f.path = y(1)
                lv.ListItems.Add , , f.hash
                If fso.FileExists(f.path) Then files.Add f
            Else
                lv.ListItems.Add , , x
            End If
        End If
        pb.value = pb.value + 1
    Next
    
    pb.value = 0
    Me.Caption = lv.ListItems.count & " hashs added"
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    List1.Width = Me.Width - List1.Left - 200
    Text2.Width = List1.Width
    lv.Width = List1.Width
    pb.Width = List1.Width
    Text2.Height = Me.Height - Text2.Top - 700
    txtFilter.Width = Me.Width - txtFilter.Left - 200
End Sub

Public Sub LV_ColumnSort(ListViewControl As ListView, Column As ColumnHeader)
     On Error Resume Next
    With ListViewControl
       If .SortKey <> Column.index - 1 Then
             .SortKey = Column.index - 1
             .SortOrder = lvwAscending
       Else
             If .SortOrder = lvwAscending Then
              .SortOrder = lvwDescending
             Else
              .SortOrder = lvwAscending
             End If
       End If
       .Sorted = -1
    End With
End Sub

Function lvGetAllElements(lv As Object, Optional divider As String = vbTab) As String
    Dim ret() As String, i As Integer, tmp As String
    Dim li 'As ListItem
    
    On Error Resume Next
    
    For i = 1 To lv.ColumnHeaders.count
        tmp = tmp & lv.ColumnHeaders(i).Text & divider
    Next
    
    push ret, tmp
    push ret, String(50, "-")
        
    For Each li In lv.ListItems
        tmp = Trim(li.Text) & divider
        For i = 1 To lv.ColumnHeaders.count - 1
            tmp = tmp & " " & Trim(li.subItems(i)) & divider
        Next
        If Right(tmp, Len(divider)) = divider Then tmp = Left(tmp, Len(tmp) - Len(divider))
        push ret, tmp
    Next
    
    lvGetAllElements = Join(ret, vbCrLf)
    
End Function

Function IsIde() As Boolean
' Brad Martinez  http://www.mvps.org/ccrp
    On Error GoTo out
    Debug.Print 1 / 0
out: IsIde = Err
End Function

Sub errorStartup()
    List1.AddItem "Designed to be run from right click menus in explorer."
    List1.AddItem "You can add bulk hash lists to lookup by right click on listview"
    List1.AddItem "If you have a private api key set, you can use Options->Bulk Download"
    Me.Show
End Sub

Private Sub Form_Load()
    
    On Error Resume Next
    
    Dim path As String
    Dim hash_mode As Boolean
    
    mnuUsePrivateKey.Checked = vt.usingPrivateKey
    
    If Not IsIde Then
        mnuBulkDownload.Enabled = vt.usingPrivateKey
        mnuSearchVT.Enabled = vt.usingPrivateKey
    End If
    
    mnuPopup.Visible = False
    
    Set vt.Timer1 = tmrDelay
    Set vt.winInet = Inet1
    Set vt.debugLog = List1
    
    txtCacheDir = GetSetting("vt", "settings", "cachedir", App.path & "\VT_Cache")
    chkCache.value = GetSetting("vt", "settings", "usecache", 0)
    
    lv.ColumnHeaders(5).Width = lv.Width - lv.ColumnHeaders(5).Left - 150
    
    mnuMoveSelected.Visible = (files.count > 0)
    
    Exit Sub

End Sub


Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    Set selli = Item
    Set scan = Item.Tag
    If scan Is Nothing Then Exit Sub
    If Len(txtFilter) > 0 Then
        If chkSearchAll.value = 1 Then
            chkSearchAll.value = 0 'this will trigger the filter change itself
        Else
            txtFilter_Change
        End If
    Else
        Text2 = scan.GetReport()
    End If
End Sub

Function GetKeyState(k As KeyCodeConstants) As Boolean
    If GetAsyncKeyState(k) Then GetKeyState = True
End Function

Private Sub lv_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    'LV_ColumnSort lv, ColumnHeader
    On Error Resume Next
    Dim li As ListItem, c As CScan
    
    If GetKeyState(vbKeyShift) Then
        If humanReadableSizes Then
            humanReadableSizes = False
            For Each li In lv.ListItems
                Set c = li.Tag
                li.subItems(3) = lpad(c.CodeSize, 12)
            Next
        Else
            humanReadableSizes = True
            For Each li In lv.ListItems
                Set c = li.Tag
                li.subItems(3) = lpad(FileSize(c.CodeSize, False))
            Next
        End If
    Else
        LV_ColumnSort lv, ColumnHeader
    End If
    
End Sub

Public Function FileSize(fpathOrSize, Optional showBytes As Boolean = True) As String
    Dim fSize As Long
    Dim szName As String
    On Error GoTo hell
    
    If fso.FileExists(CStr(fpathOrSize)) Then
        fSize = FileLen(fpathOrSize)
    Else
        fSize = CLng(fpathOrSize)
    End If
    
    If showBytes Then szName = " bytes"
    
    If fSize > 1024 Then
        fSize = fSize / 1024
        szName = " Kb"
    End If
    
    If fSize > 1024 Then
        fSize = fSize / 1024
        szName = " Mb"
    End If
    
    FileSize = fSize & szName
    
    Exit Function
hell:
    
End Function


Private Function PathForHash(hash As String) As String
    Dim f As CFile
    For Each f In files
        If f.hash = hash Then
            PathForHash = f.path
            Exit Function
        End If
    Next
End Function

Private Sub lv_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuPopup
End Sub


Private Sub mnuAddHashsFile_Click()
    On Error Resume Next
    Dim fp As String
    f = dlg.OpenDialog(AllFiles)
    If Len(f) = 0 Then Exit Sub
    Clipboard.Clear
    Clipboard.SetText fso.ReadFile(f)
    mnuAddHashs_Click
End Sub

Private Sub mnuBulkDownload_Click()
    frmBulkDownload.Show
    Unload Me
End Sub


Private Sub mnuClearCache_Click()
    
    Dim f() As String
    Dim ff
    
    If MsgBox("Are you sure you want to delete all cached reports? They take a while to query?", vbYesNo + vbInformation) = vbNo Then
        Exit Sub
    End If
    
    If fso.FolderExists(txtCacheDir) Then
        f() = fso.GetFolderFiles(txtCacheDir, "*.txt")
        For Each ff In f
            If Len(fso.FileNameFromPath(CStr(ff))) = 36 Then fso.DeleteFile CStr(ff)
        Next
    End If
    
End Sub

Private Sub mnuClearList_Click()
    cmdClear_Click
End Sub

Private Sub mnuClearSelectedFromCache_Click()
    Dim li As ListItem
    Dim fpath As String
    
    If MsgBox("Are you sure you want to delete all cached reports? They take a while to query?", vbYesNo + vbInformation) = vbNo Then
        Exit Sub
    End If
    
    If fso.FolderExists(txtCacheDir) Then
        For Each li In lv.ListItems
            If li.Selected Then
                fpath = txtCacheDir & "\" & li.Text & ".txt"
                If fso.FileExists(fpath) Then fso.DeleteFile fpath
                li.subItems(1) = Empty
                li.subItems(2) = Empty
                li.subItems(3) = Empty
            End If
        Next
    End If
    
End Sub

Private Sub mnuCopyAll_Click()
    Dim li As ListItem
    Dim r
    Dim scan As CScan
    
    On Error Resume Next
    
    r = lvGetAllElements(lv, ",")
    r = r & vbCrLf & vbCrLf
    
    For Each li In lv.ListItems
        Set scan = li.Tag
        r = r & scan.GetReport() & vbCrLf & String(60, "-") & vbCrLf & vbCrLf
    Next
    
    r = vbCrLf & "This is a temp file do a file SaveAs if you want to keep it." & vbCrLf & vbCrLf & r
    
    Dim tf As String
    tf = fso.GetFreeFileName(Environ("temp"))
    fso.writeFile tf, r
    Shell "notepad.exe """ & tf & """", vbNormalFocus
    
End Sub

Private Sub mnuCopyAllProps_Click()
    
    Dim li As ListItem
    Dim r
    Dim s As CScan
    Dim ret() As String, i As Integer, tmp() As String
    Dim fp As New CFileProperties
    Dim fs As New clsFileStream
    Dim report As String
    
    On Error Resume Next
    
    report = fso.GetFreeFileName(Environ("temp"))
    fs.fOpen report, otwriting
    fs.WriteLine vbCrLf & "This is a temp file saveAs to save"
    fs.WriteDivider
    
    For Each li In lv.ListItems
        fs.WriteLine "Hash: " & Trim(li.Text)
        
        For i = 2 To lv.ColumnHeaders.count
            fs.WriteLine lv.ColumnHeaders(i).Text & ": " & Trim(li.subItems(i - 1))
        Next
        
        Set s = li.Tag
        If Not s Is Nothing Then
            If fso.FileExists(s.LocalFilePath) Then
                fs.WriteLine "File: " & fso.FileNameFromPath(s.LocalFilePath)
                fs.WriteLine "Props: " & fp.FileInfo(s.LocalFilePath).asStr()
            End If
        End If
        
        fs.WriteDivider
    Next
    
    fs.fClose
    Shell "notepad.exe """ & report & """", vbNormalFocus
    
End Sub

Private Sub mnuCopyResult_Click()

    On Error Resume Next

    If selli Is Nothing Then
        MsgBox "Nothing selected"
        Exit Sub
    End If
    
    Dim r As String
    Dim scan As CScan
    Set scan = selli.Tag
    
    r = selli.Text & "  Detections: " & selli.subItems(1)
    If vt.usingPrivateKey Then r = r & "  Submissions: " & li.subItems(2)
    r = r & "  ScanDate: " & li.subItems(3)
    If vt.usingPrivateKey Then r = r & "  First Seen: " & li.subItems(4)
    r = r & vbCrLf & String(50, "-") & vbCrLf & scan.GetReport()
    
    Clipboard.Clear
    Clipboard.SetText r
    MsgBox Len(r) & " bytes copied to clipboard"
    
End Sub

Private Sub mnuCopyTable_Click()

    Dim li As ListItem
    Dim r
    Dim s As CScan
    Dim ret() As String, i As Integer, tmp As String
    
    On Error Resume Next
    
    For i = 1 To lv.ColumnHeaders.count
        tmp = tmp & lv.ColumnHeaders(i).Text & ","
    Next
    
    push ret, tmp & "file"
    push ret, String(50, "-")
        
    For Each li In lv.ListItems
        tmp = Trim(li.Text) & ","
        
        For i = 1 To lv.ColumnHeaders.count - 1
            tmp = tmp & " " & Trim(li.subItems(i)) & ","
        Next
        
        Set s = li.Tag
        If Not s Is Nothing Then
            If Len(s.LocalFilePath) > 0 Then
                tmp = tmp & " " & fso.FileNameFromPath(s.LocalFilePath)
            End If
        End If
        
        push ret, tmp
    Next
    
    r = Join(ret, vbCrLf)

    Clipboard.Clear
    Clipboard.SetText r
    MsgBox Len(r) & " bytes copied to clipboard"
    
End Sub

Private Sub mnuHashForDir_Click()
    On Error Resume Next
    Dim d As String, c() As String, x, isRecur As Boolean, li As ListItem, hash As New CWinHash, f As CFile
    d = dlg.FolderDialog2()
    If Len(d) = 0 Then Exit Sub
    If Not AryIsEmpty(fso.GetSubFolders(d)) Then
        isRecur = (MsgBox("Recursive search?", vbYesNo) = vbYes)
    End If
    c = fso.GetFolderFiles(d, , , isRecur)
    If AryIsEmpty(c) Then Exit Sub
    pb.Max = UBound(c)
    pb.value = 0
    For Each x In c
        Set f = New CFile
        f.hash = hash.HashFile(CStr(x))
        f.path = x
        lv.ListItems.Add , , f.hash
        files.Add f
        pb.value = pb.value + 1
    Next
    pb.value = 0
    Me.Caption = pb.Max & " hashs added"
End Sub



Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    Dim i
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function

Private Sub mnuMoveSelected_Click()
    
    On Error Resume Next
    
    Dim li As ListItem
    Dim fpath As String
    Dim newDir As String
    Dim fName As String
    Dim i As Long
    Dim moved As Long, total As Long
    
    fpath = fileFromHash(lv.ListItems(1).Text)
    If fso.FileExists(fpath) Then
         fpath = fso.GetParentFolder(fpath)
    Else
         fpath = Empty
    End If
                
    newDir = dlg.FolderDialog2(fpath)
    If Len(newDir) = 0 Then Exit Sub
    
    For i = lv.ListItems.count To 1 Step -1
        Set li = lv.ListItems(i)
        If li.Selected Then
            total = total + 1
            fpath = fileFromHash(li.Text)
            If fso.FileExists(fpath) Then
                fName = "\" & fso.FileNameFromPath(fpath)
                If Not fso.FileExists(newDir & fName) Then
                    fso.Move fpath, newDir
                    lv.ListItems.remove i
                    moved = moved + 1
                End If
            End If
        End If
    Next
    
    MsgBox moved & "/" & total & " files moved", vbInformation
    
End Sub

Function fileFromHash(hash, Optional remove As Boolean = True) As String
    Dim f As CFile
    For i = 1 To files.count
        Set f = files(i)
        If LCase(f.hash) = LCase(hash) Then
            fileFromHash = f.path
            If remove Then files.remove i
            Exit Function
        End If
    Next
End Function

Private Sub mnuPrune_Click()
    Dim li As ListItem
    On Error Resume Next
    For i = lv.ListItems.count To 1 Step -1
        Set li = lv.ListItems(i)
        If li.subItems(1) = "0" Then
            fileFromHash li.Text, True
            lv.ListItems.remove i
        End If
    Next
End Sub

Private Sub mnuRemoveSelected_Click()
    On Error Resume Next
    
    For i = lv.ListItems.count To 1 Step -1
        If lv.ListItems(i).Selected Then
            fileFromHash li.Text, True
            lv.ListItems.remove i
        End If
    Next
    
End Sub

Private Sub mnuRemoveUnsel_Click()
 On Error Resume Next
    
    For i = lv.ListItems.count To 1 Step -1
        If Not lv.ListItems(i).Selected Then
            fileFromHash li.Text, True
            lv.ListItems.remove i
        End If
    Next
    
End Sub

Private Sub mnuRescanSelected_Click()
    Dim li As ListItem
    Dim scan As CScan
    
    For Each li In lv.ListItems

        If li.Selected Then
            If Len(Trim(li.Text)) > 0 Then
            
                Set scan = vt.GetReport(li.Text)
                
                pth = PathForHash(li.Text)
                If Len(pth) > 0 Then scan.LocalFilePath = pth
                
                If Not scan.HadError Then
                    li.subItems(1) = Right("   " & scan.positives, 3)
                    li.subItems(2) = Right("   " & scan.times_submitted, 3)
                    li.subItems(3) = lpad(scan.CodeSize, 12)
                    li.subItems(4) = scan.scan_date
                    li.subItems(5) = scan.first_seen
                    li.subItems(6) = scan.verbose_msg
                    Set li.Tag = scan
                Else
                    li.subItems(1) = "Failure"
                    li.subItems(2) = Empty
                    li.subItems(3) = Empty
                    li.subItems(4) = Empty
                    li.subItems(5) = Empty
                    li.subItems(6) = Empty
                    Set li.Tag = Nothing
                    Set vt = New CVirusTotal
                End If
                
                li.EnsureVisible
                DoEvents
                
            End If
        End If
    Next
         
    
End Sub

Private Sub mnuSaveReports_Click()
    
    On Error Resume Next
    Dim li As ListItem
    Dim pf As String
    Dim scan As CScan
    Dim report As String
    
    pf = dlg.FolderDialog()
    If Len(pf) = 0 Then Exit Sub
    
    For Each li In lv.ListItems
        hash = li.Text
        Set scan = li.Tag
        
        report = "Hash: " & li.Text & vbCrLf & _
                 "Detections: " & li.subItems(1) & vbCrLf & _
                 "ScanDate: " & li.subItems(2) & vbCrLf
                 
        If vt.usingPrivateKey Then report = report & "  First Seen: " & li.subItems(3) & vbCrLf
        
        report = report & String(50, "-") & vbCrLf & vbCrLf & scan.GetReport()
                 
        fso.writeFile pf & "\VT_" & hash & ".txt", report
    Next

End Sub

Function StartsWith(blob, prefix) As Boolean
    Dim i As Long
    Dim tmp As String
    i = Len(prefix)
    If Len(blob) < i Then Exit Function
    tmp = LCase(Left(blob, i))
    If tmp = LCase(prefix) Then StartsWith = True
End Function

Private Sub mnuSearch_Click()
    Dim li As ListItem
    Dim likeSearch As Boolean, NotLikeSearch As Boolean
    Dim cs As CScan
    Dim found As Long
    Dim tmp() As String
    Dim r As String
    Dim lines()  As String
    Dim tmpFile As String
    Dim capt As String
    Dim tested As Long
    
    capt = " string: "
    find = InputBox("Enter marker to search for, to use vb like operator prefix with 'like:' or 'not like:'")
    If Len(find) = 0 Then Exit Sub
    
    If StartsWith(find, "like:") Then
        find = LCase(Trim(Mid(find, 6)))
        If InStr(find, "*") < 1 Then find = "*" & find & "*"
        likeSearch = True
        capt = " like " & find
    End If
    
    If StartsWith(find, "not like:") Then
        find = LCase(Trim(Mid(find, 10)))
        If InStr(find, "*") < 1 Then find = "*" & find & "*"
        NotLikeSearch = True
        capt = " not like " & find
    End If
    
    push tmp(), "Search for '" & find & "' " & lv.ListItems.count & " samples - " & Now & vbCrLf
    
    For Each li In lv.ListItems
        li.Selected = False
        If IsObject(li.Tag) Then
            tested = tested + 1
            Set cs = li.Tag
            r = LCase(cs.GetReport())
            If NotLikeSearch Then
                If Not r Like find Then
                    li.Selected = True
                    push tmp(), cs.GetReport()
                End If
            ElseIf likeSearch Then
                If r Like find Then
                    li.Selected = True
                    push tmp(), cs.extractDetectionsFor(find, True)
                End If
            Else
                If InStr(1, r, find, vbTextCompare) > 0 Then
                    li.Selected = True
                    push tmp(), cs.extractDetectionsFor(find)
                End If
            End If
            If li.Selected Then
                found = found + 1
                li.EnsureVisible
            End If
        End If
    Next
    
    Me.Caption = found & " matches found for " & capt & find
    
    If found > 0 Then
        tmp(0) = tmp(0) & "Found: " & found & " hits of " & tested & vbCrLf
        tmpFile = fso.GetFreeFileName(Environ("temp"))
        fso.writeFile tmpFile, Join(tmp, vbCrLf)
        Shell "notepad.exe """ & tmpFile & """", vbNormalFocus
    End If
    
End Sub

Private Sub mnuSearchVT_Click()
    frmSearch.Show
End Sub

Private Sub mnuSubmitSelected_Click()

    Dim li As ListItem
    Dim scan As CScan
    Dim pth As String
    
    List1.Clear
    List1.AddItem "Submitting selected files"
    
    For Each li In lv.ListItems

        If li.Selected Then
            If Len(Trim(li.Text)) > 0 Then
            
                pth = PathForHash(li.Text)
                If Len(pth) > 0 Then
                    Set scan = vt.SubmitFile(pth)
                    scan.response_code = 2 'manually overridden for getreport() display purposes..
                    li.subItems(1) = scan.verbose_msg
                    Set li.Tag = scan
                Else
                    List1.AddItem "No file path found for " & li.Text
                End If
                
                li.EnsureVisible
                DoEvents
                
            End If
        End If
    Next
    
End Sub

Private Sub mnuUsePrivateKey_Click()
    Dim x As String
    
    x = InputBox("By default we use a rate limited public API key. If you have access to a private api key, you may enter it here to avoid delays. " & _
                 "Enter an empty string or hit cancel to clear the private key." & vbCrLf & vbCrLf & "Your key will be stored in the registry.", _
                 "Enter private api key", _
                 vt.ReadPrivateApiKey _
        )
                 
    vt.SetPrivateApiKey x
    mnuUsePrivateKey.Checked = vt.usingPrivateKey
    
    If vt.usingPrivateKey Then
        MsgBox "Private key successfull set", vbInformation
        mnuBulkDownload.Enabled = True
        mnuSearchVT.Enabled = True
    Else
        MsgBox "You are now using the default public key which is rate limited and free for non-commercial use. " & vbCrLf & vbCrLf & "Please see the VirusTotal terms of service.", vbInformation
        mnuBulkDownload.Enabled = False
        mnuSearchVT.Enabled = False
    End If
    
End Sub

Private Sub mnuViewRaw_Click()
    On Error Resume Next
    If selli Is Nothing Then Exit Sub
    Dim scan As CScan
    Set scan = selli.Tag
    Text2 = scan.RawJson
End Sub


Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub
'
'Private Sub txtFilter_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 13 Then
'        txtFilter_Change
'        KeyCode = 0
'    End If
'End Sub

Private Sub txtFilter_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        txtFilter_Change
        KeyAscii = 0
    End If
End Sub

Private Sub chkSearchAll_Click()
    txtFilter_Change
End Sub

Private Sub txtFilter_Change()
    On Error Resume Next
    
    If chkSearchAll.value = 0 And scan Is Nothing Then
        Text2 = Empty
        Exit Sub
    End If
    
    If Len(txtFilter.Text) = 0 Then
        If scan Is Nothing Then
            Text2 = Empty
        Else
            Text1 = scan.GetReport()
        End If
        Exit Sub
    End If
    
    Dim ret() As String, r As String, hitCounter As Long, selCount As Long
    Dim li As ListItem
    Dim s As CScan
    
    If chkSearchAll.value = 1 Then
        For Each li In lv.ListItems
            Set s = li.Tag
            r = searchScan(s, txtFilter, hitCounter)
            If Len(r) Then
                push ret, vbCrLf & String(75, "-") & vbCrLf & vbCrLf & r
                selCount = selCount + 1
                li.Selected = True
            Else
                li.Selected = False
            End If
        Next
        Text2 = hitCounter & " hits across " & selCount & " samples for " & txtFilter & vbCrLf & vbCrLf & Join(ret, vbCrLf)
    Else
        Text2 = searchScan(scan, txtFilter)
    End If
    
End Sub

Private Function searchScan(s As CScan, csvText As String, Optional ByRef hitCounter As Long) As String
    
    Dim ret(), tmp() As String, x
    Dim matches() As String, m
    Dim hits As Long
    
    matches = Split(csvText, ",")
    tmp = Split(s.GetReport(), vbCrLf)
    
    'save file info
    For i = 0 To 4
        push ret, tmp(i)
    Next
    
    For Each x In tmp
        For Each m In matches
            If Len(m) > 0 Then
                If InStr(1, x, m, vbTextCompare) > 0 Then
                   push ret, x
                   hits = hits + 1
                   hitCounter = hitCounter + 1
                   Exit For
                End If
            End If
        Next
    Next
    
    If hits > 0 Then
        searchScan = Join(ret, vbCrLf)
    End If
    
End Function




