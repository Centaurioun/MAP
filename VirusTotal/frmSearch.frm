VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmSearch 
   Caption         =   "VirusTotal Search UI"
   ClientHeight    =   9030
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   17085
   LinkTopic       =   "frmSearch"
   ScaleHeight     =   9030
   ScaleWidth      =   17085
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDir 
      Height          =   285
      Left            =   4860
      OLEDropMode     =   1  'Manual
      TabIndex        =   13
      Top             =   540
      Width           =   2625
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   450
      Top             =   8415
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CheckBox chkGetReports 
      Caption         =   "Get Reports"
      Height          =   240
      Left            =   11745
      TabIndex        =   11
      Top             =   180
      Width           =   1770
   End
   Begin VB.CommandButton cmdAbort 
      Caption         =   "Abort"
      Height          =   375
      Left            =   15705
      TabIndex        =   10
      Top             =   90
      Width           =   1230
   End
   Begin VB.ComboBox cboLimit 
      Height          =   315
      Left            =   10170
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   135
      Width           =   1230
   End
   Begin VB.TextBox txtReport 
      Height          =   8025
      Left            =   10395
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   900
      Width           =   6585
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7560
      TabIndex        =   6
      Top             =   540
      Width           =   510
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   240
      Left            =   8325
      TabIndex        =   4
      Top             =   540
      Width           =   8610
      _ExtentX        =   15187
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   1
   End
   Begin VirusTotal.HistoryCombo hc 
      Height          =   375
      Left            =   270
      TabIndex        =   3
      Top             =   135
      Width           =   9150
      _ExtentX        =   16140
      _ExtentY        =   661
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   8505
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   180
      TabIndex        =   2
      Top             =   6300
      Width           =   10095
   End
   Begin VirusTotal.ucFilterList lv 
      Height          =   5325
      Left            =   180
      TabIndex        =   1
      Top             =   855
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   9393
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   375
      Left            =   14175
      TabIndex        =   0
      Top             =   90
      Width           =   1140
   End
   Begin VB.Label Label2 
      Caption         =   "Dl Dir:"
      Height          =   195
      Left            =   4320
      TabIndex        =   12
      Top             =   585
      Width           =   510
   End
   Begin VB.Label Label1 
      Caption         =   "Limit"
      Height          =   240
      Left            =   9585
      TabIndex        =   8
      Top             =   180
      Width           =   510
   End
   Begin VB.Label lblStatus 
      Height          =   285
      Left            =   225
      TabIndex        =   5
      Top             =   540
      Width           =   4020
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuClearHistory 
         Caption         =   "Clear History"
      End
      Begin VB.Menu mnuClearCacheVT 
         Caption         =   "Clear Cache VT Results"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Help"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Begin VB.Menu mnuGetReport 
         Caption         =   "Get Report"
      End
      Begin VB.Menu mnuDownloadSelected 
         Caption         =   "Download Selected"
      End
      Begin VB.Menu mnuDownloadAll 
         Caption         =   "Download All"
      End
      Begin VB.Menu mnuRawJSON 
         Caption         =   "View Raw JSON"
      End
      Begin VB.Menu mnuVisitPage 
         Caption         =   "Visit Page"
      End
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vt As New CVirusTotal
Dim txtCacheDir As String
Dim fso As New CFileSystem2
Dim dlg As New CCmnDlg
'Dim txtDir As String
Dim csr As CSearchResult
Dim selItem As ListItem

Private Sub cmdAbort_Click()
    vt.abort = True
End Sub

Private Sub cmdBrowse_Click()
    setDLDir
End Sub

Private Sub cmdMore_Click()
    If csr Is Nothing Then Exit Sub
    MsgBox "todo: search offset: " & csr.lastOffset
End Sub

Private Sub cmdSearch_Click()
     
     Dim c As CScan
     Dim li As ListItem
     Dim h
     Dim limit As Long
     
     If Len(Trim(hc.Text)) = 0 Then Exit Sub
     
     hc.RecordIfNew
     List1.Clear
     limit = CLng(cboLimit.Text)
     lblStatus = "Doing search..."
     Set csr = vt.Search(hc.Combo.Text, limit)
     
     lv.ListItems.Clear
     For Each h In csr.results
        Set li = lv.ListItems.Add()
        li.subItems(3) = h
     Next
     
     If csr.HadError Then
        lblStatus = "HadErr: " & csr.verbose_msg
        Exit Sub
     End If
     
     If lv.ListItems.count = 0 Then
        lblStatus = "No results"
        Exit Sub
     End If
     
     If chkGetReports.Value = 1 Then
        lblStatus = lv.ListItems.count & " hits - loading reports"
        
        pb.Value = 0
        pb.Max = lv.ListItems.count
        Me.Refresh
        For Each li In lv.ListItems
           If vt.abort Then Exit For
           Set c = vt.GetReport(li.subItems(3))
           li.subItems(1) = c.positives
           li.subItems(2) = c.first_seen
           Set li.Tag = c
           li.EnsureVisible
           pb.Value = pb.Value + 1
        Next
        
     End If
     
     lv.ListItems(1).EnsureVisible
     pb.Value = 0
     lblStatus = "Complete " & lv.ListItems.count & " records loaded"
     
      
End Sub

Function doDownloads(txtDir As String, Optional doAll As Boolean = False)
    
    Dim li As ListItem
    Dim dl As Boolean
    
    If lv.ListItems.count = 0 Then Exit Function
    
    If Not fso.FolderExists(txtDir) Then
        MsgBox "Set download folder", vbInformation
        Exit Function
    End If
    
    pb.Value = 0
    pb.Max = lv.ListItems.count
    vt.abort = False
    
    For Each li In lv.ListItems
        dl = doAll
        If vt.abort Then Exit For
        If li.Selected = True Then dl = True
        If doAll Or li.Selected Then
            x = Trim(li.subItems(3)) 'md5
            If Len(x) > 0 Then
                If fso.FileExists(txtDir & "\" & x) Then
                    li.Text = "Exists"
                Else
                    li.Text = vt.DownloadFile(CStr(x), txtDir)
                End If
                li.EnsureVisible
            End If
        End If
        DoEvents
        Me.Refresh
        pb.Value = pb.Value + 1
    Next
    
    pb.Value = 0
               
End Function

Function setDLDir(Optional t As String) As Boolean
    If Len(t) = 0 Then t = dlg.FolderDialog2()
    If Len(t) = 0 Then Exit Function
    txtDir = t
    SaveSetting "vt", "settings", "dl_dir", txtDir
    Me.Caption = "Virustotal Search UI - Download Dir: " & txtDir
    setDLDir = True
End Function

Private Sub Form_Load()
    
    Dim hcDat As String
    
    vt.KeepLog = True
    Set vt.debugLog = List1
    Set vt.winInet = Inet1
    Set vt.Timer1 = Timer1
    
    'mnuSetPrivateKey.Checked = vt.usingPrivateKey
    
    txtCacheDir = App.path & "\..\VT_Search_Cache"
    If Not fso.FolderExists(txtCacheDir) Then
        txtCacheDir = App.path & "\..\VT_Search_Cache"
    End If
    
    If Not fso.FolderExists(txtCacheDir) Then fso.CreateFolder txtCacheDir
    If fso.FolderExists(txtCacheDir) Then vt.report_cache_dir = txtCacheDir
    List1.AddItem "Report cache dir: " & vt.report_cache_dir
    
    txtDir = GetSetting("vt", "settings", "dl_dir")
    If fso.FolderExists(txtDir) Then
        List1.AddItem "Download Directory: " & txtDir
    Else
        txtDir = Empty
    End If
       
    hcDat = App.path & "\..\hc.dat"
    If fso.FileExists(hcDat) Then
        hc.LoadHistory hcDat
    Else
        hcDat = App.path & "\hc.dat"
        hc.LoadHistory hcDat
    End If
    
    Me.Caption = Me.Caption & " DlDir: " & txtDir & " Cache: " & txtCacheDir
    lv.SetColumnHeaders "DL,Detections,FirstSeen,Hash*", "500,1200,1200"
    mnuPopup.Visible = False
    txtReport.Font = "Courier"

    If IsIde() Then cboLimit.AddItem "5"
    cboLimit.AddItem "25"
    cboLimit.AddItem "50"
    cboLimit.AddItem "100"
    cboLimit.AddItem "300"
    cboLimit.AddItem "600"
    cboLimit.AddItem "900"
    cboLimit.AddItem "1200"
    
    cboLimit.ListIndex = 0
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    txtReport.Width = Me.Width - txtReport.Left - 200
    txtReport.Height = Me.Height - txtReport.Top - 800
    List1.Height = Me.Height - List1.Top - 800
End Sub

Private Sub Form_Unload(Cancel As Integer)
    hc.SaveHistory
End Sub

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim c As CScan
    On Error Resume Next
    Set c = Item.Tag
    txtReport = c.GetReport
    Set selItem = Item
End Sub

Private Sub lv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    mnuDownloadSelected.Enabled = (lv.selCount > 0)
    If Button = 2 Then PopupMenu mnuPopup
End Sub

Private Sub mnuClearHistory_Click()
    hc.ClearHistory
End Sub

Private Sub mnuDownloadAll_Click()
    If Len(txtDir) = 0 Then setDLDir
    If Len(txtDir) > 0 Then doDownloads txtDir, True
End Sub

Private Sub mnuDownloadSelected_Click()
    If Len(txtDir) = 0 Then setDLDir
    If Len(txtDir) > 0 Then doDownloads txtDir
End Sub

Private Sub mnuGetReport_Click()
    On Error Resume Next
    Dim c As CScan
    If selItem Is Nothing Then Exit Sub
    Set c = selItem.Tag
    If Err.Number <> 0 Then
        Set c = vt.GetReport(selItem.subItems(3))
        selItem.subItems(1) = c.positives
        selItem.subItems(2) = c.first_seen
        Set selItem.Tag = c
    End If
End Sub

Private Sub mnuHelp_Click()
    On Error Resume Next
    Shell "cmd.exe /c start https://virustotal.com/intelligence/help/", vbHide
End Sub

Private Sub mnuRawJSON_Click()
    On Error Resume Next
    Dim c As CScan
    If selItem Is Nothing Then Exit Sub
    Set c = selItem.Tag
    txtReport = c.RawJson
End Sub

Private Sub mnuVisitPage_Click()
    On Error Resume Next
    Dim c As CScan
    If selItem Is Nothing Then Exit Sub
    Set c = selItem.Tag
    c.VisitPage
End Sub

Function IsIde() As Boolean
' Brad Martinez  http://www.mvps.org/ccrp
    On Error GoTo out
    Debug.Print 1 / 0
out: IsIde = Err
End Function
