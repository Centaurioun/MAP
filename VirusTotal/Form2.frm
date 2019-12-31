VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form2 
   Caption         =   "Virus Total Sample Lookup"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9210
   LinkTopic       =   "Form2"
   ScaleHeight     =   6345
   ScaleWidth      =   9210
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   8520
      Top             =   870
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox txtFilter 
      Height          =   330
      Left            =   1395
      TabIndex        =   10
      Top             =   2115
      Width           =   7755
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit File"
      Height          =   255
      Left            =   7920
      TabIndex        =   8
      Top             =   540
      Width           =   1215
   End
   Begin VB.CommandButton cmdSaveReport 
      Caption         =   "Save Report"
      Height          =   315
      Left            =   7920
      TabIndex        =   7
      Top             =   60
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Left            =   7350
      Top             =   420
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3705
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   2550
      Width           =   9165
   End
   Begin VB.TextBox txtFile 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   840
      TabIndex        =   4
      Top             =   60
      Width           =   6945
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
      Left            =   0
      TabIndex        =   2
      Top             =   810
      Width           =   9165
   End
   Begin VB.TextBox txtHash 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   480
      Width           =   4635
   End
   Begin VB.Label Label3 
      Caption         =   "CSV Line Filters:"
      Height          =   240
      Left            =   45
      TabIndex        =   9
      Top             =   2160
      Width           =   1635
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000B&
      Caption         =   "http://virustotal.com"
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
      Height          =   195
      Left            =   5640
      MousePointer    =   1  'Arrow
      TabIndex        =   6
      Top             =   540
      Width           =   1515
   End
   Begin VB.Label Label2 
      Caption         =   "File: "
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "MD5"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   0
      Top             =   510
      Width           =   615
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Begin VB.Menu mnuCopyLine 
         Caption         =   "Copy Line"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCopyTable 
         Caption         =   "Copy Table"
      End
      Begin VB.Menu mnuViewRawJson 
         Caption         =   "View Raw Json"
      End
      Begin VB.Menu mnuGotoPage 
         Caption         =   "Goto Page"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuBulkScan 
         Caption         =   "Bulk Scan"
      End
      Begin VB.Menu mnuBulkDl 
         Caption         =   "Bulk Download"
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "Search"
      End
      Begin VB.Menu mnuSpacer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUsePrivKey 
         Caption         =   "Use Private Api Key"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim md5 As New MD5Hash
Dim vt As New CVirusTotal
Dim scan As CScan

Public Function StartFromFile(fpath As String)

    On Error Resume Next
        
    txtFile = fpath
    
    If Not FileExists(fpath) Then
        List1.AddItem "File not found"
        End
        Exit Function
    End If
    
    Me.Show
    txtHash = md5.HashFile(fpath)
    Set scan = vt.GetReport(txtHash)
    Text1 = scan.GetReport()

End Function



Public Function StartFromHash(hash As String)
    
    On Error Resume Next
    
    txtFile.Enabled = False
    txtFile.BackColor = &H8000000F
    cmdSubmit.Enabled = False
    
    If Len(hash) = 0 Then
        MsgBox "Error starting up from hash, no value specified?", vbInformation
        End
        Exit Function
    End If
    
    
    Me.Show
    txtHash = hash
    Set scan = vt.GetReport(hash)
    Text1 = scan.GetReport()
    
End Function

Private Function FileExists(p) As Boolean
    If Len(p) = 0 Then Exit Function
    If Dir(p, vbNormal Or vbHidden Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
End Function


Private Sub cmdSaveReport_Click()
    
    On Error Resume Next
    Dim pf As String
    Dim path As String
    Dim dlg As New CCmnDlg
    Dim fso As New CFileSystem2
    Dim bn As String
    
    bn = fso.GetBaseName(txtFile)
    pf = fso.GetParentFolder(txtFile)
    If Len(bn) = 0 Then bn = Mid(txtHash, 1, 5)
    bn = "VT_" & bn & ".txt"
    
    path = dlg.SaveDialog(bn, pf, "Save As")
    If Len(path) = 0 Then Exit Sub
    fso.writeFile path, Text1
    
End Sub

Private Sub cmdSubmit_Click()
   
    On Error Resume Next
    If Not FileExists(txtFile) Then
        MsgBox "File not found?"
        Exit Sub
    End If
    
    Set scan = vt.SubmitFile(txtFile)
    scan.response_code = 2 'manually overridden for getreport() display purposes..
    Text1 = scan.GetReport()

 End Sub

Private Sub Form_Load()
    Me.Show
    mnuPopup.Visible = False
    Set vt.Timer1 = Timer1
    Set vt.winInet = Inet1
    Set vt.debugLog = List1
    mnuUsePrivKey.Checked = vt.usingPrivateKey
    mnuBulkDl.Enabled = vt.usingPrivateKey
    mnuSearch.Enabled = vt.usingPrivateKey
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    List1.Width = Me.Width - List1.Left - 200
    Text1.Width = Me.Width - Text1.Left - 200
    Text1.Height = Me.Height - Text1.Top - 400
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    abort = True
    DoEvents
    Timer1.Enabled = False
    End
End Sub

Private Sub Label4_Click()
    On Error Resume Next
    Shell "cmd /c start http://virustotal.com", vbHide
    'If Err.Number <> 0 Then MsgBox Err.Description
End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuPopup
End Sub

Private Sub mnuBulkDl_Click()
    frmBulkDownload.Show
End Sub

Private Sub mnuBulkScan_Click()
    On Error Resume Next
    Form1.Show
    Unload Me
End Sub

Private Sub mnuCopyTable_Click()
    On Error Resume Next
    Dim r
    For i = 0 To List1.ListCount
        r = r & List1.List(i) & vbCrLf
    Next
    Clipboard.Clear
    Clipboard.SetText r
End Sub

Private Sub mnuGotoPage_Click()
    On Error Resume Next
    If Len(scan.permalink) > 0 Then
        Shell "cmd /c start " & scan.permalink
        'MsgBox scan.permalink
    End If
End Sub

Private Sub mnuSearch_Click()
    frmSearch.Show
End Sub

Private Sub mnuUsePrivKey_Click()
    Dim x As String
    
    x = InputBox("By default we use a rate limited public API key. If you have access to a private api key, you may enter it here to avoid delays. " & _
                 "Enter an empty string or hit cancel to clear the private key." & vbCrLf & vbCrLf & "Your key will be stored in the registry.", _
                 "Enter private api key", _
                 vt.ReadPrivateApiKey _
        )
                 
    vt.SetPrivateApiKey x
    mnuUsePrivKey.Checked = vt.usingPrivateKey
    
    If vt.usingPrivateKey Then
        MsgBox "Private key successfull set", vbInformation
        mnuBulkDl.Enabled = True
        mnuSearch.Enabled = True
    Else
        MsgBox "You are now using the default public key which is rate limited and free for non-commercial use. " & vbCrLf & vbCrLf & "Please see the VirusTotal terms of service.", vbInformation
        mnuBulkDl.Enabled = False
        mnuSearch.Enabled = False
    End If
    
End Sub

Private Sub mnuViewRawJson_Click()
    On Error Resume Next
    Text1 = scan.RawJson
End Sub

Private Sub txtFilter_Change()
    On Error Resume Next
    
    If Len(txtFilter.Text) = 0 Then
        Text1 = scan.GetReport()
        Exit Sub
    End If
    
    Dim ret(), tmp() As String, x
    Dim matches() As String, m
    
    matches = Split(txtFilter, ",")
    tmp = Split(scan.GetReport(), vbCrLf)
    
    For i = 0 To 4
        push ret, tmp(i)
    Next
    
    For Each x In tmp
        For Each m In matches
            If Len(m) > 0 Then
                If InStr(1, x, m, vbTextCompare) > 0 Then
                   push ret, x
                   Exit For
                End If
            End If
        Next
    Next
    
    If AryIsEmpty(ret) Then
        Text1 = "No results for " & txtFilter
    Else
        Text1 = Join(ret, vbCrLf)
    End If
    
    
End Sub



Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub



Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function
