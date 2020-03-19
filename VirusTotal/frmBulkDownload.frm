VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmBulkDownload 
   Caption         =   "Bulk Download"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   11415
   LinkTopic       =   "Form3"
   ScaleHeight     =   6030
   ScaleWidth      =   11415
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAbort 
      Caption         =   "Abort"
      Height          =   315
      Left            =   9210
      TabIndex        =   9
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtLimit 
      Height          =   285
      Left            =   3300
      TabIndex        =   8
      Text            =   "-1"
      Top             =   120
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Left            =   11040
      Top             =   1740
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   10860
      Top             =   1020
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   285
      Left            =   4920
      TabIndex        =   6
      Top             =   480
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ListView lv 
      Height          =   5100
      Left            =   4920
      TabIndex        =   5
      Top             =   780
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   8996
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "hash"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "status"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdDownload 
      Caption         =   "Download"
      Height          =   330
      Left            =   10140
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   285
      Left            =   8685
      TabIndex        =   3
      Top             =   135
      Width           =   465
   End
   Begin VB.TextBox txtDir 
      Height          =   285
      Left            =   5175
      OLEDropMode     =   1  'Manual
      TabIndex        =   2
      Text            =   "Supports drag and drop"
      Top             =   90
      Width           =   3435
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
      Height          =   5415
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   450
      Width           =   4710
   End
   Begin VB.Label Label2 
      Caption         =   "import"
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
      Left            =   2100
      TabIndex        =   7
      Top             =   180
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Hashs One per line or CSV                     limit                   save to  folder:"
      Height          =   240
      Left            =   60
      TabIndex        =   0
      Top             =   180
      Width           =   5010
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
      End
   End
End
Attribute VB_Name = "frmBulkDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim fso As New CFileSystem2
Dim vt As New CVirusTotal
Dim dlg As New CCmnDlg
Dim abort As Boolean

Private Sub cmdAbort_Click()
    abort = True
End Sub

'todo: async download with progress bar instead of blocking

Private Sub cmdBrowse_Click()
    Dim X As String
    X = dlg.FolderDialog2(txtDir)
    If Len(X) = 0 Then Exit Sub
    txtDir = X
End Sub

Function getLimit() As Long
    On Error Resume Next
    Dim limit As Long
    limit = CLng(txtLimit)
    If limit < 1 Then limit = 0
    getLimit = limit
End Function

Private Sub cmdDownload_Click()
    
    Dim lI As ListItem
    Dim limit As Long
    
    abort = False
    limit = getLimit()
    
    If Not fso.FolderExists(txtDir) Then
        MsgBox "Set download folder", vbInformation
        Exit Sub
    End If
    
    If InStr(txtHash, ",") > 0 Then
        txtHash = Replace(txtHash, ",", vbCrLf)
    End If
    
    txtHash = Replace(txtHash, vbTab, Empty)
    txtHash = Trim(Replace(txtHash, vbCrLf & vbCrLf, vbCrLf))
    
    lv.ListItems.Clear
    pb.Value = 0
    
    tmp = Split(txtHash, vbCrLf)
    pb.Max = UBound(tmp) + 1
    
    For Each X In tmp
        If abort Then Exit For
        If limit > 0 Then
            If pb.Value > limit Then Exit For
        End If
        X = Trim(X)
        If Len(X) > 0 Then
            Set lI = lv.ListItems.Add(, , X)
            lv.Refresh
            If fso.FileExists(txtDir & "\" & X) Then
                lI.subItems(1) = "Exists"
            Else
                lI.subItems(1) = vt.DownloadFile(CStr(X), txtDir)
            End If
            lv.Refresh
        End If
        DoEvents
        Me.Refresh
        pb.Value = pb.Value + 1
        Me.Caption = pb.Value & "/" & pb.Max
    Next
    
    pb.Value = 0
        
End Sub

Private Sub Form_Load()
    Set vt.Timer1 = Timer1
    Set vt.winInet = Inet1
    mnuPopup.Visible = False
End Sub

Private Sub Label2_Click()
    Dim tmp As String
    tmp = frmImport.ImportHashs()
    If Len(tmp) > 0 Then
        txtHash = tmp
    End If
End Sub

Private Sub lv_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvColumnSort lv, ColumnHeader
End Sub

Private Sub lv_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = 2 Then PopupMenu mnuPopup
End Sub

Private Sub mnuCopy_Click()
    Clipboard.Clear
    Clipboard.SetText lvGetAllElements(lv)
End Sub

Private Sub txtDir_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, xx As Single, y As Single)
    On Error Resume Next
    Dim X As String
    X = Data.files(1)
    If fso.FolderExists(X) Then
        txtDir = X
    End If
End Sub

