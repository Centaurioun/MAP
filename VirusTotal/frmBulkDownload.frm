VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmBulkDownload 
   Caption         =   "Bulk Download"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10410
   LinkTopic       =   "Form3"
   ScaleHeight     =   6000
   ScaleWidth      =   10410
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   4560
      Top             =   2760
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   9720
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   285
      Left            =   4005
      TabIndex        =   6
      Top             =   495
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ListView lv 
      Height          =   5100
      Left            =   3960
      TabIndex        =   5
      Top             =   810
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
      Left            =   9270
      TabIndex        =   4
      Top             =   135
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
      Left            =   3915
      OLEDropMode     =   1  'Manual
      TabIndex        =   2
      Text            =   "Supports drag and drop"
      Top             =   90
      Width           =   4695
   End
   Begin VB.TextBox txtHash 
      Height          =   5415
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   450
      Width           =   3750
   End
   Begin VB.Label Label1 
      Caption         =   "Hashs One per line or CSV                  save to  folder:"
      Height          =   240
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   3750
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

'todo: async download with progress bar instead of blocking

Private Sub cmdBrowse_Click()
    Dim x As String
    x = dlg.FolderDialog2()
    If Len(x) = 0 Then Exit Sub
    txtDir = x
End Sub

Private Sub cmdDownload_Click()
    
    Dim li As ListItem
    
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
    pb.value = 0
    
    tmp = Split(txtHash, vbCrLf)
    pb.Max = UBound(tmp) + 1
    
    For Each x In tmp
        x = Trim(x)
        If Len(x) > 0 Then
            Set li = lv.ListItems.Add(, , x)
            lv.Refresh
            If fso.FileExists(txtDir & "\" & x) Then
                li.subItems(1) = "Exists"
            Else
                li.subItems(1) = vt.DownloadFile(CStr(x), txtDir)
            End If
            lv.Refresh
        End If
        DoEvents
        Me.Refresh
        pb.value = pb.value + 1
    Next
    
    pb.value = 0
        
End Sub

Private Sub Form_Load()
    Set vt.Timer1 = Timer1
    Set vt.winInet = Inet1
End Sub

Private Sub txtDir_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, xx As Single, y As Single)
    On Error Resume Next
    Dim x As String
    x = Data.files(1)
    If fso.FolderExists(x) Then
        txtDir = x
    End If
End Sub

