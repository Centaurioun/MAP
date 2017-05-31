VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmSubmit 
   Caption         =   "VT Submit"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9825
   LinkTopic       =   "Form3"
   ScaleHeight     =   3225
   ScaleWidth      =   9825
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   7290
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   375
      Left            =   60
      TabIndex        =   2
      Top             =   2700
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer tmrDelay 
      Enabled         =   0   'False
      Interval        =   4500
      Left            =   8040
      Top             =   2400
   End
   Begin VB.CommandButton cmdAbort 
      Caption         =   "Abort"
      Height          =   435
      Left            =   8520
      TabIndex        =   1
      Top             =   2700
      Width           =   1215
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2490
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   9675
   End
End
Attribute VB_Name = "frmSubmit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vt As New CVirusTotal

Public Sub SubmitFile(fpath As String)
    Dim s As CScan
    Me.Show
    Set s = vt.SubmitFile(fpath)
    List1.AddItem s.verbose_msg
    List1.ListIndex = List1.ListCount - 1
    Me.Refresh
    DoEvents
End Sub

Public Sub SubmitBulk()
    
    On Error Resume Next
    Dim files As New Collection
    Dim s As CScan
    
    x = Clipboard.GetText
    tmp = Split(x, vbCrLf)
    For Each x In tmp
        x = Trim(x)
        If Len(x) > 0 Then
            If fso.FileExists(CStr(x)) Then files.Add x
        End If
    Next
    
    Me.Show
    List1.AddItem "Loaded " & files.count & " file paths from clipboard.."
    pb.Value = 0
    pb.Max = files.count
    
    For Each x In files
        If vt.abort Then Exit For
        Set s = vt.SubmitFile(CStr(x))
        List1.AddItem s.verbose_msg
        List1.ListIndex = List1.ListCount - 1
        pb.Value = pb.Value + 1
        Me.Refresh
        DoEvents
    Next
    
    pb.Value = 0
    List1.AddItem "Complete"
    List1.ListIndex = List1.ListCount - 1
    Me.Refresh
    DoEvents
    
End Sub

Private Sub cmdAbort_Click()
    vt.abort = True
End Sub

Private Sub Form_Load()
   Set vt.Timer1 = tmrDelay
End Sub
