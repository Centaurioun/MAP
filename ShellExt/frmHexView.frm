VERSION 5.00
Object = "{9A143468-B450-48DD-930D-925078198E4D}#1.1#0"; "hexed.ocx"
Begin VB.Form frmHexView 
   Caption         =   "Form1"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   11175
   StartUpPosition =   2  'CenterScreen
   Begin rhexed.HexEd HexEd1 
      Height          =   5535
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   9763
   End
End
Attribute VB_Name = "frmHexView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub HexView(fPath As String)
    Dim fs As Long
    Dim a As Long
    Dim base
    On Error Resume Next
    
    a = InStr(1, fPath, "/base=", vbTextCompare)
    
    If a > 0 Then
        base = Trim(Mid(fPath, a + Len("/base=")))
        fPath = Trim(Mid(fPath, 1, a - 1))
        If Len(base) <= 8 Then HexEd1.AdjustBaseOffset = CLng("&h" & base)
    End If
        
    If Not fso.FileExists(fPath) Then
        MsgBox "File not found: " & fPath, vbExclamation
        GoTo done
    End If
    
    HexEd1.LoadFile fPath
    Me.Caption = "HexViewer -  Base:" & base & " Path: " & fPath
    
done:
    Me.Visible = True
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    HexEd1.Height = Me.Height - 600
    HexEd1.Width = Me.Width - 300
End Sub
