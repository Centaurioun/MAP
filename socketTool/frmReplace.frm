VERSION 5.00
Begin VB.Form frmReplace 
   Caption         =   "Find/Replace"
   ClientHeight    =   1350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5340
   LinkTopic       =   "Form3"
   ScaleHeight     =   1350
   ScaleWidth      =   5340
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Replace"
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   945
      Width           =   1335
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
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   480
      Width           =   4335
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
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   0
      Width           =   4335
   End
   Begin VB.Label Label2 
      Caption         =   "Replace"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Find"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuCopyAll 
         Caption         =   "Copy All"
      End
   End
End
Attribute VB_Name = "frmReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public active_object As TextBox

Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_SHOWWINDOW = &H40


Public Sub LaunchReplaceForm(txtObj As TextBox)
    
    Set active_object = txtObj
    Me.Show
    
End Sub

Private Sub Command1_Click()
    active_object = Replace(active_object.Text, Text1, Text2, , , vbTextCompare)
End Sub

Private Sub Form_Load()
    Me.Icon = Form1.Icon
    SetWindowPos Me.hwnd, HWND_TOPMOST, Me.Left / 15, Me.Top / 15, Me.Width / 15, Me.Height / 15, SWP_SHOWWINDOW
End Sub


 

 
