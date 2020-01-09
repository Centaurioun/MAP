VERSION 5.00
Begin VB.Form frmImport 
   Caption         =   "Hash Import"
   ClientHeight    =   7125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14130
   Icon            =   "frmImport.frx":0000
   ScaleHeight     =   7125
   ScaleWidth      =   14130
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtElemIndex 
      Height          =   315
      Left            =   3000
      TabIndex        =   7
      Text            =   "0"
      Top             =   60
      Width           =   795
   End
   Begin VB.TextBox txtDivider 
      Height          =   315
      Left            =   780
      TabIndex        =   5
      Text            =   " "
      Top             =   60
      Width           =   675
   End
   Begin VB.CommandButton cmdComplete 
      Caption         =   "Done"
      Height          =   315
      Left            =   10260
      TabIndex        =   3
      Top             =   60
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Import"
      Height          =   315
      Left            =   8760
      TabIndex        =   2
      Top             =   60
      Width           =   1335
   End
   Begin VB.TextBox txtOut 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   3780
      Width           =   11235
   End
   Begin VB.TextBox txtIn 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   420
      Width           =   11235
   End
   Begin VB.Label Label2 
      Caption         =   "Elem Index"
      Height          =   255
      Left            =   1860
      TabIndex        =   6
      Top             =   120
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "Divider"
      Height          =   315
      Left            =   180
      TabIndex        =   4
      Top             =   120
      Width           =   675
   End
End
Attribute VB_Name = "frmImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdComplete_Click()
    Me.Visible = False
End Sub

Function ImportHashs(Optional x = "") As String

    If Len(x) = 0 Then
        x = standarize(Clipboard.GetText())
    Else
        x = standarize(x)
    End If
        
    txtIn.Text = x
    txtIn.SelStart = 0
    txtIn.SelLength = Len(txtIn)
    
    Me.Show 1
    ImportHashs = Trim(txtOut)
    Unload Me
    
End Function

Private Sub Command1_Click()
    Dim x, y() As String
    Dim tmp() As String
    Dim index As Long
    
    On Error Resume Next
    index = Abs(CLng(txtElemIndex))
    txtElemIndex = index
    
    tmp = Split(txtIn, vbCrLf)
    For Each x In tmp
        If InStr(x, txtDivider) > 0 Then
             x = Split(x, txtDivider)(index)
        End If
        push y(), x
    Next
    
    txtOut = Join(y, vbCrLf)
        
End Sub

Sub push(ary, value) 'this modifies parent ary object
    Dim x
    On Error GoTo init
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub

Function standarize(x)
    x = Replace(x, vbTab, " ")
    'If InStr(Command, "/bulk") < 1 Then x = Replace(x, ",", Empty) '/bulk command line from shellext uses hash,path\r\n format
    x = Replace(x, "'", Empty)
    x = Replace(x, """", Empty)
    x = Replace(x, ";", Empty)
    x = Replace(x, "}", Empty)
    x = Replace(x, ")", Empty)
    x = Replace(x, vbCr, Chr(5))
    x = Replace(x, vbLf, Empty)
    x = Replace(x, Chr(5), vbCrLf)
    standarize = x
End Function
 
