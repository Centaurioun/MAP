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
   Begin VB.CommandButton cmdUndo 
      Caption         =   "Undo"
      Height          =   315
      Left            =   7800
      TabIndex        =   12
      Top             =   60
      Width           =   615
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "Replace"
      Height          =   315
      Left            =   6960
      TabIndex        =   11
      Top             =   60
      Width           =   795
   End
   Begin VB.TextBox txtReplace 
      Height          =   285
      Left            =   5880
      TabIndex        =   10
      Top             =   120
      Width           =   795
   End
   Begin VB.TextBox txtFind 
      Height          =   285
      Left            =   4500
      TabIndex        =   9
      Top             =   120
      Width           =   675
   End
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
   Begin VB.Label Label3 
      Caption         =   "Find                 Replace"
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
      Height          =   315
      Left            =   4140
      TabIndex        =   8
      Top             =   120
      Width           =   2895
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
Dim undo As String

Private Sub cmdComplete_Click()
    Me.Visible = False
End Sub

Function ImportHashs(Optional X = "") As String

    If Len(X) = 0 Then
        X = standarize(Clipboard.GetText())
    Else
        X = standarize(X)
    End If
        
    txtIn.Text = X
    txtIn.SelStart = 0
    txtIn.SelLength = Len(txtIn)
    
    Me.Show 1
    ImportHashs = Trim(txtOut)
    Unload Me
    
End Function

Private Sub cmdReplace_Click()
    Dim r As String, f As String
    
    f = Replace(txtFind, "/tab", vbTab)
    f = Replace(f, "/crlf", vbCrLf)
    f = Replace(f, "/cr", vbCr)
    f = Replace(f, "/lf", vbLf)
    
    r = Replace(txtReplace, "/tab", vbTab)
    r = Replace(r, "/crlf", vbCrLf)
    r = Replace(r, "/cr", vbCr)
    r = Replace(r, "/lf", vbLf)
     
    undo = txtIn
    txtIn = Replace(txtIn, f, r)
    
End Sub

Private Sub cmdUndo_Click()
    If Len(undo) > 0 Then
        txtIn = undo
        undo = Empty
    Else
        MsgBox "No undo action available did you do a replace?", vbInformation
    End If
End Sub

Private Sub Command1_Click()
    Dim X, y() As String
    Dim tmp() As String
    Dim index As Long
    
    On Error Resume Next
    index = Abs(CLng(txtElemIndex))
    txtElemIndex = index
    
    tmp = Split(txtIn, vbCrLf)
    For Each X In tmp
        If InStr(X, txtDivider) > 0 Then
             X = Split(X, txtDivider)(index)
        End If
        push y(), X
    Next
    
    txtOut = Join(y, vbCrLf)
        
End Sub

Sub push(ary, Value) 'this modifies parent ary object
    Dim X
    On Error GoTo Init
    X = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = Value
    Exit Sub
Init:     ReDim ary(0): ary(0) = Value
End Sub

Function standarize(X)
    X = Replace(X, vbTab, " ")
    'If InStr(Command, "/bulk") < 1 Then x = Replace(x, ",", Empty) '/bulk command line from shellext uses hash,path\r\n format
    X = Replace(X, "'", Empty)
    X = Replace(X, """", Empty)
    X = Replace(X, ";", Empty)
    X = Replace(X, "}", Empty)
    X = Replace(X, ")", Empty)
    X = Replace(X, vbCr, Chr(5))
    X = Replace(X, vbLf, Empty)
    X = Replace(X, Chr(5), vbCrLf)
    standarize = X
End Function
 
Private Sub Label3_Click()
    MsgBox "Supports /tab /crlf /cr /lf", vbInformation
End Sub
