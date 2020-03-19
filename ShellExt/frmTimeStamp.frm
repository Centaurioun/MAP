VERSION 5.00
Begin VB.Form frmTimeStamp 
   Caption         =   "Time Stamp Calculator"
   ClientHeight    =   1605
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   ScaleHeight     =   1605
   ScaleWidth      =   6135
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1125
      TabIndex        =   7
      Top             =   945
      Width           =   3030
   End
   Begin VB.CommandButton cmdFromDate 
      Caption         =   "From Date"
      Height          =   330
      Left            =   4365
      TabIndex        =   5
      Top             =   585
      Width           =   1545
   End
   Begin VB.CommandButton cmdFromStamp 
      Caption         =   "From Timestamp"
      Height          =   375
      Left            =   4365
      TabIndex        =   4
      Top             =   135
      Width           =   1545
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1125
      TabIndex        =   3
      Top             =   540
      Width           =   3030
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1125
      TabIndex        =   1
      Top             =   135
      Width           =   3030
   End
   Begin VB.Label Label3 
      Caption         =   "Format"
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
      Height          =   285
      Left            =   585
      TabIndex        =   6
      Top             =   990
      Width           =   555
   End
   Begin VB.Label Label2 
      Caption         =   "Date"
      Height          =   285
      Left            =   675
      TabIndex        =   2
      Top             =   585
      Width           =   420
   End
   Begin VB.Label Label1 
      Caption         =   "TimeStamp"
      Height          =   285
      Left            =   270
      TabIndex        =   0
      Top             =   180
      Width           =   870
   End
End
Attribute VB_Name = "frmTimeStamp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFromDate_Click()
    
    On Error Resume Next
    Dim d As Date, s As String, e As String, i As Long
    
    d = CDate(Text2)
    
    If Err.Number > 0 Then
        MsgBox Err.Description
        Exit Sub
    End If
    
    i = DatetoTimeStamp(d)
    Text1 = "0x" & Hex(i)
    
End Sub

Private Sub cmdFromStamp_Click()
    
    On Error Resume Next
    
    Dim i As Long
    i = CLng(trim(Replace(Text1, "0x", "&h")))
    
    If Err.Number <> 0 Then
        MsgBox Err.Description
        Exit Sub
    End If
    
    Text2 = timeStampToDate(i, Text3)
    
End Sub

Private Sub Form_Load()
    Text3 = GetSetting("map", "shellext", "dateFormat", "mmm d yyyy h:nn:ss")
    Text1 = GetSetting("map", "shellext", "t1", 1584574200)
    Text2 = GetSetting("map", "shellext", "t2", "Mar 18 2020 23:30:00")
End Sub

Private Sub Form_Unload(Cancel As Integer)
     SaveSetting "map", "shellext", "dateFormat", Text3
     SaveSetting "map", "shellext", "t1", Text1
     SaveSetting "map", "shellext", "t2", Text2
End Sub

Private Sub Label3_Click()
    Text3 = "mmm d yyyy h:nn:ss"
End Sub

