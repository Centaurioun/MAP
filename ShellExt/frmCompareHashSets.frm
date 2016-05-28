VERSION 5.00
Begin VB.Form frmCompareHashSets 
   Caption         =   "Compare Hash Sets"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   17490
   LinkTopic       =   "Form1"
   ScaleHeight     =   7650
   ScaleWidth      =   17490
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6465
      Left            =   11310
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   450
      Width           =   5985
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Analyze"
      Height          =   525
      Left            =   9870
      TabIndex        =   2
      Top             =   7020
      Width           =   1245
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
      Height          =   6465
      Left            =   5640
      MultiLine       =   -1  'True
      OLEDropMode     =   1  'Manual
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   450
      Width           =   5445
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
      Height          =   6465
      Left            =   90
      MultiLine       =   -1  'True
      OLEDropMode     =   1  'Manual
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   420
      Width           =   5445
   End
   Begin VB.Label Label3 
      Caption         =   "Comparison Report"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11310
      TabIndex        =   6
      Top             =   150
      Width           =   5115
   End
   Begin VB.Label Label2 
      Caption         =   "Compare hash set ( one entry per line)"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5670
      TabIndex        =   5
      Top             =   120
      Width           =   5115
   End
   Begin VB.Label Label1 
      Caption         =   "Base Hash Set  ( one entry per line)"
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
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4785
   End
End
Attribute VB_Name = "frmCompareHashSets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function KeyExistsInCollection(c As Collection, val As String) As Boolean
    On Error GoTo nope
    Dim t
    t = c(val)
    KeyExistsInCollection = True
 Exit Function
nope: KeyExistsInCollection = False
End Function


Private Sub Command1_Click()

    Dim hashs As New Collection
    Dim hashs2 As New Collection
    Dim h
    
    Dim h1() As String
    Dim h2() As String
    
    h1 = Split(Text1, vbCrLf)
    h2 = Split(Text2, vbCrLf)
    
    'build a unique list of hashs in base directory set..
    For Each h In h1
        h = Trim(h)
        If Len(h) > 0 And Not KeyExistsInCollection(hashs, CStr(h)) Then
            hashs.Add h, CStr(h)
        End If
     Next
     
     'build a unique list of hashs in compare directory set..
     For Each h In h2
        h = Trim(h)
        If Len(h) > 0 And Not KeyExistsInCollection(hashs2, CStr(h)) Then
            hashs2.Add h, CStr(h)
        End If
     Next
     
    For Each h In hashs
        results = InStr(1, Text2, h)
        If results > 0 Then
            r = r & h & vbCrLf
        Else
            unique1 = unique1 & h & vbCrLf
            unique1_hits = unique1_hits + 1
        End If
    Next
    
    'now we find the files from the second directory not found in main compare dir..
    For Each h In hashs2
        results = InStr(1, Text1, h)
        If results < 1 Then
            unique2 = unique2 & h & vbCrLf
            unique2_hits = unique2_hits + 1
        End If
    Next
    
    If Len(r) > 0 Then
        
        report = "Base set:    " & UBound(h1) & " hashs / " & hashs.Count & " unique" & vbCrLf & _
                 "Compare set: " & UBound(h2) & " hashs / " & hashs2.Count & " unique" & vbCrLf & vbCrLf & _
                 "Hashs found in both sets:  " & totalHits & vbCrLf & vbCrLf & _
                 r & vbCrLf & vbCrLf & _
                 "Hashs unique to base dir: " & unique1_hits & " files" & vbCrLf & vbCrLf & _
                 unique1 & vbCrLf & vbCrLf & _
                 "Hashs unique to compare dir: " & unique2_hits & " files" & vbCrLf & vbCrLf & _
                 unique2
                
        Text3 = report
    Else
        MsgBox "There were no hash matches in these two sample sets.", vbInformation
    End If
    
     
End Sub

Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim X As Long
    X = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub


Private Sub Form_Load()
    Me.Icon = myIcon
End Sub

Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If fso.FileExists(Data.Files(1)) Then
        Text1 = fso.ReadFile(Data.Files(1))
    End If
End Sub

Private Sub Text2_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If fso.FileExists(Data.Files(1)) Then
        Text2 = fso.ReadFile(Data.Files(1))
    End If
End Sub

