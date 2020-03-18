VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
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
   Begin MSComctlLib.ProgressBar pb 
      Height          =   375
      Left            =   60
      TabIndex        =   10
      Top             =   240
      Visible         =   0   'False
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton mnuSelectHashs 
      Caption         =   "Select In Main UI"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   7080
      Width           =   1755
   End
   Begin VB.CommandButton cmdUnique 
      Caption         =   "Unique Only"
      Height          =   435
      Left            =   2940
      TabIndex        =   8
      Top             =   7020
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear All"
      Height          =   495
      Left            =   4440
      TabIndex        =   7
      Top             =   7020
      Width           =   1155
   End
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
      Height          =   6240
      Left            =   11310
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   675
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
      Height          =   6240
      Left            =   5640
      MultiLine       =   -1  'True
      OLEDropMode     =   1  'Manual
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   675
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
      Height          =   6195
      Left            =   90
      MultiLine       =   -1  'True
      OLEDropMode     =   1  'Manual
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   690
      Width           =   5445
   End
   Begin VB.Label lblSelHelp 
      Caption         =   "?"
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
      Left            =   2040
      TabIndex        =   11
      Top             =   7140
      Width           =   135
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
      Caption         =   "Compare hash set ( one entry per line)  can drop files here"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5670
      TabIndex        =   5
      Top             =   120
      Width           =   5115
   End
   Begin VB.Label Label1 
      Caption         =   "Base Hash Set  ( one entry per line)   can drop files here"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
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


Private Sub cmdClear_Click()
    Text1 = Empty
    Text2 = Empty
    Text3 = Empty
End Sub

Private Sub cmdUnique_Click()
    Dim hashs As New Collection
    Dim h
    Dim h1() As String
    Dim tmp() As String
    
    On Error Resume Next
    h1 = Split(Text1, vbCrLf)
    
    'build a unique list of hashs in base directory set..
    For Each h In h1
        h = trim(h)
        If Len(h) > 0 And Not KeyExistsInCollection(hashs, CStr(h)) Then
            hashs.Add h, CStr(h)
            push tmp, h
        End If
    Next
     
    Text1 = Join(tmp, vbCrLf)
    Me.Caption = UBound(tmp) & " unique"
    
End Sub

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
        h = trim(h)
        If Len(h) > 0 And Not KeyExistsInCollection(hashs, CStr(h)) Then
            hashs.Add h, CStr(h)
        End If
     Next
     
     'build a unique list of hashs in compare directory set..
     For Each h In h2
        h = trim(h)
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
    Dim x As Long
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub


Private Sub Form_Load()
    Me.Icon = myIcon
End Sub

Private Sub lblSelHelp_Click()
    MsgBox "Enter a list of MD5s to try to find and select in dir hash UI. Then you can Move Selected or take other action", vbInformation
End Sub

Private Sub mnuSelectHashs_Click()
    On Error Resume Next
    Dim tmp() As String, x, li As ListItem, found As Boolean, miss() As String, hits As Long, t() As String
    
    If Len(Text1.text) = 0 Then
        MsgBox "Paste the md5s you want to select above", vbInformation
        Exit Sub
    End If
        
    tmp = Split(Replace(Text1.text, vbTab, Empty), vbCrLf)
    
    For Each li In frmHash.lv.ListItems
        li.selected = False
    Next
    
    pb.max = UBound(tmp)
    pb.Visible = True
    pb.value = 0
    
    For Each li In frmHash.lv.ListItems 'probably less hashs than listitems..
        found = False
        pb.value = pb.value + 1
        For Each x In tmp
            If LCase(li.SubItems(2)) = LCase(trim(x)) Then
                li.selected = True
                found = True
                hits = hits + 1
                Exit For
            End If
        Next
        If Not found Then push miss, x
        DoEvents
    Next
    
    push t, hits & "/" & UBound(tmp) - 1 & " selected "
    
    If Not AryIsEmpty(miss) Then
        push t, "Misses: "
        push t, Join(miss, vbCrLf)
    End If
    
    pb.value = 0
    pb.Visible = False
    Text2 = Join(t, vbCrLf)
    
End Sub

Private Sub Text1_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error Resume Next
    If fso.FileExists(data.Files(1)) Then
        Text1 = fso.ReadFile(data.Files(1))
    End If
End Sub

Private Sub Text2_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error Resume Next
    If fso.FileExists(data.Files(1)) Then
        Text2 = fso.ReadFile(data.Files(1))
    End If
End Sub

