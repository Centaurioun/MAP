VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCompareHashSets 
   Caption         =   "Compare Hash Sets"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16725
   LinkTopic       =   "Form1"
   ScaleHeight     =   7650
   ScaleWidth      =   16725
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBaseRight 
      Caption         =   "g"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4380
      TabIndex        =   21
      Top             =   420
      Width           =   555
   End
   Begin VB.CommandButton cmdCmpLeft 
      Caption         =   "f"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10020
      TabIndex        =   20
      Top             =   420
      Width           =   555
   End
   Begin VB.CommandButton cmdCompareOnly 
      Caption         =   "Compare Only"
      Height          =   255
      Left            =   14820
      TabIndex        =   19
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdBaseOnly 
      Caption         =   "Base Only"
      Height          =   255
      Left            =   13560
      TabIndex        =   18
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton cmdShared 
      Caption         =   "Shared"
      Height          =   255
      Left            =   12780
      TabIndex        =   16
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10620
      TabIndex        =   13
      Top             =   420
      Width           =   555
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4980
      TabIndex        =   12
      Top             =   420
      Width           =   555
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   6195
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   10927
      _Version        =   393217
      ScrollBars      =   3
      OLEDropMode     =   1
      TextRTF         =   $"frmCompareHashSets.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   375
      Left            =   60
      TabIndex        =   7
      Top             =   60
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
      TabIndex        =   6
      Top             =   7080
      Width           =   1755
   End
   Begin VB.CommandButton cmdUnique 
      Caption         =   "Unique Only"
      Height          =   435
      Left            =   2940
      TabIndex        =   5
      Top             =   7020
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear All"
      Height          =   495
      Left            =   4440
      TabIndex        =   4
      Top             =   7020
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Analyze"
      Height          =   525
      Left            =   9870
      TabIndex        =   0
      Top             =   7020
      Width           =   1245
   End
   Begin RichTextLib.RichTextBox Text2 
      Height          =   6195
      Left            =   5700
      TabIndex        =   10
      Top             =   720
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   10927
      _Version        =   393217
      ScrollBars      =   3
      OLEDropMode     =   1
      TextRTF         =   $"frmCompareHashSets.frx":007C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox Text3 
      Height          =   6195
      Left            =   11220
      TabIndex        =   11
      Top             =   720
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   10927
      _Version        =   393217
      ScrollBars      =   3
      OLEDropMode     =   1
      TextRTF         =   $"frmCompareHashSets.frx":00F8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label4 
      Caption         =   "Copy:"
      Height          =   255
      Left            =   12240
      TabIndex        =   17
      Top             =   480
      Width           =   435
   End
   Begin VB.Label lblFile2 
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   5820
      TabIndex        =   15
      Top             =   480
      Width           =   4575
   End
   Begin VB.Label lblFile1 
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   240
      TabIndex        =   14
      Top             =   480
      Width           =   4455
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
      TabIndex        =   8
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
      TabIndex        =   3
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
      TabIndex        =   2
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
      TabIndex        =   1
      Top             =   120
      Width           =   4785
   End
End
Attribute VB_Name = "frmCompareHashSets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private both As String
Private baseOnly As String
Private compareOnly As String

Function KeyExistsInCollection(c As Collection, val As String) As Boolean
    On Error GoTo nope
    Dim t
    t = c(val)
    KeyExistsInCollection = True
 Exit Function
nope: KeyExistsInCollection = False
End Function


Private Sub cmdBaseOnly_Click()
    Clipboard.Clear
    Clipboard.SetText baseOnly
End Sub

Private Sub cmdClear_Click()
    Text1 = Empty
    Text2 = Empty
    Text3 = Empty
End Sub

Private Sub cmdCmpLeft_Click()
    Text1.text = Text2.text
    lblFile1.Caption = lblFile2.Caption
    Text2.text = Empty
    lblFile2.Caption = Empty
End Sub

Private Sub cmdCompareOnly_Click()
    Clipboard.Clear
    Clipboard.SetText compareOnly
End Sub

Private Sub cmdShared_Click()
    Clipboard.Clear
    Clipboard.SetText both
End Sub

Private Sub cmdUnique_Click()
    Dim hashs As New Collection
    Dim h
    Dim h1() As String
    Dim tmp() As String
    
    On Error Resume Next
    h1 = Split(Text1.text, vbCrLf)
    
    'build a unique list of hashs in base directory set..
    For Each h In h1
        h = trim(h)
        If Len(h) > 0 And Not KeyExistsInCollection(hashs, CStr(h)) Then
            hashs.Add h, CStr(h)
            push tmp, h
        End If
    Next
     
    Text1.text = Join(tmp, vbCrLf)
    Me.Caption = UBound(tmp) & " unique"
    
End Sub

Private Sub Command1_Click()

    Dim hashs As New CollectionEx
    Dim hashs2 As New CollectionEx
    Dim h
    Dim unique1() As String, unique2() As String
    
    Dim h1() As String
    Dim h2() As String
    
    both = Empty
    baseOnly = Empty
    compareOnly = Empty
    
    Text3.text = Empty
    h1 = Split(Text1.text, vbCrLf)
    h2 = Split(Text2.text, vbCrLf)
    
    'build a unique list of hashs in base directory set..
    For Each h In h1
        h = trim(h)
        'If Len(h) > 0 And Not  KeyExistsInCollection(hashs, CStr(h)) Then
        If Len(h) > 0 And Not hashs.keyExists(CStr(h)) Then
            hashs.Add h, CStr(h)
        End If
     Next
     
     'build a unique list of hashs in compare directory set..
     For Each h In h2
        h = trim(h)
        'If Len(h) > 0 And Not KeyExistsInCollection(hashs2, CStr(h)) Then
        If Len(h) > 0 And Not hashs2.keyExists(CStr(h)) Then
            hashs2.Add h, CStr(h)
        End If
     Next
     
    For Each h In hashs
        'results = InStr(1, Text2.text, h)
        If hashs2.keyExists(h) Then
        'If results > 0 Then
            r = r & h & vbCrLf
            totalHits = totalHits + 1
        Else
            'unique1 = unique1 & h & vbCrLf
            push unique1, h
            unique1_hits = unique1_hits + 1
        End If
    Next
    
    'now we find the files from the second directory not found in main compare dir..
    For Each h In hashs2
        'results = InStr(1, Text1.text, h)
        If Not hashs.keyExists(h) Then
        'If results < 1 Then
            'unique2 = unique2 & h & vbCrLf
            push unique2, h
            unique2_hits = unique2_hits + 1
        End If
    Next
    
    If Len(r) > 0 Then
        
        both = r
        baseOnly = Join(unique1, vbCrLf)
        compareOnly = Join(unique2, vbCrLf)
        
        report = "Base set:    " & (UBound(h1) + 1) & " hashs / " & hashs.Count & " unique" & vbCrLf & _
                 "Compare set: " & (UBound(h2) + 1) & " hashs / " & hashs2.Count & " unique" & vbCrLf & vbCrLf & _
                 "Hashs found in both sets:  " & totalHits & vbCrLf & vbCrLf & _
                 r & vbCrLf & vbCrLf & _
                 "Hashs unique to base dir: " & unique1_hits & " files" & vbCrLf & vbCrLf & _
                 Join(unique1, vbCrLf) & vbCrLf & vbCrLf & _
                 "Hashs unique to compare dir: " & unique2_hits & " files" & vbCrLf & vbCrLf & _
                 Join(unique2, vbCrLf)
                
        Text3.text = report

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

Public Sub preload(file1 As String, Optional file2 As String)
    
    lblFile1.Caption = Empty
    If Len(file1) > 0 Then
        If fso.FileExists(file1) Then
            Text1.text = fso.ReadFile(file1)
            lblFile1.Caption = fso.FileNameFromPath(file1)
        End If
    End If
    
    lblFile2.Caption = Empty
    If Len(file2) > 0 Then
        If fso.FileExists(file2) Then
            Text2.text = fso.ReadFile(file2)
            lblFile2.Caption = fso.FileNameFromPath(file2)
        End If
    End If
    
End Sub

Private Sub Command2_Click()
    Dim f As String
    On Error Resume Next
    f = dlg.OpenDialog()
    lblFile1.Caption = Empty
    If Len(f) > 0 Then
        If fso.FileExists(f) Then
            Text1.text = fso.ReadFile(f)
            lblFile1.Caption = fso.FileNameFromPath(f)
        End If
    End If
End Sub

Private Sub Command3_Click()
    Dim f As String
    On Error Resume Next
    f = dlg.OpenDialog()
    lblFile2.Caption = Empty
    If Len(f) > 0 Then
        If fso.FileExists(f) Then
            Text2.text = fso.ReadFile(f)
            lblFile2.Caption = fso.FileNameFromPath(f)
        End If
    End If
End Sub

Private Sub cmdBaseRight_Click()
    Text2.text = Text1.text
    lblFile2.Caption = lblFile1.Caption
    Text1.text = Empty
    lblFile1.Caption = Empty
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim f As Form
    
    Me.Icon = myIcon
    mnuSelectHashs.Visible = False
    
    For Each f In Forms
        If f.Name = "frmHash" Then
             mnuSelectHashs.Visible = True
             Exit For
        End If
    Next
    
    
End Sub

Private Sub lblSelHelp_Click()
    MsgBox "Enter a list of MD5s to try to find and select in dir hash UI. Then you can Move Selected or take other action", vbInformation
End Sub

Private Sub mnuSelectHashs_Click()
    On Error Resume Next
    Dim tmp() As String, X, li As ListItem, found As Boolean, miss() As String, hits As Long, t() As String
    
    
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
        For Each X In tmp
            If LCase(li.subItems(2)) = LCase(trim(X)) Then
                li.selected = True
                found = True
                hits = hits + 1
                Exit For
            End If
        Next
        If Not found Then push miss, X
        DoEvents
    Next
    
    push t, hits & "/" & UBound(tmp) - 1 & " selected "
    
    If Not AryIsEmpty(miss) Then
        push t, "Misses: "
        push t, Join(miss, vbCrLf)
    End If
    
    pb.value = 0
    pb.Visible = False
    Text2.text = Join(t, vbCrLf)
    
End Sub

Private Sub Text1_Change()
    lblFile1.Caption = Empty
End Sub

Private Sub Text2_Change()
    lblFile2.Caption = Empty
End Sub

Private Sub Text1_OLEDragDrop(data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
    On Error Resume Next
    Dim f As String
    f = data.Files(1)
    lblFile1.Caption = Empty
    If fso.FileExists(f) Then
        Text1.text = fso.ReadFile(f)
        lblFile1.Caption = fso.FileNameFromPath(f)
    End If
End Sub

Private Sub Text2_OLEDragDrop(data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
    On Error Resume Next
    Dim f As String
    f = data.Files(1)
    lblFile2.Caption = Empty
    If fso.FileExists(f) Then
        Text2.text = fso.ReadFile(f)
        lblFile2.Caption = fso.FileNameFromPath(f)
    End If
End Sub

