VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmRecursiveHashFiles 
   Caption         =   "Hash all files below "
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   14460
   LinkTopic       =   "frmRecursiveHashFiles"
   ScaleHeight     =   5895
   ScaleWidth      =   14460
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Hash Files Below"
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
      Left            =   11925
      TabIndex        =   2
      Top             =   0
      Width           =   2490
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
      Left            =   900
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Text            =   "drop here"
      Top             =   45
      Width           =   10995
   End
   Begin MSComctlLib.ListView lv 
      Height          =   5280
      Left            =   45
      TabIndex        =   0
      Top             =   450
      Width           =   14280
      _ExtentX        =   25188
      _ExtentY        =   9313
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Hash"
         Object.Width           =   8467
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "File"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Folder"
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
      Left            =   0
      TabIndex        =   3
      Top             =   45
      Width           =   1365
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Begin VB.Menu mnuCopyHashs 
         Caption         =   "Copy Hashs"
      End
      Begin VB.Menu mnuCopyCSV 
         Caption         =   "Copy CSV Results"
      End
      Begin VB.Menu mnuCopyReport 
         Caption         =   "Copy for Report"
      End
   End
End
Attribute VB_Name = "frmRecursiveHashFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fso As New CFileSystem2
Dim hash As New CWinHash

Sub RecursiveHashDir(folder As String)
    Text1 = folder
    Me.Visible = True
    Command1_Click
End Sub

Private Sub Command1_Click()
    
    On Error Resume Next
    
    If Not fso.FolderExists(Text1) Then
        MsgBox "No folder found"
        Exit Sub
    End If
    
    Dim ff() As String, f, li As ListItem
    ff() = fso.GetFolderFiles(Text1, , , True)
    
    lv.ListItems.Clear
    
    For Each f In ff
        Set li = lv.ListItems.Add()
        li.text = hash.HashFile(CStr(f))
        li.SubItems(1) = pad(Hex(FileLen(CStr(f))))
        li.SubItems(2) = f
    Next
    
End Sub

 

Private Sub Form_Resize()
    On Error Resume Next
    lv.Width = Me.Width - lv.Left - 200
    lv.Height = Me.Height - lv.top - 400
    sizeLvCol lv
End Sub

Private Sub lv_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuPopup
End Sub

Private Sub mnuCopyHashs_Click()
    Dim li As ListItem
    Dim x()
    
    push x, "hashs for: " & Text1 & vbCrLf
    
    For Each li In lv.ListItems
        push x, li.text
    Next
    
    Clipboard.Clear
    Clipboard.SetText Join(x, vbCrLf)
    
End Sub

Private Sub mnuCopyCSV_Click()
    Dim li As ListItem
    Dim x()
    
    push x, "hash,hexSize,path"
    
    For Each li In lv.ListItems
        push x, li.text & "," & li.SubItems(1) & "," & li.SubItems(2)
    Next
    
    Clipboard.Clear
    Clipboard.SetText Join(x, vbCrLf)
End Sub

Private Sub Form_Load()
    mnuPopup.Visible = False
    sizeLvCol lv
End Sub

Private Sub mnuCopyReport_Click()
    Dim li As ListItem
    Dim x()
    
    For Each li In lv.ListItems
        push x, rpad("File: ") & li.SubItems(2)
        push x, rpad("Size: ") & li.SubItems(1)
        push x, rpad("MD5: ") & li.text & vbCrLf
    Next
    
    Clipboard.Clear
    Clipboard.SetText Join(x, vbCrLf)
End Sub

Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error Resume Next
    If Not fso.FolderExists(Data.Files(1)) Then Exit Sub
    If Err.Number <> 0 Then Exit Sub
    Text1 = Data.Files(1)
End Sub





Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub
