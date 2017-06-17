VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMD5FileSearch 
   Caption         =   "File Hash Search"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13095
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   13095
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSelectFile 
      Caption         =   "..."
      Height          =   315
      Left            =   7680
      TabIndex        =   16
      Top             =   120
      Width           =   495
   End
   Begin VB.ComboBox cboHash 
      Height          =   315
      Left            =   8220
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   120
      Width           =   1275
   End
   Begin VB.CheckBox chkStopFirst 
      Caption         =   "Stop on first"
      Height          =   255
      Left            =   10800
      TabIndex        =   14
      Top             =   120
      Width           =   1155
   End
   Begin VB.TextBox txtSize 
      Height          =   315
      Left            =   6720
      TabIndex        =   13
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "default"
      Height          =   315
      Left            =   12000
      TabIndex        =   11
      Top             =   540
      Width           =   1005
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   285
      Left            =   150
      TabIndex        =   10
      Top             =   1320
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox txtResults 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3165
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   9
      Top             =   1680
      Width           =   12885
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "Scan"
      Height          =   375
      Left            =   11970
      TabIndex        =   8
      Top             =   60
      Width           =   1125
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   285
      Left            =   12000
      TabIndex        =   7
      Top             =   960
      Width           =   1005
   End
   Begin VB.TextBox txtBaseDir 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1500
      OLEDropMode     =   1  'Manual
      TabIndex        =   6
      ToolTipText     =   "supports drag and drop"
      Top             =   900
      Width           =   10305
   End
   Begin VB.CheckBox chkRecursive 
      Caption         =   "Recursive"
      Height          =   255
      Left            =   9600
      TabIndex        =   4
      Top             =   120
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.TextBox txtExt 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1470
      TabIndex        =   3
      Top             =   510
      Width           =   10335
   End
   Begin VB.TextBox txtHash 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1470
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      ToolTipText     =   "supports drag drop"
      Top             =   120
      Width           =   4395
   End
   Begin VB.Label Label4 
      Caption         =   "Size (opt)"
      Height          =   375
      Left            =   5940
      TabIndex        =   12
      Top             =   180
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Base Directory:"
      Height          =   255
      Left            =   90
      TabIndex        =   5
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Ignore Extensions"
      Height          =   285
      Left            =   60
      TabIndex        =   2
      Top             =   540
      Width           =   1395
   End
   Begin VB.Label lblHash 
      Caption         =   "Hash (full or partial)"
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
      Left            =   90
      TabIndex        =   0
      Top             =   150
      Width           =   1395
   End
End
Attribute VB_Name = "frmMD5FileSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fso As New CFileSystem2
Dim hash As New CWinHash
Dim dlg As New clsCmnDlg2
Dim abort As Boolean
Dim LoadedFile As String 'only set if user browses to set hash/size

Sub Launch(Optional baseDir As String)
    txtBaseDir = baseDir
    Me.Show
End Sub

Private Sub cboHash_Click()
    If Len(LoadedFile) > 0 Then loadHashSize
End Sub

Private Sub cmdBrowse_Click()
    txtBaseDir = dlg.FolderDialog(, Me.hwnd)
End Sub

Private Sub cmdDefault_Click()
    txtExt.text = "txt doc docx pdf zip rar 7z rtf idb pcap png jpg"
End Sub

Function SelectedHash() As HashTypes
    Dim method As HashTypes 'also supports md2/md4 but dont see any reason to complicate ui..
    If cboHash.ListIndex = 0 Then method = md5
    If cboHash.ListIndex = 1 Then method = SHA
    If cboHash.ListIndex = 2 Then method = 256
    If cboHash.ListIndex = 3 Then method = 512
    SelectedHash = method
End Function

Private Sub cmdScan_Click()
    
    Dim hashs() As String
    Dim ignore As Boolean
    Dim size As Long
    Dim i As Long
    Dim isFileNameSearch As Boolean
    
    On Error Resume Next
    
    If cmdScan.Caption = "Abort" Then
        abort = True
        Exit Sub
    End If
    
    abort = False
    size = CLng(txtSize)
    txtHash = trim(txtHash)
    txtResults = trim(txtResults)
    isFileNameSearch = (InStr(lblHash, "Hash") < 1)
    
    If Len(txtHash) = 0 Then
        If Len(txtResults) = 0 Then
            MsgBox "You must either Specify a single hash in top textbox or a list or hashs one per line in lower textbox", vbInformation
            Exit Sub
        Else
            hashs() = Split(txtResults, vbCrLf)
        End If
    Else
        push hashs(), txtHash
    End If
    
    If Not fso.FolderExists(txtBaseDir) Then
        MsgBox "Folder not found", vbInformation
        Exit Sub
    End If
    
    Dim tmp() As String, ret() As String
    Dim f, h As String, c As Long, ext As String, hh
    Dim compiled As String
    
    Me.Caption = "Enumerating files..."
    tmp() = fso.GetFolderFiles(txtBaseDir, , , CBool(chkRecursive.value))
    
    Dim method As HashTypes 'also supports md2/md4 but dont see any reason to complicate ui..
    method = SelectedHash()
    
    push ret, "Scanning " & UBound(tmp) + 1 & " files in base directory for " & UBound(hashs) + 1 & " hash(s)"
    Me.Caption = ret(0)
    
    If UBound(tmp) > 0 Then pb.max = UBound(tmp) + 1
    
    cmdScan.Caption = "Abort"
    
    For Each f In tmp
        i = i + 1
        If abort Then Exit For
        pb.value = pb.value + 1
        Me.Caption = "Scanning " & f
        ext = fso.GetExtension(f)
        If Len(ext) > 2 Then ext = Mid(ext, 2)   'remove leading dot
        If Len(ext) = 0 Or _
            InStr(1, txtExt, ext, vbTextCompare) < 1 Or _
            isFileNameSearch _
        Then 'not found in ignore list or is a file name search..
        
            If size <> 0 Then
                If FileLen(CStr(f)) <> size Then GoTo nextone
            End If
            
            If isFileNameSearch Then 'is file name search
                h = CStr(f)
            Else
                h = hash.HashFile(CStr(f), method)
            End If
            
            For Each hh In hashs                 'cycle through each hash given to match..
                hh = trim(hh)
                If Len(hh) > 0 And InStr(1, h, hh, vbTextCompare) > 0 Then
                    If isFileNameSearch Then
                        push ret, f
                        If chkStopFirst.value Then Exit For
                    Else
                        compiled = GetCompileDateOrType(CStr(f))
                        push ret, h & vbCrLf & compiled & vbCrLf & f & vbCrLf
                        If chkStopFirst.value Then Exit For
                    End If
                    c = c + 1
                End If
            Next
            
        End If
        
nextone:
        If i Mod 5 Then
            pb.Refresh
            Me.Refresh
            DoEvents
            DoEvents
        End If
        
    Next
    
    cmdScan.Caption = "Scan"
    Me.Caption = c & " hits"
    pb.value = 0
    ret(0) = ret(0) & " - " & c & " hits" & vbCrLf
    txtResults = Join(ret, vbCrLf)
    
End Sub

Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Long
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init: ReDim ary(0): ary(0) = value
End Sub

Private Sub cmdSelectFile_Click()
    LoadedFile = dlg.OpenDialog(AllFiles)
    loadHashSize
End Sub

Function loadHashSize()
    If Len(LoadedFile) > 0 Then
        Dim method As HashTypes
        method = SelectedHash()
        txtHash = hash.HashFile(LoadedFile, method)
        txtSize = FileLen(LoadedFile)
    Else
        txtHash = Empty
        txtSize = Empty
    End If
End Function

Private Sub Form_Load()
    Me.Icon = myIcon
    cmdDefault_Click
    txtExt = GetSetting("shellext", "settings", "txtExt", txtExt.text)
    cboHash.AddItem "MD5"
    cboHash.AddItem "SHA1"
    cboHash.AddItem "SHA256"
    cboHash.AddItem "SHA512"
    cboHash.ListIndex = 0
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    txtResults.Width = Me.Width - 250
    txtResults.Height = Me.Height - txtResults.top - 400
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "shellext", "settings", "txtExt", txtExt.text
End Sub

Private Sub optHash_Click(index As Integer)
    txtExt.BackColor = IIf(index = 2, &HC0C0C0, vbWhite) 'ignore list not used for file name search..
End Sub

Private Sub lblHash_Click()
    If InStr(lblHash, "Hash") > 0 Then
        lblHash = "File Name"
    Else
        lblHash = "Hash (full or partial)"
    End If
End Sub

Private Sub txtBaseDir_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error Resume Next
    
    If fso.FolderExists(Data.Files(1)) Then
        txtBaseDir.text = Data.Files(1)
    Else
        MsgBox "Only drop folders in here..", vbInformation
    End If
    
End Sub





Private Sub txtHash_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error Resume Next
    Dim f As String
    f = Data.Files(1)
    If fso.FileExists(f) Then
        LoadedFile = f
        loadHashSize
    End If
End Sub

