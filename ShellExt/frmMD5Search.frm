VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMD5FileSearch 
   Caption         =   "File Hash Search"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   11685
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optHash 
      Caption         =   "File Name"
      Height          =   375
      Index           =   2
      Left            =   7800
      TabIndex        =   14
      Top             =   90
      Width           =   1275
   End
   Begin VB.OptionButton optHash 
      Caption         =   "SHA1"
      Height          =   315
      Index           =   1
      Left            =   6810
      TabIndex        =   13
      Top             =   120
      Width           =   855
   End
   Begin VB.OptionButton optHash 
      Caption         =   "MD5"
      Height          =   285
      Index           =   0
      Left            =   6000
      TabIndex        =   12
      Top             =   120
      Value           =   -1  'True
      Width           =   705
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "default"
      Height          =   315
      Left            =   10800
      TabIndex        =   11
      Top             =   540
      Width           =   765
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   285
      Left            =   150
      TabIndex        =   10
      Top             =   1320
      Width           =   11415
      _ExtentX        =   20135
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
      Width           =   11445
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "Scan"
      Height          =   375
      Left            =   10530
      TabIndex        =   8
      Top             =   60
      Width           =   1125
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   285
      Left            =   10800
      TabIndex        =   7
      Top             =   990
      Width           =   765
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
      Width           =   9105
   End
   Begin VB.CheckBox chkRecursive 
      Caption         =   "Recursive"
      Height          =   255
      Left            =   9330
      TabIndex        =   4
      Top             =   150
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
      Width           =   9135
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
      TabIndex        =   1
      Top             =   120
      Width           =   4395
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
   Begin VB.Label Label1 
      Caption         =   "Hash (full or partial)"
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

Sub Launch(Optional baseDir As String)
    txtBaseDir = baseDir
    Me.Show
End Sub

Private Sub cmdBrowse_Click()
    txtBaseDir = dlg.FolderDialog(, Me.hwnd)
End Sub

Private Sub cmdDefault_Click()
    txtExt.Text = "txt doc docx pdf zip rar 7z rtf idb pcap png jpg"
End Sub

Private Sub cmdScan_Click()
    
    Dim hashs() As String
    Dim ignore As Boolean
    
    txtHash = Trim(txtHash)
    txtResults = Trim(txtResults)
    
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
    tmp() = fso.GetFolderFiles(txtBaseDir, , , CBool(chkRecursive.Value))
    
    Dim method As HashTypes 'also supports md2/md4 but dont see any reason to complicate ui..
    If optHash(0).Value Then method = md5 Else method = SHA
    
    push ret, "Scanning " & UBound(tmp) + 1 & " files in base directory for " & UBound(hashs) + 1 & " hash(s)"
    Me.Caption = ret(0)
    
    If UBound(tmp) > 0 Then pb.max = UBound(tmp) + 1
    
    For Each f In tmp
        pb.Value = pb.Value + 1
        Me.Caption = "Scanning " & f
        ext = fso.GetExtension(f)
        If Len(ext) > 2 Then ext = Mid(ext, 2)   'remove leading dot
        If Len(ext) = 0 Or _
            InStr(1, txtExt, ext, vbTextCompare) < 1 Or _
            optHash(2).Value = True _
        Then 'not found in ignore list or is a file name search..
        
            If optHash(2).Value = True Then 'is file name search
                h = CStr(f)
            Else
                h = hash.HashFile(CStr(f), method)
            End If
            
            For Each hh In hashs                 'cycle through each hash given to match..
                hh = Trim(hh)
                If Len(hh) > 0 And InStr(1, h, hh, vbTextCompare) > 0 Then
                    If optHash(2).Value = True Then
                        push ret, f
                    Else
                        compiled = GetCompileDateOrType(CStr(f))
                        push ret, h & vbCrLf & compiled & vbCrLf & f & vbCrLf
                    End If
                    c = c + 1
                End If
            Next
            
        End If
        pb.Refresh
        Me.Refresh
        DoEvents
        DoEvents
    Next
    
    Me.Caption = c & " hits"
    pb.Value = 0
    ret(0) = ret(0) & " - " & c & " hits" & vbCrLf
    txtResults = Join(ret, vbCrLf)
    
End Sub

Sub push(ary, Value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Long
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = Value
    Exit Sub
init: ReDim ary(0): ary(0) = Value
End Sub

Private Sub Form_Load()
    Me.Icon = myIcon
    cmdDefault_Click
    txtExt = GetSetting("shellext", "settings", "txtExt", txtExt.Text)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    txtResults.Width = Me.Width - 250
    txtResults.Height = Me.Height - txtResults.Top - 400
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "shellext", "settings", "txtExt", txtExt.Text
End Sub

Private Sub optHash_Click(index As Integer)
    txtExt.BackColor = IIf(index = 2, &HC0C0C0, vbWhite) 'ignore list not used for file name search..
End Sub

Private Sub txtBaseDir_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error Resume Next
    
    If fso.FolderExists(Data.Files(1)) Then
        txtBaseDir.Text = Data.Files(1)
    Else
        MsgBox "Only drop folders in here..", vbInformation
    End If
    
End Sub





