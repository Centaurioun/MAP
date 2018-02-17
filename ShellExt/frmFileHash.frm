VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmFileHash 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Hash"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5850
      Top             =   1110
   End
   Begin VB.Timer Timer1 
      Left            =   5850
      Top             =   675
   End
   Begin VB.PictureBox pictIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   5580
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Frame fraLower 
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   60
      TabIndex        =   0
      Top             =   1260
      Width           =   5835
      Begin VB.CommandButton cmdCopyHash 
         Caption         =   "Copy Hash"
         Height          =   345
         Left            =   3060
         TabIndex        =   2
         Top             =   0
         Width           =   1125
      End
      Begin VB.CommandButton cmdCopyAll 
         Caption         =   "Copy All"
         Height          =   345
         Left            =   4560
         TabIndex        =   1
         Top             =   0
         Width           =   1215
      End
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   1095
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1931
      _Version        =   393217
      BackColor       =   -2147483633
      BorderStyle     =   0
      Appearance      =   0
      TextRTF         =   $"frmFileHash.frx":0000
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
   Begin VB.Menu mnuPopup 
      Caption         =   "Actions"
      Begin VB.Menu mnuNameMD5 
         Caption         =   "Rename to MD5"
      End
      Begin VB.Menu mnuStrings 
         Caption         =   "Strings"
      End
      Begin VB.Menu mnuFileProps 
         Caption         =   "File Properties"
      End
      Begin VB.Menu mnuOffsetCalc 
         Caption         =   "Offset Calculator (32/64bit)"
      End
      Begin VB.Menu mnuPEVerInfo 
         Caption         =   "PE Version Info"
      End
      Begin VB.Menu mnuSpacer 
         Caption         =   "-"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMoreHashs 
      Caption         =   "InfoLevel"
      Begin VB.Menu mnuCopyHashMore 
         Caption         =   "MD5"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCopyHashMore 
         Caption         =   "SHA1"
         Index           =   1
      End
      Begin VB.Menu mnuCopyHashMore 
         Caption         =   "SHA256"
         Index           =   2
      End
      Begin VB.Menu mnuCopyHashMore 
         Caption         =   "SHA512"
         Index           =   3
      End
      Begin VB.Menu mnuCopyHashMore 
         Caption         =   "File Properties"
         Index           =   4
      End
      Begin VB.Menu mnuCopyHashMore 
         Caption         =   "VirusTotal"
         Index           =   5
      End
      Begin VB.Menu mnuCopyHashMore 
         Caption         =   "PE Version"
         Index           =   6
      End
      Begin VB.Menu mnuCopyHashMore 
         Caption         =   "Entropy"
         Index           =   7
      End
      Begin VB.Menu mnuCopyHashMore 
         Caption         =   "Detect It Easy"
         Index           =   8
      End
      Begin VB.Menu mnuCopyHashMore 
         Caption         =   "imphash"
         Index           =   9
      End
   End
   Begin VB.Menu mnuVTParent 
      Caption         =   "VirusTotal"
      Begin VB.Menu mnuVT 
         Caption         =   "View Results"
      End
      Begin VB.Menu mnuGotoScan 
         Caption         =   "Goto Scan Page"
      End
      Begin VB.Menu mnuSubmitToVT 
         Caption         =   "Submit To VirusTotal"
      End
      Begin VB.Menu mnuSpacer2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVTCache 
         Caption         =   "Cache Results"
      End
   End
   Begin VB.Menu mnuExternal 
      Caption         =   "External"
      Begin VB.Menu mnuKryptoAnalyzer 
         Caption         =   "Krypto Analyzer"
      End
      Begin VB.Menu mnuCorFlags 
         Caption         =   ".NET Force 32Bit"
      End
      Begin VB.Menu mnuDllChar 
         Caption         =   "Dll Characteristics"
      End
      Begin VB.Menu mnuSearchFileName 
         Caption         =   "Google File Name"
      End
      Begin VB.Menu mnuSearchHash 
         Caption         =   "Google File MD5"
      End
      Begin VB.Menu mnuExt 
         Caption         =   "Edit Cfg"
         Index           =   0
      End
      Begin VB.Menu mnuExt 
         Caption         =   "-"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmFileHash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'updated to sppe3 should now all be x64 safe
'note: diescan wont work on xpsp0 because of msvcr100.dll dependancies, known ok on xpsp3

Dim myMd5 As String
Dim LoadedFile As String
Dim isPE As Boolean
Dim scan As CScan
Dim vt As New CVirusTotal
Dim hashs() 'checked menu names (infolevel)
Dim vt_cache As String
Dim pe As New CPEEditor

'Dim WithEvents subclass As CSubclass2
'Dim kanal As CWindow

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hINst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Const WM_COMMAND = &H111

Dim lastX As Long
Dim lastY As Long
Dim curWord As String
Dim startPos As Long
Dim links()

Enum opts
    oMD5 = 0
    oSHA1
    oSHA256
    oSHA512
    oFileProps
    oVT
    oPEVer
    oEntropy
    oDiE
    oImpHash
End Enum


Private Sub mnuDllChar_Click()
    frmDllCharacteristics.LoadFile LoadedFile
End Sub

Private Sub Text1_Click()
    
    Dim marker As String
    Dim i As Long
    Dim resHackPath As String
    
    marker = vbCrLf & "   "
    
    If Screen.MousePointer = vbArrow Then
         
         If curWord = "Exports:" Then
            If pe.isLoaded Then
                'output is ready for use with lib.exe /def:in.def /out:out.lib
                frmPEVersion.ShowReport LoadedFile, "LIBRARY """ & fso.GetBaseName(LoadedFile) & """" & vbCrLf & "EXPORTS " & marker & c2s(pe.Exports.dumpNames, marker)
            End If
         End If
         
         If curWord = "Resources:" Then
            'if they have a copy of reshacker in map dir then it overrides our internal one..
            resHackPath = App.path & IIf(IsIde, "\..", Empty) & "\ResHacker.exe"
            If pe.isLoaded And Not fso.FileExists(resHackPath) Then
                frmResViewer.ShowResources pe
            Else
                launchResHack LoadedFile
            End If
         End If
         
         If curWord = "Scan_Date:" Then mnuGotoScan_Click
         If curWord = "Detections:" Then mnuVt_Click
         
         Screen.MousePointer = vbNormal
    End If
    
End Sub

Function launchResHack(target As String)
    
    On Error Resume Next
    
    Dim exe As String, i As Long, resHackPath
    
    If Not fso.FileExists(target) Then Exit Function
    
    exe = App.path & IIf(IsIde, "\..", Empty) & "\ResHacker.exe"
    
    If fso.FileExists(exe) Then
        Shell exe & """" & target & """", vbNormalFocus
    Else
    
        'find it from the user defined externals
        For i = 0 To mnuExt.Count
            If InStr(1, mnuExt(i).Caption, "reshack", vbTextCompare) > 0 Then
                resHackPath = mnuExt(i).Tag
                Exit For
            End If
        Next
        
        If InStr(1, resHackPath, "%app_path%", vbTextCompare) > 0 Then
            resHackPath = Replace(resHackPath, "%app_path%", App.path)
        End If
        
        If Len(resHackPath) = 0 Then
            MsgBox "If you add reshack to your external apps entry this will start it for you on click"
        ElseIf fso.FileExists(trim(Replace(resHackPath, "%1", Empty))) Then
            resHackPath = Replace(resHackPath, "%1", """" & LoadedFile & """")
            Shell resHackPath, vbNormalFocus
        Else
            MsgBox "ResHack external app path is not a valid file path on this system" & vbCrLf & _
                    "choose external -> edit cfg to edit it"
        End If
        
    End If
            
    
End Function

Function hilite_keywords()
    On Error Resume Next
    Dim z, a
    For Each z In links
        a = InStr(1, Text1.text, z)
        If a > 0 Then
            Text1.SelStart = a - 1
            Text1.SelLength = Len(z) - 1
            'Text1.SelColor = &HC0C0C0C0
            Text1.SelUnderline = True
        End If
    Next
End Function

Private Function isOpt(o As opts) As Boolean
    isOpt = mnuCopyHashMore(o).Checked
End Function

Function ShowIcon(ByVal fileName As String, ByVal hDC As Long, Optional ByVal iconIndex As Long = 0, Optional ByVal x As Long = 0, Optional ByVal Y As Long = 0) As Boolean
    Dim hIcon As Long
    hIcon = ExtractIcon(App.hInstance, fileName, iconIndex)

    If hIcon Then
        DrawIcon hDC, x, Y, hIcon
        ShowIcon = True
    End If
    
End Function

Function hasInterestingResources() As Boolean
    Dim r As CResData, i As Long
    For i = 1 To pe.Resources.Entries.Count
        Set r = pe.Resources.Entries(i)
        If InStr(1, r.path, "icon", vbTextCompare) < 1 And InStr(1, r.path, "version", vbTextCompare) < 1 Then
            hasInterestingResources = True
            Exit Function
        End If
    Next
End Function

Sub ShowFileStats(fPath As String)
    
    On Error Resume Next
    Dim ret() As String
    Dim istype As Boolean
    Dim compiled As String
    Dim fs As Long, sz As Long
    Dim fname As String
    Dim mySHA As String
    Dim Sections As String
    Dim tmp As String
    Dim isX64 As Boolean
    Dim a, b
    
    'cmdExports.Visible = False
    'cmdRes.Visible = False
        
    LoadedFile = fPath
    fs = DisableRedir()
    myMd5 = hash.HashFile(fPath)
    pe.LoadFile fPath
    
    If myMd5 = fso.FileNameFromPath(fPath) Then
        mnuNameMD5.Enabled = False
    End If
    
    sz = FileLen(fPath)
    RevertRedir fs
    
    fname = fso.FileNameFromPath(fPath)
    If Len(fname) > 50 Then
        fname = GetShortName(fPath)
        fname = fso.FileNameFromPath(fname)
    End If
    
    If LCase(fname) <> LCase(myMd5) Then
        push ret(), rpad("File:") & fname
    End If
    
    push ret(), rpad("Size:") & sz
    push ret(), rpad("MD5:") & myMd5
    
    If isOpt(oSHA1) Then
        push ret(), rpad("SHA1:") & hash.HashFile(fPath, SHA, HexFormat)
    End If
    
    If isOpt(oSHA256) Then
        tmp = hash.HashFile(fPath, 256, HexFormat)
        If Len(tmp) = 0 Then tmp = "must update vbdevkit.dll"
        push ret(), rpad("SHA256:") & tmp
    End If
    
    If isOpt(oSHA512) Then
        tmp = hash.HashFile(fPath, 512, HexFormat)
        If Len(tmp) > 0 Then
            a = Mid(tmp, 1, 64)
            b = Mid(tmp, 65)
            push ret(), rpad("SHA512:") & a
            push ret(), rpad(" ") & b
        Else
            push ret, rpad("SHA512:") & "must update vbdevkit.dll"
        End If
    End If
    
    compiled = GetCompileDateOrType(fPath, istype, isPE, isX64)
    push ret(), IIf(istype, rpad("FileType: "), rpad("Compiled:")) & compiled
    
    If isPE Then
        
        If InStr(compiled, ".NET") > 0 Then
            If InStr(compiled, "AnyCPU") > 0 Or InStr(compiled, "64 Bit") > 0 Then
                mnuCorFlags.Enabled = True
            End If
        End If
                        
        'If InStr(LCase(compiled), "dll") > 0 Then  mnuDllChar.Enabled = True
            
        'If pe.LoadFile(fPath, Sections) Then 'little wasteful to load the pe twice (compile date 1st) but managable..
        '    If Len(Sections) > 0 Then push ret(), "Sections: " & Sections
        'End If
        
        If Len(pe.DebugDirectory.pdbPath) > 0 Then push ret(), rpad("PDB: ") & pe.DebugDirectory.pdbPath
            
        Dim fp As CFileProps
        Set fp = FileProps.FileInfo(fPath) 'should we include more here? we need a config pane now :(
        If Len(fp.FileVersion) > 0 Then
            push ret(), rpad("Version:") & fp.FileVersion
        End If
        
    End If
    
    If isOpt(oEntropy) Then push ret, "Entropy:  " & fileEntropy(fPath)
    
    push ret(), Empty
    
    If pe.isLoaded Then
        If pe.Exports.functions.Count > 0 Then push ret(), "Exports:  " & pe.Exports.functions.Count
        If hasInterestingResources Then push ret(), "Resources: " & pe.Resources.Entries.Count & " - " & pe.Resources.size & " bytes"
    End If
    
    If isPE And isOpt(oPEVer) Then push ret(), PEVersionReport(pe, True)
    If pe.isLoaded And isOpt(oImpHash) Then push ret(), "ImpHash:  " & LCase(pe.impHash())
    
    If isOpt(oDiE) Then
        If DiEScan(LoadedFile, tmp) Then push ret(), "DiE:      " & tmp
    End If
    
    Dim v As SigResults
    Dim subject As String, issuer As String
    v = VerifyFileSignature(fPath)
    If isSigned(v) Then
        push ret(), rpad("Signature ") & SigToStr(v)
        If GetSigner(fPath, issuer, subject) Then
            If Len(subject) > 0 Then push ret(), rpad("Subject:") & subject
            If Len(issuer) > 0 Then push ret(), rpad("Issuer:") & issuer
        End If
    End If
    
    If isOpt(oVT) Then
       If scan Is Nothing Then
            Set scan = vt.GetReport(myMd5)
       End If
       mnuGotoScan.Enabled = (Len(scan.permalink) > 0)
       mnuVT.Enabled = mnuGotoScan.Enabled
       push ret(), Replace(scan.BriefReport(), "Scan Date", "Scan_Date")
    End If

    If isOpt(oFileProps) Then
       push ret(), vbCrLf & FileProps.FileInfo(fPath).asStr()
    End If
        
    mnuFileProps.Enabled = isPE
    mnuOffsetCalc.Enabled = isPE
    'If mnuOffsetCalc.Enabled Then
    '    If isX64 Then mnuOffsetCalc.Enabled = False
    'End If
    
    mnuPEVerInfo.Enabled = isPE
     
    Text1 = Join(ret, vbCrLf)
    
    hilite_keywords
    Me.fontname = Text1.Font.Name
    Me.FontSize = Text1.Font.size
    Text1.Height = TextHeight(Text1.text) + 200
    Text1.Width = TextWidth(Text1.text) + 200
    Me.Height = Text1.top + Text1.Height + fraLower.Height + 700
    Me.Width = Text1.Width + Text1.Left + 400
    fraLower.top = Me.Height - fraLower.Height - 750
    
    If ShowIcon(fPath, pictIcon.hDC) Then
        'Me.Width = Me.Width + pictIcon.Width '+ 50
        pictIcon.Left = Me.Width - pictIcon.Width
        pictIcon.Visible = True
    End If
    
    Dim minWidth  As Long
    minWidth = fraLower.Width + fraLower.Left + 300
    If Me.Width < minWidth Then Me.Width = minWidth
    
    Me.Show '1 why was using a modal show? any reason?? made popup menus on subforms not show up..
        
End Sub

Function c2a(c As Collection) As Variant()
    Dim tmp(), x
    For Each x In c
        push tmp, x
    Next
    c2a = tmp
End Function

Function c2s(c As Collection, Optional delimiter = ",") As String
    Dim tmp()
    tmp = c2a(c)
    c2s = Join(tmp, delimiter)
End Function

Private Sub cmdCopyAll_Click()
    Clipboard.Clear
    Clipboard.SetText Text1.text
    Unload Me
    End
End Sub

Private Sub cmdCopyHash_Click()
    Clipboard.Clear
    Clipboard.SetText myMd5
    Unload Me
    End
End Sub

Private Sub cmdVT_Click()
    On Error Resume Next
    Dim vt As String
    vt = App.path & IIf(IsIde(), "\..\", "") & "\virustotal.exe"
    If Not fso.FileExists(vt) Then
        MsgBox "VirusTotal app not found?: " & vt, vbInformation
        Exit Sub
    End If
    'Shell vt & " /hash " & myMd5
    Shell vt & " " & LoadedFile 'so submit button is active..
End Sub

Private Sub Form_Load()
    
    'Set subclass = New CSubclass2
    pictIcon.BackColor = &H8000000F
    
    mnuVTCache.Checked = GetMySetting("mnuVTCache", True)
    vt_cache = Environ("temp") & "\vt_cache"
    If Not fso.FolderExists(vt_cache) Then MkDir vt_cache
    If mnuVTCache.Checked Then vt.report_cache_dir = vt_cache
        
    Me.Icon = myIcon
    vt.TimerObj = Timer1
    mnuCorFlags.Enabled = False
    'mnuDllChar.Enabled = False
    
    Dim ext As String, tmp() As String, x, i
    
    ext = App.path & IIf(IsIde(), "\..\", "") & "\shellext.external.txt"
    If fso.FileExists(ext) Then
        ext = fso.ReadFile(ext)
        tmp = Split(ext, vbCrLf)
        For Each x In tmp
            AddExternal CStr(x)
        Next
    End If
    
    On Error Resume Next
    
    hashs = Array("MD5", "SHA1", "SHA256", "SHA512", "FileProps", "VirusTotal", "PEVersion", "Entropy", "DiE", "imphash")
    links = Array("Exports:", "Resources:", "Scan_Date:", "Detections:")
    
    For i = 0 To mnuCopyHashMore.Count - 1
        mnuCopyHashMore(i).Checked = CBool(GetMySetting(hashs(i), IIf(i = 0, True, False)))
    Next
    
    
End Sub



Private Sub mnuCopyHashMore_Click(index As Integer)
 
    mnuCopyHashMore(index).Checked = Not mnuCopyHashMore(index).Checked
    SaveMySetting hashs(index), mnuCopyHashMore(index).Checked
    Me.ShowFileStats LoadedFile
    
End Sub

Private Sub mnuCorFlags_Click()
    Dim exe As String
    Dim f As String
    
    On Error Resume Next
    
    exe = App.path & "\CorFlags.exe"
    If Not fso.FileExists(exe) Then exe = fso.GetParentFolder(App.path) & "\CorFlags.exe"
    If Not fso.FileExists(exe) Then
        MsgBox "Corflags.exe not found? " & exe
        Exit Sub
    End If
    
    f = fso.GetParentFolder(LoadedFile) & "\" & fso.GetBaseName(LoadedFile) & ".32bit"
    If fso.FileExists(f) Then Kill f
    FileCopy LoadedFile, f
    
    Shell """" & exe & """ """ & f & """ /32Bit+"
    Sleep 500
    ShowFileStats f
    
End Sub

Private Sub mnuExt_Click(index As Integer)
    On Error GoTo hell
    Dim cmd As String
    
    If index = 0 Then
        cmd = App.path & IIf(IsIde(), "\..\", "") & "\shellext.external.txt"
        Shell "notepad.exe " & GetShortName(cmd), vbNormalFocus
    Else
        cmd = mnuExt(index).Tag
        cmd = Replace(cmd, "%1", GetShortName(LoadedFile))
        cmd = Replace(cmd, "%app_path%", App.path & IIf(IsIde(), "\..\", "\"))
        Shell cmd, vbNormalFocus
    End If
    
    Exit Sub
hell:
    MsgBox "Error launching program cmdline: " & vbCrLf & vbCrLf & cmd, vbInformation
        
End Sub

Private Sub mnuFileProps_Click()
    On Error Resume Next
    Dim fs As Long, f As String, tmp As String
    fs = DisableRedir()
    tmp = FileProps.FileInfo(LoadedFile).asStr()
    RevertRedir fs
    If Len(tmp) = 0 Then Exit Sub
    f = fso.GetFreeFileName(Environ("temp"))
    fso.WriteFile f, vbCrLf & vbCrLf & tmp
    Shell "notepad.exe """ & f & """", vbNormalFocus
End Sub

Private Sub mnuGotoScan_Click()
    If scan Is Nothing Then
        Set scan = vt.GetReport(myMd5)
    End If
    If scan.HadError Then
        MsgBox scan.BriefReport
    Else
        scan.VisitPage
    End If
End Sub

Private Sub mnuKryptoAnalyzer_Click()
    'Set kanal = Nothing
    'Timer2.Tag = 0
    'Timer2.enabled = True
    LaunchPeidPlugin "kanal.dll", LoadedFile 'this is a modal dialog so we have to enable timer first...
End Sub

Private Sub mnuNameMD5_Click()
    On Error Resume Next
    Dim fNew As String
    fNew = fso.GetParentFolder(LoadedFile) & "\" & myMd5
    If fso.FileExists(fNew) Then
        MsgBox "A file named the md5 already exists in the target directory", vbExclamation
        Exit Sub
    End If
    Name LoadedFile As fNew
    If Err.Number = 0 Then
        LoadedFile = fNew
        ShowFileStats fNew
    Else
        MsgBox "Error renaming file: " & Err.Description
    End If
End Sub

Private Sub mnuOffsetCalc_Click()
    On Error Resume Next
    Dim pe As New CPEEditor
    If pe.LoadFile(LoadedFile) Then
        pe.ShowOffsetCalculator
        'frmOffsets.Initilize pe
    Else
        MsgBox "Failed to load pe file: " & pe.errMessage
    End If
    If Err.Number <> 0 Then
        MsgBox "Error: " & Err.Description
    End If
End Sub

Private Sub mnuPEVerInfo_Click()
    frmPEVersion.ShowReport LoadedFile
End Sub

Private Sub mnuSearchFileName_Click()
    Dim f As String
    f = fso.FileNameFromPath(LoadedFile)
    Google f, Me.hwnd
End Sub

Private Sub mnuSearchHash_Click()
    Google myMd5, Me.hwnd
End Sub

Private Sub mnuStrings_Click()
    On Error Resume Next
    Dim exe As String
    exe = App.path & IIf(IsIde(), "\..\", "") & "\shellext.exe"
    Shell exe & " """ & LoadedFile & """ /peek"
End Sub

Private Sub mnuSubmitToVT_Click()
    On Error Resume Next
    Dim vt As String
    vt = App.path & IIf(IsIde(), "\..\", "") & "\virustotal.exe"
    If Not fso.FileExists(vt) Then
        MsgBox "VirusTotal app not found?: " & vt, vbInformation
        Exit Sub
    End If
    Shell vt & " /submit " & LoadedFile
End Sub

Private Sub mnuVt_Click()
    
    On Error Resume Next
    Dim tmp As String
    Dim a As Long
    
    If scan Is Nothing Then
        cmdVT_Click
        Exit Sub
    End If
    
    If scan.HadError Then
        MsgBox scan.BriefReport
    Else
        tmp = Text1.text
        a = InStr(tmp, "Scan_Date")
        If a > 3 Then tmp = Mid(tmp, 1, a - 3)
        frmPEVersion.ShowReport LoadedFile, tmp & vbCrLf & vbCrLf & scan.GetReport()
    End If
    
End Sub

Sub AddExternal(cmd As String)
     
    Dim i As Integer, tmp() As String
    
    cmd = trim(cmd)
    If Len(cmd) = 0 Then Exit Sub
    If VBA.Left(cmd, 1) = "#" Then Exit Sub
    
    tmp = Split(cmd, "=", 2)
    
    If UBound(tmp) <> 1 Then
        MsgBox "Invalid external menu entry. format is menu_text=command_line" & vbCrLf & vbCrLf & cmd
        Exit Sub
    End If
    
    i = mnuExt.Count
    Load mnuExt(i)
    mnuExt(i).Caption = trim(tmp(0))
    mnuExt(i).Visible = True
    mnuExt(i).Tag = trim(tmp(1))
    
End Sub

Private Sub mnuVTCache_Click()
    mnuVTCache.Checked = Not mnuVTCache.Checked
    SaveMySetting "mnuVTCache", mnuVTCache.Checked
    vt.report_cache_dir = IIf(mnuVTCache.Checked, vt_cache, Empty)
End Sub



'Private Sub Timer2_Timer()
'
'    Dim c As Collection
'    Dim w As Cwindow
'
'    'Debug.Print "Timer2"
'
'    If Timer2.Tag = 7 Then      ' x attempts...
'        Timer2.enabled = False
'        Exit Sub
'    End If
'
'    Timer2.Tag = Timer2.Tag + 1
'
'    Set c = ChildWindows()
'    For Each w In c
'        If VBA.Left(w.Caption, 5) = "KANAL" Then
'            'Debug.Print Now & " - found kanal window attaching to " & w.hwnd & " (" & Hex(w.hwnd) & " )"
'            w.Caption = w.Caption & "+" & Timer2.Tag
'            Set kanal = w
'            subclass.AttachMessage w.hWnd, WM_COMMAND
'            'Debug.Print "Disabling timer now..."
'            Timer2.enabled = False
'            Exit Sub
'        End If
'    Next
'
'End Sub
'
'Private Sub subclass_MessageReceived(hWnd As Long, wMsg As Long, wParam As Long, lParam As Long, Cancel As Boolean)
'
'        Dim c2 As Collection, wTv As Cwindow, tmp As String
'
'        'we hooked WM_COMMAND for the kanal window
'        'we are looking for BN_Clicked in the hiword of wParam and our button id (1022) in the low word
'        'since bn_clicked = 0 we are going to take a shortcut and just look for 1022 in wparam
'
'        'Debug.Print Now & " - received WM_COMMAND message for hwnd " & hwnd & " wParam = " & wParam
'
'        If wParam = 1022 Then 'our new copy button
'            If kanal Is Nothing Then Exit Sub
'            If kanal.isValid Then
'                Set wTv = kanal.FindChild("SysTreeView32")
'                If wTv.isValid Then
'                     Set c2 = wTv.CopyRemoteTreeView()
'                     tmp = ColToStr(c2)
'                     Clipboard.Clear
'                     Clipboard.SetText tmp
'                     'kanal.CloseWindow
'                     MsgBox "Saved " & Len(tmp) & " bytes!", vbInformation
'                End If
'            End If
'        End If
'
'End Sub

'Private Sub mnuDllCharAction_Click(index As Integer)
'
'    '0 show current, 1 remove aslr -d, 2 remove dep -n, 3 remove sigcheck -f,
'
'    Dim exe As String
'    Dim f As String
'    Dim opt As String
'    Dim author As String
'    Dim newFile As Boolean
'
'    author = "setdllcharacteristics.exe\nAuthor: Didier Stevens\n" & _
'             "Site: http://didierstevens.com\n" & _
'             "Source: public domain, no Copyright, Use at your own risk\n\n"
'
'    author = Replace(author, "\n", vbCrLf)
'
'    opt = Array("", "-d", "-n", "-f")(index)
'
'    On Error Resume Next
'
'    exe = App.path & "\setdllcharacteristics.exe"
'    If Not fso.FileExists(exe) Then exe = fso.GetParentFolder(App.path) & "\setdllcharacteristics.exe"
'    If Not fso.FileExists(exe) Then
'        MsgBox "setdllcharacteristics.exe not found? " & exe
'        Exit Sub
'    End If
'
'    If fso.GetExtension(LoadedFile) = ".dllmod" Or index = 0 Then 'or about read only..
'        f = LoadedFile
'    Else
'        f = fso.GetParentFolder(LoadedFile) & "\" & fso.GetBaseName(LoadedFile) & ".dllmod"
'        If fso.FileExists(f) Then Kill f
'        FileCopy LoadedFile, f
'        newFile = True
'    End If
'
'    Dim cmd As New CCmdOutput
'
'    If Not cmd.GetCommandOutput(exe, opt & " """ & f & """") Then
'        opt = "Failed to launch setdllcharacteristics.exe"
'    Else
'        opt = cmd.result
'    End If
'
'    opt = "File: " & f & vbCrLf & vbCrLf & opt
'    If index = 0 Then opt = author & opt
'
'    If newFile Then
'        opt = opt & vbCrLf & vbCrLf & "Reload as current?"
'        If MsgBox(opt, vbInformation + vbYesNo) = vbYes Then
'            ShowFileStats f
'        End If
'    Else
'        MsgBox opt, vbInformation
'    End If
'
'
'    'Shell """" & exe & """" & opt & " """ & f & """"
'
'
'
'End Sub


Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim z
    On Error Resume Next
    
    lastX = x
    lastY = Y
    
    'this code below is a little processor heavy but what else are we doing...
    curWord = WordUnderCursor(Text1, x, Y, startPos)
    'Debug.Print curWord
    
    For Each z In links
        If z = curWord Then
            Screen.MousePointer = vbArrow
            Exit Sub
        End If
    Next
    
    Screen.MousePointer = vbNormal
        
End Sub
