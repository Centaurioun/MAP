VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Install Shell Extensions"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer tmrCloseWithHexEditor 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   3945
   End
   Begin VB.PictureBox pict 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   3195
      Left            =   120
      Picture         =   "frmMain.frx":030A
      ScaleHeight     =   3135
      ScaleWidth      =   5505
      TabIndex        =   6
      Top             =   60
      Width           =   5565
   End
   Begin VB.Frame Frame1 
      Caption         =   " Options "
      Height          =   1110
      Left            =   90
      TabIndex        =   2
      Top             =   3300
      Width           =   5595
      Begin VB.CheckBox chkUseSha256 
         Caption         =   "Use SHA256 as Default"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   855
         TabIndex        =   7
         Top             =   630
         Width           =   2940
      End
      Begin VB.CommandButton cmdMinLen 
         Caption         =   "Set"
         Height          =   285
         Left            =   4680
         TabIndex        =   5
         Top             =   210
         Width           =   795
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4050
         TabIndex        =   4
         Text            =   "4"
         Top             =   210
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Strings min match length"
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
         Index           =   0
         Left            =   870
         TabIndex        =   3
         Top             =   225
         Width           =   2955
      End
   End
   Begin VB.CommandButton cmdInstallRegKeys 
      Caption         =   "Install"
      Height          =   315
      Left            =   4635
      TabIndex        =   1
      Top             =   4530
      Width           =   1035
   End
   Begin VB.CommandButton cmdRemoveRegKeys 
      Caption         =   "Remove"
      Height          =   315
      Left            =   3240
      TabIndex        =   0
      Top             =   4530
      Width           =   1035
   End
   Begin VB.Label Label2 
      Caption         =   "Shell extensions:"
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
      Index           =   1
      Left            =   1035
      TabIndex        =   8
      Top             =   4575
      Width           =   2100
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'Author: dzzie@yahoo.com
'
'Purpose: small utility to add 3 shell extensions to explorer right click
'         context menus.
'
'         1) "Strings" contex menu item added for files
'                reads through the file and extracts all ascii and unicode strings
'                matching minimum predefined length. Results displayed in a popup form.
'                Uses the MS VBscript Regexp library should be pretty quick.
'
'         2) "Hash Files" contex menu item added for folders
'                enumerates all files in folder and pops up a form listing their names,
'                file sizes, and MD5 hash values. Also allows you to delete files from
'                the UI. Very useful for sorting directories full of malcode sample which
'                may contain duplicates.
'
'         3) "Decompile" context menu item added for chm files
'               this uses the -decompile option for hh.exe to decompile
'               the chm file you select into ./chm_src
'
'         4) "MD5 Hash" context menu added for all file types (added 12.15.05)
'               -bug fix 9/7/07 some ms service pack broke my vbdevkit md5 code..fixed now :-\
'
'         5) "Virus Total" context menu added for all file types (added 4-19-12)
'
'         6) "Submit to VirusTotal" context menu added for all file types (added 11-11-13)
'
'
'License: Copyright (C) 2005 David Zimmer <david@idefense.com, dzzie@yahoo.com>
'
'         This program is free software; you can redistribute it and/or modify it
'         under the terms of the GNU General Public License as published by the Free
'         Software Foundation; either version 2 of the License, or (at your option)
'         any later version.
'
'         This program is distributed in the hope that it will be useful, but WITHOUT
'         ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or
'         FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for
'         more details.
'
'         You should have received a copy of the GNU General Public License along with
'         this program; if not, write to the Free Software Foundation, Inc., 59 Temple
'         Place, Suite 330, Boston, MA 02111-1307 USA

Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long

Const copy = "*\shell\Copy Path\command"
Const copy2 = "Folder\shell\Copy Path\command"

Const peek = "*\shell\Strings\command"
Const hash = "Folder\shell\Hash Files\command"
Const hSearch = "Folder\shell\Hash Search\command"
Const hashSets = "Folder\shell\Compare HashSets\command"
Const deco = "chm.file\shell\Decompile\command"
Const m5 = "*\shell\Md5 Hash\command"
Const vt = "*\shell\Virus Total\command"
Const vtsubmit = "*\shell\Submit to VirusTotal\command"
Const tlb = "dllfile\shell\Type Library Viewer\command"
Const tlb2 = "ocxfile\shell\Type Library Viewer\command"
Const tlb3 = "tlbfile\shell\Type Library Viewer\command"

Private autoInstall As Boolean
Private formLoaded As Boolean

Function ap() As String
    ap = App.path
    If isIde() Then ap = fso.GetParentFolder(ap)
End Function

Sub InstallRegKeys()
    
    Dim cmdline_1 As String
    Dim cmdline_2 As String
    Dim cmdline_3 As String
    Dim cmdline_4 As String
    Dim cmdline_5 As String
    Dim cmdline_6 As String
    Dim cmdline_7 As String
    Dim cmdline_8 As String
    Dim cmdline_9 As String
    Dim cmdline_10 As String
    
    Dim reg As New clsRegistry2
    
    'note app.path will be wrong value to use in IDE unless you actually compile
    'a version to app.path, default compile dir is /source dir/../
    
    cmdline_1 = """" & ap() & "\shellext.exe"" ""%1"" /peek"
    cmdline_2 = """" & ap() & "\shellext.exe"" ""%1"" /hash"
    cmdline_3 = """" & ap() & "\shellext.exe"" ""%1"" /deco"
    cmdline_4 = """" & ap() & "\shellext.exe"" ""%1"" /md5f"
    cmdline_5 = """" & ap() & "\virustotal.exe"" ""%1"""
    cmdline_6 = """" & ap() & "\virustotal.exe"" ""%1"" /submit"
    cmdline_7 = """" & ap() & "\shellext.exe"" ""%1"" /hsch"
    cmdline_8 = """" & ap() & "\tlbViewer.exe"" ""%1"""
    cmdline_9 = """" & ap() & "\shellext.exe"" ""%1"" /hset"
    cmdline_10 = """" & ap() & "\shellext.exe"" ""%1"" /copy"
    
    On Error GoTo hell
    
    reg.hive = HKEY_CLASSES_ROOT
    
    If reg.CreateKey(copy) Then
        reg.SetValue copy, "", cmdline_10, REG_SZ
    Else
        If Not autoInstall Then
            MsgBox "You may not have permission to write to HKCR", vbExclamation
        End If
        Exit Sub
    End If
    
    If reg.CreateKey(copy2) Then
        reg.SetValue copy2, "", cmdline_10, REG_SZ
    Else
        If Not autoInstall Then
            MsgBox "You may not have permission to write to HKCR", vbExclamation
        End If
        Exit Sub
    End If
    
    If reg.CreateKey(peek) Then
        reg.SetValue peek, "", cmdline_1, REG_SZ
    Else
        If Not autoInstall Then
            MsgBox "You may not have permission to write to HKCR", vbExclamation
        End If
        Exit Sub
    End If
    
    If reg.CreateKey(hash) Then
        reg.SetValue hash, "", cmdline_2, REG_SZ
    End If
    
    If reg.CreateKey(deco) Then
        reg.SetValue deco, "", cmdline_3, REG_SZ
    End If
    
    If reg.CreateKey(m5) Then
        reg.SetValue m5, "", cmdline_4, REG_SZ
    End If
    
    If reg.CreateKey(vt) Then
        reg.SetValue vt, "", cmdline_5, REG_SZ
    End If
    
    If reg.CreateKey(vtsubmit) Then
        reg.SetValue vtsubmit, "", cmdline_6, REG_SZ
    End If
    
    If reg.CreateKey(hSearch) Then
        reg.SetValue hSearch, "", cmdline_7, REG_SZ
    End If
    
    If reg.CreateKey(tlb) Then
        reg.SetValue tlb, "", cmdline_8, REG_SZ
    End If
    
    If reg.CreateKey(tlb2) Then
        reg.SetValue tlb2, "", cmdline_8, REG_SZ
    End If
    
    If reg.CreateKey(tlb3) Then
        reg.SetValue tlb3, "", cmdline_8, REG_SZ
    End If
    
    If reg.CreateKey(hashSets) Then
        reg.SetValue hashSets, "", cmdline_9, REG_SZ
    End If
    
    If Not autoInstall Then MsgBox "Entries Added", vbInformation
    End
    
hell: If Not autoInstall Then MsgBox "Error adding keys: " & Err.Description

End Sub


Private Sub chkUseSha256_Click()
    If Not formLoaded Then Exit Sub
    SaveMySetting "mnuUseSHA256.Checked", chkUseSha256.value
End Sub

Private Sub cmdInstallRegKeys_Click()
    
    If IsVistaPlus() Then
        If Not IsUserAnAdministrator() Then
            MsgBox "Must be an admin user to install these settings can not elevate.", vbInformation
        Else
            RunElevated App.path & "\shellext.exe", essSW_HIDE, , "/install"
        End If
    Else
        InstallRegKeys
    End If
                
End Sub

Private Sub cmdRemoveRegKeys_Click()
    
    If IsVistaPlus() Then
        If Not IsUserAnAdministrator() Then
            MsgBox "Must be an admin user to remove these settings can not elevate.", vbInformation
        Else
             RunElevated App.path & "\shellext.exe", essSW_HIDE, , "/remove"
        End If
    Else
        RemoveRegKeys
    End If
    
End Sub


Private Sub cmdMinLen_Click()

    If Not IsNumeric(Text1) Then
        MsgBox "String Length must be numeric", vbInformation
        Exit Sub
    End If
    
    On Error Resume Next
    minStrLen = CLng(Text1)
    If Len(minStrLen) = 0 Then minStrLen = 4
    SaveMySetting "minStrLen", minStrLen
    
End Sub

Function RemoveRegKeys()

    Dim reg As New clsRegistry2
    Dim a As Boolean, b As Boolean, c As Boolean
    
    reg.hive = HKEY_CLASSES_ROOT
    
    a = True: b = True: c = True
    
    If reg.keyExists(peek) Then
        a = reg.DeleteKey(peek)
        a = reg.DeleteKey("*\shell\Strings\")
    End If
    
    If reg.keyExists(m5) Then
        a = reg.DeleteKey(m5)
        a = reg.DeleteKey("*\shell\Md5 Hash\")
    End If
    
    If reg.keyExists(vt) Then
        a = reg.DeleteKey(vt)
        a = reg.DeleteKey("*\shell\Virus Total\")
    End If
    
    If reg.keyExists(vtsubmit) Then
        a = reg.DeleteKey(vtsubmit)
        a = reg.DeleteKey("*\shell\Submit to VirusTotal\")
    End If
    
    If reg.keyExists(hash) Then
        b = reg.DeleteKey(hash)
        b = reg.DeleteKey("Folder\shell\Hash Files")
    End If
    
    If reg.keyExists(hSearch) Then
        b = reg.DeleteKey(hSearch)
        b = reg.DeleteKey("Folder\shell\Hash Search")
    End If
    
    If reg.keyExists(hashSets) Then
        b = reg.DeleteKey(hashSets)
        b = reg.DeleteKey("Folder\shell\Compare HashSets")
    End If
    
    If reg.keyExists(deco) Then
       c = reg.DeleteKey(deco)
       c = reg.DeleteKey("chm.file\shell\Decompile")
    End If
    
    If reg.keyExists(tlb) Then
       c = reg.DeleteKey(tlb)
       c = reg.DeleteKey("dllfile\shell\Type Library Viewer\")
    End If
    
    If reg.keyExists(tlb2) Then
       c = reg.DeleteKey(tlb2)
       c = reg.DeleteKey("ocxfile\shell\Type Library Viewer\")
    End If
    
    If reg.keyExists(tlb3) Then
       c = reg.DeleteKey(tlb3)
       c = reg.DeleteKey("tlbfile\shell\Type Library Viewer\")
    End If
    
    If reg.keyExists(copy) Then
        a = reg.DeleteKey(copy)
        a = reg.DeleteKey("*\shell\Copy Path\")
    End If
    
    If reg.keyExists(copy2) Then
        a = reg.DeleteKey(copy2)
        a = reg.DeleteKey("Folder\shell\Copy Path\")
    End If
    
    If Not autoInstall Then
        If a And b And c Then
            MsgBox "Keys deleted        ", vbInformation
        Else
            MsgBox "Could not delete all regkeys", vbExclamation
        End If
    End If
    
    End
    
End Function

Private Sub Form_Load()

    On Error Resume Next
    
    Dim mode As Long
    Dim cmd As String
    Dim lastCmd As String
    Dim isLastCmd As Boolean
    
    'MsgBox "frmmain.load"
    
    Set myIcon = Me.Icon 'this prevents sub forms from accidently recalling frmMain.form_load if it unloads, but they want to use its main icon as their own..
    
    pict.CurrentY = 10
    pict.Print " All files: " & vbCrLf & _
               "    Strings" & vbCrLf & _
               "    Md5 Hash" & vbCrLf & _
               "    VirusTotal" & vbCrLf & _
               "    Submit to VT" & vbCrLf & _
               "" & vbCrLf & _
               " All folders:" & vbCrLf & _
               "    Hash Files" & vbCrLf & _
               "    Hash Search" & vbCrLf & _
               "    Compare HashSets" & vbCrLf & _
               "" & vbCrLf & _
               " Dll/OCX/TLB Files: " & vbCrLf & _
               "    Type Library Viewer" & vbCrLf & _
               "" & vbCrLf & _
               " CHM Files: Decompile"

                 
    lastCmd = GetMySetting("lastCMD", "")
    
    If isIde() And Len(lastCmd) > 0 Then
        cmd = Replace(lastCmd, """", "")
        isLastCmd = True
    Else
        cmd = Replace(Command, """", "")
    End If
    
    On Error Resume Next
    minStrLen = CLng(GetMySetting("minStrLen", 4))
    If Len(minStrLen) = 0 Then minStrLen = 4
    Text1 = minStrLen
    
    'frmFileHash.ShowFileStats "c:\windows\notepad.exe"
    'Exit Sub
    
    'frmStrings.ParseFile "c:\peEditor.exe"
    'Exit Sub
    
    If Len(cmd) > 0 Then
        If VBA.Right(cmd, 5) = "/peek" Then mode = 1
        If VBA.Right(cmd, 5) = "/hash" Then mode = 2
        If VBA.Right(cmd, 5) = "/deco" Then mode = 3
        If VBA.Right(cmd, 5) = "/md5f" Then mode = 4
        If VBA.Right(cmd, 5) = "/hsch" Then mode = 7
        If VBA.Right(cmd, 5) = "/hexv" Then mode = 8 'if /base is supplied it must be before /hexv
        If VBA.Right(cmd, 5) = "/hset" Then mode = 9
        If VBA.Right(cmd, 5) = "/copy" Then mode = 10
        
        If VBA.Right(cmd, 8) = "/install" Then
            mode = 5 'required for Vista run elevated mode
            autoInstall = True
        End If
        
        If VBA.Right(cmd, 7) = "/remove" Then
            mode = 6
            autoInstall = True
        End If
        
        cmd = trim(Mid(cmd, 1, Len(cmd) - 5)) '<-- ** this is why those with path args are 5 chars long **
        
        If mode = 5 Or mode = 6 Then
            If IsVistaPlus() And Not IsProcessElevated() Then
                MsgBox "Process must be elevated for this option to work..", vbInformation
                'but cant hurt to try it anyway right...
            End If
        End If
        
        If Not isLastCmd Then SaveMySetting "lastCmd", Command
        
        Select Case mode
            Case 1: frmStrings.ParseFile cmd
            Case 2: frmHash.HashDir cmd
            Case 3: DecompileChm cmd
            Case 4: frmFileHash.ShowFileStats cmd
            Case 5: InstallRegKeys
            Case 6: RemoveRegKeys
            Case 7: frmMD5FileSearch.Launch cmd
            Case 8: frmHexView.HexView cmd
            Case 9: frmCompareHashSets.Show
            Case 10: Clipboard.Clear: Clipboard.SetText Replace(cmd, """", Empty)
            Case Else: MsgBox "Unknown Option: " & Command & vbCrLf & "Last5 = " & Right(cmd, 5), vbExclamation
        End Select
        
       Unload Me
        
    Else
        chkUseSha256.value = IIf(CBool(GetMySetting("mnuUseSHA256.Checked", 0)), 1, 0)
        Me.Visible = True
        formLoaded = True
    End If
    
    
End Sub



Sub DecompileChm(pth As String)
    On Error GoTo hell
    
    Dim pf As String
    Dim cmd As String
    Dim tmp As String
    Dim fn As String
    
    pf = fso.GetParentFolder(pth)
        
    If InStr(pth, " ") < 1 Then
            pf = pf & "\chm_src"
    Else 'hh bugs! cant handle spaces in path or " this sucks...
    
        tmp = Environ("TEMP")
        If Len(tmp) = 0 Then
            tmp = Environ("TMP")
            If Len(tmp) = 0 Then
                MsgBox "Chm path has space char in it and Enviroment variable TEMP not set sorry exiting"
                Exit Sub
            End If
        End If
        
        If Not fso.FolderExists(tmp) Then
            MsgBox "TEMP variable points to invalid directory?"
            Exit Sub
        End If
        
        fn = fso.FileNameFromPath(pth)
        If InStr(fn, " ") > 0 Then fn = Replace(fn, " ", "")
        
        fn = tmp & "\" & fn
        If fso.FileExists(fn) Then Kill fn
        FileCopy pth, fn
        
        tmp = tmp & "\chm_src"
        tmp = Replace(tmp, "\\", "\")
        If fso.FolderExists(tmp) Then fso.DeleteFolder tmp
        
        pf = tmp
        pth = fn
    End If
    
    If Not fso.FolderExists(pf) Then MkDir pf
    
    cmd = "hh -decompile " & pf & " " & pth
    'InputBox "", , cmd
    
    Shell cmd
    Shell "explorer " & pf, vbNormalFocus
    
    Exit Sub
hell: MsgBox "Error Decompiling CHM: " & Err.Description
End Sub

 
