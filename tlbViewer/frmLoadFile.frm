VERSION 5.00
Begin VB.Form frmLoadFile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select COM Server"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   6750
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optProgIDScan 
      Caption         =   "Search by ProgID"
      Height          =   255
      Left            =   2325
      TabIndex        =   6
      Top             =   1755
      Width           =   2355
   End
   Begin VB.OptionButton optScanDirectory 
      Caption         =   "Scan a directory for registered COM servers"
      Height          =   315
      Left            =   2325
      TabIndex        =   5
      Top             =   900
      Width           =   4035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3315
      Left            =   120
      Picture         =   "frmLoadFile.frx":0000
      ScaleHeight     =   3315
      ScaleWidth      =   1815
      TabIndex        =   4
      Top             =   60
      Width           =   1815
   End
   Begin VB.CommandButton cndNext 
      Caption         =   "Next"
      Height          =   375
      Left            =   5445
      TabIndex        =   3
      Top             =   2925
      Width           =   1095
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Manually Enter the GUID "
      Height          =   255
      Left            =   2340
      TabIndex        =   2
      Top             =   1350
      Width           =   2415
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Choose ActiveX dll or ocx file directly"
      Height          =   255
      Left            =   2325
      TabIndex        =   1
      Top             =   540
      Value           =   -1  'True
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Step 1 - Select the COM Server you wish to test. "
      Height          =   255
      Left            =   1980
      TabIndex        =   0
      Top             =   60
      Width           =   3615
   End
End
Attribute VB_Name = "frmLoadFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'License:   GPL
'Copyright: 2005 iDefense a Verisign Company
'Site:      http://labs.idefense.com
'
'Author:  David Zimmer <david@idefense.com, dzzie@yahoo.com>
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

Private Declare Function ExpandEnvironmentStrings _
   Lib "kernel32" Alias "ExpandEnvironmentStringsA" _
   (ByVal lpSrc As String, ByVal lpDst As String, _
   ByVal nSize As Long) As Long
   
Private Sub Form_Load()
    
    Dim c As String
    c = Replace(Command, """", "")
    If fso.FileExists(c) Then
        frmtlbViewer.LoadFile c
        Unload Me
    Else
        If Len(c) > 0 Then
            MsgBox "File not found: " & c, vbInformation
        End If
    End If
    
End Sub
   
Public Function ResolvePath(strInput As String) As String
    On Error Resume Next
    Dim result As Long
    Dim strOutput As String
    '' Two calls required, one to get expansion buffer length first then do expansion
    result = ExpandEnvironmentStrings(strInput, strOutput, result)
    strOutput = Space$(result)
    result = ExpandEnvironmentStrings(strInput, strOutput, result)
    If result > 1 Then strOutput = Mid(strOutput, 1, result - 1)
    ResolvePath = strOutput
End Function

Private Sub cndNext_Click()
    Dim f As String, ff As String
    Dim Files() As String
    Dim key As String
    
    If Option1.value Then 'select file
        f = dlg.OpenDialog(AllFiles)
        If Len(f) = 0 Then Exit Sub
        frmtlbViewer.LoadFile f
        Unload Me
        
    ElseIf Option2.value Then 'load from guid
        
        ff = Trim(InputBox("Enter GUID you wish to analyze")) ', , "{4CECCEB2-8359-11D0-A34E-00AA00BDCDFD}")) '"05589FA1-C356-11CE-BF01-00AA0055595A"))
        If Len(ff) = 0 Then Exit Sub
        
        reg.hive = HKEY_CLASSES_ROOT
        
        If Right(ff, 1) <> "}" Then ff = ff & "}"
        If Left(ff, 1) <> "{" Then ff = "{" & ff
        
        f = "\CLSID\" & ff
        
        If reg.keyExists(f) Then
            f = f & "\InProcServer32"
            If reg.keyExists(f) Then
                f = reg.ReadValue(f, "")
                f = StripQuotes(f)
                If InStr(f, "%") > 0 Then f = ResolvePath(f)
                If fso.FileExists(f) Then
                    frmtlbViewer.LoadFile f, ff
                Else
                    MsgBox "Clsid found, but file not: " & f
                End If
            Else
                MsgBox "Could not find its InProcServer32 entry", vbInformation
            End If
        Else
            MsgBox "Could not locate this GUID on your system", vbInformation
        End If
        
    ElseIf optScanDirectory.value Then
        
        f = dlg.FolderDialog
        If Len(f) = 0 Then Exit Sub
        If Not frmScanDir.ShowServersForPath(f) Then
            MsgBox "No COM Servers found in " & f, vbInformation
        Else
            Unload Me
        End If
        
    End If
    
End Sub


