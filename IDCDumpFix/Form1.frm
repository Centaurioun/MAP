VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "IAT DumpFix - Generate IDC file from olly dump for CALL PTR and JMP IATs"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8880
   HelpContextID   =   8
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   8880
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkHeader 
      Caption         =   "main()"
      Height          =   285
      Left            =   1800
      TabIndex        =   12
      Top             =   810
      Width           =   3030
   End
   Begin VB.CommandButton cmdPaste 
      Caption         =   "Paste"
      Height          =   375
      Left            =   7695
      TabIndex        =   11
      Top             =   135
      Width           =   1140
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4455
      TabIndex        =   10
      Top             =   135
      Width           =   480
   End
   Begin VB.CheckBox chkshowfails 
      Caption         =   "Show failed lines"
      Height          =   330
      Left            =   45
      TabIndex        =   9
      Top             =   765
      Width           =   1635
   End
   Begin VB.CheckBox chkUniqueOnly 
      Caption         =   "Unique Only (disable for delphi apps)"
      Height          =   330
      Left            =   1845
      TabIndex        =   8
      Top             =   495
      Value           =   1  'Checked
      Width           =   2940
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Make Unk"
      Height          =   255
      Left            =   45
      TabIndex        =   7
      Top             =   495
      Width           =   1335
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "?"
      Height          =   375
      HelpContextID   =   8
      Left            =   3300
      TabIndex        =   6
      Top             =   120
      Width           =   435
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save As"
      Height          =   375
      Left            =   5025
      TabIndex        =   5
      Top             =   120
      Width           =   1395
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Kill Lines Like"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   120
      Width           =   1275
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      TabIndex        =   3
      Text            =   "*EAX,EAX*"
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Copy"
      Height          =   375
      Left            =   6435
      TabIndex        =   2
      Top             =   135
      Width           =   1230
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4920
      Left            =   0
      MultiLine       =   -1  'True
      OLEDropMode     =   1  'Manual
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   1155
      Width           =   8775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate IDC"
      Height          =   420
      Left            =   7650
      TabIndex        =   0
      Top             =   585
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this was made about as quick as they come
'does what it needs to, no frills
'
'Author: david@idefense.com
'
'Purpose: small tool used for speed RE of packed binaries.
'         This tools gives you an easy way to make a disasm
'         readable after you have done a raw dump from memory
'         without requiring the time to rebuild the pe for a
'         clean disasm etc..See ? for more details
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

Dim unique As Collection
Dim dlg As New clsCmnDlg
Dim fso As New CFileSystem2

Private Sub cmdHelp_Click()
    SendKeys "{F1}"
End Sub

Private Sub cmdOpen_Click()
    Dim f As String
    f = dlg.OpenDialog(AllFiles)
    If Len(f) = 0 Then Exit Sub
    Text2 = fso.ReadFile(f)
End Sub

Private Sub cmdPaste_Click()
    Text2 = Clipboard.GetText
End Sub

Private Sub cmdSave_Click()
    On Error Resume Next
    Dim f As String
    dlg.SetCustomFilter "IDC Files", "*.idc"
    f = dlg.SaveDialog(CustomFilter, "", "Save As:")
    If Len(f) = 0 Then Exit Sub
    fso.WriteFile f, Text2
    If Err.Number = 0 Then
        MsgBox "Saved Successfuly", vbInformation
    End If
End Sub

Private Sub Command1_Click()
    On Error Resume Next
    hits = 0
    
    header = "#define UNLOADED_FILE   1" & vbCrLf & _
             "#include <idc.idc>" & vbCrLf & vbCrLf & _
             "static main(void) {" & vbCrLf

    Set unique = New Collection
    
    f = Text2
    f = Split(f, vbCrLf) 'lines
    
    Dim addr As String
    Dim import As String
    Dim fails As String
    
    For i = 0 To UBound(f)
    
        addr = Empty
        import = Empty
        l = f(i)
        
        If Len(Trim(l)) = 0 Then GoTo nextone
        
        If InStr(l, "CALL") > 0 And InStr(l, "PTR") > 0 Then 'style 2
            ImportStyleCallPtr l, addr, import
        ElseIf InStr(l, "CALL") > 0 Then
            ImportStyleCall l, addr, import
        ElseIf InStr(l, "JMP") > 0 Then 'style 1
            ImportStyleJmp l, addr, import
        ElseIf InStr(l, ".") > 0 Then
            PointerTable l, addr, import
        ElseIf InStr(l, "!") > 0 Then ' windbg format
            PointerTable l, addr, import
        ElseIf isIDAImport(l) Then
            IDAImport l, addr, import
        ElseIf isHexRaysGetProc(l) Then
            HexRaysGetProc l, addr, import
        Else
            ' This is to prevent lines that are invalid from producing duplicates later.
            import = ""
        End If

        If Len(import) = 0 Or Len(addr) = 0 Then
            fails = fails & "line:" & i & " = " & l & vbCrLf
            GoTo nextone
        End If

        addr = Replace(addr, "`", Empty) 'windbg x64 splitter
        lZeroTrim addr
        import = import & "_"
        
        Err.Clear
        unique.Add "0x" & addr, CStr(import)
        
        If Err.Number <> 0 And chkUniqueOnly.Value = 0 Then
            base = import
            For j = 1 To 30
                Err.Clear
                unique.Add "0x" & addr, base & "_" & j
                If Err.Number = 0 Then
                    import = base & "_" & j
                    Exit For
                End If
            Next
        End If
        
        If Err.Number = 0 Then
            import = Replace(import, "-", "_") 'some chars are reserved for IDA names
            If Check1.Value Then tmp = tmp & vbTab & "MakeUnkn(0x" & addr & ",1);" & vbCrLf 'MakeName(0X4010E8,  "THISISMYSUB_2");
            tmp = tmp & vbTab & "MakeName(0x" & addr & ",""" & import & """);" & vbCrLf
            hits = hits + 1
        End If
        
        
nextone:
    Next

    note = vbCrLf & "// " & hits & " of " & (UBound(f) + 1) & " lines added" & vbCrLf & vbCrLf
    If chkshowfails.Value = 1 Then note = note & "/* failed lines:" & vbCrLf & fails & "*/" & vbCrLf & vbCrLf
     
    If chkHeader.Value Then
        Text2 = note & header & tmp & "}"
    Else
        Text2 = note & Replace(tmp, vbTab, Empty)
    End If
    
End Sub

Sub lZeroTrim(l)
    On Error GoTo hell
    While Left(l, 1) = "0"
        l = Mid(l, 2)
    Wend
hell:
End Sub

Function isHexRaysGetProc(l) As Boolean
    
    If VBA.Left(l, 6) <> "dword_" Then Exit Function
    If InStr(l, "GetProcAddress") < 1 Then Exit Function
    isHexRaysGetProc = True
     
End Function

Function HexRaysGetProc(fileLine, addrVar, importNameVar)
    ' dword_414F60 = (int)GetProcAddress(v0, aGettempfilenam);
    Dim tmp
    On Error GoTo hell
    
    tmp = Replace(fileLine, vbTab, Empty)
    tmp = Split(Trim(tmp), " ")
    addrVar = Replace(tmp(0), "dword_", Empty)
    importNameVar = tmp(UBound(tmp))
    If Right(importNameVar, 1) = ";" Then importNameVar = Mid(importNameVar, 1, Len(importNameVar) - 1)
    If Right(importNameVar, 1) = ")" Then importNameVar = Mid(importNameVar, 1, Len(importNameVar) - 1)
    If VBA.Left(importNameVar, 1) = "a" Then importNameVar = Mid(importNameVar, 2)
    
    Exit Function
hell:
     ' addrVar = Empty
     ' importNameVar = Empty
     
End Function

Function IDAImport(fileLine, addrVar, importNameVar)
   '000000018000F000  000007FEFDF3B5A0  ADVAPI32:advapi32_SetSecurityDescriptorDacl
   a = InStr(fileLine, " ")
   If a < 2 Then Exit Function
   addrVar = Mid(fileLine, 1, a)
   
   a = InStrRev(fileLine, "_")
   If a < 1 Then Exit Function
   importNameVar = Mid(fileLine, a + 1)
   
End Function

Function isIDAImport(l) As Boolean
    '000000018000F000  000007FEFDF3B5A0  ADVAPI32:advapi32_SetSecurityDescriptorDacl
    
    l = Trim(l)
    If CountOccurances(l, " ") <> 4 Then Exit Function
    If InStr(l, ":") < 1 Then Exit Function
    If InStr(l, "_") < 1 Then Exit Function
    
    isIDAImport = True
    
End Function

Function CountOccurances(it, find) As Integer
    Dim tmp() As String
    If InStr(1, it, find, vbTextCompare) < 1 Then CountOccurances = 0: Exit Function
    tmp = Split(it, find, , vbTextCompare)
    CountOccurances = UBound(tmp)
End Function

Sub PointerTable(fileLine, addrVar, importNameVar)
    ' ollydbg format
    '43434394 >7C91137A  ntdll.RtlDeleteCriticalSection
    ' windbg format
    '000b41bc  7c81473b kernel32!MoveFileExW
    l = Split(fileLine, " ")
    addrVar = l(0)
    importNameVar = l(UBound(l))
    
    ' ollystyle
    a = InStr(importNameVar, ".")
    If a > 0 Then
        importNameVar = Mid(importNameVar, a + 1)
    End If
    
    ' windbg sytle
    a = InStr(importNameVar, "!")
    If a > 0 Then
        importNameVar = Mid(importNameVar, a + 1)
    End If
    
    If KeyExistsInCollection(unique, CStr(importNameVar)) Then
        importNameVar = Empty
        addrVar = Empty
    End If
    
End Sub


'all variables byref modificed here
Sub ImportStyleJmp(fileLine, addrVar, importNameVar)
    '00402A98  FF25 7CF14100  JMP DWORD PTR DS:[41F17C] ; ADVAPI32.AdjustTokenPrivileges
    '--------                                             ------------------------------
    l = Split(fileLine, " ") 'words (we want first(address) and last (api name)
    addrVar = l(0)
    importNameVar = l(UBound(l))
    
    a = InStr(importNameVar, ".")
    If a > 0 Then
        importNameVar = Mid(importNameVar, a + 1)
    End If
    
    If KeyExistsInCollection(unique, CStr(importNameVar)) Then
        importNameVar = Empty
        addrVar = Empty
    End If
    
End Sub



'all variables byref modificed here
Sub ImportStyleCallPtr(fileLine, addrVar, importNameVar)
    '00401000   CALL DWORD PTR DS:[405100]                KERNEL32.FreeConsole
    '                              ------                 --------------------
    
    l = Split(Trim(fileLine), " ") 'words (we want first(address) and last (api name)
    importNameVar = l(UBound(l))
    
    a = InStr(importNameVar, ".")
    If a > 0 Then
        importNameVar = Mid(importNameVar, a + 1)
    End If
    
    a = InStr(fileLine, "[")
    b = InStr(fileLine, "]")
    If a > 0 And b > a Then
        a = a + 1
        addrVar = Mid(fileLine, a, b - a)
    End If
    
    If KeyExistsInCollection(unique, CStr(importNameVar)) Then
        importNameVar = Empty
        addrVar = Empty
    End If
    
End Sub

Sub ImportStyleCall(fileLine, addrVar, importNameVar)
    '00402330   CALL 13_5k.00403F66                       urlmon.URLDownloadToFileA
    '                      --------                       -------------------------
    On Error Resume Next
    
    l = Split(Trim(fileLine), " ") 'words (we want first(address) and last (api name)
    importNameVar = l(UBound(l))
    
    a = InStrRev(importNameVar, ".")
    If a > 0 Then
        importNameVar = Mid(importNameVar, a + 1)
    End If
    
    a = InStr(fileLine, ".")
    If a > 25 Then 'module name not call x.
        a = 0
        b = InStr(1, fileLine, "CALL ")
    Else
        b = InStr(a, fileLine, " ")
    End If
    
    If a > 0 And b > a Then
        a = a + 1
        addrVar = Mid(fileLine, a, b - a)
    ElseIf a < 1 And b > 0 Then
        b = b + 6
        addrVar = Mid(fileLine, b, InStr((b + 1), fileLine, " ") - b)
        addrVar = Replace(addrVar, vbTab, "")
        addrVar = Replace(addrVar, "]", "")
    Else
        addrVar = ""
        importNameVar = ""
        Exit Sub
    End If
        
    If KeyExistsInCollection(unique, CStr(importNameVar)) Then
        importNameVar = Empty
        addrVar = Empty
    End If
    
End Sub



Private Sub Command2_Click()
    Clipboard.Clear
    Clipboard.SetText Text2.Text
End Sub

Function KeyExistsInCollection(c As Collection, val As String) As Boolean
    On Error GoTo nope
    Dim t
    t = c(val)
    KeyExistsInCollection = True
 Exit Function
nope: KeyExistsInCollection = False
End Function
    
Private Sub Form_Load()
    
    On Error Resume Next
    
    Dim c As String
    c = Replace(Command, """", Empty)
    If fso.FileExists(c) Then
        f = c
        Text2 = fso.ReadFile(f)
        Command1_Click
    End If
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With Text2
        .Width = Me.Width - .Left - 200
        .Height = Me.Height - .Top - 500
    End With
End Sub

Private Sub Text2_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Not fso.FileExists(Data.Files(1)) Then
        MsgBox "Files only"
        Exit Sub
    End If

    f = Data.Files(1)
    Text2 = fso.ReadFile(f)

End Sub



 

Private Sub Command3_Click()
    
    If Len(Text1) = 0 Then
        MsgBox "Enter expression to match, uses VB LIKE keyword", vbInformation
        Exit Sub
    End If
    
    tmp = Split(Text2, vbCrLf)
    For i = 0 To UBound(tmp)
        If tmp(i) Like Text1 Then tmp(i) = ""
    Next
    
    tmp = Join(tmp, vbCrLf)
    tmp = Replace(tmp, vbCrLf & vbCrLf, vbCrLf)
    Text2 = tmp
    
    
    
End Sub

Private Sub Text2_DblClick()
    Text2 = Clipboard.GetText
End Sub












