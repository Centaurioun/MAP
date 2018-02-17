VERSION 5.00
Begin VB.Form frmDllCharacteristics 
   Caption         =   "Modify OptionalHeader.DLLCharacteristics"
   ClientHeight    =   2010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   ScaleHeight     =   2010
   ScaleWidth      =   5250
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   660
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   60
      Width           =   4455
   End
   Begin VB.CheckBox chkAddExt 
      Caption         =   "Add .mod ext"
      Height          =   330
      Left            =   720
      TabIndex        =   10
      Top             =   1440
      Value           =   1  'Checked
      Width           =   1590
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   465
      Left            =   2520
      TabIndex        =   9
      Top             =   1350
      Width           =   1275
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   465
      Left            =   3960
      TabIndex        =   8
      Top             =   1350
      Width           =   1185
   End
   Begin VB.CheckBox chkIntegrity 
      Caption         =   "Force Integrity"
      Height          =   285
      Left            =   2835
      TabIndex        =   7
      Top             =   900
      Width           =   1410
   End
   Begin VB.CheckBox chkASLR 
      Caption         =   "ASLR"
      Height          =   330
      Left            =   720
      TabIndex        =   6
      Top             =   855
      Width           =   870
   End
   Begin VB.CheckBox chkDep 
      Caption         =   "DEP"
      Height          =   285
      Left            =   1845
      TabIndex        =   5
      Top             =   900
      Width           =   780
   End
   Begin VB.Label Label4 
      Caption         =   "File: "
      Height          =   195
      Left            =   60
      TabIndex        =   11
      Top             =   120
      Width           =   435
   End
   Begin VB.Label Label3 
      Caption         =   "Flags:"
      Height          =   240
      Left            =   90
      TabIndex        =   4
      Top             =   900
      Width           =   510
   End
   Begin VB.Label lblNew 
      Height          =   285
      Left            =   3330
      TabIndex        =   3
      Top             =   435
      Width           =   1275
   End
   Begin VB.Label Label2 
      Caption         =   "New: "
      Height          =   285
      Left            =   2835
      TabIndex        =   2
      Top             =   435
      Width           =   420
   End
   Begin VB.Label lblCur 
      Height          =   285
      Left            =   855
      TabIndex        =   1
      Top             =   435
      Width           =   1230
   End
   Begin VB.Label Label1 
      Caption         =   "Current:"
      Height          =   240
      Left            =   60
      TabIndex        =   0
      Top             =   435
      Width           =   645
   End
End
Attribute VB_Name = "frmDllCharacteristics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim curVal As Integer
Dim finalOffset As Long
Dim newVal As Integer
Dim chkActionEnabled As Boolean
Dim LoadedFile As String

'Thanks to: https://blog.didierstevens.com/2010/10/17/setdllcharacteristics/
'AllFlags: https://github.com/0xd4d/dnlib/blob/master/src/PE/DllCharacteristics.cs

Property Get DEP() As Boolean
    DEP = ((newVal And &H100) = &H100)
    lblNew = "0x" & Hex(newVal)
End Property

Property Get ASLR() As Boolean
    ASLR = ((newVal And &H40) = &H40)
    lblNew = "0x" & Hex(newVal)
End Property

Property Get FORCE_INTEGRITY() As Boolean
    FORCE_INTEGRITY = ((newVal And &H80) = &H80)
    lblNew = "0x" & Hex(newVal)
End Property

Property Let DEP(v As Boolean)
    If v Then
        newVal = newVal Or &H100
    Else
        newVal = newVal Xor &H100
    End If
    lblNew = "0x" & Hex(newVal)
End Property

Property Let ASLR(v As Boolean)
    If v Then
        newVal = newVal Or &H40
    Else
        newVal = newVal Xor &H40
    End If
    lblNew = "0x" & Hex(newVal)
End Property

Property Let FORCE_INTEGRITY(v As Boolean)
    If v Then
        newVal = newVal Or &H80
    Else
        newVal = newVal Xor &H80
    End If
    lblNew = "0x" & Hex(newVal)
End Property

Public Function LoadFile(pth As String)
     
    Dim f As Long, mz As Integer, peOffset As Long, peSig As Integer
    Const mz_pe_offset = &H3C
    Const dllChar_offset = &H5E
    
    On Error GoTo hell
    
    Me.Visible = True
    finalOffset = 0
    curVal = 0
    newVal = 0
    lblCur = Empty
    lblNew = Empty
    chkDep.value = 0
    chkASLR.value = 0
    chkIntegrity.value = 0
    chkActionEnabled = False
    LoadedFile = Empty
    txtFile = pth
    cmdSave.Enabled = False
    
    If Not FileExists(pth) Then
        Caption = "File not found: " & pth
        Exit Function
    End If
    
    LoadedFile = pth
    
    f = FreeFile
    Open pth For Binary Access Read As f
    Get f, , mz
    
    If mz <> &H5A4D Then
        Caption = "MZ marker not found"
        GoTo hell
    End If
    
    Get f, mz_pe_offset + 1, peOffset
    Get f, peOffset + 1, peSig
    
    If peSig <> &H4550 Then
        Caption = "PE marker not found"
        GoTo hell
    End If
    
    finalOffset = peOffset + dllChar_offset + 1
    Get f, finalOffset, curVal
    newVal = curVal
    lblCur = "0x" & Hex(curVal)
    
    If DEP Then chkDep.value = 1
    If ASLR Then chkASLR.value = 1
    If FORCE_INTEGRITY Then chkIntegrity.value = 1
    chkActionEnabled = True
    cmdSave.Enabled = True
    
hell:
    If Err.Number <> 0 Then Caption = "Err: " & Err.Description
    Close f

End Function





Private Function FileExists(path As String) As Boolean
  On Error GoTo hell
    
  If Len(path) = 0 Then Exit Function
  If Right(path, 1) = "\" Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
  
  Exit Function
hell: FileExists = False
End Function


Private Sub chkASLR_Click()
    If Not chkActionEnabled Then Exit Sub
    ASLR = (chkASLR.value = 1)
End Sub

Private Sub chkDep_Click()
    If Not chkActionEnabled Then Exit Sub
    DEP = (chkDep.value = 1)
End Sub

Private Sub chkIntegrity_Click()
    If Not chkActionEnabled Then Exit Sub
    FORCE_INTEGRITY = (chkIntegrity.value = 1)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    
    If curVal = newVal Then
        MsgBox "No Changes", vbInformation
        Exit Sub
    End If
    
    On Error GoTo hell
    If Not FileExists(LoadedFile) Then
        MsgBox "File not found", vbInformation
        Exit Sub
    End If
    
    Dim f As Long, pth As String, m As VbMsgBoxResult
    
    If chkAddExt.value = 1 And LCase(Right(LoadedFile, 4)) <> ".mod" Then
        pth = LoadedFile & ".mod"
        If FileExists(pth) Then
            m = MsgBox(".mod file already exists should I delete it?", vbYesNo)
            If m = vbNo Then
                Caption = "Aborted delete .mod file first"
                Exit Sub
            End If
            If m = vbYes Then Kill pth
        End If
        FileCopy LoadedFile, pth
    Else
        pth = LoadedFile
    End If
        
    f = FreeFile
    Open pth For Binary As f
    Put f, finalOffset, newVal
    Close f
    
    LoadFile pth
    
    Exit Sub
hell:
    If f <> 0 Then Close f
    Me.Caption = "Err: " & Err.Description
    
End Sub

 
