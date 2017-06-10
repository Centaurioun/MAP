VERSION 5.00
Begin VB.Form frmPEVersion 
   Caption         =   "PE Header Version Requirements"
   ClientHeight    =   2385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3765
   LinkTopic       =   "Form1"
   ScaleHeight     =   2385
   ScaleWidth      =   3765
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Enable On XP"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1800
      Width           =   2895
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
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmPEVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LoadedFile As String
Dim LastError As String

Sub ShowReport(fPath As String, Optional text As String)

    LoadedFile = fPath
    
    If Len(text) = 0 Then
        Text1 = PEVersionReport
    Else
        Me.Caption = "Data Display"
        Command1.Visible = False
        Text1 = text
        RestoreFormSizeAnPosition Me
        SetWindowTopMost Me
    End If
    
    Me.Visible = True
End Sub
 
Private Sub Command1_Click()
    If ResetVersionsBetaTest(LoadedFile) Then
        MsgBox "Ok we generated a test image now named " & fso.FileNameFromPath(LoadedFile) & ".xp", vbInformation
    Else
        MsgBox "Error: " & LastError, vbInformation
    End If
End Sub

Function ResetVersionsBetaTest(ByVal fPath As String) As Boolean
    On Error GoTo hell
        
        Dim i As Long
        Dim f As Long
        Dim buf(20) As Byte
        Dim sBuf As String
        Dim fs As Long
        
        Dim DOSHEADER As IMAGEDOSHEADER
        Dim NTHEADER As IMAGE_NT_HEADERS
        Dim opt As IMAGE_OPTIONAL_HEADER
        Dim opt64 As IMAGE_OPTIONAL_HEADER_64
        
        LastError = Empty
        If Not fso.FileExists(fPath) Then
            LastError = "File not found"
            Exit Function
        End If
        
        fs = DisableRedir()
        
        Dim newF As String
        newF = fPath & ".xp"
        If fso.FileExists(newF) Then Kill newF
        FileCopy fPath, newF
        fPath = newF
        
        f = FreeFile
        
        Open fPath For Binary As f
        Get f, , DOSHEADER
        
        If DOSHEADER.e_magic <> &H5A4D Then
            RevertRedir fs
            LastError = "MZ header not found"
            Exit Function
        End If
        
        Get f, DOSHEADER.e_lfanew + 1, NTHEADER
        
        If NTHEADER.Signature <> "PE" & Chr(0) & Chr(0) Then
            RevertRedir fs
            LastError = "PE header not found"
            Exit Function
        End If
        
        If is64Bit(NTHEADER.FileHeader.Machine) Then
            Get f, , opt64
            With opt64
                .MajorOperatingSystemVersion = 6
                .MinorOperatingSystemVersion = 0
                .MajorImageVersion = 1
                .MinorImageVersion = 0
                .MajorSubsystemVersion = 4
                .MinorSubsystemVersion = 0
                .MajorLinkerVersion = 6
                .MinorLinkerVersion = 0
            End With
            Seek f, DOSHEADER.e_lfanew + 1 + Len(NTHEADER)
            Put f, , opt64
        Else
            Get f, , opt
            With opt
                .MajorOperatingSystemVersion = 6
                .MinorOperatingSystemVersion = 0
                .MajorImageVersion = 1
                .MinorImageVersion = 0
                .MajorSubsystemVersion = 4
                .MinorSubsystemVersion = 0
                .MajorLinkerVersion = 6
                .MinorLinkerVersion = 0
            End With
            Seek f, DOSHEADER.e_lfanew + 1 + Len(NTHEADER)
            Put f, , opt
        End If
        
        Close f
        RevertRedir fs
        ResetVersionsBetaTest = True
        
Exit Function
hell:
    LastError = Err.Description
    Close f
    RevertRedir fs
End Function

Private Sub Form_Resize()
    On Error Resume Next
    If Command1.Visible Then Exit Sub
    '    Command1.top = Me.Height - Command1.Height - 200
    '    Text1.Height = Me.Height - Text1.top - Command1.top - 200
    'Else
        Text1.Height = Me.Height - Text1.top - 200
    'End If
    Text1.Width = Me.Width - Text1.Left - 200
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormSizeAnPosition Me
End Sub
