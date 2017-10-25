VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmPEVersion 
   Caption         =   "PE Header Version Requirements"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   4665
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox Text1 
      Height          =   1545
      Left            =   90
      TabIndex        =   3
      Top             =   585
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   2725
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmPEVersion.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   90
      TabIndex        =   1
      Top             =   45
      Width           =   3525
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   450
         TabIndex        =   2
         Top             =   45
         Width           =   2850
      End
      Begin VB.Label Label1 
         Caption         =   "Find: "
         Height          =   240
         Left            =   45
         TabIndex        =   4
         Top             =   45
         Width           =   420
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enable On XP"
      Height          =   375
      Left            =   315
      TabIndex        =   0
      Top             =   2205
      Width           =   2895
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
 
Private Sub txtFind_Change()
    Me.Caption = ColorText(Text1, txtFind, vbBlue) & " hits"
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


