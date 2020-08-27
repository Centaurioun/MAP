VERSION 5.00
Begin VB.Form frmApiKey 
   Caption         =   "Set ApiKey"
   ClientHeight    =   1980
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8715
   LinkTopic       =   "Form3"
   ScaleHeight     =   1980
   ScaleWidth      =   8715
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   330
      Left            =   7200
      TabIndex        =   5
      Top             =   1530
      Width           =   1410
   End
   Begin VB.TextBox txtUserPrivateKey 
      Height          =   285
      Left            =   2070
      TabIndex        =   4
      Top             =   1080
      Width           =   6450
   End
   Begin VB.TextBox txtUserPublicKey 
      Height          =   330
      Left            =   2070
      TabIndex        =   3
      Top             =   675
      Width           =   6450
   End
   Begin VB.OptionButton optKey 
      Caption         =   "Private Key"
      Height          =   240
      Index           =   2
      Left            =   135
      TabIndex        =   2
      Top             =   1125
      Width           =   1815
   End
   Begin VB.OptionButton optKey 
      Caption         =   "Personal public key"
      Height          =   330
      Index           =   1
      Left            =   135
      TabIndex        =   1
      Top             =   675
      Width           =   1770
   End
   Begin VB.OptionButton optKey 
      Caption         =   "Default Public key"
      Height          =   375
      Index           =   0
      Left            =   135
      TabIndex        =   0
      Top             =   180
      Value           =   -1  'True
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   $"frmApiKey.frx":0000
      Height          =   420
      Left            =   495
      TabIndex        =   6
      Top             =   1530
      Width           =   6045
   End
End
Attribute VB_Name = "frmApiKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vt As CVirusTotal

Sub init(owner As CVirusTotal)
    Set vt = owner
    Me.Visible = True
End Sub

Private Sub Command1_Click()
    Dim i As Long

    SaveSetting "VirusTotal", "config", "key_option", 0
    For i = 0 To optKey.count
        If optKey(i).Value Then
            SaveSetting "VirusTotal", "config", "key_option", i
            Exit For
        End If
    Next
    
    txtUserPrivateKey = Trim(txtUserPrivateKey)
    txtUserPublicKey = Trim(txtUserPublicKey)
     
    Select Case i
        Case 1:
                If Len(txtUserPublicKey) = 0 Then
                    MsgBox "Must enter a User Public key"
                    Exit Sub
                End If
                
        Case 2:
                If Len(txtUserPrivateKey) = 0 Then
                    MsgBox "Must enter a Private key"
                    Exit Sub
                End If
    End Select
    
    SaveSetting "VirusTotal", "config", "private_api_key", txtUserPrivateKey
    SaveSetting "VirusTotal", "config", "user_public_api_key", txtUserPublicKey
    
    Select Case i
        Case 0: vt.SetPrivateApiKey Empty
        Case 1: vt.SetPrivateApiKey "user:" & txtUserPublicKey
        Case 2: vt.SetPrivateApiKey txtUserPrivateKey
    End Select
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    Dim i As Long
    
    On Error Resume Next
    txtUserPrivateKey = GetSetting("VirusTotal", "config", "private_api_key", "")
    txtUserPublicKey = GetSetting("VirusTotal", "config", "user_public_api_key", "")
    i = GetSetting("VirusTotal", "config", "key_option", 0)
    If i > optKey.count Then i = 0
    optKey(i).Value = True
    
End Sub
