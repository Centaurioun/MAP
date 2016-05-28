VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Socket Tool"
   ClientHeight    =   9570
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   ScaleHeight     =   9570
   ScaleWidth      =   10350
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optTcp 
      Caption         =   "UDP"
      Height          =   240
      Index           =   1
      Left            =   6030
      TabIndex        =   24
      Top             =   90
      Width           =   690
   End
   Begin VB.OptionButton optTcp 
      Caption         =   "TCP"
      Height          =   240
      Index           =   0
      Left            =   5130
      TabIndex        =   23
      Top             =   90
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.CheckBox chkAutoScroll 
      Caption         =   "Auto Scroll Response"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   5040
      TabIndex        =   18
      Top             =   4065
      Width           =   2235
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   255
      Left            =   9060
      TabIndex        =   17
      Top             =   4065
      Width           =   1155
   End
   Begin VB.CheckBox chkLastResponseOnly 
      Caption         =   "Show only last response"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2760
      TabIndex        =   16
      Top             =   4065
      Value           =   1  'Checked
      Width           =   2235
   End
   Begin VB.Frame Frame2 
      Caption         =   "Clipboards"
      ForeColor       =   &H00FF0000&
      Height          =   2595
      Left            =   0
      TabIndex        =   14
      Top             =   6840
      Width           =   10215
      Begin VB.CommandButton cmdSaveNew 
         Caption         =   "Save new"
         Height          =   285
         Left            =   8955
         TabIndex        =   29
         Top             =   2250
         Width           =   1095
      End
      Begin VB.CommandButton cmdLoadFromTop 
         Caption         =   "Load from top"
         Height          =   285
         Left            =   4230
         TabIndex        =   28
         Top             =   2250
         Width           =   1320
      End
      Begin VB.CommandButton cmdCopyTop 
         Caption         =   "Copy to top"
         Height          =   285
         Left            =   5625
         TabIndex        =   27
         Top             =   2250
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   285
         Left            =   6795
         TabIndex        =   26
         Top             =   2250
         Width           =   1005
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   285
         Left            =   7875
         TabIndex        =   25
         Top             =   2250
         Width           =   1005
      End
      Begin MSComctlLib.ListView lv 
         Height          =   2220
         Left            =   135
         TabIndex        =   22
         Top             =   225
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   3916
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "sz"
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "data"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox txtClip 
         Height          =   1965
         Left            =   2520
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   225
         Width           =   7515
      End
   End
   Begin VB.CheckBox chkHexResponse 
      Caption         =   "HexDump"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   960
      TabIndex        =   13
      Top             =   4065
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Send Data "
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   45
      TabIndex        =   10
      Top             =   450
      Width           =   10170
      Begin VB.OptionButton optEscape 
         Caption         =   "Raw as is"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   21
         Top             =   270
         Value           =   -1  'True
         Width           =   1140
      End
      Begin VB.OptionButton optEscape 
         Caption         =   "Convert from Hex String"
         Height          =   240
         Index           =   1
         Left            =   1530
         TabIndex        =   20
         Top             =   270
         Width           =   2085
      End
      Begin VB.CheckBox chkStripCRLF 
         Caption         =   "strip CRLF on send"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   8055
         TabIndex        =   12
         Top             =   225
         Width           =   1695
      End
      Begin VB.TextBox txtEscapeChar 
         Height          =   285
         Left            =   5310
         TabIndex        =   11
         Text            =   "%"
         Top             =   225
         Width           =   435
      End
      Begin VB.OptionButton optEscape 
         Caption         =   "unescape char"
         Height          =   240
         Index           =   0
         Left            =   3870
         TabIndex        =   19
         Top             =   270
         Width           =   2130
      End
   End
   Begin VB.TextBox txtLog 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2385
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   4410
      Width           =   10215
   End
   Begin MSWinsockLib.Winsock ws 
      Left            =   3060
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   315
      Left            =   9270
      TabIndex        =   5
      Top             =   45
      Width           =   900
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   315
      Left            =   8100
      TabIndex        =   4
      Top             =   45
      Width           =   945
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   315
      Left            =   6930
      TabIndex        =   3
      Top             =   45
      Width           =   945
   End
   Begin VB.TextBox txtData 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2835
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1185
      Width           =   10275
   End
   Begin VB.TextBox txtPort 
      Height          =   255
      Left            =   2115
      TabIndex        =   1
      Text            =   "3127"
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtIp 
      Height          =   255
      Left            =   375
      TabIndex        =   0
      Text            =   "127.0.0.1"
      Top             =   120
      Width           =   1275
   End
   Begin VB.Label lblReceived 
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   7425
      TabIndex        =   30
      Top             =   4095
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "Response"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   9
      Top             =   4065
      Width           =   795
   End
   Begin VB.Label lblIp 
      Caption         =   "Port"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   1755
      TabIndex        =   7
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblIp 
      Caption         =   "Ip"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   45
      TabIndex        =   6
      Top             =   120
      Width           =   255
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuLoadFile 
         Caption         =   "LoadFile"
      End
      Begin VB.Menu mnuReplace 
         Caption         =   "Replace"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: david@idefense.com
'
'Purpose: small tool to send text or binary data to backdoor server ports
'         to help in testing/probing their functionality.
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


'aghh dearest mydoom, my first outbreak as a responder..
'...how I miss thee simplicity of yonder year....
'
'Dim login(4) As Byte
'   login(0) = &H85
'   login(1) = &H13
'   login(2) = &H3C
'   login(3) = &H9E
'   login(4) = &HA2


Dim fso As New CFileSystem2
Dim dlg As New clsCmnDlg2

Private Type struc
    ip As String
    port As Long
    datachar As String
    datarep As Long
    onsend As Byte
    escChar As String
    shellcode As String
    stripCRLF As Byte
    hexdump As Byte
    autoscroll As Byte
    freshView As Byte
    data As String
    finalsz As Long
    isUdp As Byte
    sendMode As Byte
    clips(30) As String
End Type

Dim selLi As ListItem
Private settings As struc
Dim received As Long

Private Sub SaveSettings()
   On Error Resume Next
   Dim f As Long
   
   Dim i As Long
   Dim li As ListItem
        
   f = FreeFile
   
   With settings
   
        For Each li In lv.ListItems
             If Len(li.SubItems(1)) > 0 Then
                 .clips(i) = li.SubItems(1)
                 i = i + 1
             End If
        Next
        
        While i < UBound(.clips)
            .clips(i) = Empty
            i = i + 1
        Wend
        
        .ip = txtIp
        .port = txtPort
        .datachar = txtChar
        .datarep = txtNumReps
        .onsend = chkInsertData
        .escChar = txtEscapeChar
        .shellcode = txtData
        .stripCRLF = chkStripCRLF
        .data = txtData
        .freshView = chkLastResponseOnly.value
        .hexdump = chkHexResponse.value
        .autoscroll = chkAutoScroll.value
        .finalsz = CLng(txtFinalSize)
        If optEscape(0).value Then .sendMode = 0
        If optEscape(1).value Then .sendMode = 1
        If optEscape(2).value Then .sendMode = 2
        If optTcp(0).value Then .isUdp = 0 Else .isUdp = 1
    End With
      
   Open App.path & "\options.dat" For Binary As f
   Put f, , settings
   Close f
   
End Sub

Private Sub LoadSettings()
   
   If Not FileExists(App.path & "\options.dat") Then Exit Sub
      
   On Error Resume Next
   Dim f As Long
   f = FreeFile
   
   Open App.path & "\options.dat" For Binary As f
   Get f, , settings
   Close f
   
   Dim i As Long, li As ListItem
   
   With settings
        txtIp = .ip
        txtPort = .port
        txtChar = .datachar
        txtNumReps = .datarep
        chkInsertData = .onsend
        txtEscapeChar = .escChar
        txtData = .shellcode
        chkStripCRLF = .stripCRLF
        txtData = .data
        chkLastResponseOnly.value = .freshView
        chkHexResponse.value = .hexdump
        chkAutoScroll.value = .autoscroll
        txtFinalSize = .finalsz
        optEscape(.sendMode).value = True
        optTcp(.isUdp).value = True
        
        For i = 0 To UBound(.clips)
            If Len(.clips(i)) > 0 Then
                Set li = lv.ListItems.Add(, , Hex(Len(.clips(i))))
                li.SubItems(1) = .clips(i)
            End If
        Next
        
    End With
   
   
End Sub

Private Sub cmdClear_Click()
    txtLog.Text = ""
End Sub

Private Sub cmdClose_Click()
   On Error Resume Next
   ws.Close
   Me.Caption = IIf(Err.Number = 0, "Closed...", "Error: " & Err.Description)
End Sub

Private Sub cmdConnect_Click()
    On Error Resume Next
    ws.Close
    ws.Protocol = IIf(optTcp(0).value, sckTCPProtocol, sckUDPProtocol)
    ws.Connect txtIp, CLng(txtPort)
    Me.Caption = IIf(Err.Number = 0, "Connected...", "Error: " & Err.Description)
End Sub

 


Private Sub cmdLoadFile_Click()
    On Error Resume Next
    If Not FileExists(txtBinary) Then
        MsgBox "Nofile"
    Else
        ReadFile (txtBinary)
        lblIp(4).Caption = "Ready"
        lblFileSize.Caption = "CurSize: " & UBound(myBinary)
    End If
End Sub

 

Private Sub cmdCopyTop_Click()
    txtData = txtClip
End Sub

Private Sub cmdDelete_Click()
    If Not selLi Is Nothing Then
        lv.ListItems.Remove selLi.Index
        Set selLi = Nothing
    End If
    txtClip.Text = Empty
End Sub

Private Sub cmdLoadFromTop_Click()
    txtClip = txtData
End Sub

Private Sub cmdSave_Click()
    If Len(txtClip) = 0 Then Exit Sub
    If selLi Is Nothing Then
        Set selLi = lv.ListItems.Add(, , Hex(Len(txtClip)))
        selLi.SubItems(1) = txtClip
    Else
        selLi.SubItems(1) = txtClip
        selLi.Text = Hex(Len(txtClip))
    End If
End Sub

Private Sub cmdSaveNew_Click()
    If lv.ListItems.Count = 29 Then
        MsgBox "Sorry you hit the max!"
        Exit Sub
    End If
    If Len(txtClip) = 0 Then Exit Sub
    Set selLi = lv.ListItems.Add(, , Hex(Len(txtClip)))
    selLi.SubItems(1) = txtClip
End Sub

Private Sub cmdSend_Click()
    On Error GoTo hell
    
    Dim x As Long
    Dim b() As Byte
    
    received = 0
    buf = txtData
    
    If chkStripCRLF.value = 1 Then
        buf = Replace(buf, vbCrLf, "")
    End If
    
    'If chkInsertData.value = 1 Then
    '    buf = Replace(buf, "[DATA]", String(txtNumReps, Chr("&h" & txtChar)))
    'End If
    
    If optEscape(0).value = True Then
        buf = Escape(buf)
    ElseIf optEscape(1).value = True Then
        If Not HexStringUnescape(buf, b(), True) Then
            MsgBox "had errors converting raw hex input to bytes, aborting"
            Exit Sub
        End If
        buf = StrConv(b(), vbUnicode)
    End If
        
    ws.SendData buf
    Me.Caption = "Connected: " & Len(buf) & " Bytes Sent (" & Hex(Len(buf)) & ")"
    
    Exit Sub
hell:
    MsgBox Err.Description
    Me.Caption = "Error"
End Sub











Function Escape(it)
    Dim f(): Dim c()
    n = Replace(it, "+", " ")
    If InStr(n, txtEscapeChar) > 0 Then
        t = Split(n, txtEscapeChar)
        For i = 0 To UBound(t)
            a = Left(t(i), 2)
            b = IsHex(a)
            If b <> Empty Then
                push f(), txtEscapeChar & a
                push c(), b
            End If
        Next
        For i = 0 To UBound(f)
            n = Replace(n, f(i), c(i))
        Next
    End If
    Escape = n
End Function

Private Function IsHex(it)
    On Error GoTo out
      IsHex = Chr(Int("&H" & it))
    Exit Function
out:  IsHex = Empty
End Function

Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub
 
 

Private Sub cmdWHoleHit_Click()
   On Error GoTo hell
   
   x = UBound(myBinary) 'err if not loaded
   ws.SendData myBinary()
 
 Exit Sub
hell:  MsgBox Err.Description
End Sub


 
Private Sub Command4_Click()
    Dim s As String
    s = txtClip
    txtClip = txtData
    txtData = s
End Sub

Private Sub Command5_Click()
 Dim s As String
    s = txtClip2
    txtClip2 = txtData
    txtData = s
End Sub

Private Sub Form_Load()
    LoadSettings
    'txtBinary = App.path & "\test.exe"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSettings
End Sub

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set selLi = Item
    txtClip = selLi.SubItems(1)
End Sub

Private Sub mnuLoadFile_Click()
    On Error Resume Next
    Dim p As String
    p = dlg.OpenDialog(AllFiles)
    If Len(p) = 0 Then Exit Sub
    p = fso.ReadFile(p)
    txtData = sdump(p)
End Sub

Private Sub txtBinary_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error Resume Next
    txtBinary = data.Files(1)
End Sub


Private Sub mnuReplace_Click()
    frmReplace.LaunchReplaceForm txtData
End Sub

Private Sub ws_Close()
    Me.Caption = "Socket Closed..."
End Sub

Private Sub ws_DataArrival(ByVal bytesTotal As Long)
    Dim s As String
    On Error Resume Next
    
    received = received + 1
    lblReceived = "Received: " & received
    
    ws.GetData s
    
    If chkLastResponseOnly.value = 1 Then txtLog = ""
    
    If chkHexResponse.value = 1 Then
        txtLog = txtLog & hexdump(s) & vbCrLf
    Else
        s = Replace(s, Chr(0), "\x00")
        txtLog = txtLog & s
    End If
    
    If chkAutoScroll.value Then txtLog.SelStart = Len(txtLog)
    
End Sub



Function ReadFile(filename)
  f = FreeFile
   ReDim myBinary(FileLen(filename))
   Open filename For Binary As #f        ' Open file.(can be text or image)
     Get f, , myBinary() ' Get entire Files data
   Close #f
   
End Function

Function FileExists(path) As Boolean
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True _
  Else FileExists = False
End Function

Private Sub ws_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Me.Caption = "Error: " & Description
End Sub




Function hexdump(it)
    Dim my, i, c, s, a, b
    Dim lines() As String
    
    my = ""
    For i = 1 To Len(it)
        a = Asc(Mid(it, i, 1))
        c = Hex(a)
        c = IIf(Len(c) = 1, "0" & c, c)
        b = b & IIf(a > 33 And a < 123, Chr(a), ".")
        my = my & c & " "
        If i Mod 16 = 0 Then
            push lines(), my & "  [" & b & "]"
            my = Empty
            b = Empty
        End If
    Next
    
    If Len(b) > 0 Then
        If Len(my) < 48 Then
            my = my & String(48 - Len(my), " ")
        End If
        If Len(b) < 16 Then
             b = b & String(16 - Len(b), " ")
        End If
        push lines(), my & "  [" & b & "]"
    End If
        
    If Len(it) < 16 Then
        hexdump = my & "  [" & b & "]"
    Else
        hexdump = Join(lines, vbCrLf)
    End If
    
    
End Function


Public Function HexStringUnescape(str, ByRef r() As Byte, Optional stripWhite As Boolean = False) As Boolean

    Dim ret As String
    Dim x As String
    Dim errCount As Long
    Dim b As Byte
    
    On Error Resume Next
    
    Erase r()
    
    If stripWhite Then
        str = Replace(str, " ", Empty)
        str = Replace(str, vbCrLf, Empty)
        str = Replace(str, vbCr, Empty)
        str = Replace(str, vbLf, Empty)
        str = Replace(str, vbTab, Empty)
        str = Replace(str, Chr(0), Empty)
    End If

    For i = 1 To Len(str) Step 2 'this is to agressive for headers...
        x = Mid(str, i, 2)
        If isHexChar(x, b) Then
            bpush r(), b
        Else
            errCount = errCount + 1
            s_bpush r(), x
        End If
    Next

    HexStringUnescape = (errCount = 0)
     
End Function


Public Function isHexChar(hexValue As String, Optional b As Byte) As Boolean
    On Error Resume Next
    Dim v As Long
    
    
    If Len(hexValue) = 0 Then GoTo nope
    If Len(hexValue) > 2 Then GoTo nope 'expecting hex char code like FF or 90
    
    v = CLng("&h" & hexValue)
    If Err.Number <> 0 Then GoTo nope 'invalid hex code
    
    b = CByte(v)
    If Err.Number <> 0 Then GoTo nope  'shouldnt happen.. > 255 cant be with len() <=2 ?

    isHexChar = True
    
    Exit Function
nope:
    Err.Clear
    isHexChar = False
End Function


Private Sub bpush(bAry() As Byte, b As Byte) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Long
    
    x = UBound(bAry) '<-throws Error If Not initalized
    ReDim Preserve bAry(UBound(bAry) + 1)
    bAry(UBound(bAry)) = b
    
    Exit Sub

init:
    ReDim bAry(0)
    bAry(0) = b
    
End Sub

Private Sub s_bpush(bAry() As Byte, sValue As String)
    Dim tmp() As Byte
    Dim i As Long
    tmp() = StrConv(sValue, vbFromUnicode, LANG_US)
    For i = 0 To UBound(tmp)
        bpush bAry, tmp(i)
    Next
End Sub

Function sdump(it)
    Dim b() As Byte
    Dim t() As String
    b() = StrConv(it, vbFromUnicode)
    For i = 0 To UBound(b)
        push t, sHex(b(i))
    Next
    sdump = Join(t, "")
End Function

Function sHex(x) As String
    sHex = Right("0" & Hex(x), 2)
End Function

