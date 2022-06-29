VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmtlbViewer 
   Caption         =   "Type Library Viewer"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   9900
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   315
      Left            =   8520
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
   Begin RichTextLib.RichTextBox text2 
      Height          =   4830
      Left            =   3330
      TabIndex        =   4
      Top             =   495
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   8520
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmTlbViewer.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   330
      Left            =   8910
      TabIndex        =   3
      Top             =   90
      Width           =   960
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2040
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTlbViewer.frx":007D
            Key             =   "const"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTlbViewer.frx":018F
            Key             =   "event"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTlbViewer.frx":02A1
            Key             =   "class"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTlbViewer.frx":03B3
            Key             =   "interface"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTlbViewer.frx":04C5
            Key             =   "lib"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTlbViewer.frx":05D7
            Key             =   "sub"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTlbViewer.frx":06E9
            Key             =   "module"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTlbViewer.frx":07FB
            Key             =   "value"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTlbViewer.frx":090D
            Key             =   "prop"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTlbViewer.frx":0A1F
            Key             =   "control"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tv 
      Height          =   4815
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   3240
      _ExtentX        =   5715
      _ExtentY        =   8493
      _Version        =   393217
      Indentation     =   176
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   120
      Width           =   7425
   End
   Begin VB.Label Label2 
      Caption         =   "COM Server"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Begin VB.Menu mnuExpandAll 
         Caption         =   "Expand All"
      End
      Begin VB.Menu mnuCollapseAll 
         Caption         =   "Collapse All"
      End
      Begin VB.Menu mnuStringScanner 
         Caption         =   "Scan for Strings"
      End
      Begin VB.Menu mnuFullProtos 
         Caption         =   "Compact Prototypes"
      End
      Begin VB.Menu mnuCopyFuncNames 
         Caption         =   "Copy Names"
      End
      Begin VB.Menu mnuShowVOff 
         Caption         =   "Show VTable Offsets"
      End
   End
End
Attribute VB_Name = "frmtlbViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'License:   GPL
'Copyright: 2005 iDefense a Verisign Company
'Site:      http://labs.idefense.com
'
'Author:  David Zimmer <david@idefense.com, dzzie@yahoo.com>
'
'         disassembler functionality provided by olly.dll which
'         is a modified version of the OllyDbg GPL source from
'         Oleh Yuschuk Copyright (C) 2001 - http://ollydbg.de
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

Public tlb As New CTlbParse
Public ActiveNode As Node
Public FilterGUID As String

Public lngs As Collection
Public ints As Collection
Public dbls As Collection
Public strs As Collection
Public cust As Collection
 
Private LiveLoadWarned As Boolean
Private MoreMode As Boolean

Private Sub cmdBrowse_Click()
    On Error Resume Next
    Dim p As String
    p = fso.GetParentFolder(Text1)
    p = dlg.OpenDialog(AllFiles, p, "Open File", Me.hwnd)
    Text1 = p
    cmdLoad_Click
End Sub

Private Sub cmdLoad_Click()
    LoadFile Text1
End Sub

Function LoadFile(fpath As String, Optional onlyShowGuid As String) As Boolean
    
    Me.Visible = True
    Text1 = ExpandPath(fpath)
    
    Dim c As CClass
    Dim i As CInterface
    Dim m As CMember
    Dim pi As ParameterInfo
    
    Dim n0 As Node
    Dim n1 As Node
    Dim n2 As Node
    Dim n3 As Node
    
    Dim mMembers As Long
    Dim mInterfaces As Long
    Dim x As Long
    Dim loaded As Boolean
    
    Set tlb = Nothing
    Set tlb = New CTlbParse
    
    FilterGUID = onlyShowGuid
    
    If Len(onlyShowGuid) > 0 Then
        Me.caption = Me.caption & "  Only showing class " & onlyShowGuid
    End If
    
    tv.Nodes.Clear
     
    text2 = Empty
    
    loaded = tlb.LoadFile(Text1, onlyShowGuid)
                
    If loaded Then
        
        Set n0 = tv.Nodes.Add(, , , tlb.LibName, "lib")
        
        For Each c In tlb.mClasses
        
            If Len(onlyShowGuid) > 0 Then
                If InStr(1, c.GUID, onlyShowGuid, vbTextCompare) < 1 Then
                    GoTo nextOne
                End If
            End If
            
            Set n1 = tv.Nodes.Add(n0, tvwChild, , c.name, IIf(c.isControl, "control", "class"))
            Set n1.Tag = c
            mInterfaces = 0
            mMembers = 0
            For Each i In c.mInterfaces
                mInterfaces = mInterfaces + 1
                Set n2 = tv.Nodes.Add(n1, tvwChild, , i.name, "interface")
                Set n2.Tag = i
                For Each m In i.mMembers
                    Set n3 = tv.Nodes.Add(n2, tvwChild, , m.mMemberInfo.name, IIf(m.CallType > 1, "prop", "sub"))
                    Set n3.Tag = m
                    mMembers = mMembers + 1
                    'If ObjPtr(n3) And Not m.SupportsFuzzing Then n3.ForeColor = &H606060
                    Set n3 = Nothing
                Next
                n2.Sorted = True
            Next
            If mInterfaces = 0 Or mMembers = 0 Then n1.Tag = Empty
            n1.Sorted = True
nextOne:
        Next
        
        For x = tv.Nodes.Count To 1 Step -1
            If tv.Nodes(x).index <> n0.index Then
                If Not IsObject(tv.Nodes(x).Tag) Then
                    tv.Nodes.Remove x
                Else
                    If TypeName(tv.Nodes(x).Tag) = "CInterface" Then
                        If tv.Nodes(x).Children = 0 Then tv.Nodes.Remove x
                    End If
                End If
            End If
        Next
        
        n0.Expanded = True
        
        If tlb.types.Count > 0 Then
            Dim t As CRecord
            Set n1 = tv.Nodes.Add(n0, tvwChild, , "Structs", "prop")
            For Each t In tlb.types
                Set n2 = tv.Nodes.Add(n1, tvwChild, , t.name, "prop")
                Set n2.Tag = t
            Next
        End If
        
        If tlb.enums.Count > 0 Then
            Dim e As CEnum
            Set n1 = tv.Nodes.Add(n0, tvwChild, , "Enums", "const")
            For Each e In tlb.enums
                Set n2 = tv.Nodes.Add(n1, tvwChild, , e.name, "const")
                Set n2.Tag = e
            Next
        End If
        
        
    End If
    
    If tv.Nodes.Count = 0 Then
        text2 = tlb.ErrMsg
        LoadFile = False
    Else
        LoadFile = True
        If tv.Nodes.Count > 3 Then
            mnuCollapseAll_Click
        Else
            mnuExpandAll_Click
        End If
        Me.Visible = True
        Me.ZOrder 0
    End If

End Function

Private Sub Form_Load()
    mnuPopup.Visible = False
    LoadKillBittedControlList
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    tv.Height = Me.Height - tv.top - 450
    text2.Height = tv.Height
    text2.Width = Me.Width - text2.Left - 200
End Sub

Private Sub Form_Unload(Cancel As Integer)
    LiveLoadWarned = False
End Sub

Private Sub mnuCollapseAll_Click()
    Dim n As Node
    For Each n In tv.Nodes
        If n.Children > 0 Then n.Expanded = False
    Next
    tv.Nodes(1).Expanded = True
    tv.Nodes(1).EnsureVisible
End Sub

Private Sub mnuCopyFuncNames_Click()
    
    If ActiveNode Is Nothing Then Exit Sub
    If TypeName(ActiveNode.Tag) <> "CInterface" Then Exit Sub
    
    Dim c As CMember
    Dim i As CInterface
    Dim tmp() As String
    
    Set i = ActiveNode.Tag

    For Each c In i.mMembers
        push tmp, c.name
    Next
    
    Clipboard.Clear
    Clipboard.SetText Join(tmp, vbCrLf)
    
    MsgBox i.mMembers.Count & " names copied.", vbInformation

End Sub

Private Sub mnuExpandAll_Click()
    Dim n As Node
    For Each n In tv.Nodes
        If n.Children > 0 Then n.Expanded = True
    Next
    tv.Nodes(1).EnsureVisible
End Sub

Private Sub mnuShowAllClasses_Click()
    If Len(FilterGUID) > 0 Then
        FilterGUID = Empty
        LoadFile Text1
    End If
End Sub


Private Sub mnuFullProtos_Click()

    mnuFullProtos.Checked = Not mnuFullProtos.Checked

'    If Not mnuFullProtos.Checked Then
'        mnuFullProtos.Checked = True
'    Else
'        If mnuFullProtos.Checked = True And InStr(mnuFullProtos.caption, "Compact") < 1 Then
'            mnuFullProtos.caption = "Compact Protos"
'        ElseIf mnuFullProtos.Checked = True And InStr(mnuFullProtos.caption, "Compact") >= 1 Then
'            mnuFullProtos.Checked = False
'            mnuFullProtos.caption = "Full Prototypes"
'        Else
'            mnuFullProtos.Checked = False
'        End If
'
'    End If
        
End Sub

Private Sub mnuShowVOff_Click()
    mnuShowVOff.Checked = Not mnuShowVOff.Checked
End Sub

Private Sub Text1_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    Text1 = data.files(1)
End Sub

 
Private Sub text2_Change()
    On Error Resume Next
    modSyntaxHighlighting.SyntaxHighlight text2
End Sub

Private Sub tv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    mnuCopyFuncNames.Enabled = False
    If Not ActiveNode Is Nothing Then
        'If ActiveNode.Children > 0 Then mnuCopyFuncNames.Enabled = True
        If TypeName(ActiveNode.Tag) = "CInterface" Then mnuCopyFuncNames.Enabled = True
    End If
    If Button = 2 Then PopupMenu mnuPopup
End Sub

Private Sub tv_NodeClick(ByVal Node As MSComctlLib.Node)
    
    Dim c As CMember
    Dim i As CInterface
    Dim cc As CClass
    Dim tmp() As String
    Dim report As String
    On Error Resume Next
    
    Set ActiveNode = Node
       
    If Node.index = 1 Then
        push tmp(), "Loaded File: " & Text1
        push tmp(), "Name:        " & tlb.LibName
        If Len(tlb.tli.GUID) > 0 Then
            push tmp(), "Lib GUID:    " & tlb.tli.GUID
            push tmp(), "Version:     " & tlb.tli.MajorVersion & "." & tlb.tli.MinorVersion
        End If
        push tmp(), "Lib Classes: " & tlb.NumClassesInLib
        text2 = vbCrLf & Join(tmp, vbCrLf)
    End If
    
    If TypeName(Node.Tag) = "CMember" Then
        Set c = Node.Tag
        report = IIf(mnuShowVOff.Checked, c.vTableOffset & "|" & Hex(c.vTableOffset) & vbCrLf, Empty) & c.ProtoString
        text2 = vbCrLf & report
        
    End If
    
    If TypeName(Node.Tag) = "CInterface" Then
        Set i = Node.Tag
        push tmp, "Interface " & i.name & i.DerivedString
        push tmp, "Default Interface: " & i.isDefault
        'push tmp, "Public: " & i.isPublic()
        'push tmp, "Dual: " & i.isDual()
        'push tmp, "Creatable: " & i.isCreatable()
        'push tmp, "Licensed: " & i.isLicensed()
        push tmp, "Members : " & i.mMembers.Count
        
        push tmp, Empty
        'If mnuFullProtos.Checked Then push tmp, Empty
        
        For Each c In i.mMembers
            'If mnuFullProtos.Checked Then
                report = c.ProtoString
                If mnuFullProtos.Checked Then
                'If InStr(mnuFullProtos.caption, "Compact") >= 1 Then
                    report = Replace(Replace(report, vbTab, Empty), vbCrLf, Empty)
                Else
                    report = report & vbCrLf
                End If
                If mnuShowVOff.Checked Then report = "[" & c.vTableOffset & "|" & Hex(c.vTableOffset) & "] " & report
                push tmp, report
            'Else
            '    push tmp, vbTab & c.mMemberInfo.Name
            'End If
        Next
        text2 = vbCrLf & Join(tmp, vbCrLf)
    End If
    
    If TypeName(Node.Tag) = "CClass" Then
        Set cc = Node.Tag
        push tmp, "Class " & cc.name
        push tmp, "GUID: " & cc.GUID
        push tmp, "Number of Interfaces: " & cc.mInterfaces.Count
        
        If cc.mInterfaces.Count > 0 Then
            For Each i In cc.mInterfaces
                push tmp, vbTab & i.name & " - " & i.GUID & " " & i.DerivedString
            Next
        End If
        
        push tmp, "Default Interface: " & cc.DefaultInterface
        push tmp, "KillBitSet: " & cc.KillBitSet
        push tmp, vbCrLf
        
        If Not cc.isRegisteredOnSystem Then
            push tmp, "Control not registered on system"
        Else
            push tmp, "RegKey Safe for Script: " & cc.SafeForScripting
            push tmp, "RegkeySafe for Init: " & cc.SafeForInitilization
            If cc.isDesignTime Then push tmp, "Design Time Editable"
            If cc.isDotNet Then push tmp, "Created in .NET"
            If cc.isInsertable Then push tmp, "Insertable"
            If cc.isControl Then push tmp, "Control"
        End If
                
        text2 = vbCrLf & Join(tmp, vbCrLf)
    End If
    
    
     If TypeName(Node.Tag) = "CRecord" Then
        Dim cr As CRecord
        Set cr = Node.Tag
        text2 = cr.dump()
     End If
     
     If TypeName(Node.Tag) = "CEnum" Then
        Dim ce As CEnum
        Set ce = Node.Tag
        text2 = ce.dump()
     End If
     
     
End Sub

Private Sub mnuStringScanner_Click()
    Dim i As Long
    Dim tmp() As String
    Dim n As Node
    Dim m As CMember
    Dim a As CArgument
    Dim match As String
    
    On Error Resume Next
    
    match = InputBox("Enter comma delimited substrings to find", , "file,path,url,key")
    If Len(match) = 0 Then Exit Sub
    
    For Each n In tv.Nodes
        If IsObject(n.Tag) Then
            If TypeName(n.Tag) = "CMember" Then
                Set m = n.Tag
                If AnyOfTheseInstr(m.mMemberInfo.name, match) Then
                    push tmp, "Clsid: " & m.ClassGUID & " function: " & m.mMemberInfo.name
                End If
                For Each a In m.Args
                    If AnyOfTheseInstr(a.name, match) Then
                        push tmp, "Clsid: " & m.ClassGUID & " function: " & m.mMemberInfo.name & " Argument: " & a.name
                    End If
                Next
            End If
        End If
    Next
                                     
    If Not AryIsEmpty(tmp) Then
       frmMsg.Display "Search results for match string: " & match & vbCrLf & vbCrLf & Join(tmp, vbCrLf)
    Else
        MsgBox "no string matchs found for function names or arguments :(", vbInformation
    End If
                    
End Sub

Function GetParentClass(member As Node) As CClass

    Dim rep As Long
    Dim mNode As Node
    Dim cc As CClass
    On Error Resume Next
    
    Set mNode = member
top:

    If TypeName(mNode.Tag) = "CClass" Then
        Set cc = mNode.Tag
        Set GetParentClass = cc
    Else
        rep = rep + 1
        If rep < 3 Then
            Set mNode = mNode.Parent
            GoTo top
        End If
    End If



End Function

Sub ScanElementsFor(match As String, tmp() As String, alerted As Collection)
    On Error Resume Next
    Dim key As String
    Dim n As Node
    Dim m As CMember
    Dim a As CArgument
    
    For Each n In tv.Nodes
        If IsObject(n.Tag) Then
            If TypeName(n.Tag) = "CMember" Then
                Set m = n.Tag
                If AnyOfTheseInstr(m.mMemberInfo.name, match) Then
                    key = m.ClassGUID & "." & m.mMemberInfo.name
                    If Not KeyExistsInCollection(alerted, key) Then
                        alerted.Add key, key
                        push tmp, "Library: " & tv.Nodes(1).Text & " - " & Text1
                        push tmp, "Class: " & GetParentClass(n).name & "  " & m.ClassGUID & vbCrLf
                        push tmp, m.ProtoString & vbCrLf
                        push tmp, String(40, "-")
                    End If
                End If
                For Each a In m.Args
                    If AnyOfTheseInstr(a.name, match) Then
                        key = m.ClassGUID & "." & m.mMemberInfo.name
                        If Not KeyExistsInCollection(alerted, key) Then
                            alerted.Add key, key
                            push tmp, "Library: " & tv.Nodes(1).Text & " - " & Text1
                            push tmp, "Class: " & GetParentClass(n).name & "  " & m.ClassGUID & vbCrLf
                            push tmp, m.ProtoString & vbCrLf
                            push tmp, String(40, "-")
                        End If
                    End If
                Next
            End If
        End If
    Next
    
End Sub

