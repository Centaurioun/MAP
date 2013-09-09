VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmOffsets 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   6780
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtOffset 
      BackColor       =   &H80000004&
      Height          =   315
      Left            =   1380
      TabIndex        =   8
      Top             =   2400
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "FileOffset"
      Height          =   315
      Index           =   2
      Left            =   60
      TabIndex        =   7
      Top             =   2340
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "RVA"
      Height          =   315
      Index           =   1
      Left            =   60
      TabIndex        =   6
      Top             =   1980
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "VirtAddress"
      Height          =   315
      Index           =   0
      Left            =   60
      TabIndex        =   5
      Top             =   1620
      Width           =   1215
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
      Height          =   375
      Left            =   5580
      TabIndex        =   4
      Top             =   1620
      Width           =   1155
   End
   Begin VB.TextBox txtRVA 
      Height          =   315
      Left            =   1380
      TabIndex        =   3
      Top             =   1980
      Width           =   1095
   End
   Begin VB.TextBox txtVA 
      BackColor       =   &H80000004&
      Height          =   315
      Left            =   1380
      TabIndex        =   2
      Top             =   1620
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvSect 
      Height          =   1515
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   2672
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Virtual Addr"
         Object.Width           =   1588
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Virtual Size"
         Object.Width           =   1588
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "RawOffset"
         Object.Width           =   1588
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "RawSize"
         Object.Width           =   1588
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Attributes"
         Object.Width           =   1235
      EndProperty
   End
   Begin VB.Label lblEntryPoint 
      Caption         =   "Entry Point: "
      Height          =   315
      Left            =   2940
      TabIndex        =   10
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label lblSection 
      Height          =   255
      Left            =   3780
      TabIndex        =   9
      Top             =   2460
      Width           =   2715
   End
   Begin VB.Label Label1 
      Caption         =   "Section:             Bytes :"
      Height          =   255
      Left            =   2940
      TabIndex        =   1
      Top             =   2460
      Width           =   735
   End
End
Attribute VB_Name = "frmOffsets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'dzzie@yahoo.com
'http://sandsprite.com

Private selIndex As Long
Private ImageBase As Long
Private mParent As CPEEditor

Sub Initilize(parent As CPEEditor, Optional modal = True)
    
    selIndex = 1
    FilloutListView lvSect, parent.Sections
    
    Set mParent = parent
    
    ImageBase = mParent.ImageBase
    Me.Caption = "ImageBase: " & Hex(ImageBase)
    
    lblEntryPoint = "Entry Point: " & Hex(parent.EntryPoint)
    txtRVA.Text = Hex(parent.EntryPoint)
    cmdCalculate_Click
    
    If modal Then modal = 1 Else modal = 0
    
    Me.Show modal
    
End Sub


Private Sub cmdCalculate_Click()
    Dim va As Long
    Dim fo As Long
    Dim rva As Long
    Dim sectName As String
    
    On Error Resume Next
    
    Select Case selIndex
        Case 0:  'virtual address
                If Not GetHextxt(txtVA, va) Then Exit Sub
                
                If va < ImageBase Then
                    MsgBox "VA is below Imagebase"
                    Exit Sub
                End If
                
                rva = va - ImageBase
                fo = mParent.RvaToOffset(rva, , sectName)
                
                txtRVA = Hex(rva)
                txtOffset = Hex(fo)
        Case 1: 'rva
                If Not GetHextxt(txtRVA, rva) Then Exit Sub
                
                va = rva + ImageBase
                fo = mParent.RvaToOffset(rva, , sectName)
                
                txtVA = Hex(va)
                txtOffset = Hex(fo)
        Case 2: 'file offset
                If Not GetHextxt(txtOffset, fo) Then Exit Sub
                
                rva = mParent.OffsetToRVA(fo, sectName)
                va = rva + ImageBase
              
                txtRVA = Hex(rva)
                txtVA = Hex(va)
    End Select
        
    lblSection.Caption = sectName
    
End Sub


 

Private Sub Option1_Click(index As Integer)

    Enable txtVA, False
    Enable txtRVA, False
    Enable txtOffset, False
    
    Select Case index
        Case 0: Enable txtVA
        Case 1: Enable txtRVA
        Case 2: Enable txtOffset
    End Select
        
    selIndex = index
End Sub

Sub FilloutListView(lv As Object, Sections As Collection)
        
    If Sections.Count = 0 Then
        MsgBox "Sections not loaded yet"
        Exit Sub
    End If
    
    Dim cs As CSection, li As Object 'ListItem
    lv.ListItems.Clear
    
    For Each cs In Sections
        Set li = lv.ListItems.Add(, , cs.nameSec)
        li.SubItems(1) = Hex(cs.VirtualAddress)
        li.SubItems(2) = Hex(cs.VirtualSize)
        li.SubItems(3) = Hex(cs.PointerToRawData)
        li.SubItems(4) = Hex(cs.SizeOfRawData)
        li.SubItems(5) = Hex(cs.Characteristics)
    Next
    
    Dim i As Integer
    For i = 1 To lv.ColumnHeaders.Count
        lv.ColumnHeaders(i).Width = 1000
    Next
    With lv.ColumnHeaders(i - 1)
        .Width = lv.Width - .Left - 100
    End With
    
    
End Sub

Function GetHextxt(t As TextBox, v As Long) As Boolean
    
    On Error Resume Next
    v = CLng("&h" & t)
    If Err.Number > 0 Then
        MsgBox "Error " & t.Text & " is not valid hex number", vbInformation
        Exit Function
    End If
    
    GetHextxt = True
    
End Function

Sub Enable(t As TextBox, Optional enabled = True)
    t.BackColor = IIf(enabled, vbWhite, &H80000004)
    't.enabled = enabled
    t.Text = Empty
End Sub
