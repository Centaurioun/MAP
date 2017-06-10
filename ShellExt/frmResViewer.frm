VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{9A143468-B450-48DD-930D-925078198E4D}#1.1#0"; "hexed.ocx"
Begin VB.Form frmResViewer 
   Caption         =   "Resource Viewer"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14910
   LinkTopic       =   "Form1"
   ScaleHeight     =   7155
   ScaleWidth      =   14910
   StartUpPosition =   3  'Windows Default
   Begin rhexed.HexEd he 
      Height          =   6855
      Left            =   3600
      TabIndex        =   1
      Top             =   120
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   12091
   End
   Begin MSComctlLib.ListView lv 
      Height          =   6855
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   12091
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
         Text            =   "Size"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Path"
         Object.Width           =   4410
      EndProperty
   End
End
Attribute VB_Name = "frmResViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pe As CPEEditor

Sub ShowResources(pee As CPEEditor)
    Set pe = pee
    
    Dim r As CResourceEntry
    Dim li As ListItem
    
    For Each r In pe.Resources.Entries
        Set li = lv.ListItems.Add(, , Hex(r.size))
        Set li.Tag = r
        li.SubItems(1) = r.path
        Debug.Print r.report
    Next
    
    Me.Show
    
End Sub
 
Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim r As CResourceEntry
    Dim b() As Byte
    
    Set r = Item.Tag
    If pe.Resources.GetResourceData(Item.text, b) Then
        he.LoadByteArray b
    Else
        he.LoadString ""
    End If
    
End Sub
