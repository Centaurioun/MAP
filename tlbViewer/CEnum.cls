VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEnum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public name As String
Public members As New Collection
Public values As New Collection

Public Function nameValuePair(index As Long)
    On Error Resume Next
    nameValuePair = members(index) & " = " & values(index)
End Function

Public Function dump() As String
    On Error Resume Next
    Dim t() As String, v
    push t, "" 'syntax highlight bug
    push t, "Public Enum " & name
    For i = 1 To members.Count
        v = values(i)
        If CLng(v) > 9 Or CLng(v) < 0 Then v = "&H" & Hex(v)
        push t, vbTab & members(i) & " = " & v
    Next
    push t, "End Enum"
    dump = Join(t, vbCrLf)
End Function
