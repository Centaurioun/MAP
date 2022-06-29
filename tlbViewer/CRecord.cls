VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public name As String
Public members As New Collection
Public types As New Collection
Public typedName As New Collection

Public Function dump() As String
    On Error Resume Next
    Dim t() As String
    push t, "" 'syntax highlight bug
    push t, "Public Type " & name
    For i = 1 To members.Count
        push t, vbTab & members(i) & " as " & typedName(i)
    Next
    push t, "End Type"
    dump = Join(t, vbCrLf)
End Function

