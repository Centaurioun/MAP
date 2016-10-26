VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSearchResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public criteria
Public response_code
Public verbose_msg

Public RawJson As String
Public HTTP_ResponseCode As Long
Public HTTP_StatusMessage As String
Public HTTP_HEADERS As String

Public results As New Collection  'of hash string
Public HadError As Boolean
'Public Errors As New Collection

'
'{"response_code": 0, "verbose_msg": "No samples matching the search criteria"}
'{"response_code": -1, "verbose_msg": "Unbalanced modifier or operator"}
'
'{"response_code": 1, "hashes": ["b531a5035025f889ec51565284de94548ac17256442c07303ab3f735894adc4b",
'"cfe45b8302d116595b28569bca52329733d9944b19b52ec8f1db253e33106ef1",
'"eddb2bdc8a6d4bd23af473062593a1c4eb3d8627d8a033afab30b6099e307b24",
'"7804ad9b96fedd038e93e198a43a813da5d7de1024e69f6fc707ff929409d651",
'"739fbb0e06e8d6a425484bd612ca84b10484c097adf931afc78301d172292107"],
'"verbose_msg": "Found samples matching the required properties"}
'
'
'{"response_code": 1, "offset": "RmFsc2U6Q3VBRkN1b0R...XBJQ2hKVXlBaWFXNWtaWGhmY=",
'"hashes": ["b78f93c8e3ef37c2b0e0f45d44186dde4962e36bf5386183e834b05a3150deb2",
'"298fdf9f25b824a1087a8ccb20070f2d19cb5ada3794cd5c7b8b6a49f2dd5213",
'"eea630eb0ef9956bbd4c5e943e14d4970a8ee2de8591501f9bc738792516fcb1",
'"696ba47fd7c282a6a24719e559467d7effe35959daa1f4fedb0e33583589de39"],
'"verbose_msg": "Found samples matching the required properties"
'}

Public lastOffset As String

Friend Function CopyAndAppend(c As CSearchResult)
    Dim r
    
    With c
        Me.criteria = .criteria
        Me.response_code = .response_code
        Me.verbose_msg = .verbose_msg
        Me.RawJson = .RawJson
        Me.HTTP_ResponseCode = .HTTP_ResponseCode
        Me.HTTP_StatusMessage = .HTTP_StatusMessage
        Me.HadError = .HadError
           
        For Each r In .results
            results.Add r
        Next
    End With
    
End Function

Friend Function TruncateResults(limit As Long)
        
    On Error Resume Next
    Dim i As Long
    
    For i = results.count To 1 Step -1
        If results.count = limit Then Exit For
        results.Remove i
    Next
        
End Function

Friend Function ParseSearchResults(json)
    
    Dim a As Long, b As Long
    Dim hashs As String, tmp() As String, t
    
    RawJson = json
    
    Const rc_marker = "response_code"":"
    Const msg_marker = "verbose_msg"":"
    Const offset_marker = "offset"":"
    
    a = InStr(json, rc_marker)
    If a > 0 Then
        a = a + Len(rc_marker)
        b = InStr(a, json, ",")
        If b > 0 Then
            response_code = Trim(Mid(json, a, b - a))
        End If
    End If
    
    a = InStr(json, msg_marker)
    If a > 0 Then
        a = a + Len(msg_marker)
        b = InStr(a, json, ",")
        If b > 0 Then
            verbose_msg = Trim(Mid(json, a, b - a))
        End If
    End If
    
    If response_code <> 1 Then
        HadError = True
        Exit Function
    End If
    
    a = InStr(json, offset_marker)
    If a > 0 Then
        a = a + Len(offset_marker)
        b = InStr(a, json, ",")
        If b > 0 Then
            lastOffset = Trim(Mid(json, a, b - a))
        End If
    End If
    
    a = InStr(json, "[")
    If a > 0 Then
        a = a + 1
        b = InStr(a, json, "]")
        If b > 0 Then
            hashs = Trim(Mid(json, a, b - a))
        End If
    End If
    
    If Len(hashs) > 0 Then
        hashs = Replace(hashs, """", Empty)
        tmp = Split(hashs, ",")
        For Each t In tmp
            results.Add Trim(t)
        Next
    End If
    
End Function

