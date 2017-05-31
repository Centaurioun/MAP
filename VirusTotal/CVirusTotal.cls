VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVirusTotal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'8.5.16 - includes first seen date in default report if using private api key

'todo switch to something threaded and async for DownloadFile...


Public WithEvents winInet As Inet
Attribute winInet.VB_VarHelpID = -1

'still used for file submit
Private http As New XMLHTTP

Const URL_API_BASIS = "http://www.virustotal.com/vtapi/v2/"
Const URL_SCAN_FILE = "file/scan"
Const URL_FILE_REPORT = "file/report"
Const URL_SCAN_URL = "url/scan"
Const URL_URL_REPORT = "url/report"
Const URL_MAKE_COMMENT = "comments/put"
Const URL_SEARCH = "file/search"

Private API_KEY As String
Const PUBLIC_API_KEY = "a949ea9c64e7145a065b0e562673a66216a132712e958168c8c00ee5f451485b"

Public report_cache_dir As String 'if this is set it will cache the results here in case looked up again.
Public debugLog As Object         'set to a listbox for live logging, or it will default to a collection
Private m_Abort As Boolean
Public KeepLog As Boolean

Public WithEvents Timer1 As Timer 'we cant make this public it will switch to object type in typelib
Attribute Timer1.VB_VarHelpID = -1
Private m_usingPrivateKey As Boolean

Private Declare Function URLDownloadToFile Lib "urlmon" Alias _
    "URLDownloadToFileA" (ByVal pCaller As Long, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As Long, _
    ByVal lpfnCB As Long) As Long

Private NeedsDelay As Boolean
Private LastQueryTime As Long
Private m_delayInterval As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const LB_GETCOUNT = &H18B
Private Const LB_SETTOPINDEX = &H197

Dim repeatOnce As Boolean
Const LANG_US = &H409

Property Get abort() As Boolean
    abort = m_Abort
End Property

Property Let abort(v As Boolean)
    m_Abort = v
    If v Then winInet.Cancel
End Property

Property Get usingPrivateKey() As Boolean
    usingPrivateKey = m_usingPrivateKey
End Property

'MUST BE SET!
'Property Set TimerObj(tmr As Timer)
'    On Error Resume Next
'    Set Timer1 = tmr
'End Property

Public Function RawDownloadFile(URL, LocalFilename) As Boolean
    Dim lngRetVal As Long
    lngRetVal = URLDownloadToFile(0, CStr(URL), CStr(LocalFilename), 0, 0)
    If lngRetVal = 0 Then RawDownloadFile = True
End Function

Property Get delayInterval() As Long
    delayInterval = m_delayInterval
End Property

Property Let delayInterval(x As Long)
    If Not usingPrivateKey Then
        m_delayInterval = x
    End If
End Property

Private Sub debugLogger(msg)
    
    On Error Resume Next
    
    If debugLog Is Nothing Then
        Set debugLog = New Collection
    End If
    
    If TypeName(debugLog) = "ListBox" Then
        debugLog.AddItem msg
        LB_ScrollTo debugLog.hwnd
        debugLog.Refresh
    ElseIf TypeName(debugLog) = "Collection" Then
        debugLog.Add msg
    End If
        
End Sub

Private Sub clearDebugLog()

    On Error Resume Next
    
    If debugLog Is Nothing Then
        Set debugLog = New Collection
    ElseIf TypeName(debugLog) = "ListBox" Then
        If Not KeepLog Then debugLog.Clear
    ElseIf TypeName(debugLog) = "Collection" Then
        If Not KeepLog Then Set debugLog = New Collection
    End If
    
End Sub

Public Function DownloadFile(hash As String, toDir As String) As Boolean
    
    On Error GoTo hell
        
    If Len(API_KEY) = 0 Then Class_Initialize
    If Not usingPrivateKey Then Exit Function
    If Not FolderExists(toDir) Then Exit Function
    
'private api: GET /vtapi/v2/file/download
'  apikey
'  hash    The md5/sha1/sha256 hash of the file you want to download.

    
    Dim URL As String
    Dim b() As Byte
    Dim hFile As Long
    Dim path As String
    Dim fSize As Long
    
    URL = "https://www.virustotal.com/vtapi/v2/file/download?apikey=" & API_KEY & "&hash=" & hash
    path = toDir & "\" & hash
    
    If FileExists(path) Then
        DownloadFile = True
        Exit Function
    End If
    
    With winInet
        .Execute URL, "GET"
     
        While .StillExecuting
           DoEvents
        Wend
        
        b() = .GetChunk(1024, icByteArray)
        'If Err.Number = 0 Then
            If UBound(b) <> -1 Then
                hFile = FreeFile
                Open path For Binary As hFile
                Put hFile, , b()
                
                Do
                    b() = .GetChunk(1024, icByteArray)
                    'If Err.Number <> 0 Then Exit Do
                    If UBound(b) = -1 Then Exit Do
                    Put hFile, , b()
                Loop
                fSize = LOF(hFile) - 1
                Close hFile
            End If
        'End If
            
        If Not FileExists(path) Then Exit Function
        
        If fSize = 0 Then
            delFile path
            Exit Function
        End If
        
        DownloadFile = True
        
    End With
    
    Exit Function
hell: DownloadFile = False
End Function

Private Function delFile(f)
    On Error Resume Next
    Kill f
End Function

'this class supports a couple nice things..
'1) if you set the cache dir, it will save raw json there as optimization in case you lookup the same file
'    file again. this will sve you the 4sec delay per request.
'2) this class tracks when the last request was made, and implements its own optimized delay so that the
'   4 second threshold is not exceeded.
'3) since vb doesnt support json, a CScan object is always passed back with details of the scan. parsing
'    was just done in VB and is VT json specific and maybe a bit shady as it was done quickly..


Function GetReport(hash As String) As CScan

    On Error Resume Next
    
    Dim scan As New CScan
    Dim my_json As String
    Dim sStatus As String
    Dim status As Long
    Dim curtime As Long
    Dim success As Boolean
    Dim errorMsg As String
    Dim headers As String
    
    abort = False

    If Len(API_KEY) = 0 Then Class_Initialize
    
    If Timer1 Is Nothing Then
        MsgBox "Must set Timer1 reference before using GetReport!"
        Exit Function
    End If
    
    
    clearDebugLog
    Timer1.Enabled = False
    Timer1.Interval = m_delayInterval

    Set http = New XMLHTTP
    
    If http Is Nothing Then
        errorMsg = "Could not create XMLHTTP Object"
        GoTo errorExit
    End If
        
    If FolderExists(report_cache_dir) Then
        If FileExists(report_cache_dir & "\" & hash & ".txt") Then
            debugLogger "Cached report found for " & hash
            my_json = ReadFile(report_cache_dir & "\" & hash & ".txt")
            GoTo parse_json
        End If
    End If
            
    DelayIfNeeded
    
    debugLogger "Connecting to VirusTotal to query report for " & hash
    
    DoEvents
    
    NeedsDelay = True 'need 4 second delay for next if using public api key
    
    LastQueryTime = GetTickCount
    success = Get_Web_Report(hash, my_json, sStatus, status, headers)
    scan.HTTP_ResponseCode = status
    scan.HTTP_StatusMessage = sStatus
    scan.HTTP_HEADERS = headers
    
    If Not success Then
        errorMsg = "Web request failure - status code: " & status & " " & sStatus
        GoTo errorExit
    End If
   
    debugLogger "Query successful for hash: " & hash
         
    If FolderExists(report_cache_dir) Then
        writeFile report_cache_dir & "\" & hash & ".txt", my_json
    End If
    
parse_json:
    scan.ParseResponse my_json
    Set GetReport = scan
    
    debugLogger "Detections: " & scan.positives & " Response Code: " & scan.response_code
    If scan.HadError Then debugLogger "Had JSON Parsing error..."
     
    
Exit Function

errorExit:

    debugLogger errorMsg
    scan.Errors.Add errorMsg
    scan.HadError = True
    Set GetReport = scan
    
End Function


'search results are returned in blocks of 300 and contain only sha256 hashs...
Function Search(ByVal criteria, Optional ByVal limit As Long = 300) As CSearchResult
    
    Dim cs As CSearchResult
    Dim cs2 As CSearchResult
        
    
    abort = False
    debugLogger "Starting search for: " & criteria & " Limit: " & limit
    Set cs = Search_internal(criteria)
    Set Search = cs

    If Not cs.HadError Then
        Do While cs.results.count < limit
            If abort Then Exit Do
            DoEvents
            debugLogger "Loading next block of search results current count: " & cs.results.count & " Limit: " & limit
            DoEvents
            If Len(cs.lastOffset) = 0 Then Exit Do             'no more results...
            Set cs2 = Search_internal(criteria, cs.lastOffset)
            If cs2.HadError Then Exit Do
            cs.CopyAndAppend cs2
        Loop
    End If
    
    If cs.results.count > limit Then cs.TruncateResults limit
    
    
End Function

Private Function Search_internal(ByVal criteria As String, Optional ByVal offset As String) As CSearchResult

    On Error Resume Next
    Dim my_json As String
    Dim sStatus As String
    Dim status As Long
    Dim curtime As Long
    Dim success As Boolean
    Dim errorMsg As String
    Dim headers As String
    
    Dim csr As New CSearchResult
    
    csr.criteria = criteria
    Set Search_internal = csr
    abort = False
    If Len(API_KEY) = 0 Then Class_Initialize
    If Len(Trim(criteria)) = 0 Then Exit Function
    
    If Not m_usingPrivateKey Then
        errorMsg = "Search requires a private API key"
        GoTo errorExit
    End If
    
    If Timer1 Is Nothing Then
        MsgBox "Must set Timer1 reference before using GetReport!"
        Exit Function
    End If
    
    clearDebugLog
    Timer1.Enabled = False
    Timer1.Interval = m_delayInterval
    
    debugLogger "Connecting to VirusTotal to search for '" & criteria & "'"
    NeedsDelay = False  'this is private api only
    LastQueryTime = GetTickCount
    
    success = DoWebSearch(criteria, my_json, sStatus, status, headers, offset)
    csr.HTTP_ResponseCode = status
    csr.HTTP_StatusMessage = sStatus
    csr.HTTP_HEADERS = headers
    
    If Not success Then
        errorMsg = "Web request failure - status code: " & status & " " & sStatus
        GoTo errorExit
    End If
   
    debugLogger "Search successful for criteria: " & criteria
         
parse_json:
    csr.ParseSearchResults my_json
    
    debugLogger "Parsed Results: " & csr.results.count
    'If scan.HadError Then debugLogger "Had JSON Parsing error..."
    
Exit Function

errorExit:
    debugLogger errorMsg
    
End Function

Private Function DoWebSearch(criteria, out_response As String, out_status As String, out_statusCode As Long, outHeaders As String, Optional ByVal offset As String) As Boolean
    
    Err.Clear
        
    Dim x As Variant
    Dim s As String
    
    out_status = Empty
    out_response = Empty
    outHeaders = False
    repeatOnce = True
    
tryAgain:

On Error GoTo hell

    Dim strUrl As String, strFormData As String
    Dim var_data As Variant, str_data As String
     
    strUrl = URL_API_BASIS & URL_SEARCH
    s = "key=" & API_KEY & "&query=" & criteria
    If Len(offset) > 0 Then s = s & "&offset=" & offset
    strFormData = s
    
    'winInet.RequestTimeout = 3000
    winInet.Execute strUrl, "POST", strFormData, "Content-Type: application/x-www-form-urlencoded"
    
    While winInet.StillExecuting
        DoEvents
    Wend
    
    'wininet.responsecode doesnt always work?
    outHeaders = winInet.GetHeader
    ParseHeader outHeaders, out_status, out_statusCode
    
    var_data = winInet.GetChunk(1024, icString)
    str_data = str_data & var_data

    Do
        DoEvents
        var_data = winInet.GetChunk(1024, icString)
        If Len(var_data) = 0 Then Exit Do
        str_data = str_data & var_data
    Loop
     
    out_response = str_data
    If out_statusCode = 200 Then DoWebSearch = True
    repeatOnce = True
     
hell:
    DoEvents
    If Err.Number <> 0 Then
        debugLogger "Error in DoWebSearch Line: " & Erl & " desc: " & Err.Description & " Status: " & out_statusCode
    End If
    
End Function

Private Function ParseHeader(ByVal h As String, ByRef out_status As String, ByRef out_statusCode As Long)
    
    On Error Resume Next
    
    Dim a As Long, b As Long, tmp As String
    a = InStr(h, vbLf)
    If a > 1 Then
        h = Mid(h, 1, a - 1)
        If Right(h, 1) = vbCr Then h = Mid(h, 1, Len(h) - 1)
        a = InStr(h, " ")
        If a > 0 Then
            a = a + 1
            b = InStr(a, h, " ")
            If b > 0 Then
                tmp = Trim(Mid(h, a, b - a))
                out_statusCode = CLng(tmp)
                out_status = Trim(Mid(h, b))
            End If
        End If
    End If
                
End Function

Private Function Get_Web_Report(hash, out_response As String, out_status As String, out_statusCode As Long, outHeaders As String) As Boolean
    
    Err.Clear
        
    Dim x As Variant
    out_status = Empty
    out_response = Empty
    outHeaders = Empty
    repeatOnce = True
    
tryAgain:

On Error GoTo hell

    Dim strUrl As String, strFormData As String
    Dim var_data As Variant, str_data As String
     
    strUrl = URL_API_BASIS & URL_FILE_REPORT
    s = "key=" & API_KEY & "&resource=" & hash & IIf(usingPrivateKey, "&allinfo=1", "")
    If Len(offset) > 0 Then s = s & "&offset=" & offset
    strFormData = s
    
    'winInet.RequestTimeout = 3000
    winInet.Execute strUrl, "POST", strFormData, "Content-Type: application/x-www-form-urlencoded"
    
    While winInet.StillExecuting
        DoEvents
    Wend
    
    'wininet.responsecode doesnt always work?
    outHeaders = winInet.GetHeader
    ParseHeader outHeaders, out_status, out_statusCode
    
    var_data = winInet.GetChunk(1024, icString)
    str_data = str_data & var_data

    Do
        DoEvents
        var_data = winInet.GetChunk(1024, icString)
        If Len(var_data) = 0 Then Exit Do
        str_data = str_data & var_data
    Loop
     
    out_response = str_data
    If out_statusCode = 200 Then Get_Web_Report = True
    
hell:
    DoEvents
    If Err.Number <> 0 Then
        debugLogger "Error in Get_Web_Report Line: " & Erl & " desc: " & Err.Description & " Status: " & out_status
    End If
    
End Function

Function AddComment(hash, comment, out_response As String, out_status As String, out_statusCode As Long, out_headers As String) As Boolean
    
    abort = False
    If Len(API_KEY) = 0 Then Class_Initialize

    Err.Clear
    On Error GoTo hell

    Dim x As Variant
    out_status = Empty
    out_response = Empty
    out_statusCode = Empty
    out_headers = Empty
    
    Dim strUrl As String, strFormData As String
    Dim var_data As Variant, str_data As String
     
    strUrl = URL_API_BASIS & URL_MAKE_COMMENT
    s = "key=" & API_KEY & "&resource=" & hash & "&comment=" & comment
    strFormData = s
    
    'winInet.RequestTimeout = 3000
    winInet.Execute strUrl, "POST", strFormData, "Content-Type: application/x-www-form-urlencoded"
    
    While winInet.StillExecuting
        DoEvents
    Wend
    
    'wininet.responsecode doesnt always work?
    out_headers = winInet.GetHeader
    ParseHeader out_headers, out_status, out_statusCode
    
    var_data = winInet.GetChunk(1024, icString)
    str_data = str_data & var_data

    Do
        DoEvents
        var_data = winInet.GetChunk(1024, icString)
        If Len(var_data) = 0 Then Exit Do
        str_data = str_data & var_data
    Loop
     
    out_response = str_data
    If out_statusCode = 200 Then AddComment = True
    
hell:
    DoEvents
    If Err.Number <> 0 Then
        debugLogger "Error in AddComment Line: " & Erl & " desc: " & Err.Description
    End If
    
End Function

Private Function FolderExists(path) As Boolean
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbDirectory) <> "" Then FolderExists = True _
  Else FolderExists = False
End Function

Private Function FileExists(p) As Boolean
    If Len(p) = 0 Then Exit Function
    If Dir(p, vbNormal Or vbHidden Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
End Function

Private Function ReadFile(filename) As String 'this one should be binary safe...
  On Error GoTo hell
  f = FreeFile
  Dim b() As Byte
  Open filename For Binary As #f
  ReDim b(LOF(f) - 1)
  Get f, , b()
  Close #f
  ReadFile = StrConv(b(), vbUnicode, LANG_US)
  Exit Function
hell:   ReadFile = ""
End Function

Private Function writeFile(path, it) As Boolean 'this one should be binary safe...
    On Error GoTo hell
    Dim b() As Byte
    If FileExists(path) Then Kill path
    f = FreeFile
    b() = StrConv(it, vbFromUnicode, LANG_US)
    Open path For Binary As #f
    Put f, , b()
    Close f
    writeFile = True
    Exit Function
hell: writeFile = False
End Function

Function SetPrivateApiKey(ByVal key As String)

    key = Trim(key)
    
    If key <> API_KEY Then
        SaveSetting "VirusTotal", "config", "private_api_key", key
    End If
    
    If Len(key) > 0 Then
        API_KEY = key
        m_delayInterval = 0
        m_usingPrivateKey = True
    Else
        API_KEY = PUBLIC_API_KEY
        m_delayInterval = 17300 '4 requests per minute max
        m_usingPrivateKey = False
    End If
    
End Function

Function ReadPrivateApiKey() As String
    ReadPrivateApiKey = GetSetting("VirusTotal", "config", "private_api_key", "")
End Function

Private Sub Class_Initialize()
    NeedsDelay = False
    SetPrivateApiKey ReadPrivateApiKey()
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
End Sub

Private Function DelayIfNeeded()

    If usingPrivateKey Then Exit Function
    
    curtime = GetTickCount
    If LastQueryTime = 0 Then
        LastQueryTime = curtime
        NeedsDelay = False
    Else
        If NeedsDelay = True Then
            Interval = curtime - LastQueryTime
            If Interval < m_delayInterval Then
                Interval = m_delayInterval - Interval
                debugLogger "Must delay for " & Interval & " milliseconds"
                Timer1.Interval = Interval
                Timer1.Enabled = True
                While Timer1.Enabled
                    DoEvents
                    Sleep 20
                    If abort Then Timer1.Enabled = False
                Wend
                NeedsDelay = False
                LastQueryTime = curtime
            End If
        Else
            NeedsDelay = False
            LastQueryTime = curtime
        End If
    End If
    
End Function

'uses timer but not wininet, still uses xmlhttp
Function SubmitFile(fpath As String) As CScan

    'base example from: http://wqweto.wordpress.com/2011/07/12/vb6-using-wininet-to-post-binary-file/
    Const STR_BOUNDARY  As String = "3fbd04f5-b1ed-4060-99b9-fca7ff59c113"
    Dim sPostData       As String
    Dim fName As String
    Dim fData As String
    
    On Error Resume Next
    
    abort = False
    If Len(API_KEY) = 0 Then Class_Initialize
    
    Dim scan As New CScan
    scan.LocalFilePath = fpath
    
    Dim my_json As String
    Dim sStatus As String
    Dim status As Long
    Dim curtime As Long
    Dim success As Boolean
    Dim errorMsg As String
    
    If Timer1 Is Nothing Then
        MsgBox "Must set Timer1 reference before using SubmitFile!"
        Exit Function
    End If
    
    
    clearDebugLog
    Timer1.Enabled = False
    Timer1.Interval = m_delayInterval
    
    If Not FileExists(fpath) Then
        errorMsg = "Upload file does not exist: " & fpath
        GoTo errorExit
    End If
    
    Set http = New XMLHTTP
    
    If http Is Nothing Then
        errorMsg = "Could not create XMLHTTP Object"
        GoTo errorExit
    End If
    
    fName = fpath
    If InStr(fName, "\") > 0 Then fName = Mid$(fpath, InStrRev(fpath, "\") + 1)
    
    fData = ReadFile(fpath)
    
    sPostData = "--" & STR_BOUNDARY & vbCrLf & _
        "Content-Disposition: form-data; name=""file""; filename=""" & fName & """" & vbCrLf & _
        "Content-Type: application/octet-stream" & vbCrLf & vbCrLf & _
        fData & vbCrLf & _
        "--" & STR_BOUNDARY & vbCrLf & _
        "Content-Disposition: form-data; name=""key"" " & vbCrLf & vbCrLf & _
        API_KEY & vbCrLf & _
        "--" & STR_BOUNDARY & "--"
    
    DelayIfNeeded
    
    debugLogger "Connecting to VirusTotal to submit file: " & fName
    NeedsDelay = True 'need 4 second delay for next
    DoEvents
    LastQueryTime = GetTickCount
    
    http.Open "POST", URL_API_BASIS & URL_SCAN_FILE, False
    http.setRequestHeader "Content-Type", "multipart/form-data; boundary=" & STR_BOUNDARY
    http.send pvToByteArray(sPostData) 'blocking call until it returns..
    
    DoEvents
     
    scan.HTTP_ResponseCode = http.status
    scan.HTTP_StatusMessage = http.statusText
    scan.ParseResponse http.responseText
    Set SubmitFile = scan
    
    'Response Codes:
    ' -4: file not found.
    ' -3: public API request rate exceeded.
    ' -2: resource is currently being analyzed.
    ' -1: you do not have the required privileges (wrong API key?).
     
    debugLogger "Result: " & scan.verbose_msg & " Response Code: " & scan.response_code
      
    Exit Function

errorExit:

    debugLogger errorMsg
    scan.Errors.Add errorMsg
    scan.HadError = True
    Set SubmitFile = scan
    
    
End Function


Private Function pvToByteArray(sText As String) As Byte()
    pvToByteArray = StrConv(sText, vbFromUnicode, LANG_US)
End Function

Private Function LB_Count(hListBox As Long) As Long
   LB_Count = SendMessage(hListBox, LB_GETCOUNT, ByVal CLng(0), ByVal CLng(0))
End Function

Private Sub LB_ScrollTo(hListBox As Long, Optional ByVal index As Long = -1)
    If index = -1 Then index = LB_Count(hListBox)
    SendMessage hListBox, LB_SETTOPINDEX, ByVal index - 1, ByVal 0
End Sub


