Attribute VB_Name = "FileProps"
'License:   GPL
'Copyright: 2005 iDefense a Verisign Company
'Site:      http://labs.idefense.com
'
'Author:    David Zimmer <david@idefense.com, dzzie@yahoo.com>
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

'Used in several projects do not change interface!

 
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hWndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal path As String, ByVal cbBytes As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, ByVal Source As Long, ByVal Length As Long)
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Public Type FILEPROPERTIE
    CompanyName As String
    FileDescription As String
    FileVersion As String
    InternalName As String
    LegalCopyright As String
    OrigionalFileName As String
    ProductName As String
    ProductVersion As String
    LanguageID As String
End Type

Private Type COPYDATASTRUCT
    dwFlag As Long
    cbSize As Long
    lpData As Long
End Type

Private Const WM_COPYDATA = &H4A
Private Const WM_DISPLAY_TEXT = 3

Private Const LANG_BULGARIAN = &H2
Private Const LANG_CHINESE = &H4
Private Const LANG_CROATIAN = &H1A
Private Const LANG_CZECH = &H5
Private Const LANG_DANISH = &H6
Private Const LANG_DUTCH = &H13
Private Const LANG_ENGLISH = &H9
Private Const LANG_FINNISH = &HB
Private Const LANG_FRENCH = &HC
Private Const LANG_GERMAN = &H7
Private Const LANG_GREEK = &H8
Private Const LANG_HUNGARIAN = &HE
Private Const LANG_ICELANDIC = &HF
Private Const LANG_ITALIAN = &H10
Private Const LANG_JAPANESE = &H11
Private Const LANG_KOREAN = &H12
Private Const LANG_NEUTRAL = &H0
Private Const LANG_NORWEGIAN = &H14
Private Const LANG_POLISH = &H15
Private Const LANG_PORTUGUESE = &H16
Private Const LANG_ROMANIAN = &H18
Private Const LANG_RUSSIAN = &H19
Private Const LANG_SLOVAK = &H1B
Private Const LANG_SLOVENIAN = &H24
Private Const LANG_SPANISH = &HA
Private Const LANG_SWEDISH = &H1D
Private Const LANG_TURKISH = &H1F

Private Type sockaddr_in
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero As String * 8
End Type

Private Type sockaddr_gen
    AddressIn As sockaddr_in
    filler(0 To 7) As Byte
End Type

Private Type INTERFACE_INFO
    iiFlags  As Long
    iiAddress As sockaddr_gen
    iiBroadcastAddress As sockaddr_gen
    iiNetmask As sockaddr_gen
End Type

Private Type INTERFACEINFO
    iInfo(0 To 7) As INTERFACE_INFO
End Type

Private Const WSADESCRIPTION_LEN As Long = 256
Private Const WSASYS_STATUS_LEN  As Long = 128

Private Type WSAData
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * WSADESCRIPTION_LEN
    szSystemStatus As String * WSASYS_STATUS_LEN
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Private Declare Function socket Lib "ws2_32.dll" (ByVal af As Long, ByVal s_type As Long, ByVal Protocol As Long) As Long
Private Declare Function closesocket Lib "ws2_32.dll" (ByVal s As Long) As Long
Private Declare Function WSAIoctl Lib "ws2_32.dll" (ByVal s As Long, ByVal dwIoControlCode As Long, lpvInBuffer As Any, ByVal cbInBuffer As Long, lpvOutBuffer As Any, ByVal cbOutBuffer As Long, lpcbBytesReturned As Long, lpOverlapped As Long, lpCompletionRoutine As Long) As Long
Private Declare Sub CopyMemory2 Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, ByVal pSrc As Long, ByVal ByteLen As Long)
Private Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVR As Long, lpWSAD As WSAData) As Long

'DiE Detect it Easy signature scanner: http://ntinfo.biz/
Private hDieDll As Long
Private Const DIE_SHOWERRORS = &H1
Private Const DIE_SHOWOPTIONS = &H2
Private Const DIE_SHOWVERSION = &H4
Private Const DIE_SHOWENTROPY = &H8
Private Const DIE_SINGLELINEOUTPUT = &H10
Private Const DIE_SHOWFILEFORMATONCE = &H20
Private Declare Function DiEScanA Lib "diedll.dll" Alias "_DIE_scanA@16" (ByVal fileName As String, ByVal buf As String, ByVal bufSz As Long, ByVal flags As Long) As Long
Private Declare Function dieScanEx Lib "diedll.dll" Alias "_DIE_scanExA@20" (ByVal fileName As String, ByVal buf As String, ByVal bufSz As Long, ByVal flags As Long, ByVal dbPath As String) As Long
Private Declare Function dieVer Lib "diedll.dll" Alias "_DIE_versionA@0" () As Long
Private Declare Function SetDllDirectory Lib "kernel32" Alias "SetDllDirectoryA" (ByVal path As String) As Long

'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long


'Private Type MungeDbl
'    Value As Currency
'End Type
'
'Private Type Munge2Long
'    LoValue As Long
'    HiValue As Long
'End Type
'
'Function x64ToHex(v As Currency) As String
'    Dim c As MungeDbl
'    Dim l As Munge2Long
'    c.Value = v
'    LSet l = c
'    If l.HiValue = 0 Then
'        x64ToHex = Hex(l.LoValue)
'    Else
'        x64ToHex = Hex(l.HiValue) & Right("00000000" & Hex(l.LoValue), 8)
'    End If
'End Function
'
''handles hex strings for 32bit and 64 bit numbers, leading 00's on high part not required, of course they are on lo if there is a high..
'Function HextoX64(s As String) As Currency
'    Dim c As MungeDbl
'    Dim l As Munge2Long
'
'    Dim lo As String, hi As String
'    If Len(s) <= 8 Then
'        l.LoValue = CLng("&h" & s)
'    Else
'        lo = Right(s, 8)
'        hi = Left(s, Len(s) - 8)
'        l.LoValue = CLng("&h" & lo)
'        l.HiValue = CLng("&h" & hi)
'    End If
'
'    LSet c = l
'    HextoX64 = c.Value
'
'End Function

'Function LaunchStrings(data As String, Optional isPath As Boolean = False)
'
'    Dim b() As Byte
'    Dim f As String
'    Dim exe As String
'    Dim h As Long
'
'    On Error Resume Next
'
'    exe = App.path & IIf(IsIde(), "\..\..", "") & "\shellext.exe"
'    If Not fso.FileExists(exe) Then
'        MsgBox "Could not launch strings shellext not found", vbInformation
'        Exit Function
'    End If
'
'    If isPath Then
'        If fso.FileExists(data) Then
'            f = data
'        Else
'            MsgBox "Can not launch strings, File not found: " & data, vbInformation
'        End If
'    Else
'        b() = StrConv(dataOrPath, vbFromUnicode, LANG_US)
'        f = fso.GetFreeFileName(Environ("temp"), ".bin")
'        h = FreeFile
'    End If
'
'    Open f For Binary As h
'    Put h, , b()
'    Close h
'
'    Shell exe & " """ & f & """ /peek"
'
'End Function
 
Private Function LoadDie() As Boolean
    Dim p As String
    
    If hDieDll = 0 Then
        p = App.path & IIf(IsIde(), "\..", "") & "\die\diedll.dll"
        If fso.FileExists(p) Then
            SetDllDirectory fso.GetParentFolder(p) 'requires: xp sp1 which is fine
            hDieDll = LoadLibrary(p)               'msvcr100.dll actually wont load on xpsp0 anyway...
        End If
    End If
    
    LoadDie = (hDieDll <> 0)
    
End Function

Function DieVersion() As String
    Dim addr As Long, leng As Long, b() As Byte
    
    If Not LoadDie Then Exit Function
    
    addr = dieVer()
    If addr Then
        leng = lstrlenA(addr)
        If leng > 0 Then
            ReDim b(1 To leng)
            CopyMemory ByVal VarPtr(b(1)), ByVal addr, leng
            DieVersion = StrConv(b, vbUnicode, &H409)
        End If
    End If
    
End Function

Function DiEScan(fPath As String)
    Dim v As Long
    Dim buf As String
    Dim flags As Long
    Dim a As Long
    Dim tmp() As String
    Dim x
    Const et = "Entropy"
    
    On Error GoTo hell
    
    If Not LoadDie Then Exit Function
    
    flags = DIE_SHOWOPTIONS Or DIE_SHOWVERSION Or DIE_SINGLELINEOUTPUT Or DIE_SHOWENTROPY
    buf = String(1000, Chr(0))
    v = DiEScanA(fPath, buf, Len(buf), flags)
    
    a = InStr(buf, Chr(0))
    If a > 0 Then buf = Left(buf, a - 1)
    buf = Replace(buf, vbLf, vbCrLf)
    DiEScan = buf
    
'    tmp = Split(buf, ";")
'
'    For i = 0 To UBound(tmp)
'        tmp(i) = trim(tmp(i))
'        If Left(tmp(i), Len(et)) = et Then
'            entropy = trim(Mid(tmp(i), Len(et) + 2))
'            tmp(i) = Empty
'            Exit For
'        End If
'    Next
    
'    DiEScan = Join(tmp, ";")
    
    Exit Function
hell:
    DiEScan = Err.Description
End Function


Function HexDump(ByVal str, Optional hexOnly = 0, Optional offset As Long = 0) As String
    Dim s() As String, chars As String, tmp As String
    On Error Resume Next
    Dim ary() As Byte
   
    str = " " & str
    ary = StrConv(str, vbFromUnicode, LANG_US)
    
    chars = "   "
    For i = 1 To UBound(ary)
        tt = Hex(ary(i))
        If Len(tt) = 1 Then tt = "0" & tt
        tmp = tmp & tt & " "
        x = ary(i)
        'chars = chars & IIf((x > 32 And x < 127) Or x > 191, Chr(x), ".") 'x > 191 causes \x0 problems on non us systems... asc(chr(x)) = 0
        chars = chars & IIf((x > 32 And x < 127), Chr(x), ".")
        If i > 1 And i Mod 16 = 0 Then
            h = Hex(offset)
            While Len(h) < 6: h = "0" & h: Wend
            If hexOnly = 0 Then
                push s, h & "   " & tmp & chars
            Else
                push s, tmp
            End If
            offset = offset + 16
            tmp = Empty
            chars = "   "
        End If
    Next
    'if read length was not mod 16=0 then
    'we have part of line to account for
    If tmp <> Empty Then
        If hexOnly = 0 Then
            h = Hex(offset)
            While Len(h) < 6: h = "0" & h: Wend
            h = h & "   " & tmp
            While Len(h) <= 56: h = h & " ": Wend
            push s, h & chars
        Else
            push s, tmp
        End If
    End If
    
    HexDump = Join(s, vbCrLf)
    
    If hexOnly <> 0 Then
        HexDump = Replace(HexDump, " ", "")
        HexDump = Replace(HexDump, vbCrLf, "")
    End If
    
End Function

Function QuickInfo(fileName As String, Optional showIfBlank = True)
    Dim f As FILEPROPERTIE
    Dim tmp() As String
    
    f = FileInfo(fileName)
    
    If Len(f.CompanyName) > 0 Or showIfBlank Then push tmp, "CompanyName      " & f.CompanyName
    If Len(f.FileDescription) > 0 Or showIfBlank Then push tmp, "FileDescription  " & f.FileDescription
    If Len(f.FileVersion) > 0 Or showIfBlank Then push tmp, "FileVersion      " & f.FileVersion
    If Len(f.InternalName) > 0 Or showIfBlank Then push tmp, "InternalName     " & f.InternalName
    If Len(f.LegalCopyright) > 0 Or showIfBlank Then push tmp, "LegalCopyright   " & f.LegalCopyright
    If Len(f.OrigionalFileName) > 0 Or showIfBlank Then push tmp, "OriginalFilename " & f.OrigionalFileName
    If Len(f.ProductName) > 0 Or showIfBlank Then push tmp, "ProductName      " & f.ProductName
    If Len(f.ProductVersion) > 0 Or showIfBlank Then push tmp, "ProductVersion   " & f.ProductVersion
                
    QuickInfo = Join(tmp, vbCrLf)

End Function

Public Function FileInfo(Optional ByVal PathWithFilename As String) As FILEPROPERTIE
    ' return file-properties of given file  (EXE , DLL , OCX)
    'http://support.microsoft.com/default.aspx?scid=kb;en-us;160042
    
    If Len(PathWithFilename) = 0 Then
        Exit Function
    End If
    
    Dim lngBufferlen As Long
    Dim lngDummy As Long
    Dim lngRc As Long
    Dim lngVerPointer As Long
    Dim lngHexNumber As Long
    Dim bytBuffer() As Byte
    Dim bytBuff() As Byte
    Dim strBuffer As String
    Dim strLangCharset As String
    Dim strVersionInfo(7) As String
    Dim strTemp As String
    Dim intTemp As Integer
           
    ReDim bytBuff(500)
    
    ' size
    lngBufferlen = GetFileVersionInfoSize(PathWithFilename, lngDummy)
    If lngBufferlen > 0 Then
    
       ReDim bytBuffer(lngBufferlen)
       lngRc = GetFileVersionInfo(PathWithFilename, 0&, lngBufferlen, bytBuffer(0))
       
       If lngRc <> 0 Then
          lngRc = VerQueryValue(bytBuffer(0), "\VarFileInfo\Translation", lngVerPointer, lngBufferlen)
          If lngRc <> 0 Then
             'lngVerPointer is a pointer to four 4 bytes of Hex number,
             'first two bytes are language id, and last two bytes are code
             'page. However, strLangCharset needs a  string of
             '4 hex digits, the first two characters correspond to the
             'language id and last two the last two character correspond
             'to the code page id.
             MoveMemory bytBuff(0), lngVerPointer, lngBufferlen
             lngHexNumber = bytBuff(2) + bytBuff(3) * &H100 + bytBuff(0) * &H10000 + bytBuff(1) * &H1000000
             strLangCharset = Hex(lngHexNumber)
             'now we change the order of the language id and code page
             'and convert it into a string representation.
             'For example, it may look like 040904E4
             'Or to pull it all apart:
             '04------        = SUBLANG_ENGLISH_USA
             '--09----        = LANG_ENGLISH
             ' ----04E4 = 1252 = Codepage for Windows:Multilingual
             Do While Len(strLangCharset) < 8
                 strLangCharset = "0" & strLangCharset
             Loop
             
             If Mid(strLangCharset, 2, 2) = LANG_ENGLISH Then
               strLangCharset2 = "English (US)"
             End If

             If Mid(strLangCharset, 2, 2) = LANG_BULGARIAN Then strLangCharset2 = "Bulgarian"
             If Mid(strLangCharset, 2, 2) = LANG_FRENCH Then strLangCharset2 = "French"
             If Mid(strLangCharset, 2, 2) = LANG_NEUTRAL Then strLangCharset2 = "Neutral"

             Do While Len(strLangCharset) < 8
                 strLangCharset = "0" & strLangCharset
             Loop

             ' assign propertienames
             strVersionInfo(0) = "CompanyName"
             strVersionInfo(1) = "FileDescription"
             strVersionInfo(2) = "FileVersion"
             strVersionInfo(3) = "InternalName"
             strVersionInfo(4) = "LegalCopyright"
             strVersionInfo(5) = "OriginalFileName"
             strVersionInfo(6) = "ProductName"
             strVersionInfo(7) = "ProductVersion"
             
             Dim n As Long
             
             ' loop and get fileproperties
             For intTemp = 0 To 7
                strBuffer = String$(800, 0)
                strTemp = "\StringFileInfo\" & strLangCharset & "\" & strVersionInfo(intTemp)
                lngRc = VerQueryValue(bytBuffer(0), strTemp, lngVerPointer, lngBufferlen)
                If lngRc <> 0 Then
                   ' get and format data
                   lstrcpy strBuffer, lngVerPointer
                   n = InStr(strBuffer, Chr(0)) - 1
                   If n > 0 Then
                        strBuffer = Mid$(strBuffer, 1, n)
                        strBuffer = Replace(strBuffer, Chr(0), Empty)
                        strVersionInfo(intTemp) = trim(strBuffer)
                   End If
                 Else
                   ' property not found
                   strVersionInfo(intTemp) = ""
                End If
             Next intTemp
             
          End If
       End If
    End If
    
    ' assign array to user-defined-type
    FileInfo.CompanyName = strVersionInfo(0)
    FileInfo.FileDescription = strVersionInfo(1)
    FileInfo.FileVersion = strVersionInfo(2)
    FileInfo.InternalName = strVersionInfo(3)
    FileInfo.LegalCopyright = strVersionInfo(4)
    FileInfo.OrigionalFileName = strVersionInfo(5)
    FileInfo.ProductName = strVersionInfo(6)
    FileInfo.ProductVersion = strVersionInfo(7)
    FileInfo.LanguageID = strLangCharset2
    
End Function




Private Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Integer
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub

 



Private Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function


Function GetAllElements(lv As ListView, Optional selOnly As Boolean = False) As String
    Dim ret() As String, i As Integer, tmp As String
    Dim li As ListItem

    For i = 1 To lv.ColumnHeaders.Count
        tmp = tmp & lv.ColumnHeaders(i).text & vbTab
    Next

    push ret, tmp
    push ret, String(50, "-")

    For Each li In lv.ListItems
        
        tmp = Empty
        
        If selOnly Then
            If li.selected Then tmp = li.text & vbTab
        Else
            tmp = li.text & vbTab
        End If
        
        For i = 1 To lv.ColumnHeaders.Count - 1
            If selOnly Then
                If li.selected Then tmp = tmp & li.SubItems(i) & vbTab
            Else
                tmp = tmp & li.SubItems(i) & vbTab
            End If
        Next
        
        If Len(tmp) > 0 Then push ret, tmp
        
    Next

    GetAllElements = Join(ret, vbCrLf)

End Function

Function GetAllText(lv As ListView, Optional subItemRow As Long = 0) As String
    Dim i As Long
    Dim tmp As String, x As String
    
    For i = 1 To lv.ListItems.Count
        If subItemRow = 0 Then
            x = lv.ListItems(i).text
            If Len(x) > 0 Then
                tmp = tmp & x & vbCrLf
            End If
        Else
            x = lv.ListItems(i).SubItems(subItemRow)
            If Len(x) > 0 Then
                tmp = tmp & x & vbCrLf
            End If
        End If
    Next
    
    GetAllText = tmp
End Function

Function IsIde() As Boolean
    On Error GoTo hell
    Debug.Print 1 \ 0
Exit Function
hell: IsIde = True
End Function

Sub SetLiColor(li As ListItem, newcolor As Long)
    Dim f As ListSubItem
'    On Error Resume Next
    li.ForeColor = newcolor
    For Each f In li.ListSubItems
        f.ForeColor = newcolor
    Next
End Sub

'ported from Detect It Easy - Binary::calculateEntropy
'   https://github.com/horsicq/DIE-engine/blob/master/binary.cpp#L2319
Function fileEntropy(pth As String, Optional offset As Long = 0, Optional leng As Long = -1) As Single
    
    Dim sz As Long
    Dim fEntropy As Single
    Dim bytes(255) As Single
    Dim temp As Single
    Dim nSize As Long
    Dim nTemp As Long
    Const BUFFER_SIZE = &H1000
    Dim buf() As Byte
    Dim f As Long
    
    On Error Resume Next
    
    f = FreeFile
    Open pth For Binary Access Read As f
    If Err.Number <> 0 Then GoTo ret0
    
    sz = LOF(f) - 1
    
    If leng = 0 Then GoTo ret0
    
    If leng = -1 Then
        leng = sz - offset
        If leng = 0 Then GoTo ret0
    End If
    
    If offset >= sz Then GoTo ret0
    If offset + leng > sz Then GoTo ret0
    
    Seek f, offset
    nSize = leng
    fEntropy = 1.44269504088896
    ReDim buf(BUFFER_SIZE)
    
    'read the file in chunks and count how many times each byte value occurs
    While (nSize > 0)
        nTemp = IIf(nSize < BUFFER_SIZE, nSize, BUFFER_SIZE)
        If nTemp <> BUFFER_SIZE Then ReDim buf(nTemp) 'last chunk, partial buffer
        Get f, , buf()
        For i = 0 To UBound(buf)
            bytes(buf(i)) = bytes(buf(i)) + 1
        Next
        nSize = nSize - nTemp
    Wend
    
    For i = 0 To UBound(bytes)
        temp = bytes(i) / CSng(leng)
        If temp <> 0 Then
            fEntropy = fEntropy + (-Log(temp) / Log(2)) * bytes(i)
        End If
    Next
    
    Close f
    fileEntropy = fEntropy / CSng(leng)
    
Exit Function
ret0:
    Close f
End Function


Function memEntropy(buf() As Byte, Optional offset As Long = 0, Optional leng As Long = -1) As Single
    
    Dim sz As Long
    Dim fEntropy As Single
    Dim bytes(255) As Single
    Dim temp As Single
    Const BUFFER_SIZE = &H1000
    
    sz = UBound(buf)
    
    If leng = 0 Then GoTo ret0
    If leng = -1 Then
        leng = sz - offset
        If leng = 0 Then GoTo ret0
    End If
    
    If offset >= sz Then GoTo ret0
    If offset + leng > sz Then GoTo ret0
    
    fEntropy = 1.44269504088896
    
    While (offset < sz)
        'count each byte value occurance
        bytes(buf(offset)) = bytes(buf(offset)) + 1
        offset = offset + 1
    Wend
    
    For i = 0 To UBound(bytes)
        temp = bytes(i) / CSng(leng)
        If temp <> 0 Then
            fEntropy = fEntropy + (-Log(temp) / Log(2)) * bytes(i)
        End If
    Next
    
    memEntropy = fEntropy / CSng(leng)
    
Exit Function
ret0:
End Function

