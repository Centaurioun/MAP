Attribute VB_Name = "modMisc"
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
'Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
'Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
'Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal path As String, ByVal cbBytes As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, ByVal Source As Long, ByVal Length As Long)
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Private Type COPYDATASTRUCT
    dwFlag As Long
    cbSize As Long
    lpData As Long
End Type

Private Const WM_COPYDATA = &H4A
Private Const WM_DISPLAY_TEXT = 3

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

Global FileProps As New CFileProperties

'Private Type debugDir
'    Characteristics As Long
'    timestamp As Long
'    major As Integer
'    min As Integer
'    dbgtype As Long
'    sizeofData As Long
'    adrRawData As Long
'    ptrRawData As Long
'End Type


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

Function DiEScan(fPath As String, ByRef outVal) As Boolean
    Dim v As Long
    Dim buf As String
    Dim flags As Long
    Dim a As Long
    Dim tmp() As String
    Dim x
    Const et = "Entropy"
    
    On Error GoTo hell
    
    outVal = Empty
    If Not LoadDie Then Exit Function
    
    flags = DIE_SHOWOPTIONS Or DIE_SHOWVERSION Or DIE_SINGLELINEOUTPUT 'Or DIE_SHOWENTROPY
    buf = String(1000, Chr(0))
    'v = DiEScanA(fPath, buf, Len(buf), flags)
    v = dieScanEx(fPath, buf, Len(buf), flags, App.path & IIf(IsIde(), "\..", "") & "\die\db\")
    
    a = InStr(buf, Chr(0))
    If a > 0 Then buf = Left(buf, a - 1)
    buf = Replace(buf, vbLf, vbCrLf)
    tmp = Split(buf, ";")
    buf = Join(tmp, vbCrLf & vbTab & "   ")
    outVal = buf
    
    DiEScan = (InStr(1, buf, "Nothing found", vbBinaryCompare) < 1)
    
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
    outVal = Err.Description
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

'Function pdbPath(pe As CPEEditor, outRet As String) As Boolean
'
'    Dim rvaDebug As Long
'    Dim rawDebug As Long
'    Dim f As Long
'    Dim dd As debugDir
'    Dim bb() As Byte
'    Dim tmp As String
'    Dim a As Long, b As Long
'
'    On Error GoTo hell
'    outRet = Empty
'    If Not pe.isLoaded Then Exit Function
'
'    rvaDebug = pe.OptionalHeader.ddVirtualAddress(Debug_Data)
'    If rvaDebug = 0 Then Exit Function
'
'    rawDebug = pe.RvaToOffset(rvaDebug)
'    If rawDebug = 0 Then Exit Function
'
'    f = FreeFile
'    Open pe.LoadedFile For Binary As f
'    Get f, rawDebug + 1, dd
'
'    ReDim bb(dd.sizeofData)
'    Get f, dd.ptrRawData + 1, bb()
'    Close f
'    f = 0
'
'    tmp = StrConv(bb, vbUnicode, LANG_US)
'    a = InStr(1, tmp, ".pdb", vbTextCompare)
'    If a < 1 Then Exit Function
'    a = a + 4
'
'    b = InStrRev(tmp, ":\", a)
'    If b < 1 Then Exit Function
'    b = b - 1
'
'    outRet = Mid(tmp, b, a - b)
'    pdbPath = True
'
'    Exit Function
'hell:
'    On Error Resume Next
'    If f <> 0 Then Close f
'
'End Function
