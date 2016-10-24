Attribute VB_Name = "Module1"
Option Explicit
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

Global fso As New clsFileSystem
Global hash As New CWinHash
Global dlg As New CCmnDlg 'clsCmnDlg
Global minStrLen As Long
Global Const LANG_US = &H409
Global myIcon As IPictureDisp

Public Const IMAGE_NT_OPTIONAL_HDR32_MAGIC = &H10B

Public Type IMAGEDOSHEADER
    e_magic As Integer
    e_cblp As Integer
    e_cp As Integer
    e_crlc As Integer
    e_cparhdr As Integer
    e_minalloc As Integer
    e_maxalloc As Integer
    e_ss As Integer
    e_sp As Integer
    e_csum As Integer
    e_ip As Integer
    e_cs As Integer
    e_lfarlc As Integer
    e_ovno As Integer
    e_res(1 To 4) As Integer
    e_oemid As Integer
    e_oeminfo As Integer
    e_res2(1 To 10)    As Integer
    e_lfanew As Long
End Type

Public Type IMAGE_FILE_HEADER
    Machine As Integer
    NumberOfSections As Integer
    TimeDateStamp As Long
    PointerToSymbolTable As Long
    NumberOfSymbols As Long
    SizeOfOptionalHeader As Integer
    Characteristics As Integer
End Type

Public Type IMAGE_DATA_DIRECTORY
    VirtualAddress As Long
    size As Long
End Type

Public Type IMAGE_OPTIONAL_HEADER
    Magic As Integer
    MajorLinkerVersion As Byte
    MinorLinkerVersion As Byte
    SizeOfCode As Long
    SizeOfInitializedData As Long
    SizeOfUninitializedData As Long
    AddressOfEntryPoint As Long
    BaseOfCode As Long
    BaseOfData As Long
    ImageBase As Long
    SectionAlignment As Long
    FileAlignment As Long
    MajorOperatingSystemVersion As Integer
    MinorOperatingSystemVersion As Integer
    MajorImageVersion As Integer
    MinorImageVersion As Integer
    MajorSubsystemVersion As Integer
    MinorSubsystemVersion As Integer
    Win32VersionValue As Long
    SizeOfImage As Long
    SizeOfHeaders As Long
    CheckSum As Long
    Subsystem As Integer
    DllCharacteristics As Integer
    SizeOfStackReserve As Long
    SizeOfStackCommit As Long
    SizeOfHeapReserve As Long
    SizeOfHeapCommit As Long
    LoaderFlags As Long
    NumberOfRvaAndSizes As Long
    DataDirectory(0 To 15) As IMAGE_DATA_DIRECTORY
End Type

Public Type IMAGE_OPTIONAL_HEADER_64
    Magic As Integer
    MajorLinkerVersion As Byte
    MinorLinkerVersion As Byte
    SizeOfCode As Long
    SizeOfInitializedData As Long
    SizeOfUninitializedData As Long
    AddressOfEntryPoint As Long
    BaseOfCode As Long
    'BaseOfData As Long                        'this was removed for pe32+
    ImageBase As Double                        'changed
    SectionAlignment As Long
    FileAlignment As Long
    MajorOperatingSystemVersion As Integer
    MinorOperatingSystemVersion As Integer
    MajorImageVersion As Integer
    MinorImageVersion As Integer
    MajorSubsystemVersion As Integer
    MinorSubsystemVersion As Integer
    Win32VersionValue As Long
    SizeOfImage As Long
    SizeOfHeaders As Long
    CheckSum As Long
    Subsystem As Integer
    DllCharacteristics As Integer
    SizeOfStackReserve As Double                        'changed
    SizeOfStackCommit As Double                        'changed
    SizeOfHeapReserve As Double                        'changed
    SizeOfHeapCommit As Double                        'changed
    LoaderFlags As Long
    NumberOfRvaAndSizes As Long
    DataDirectory(0 To 15) As IMAGE_DATA_DIRECTORY
End Type

Public Type IMAGE_NT_HEADERS
    Signature As String * 4
    FileHeader As IMAGE_FILE_HEADER
    'OptionalHeader As IMAGE_OPTIONAL_HEADER
End Type

'Enum eDATA_DIRECTORY
'    Export_Table = 0
'    Import_Table = 1
'    Resource_Table = 2
'    Exception_Table = 3
'    Certificate_Table = 4
'    Relocation_Table = 5
'    Debug_Data = 6
'    Architecture_Data = 7
'    Machine_Value = 8
'    TLS_Table = 9
'    Load_Configuration_Table = 10
'    Bound_Import_Table = 11
'    Import_Address_Table = 12
'    Delay_Import_Descriptor = 13
'    CLI_Header = 14
'    Reserved = 15
'End Enum

Public Enum tmMsgs
        EM_UNDO = &HC7
        EM_CANUNDO = &HC6
        EM_SETWORDBREAKPROC = &HD0
        EM_SETTABSTOPS = &HCB
        EM_SETSEL = &HB1
        EM_SETRECTNP = &HB4
        EM_SETRECT = &HB3
        EM_SETREADONLY = &HCF
        EM_SETPASSWORDCHAR = &HCC
        EM_SETMODIFY = &HB9
        EM_SCROLLCARET = &HB7
        EM_SETHANDLE = &HBC
        EM_SCROLL = &HB5
        EM_REPLACESEL = &HC2
        EM_LINESCROLL = &HB6
        EM_LINELENGTH = &HC1
        EM_LINEINDEX = &HBB
        EM_LINEFROMCHAR = &HC9
        EM_LIMITTEXT = &HC5
        EM_GETWORDBREAKPROC = &HD1
        EM_GETTHUMB = &HBE
        EM_GETRECT = &HB2
        EM_GETSEL = &HB0
        EM_GETPASSWORDCHAR = &HD2
        EM_GETMODIFY = &HB8
        EM_GETLINECOUNT = &HBA
        EM_GETLINE = &HC4
        EM_GETHANDLE = &HBD
        EM_GETFIRSTVISIBLELINE = &HCE
        EM_FMTLINES = &HC8
        EM_EMPTYUNDOBUFFER = &HCD
        EM_SETMARGINS = &HD3
End Enum

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpszOp As String, ByVal lpszFile As String, ByVal lpszParams As String, ByVal LpszDir As String, ByVal FsShowCmd As Long) As Long

Private Declare Function Wow64DisableWow64FsRedirection Lib "kernel32.dll" (ByRef old As Long) As Long
Private Declare Function Wow64RevertWow64FsRedirection Lib "kernel32.dll" (ByRef old As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Dim firstHandle As Long

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long
Private Const VER_PLATFORM_WIN32_NT = 2
Private Const TOKEN_READ As Long = &H20008
Private Const TOKEN_QUERY As Long = &H8
Private Const TOKEN_ELEVATION_TYPE As Long = 18
Private Declare Function IsUserAnAdmin Lib "shell32" Alias "#680" () As Integer
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function GetTokenInformation Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal TokenInformationClass As Long, TokenInformation As Any, ByVal TokenInformationLength As Long, ReturnLength As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

'vista+ only
Private Type TOKEN_ELEVATION
    TokenIsElevated As Long
End Type

Private Enum TOKEN_INFORMATION_CLASS
    TokenUser = 1
    TokenGroups
    TokenPrivileges
    TokenOwner
    TokenPrimaryGroup
    TokenDefaultDacl
    TokenSource
    TokenType
    TokenImpersonationLevel
    TokenStatistics
    TokenRestrictedSids
    TokenSessionId
    TokenGroupsAndPrivileges
    TokenSessionReference
    TokenSandBoxInert
    TokenAuditPolicy
    TokenOrigin
    tokenElevationType
    TokenLinkedToken
    TokenElevation
    TokenHasRestrictions
    TokenAccessInformation
    TokenVirtualizationAllowed
    TokenVirtualizationEnabled
    TokenIntegrityLevel
    TokenUIAccess
    TokenMandatoryPolicy
    TokenLogonSid
    MaxTokenInfoClass  'MaxTokenInfoClass should always be the last enum
End Enum

Private Type SHELLEXECUTEINFO
        cbSize        As Long
        fMask         As Long
        hWnd          As Long
        lpVerb        As String
        lpFile        As String
        lpParameters  As String
        lpDirectory   As String
        nShow         As Long
        hInstApp      As Long
        lpIDList      As Long     'Optional
        lpClass       As String   'Optional
        hkeyClass     As Long     'Optional
        dwHotKey      As Long     'Optional
        hIcon         As Long     'Optional
        hProcess      As Long     'Optional
End Type

Private Declare Function ShellExecuteEx Lib "shell32.dll" Alias "ShellExecuteExA" (lpSEI As SHELLEXECUTEINFO) As Long

Public Enum EShellShowConstants
        essSW_HIDE = 0
        essSW_SHOWNORMAL = 1
        essSW_SHOWMINIMIZED = 2
        essSW_MAXIMIZE = 3
        essSW_SHOWMAXIMIZED = 3
        essSW_SHOWNOACTIVATE = 4
        essSW_SHOW = 5
        essSW_MINIMIZE = 6
        essSW_SHOWMINNOACTIVE = 7
        essSW_SHOWNA = 8
        essSW_RESTORE = 9
        essSW_SHOWDEFAULT = 10
End Enum

Public OSVersion As String
Public LinkerVersion As String
Public ImageVersion As String
Public SubSysVersion As String

Function PEVersionReport(Optional compact As Boolean = False) As String
    Dim tmp() As String
    
    If OSVersion = Empty Then Exit Function 'marker that its not set...
    
    If compact Then
        PEVersionReport = "PEVersion:" & "OS:" & OSVersion & " Link:" & LinkerVersion & _
                                         " Img:" & ImageVersion & " SubSys:" & SubSysVersion
    Else
        push tmp, rpad("OSVersion: ", 16) & OSVersion
        push tmp, rpad("LinkerVersion: ", 16) & LinkerVersion
        push tmp, rpad("ImageVersion: ", 16) & ImageVersion
        push tmp, rpad("SubSysVersion: ", 16) & SubSysVersion
        PEVersionReport = Join(tmp, vbCrLf)
    End If
    
End Function

Private Sub SetVersions64(ih As IMAGE_OPTIONAL_HEADER_64)
    With ih
        LinkerVersion = .MajorLinkerVersion & "." & .MinorLinkerVersion
        OSVersion = .MajorOperatingSystemVersion & "." & .MinorOperatingSystemVersion
        ImageVersion = .MajorImageVersion & "." & .MinorImageVersion
        SubSysVersion = .MajorSubsystemVersion & "." & .MinorSubsystemVersion
    End With
End Sub

Private Sub SetVersions(ih As IMAGE_OPTIONAL_HEADER)
    With ih
        LinkerVersion = .MajorLinkerVersion & "." & .MinorLinkerVersion
        OSVersion = .MajorOperatingSystemVersion & "." & .MinorOperatingSystemVersion
        ImageVersion = .MajorImageVersion & "." & .MinorImageVersion
        SubSysVersion = .MajorSubsystemVersion & "." & .MinorSubsystemVersion
    End With
End Sub

Public Function RunElevated(ByVal FilePath As String, Optional ShellShowType As EShellShowConstants = essSW_SHOWNORMAL, Optional ByVal hWndOwner As Long = 0, Optional EXEParameters As String = "") As Boolean
    Dim SEI As SHELLEXECUTEINFO
    Const SEE_MASK_DEFAULT = &H0
    
    On Error GoTo Err

    'Fill the SEI structure
    With SEI
        .cbSize = Len(SEI)                  ' Bytes of the structure
        .fMask = SEE_MASK_DEFAULT           ' Check MSDN for more info on Mask
        .lpFile = FilePath                  ' Program Path
        .nShow = ShellShowType              ' How the program will be displayed
        .lpDirectory = PathGetFolder(FilePath)
        .lpParameters = EXEParameters       ' Each parameter must be separated by space. If the lpFile member specifies a document file, lpParameters should be NULL.
        .hWnd = hWndOwner                   ' Owner window handle
        .lpVerb = "runas"
    End With

    RunElevated = ShellExecuteEx(SEI)   ' Execute the program, return success or failure

    Exit Function
Err:
    RunElevated = False
End Function

Private Function PathGetFolder(s) As String
    If InStr(1, s, "\") > 0 Then
        PathGetFolder = Mid(s, 1, InStrRev(s, "\"))
    End If
End Function

Public Function IsVistaPlus() As Boolean
    Dim OSVersion As OSVERSIONINFO
    OSVersion.dwOSVersionInfoSize = Len(OSVersion)
    If GetVersionEx(OSVersion) = 0 Then Exit Function
    If OSVersion.dwPlatformId <> VER_PLATFORM_WIN32_NT Or OSVersion.dwMajorVersion < 6 Then Exit Function
    IsVistaPlus = True
End Function

Public Function IsUserAnAdministrator() As Boolean
    'http://www.davidmoore.info/2011/06/20/how-to-check-if-the-current-user-is-an-administrator-even-if-uac-is-on/
    Dim result As Long
    Dim hProcessID As Long
    Dim hToken As Long
    Dim lReturnLength As Long
    Dim tokenElevationType As Long
    
    On Error GoTo hell
    
    IsUserAnAdministrator = False
    
    If IsUserAnAdmin() Then
        IsUserAnAdministrator = True
        Exit Function
    End If
    
    'If we’re on Vista onwards, check for UAC elevation token
    'as we may be an admin but we’re not elevated yet, so the
    'IsUserAnAdmin() function will return false
    Dim OSVersion As OSVERSIONINFO
    OSVersion.dwOSVersionInfoSize = Len(OSVersion)
    
    If GetVersionEx(OSVersion) = 0 Then Exit Function
    
    'If the user is not on Vista or greater, then there’s no UAC, so don’t bother checking.
    If OSVersion.dwPlatformId <> VER_PLATFORM_WIN32_NT Or OSVersion.dwMajorVersion < 6 Then Exit Function
   
    hProcessID = GetCurrentProcess() 'get the token for the current process
    If hProcessID = 0 Then Exit Function
    
    If OpenProcessToken(hProcessID, TOKEN_READ, hToken) = 1 Then
        result = GetTokenInformation(hToken, TOKEN_ELEVATION_TYPE, tokenElevationType, 4, lReturnLength)
        If result <> 0 Then
             If tokenElevationType <> 1 Then IsUserAnAdministrator = True
        End If
        CloseHandle hToken
    End If
    

Exit Function
hell:
    
End Function

Function IsProcessElevated() As Boolean

    Dim fIsElevated As Boolean
    Dim dwError As Long
    Dim hToken As Long

    'Open the primary access token of the process with TOKEN_QUERY.
    If OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, hToken) = 0 Then GoTo cleanup
     
    Dim elevation As TOKEN_ELEVATION
    Dim dwSize As Long
    If GetTokenInformation(hToken, TOKEN_INFORMATION_CLASS.TokenElevation, elevation, Len(elevation), dwSize) = 0 Then
        'When the process is run on operating systems prior to Windows Vista, GetTokenInformation returns FALSE with the
        'ERROR_INVALID_PARAMETER error code because TokenElevation is not supported on those operating systems.
         dwError = Err.LastDllError
         GoTo cleanup
    End If

    fIsElevated = IIf(elevation.TokenIsElevated = 0, False, True)

cleanup:
    If hToken Then CloseHandle (hToken)
    'if ERROR_SUCCESS <> dwError then err.Raise
    IsProcessElevated = fIsElevated
End Function

Function pad(v, Optional l As Long = 8)
    On Error GoTo hell
    Dim X As Long
    X = Len(v)
    If X < l Then
        pad = String(l - X, " ") & v
    Else
hell:
        pad = v
    End If
End Function

Function rpad(v, Optional l As Long = 10)
    On Error GoTo hell
    Dim X As Long
    X = Len(v)
    If X < l Then
        rpad = v & String(l - X, " ")
    Else
hell:
        rpad = v
    End If
End Function

Public Sub LV_ColumnSort(ListViewControl As ListView, Column As ColumnHeader)
     On Error Resume Next
    With ListViewControl
       If .SortKey <> Column.index - 1 Then
             .SortKey = Column.index - 1
             .SortOrder = lvwAscending
       Else
             If .SortOrder = lvwAscending Then
              .SortOrder = lvwDescending
             Else
              .SortOrder = lvwAscending
             End If
       End If
       .Sorted = -1
    End With
End Sub

Public Function GetShortName(sFile As String) As String
    Dim sShortFile As String * 67
    Dim lResult As Long
    
    'the path must actually exist to get the short path name !!
    If Not fso.FileExists(sFile) Then 'fso.WriteFile sFile, ""
        GetShortName = sFile
        Exit Function
    End If
        
    'Make a call to the GetShortPathName API
    lResult = GetShortPathName(sFile, sShortFile, _
    Len(sShortFile))

    'Trim out unused characters from the string.
    GetShortName = Left$(sShortFile, lResult)
    
    If Len(GetShortName) = 0 Then GetShortName = sFile

End Function

Function DisableRedir() As Long
    
    If firstHandle <> 0 Then Exit Function 'defaults to 0 on subsequent calls...
    
    If GetProcAddress(GetModuleHandle("kernel32.dll"), "Wow64DisableWow64FsRedirection") = 0 Then
        Exit Function
    End If
    
    Dim r As Long, lastRedir As Long
    r = Wow64DisableWow64FsRedirection(lastRedir)
    firstHandle = IIf(r <> 0, lastRedir, 0)
    DisableRedir = firstHandle
    
End Function

Function RevertRedir(old As Long) As Boolean 'really only reverts firstHandle when called...
    
    If old = 0 Then Exit Function
    If old <> firstHandle Then Exit Function
    
    If GetProcAddress(GetModuleHandle("kernel32.dll"), "Wow64RevertWow64FsRedirection") = 0 Then
        Exit Function
    End If
    
    Dim r As Long
    r = Wow64RevertWow64FsRedirection(old)
    If r <> 0 Then RevertRedir = True
    firstHandle = 0
    
End Function


Function Google(hash As String, Optional hWnd As Long = 0)
    Const u = "http://www.google.com/#hl=en&output=search&q="
    ShellExecute hWnd, "Open", u & hash, "", "C:\", 1
End Function

Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim X As Long
    X = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub

Sub SaveMySetting(key, value)
    SaveSetting "iDefense", "ShellExt", key, value
End Sub

Function GetMySetting(key, def)
    GetMySetting = GetSetting("iDefense", "ShellExt", key, def)
End Function

Sub SaveFormSizeAnPosition(f As Form)
    Dim s As String
    If f.WindowState <> 0 Then Exit Sub 'vbnormal
    s = f.Left & "," & f.Top & "," & f.Width & "," & f.Height
    SaveMySetting f.name & "_pos", s
End Sub

Sub RestoreFormSizeAnPosition(f As Form)
    On Error GoTo hell
    Dim s
    
    s = GetMySetting(f.name & "_pos", "")
    
    If Len(s) = 0 Then Exit Sub
    If occuranceCount(s, ",") <> 3 Then Exit Sub
    
    s = Split(s, ",")
    f.Left = s(0)
    f.Top = s(1)
    f.Width = s(2)
    f.Height = s(3)
    
    Exit Sub
hell:
End Sub

Function occuranceCount(haystack, match) As Long
    On Error Resume Next
    Dim tmp
    tmp = Split(haystack, match, , vbTextCompare)
    occuranceCount = UBound(tmp) + 1
    If Err.Number <> 0 Then occuranceCount = 0
End Function

Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    Dim i As Long
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function

Private Function CompiledDate(stamp As Double) As String

    On Error Resume Next
    Dim base As Date
    Dim compiled As Date
    
    base = DateSerial(1970, 1, 1)
    compiled = DateAdd("s", stamp, base)
    CompiledDate = Format(compiled, "ddd, mmm d yyyy, h:nn:ss ")

End Function

Function GetCompileDateOrType(fPath As String, Optional ByRef out_isType As Boolean, Optional ByRef out_isPE As Boolean) As String
    On Error GoTo hell
        
        Dim i As Long
        Dim f As Long
        Dim buf(20) As Byte
        Dim sBuf As String
        Dim fs As Long
        
        Dim DOSHEADER As IMAGEDOSHEADER
        Dim NTHEADER As IMAGE_NT_HEADERS
        Dim opt As IMAGE_OPTIONAL_HEADER
        Dim opt64 As IMAGE_OPTIONAL_HEADER_64
        Dim isNative As Boolean
        Dim cli As Long
  
        OSVersion = Empty 'module level
        out_isType = False
        
        fs = DisableRedir()
        If Not fso.FileExists(fPath) Then Exit Function
            
        f = FreeFile
        
        Open fPath For Binary Access Read As f
        Get f, , DOSHEADER
        
        If DOSHEADER.e_magic <> &H5A4D Then
            Get f, 1, buf()
            Close f
            sBuf = StrConv(buf(), vbUnicode, LANG_US)
            GetCompileDateOrType = DetectFileType(sBuf, fPath)
            out_isType = True
            RevertRedir fs
            Exit Function
        End If
        
        Get f, DOSHEADER.e_lfanew + 1, NTHEADER
        
        If NTHEADER.Signature <> "PE" & Chr(0) & Chr(0) Then
            Get f, 1, buf()
            Close f
            sBuf = StrConv(buf(), vbUnicode, LANG_US)
            GetCompileDateOrType = DetectFileType(sBuf, fPath)
            out_isType = True
            RevertRedir fs
            Exit Function
        End If
        
        out_isPE = True
        GetCompileDateOrType = CompiledDate(CDbl(NTHEADER.FileHeader.TimeDateStamp))
        
        If is64Bit(NTHEADER.FileHeader.Machine) Then
            Get f, , opt64
            SetVersions64 opt64
            cli = opt64.DataDirectory(eDATA_DIRECTORY.CLI_Header).VirtualAddress
            If opt64.Subsystem = 1 Then isNative = True
            GetCompileDateOrType = GetCompileDateOrType & " - 64 Bit"
        Else
            Get f, , opt
            SetVersions opt
            cli = opt.DataDirectory(eDATA_DIRECTORY.CLI_Header).VirtualAddress
            If opt.Subsystem = 1 Then isNative = True
            If is32Bit(NTHEADER.FileHeader.Machine) Then GetCompileDateOrType = GetCompileDateOrType & " - 32 Bit"
        End If
        
        Close f
        RevertRedir fs

        GetCompileDateOrType = GetCompileDateOrType & GetDotNetAttributes(fPath, cli)

        If isNative Then
            GetCompileDateOrType = GetCompileDateOrType & " Native"
        Else
            GetCompileDateOrType = GetCompileDateOrType & isExe_orDll(NTHEADER.FileHeader.Characteristics)
        End If
        
        GetCompileDateOrType = Replace(GetCompileDateOrType, "  ", " ")
        
Exit Function
hell:
    
    Close f
    out_isType = True
    GetCompileDateOrType = Err.Description
    RevertRedir fs
End Function

Private Function GetDotNetAttributes(fPath As String, cli As Long) As String
    
    'ok we are going to need to load it more fully...
    Dim pe As New CPEEditor
    Dim tmp As String
    Dim fs As Long
    
    If cli = 0 Then Exit Function
    
    tmp = " .NET"
    
    fs = DisableRedir()
    If pe.LoadFile(fPath) Then
        tmp = tmp & " " & pe.dotNetVersion
        If pe.isDotNetAnyCpu Then tmp = tmp & " AnyCPU "
    End If
    RevertRedir fs
    
    GetDotNetAttributes = tmp

End Function

Private Function isExe_orDll(chart As Integer) As String
    'IMAGE_FILE_DLL 0x2000, IMAGE_FILE_EXECUTABLE_IMAGE x0002
    Dim isExecutable As Boolean
    Dim isDll As Boolean
    
    If (chart And 2) = 2 Then
        isExecutable = True
        If (chart And &H2000) = &H2000 Then
            isDll = True
            isExe_orDll = " DLL"
        Else
            isExe_orDll = " EXE"
        End If
    End If
    
End Function


Function is64Bit(m As Integer) As Boolean
    If m = &H8664 Or m = &H200 Then 'AMD64 or IA64
        is64Bit = True
    End If
End Function

Private Function is32Bit(m As Integer) As Boolean
    If m = &H14C Then '386
        is32Bit = True
    End If
End Function

 

Private Function DetectFileType(buf As String, fname As String) As String
    Dim dot As Long
    On Error GoTo hell
    
    If VBA.Left(buf, 2) = "PK" Then '1)"PK\003\004" , 2) "PK\005\006" (empty archive), or "PK\007\008" (spanned archieve).
        DetectFileType = "Zip file"
    ElseIf InStr(1, buf, "%PDF", vbTextCompare) > 0 Then
        DetectFileType = "Pdf File"
    'ElseIf VBA.Left(buf, 8) = Chr(&HD0) & Chr(&HCF) & Chr(&H11) & Chr(&HE0) & _
    '                          Chr(&HA1) & Chr(&HB1) & Chr(&H1A) & Chr(&HE1) Then
    '    DetectFileType = "MSI Installer"
    ElseIf VBA.Left(buf, 4) = Chr(&HD0) & Chr(&HCF) & Chr(&H11) & Chr(&HE0) Then
        DetectFileType = "Ole Document"
    ElseIf VBA.Left(buf, 4) = "L" & Chr(0) & Chr(0) & Chr(0) Then
        DetectFileType = "Link File"
    ElseIf VBA.Left(buf, 3) = "CWS" Then
        DetectFileType = "Compressed SWF File"
    ElseIf VBA.Left(buf, 3) = "FWS" Then
        DetectFileType = "SWF File"
    ElseIf VBA.Left(buf, 3) = "ZWS" Then
        DetectFileType = "LZMA Compressed SWF File"
    ElseIf VBA.Left(buf, 4) = "Rar!" Then
        DetectFileType = "RAR File"
    ElseIf VBA.Left(buf, 5) = "{\rtf" Then
        DetectFileType = "RTF Document"
    Else
        dot = InStrRev(fname, ".")
        If dot > 0 And dot <> Len(fname) Then
            DetectFileType = UCase(Mid(fname, dot + 1)) & " File"
            If Len(DetectFileType) > 12 Then DetectFileType = "Unknown File Type." '<-- subtle identifier ending period
        Else
            DetectFileType = "Unknown File Type"
        End If
    End If
    
    Exit Function
hell: DetectFileType = "Unknown FileType" '<-- subtle error identifier in missing space...
      Err.Clear
    
End Function


Sub ScrollToLine(t As Object, X As Integer)
     X = X - TopLineIndex(t)
     ScrollIncremental t, , X
End Sub

Sub ScrollIncremental(t As Object, Optional horz As Integer = 0, Optional vert As Integer = 0)
    'lParam&  The low-order 2 bytes specify the number of vertical
    '          lines to scroll. The high-order 2 bytes specify the
    '          number of horizontal columns to scroll. A positive
    '          value for lParam& causes text to scroll upward or to the
    '          left. A negative value causes text to scroll downward or
    '          to the right.
    ' r&       Indicates the number of lines actually scrolled.
    
    Dim r As Long
    r = CLng(&H10000 * horz) + vert
    r = SendMessage(t.hWnd, EM_LINESCROLL, 0, ByVal r)

End Sub

Function TopLineIndex(X As Object) As Long
    TopLineIndex = SendMessage(X.hWnd, EM_GETFIRSTVISIBLELINE, 0, ByVal 0&) + 1
End Function


Function sizeLvCol(lv As ListView)
    On Error Resume Next
    lv.ColumnHeaders(lv.ColumnHeaders.count).Width = lv.Width - lv.ColumnHeaders(lv.ColumnHeaders.count - 1).Left - 100
End Function
