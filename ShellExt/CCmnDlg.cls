VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCmnDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Author:   dzzie@yahoo.com
'Site:     http://sandsprite.com

Option Explicit

Const LANG_US = &H409

Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type oColorDlg
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Public Enum FilterTypes
    textFiles = 0
    htmlFiles = 1
    exeFiles = 2
    zipFiles = 3
    AllFiles = 4
    CustomFilter = 5
End Enum

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As oColorDlg) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)

Private Const LF_FACESIZE = 32

Private Type LOGFONT
   lfHeight As Long
   lfWidth As Long
   lfEscapement As Long
   lfOrientation As Long
   lfWeight As Long
   lfItalic As Byte
   lfUnderline As Byte
   lfStrikeOut As Byte
   lfCharSet As Byte
   lfOutPrecision As Byte
   lfClipPrecision As Byte
   lfQuality As Byte
   lfPitchAndFamily As Byte
   lfFaceName(LF_FACESIZE) As Byte
End Type

Private Type udtFont
   lStructSize As Long
   hWndOwner As Long           '  caller's window handle
   hdc As Long                 '  printer DC/IC or NULL
   lpLogFont As Long            ' LOGFONT          '  ptr. to a LOGFONT struct
   iPointSize As Long          '  10 * size in points of selected font
   flags As Long               '  enum. type flags
   rgbColors As Long           '  returned text color
   lCustData As Long           '  data passed to hook fn.
   lpfnHook As Long            '  ptr. to hook function
   lpTemplateName As String    '  custom template name
   hInstance As Long           '  instance handle of.EXE that
                               '    contains cust. dlg. template
   lpszStyle As String         '    return the style field here
                               '    must be LF_FACESIZE or bigger
   nFontType As Integer        '  same value reported to the EnumFonts
                               '    call back with the extra FONTTYPE_
                               '    bits added
   MISSING_ALIGNMENT As Integer
   nSizeMin As Long            '  minimum pt size allowed &
   nSizeMax As Long            '  max pt size allowed if
                               '    CF_LIMITSIZE is used
End Type

Private Declare Function SelectFont Lib "comdlg32.dll" Alias "ChooseFontA" (uf As udtFont) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "USER32" (ByVal hWnd As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long

Private Const GMEM_FIXED = &H0
Private Const GMEM_ZEROINIT = &H40
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)
Private Const CF_SCREENFONTS = &H1
Private Const CF_PRINTERFONTS = &H2
Private Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Private Const CF_SHOWHELP = &H4&
Private Const CF_ENABLEHOOK = &H8&
Private Const CF_ENABLETEMPLATE = &H10&
Private Const CF_ENABLETEMPLATEHANDLE = &H20&
Private Const CF_INITTOLOGFONTSTRUCT = &H40&
Private Const CF_USESTYLE = &H80&
Private Const CF_EFFECTS = &H100&
Private Const CF_APPLY = &H200&
Private Const CF_ANSIONLY = &H400&
Private Const CF_SCRIPTSONLY = CF_ANSIONLY
Private Const CF_NOVECTORFONTS = &H800&
Private Const CF_NOOEMFONTS = CF_NOVECTORFONTS
Private Const CF_NOSIMULATIONS = &H1000&
Private Const CF_LIMITSIZE = &H2000&
Private Const CF_FIXEDPITCHONLY = &H4000&
Private Const CF_WYSIWYG = &H8000 '  must also have CF_SCREENFONTS CF_PRINTERFONTS
Private Const CF_FORCEFONTEXIST = &H10000
Private Const CF_SCALABLEONLY = &H20000
Private Const CF_TTONLY = &H40000
Private Const CF_NOFACESEL = &H80000
Private Const CF_NOSTYLESEL = &H100000
Private Const CF_NOSIZESEL = &H200000
Private Const CF_SELECTSCRIPT = &H400000
Private Const CF_NOSCRIPTSEL = &H800000
Private Const CF_NOVERTFONTS = &H1000000
Private Const LOGPIXELSY = 90

Private o As OPENFILENAME
Private filters(6) As String
Private extensions(6) As String
Private errOnCancel As Boolean

Property Let ErrorOnCancel(bln As Boolean)
    errOnCancel = bln
End Property

Property Get ErrorOnCancel() As Boolean
    ErrorOnCancel = errOnCancel
End Property

Sub SetCustomFilter(displayText As String, Optional wildCardExtMatch = "*.*")
    filters(5) = "____" + Chr$(0) + "___" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    filters(5) = Replace(filters(5), "____", displayText)
    filters(5) = Replace(filters(5), "___", wildCardExtMatch)
    extensions(5) = Replace(wildCardExtMatch, "*", "")
End Sub

Private Sub Class_Initialize()
   
    'If Not isRegistered And Not isInitalized Then TellThemAllAboutIt
    
    filters(0) = "Text Files (*.txt)" + Chr$(0) + "*.txt" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    filters(1) = "Html Files (*.htm*)" + Chr$(0) + "*.htm*" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    filters(2) = "Exe Files (*.exe)" + Chr$(0) + "*.exe" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    filters(3) = "Zip Files (*.zip)" + Chr$(0) + "*.zip" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    filters(4) = "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)

    extensions(0) = "txt"
    extensions(1) = "html"
    extensions(2) = "exe"
    extensions(3) = "zip"
    extensions(4) = "bin"
    
End Sub

Function OpenDialog(filt As FilterTypes, Optional initDir As String, Optional title As String, Optional pHwnd As Long = 0) As String
    o.lStructSize = Len(o)
    o.hWndOwner = pHwnd
    o.hInstance = 0
    o.lpstrFilter = filters(filt)
    o.lpstrFile = Space$(254)
    o.nMaxFile = 255
    o.lpstrFileTitle = Space$(254)
    o.nMaxFileTitle = 255
    o.lpstrInitialDir = initDir
    o.lpstrTitle = title
    o.flags = 0

    OpenDialog = IIf(GetOpenFileName(o), trim$(o.lpstrFile), "")
    OpenDialog = Replace(OpenDialog, Chr(0), Empty)
    If Len(OpenDialog) = 0 And errOnCancel Then Err.Raise 1, "OpenDialog", "Cancel"
    
End Function

Function SaveDialog(filt As FilterTypes, Optional initDir As String, Optional title As String = "", Optional ConfirmOvewrite As Boolean = True, Optional pHwnd As Long = 0, Optional defaultFileName As String) As String
    o.lStructSize = Len(o)
    o.hWndOwner = pHwnd
    o.hInstance = pHwnd
    o.lpstrFilter = filters(filt)
    o.lpstrFile = Space$(254)
    o.nMaxFile = 255
    o.lpstrFileTitle = Space$(254)
    o.nMaxFileTitle = 255
    o.lpstrInitialDir = initDir
    o.lpstrTitle = title
    o.lpstrDefExt = extensions(filt)
    o.flags = 0
    If Len(defaultFileName) > 0 Then
        o.lpstrFile = defaultFileName & Space$(254)
        o.nMaxFile = Len(o.lpstrFile) + 1
    End If

    Dim tmp As String
    tmp = IIf(GetSaveFileName(o), trim$(o.lpstrFile), "")
    If ConfirmOvewrite And tmp <> "" Then
        If FileExists(tmp) Then
            If MsgBox("File Already Exists" & vbCrLf & vbCrLf & "Are you sure you wish to overwrite existing file?", vbYesNo + vbExclamation, "Confirm Overwrite") = vbYes Then SaveDialog = tmp
        Else
            SaveDialog = tmp
        End If
    Else
       SaveDialog = tmp
    End If
    
    SaveDialog = Replace(SaveDialog, Chr(0), Empty)
    
    If Len(SaveDialog) = 0 And errOnCancel Then Err.Raise 1, "SaveDialog", "Cancel"
    
End Function

Function ColorDialog(Optional pHwnd As Long) As Long
    Dim c As oColorDlg
    Dim cColors() As Byte

    c.lStructSize = Len(c)
    c.hWndOwner = pHwnd
    c.hInstance = App.hInstance
    c.lpCustColors = StrConv(cColors, vbUnicode, LANG_US)
    c.flags = 0

    If ChooseColor(c) <> 0 Then
        ColorDialog = c.rgbResult
        cColors = StrConv(c.lpCustColors, vbFromUnicode, LANG_US)
    Else
        ColorDialog = -1
        If errOnCancel Then Err.Raise 1, "ShowColor", "Cancel"
    End If
    
End Function

Function FolderDialog(Optional initDir As String, Optional pHwnd As Long = 0) As String
    Dim bInfo As BrowseInfo, ret As String, ptrList As Long, nullChar As Long

    With bInfo
        .hWndOwner = pHwnd
        .lpszTitle = lstrcat(initDir, "") 'returns memaddress
        .ulFlags = 1                      'only directories
    End With

    ptrList = SHBrowseForFolder(bInfo)
    If ptrList Then
        ret = String$(260, 0)
        SHGetPathFromIDList ptrList, ret 'Get the path from the IDList
        CoTaskMemFree ptrList            'free the block of memory
        nullChar = InStr(ret, vbNullChar)
        If nullChar > 0 Then ret = Left$(ret, nullChar - 1)
    End If

    FolderDialog = Replace(ret, Chr(0), Empty)
    
    If Len(ret) = 0 And errOnCancel Then Err.Raise 1, "ChooseFolder", "Cancel"

End Function

Private Function FileExists(path) As Boolean
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
End Function

Function ChooseFont(Optional obj As Object) As CFont
    Set ChooseFont = frmSelectFont.SelectFont(Me, obj)
End Function


'Function ChooseFont(frm As Form, Optional obj As Object, Optional hWnd As Long) As CFont
'
'   Dim CF As udtFont, hMem As Long, LF As LOGFONT, aFontName As String
'   Dim f As New CFont, tmp() As Byte, lLogFont As Long, flags As Long
'   Dim l As Long
'
'   Set ChooseFont = f
'
'   On Error Resume Next
'   aFontName = obj.Font.name & String(LF_FACESIZE, Chr(0))
'   tmp = StrConv(aFontName, vbFromUnicode, &H409)
'
'   flags = CF_BOTH
'
'   frm.ScaleMode = vbPixels
'   frm.FontName = obj.Font.name
'   frm.FontSize = obj.Font.size
'
'   l = obj.Font.size
'   LF.lfHeight = l
'   Debug.Print l
'
'   LF.lfWeight = obj.Font.Weight
'   LF.lfItalic = obj.Font.Italic * -1
'   LF.lfUnderline = obj.Font.Underline * -1
'   CopyMemory ByVal VarPtr(LF.lfFaceName(0)), ByVal VarPtr(tmp(0)), LF_FACESIZE
'
'   hMem = GlobalAlloc(GPTR, Len(LF))
'
'   If hMem <> 0 Then
'        lLogFont = GlobalLock(hMem)
'        If lLogFont <> 0 Then
'            CopyMemory ByVal lLogFont, LF, Len(LF)
'            flags = flags Or CF_INITTOLOGFONTSTRUCT
'        End If
'   End If
'
'   CF.hInstance = App.hInstance
'   CF.hWndOwner = hWnd
'   CF.lpLogFont = lLogFont
'   CF.lStructSize = Len(CF)
'   CF.flags = flags
'
''  If ofo.UseEffects Then .flags = (.flags Or CF_EFFECTS)
''  If ofo.UseFixedPitchOnly Then .flags = (.flags Or CF_FIXEDPITCHONLY)
''  If ofo.UseNoVectorFonts Then .flags = (.flags Or CF_NOVECTORFONTS)
''  If ofo.UsePrinterFonts Then .flags = (.flags Or CF_PRINTERFONTS)
''  If ofo.UseScreenFonts Then .flags = (.flags Or CF_SCREENFONTS)
''  If ofo.UseTrueTypeOnly Then .flags = (.flags Or CF_TTONLY)
'
'    If SelectFont(CF) Then
'       CopyMemory LF, ByVal hMem, Len(LF)
'       aFontName = String(LF_FACESIZE, Chr(0))
'       CopyMemory ByVal aFontName, LF.lfFaceName(0), LF_FACESIZE
'        With f
'           .name = CString(aFontName)
'           .Bold = LF.lfWeight
'           .Italic = LF.lfItalic
'           .size = CF.iPointSize / 10
'           .Underline = LF.lfUnderline
'           .Charset = LF.lfCharSet
'           .Strikethrough = LF.lfStrikeOut
'        End With
'    End If
'
'   If lLogFont Then GlobalUnlock lLogFont
'   If hMem Then GlobalFree hMem
'
'End Function

Function AvailableFonts() As String()
    Dim x As Integer
    Dim ret() As String
    
    For x = 0 To Screen.FontCount - 1
        push ret, Screen.Fonts(x)
    Next

    AvailableFonts = ret()
    
End Function

Private Function CString(aStr As String) As String
   CString = ""
    Dim k As Long
    k = InStr(aStr, Chr$(0))
    If k Then
       CString = Left$(aStr, k - 1)
    End If
End Function

Private Function MulDiv(In1 As Long, In2 As Long, In3 As Long) As Long
Dim lngTemp As Long
  On Error GoTo MulDiv_err
  If In3 <> 0 Then
    lngTemp = In1 * In2
    lngTemp = lngTemp / In3
  Else
    lngTemp = -1
  End If
MulDiv_end:
  MulDiv = lngTemp
  Exit Function
MulDiv_err:
  lngTemp = -1
  Resume MulDiv_err
End Function

