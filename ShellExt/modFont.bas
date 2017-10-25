Attribute VB_Name = "modFont"
Option Explicit

'ported from
'   http://www.catch22.net/sites/default/files/enumfixedfonts.c
'
'more info:
'    http://www.jasinskionline.com/windowsapi/ref/e/enumfontfamiliesex.html


'Private Const GMEM_MOVEABLE = &H2
'Private Const GMEM_ZEROINIT = &H40
'Private Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)

Private Const LF_FACESIZE = 32

'Private Const FW_BOLD = 700

'Private Const CF_APPLY = &H200&
'Private Const CF_ANSIONLY = &H400&
'Private Const CF_TTONLY = &H40000
'Private Const CF_EFFECTS = &H100&
'Private Const CF_ENABLETEMPLATE = &H10&
'Private Const CF_ENABLETEMPLATEHANDLE = &H20&
'Private Const CF_FIXEDPITCHONLY = &H4000&
'Private Const CF_FORCEFONTEXIST = &H10000
'Private Const CF_INITTOLOGFONTSTRUCT = &H40&
'Private Const CF_LIMITSIZE = &H2000&
'Private Const CF_NOFACESEL = &H80000
'Private Const CF_NOSCRIPTSEL = &H800000
'Private Const CF_NOSTYLESEL = &H100000
'Private Const CF_NOSIZESEL = &H200000
'Private Const CF_NOSIMULATIONS = &H1000&
'Private Const CF_NOVECTORFONTS = &H800&
'Private Const CF_NOVERTFONTS = &H1000000
'Private Const CF_OEMTEXT = 7
'Private Const CF_PRINTERFONTS = &H2
'Private Const CF_SCALABLEONLY = &H20000
'Private Const CF_SCREENFONTS = &H1
'Private Const CF_SCRIPTSONLY = CF_ANSIONLY
'Private Const CF_SELECTSCRIPT = &H400000
'Private Const CF_SHOWHELP = &H4&
'Private Const CF_USESTYLE = &H80&
'Private Const CF_WYSIWYG = &H8000
'Private Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
'Private Const CF_NOOEMFONTS = CF_NOVECTORFONTS

Private Const LOGPIXELSY = 90

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

'Private Type FONTSTRUC
'  lStructSize As Long
'  hWnd As Long
'  hDC As Long
'  lpLogFont As Long
'  iPointSize As Long
'  Flags As Long
'  rgbColors As Long
'  lCustData As Long
'  lpfnHook As Long
'  lpTemplateName As String
'  hInstance As Long
'  lpszStyle As String
'  nFontType As Integer
'  MISSING_ALIGNMENT As Integer
'  nSizeMin As Long
'  nSizeMax As Long
'End Type

Private Type TEXTMETRIC
  tmHeight As Long
  tmAscent As Long
  tmDescent As Long
  tmInternalLeading As Long
  tmExternalLeading As Long
  tmAveCharWidth As Long
  tmMaxCharWidth As Long
  tmWeight As Long
  tmOverhang As Long
  tmDigitizedAspectX As Long
  tmDigitizedAspectY As Long
  tmFirstChar As Byte
  tmLastChar As Byte
  tmDefaultChar As Byte
  tmBreakChar As Byte
  tmItalic As Byte
  tmUnderlined As Byte
  tmStruckOut As Byte
  tmPitchAndFamily As Byte
  tmCharSet As Byte
End Type

Private Type POINTAPI
    x As Long
    Y As Long
End Type

Const EM_CHARFROMPOS& = &HD7

'Private Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As FONTSTRUC) As Long
'Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
'Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function EnumFontFamiliesEx Lib "gdi32" Alias "EnumFontFamiliesExA" (ByVal hDC As Long, lpLogFont As LOGFONT, ByVal lpEnumFontProc As Long, ByVal lParam As Long, ByVal dw As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Private Const ANSI_CHARSET = 0
Private Const DEFAULT_PITCH = 0
Private Const VARIABLE_PITCH = 2
Private Const FIXED_PITCH = 1
Private Const FF_DONTCARE = 0
Private Const TRUETYPE_FONTTYPE = &H4
Private Const DEFAULT_CHARSET = 1
Private Const HWND_TOPMOST = -1

Private sizes As Collection
Private hDC As Long

Function ColorText(ByRef rtf As RichTextBox, Optional ByRef find As String, Optional ByVal Color) As Long

    Dim i As Long
    Dim l As Long
    Dim start As Long
    Dim cnt As Long
    
    start = rtf.SelStart
    LockWindowUpdate rtf.hwnd
    
    rtf.SelStart = 0
    rtf.SelLength = Len(rtf.text)
    rtf.SelColor = vbBlack
    rtf.SelBold = False
    rtf.SelFontSize = rtf.Font.size
        
    If Len(find) > 0 Then
        l = Len(find)
        i = rtf.find(find)
        Do While i >= 0
            rtf.SelStart = i
            rtf.SelLength = l
            rtf.SelColor = Color
            rtf.SelBold = True
            rtf.SelFontSize = rtf.Font.size + 4
            i = rtf.find(find, i + l)
            cnt = cnt + 1
        Loop
    End If
    
    rtf.SelStart = start
    rtf.SelLength = 0
    LockWindowUpdate 0
    ColorText = cnt
    
End Function

Sub SetWindowTopMost(f As Form, Optional top As Boolean = True)
   SetWindowPos f.hwnd, IIf(top, HWND_TOPMOST, 1), f.Left / 15, f.top / 15, f.Width / 15, f.Height / 15, Empty
End Sub


Function EnumFontSizes(fontname As String) As Collection
    Dim lf As LOGFONT
    Dim b() As Byte
   
    Set sizes = New Collection
    Set EnumFontSizes = sizes
    
    lf.lfHeight = 0
    lf.lfCharSet = DEFAULT_CHARSET
    lf.lfPitchAndFamily = DEFAULT_PITCH Or FF_DONTCARE
    b() = StrConv(fontname & Chr(0), vbFromUnicode)
    CopyMemory ByVal VarPtr(lf.lfFaceName(0)), ByVal VarPtr(b(0)), UBound(b) + 1
    hDC = GetDC(0)
    
    EnumFontFamiliesEx hDC, lf, AddressOf EnumFontCallBack, 0, 0
    ReleaseDC 0, hDC
    
End Function

Public Function EnumFontCallBack(ByVal plf As Long, ByVal ptm As Long, ByVal fontType As Long, ByVal lParam As Long) As Long
    Dim x, trueTypeSizes, pointsize As Long, logSize As Long
    Dim tm As TEXTMETRIC
    
    On Error Resume Next
    
    If fontType <> TRUETYPE_FONTTYPE Then
        CopyMemory ByVal VarPtr(tm), ByVal ptm, Len(tm)
        logSize = tm.tmHeight - tm.tmInternalLeading
        pointsize = MulDiv(logSize, 72, GetDeviceCaps(hDC, LOGPIXELSY))
        sizes.Add pointsize, "sz:" & pointsize 'only adds unique sizes resume next required
        EnumFontCallBack = 1
    Else
       trueTypeSizes = Array(8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72)
       For Each x In trueTypeSizes
            sizes.Add x
       Next
       EnumFontCallBack = 0
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
  
'MulDiv_end:
  MulDiv = lngTemp
  Exit Function
  
MulDiv_err:
  lngTemp = -1
  Resume MulDiv_err
End Function

Function WordUnderCursor(MyRTB As RichTextBox, x As Single, Y As Single, startPos As Long) As String
    Dim MyPoint As POINTAPI
    Dim MyPos As Integer
    Dim MyStartPos As Integer
    Dim MyEndPos As Integer
    Dim MyCharacter As Integer
    Dim TextLength As Integer
    Dim MyText As String
    
    On Error Resume Next
    MyPoint.x = x \ Screen.TwipsPerPixelX
    MyPoint.Y = Y \ Screen.TwipsPerPixelY
    MyPos = SendMessage(MyRTB.hwnd, EM_CHARFROMPOS, 0&, MyPoint)
    
    If MyPos <= 0 Then Exit Function
    
    MyText = MyRTB.text
    
    For MyStartPos = MyPos To 1 Step -1
        MyCharacter = Asc(Mid$(MyRTB.text, MyStartPos, 1))
        If Not isAlpha(MyCharacter) Then Exit For
           
    Next
    
    MyStartPos = MyStartPos + 1
    TextLength = Len(MyText)
    
    For MyEndPos = MyPos To TextLength
        MyCharacter = Asc(Mid$(MyText, MyEndPos, 1))
        If Not isAlpha(MyCharacter) Then Exit For
    Next
    
    MyEndPos = MyEndPos - 1
    If MyStartPos <= MyEndPos Then
        startPos = MyStartPos - 1
        WordUnderCursor = Mid$(MyText, MyStartPos, MyEndPos - MyStartPos + 1)
    End If
        
End Function

Function isAlpha(c As Integer) As Boolean

    If c >= Asc("a") And c <= Asc("z") Then
        isAlpha = True
        Exit Function
    End If
    
    If c >= Asc("A") And c <= Asc("Z") Then
        isAlpha = True
        Exit Function
    End If
    
    If c >= Asc("0") And c <= Asc("9") Then
        isAlpha = True
        Exit Function
    End If
    
    If c = Asc(":") Or c = Asc("_") Then
        isAlpha = True
        Exit Function
    End If
    
    
    
End Function

