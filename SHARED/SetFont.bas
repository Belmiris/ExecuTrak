Attribute VB_Name = "Module1"
Option Explicit
    'Constants for API call
    Private Const LF_FACESIZE = 32
    Private Const LOGPIXELSX = 88    '  Logical pixels/inch in X
    Private Const LOGPIXELSY = 90    '  Logical pixels/inch in Y
    Private Const ASPECTX = 40       '  Length of the X leg
    Private Const ASPECTXY = 44      '  Length of the hypotenuse
    Private Const ASPECTY = 42       '  Length of the Y leg
    Private Const ORIENT_NORMAL = 1
    Private Const ORIENT_RIGHT = 2
    Private Const ORIENT_LEFT = 3
    Private Const ORIENT_DOWN = 4
    
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
    Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
    Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
    Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
    Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
    Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long

    Public Type tpFont
        Name As String
        Size As Integer
        Bold As Boolean
    End Type

    
    Private fntLastFontUsed As tpFont
    Private hCurrFont As Long
    Private hPrevFont As Long


Private Sub subSetFont(objOutputTo As Object, _
                       sFontName As String, _
                       ByVal nFontSize As Integer, _
                       ByVal bFontBold As Boolean)
    Dim font As LOGFONT
    
    With fntLastFontUsed
        If sFontName = .Name Then
            If nFontSize = .Size Then
                If bFontBold = .Bold Then
                    Exit Sub
                End If
            End If
        End If
        .Name = sFontName
        .Size = nFontSize
        .Bold = bFontBold
    End With
    'font.lfEscapement = nRotate * 10   ' 180-degree rotation
    #If Win32 Then
        Dim i As Integer
        For i = 1 To Len(sFontName)
            font.lfFaceName(i - 1) = Asc(Mid(sFontName, i, 1))
        Next
        font.lfFaceName(i) = 0
    #Else
        font.lfFaceName = sFontName & Chr$(0) 'Null character at end
    #End If
    font.lfHeight = -nFontSize * GetDeviceCaps(objOutputTo.hdc, LOGPIXELSY) / 72 ' one inch contains 72 points.
    If bFontBold Then
        font.lfWeight = 700
    Else
        font.lfWeight = 400
    End If
    
    If hCurrFont <> 0 Then
        DeleteObject hCurrFont
    End If
    hCurrFont = CreateFontIndirect(font)

    If hPrevFont = 0 Then
        hPrevFont = SelectObject(objOutputTo.hdc, hCurrFont)
    Else
        SelectObject objOutputTo.hdc, hCurrFont
    End If
    
End Sub


Private Sub subRestoreFont(objOutputTo As Object)
    SelectObject objOutputTo.hdc, hPrevFont
    DeleteObject hCurrFont
End Sub


