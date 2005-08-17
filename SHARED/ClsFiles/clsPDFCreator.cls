VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPDFCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================================
' Descrizione.....: Classe per la generazione di file PDF 1.3
' Nome dei File...: clsPDFCreator.cls
' Data............: 13/05/2005
' Versione........: 1.2.3 (LIGHT)
' Sistema.........: Visual Basic 6.0 Pro - SP 5
' Testato su......: Windows 2000 Professional - SP 4.
' Scritto da......: Luigi Micco ® eLLeMMe
' E-Mail..........: l(dot)micco(at)tiscali(dot)it
' Web site........: http://xoomer.virgilio.it/ellemme
'================================================================
' (C) 2004  - Questo software è di libero utilizzo per scopi  NON
'             COMMERCIALI.  Qualsiasi altro  uso, senza  espressa
'             autorizzazione dell'autore non  è consentito. Se il
'             software viene  ridistributo,  non eliminare questa
'             nota.
'================================================================

' Attributi e Variabili pubbliche della classe
' --------------------------------------------
' Margin
' PaperWidth
' PaperHeight
' Pages
' Scalemode
' PaperSize
' Orientation
' Title

' Procedure e Funzioni Pubbliche della classe
' -------------------------------------------
' InitPDFFile(Optional strFileName As String)
' ClosePDFFile()
' StartPage() As Integer
' EndPage()
' SetWordSpacing(W As Single)
' SetCharSpacing(W As Single)
' SetTextHorizontalScaling(W As Single)
' SetTextRenderingMode(W As Single)
' DrawText(x As Single, y As Single, strTemp As String,
'          FontName As String, Fontsize As Single, _
'          Optional Align As pdfTextAlign = pdfAlignLeft, _
'          Optional Rotate As Single = 0)
' SetColorStroke(rgb As Long)
' SetColorFill(rgb As Long)
' SetDash(dash_on As Single, Optional dash_off As Single)
' SetLineWidth(W As Single)
' SetLineCap(W As Integer)
' SetLineJoin(W As Integer)
' SetMiterLimit(W As Single)
' MoveTo(x As Single, y As Single)
' LineTo(x As Single, y As Single, _
'        Optional Options As pdfPathOptions = pdfPathOptions.Stroked)
' Path(Optional Options As pdfPathOptions = pdfPathOptions.Nil)
' Rectangle(x As Single, y As Single, _
'           xdim As Single, ydim As Single, _
'           Optional Options As pdfPathOptions = pdfPathOptions.Stroked, _
'           Optional Ray As Single)
' Curve(X1 As Single, Y1 As Single, _
'       X2 As Single, Y2 As Single, _
'       X3 As Single, X3 As Single, _
'       Optional Options As pdfPathOptions = pdfPathOptions.Stroked)
' DrawCircle(x As Single, y As Single, _
'            Ray As Single, _
'            Optional Options As pdfPathOptions = pdfPathOptions.Stroked)
' Arc(x As Single, y As Single, _
'    Ray As Single, _
'    Optional StartAngle As Single = 0, _
'    Optional EndAngle As Single = 360, _
'    Optional Ratio As Single = 1, _
'    Optional Pie As Boolean = False, _
'    Optional Rotate As Single = 0, _
'    Optional Quality As Integer = 1, _
'    Optional Options As pdfPathOptions = pdfPathOptions.Stroked)
' LoadImgFromBMPFile(Name As String, _
'                   FileName As String, _
'                   Optional ColorSpace As pdfColorSpace = pdfRGB) As Boolean
' LoadImgFromBuffer(Name As String, _
'                  ByRef ImgBuf() As Byte, _
'                  ByRef ImgColor As String, _
'                  ImgWidth As Long, _
'                  ImgHeight As Long, _
'                  ImgBPP As Byte, _
'                  Optional ColorSpace As pdfColorSpace = pdfRGB)
' DrawImg(Name As String, DestX As Single, DestY As Single, _
'        ImgWidth As Single, ImgHeight As Single)
' LoadFontStandard(Name As String, BaseFont As String, _
'                 Optional Options As pdfFontStyle = pdfNormal)
' LoadFont(Name As String, BaseFont As String, _
'         Optional Options As pdfFontStyle = pdfNormal)
' StartObject(Name As String, Optional Options As pdfObjectType = pdfNull)
' EndObject()
' DrawObject(Name As String)
' Length(Phrase As String, FontName As String, Fontsize As Single) As Single
' LeggeBMP(FileName As String, _
'         ByRef ImgBuf() As Byte, _
'         ByRef ImgColor As String, _
'         ByRef ImgWidth As Long, _
'         ByRef ImgHeight As Long, _
'         ByRef ImgBPP As Byte) As Boolean

Option Explicit
Option Base 1

' BaseFont Type 1
    'CourierNew ,Italic ,Bold ,BoldItalic
    'Arial ,Italic ,Bold ,BoldItalic
    'TimesNewRoman ,Italic ,Bold ,BoldItalic
    'ZapfDingbats
    'Symbol

' BaseFont TrueType
    'CourierNew ,Italic ,Bold ,BoldItalic
    'Arial ,Italic ,Bold ,BoldItalic
    'TimesNewRoman ,Italic ,Bold ,BoldItalic
    'Symbol

Private Const PDF_VERSION = "%PDF-1.3"
Private Const AUTOR = "eLLeMMe"
Private Const CREATOR = "clsPDFCreator LIGHT Version"
Private Const COPYRIGHT = "© 2003-2005 eLLeMMe"

' Enum
Public Enum pdfPathOptions
 Stroked = &H1
 Closed = &H2
 Filled = &H4
 Nil = &H0
End Enum

Public Enum pdfColorSpace
 pdfRGB = 0
 pdfGrayScale = 1
End Enum

Public Enum pdfFontStyle
 pdfNormal = 1
 pdfBold = 2
 pdfItalic = 3
 pdfBoldItalic = 4
End Enum

Public Enum pdfTextAlign
 pdfAlignLeft = 1
 pdfAlignRight = 2
 pdfCenter = 3
End Enum

Public Enum pdfPaperSize
  pdfA4 = 0
  pdfA3 = 1
  pdf8_5x11 = 2
  pdf8_5x12 = 3
  pdfUser = &HFF
End Enum

Public Enum pdfScaleMode
  pdf72PxInch = 1
  pdfInch = 2
  pdfMillimeter = 3
  pdfCentimeter = 4
End Enum

Public Enum pdfPrintOrientation
  pdfPortrait = 1
  pdfLandscape = 2
End Enum

Public Enum pdfObjectType
  pdfNull = &H0
  pdfFirstPage = &H1
  pdfEvenPages = &H2
  pdfOddPages = &H4
  pdfNotFirstPage = &H8
  pdfAllPages = pdfEvenPages Or pdfOddPages Or pdfFirstPage
End Enum


'Internal Variables
Private mvarFileNumber As Integer
Private mvarTitle As String
Private mvarPaperSize As pdfPaperSize
Private mvarPageWidth As Single
Private mvarPageHeight As Single
Private mvarScaleMode As pdfScaleMode
Private mvarOrientation As pdfPrintOrientation
Private mvarPages As Integer
Private mvarMargin As Single
Private mvarFileName As String 'Copia locale.
Private mvarEncodeASCII85 As Boolean
Private strPage As String
Private strFont As String
Private strFontname As String
Private strImg As String
Private strObject As String
Private strObjectForPage As String
Private intResource As Integer
Private intPages As Integer
Private intObject As Integer
Private intMaxObject As Long
Private intOffsetTemp As Long
Private intUsedFont As Integer

Private mvarCharSpacing As Single
Private mvarWordSpacing As Single
Private mvarTextScaling As Single

Private Type FontDescriptor
  BaseFont As String
  FirstChar As Byte
  LastChar As Byte
  Param As String
  Widths(1 To 255) As Long
  MissingWidth As Long
End Type

Private Type ObjDescriptor
  Name As String
  Options As pdfObjectType
End Type

'Data Arrays
Private arrXREF() As String
Private arrOBJECT() As ObjDescriptor
Private arrFONT() As FontDescriptor

Private m_strSourceFile As String
Private m_strTargetFile As String
Private m_strFontname As String
Private m_intRPP As Long
Private m_intFontSize As Single
Private m_PageBreakBefore As String
'

Private Sub Class_Initialize()
  Dim xTemp As pdfFontStyle
  Dim i As Integer
  
  mvarFileNumber = 0
  mvarEncodeASCII85 = False
  Title = ""
  ScaleMode = pdf72PxInch
  PaperSize = pdfA4
  Orientation = pdfPortrait
  Margin = 0
  mvarPages = 0
  intUsedFont = 0
 
m_strSourceFile = ""
m_strTargetFile = ""
m_strFontname = "Courier New"
m_intRPP = 66
m_intFontSize = 6.5
m_PageBreakBefore = ""
 
End Sub

'
Private Sub Class_Terminate()
'  TempDelete
End Sub

'
Public Property Let EncodeASCII85(ByVal vData As Boolean)
  mvarEncodeASCII85 = vData
End Property

'
Public Property Let Margin(ByVal vData As Single)
  mvarMargin = ToSpace(vData)
End Property

'
Public Property Let PaperWidth(ByVal vData As Single)
  mvarPageWidth = ToSpace(vData)
End Property

'
Public Property Let PaperHeight(ByVal vData As Single)
  mvarPageHeight = ToSpace(vData)
End Property

'
Public Property Get Pages() As Integer
  Pages = mvarPages
End Property

'
Public Property Let ScaleMode(ByVal New_ScaleMode As pdfScaleMode)
  mvarScaleMode = New_ScaleMode
End Property

'
Public Property Let PaperSize(New_PaperSize As pdfPaperSize)

  mvarPaperSize = New_PaperSize
  Select Case mvarPaperSize
    Case pdfPaperSize.pdf8_5x12     'Letter, 8 1/2 x 12 in.
        mvarPageWidth = 612
        mvarPageHeight = 864
    
    Case pdfPaperSize.pdf8_5x11     'Letter, 8 1/2 x 11 in.
        mvarPageWidth = 612
        mvarPageHeight = 792
    
    Case pdfPaperSize.pdfA3 'A3, 297 x 420 mm
        mvarPageWidth = 842
        mvarPageHeight = 1190.5
    
    Case pdfPaperSize.pdfA4 'A4, 210 x 297 mm
        mvarPageWidth = 595.2
        mvarPageHeight = 842
    
    Case pdfPaperSize.pdfUser 'User-defined
  
  End Select
    
End Property

'
Public Property Let Orientation(NewOrientation As pdfPrintOrientation)
  mvarOrientation = NewOrientation
End Property

'
Public Property Let Title(New_Title As String)
  mvarTitle = New_Title
End Property

'
Public Sub InitPDFFile(Optional strFileName As String)

  If mvarFileNumber <> 0 Then Close #mvarFileNumber
  
  If strFileName = "" Then strFileName = App.Path & "\~" & Hex(Val(Format(Now(), "dhhmmss") & Trim(Str(Int(Rnd() * 10))))) & ".pdf"
  
  mvarFileNumber = FreeFile
  mvarFileName = strFileName
  Open mvarFileName For Output As #mvarFileNumber
  
  WriteObj PDF_VERSION
   
  ' inizializza gli oggetti
  intMaxObject = 0
    
    ' 1 Attribute
  InitObj (1)
  WriteObj "<<" & vbCr & _
           "/Title (" & ToPdfStr(mvarTitle) & ")" & vbCr & _
           "/Author (" & ToPdfStr(AUTOR) & ")" & vbCr & _
           "/Creator  (" & ToPdfStr(CREATOR) & ")" & vbCr & _
           "/Producer (" & ToPdfStr(COPYRIGHT) & ")" & vbCr & _
           "/CreationDate (D:" & Format(Now(), "yyyymmddhhmmss") & "+01'00')" & vbCr & _
           ">>"
  EndObj
  
  ' 2 Catalog
  InitObj (2)
  WriteObj "<<" & vbCr & _
           "/Type /Catalog" & vbCr & _
           "/Pages 3 0 R" & vbCr & _
           "/PageLayout /OneColumn" & vbCr & _
           "/PageMode /UseNone" & vbCr & _
           "/Lang (it) " & vbCr & _
           "/ViewerPreferences << /HideToolbar false " & _
                                 "/DisplayDocTitle true " & _
                                 "/HideWindowUI false >>" & vbCr & _
           ">>"
  EndObj
  
  ' Lascia spazio per gli oggetti 3, 4
  intMaxObject = 4

End Sub

Public Sub ClosePDFFile()
  Dim i As Integer
  Dim intTemp As Long

  ' 4 Resource
  intResource = InitObj(4)
  WriteObj "<<" & vbCr & _
           IIf(strFont <> "", "/Font <<" & vbCr & strFont & ">>" & vbCr, "") & _
           "/ProcSet [/PDF /Text" & IIf(strImg <> "", " /ImageB /ImageC /ImageI", "") & " ]" & vbCr & _
           IIf((strImg <> "") Or (strObject <> ""), "/XObject <<" & vbCr & strImg & vbCr & strObject & ">>" & vbCr, "") & _
           ">>"
  EndObj
  
  ' 3 Page
  intPages = InitObj(3)
  WriteObj "<<" & vbCr & _
           "/Type /Pages" & vbCr & _
           "/Count " & mvarPages & vbCr & _
           "/MediaBox [0 0 " & _
                      ToStr(mvarPageWidth) & " " & _
                      ToStr(mvarPageHeight) & "]" & vbCr & _
           "/CropBox [" & ToStr(mvarMargin) & " " & _
                     ToStr(mvarMargin) & " " & _
                     ToStr(mvarPageWidth - mvarMargin) & " " & _
                     ToStr(mvarPageHeight - mvarMargin) & "]" & vbCr & _
           IIf(mvarOrientation = pdfLandscape, "/Rotate 90", "") & vbCr & _
           "/Kids [" & strPage & " ]" & vbCr & _
           "/Resources " & CStr(intResource) & " 0 R" & vbCr & _
           ">>"
  EndObj
  
  intTemp = Seek(mvarFileNumber) - 1
  WriteObj "xref"
  WriteObj "0 " & CStr(intMaxObject + 1)
  WriteObj "0000000000 65535 f"
  For i = 1 To intMaxObject
    WriteObj arrXREF(i)
  Next
  
  WriteObj "trailer" & vbCr & _
           "<<" & vbCr & _
           "/Size " & intMaxObject & vbCr & _
           "/Info 1 0 R" & vbCr & _
           "/Root 2 0 R" & vbCr & _
           ">>"
  
  WriteObj "startxref" & vbCr & _
           CStr(intTemp)
           
  Print #mvarFileNumber, "%%EOF" & vbCr ';
  Close #mvarFileNumber
End Sub

Private Function InitObj(Optional intObj As Integer) As Integer
  
  If intObj = 0 Then intObj = intMaxObject + 1
  If intObj > intMaxObject Then intMaxObject = intObj
  ReDim Preserve arrXREF(1 To intMaxObject)
  
  arrXREF(intObj) = Right$("0000000000" & (Seek(mvarFileNumber) - 1), 10) & " 00000 n"
  WriteObj CStr(intObj) & " 0 obj"
  InitObj = intObj
End Function

Private Sub EndObj()
  Print #mvarFileNumber, "endobj"
End Sub

Private Sub WriteObj(strTemp As String)
  Print #mvarFileNumber, strTemp
End Sub

Public Function StartPage() As Integer
  Dim intPage As Integer
  
  mvarPages = mvarPages + 1
  intPage = InitObj
  WriteObj "<<" & vbCr & _
           "/Type /Page" & vbCr & _
           "/Parent 3 0 R" & vbCr & _
           "/Contents " & CStr(intPage + 1) & " 0 R" & vbCr & _
           ">>"
  EndObj
  
  strPage = strPage & CStr(intPage) & " 0 R "
     
  InitObj (intPage + 1)
  WriteObj "<< /Length " & CStr(intPage + 2) & " 0 R >>" & vbCr & _
           "stream"
  intOffsetTemp = Seek(mvarFileNumber)
 
  ' Verifica se ci sono oggetti comuni da inserire sullo sfondo della pagina
  InsertObjectOnPage

  mvarCharSpacing = 0
  mvarWordSpacing = 0
  mvarTextScaling = 100
  StartPage = mvarPages
End Function

'
Public Sub EndPage()
  
  intOffsetTemp = Seek(mvarFileNumber) - intOffsetTemp
  WriteObj "endstream"
  EndObj
  
  ' Scrive la lunghezza
  InitObj
  WriteObj CStr(intOffsetTemp)
  EndObj
End Sub

'
Public Sub SetWordSpacing(W As Single)
  mvarWordSpacing = W
  WriteObj ToStr(mvarWordSpacing) & " Tw"
End Sub

'
Public Sub SetCharSpacing(W As Single)
  mvarCharSpacing = W
  WriteObj ToStr(mvarCharSpacing) & " Tc"
End Sub

'
Public Sub SetTextHorizontalScaling(W As Single)
  mvarTextScaling = W
  WriteObj ToStr(mvarTextScaling) & " Tz"
End Sub

'
Public Sub SetTextRenderingMode(W As Single)
  If (W >= 0) And (W <= 2) Then WriteObj ToStr(W) & " Tr"
End Sub

'
Public Sub DrawText(x As Single, y As Single, _
                    strTemp As String, _
                    FontName As String, FontSize As Single, _
                    Optional Align As pdfTextAlign = pdfAlignLeft, _
                    Optional Rotate As Single = 0)
  Dim rad As Single
  Dim PI As Single
  Dim sTeta As Single
  Dim cTeta As Single
  Dim C As Single
  Dim l As Single
  
  Select Case Align
    Case pdfAlignLeft
    
    Case pdfAlignRight
      l = length(strTemp, FontName, FontSize)
      x = x - l
    Case pdfCenter
      l = length(strTemp, FontName, FontSize)
      x = x - l / 2
  End Select
  
  WriteObj "BT"
  WriteObj "/" & FontName & " " & ToStr(FontSize) & " Tf"
  If Rotate <> 0 Then
    PI = 3.141592
    C = PI / 180
    sTeta = sIn(C * Rotate)
    cTeta = Cos(C * Rotate)
    WriteObj ToStr(cTeta) & " " & ToStr(sTeta) & " " & _
             ToStr(-sTeta) & " " & ToStr(cTeta) & _
             " " & ToStr(ToSpace(x)) & " " & ToStr(ToSpace(y)) & " Tm"
  Else
    WriteObj ToStr(ToSpace(x)) & " " & ToStr(ToSpace(y)) & " Td"
  End If
  
  WriteObj "(" & ToPdfStr(strTemp) & ") Tj"
  WriteObj "ET"
End Sub

'
Public Sub SetColorStroke(rgb As Long)
  Dim r As Byte
  Dim G As Byte
  Dim b As Byte
  
  If (rgb <= 0) And (rgb >= -255) Then
    WriteObj ToStr(-rgb / 255) & " G"
  Else
    r = (rgb Mod 256)
    G = ((rgb \ 256) Mod 256)
    b = ((rgb \ 65536) Mod 256)
    WriteObj ToStr(r / 255) & " " & ToStr(G / 255) & " " & ToStr(b / 255) & " RG"
  End If
End Sub

'
Public Sub SetColorFill(rgb As Long)
  Dim r As Byte
  Dim G As Byte
  Dim b As Byte
  
  If (rgb <= 0) And (rgb >= -255) Then
    WriteObj ToStr(-rgb / 255) & " g"
  Else
    r = (rgb Mod 256)
    G = ((rgb \ 256) Mod 256)
    b = ((rgb \ 65536) Mod 256)
    WriteObj ToStr(r / 255) & " " & ToStr(G / 255) & " " & ToStr(b / 255) & " rg"
  End If
End Sub

'
Public Sub SetDash(dash_on As Single, Optional dash_off As Single)
  If (dash_on = 0) And (dash_off = 0) Then
    WriteObj "[ ] 0 d"
  Else
    WriteObj "[" & ToStr(ToSpace(dash_on)) & " " & ToStr(ToSpace(dash_off)) & "] 0 d"
  End If
End Sub

'
Public Sub SetLineWidth(W As Single)
  WriteObj ToStr(ToSpace(W)) & " w"
End Sub

'
Public Sub SetLineCap(W As Integer)
  If (W >= 0) And (W <= 2) Then WriteObj ToStr(W) & " J"
End Sub

'
Public Sub SetLineJoin(W As Integer)
  If (W >= 0) And (W <= 2) Then WriteObj ToStr(W) & " j"
End Sub

'
Public Sub SetMiterLimit(W As Single)
  If (W >= 1) Then WriteObj ToStr(W) & " M"
End Sub

'
Public Sub MoveTo(x As Single, y As Single)
  WriteObj ToStr(ToSpace(x)) & " " & ToStr(ToSpace(y)) & " m"
End Sub

'
Public Sub LineTo(x As Single, y As Single, _
                  Optional Options As pdfPathOptions = pdfPathOptions.Stroked)
  WriteObj ToStr(ToSpace(x)) & " " & ToStr(ToSpace(y)) & " l"
  Call Path(Options)
End Sub

Private Sub Path(Optional Options As pdfPathOptions = pdfPathOptions.Nil)
  
  If (Options And pdfPathOptions.Filled) <> 0 Then
    WriteObj "f"
  Else
    If (Options And pdfPathOptions.Closed) <> 0 Then WriteObj "h"
    If (Options And pdfPathOptions.Stroked) <> 0 Then WriteObj "S"
  End If

End Sub

'
Public Sub Rectangle(x As Single, y As Single, _
                     xdim As Single, ydim As Single, _
                     Optional Options As pdfPathOptions = pdfPathOptions.Stroked, _
                     Optional Ray As Single)
  Dim sR As Single
                     
  WriteObj "n"
  If Ray > 0 Then
    If Ray > (xdim / 2) Then Ray = xdim / 2
    If Ray > (ydim / 2) Then Ray = ydim / 2
    sR = 0.55 * Ray
    MoveTo x + Ray, y
    LineTo x + xdim - Ray, y, Nil
    Curve (x + xdim - Ray + sR), y, x + xdim, y + Ray - sR, x + xdim, y + Ray, Nil
    LineTo x + xdim, y + ydim - Ray, Nil
    Curve x + xdim, y + ydim - Ray + sR, x + xdim - Ray + sR, y + ydim, x + xdim - Ray, y + ydim, Nil
    LineTo x + Ray, y + ydim, Nil
    Curve x + Ray - sR, y + ydim, x, y + ydim - Ray + sR, x, y + ydim - Ray, Nil
    LineTo x, y + Ray, Nil
    Curve x, y + Ray - sR, x + Ray - sR, y, x + Ray, y, Nil
  Else
    WriteObj ToStr(ToSpace(x)) & " " & ToStr(ToSpace(y)) & " " & _
             ToStr(ToSpace(xdim)) & " " & ToStr(ToSpace(ydim)) & " re"
  End If
  
  Call Path(Options)
End Sub

'
Public Sub Curve(X1 As Single, Y1 As Single, _
                 X2 As Single, Y2 As Single, _
                 X3 As Single, Y3 As Single, _
                 Optional Options As pdfPathOptions = pdfPathOptions.Stroked)

  WriteObj ToStr(ToSpace(X1)) & " " & ToStr(ToSpace(Y1)) & " " & _
           ToStr(ToSpace(X2)) & " " & ToStr(ToSpace(Y2)) & " " & _
           ToStr(ToSpace(X3)) & " " & ToStr(ToSpace(Y3)) & " c"
  Call Path(Options)
End Sub

'
Public Sub DrawCircle(x As Single, y As Single, _
                      Ray As Single, _
                      Optional Options As pdfPathOptions = pdfPathOptions.Stroked)
  
  MoveTo x, y - Ray
  Curve x + 0.55 * Ray, y - Ray, x + Ray, y - 0.55 * Ray, x + Ray, y, Nil
  Curve x + Ray, y + 0.55 * Ray, x + 0.55 * Ray, y + Ray, x, y + Ray, Nil
  Curve x - 0.55 * Ray, y + Ray, x - Ray, y + 0.55 * Ray, x - Ray, y, Nil
  Curve x - Ray, y - 0.55 * Ray, x - 0.55 * Ray, y - Ray, x, y - Ray, Nil
  Call Path(Options)
End Sub

'
Public Sub Arc(x As Single, y As Single, _
               Ray As Single, _
               Optional StartAngle As Single = 0, _
               Optional EndAngle As Single = 360, _
               Optional Ratio As Single = 1, _
               Optional Pie As Boolean = False, _
               Optional Rotate As Single = 0, _
               Optional Quality As Integer = 1, _
               Optional Options As pdfPathOptions = pdfPathOptions.Stroked)
  
  Dim i As Single
  Dim rad As Single
  Dim PI As Single
  Dim sTeta As Single
  Dim cTeta As Single
  Dim C As Single
  Dim X1, X2 As Single
  Dim Y1, Y2 As Single
  
  WriteObj "n"
  MoveTo x, y
  If (Options And pdfPathOptions.Filled) <> 0 Then Pie = True
  PI = 3.141592
  C = PI / 180
  sTeta = sIn(-C * Rotate)
  cTeta = Cos(-C * Rotate)
  
  For i = StartAngle To EndAngle Step Quality
    rad = C * i
    X2 = Ray * Cos(rad)
    Y2 = (Ray * Ratio) * sIn(rad)
    
    X1 = X2 * cTeta + Y2 * sTeta
    Y1 = -X2 * sTeta + Y2 * cTeta
    If (i = StartAngle) And (Not Pie) Then MoveTo x + X1, y + Y1
    LineTo x + X1, y + Y1, Nil
  Next
    
  If Pie Then LineTo x, y, Nil
  Call Path(Options)
End Sub

'
Public Function LoadImgFromBMPFile(Name As String, _
                                   FileName As String, _
                                   Optional ColorSpace As pdfColorSpace = pdfRGB) As Boolean
  Dim ImgBuf() As Byte
  Dim ImgColor() As Byte
  Dim blnFlag As Boolean
  Dim ImgWidth As Long
  Dim ImgHeight As Long
  Dim ImgBPP As Byte
  
  blnFlag = LeggeBMP(FileName, ImgBuf(), ImgColor(), ImgWidth, ImgHeight, ImgBPP, ColorSpace)
  If blnFlag Then LoadImgFromBuffer Name, ImgBuf(), ImgColor(), ImgWidth, ImgHeight, ImgBPP, ColorSpace
  LoadImgFromBMPFile = blnFlag
End Function


'
Public Sub LoadImgFromBuffer(Name As String, _
                             ByRef ImgBuf() As Byte, _
                             ByRef ImgColor() As Byte, _
                             ImgWidth As Long, _
                             ImgHeight As Long, _
                             ImgBPP As Byte, _
                             Optional ColorSpace As pdfColorSpace = pdfRGB)
                   
  Dim BitPerPixel As Byte
'
  Dim sPixel As String
  Dim sColor As String
  Dim lngGray As Long
  Dim i As Integer
  Dim y As Long
  
  Dim strDevice As String
  Dim BPP As Byte
  
  BPP = 8 / ImgBPP
  
  i = InitObj
  WriteObj "<<" & vbCr & _
           "/Type /XObject" & vbCr & _
           "/Subtype /Image" & vbCr & _
           "/Name /" & Name & vbCr & _
           "/Width " & CStr(ImgWidth) & vbCr & _
           "/Height " & CStr(ImgHeight) & _
           IIf(mvarEncodeASCII85, vbCr & "/Filter /ASCII85Decode", "")

  strDevice = IIf(ColorSpace = pdfRGB, "DeviceRGB", "DeviceGray")
  
  Select Case ImgBPP
    Case 24
      WriteObj "/BitsPerComponent 8" & vbCr & _
               "/ColorSpace /" & strDevice & vbCr & _
               "/Length " & CStr(i + 1) & " 0 R"
    
    Case 8, 4, 1
      WriteObj "/BitsPerComponent " & CStr(ImgBPP) & vbCr & _
               "/ColorSpace [/Indexed /" & strDevice & " " & _
               CStr((2 ^ ImgBPP) - 1) & " " & CStr(i + 2) & " 0 R]" & vbCr & _
               "/Length " & CStr(i + 1) & " 0 R"
    
  End Select
  
  ' Now I use StrConv: thanks to Adriano Belfiore <for_program(at)yahoo(dot)it>
  sPixel = IIf(mvarEncodeASCII85, ToASCII85(ImgBuf), StrConv(ImgBuf, vbUnicode))
  
  WriteObj ">>" & vbCr & _
           "stream" & vbCr & _
           sPixel & vbCr & _
           "endstream"
  EndObj
  
  strImg = strImg & "/" & Name & " " & CStr(i) & " 0 R " & vbCr

  InitObj
  WriteObj CStr(Len(sPixel) - IIf(mvarEncodeASCII85, 2, 0))
  EndObj

  If ImgBPP <= 8 Then
    InitObj
    
    sColor = IIf(mvarEncodeASCII85, ToASCII85(ImgColor), StrConv(ImgColor, vbUnicode))
    
    WriteObj "<<" & _
             IIf(mvarEncodeASCII85, "/Filter /ASCII85Decode" & vbCr, "") & _
             "/Length " & CStr(Len(sColor) - IIf(mvarEncodeASCII85, 2, 0)) & " >>" & vbCr & _
             "stream" & vbCr & _
             sColor & vbCr & _
             "endstream"
    EndObj
  End If


End Sub



'
Public Sub DrawImg(Name As String, _
                   DestX As Single, DestY As Single, _
                   ImgWidth As Single, ImgHeight As Single)

  WriteObj "q" & vbCr & _
           ToStr(ToSpace(ImgWidth)) & " " & " 0 0 -" & ToStr(ToSpace(ImgHeight)) & _
           " " & ToStr(ToSpace(DestX)) & " " & ToStr(ToSpace(DestY)) & " cm" & vbCr & _
           "/" & Name & " Do" & vbCr & _
           "Q"

End Sub

' Carica 1 dei 14 font base Tipo1
Public Sub LoadFontStandard(Name As String, BaseFont As String, _
                    Optional Options As pdfFontStyle = pdfNormal)
  Dim i As Integer
  Dim sTemp As String
  
  BaseFont = Replace(BaseFont, " ", "")
  If Options = pdfBold Then
    sTemp = ",Bold"
  ElseIf Options = pdfBoldItalic Then
    sTemp = ",BoldItalic"
  ElseIf Options = pdfItalic Then
    sTemp = ",Italic"
  End If
  
  i = InitObj
  WriteObj "<< /Type /Font " & _
           "/Subtype /Type1 " & _
           "/Name /" & Name & " " & _
           "/BaseFont /" & BaseFont & sTemp & " " & _
           "/Encoding /WinAnsiEncoding >>"
  EndObj
  strFont = strFont & "/" & Name & " " & CStr(i) & " 0 R " & vbCr
End Sub

' Carica un font tipo TrueType
Public Sub LoadFont(Name As String, BaseFont As String, _
                    Optional Options As pdfFontStyle = pdfNormal)
  Dim i As Integer
  Dim j As Integer
  Dim sTemp As String
  
  intUsedFont = intUsedFont + 1
  ReDim Preserve arrFONT(intUsedFont)
  
  BaseFont = UCase(Replace(BaseFont, " ", ""))
  
  Select Case BaseFont
    Case "TIMESNEWROMAN"
      arrFONT(intUsedFont) = CreateFontTimes(Options)
    
    Case "COURIERNEW"
      arrFONT(intUsedFont) = CreateFontCourier(Options)
        
    Case "SYMBOL"
      arrFONT(intUsedFont) = CreateFontSymbol(Options)
      
    Case Else
      arrFONT(intUsedFont) = CreateFontArial(Options)
    
  End Select
  
  With arrFONT(intUsedFont)
    i = InitObj
    WriteObj "<< /Type /Font " & vbCr & _
             "/Subtype /TrueType" & vbCr & _
             "/Name /" & Name & vbCr & _
             "/BaseFont /" & .BaseFont & vbCr & _
             "/FirstChar " & CStr(.FirstChar) & vbCr & _
             "/LastChar " & CStr(.LastChar) & vbCr & _
             "/FontDescriptor " & CStr(i + 1) & " 0 R" & vbCr & _
             "/Encoding /WinAnsiEncoding" & vbCr & _
             "/Widths ["
    
    For j = .FirstChar To .LastChar
      sTemp = sTemp & ToStr(.Widths(j)) & " "
      If (((j - .FirstChar + 1) Mod 16) = 0) Or (j = .LastChar) Then
        WriteObj sTemp
        sTemp = ""
      End If
    Next
    
    WriteObj "] >>"
    EndObj
    strFont = strFont & "/" & Name & " " & CStr(i) & " 0 R " & vbCr
    strFontname = strFontname & "<" & Name & ">" & Right("0000" & CStr(intUsedFont), 4) & ";"
    
    i = InitObj
    WriteObj "<<" & vbCr & _
             "/Type /FontDescriptor" & vbCr & _
             "/FontName /" & .BaseFont & vbCr & _
             .Param & vbCr & _
             ">>"
    EndObj
  End With
End Sub


Public Sub StartObject(Name As String, Optional Options As pdfObjectType = pdfNull)
  Dim i As Integer
  
  i = InitObj
  WriteObj "<<" & vbCr & _
           "/Type /XObject" & vbCr & _
           "/Subtype /Form" & vbCr & _
           "/FormType 1" & vbCr & _
           "/Name /" & Name & vbCr & _
           "/BBox [" & ToStr(mvarMargin) & " " & _
                     ToStr(mvarMargin) & " " & _
                     ToStr(mvarPageWidth - mvarMargin) & " " & _
                     ToStr(mvarPageHeight - mvarMargin) & "]" & vbCr & _
           "/Matrix [1 0 0 1 0 0]" & vbCr & _
           "/Length " & CStr(i + 1) & " 0 R >>" & vbCr & _
           "stream"
  intOffsetTemp = Seek(mvarFileNumber)

  strObject = strObject & "/" & Name & " " & CStr(i) & " 0 R " & vbCr
  
  intObject = intObject + 1
  ReDim Preserve arrOBJECT(1 To intObject)
  With arrOBJECT(intObject)
    .Name = Name
    .Options = Options
  End With
 
End Sub

Public Sub EndObject()
  
  intOffsetTemp = Seek(mvarFileNumber) - intOffsetTemp
  WriteObj "endstream"
  EndObj
  
  ' Scrive la lunghezza
  InitObj
  WriteObj CStr(intOffsetTemp)
  EndObj
End Sub

'
Public Sub DrawObject(Name As String)
  WriteObj "/" & Name & " Do"
End Sub



' UTILITA': Funzioni di utilita'
Private Function ToSpace(Valore As Single) As Single
  
  ' Convert value to 72 pixel per inch
  Select Case mvarScaleMode
    
    Case pdfScaleMode.pdf72PxInch
      ToSpace = Valore
     
    Case pdfScaleMode.pdfInch
      ToSpace = Valore * 72
      
    Case pdfScaleMode.pdfCentimeter
      ToSpace = (Valore * 72) / 2.54
      
    Case pdfScaleMode.pdfMillimeter
      ToSpace = (Valore * 72) / 25.4
       
  End Select
 
End Function

Private Function ToUser(Valore As Single) As Single
  
  ' Convert value to 72 pixel per inch
  Select Case mvarScaleMode
    
    Case pdfScaleMode.pdf72PxInch
      ToUser = Valore
     
    Case pdfScaleMode.pdfInch
      ToUser = Valore / 72
      
    Case pdfScaleMode.pdfCentimeter
      ToUser = 2.54 * (Valore / 72)
      
    Case pdfScaleMode.pdfMillimeter
      ToUser = 25.4 * (Valore / 72)
       
  End Select
 
End Function

Private Function ToStr(Valore As Variant, Optional Dec As Integer = 3) As String
  ToStr = Replace(CStr(Round(Valore, Dec)), ",", ".")
End Function


Public Function length(Phrase As String, FontName As String, FontSize As Single) As Single

  Dim k As Single
  Dim i As Integer
  Dim C As Byte
  Dim l As Integer
  Dim j As Single
  Dim UsedFont As Integer
  Dim Options As Integer
  
  k = 0
  l = Len(Phrase)
  j = 0
  
  
  k = InStr(1, strFontname, "<" & FontName & ">", vbTextCompare)
  If k > 0 Then
    k = k + Len(FontName) + 2
    UsedFont = Val(Mid(strFontname, k, 4))
  
    With arrFONT(UsedFont)
      For i = 1 To l
        C = Asc(Mid(Phrase, i, 1))
        k = k + IIf((C >= .FirstChar) And (C <= .LastChar), .Widths(C), .MissingWidth)
        If C = 32 Then j = j + 1 ' conta gli spazi
      Next
    End With
    length = ToUser(((k * FontSize / 1000) + _
                     (j * mvarWordSpacing) + _
                     (l * mvarCharSpacing)) * _
                     (mvarTextScaling / 100))
  End If

End Function


Private Function ToASCII85(ByRef InBuf() As Byte) As String

  Dim i As Byte
  
  Dim m As Double
  Dim q As Long
  Dim l As Long
  Dim k As Long
  Dim j As Long
  Dim lngLen As Long
  
  Dim arrASCII() As Byte
      
  lngLen = UBound(InBuf)
  
  ReDim arrASCII(1 To Int(lngLen * 1.3) + 1)
    
  k = 1
  j = 1
  While k <= lngLen
    l = (lngLen - k) + 1
    m = 0
    
    For i = 1 To 4
      If i > l Then Exit For
      m = m + InBuf(k + i - 1) * (256 ^ (4 - i))
    Next
        
    If (m = 0) And (l = 4) Then
      arrASCII(j) = Asc("z")
      j = j + 1
    Else
      
      For i = 1 To 4
        q = Int(m / 85 ^ (5 - i))
        m = m - q * (85 ^ (5 - i))
        arrASCII(j + i - 1) = 33 + q
      Next
      
      arrASCII(j + 5 - 1) = 33 + m
      j = j + IIf(l < 4, l + 1, 5)
    End If
    
    k = k + 4
  Wend
    
  ReDim Preserve arrASCII(1 To j)
  
  ToASCII85 = StrConv(arrASCII, vbUnicode) & "~>"
End Function

'
Public Function LeggeBMP(FileName As String, _
                         ByRef ImgBuf() As Byte, _
                         ByRef ImgColor() As Byte, _
                         ByRef ImgWidth As Long, _
                         ByRef ImgHeight As Long, _
                         ByRef ImgBPP As Byte, _
                         Optional ColorSpace As pdfColorSpace = pdfRGB) As Boolean

  ' BITMAPFILEHEADER_Type
  Dim bfType          As String * 2     ' The string “BM” (hex value &H424D).
  Dim bfSize          As Long           ' The size of the file, measured in [Bytes].
  Dim bfDummy         As Integer        ' Not used, set to zero.
  Dim bfOffBits       As Long           ' The start offset of the bitmap data in the file.
  
  ' BITMAPINFOHEADER
  Dim biSize            As Long         ' 40 (the size of this structure).
  Dim biWidth           As Long         ' The width of the bitmap in pixels.
  Dim biHeight          As Long         ' The height of the bitmap in pixels.
  Dim biPlanes          As Integer      ' 1 (DIBs always have one plane).
  Dim biBitCount        As Integer      ' 1 for monochrome, 4 for 16 colors, 8 for 256 color, 24 for 24-bit RGB color.
  Dim biCompression     As Long         ' Specifies the type of compression for compressed
  Dim biSizeImage       As Long         ' The size of the image in bytes.
  Dim biXPelsPerMeter   As Long         ' Number of horizontal pixels per meter for
  Dim biYPelsPerMeter   As Long         ' Number of vertical pixels per meter for
  Dim biClrUsed         As Long         ' Number of entries in the DIB color table
  Dim biClrImportant    As Long         ' Number of entries in the DIB color table that
  
  
  Dim C As Long
  Dim fb As Integer
  Dim XBMP As Long
  Dim BPP As Byte
  Dim i As Long
  Dim kk As Long
  Dim blnFlag As Boolean
  Dim TempImg() As Byte
  Dim TempCol() As Byte ' RGBQUAD_Type
  Dim lngGray As Long
  
  fb = FreeFile
  Open FileName For Binary As #fb
  
  ' BITMAPFILEHEADER
  Get #fb, , bfType
  Get #fb, , bfSize
  Get #fb, , bfDummy
  Get #fb, , bfDummy
  Get #fb, , bfOffBits
  
  LeggeBMP = False
  
  If bfType = "BM" Then
    
    ' BITMAPINFOHEADER
    Get #fb, , biSize
    Get #fb, , biWidth
    Get #fb, , biHeight
    Get #fb, , biPlanes
    Get #fb, , biBitCount
    Get #fb, , biCompression
    Get #fb, , biSizeImage
    Get #fb, , biXPelsPerMeter
    Get #fb, , biYPelsPerMeter
    Get #fb, , biClrUsed
    Get #fb, , biClrImportant
    
    
    BPP = biBitCount
    
    If BPP <= 8 Then
        
      ' legge la palette di colori
      ReDim TempCol(1 To (2 ^ BPP) * 4)
      Get #fb, , TempCol()
      
      If ColorSpace = pdfRGB Then
        ReDim ImgColor(1 To 3 * (2 ^ BPP))
        For C = 0 To (2 ^ BPP) - 1
          ImgColor(C * 3 + 1) = TempCol(C * 4 + 1) ' red
          ImgColor(C * 3 + 2) = TempCol(C * 4 + 2) ' green
          ImgColor(C * 3 + 3) = TempCol(C * 4 + 3) ' blue
        Next
      Else
        ReDim ImgColor(1 To (2 ^ BPP))
        For C = 0 To (2 ^ BPP) - 1
          lngGray = (0.33 * TempCol(C * 4 + 1) + 0.59 * TempCol(C * 4 + 2) + 0.11 * TempCol(C * 4 + 3))
          ImgColor(C + 1) = IIf(lngGray > 255, 255, lngGray)
        Next
      End If
    End If
  
    XBMP = ((biWidth * BPP / 8) + 3&) And &HFFFFFFFC      ' [Bytes].
    
    ImgWidth = biWidth
    ImgHeight = biHeight
    ImgBPP = CByte(biBitCount)
    
    ReDim TempImg(1 To biHeight * XBMP)
    Get #fb, bfOffBits + 1, TempImg()
    Close #fb
    
    ReDim ImgBuf(1 To biHeight * XBMP)
    kk = 0
    
    If BPP > 8 Then
      blnFlag = ((biWidth Mod 4) <> 0)
      
      If ColorSpace = pdfRGB Then
        For C = 1 To UBound(TempImg) - 1 Step 3
          ImgBuf(3 * kk + 1) = TempImg(C + 2)
          ImgBuf(3 * kk + 2) = TempImg(C + 1)
          ImgBuf(3 * kk + 3) = TempImg(C)
          If (((kk + 1) Mod biWidth) = 0) And blnFlag Then C = C + (biWidth Mod 4)
          kk = kk + 1
        Next
      
      Else
        For C = 1 To UBound(TempImg) - 1 Step 3
          lngGray = 0.33 * TempImg(C + 2) + 0.59 * TempImg(C + 1) + 0.11 * TempImg(C)
          ImgBuf(kk + 1) = IIf(lngGray > 255, 255, lngGray)
          If (((kk + 1) Mod biWidth) = 0) And blnFlag Then C = C + (biWidth Mod 4)
          kk = kk + 1
        Next
        ReDim Preserve ImgBuf(1 To kk)
      
      End If
      
    ElseIf BPP <= 8 Then
      blnFlag = (biWidth Mod IIf(BPP = 8, 4, 8)) <> 0
      For i = 1 To UBound(TempImg)
        ImgBuf(kk + 1) = TempImg(i)
        If ((kk + 1) Mod Int((biWidth + (8 / BPP) - 1) / (8 / BPP))) = 0 And blnFlag Then
          i = i + (XBMP - (i Mod XBMP))
        End If
        kk = kk + 1
      Next
      ReDim Preserve ImgBuf(1 To kk)
    
    End If
    
    LeggeBMP = True
  End If

  If fb > 0 Then Close #fb

End Function

Private Function ToPdfStr(Temp As String) As String
  ToPdfStr = Replace(Replace(Replace(Temp, "\", "\\"), "(", "\("), ")", "\)")
End Function

Private Function CreateFontCourier(Style As pdfFontStyle) As FontDescriptor
  Dim i As Integer
  Dim awTemp()
  Dim fdTemp As FontDescriptor
  
  ' Courier New
  With fdTemp
    .BaseFont = "CourierNew"
    .FirstChar = 32
    .LastChar = 255
    .MissingWidth = 600
  End With
  
  Select Case Style
    Case pdfFontStyle.pdfNormal
      awTemp = Array(600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, _
                    600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, _
                    600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, _
                    600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600)
      fdTemp.Param = "/Flags 34 /FontBBox [-250 -300 720 1000] " & _
                     "/MissingWidth 600 /StemV 109 " & _
                     "/StemH 109 /ItalicAngle 0 /CapHeight 833 /XHeight 417 " & _
                     "/Ascent 833 /Descent -300 /Leading 133 " & _
                     "/MaxWidth 600 /AvgWidth 600"

    Case pdfFontStyle.pdfBold
      fdTemp.BaseFont = fdTemp.BaseFont & ",Bold"
      awTemp = Array(600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, _
                    600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, _
                    600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, _
                    600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600)
      fdTemp.Param = "/Flags 16418 /FontBBox [-250 -300 720 1000] " & _
                     "/MissingWidth 600 /StemV 191 " & _
                     "/StemH 191 /ItalicAngle 0 /CapHeight 833 /XHeight 417 " & _
                     "/Ascent 833 /Descent -300 /Leading 133 " & _
                     "/MaxWidth 600 /AvgWidth 600"

    Case pdfFontStyle.pdfItalic
      fdTemp.BaseFont = fdTemp.BaseFont & ",Italic"
      awTemp = Array(600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, _
                    600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, _
                    600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, _
                    600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600)
      fdTemp.Param = "/Flags 98 /FontBBox [-250 -300 720 1000] " & _
                     "/MissingWidth 600 /StemV 109 " & _
                     "/StemH 109 /ItalicAngle -11 /CapHeight 833 /XHeight 417 " & _
                     "/Ascent 833 /Descent -300 /Leading 133 " & _
                     "/MaxWidth 600 /AvgWidth 600"

    Case pdfFontStyle.pdfBoldItalic
      fdTemp.BaseFont = fdTemp.BaseFont & ",BoldItalic"
      awTemp = Array(600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, _
                    600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, _
                    600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, _
                    600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600)
      fdTemp.Param = "/Flags 16482 /FontBBox [-250 -300 720 1000] " & _
                     "/MissingWidth 600 /StemV 191 " & _
                     "/StemH 191 /ItalicAngle -11 /CapHeight 833 /XHeight 417 " & _
                     "/Ascent 833 /Descent -300 /Leading 133 " & _
                     "/MaxWidth 600 /AvgWidth 600"
  End Select
        
  For i = fdTemp.FirstChar To fdTemp.LastChar
    fdTemp.Widths(i) = awTemp(i - fdTemp.FirstChar + 1)
  Next
  
  CreateFontCourier = fdTemp

End Function

Private Function CreateFontTimes(Style As pdfFontStyle) As FontDescriptor
  Dim i As Integer
  Dim awTemp()
  Dim fdTemp As FontDescriptor
  
  ' Times New Roman
  With fdTemp
    .BaseFont = "TimesNewRoman"
    .FirstChar = 32
    .LastChar = 255
    .MissingWidth = 333

  End With
  
  Select Case Style
    Case pdfFontStyle.pdfNormal
      awTemp = Array(250, 333, 408, 500, 500, 833, 778, 180, 333, 333, 500, 564, 250, 333, 250, 278, 500, 500, 500, 500, 500, 500, 500, 500, 500, 500, 278, 278, 564, 564, 564, 444, 921, 722, 667, 667, 722, 611, 556, 722, 722, 333, 389, 722, 611, 889, 722, 722, 556, 722, 667, 556, 611, 722, 722, 944, 722, 722, 611, 333, 278, 333, 469, 500, _
                     333, 444, 500, 444, 500, 444, 333, 500, 500, 278, 278, 500, 278, 778, 500, 500, 500, 500, 333, 389, 278, 500, 500, 722, 500, 500, 444, 480, 200, 480, 541, 778, 500, 778, 333, 500, 444, 1000, 500, 500, 333, 1000, 556, 333, 889, 778, 611, 778, 778, 333, 333, 444, 444, 350, 500, 1000, 333, 980, 389, 333, 722, 778, 444, 722, _
                     250, 333, 500, 500, 500, 500, 200, 500, 333, 760, 276, 500, 564, 333, 760, 500, 400, 549, 300, 300, 333, 576, 453, 250, 333, 300, 310, 500, 750, 750, 750, 444, 722, 722, 722, 722, 722, 722, 889, 667, 611, 611, 611, 611, 333, 333, 333, 333, 722, 722, 722, 722, 722, 722, 722, 564, 722, 722, 722, 722, 722, 722, 556, 500, _
                     444, 444, 444, 444, 444, 444, 667, 444, 444, 444, 444, 444, 278, 278, 278, 278, 500, 500, 500, 500, 500, 500, 500, 549, 500, 500, 500, 500, 500, 500, 500, 500)
      fdTemp.Param = "/Flags 34 /FontBBox [-250 -216 1200 1000] " & _
                     "/MissingWidth 333 /StemV 73 " & _
                     "/StemH 73 /ItalicAngle 0 /CapHeight 891 /XHeight 446 " & _
                     "/Ascent 891 /Descent -216 /Leading 149 " & _
                     "/MaxWidth 1000 /AvgWidth 401"

    Case pdfFontStyle.pdfBold
      fdTemp.BaseFont = fdTemp.BaseFont & ",Bold"
      awTemp = Array(250, 333, 555, 500, 500, 1000, 833, 278, 333, 333, 500, 570, 250, 333, 250, 278, 500, 500, 500, 500, 500, 500, 500, 500, 500, 500, 333, 333, 570, 570, 570, 500, 930, 722, 667, 722, 722, 667, 611, 778, 778, 389, 500, 778, 667, 944, 722, 778, 611, 778, 722, 556, 667, 722, 722, 1000, 722, 722, 667, 333, 278, 333, 581, 500, _
                     333, 500, 556, 444, 556, 444, 333, 500, 556, 278, 333, 556, 278, 833, 556, 500, 556, 556, 444, 389, 333, 556, 500, 722, 500, 500, 444, 394, 220, 394, 520, 778, 500, 778, 333, 500, 500, 1000, 500, 500, 333, 1000, 556, 333, 1000, 778, 667, 778, 778, 333, 333, 500, 500, 350, 500, 1000, 333, 1000, 389, 333, 722, 778, 444, 722, _
                     250, 333, 500, 500, 500, 500, 220, 500, 333, 747, 300, 500, 570, 333, 747, 500, 400, 549, 300, 300, 333, 576, 540, 250, 333, 300, 330, 500, 750, 750, 750, 500, 722, 722, 722, 722, 722, 722, 1000, 722, 667, 667, 667, 667, 389, 389, 389, 389, 722, 722, 778, 778, 778, 778, 778, 570, 778, 722, 722, 722, 722, 722, 611, 556, _
                     500, 500, 500, 500, 500, 500, 722, 444, 444, 444, 444, 444, 278, 278, 278, 278, 500, 556, 500, 500, 500, 500, 500, 549, 500, 556, 556, 556, 556, 500, 556, 500)
      fdTemp.Param = "/Flags 16418 /FontBBox [-250 -216 1201 1000] " & _
                     "/MissingWidth 333 /StemV 136 " & _
                     "/StemH 136 /ItalicAngle 0 /CapHeight 891 /XHeight 446 " & _
                     "/Ascent 891 /Descent -216 /Leading 149 " & _
                     "/MaxWidth 1001 /AvgWidth 401"

    Case pdfFontStyle.pdfItalic
      fdTemp.BaseFont = fdTemp.BaseFont & ",Italic"
      awTemp = Array(250, 333, 420, 500, 500, 833, 778, 214, 333, 333, 500, 675, 250, 333, 250, 278, 500, 500, 500, 500, 500, 500, 500, 500, 500, 500, 333, 333, 675, 675, 675, 500, 920, 611, 611, 667, 722, 611, 611, 722, 722, 333, 444, 667, 556, 833, 667, 722, 611, 722, 611, 500, 556, 722, 611, 833, 611, 556, 556, 389, 278, 389, 422, 500, _
                     333, 500, 500, 444, 500, 444, 278, 500, 500, 278, 278, 444, 278, 722, 500, 500, 500, 500, 389, 389, 278, 500, 444, 667, 444, 444, 389, 400, 275, 400, 541, 778, 500, 778, 333, 500, 556, 889, 500, 500, 333, 1000, 500, 333, 944, 778, 556, 778, 778, 333, 333, 556, 556, 350, 500, 889, 333, 980, 389, 333, 667, 778, 389, 556, _
                     250, 389, 500, 500, 500, 500, 275, 500, 333, 760, 276, 500, 675, 333, 760, 500, 400, 549, 300, 300, 333, 576, 523, 250, 333, 300, 310, 500, 750, 750, 750, 500, 611, 611, 611, 611, 611, 611, 889, 667, 611, 611, 611, 611, 333, 333, 333, 333, 722, 667, 722, 722, 722, 722, 722, 675, 722, 722, 722, 722, 722, 556, 611, 500, _
                     500, 500, 500, 500, 500, 500, 667, 444, 444, 444, 444, 444, 278, 278, 278, 278, 500, 500, 500, 500, 500, 500, 500, 549, 500, 500, 500, 500, 500, 444, 500, 444)
      fdTemp.Param = "/Flags 98 /FontBBox [-250 -216 1200 1000] " & _
                     "/MissingWidth 333 /StemV 73 " & _
                     "/StemH 73 /ItalicAngle -11 /CapHeight 891 /XHeight 446 " & _
                     "/Ascent 891 /Descent -216 /Leading 149 " & _
                     "/MaxWidth 1000 /AvgWidth 402"

    Case pdfFontStyle.pdfBoldItalic
      fdTemp.BaseFont = fdTemp.BaseFont & ",BoldItalic"
      awTemp = Array(250, 389, 555, 500, 500, 833, 778, 278, 333, 333, 500, 570, 250, 333, 250, 278, 500, 500, 500, 500, 500, 500, 500, 500, 500, 500, 333, 333, 570, 570, 570, 500, 832, 667, 667, 667, 722, 667, 667, 722, 778, 389, 500, 667, 611, 889, 722, 722, 611, 722, 667, 556, 611, 722, 667, 889, 667, 611, 611, 333, 278, 333, 570, 500, _
                     333, 500, 500, 444, 500, 444, 333, 500, 556, 278, 278, 500, 278, 778, 556, 500, 500, 500, 389, 389, 278, 556, 444, 667, 500, 444, 389, 348, 220, 348, 570, 778, 500, 778, 333, 500, 500, 1000, 500, 500, 333, 1000, 556, 333, 944, 778, 611, 778, 778, 333, 333, 500, 500, 350, 500, 1000, 333, 1000, 389, 333, 722, 778, 389, 611, _
                     250, 389, 500, 500, 500, 500, 220, 500, 333, 747, 266, 500, 606, 333, 747, 500, 400, 549, 300, 300, 333, 576, 500, 250, 333, 300, 300, 500, 750, 750, 750, 500, 667, 667, 667, 667, 667, 667, 944, 667, 667, 667, 667, 667, 389, 389, 389, 389, 722, 722, 722, 722, 722, 722, 722, 570, 722, 722, 722, 722, 722, 611, 611, 500, _
                     500, 500, 500, 500, 500, 500, 722, 444, 444, 444, 444, 444, 278, 278, 278, 278, 500, 556, 500, 500, 500, 500, 500, 549, 500, 556, 556, 556, 556, 444, 500, 444)
      fdTemp.Param = "/Flags 16482 /FontBBox [-250 -216 1200 1000] " & _
                     "/MissingWidth 333 /StemV 131 " & _
                     "/StemH 131 /ItalicAngle -11 /CapHeight 891 /XHeight 446 " & _
                     "/Ascent 891 /Descent -216 /Leading 149 " & _
                     "/MaxWidth 1000 /AvgWidth 412"
  End Select
        
  For i = fdTemp.FirstChar To fdTemp.LastChar
    fdTemp.Widths(i) = awTemp(i - fdTemp.FirstChar + 1)
  Next
  
  CreateFontTimes = fdTemp

End Function


Private Function CreateFontSymbol(Style As pdfFontStyle) As FontDescriptor
  Dim i As Integer
  Dim awTemp()
  Dim fdTemp As FontDescriptor
  
  ' Symbol
  With fdTemp
    .BaseFont = "Symbol"
    .FirstChar = 30
    .LastChar = 255
    .MissingWidth = 332
  End With
  
  Select Case Style
    Case pdfFontStyle.pdfNormal
      awTemp = Array(600, 600, 250, 333, 713, 500, 549, 833, 778, 439, 333, 333, 500, 549, 250, 549, 250, 278, 500, 500, 500, 500, 500, 500, 500, 500, 500, 500, 278, 278, 549, 549, 549, 444, 549, 722, 667, 722, 612, 611, 763, 603, 722, 333, 631, 722, 686, 889, 722, 722, 768, 741, 556, 592, 611, 690, 439, 768, 645, 795, 611, 333, 863, 333, _
                     658, 500, 500, 631, 549, 549, 494, 439, 521, 411, 603, 329, 603, 549, 549, 576, 521, 549, 549, 521, 549, 603, 439, 576, 713, 686, 493, 686, 494, 480, 200, 480, 549, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, _
                     600, 600, 600, 620, 247, 549, 167, 713, 500, 753, 753, 753, 753, 1042, 987, 603, 987, 603, 400, 549, 411, 549, 549, 713, 494, 460, 549, 549, 549, 549, 1000, 603, 1000, 658, 823, 686, 795, 987, 768, 768, 823, 768, 768, 713, 713, 713, 713, 713, 713, 713, 768, 713, 790, 790, 890, 823, 549, 250, 713, 603, 603, 1042, 987, 603, _
                     987, 603, 494, 329, 790, 790, 786, 713, 384, 384, 384, 384, 384, 384, 494, 494, 494, 494, 600, 329, 274, 686, 686, 686, 384, 384, 384, 384, 384, 384, 494, 494, 494, 600)
      fdTemp.Param = "/Flags 6 /FontBBox [-250 -220 1246 1005] " & _
                     "/MissingWidth 332 /StemV 109 " & _
                     "/StemH 109 /ItalicAngle 0 /CapHeight 1005 /XHeight 503 " & _
                     "/Ascent 1005 /Descent -220 /Leading 225 " & _
                     "/MaxWidth 1038 /AvgWidth 601"

    Case pdfFontStyle.pdfBold
      fdTemp.BaseFont = fdTemp.BaseFont & ",Bold"
      awTemp = Array(600, 600, 250, 333, 713, 500, 549, 833, 778, 439, 333, 333, 500, 549, 250, 549, 250, 278, 500, 500, 500, 500, 500, 500, 500, 500, 500, 500, 278, 278, 549, 549, 549, 444, 549, 722, 667, 722, 612, 611, 763, 603, 722, 333, 631, 722, 686, 889, 722, 722, 768, 741, 556, 592, 611, 690, 439, 768, 645, 795, 611, 333, 863, 333, _
                     658, 500, 500, 631, 549, 549, 494, 439, 521, 411, 603, 329, 603, 549, 549, 576, 521, 549, 549, 521, 549, 603, 439, 576, 713, 686, 493, 686, 494, 480, 200, 480, 549, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, _
                     600, 600, 600, 620, 247, 549, 167, 713, 500, 753, 753, 753, 753, 1042, 987, 603, 987, 603, 400, 549, 411, 549, 549, 713, 494, 460, 549, 549, 549, 549, 1000, 603, 1000, 658, 823, 686, 795, 987, 768, 768, 823, 768, 768, 713, 713, 713, 713, 713, 713, 713, 768, 713, 790, 790, 890, 823, 549, 250, 713, 603, 603, 1042, 987, 603, _
                     987, 603, 494, 329, 790, 790, 786, 713, 384, 384, 384, 384, 384, 384, 494, 494, 494, 494, 600, 329, 274, 686, 686, 686, 384, 384, 384, 384, 384, 384, 494, 494, 494, 600)
      fdTemp.Param = "/Flags 16390 /FontBBox [-250 -220 1246 1005] " & _
                     "/MissingWidth 332 /StemV 191 " & _
                     "/StemH 191 /ItalicAngle 0 /CapHeight 1005 /XHeight 503 " & _
                     "/Ascent 1005 /Descent -220 /Leading 225 " & _
                     "/MaxWidth 1038 /AvgWidth 600"

    Case pdfFontStyle.pdfItalic
      fdTemp.BaseFont = fdTemp.BaseFont & ",Italic"
      awTemp = Array(600, 600, 250, 333, 713, 500, 549, 833, 778, 439, 333, 333, 500, 549, 250, 549, 250, 278, 500, 500, 500, 500, 500, 500, 500, 500, 500, 500, 278, 278, 549, 549, 549, 444, 549, 722, 667, 722, 612, 611, 763, 603, 722, 333, 631, 722, 686, 889, 722, 722, 768, 741, 556, 592, 611, 690, 439, 768, 645, 795, 611, 333, 863, 333, _
                     658, 500, 500, 631, 549, 549, 494, 439, 521, 411, 603, 329, 603, 549, 549, 576, 521, 549, 549, 521, 549, 603, 439, 576, 713, 686, 493, 686, 494, 480, 200, 480, 549, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, _
                     600, 600, 600, 620, 247, 549, 167, 713, 500, 753, 753, 753, 753, 1042, 987, 603, 987, 603, 400, 549, 411, 549, 549, 713, 494, 460, 549, 549, 549, 549, 1000, 603, 1000, 658, 823, 686, 795, 987, 768, 768, 823, 768, 768, 713, 713, 713, 713, 713, 713, 713, 768, 713, 790, 790, 890, 823, 549, 250, 713, 603, 603, 1042, 987, 603, _
                     987, 603, 494, 329, 790, 790, 786, 713, 384, 384, 384, 384, 384, 384, 494, 494, 494, 494, 600, 329, 274, 686, 686, 686, 384, 384, 384, 384, 384, 384, 494, 494, 494, 600)
      fdTemp.Param = "/Flags 70 /FontBBox [-250 -220 1246 1005] " & _
                     "/MissingWidth 332 /StemV 109 " & _
                     "/StemH 109 /ItalicAngle -11 /CapHeight 1005 /XHeight 503 " & _
                     "/Ascent 1005 /Descent -220 /Leading 225 " & _
                     "/MaxWidth 1038 /AvgWidth 600"

    Case pdfFontStyle.pdfBoldItalic
      fdTemp.BaseFont = fdTemp.BaseFont & ",BoldItalic"
      awTemp = Array(600, 600, 250, 333, 713, 500, 549, 833, 778, 439, 333, 333, 500, 549, 250, 549, 250, 278, 500, 500, 500, 500, 500, 500, 500, 500, 500, 500, 278, 278, 549, 549, 549, 444, 549, 722, 667, 722, 612, 611, 763, 603, 722, 333, 631, 722, 686, 889, 722, 722, 768, 741, 556, 592, 611, 690, 439, 768, 645, 795, 611, 333, 863, 333, _
                     658, 500, 500, 631, 549, 549, 494, 439, 521, 411, 603, 329, 603, 549, 549, 576, 521, 549, 549, 521, 549, 603, 439, 576, 713, 686, 493, 686, 494, 480, 200, 480, 549, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, 600, _
                     600, 600, 600, 620, 247, 549, 167, 713, 500, 753, 753, 753, 753, 1042, 987, 603, 987, 603, 400, 549, 411, 549, 549, 713, 494, 460, 549, 549, 549, 549, 1000, 603, 1000, 658, 823, 686, 795, 987, 768, 768, 823, 768, 768, 713, 713, 713, 713, 713, 713, 713, 768, 713, 790, 790, 890, 823, 549, 250, 713, 603, 603, 1042, 987, 603, _
                     987, 603, 494, 329, 790, 790, 786, 713, 384, 384, 384, 384, 384, 384, 494, 494, 494, 494, 600, 329, 274, 686, 686, 686, 384, 384, 384, 384, 384, 384, 494, 494, 494, 600)
      fdTemp.Param = "/Flags 16454 /FontBBox [-250 -220 1246 1005] " & _
                     "/MissingWidth 332 /StemV 191 " & _
                     "/StemH 191 /ItalicAngle -11 /CapHeight 1005 /XHeight 503 " & _
                     "/Ascent 1005 /Descent -220 /Leading 225 " & _
                     "/MaxWidth 1038 /AvgWidth 600"
  End Select
        
  For i = fdTemp.FirstChar To fdTemp.LastChar
    fdTemp.Widths(i) = awTemp(i - fdTemp.FirstChar + 1)
  Next
  
  CreateFontSymbol = fdTemp
    
End Function


Private Function CreateFontArial(Style As pdfFontStyle) As FontDescriptor
  Dim i As Integer
  Dim awTemp()
  Dim fdTemp As FontDescriptor
  
  ' Arial
  With fdTemp
    .BaseFont = "Arial"
    .FirstChar = 32
    .LastChar = 255
    .MissingWidth = 272
  End With
  
  Select Case Style
    Case pdfFontStyle.pdfNormal
      awTemp = Array(278, 278, 355, 556, 556, 889, 667, 191, 333, 333, 389, 584, 278, 333, 278, 278, 556, 556, 556, 556, 556, 556, 556, 556, 556, 556, 278, 278, 584, 584, 584, 556, 1015, 667, 667, 722, 722, 667, 611, 778, 722, 278, 500, 667, 556, 833, 722, 778, 667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 278, 278, 278, 469, 556, _
                     333, 556, 556, 500, 556, 556, 278, 556, 556, 222, 222, 500, 222, 833, 556, 556, 556, 556, 333, 500, 278, 556, 500, 722, 500, 500, 500, 334, 260, 334, 584, 750, 556, 750, 222, 556, 333, 1000, 556, 556, 333, 1000, 667, 333, 1000, 750, 611, 750, 750, 222, 222, 333, 333, 350, 556, 1000, 333, 1000, 500, 333, 944, 750, 500, 667, _
                     278, 333, 556, 556, 556, 556, 260, 556, 333, 737, 370, 556, 584, 333, 737, 552, 400, 549, 333, 333, 333, 576, 537, 278, 333, 333, 365, 556, 834, 834, 834, 611, 667, 667, 667, 667, 667, 667, 1000, 722, 667, 667, 667, 667, 278, 278, 278, 278, 722, 722, 778, 778, 778, 778, 778, 584, 778, 722, 722, 722, 722, 667, 667, 611, _
                     556, 556, 556, 556, 556, 556, 889, 500, 556, 556, 556, 556, 278, 278, 278, 278, 556, 556, 556, 556, 556, 556, 556, 549, 611, 556, 556, 556, 556, 500, 556, 500)
      fdTemp.Param = "/Flags 32 /FontBBox [-250 -221 1190 1000] " & _
                     "/MissingWidth 272 /StemV 80 " & _
                     "/StemH 80 /ItalicAngle 0 /CapHeight 905 /XHeight 453 " & _
                     "/Ascent 905 /Descent -212 /Leading 150 " & _
                     "/MaxWidth 992 /AvgWidth 441"

    Case pdfFontStyle.pdfBold
      fdTemp.BaseFont = fdTemp.BaseFont & ",Bold"
      awTemp = Array(278, 333, 474, 556, 556, 889, 722, 238, 333, 333, 389, 584, 278, 333, 278, 278, 556, 556, 556, 556, 556, 556, 556, 556, 556, 556, 333, 333, 584, 584, 584, 611, 975, 722, 722, 722, 722, 667, 611, 778, 722, 278, 556, 722, 611, 833, 722, 778, 667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 333, 278, 333, 584, 556, _
                     333, 556, 611, 556, 611, 556, 333, 611, 611, 278, 278, 556, 278, 889, 611, 611, 611, 611, 389, 556, 333, 611, 556, 778, 556, 556, 500, 389, 280, 389, 584, 750, 556, 750, 278, 556, 500, 1000, 556, 556, 333, 1000, 667, 333, 1000, 750, 611, 750, 750, 278, 278, 500, 500, 350, 556, 1000, 333, 1000, 556, 333, 944, 750, 500, 667, _
                     278, 333, 556, 556, 556, 556, 280, 556, 333, 737, 370, 556, 584, 333, 737, 552, 400, 549, 333, 333, 333, 576, 556, 278, 333, 333, 365, 556, 834, 834, 834, 611, 722, 722, 722, 722, 722, 722, 1000, 722, 667, 667, 667, 667, 278, 278, 278, 278, 722, 722, 778, 778, 778, 778, 778, 584, 778, 722, 722, 722, 722, 667, 667, 611, _
                     556, 556, 556, 556, 556, 556, 889, 556, 556, 556, 556, 556, 278, 278, 278, 278, 611, 611, 611, 611, 611, 611, 611, 549, 611, 611, 611, 611, 611, 556, 611, 556)
      fdTemp.Param = "/Flags 16416 /FontBBox [-250 -212 1120 1000] " & _
                     "/MissingWidth 311 /StemV 153 " & _
                     "/StemH 153 /ItalicAngle 0 /CapHeight 905 /XHeight 453 " & _
                     "/Ascent 905 /Descent -212 /Leading 150 " & _
                     "/MaxWidth 933 /AvgWidth 479"
      fdTemp.MissingWidth = 311
                     

    Case pdfFontStyle.pdfItalic
      fdTemp.BaseFont = fdTemp.BaseFont & ",Italic"
      awTemp = Array(278, 278, 355, 556, 556, 889, 667, 191, 333, 333, 389, 584, 278, 333, 278, 278, 556, 556, 556, 556, 556, 556, 556, 556, 556, 556, 278, 278, 584, 584, 584, 556, 1015, 667, 667, 722, 722, 667, 611, 778, 722, 278, 500, 667, 556, 833, 722, 778, 667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 278, 278, 278, 469, 556, _
                     333, 556, 556, 500, 556, 556, 278, 556, 556, 222, 222, 500, 222, 833, 556, 556, 556, 556, 333, 500, 278, 556, 500, 722, 500, 500, 500, 334, 260, 334, 584, 750, 556, 750, 222, 556, 333, 1000, 556, 556, 333, 1000, 667, 333, 1000, 750, 611, 750, 750, 222, 222, 333, 333, 350, 556, 1000, 333, 1000, 500, 333, 944, 750, 500, 667, _
                     278, 333, 556, 556, 556, 556, 260, 556, 333, 737, 370, 556, 584, 333, 737, 552, 400, 549, 333, 333, 333, 576, 537, 278, 333, 333, 365, 556, 834, 834, 834, 611, 667, 667, 667, 667, 667, 667, 1000, 722, 667, 667, 667, 667, 278, 278, 278, 278, 722, 722, 778, 778, 778, 778, 778, 584, 778, 722, 722, 722, 722, 667, 667, 611, _
                     556, 556, 556, 556, 556, 556, 889, 500, 556, 556, 556, 556, 278, 278, 278, 278, 556, 556, 556, 556, 556, 556, 556, 549, 611, 556, 556, 556, 556, 500, 556, 500)
      fdTemp.Param = "/Flags 96 /FontBBox [-250 -212 1134 1000] " & _
                     "/MissingWidth 259 /StemV 80 " & _
                     "/StemH 80 /ItalicAngle -11 /CapHeight 905 /XHeight 453 " & _
                     "/Ascent 905 /Descent -212 /Leading 150 " & _
                     "/MaxWidth 945 /AvgWidth 441"
      fdTemp.MissingWidth = 259

    Case pdfFontStyle.pdfBoldItalic
      fdTemp.BaseFont = fdTemp.BaseFont & ",BoldItalic"
      awTemp = Array(278, 333, 474, 556, 556, 889, 722, 238, 333, 333, 389, 584, 278, 333, 278, 278, 556, 556, 556, 556, 556, 556, 556, 556, 556, 556, 333, 333, 584, 584, 584, 611, 975, 722, 722, 722, 722, 667, 611, 778, 722, 278, 556, 722, 611, 833, 722, 778, 667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 333, 278, 333, 584, 556, _
                     333, 556, 611, 556, 611, 556, 333, 611, 611, 278, 278, 556, 278, 889, 611, 611, 611, 611, 389, 556, 333, 611, 556, 778, 556, 556, 500, 389, 280, 389, 584, 750, 556, 750, 278, 556, 500, 1000, 556, 556, 333, 1000, 667, 333, 1000, 750, 611, 750, 750, 278, 278, 500, 500, 350, 556, 1000, 333, 1000, 556, 333, 944, 750, 500, 667, _
                     278, 333, 556, 556, 556, 556, 280, 556, 333, 737, 370, 556, 584, 333, 737, 552, 400, 549, 333, 333, 333, 576, 556, 278, 333, 333, 365, 556, 834, 834, 834, 611, 722, 722, 722, 722, 722, 722, 1000, 722, 667, 667, 667, 667, 278, 278, 278, 278, 722, 722, 778, 778, 778, 778, 778, 584, 778, 722, 722, 722, 722, 667, 667, 611, _
                     556, 556, 556, 556, 556, 556, 889, 556, 556, 556, 556, 556, 278, 278, 278, 278, 611, 611, 611, 611, 611, 611, 611, 549, 611, 611, 611, 611, 611, 556, 611, 556)
      fdTemp.Param = "/Flags 16480 /FontBBox [-250 -212 1120 1000] " & _
                     "/MissingWidth 311 /StemV 153 " & _
                     "/StemH 153 /ItalicAngle -11 /CapHeight 905 /XHeight 453 " & _
                     "/Ascent 905 /Descent -212 /Leading 150 " & _
                     "/MaxWidth 933 /AvgWidth 479"
      fdTemp.MissingWidth = 311
  End Select
  
  For i = fdTemp.FirstChar To fdTemp.LastChar
    fdTemp.Widths(i) = awTemp(i - fdTemp.FirstChar + 1)
  Next
  
  CreateFontArial = fdTemp
  
End Function

' Verifica se ci sono oggetti comuni da inserire sulla pagina
Private Sub InsertObjectOnPage()
  Dim i As Integer
  
  If intObject > 0 Then
    For i = 1 To intObject
      With arrOBJECT(i)
        If (((.Options And pdfAllPages) = pdfAllPages) Or _
           (((.Options And pdfEvenPages) <> 0) And ((mvarPages Mod 2) = 0)) Or _
           (((.Options And pdfOddPages) <> 0) And ((mvarPages Mod 2) <> 0)) And (Not _
           (((.Options And pdfNotFirstPage) <> 0) And (mvarPages = 1)))) Then
          DrawObject .Name
        End If
      End With
    Next
  End If

End Sub

''''''''''''''''''''  INTERFACE PROPERTIES AND FUNCTIONS  '''''''''''''''''''''''''''''''''''
Public Property Let PageBreakBefore(ByVal sPageBr As String)
   m_PageBreakBefore = sPageBr
End Property

Public Property Get PageBreakBefore() As String
   PageBreakBefore = m_PageBreakBefore
End Property

Public Property Get SourceFile() As Variant
    SourceFile = m_strSourceFile
End Property

Public Property Let SourceFile(ByVal vNewValue As Variant)
  m_strSourceFile = vNewValue
End Property

Public Property Get TargetFile() As Variant
    TargetFile = m_strTargetFile
End Property

Public Property Let TargetFile(ByVal vNewValue As Variant)
    m_strTargetFile = vNewValue
End Property

Public Property Get RecordsPerPage() As Variant
    RecordsPerPage = m_intRPP
End Property

Public Property Let RecordsPerPage(ByVal vNewValue As Variant)
    m_intRPP = vNewValue
End Property

Public Property Get Font() As Variant
    Font = m_strFontname
End Property

Public Property Let Font(ByVal vNewValue As Variant)
    m_strFontname = vNewValue
End Property

Public Property Get FontSize() As Variant
    FontSize = m_intFontSize
End Property

Public Property Let FontSize(ByVal vNewValue As Variant)
    m_intFontSize = vNewValue
End Property

Public Function CreatePDFFile() As String
    Dim objFile As FileServer
    Dim strRow As String
    Dim i As Long
    Dim j As Integer
    Dim nLeft As Integer
    Dim m As Integer
    
    On Error GoTo errTrap
    
    If Not isFileExist() Then
       CreatePDFFile = m_strSourceFile & " File does not exist!"
       Exit Function
    End If
    
    Set objFile = New FileServer

    If Not objFile.PrepareFile(m_strSourceFile) Then
       CreatePDFFile = "Failed to open " & m_strSourceFile & "!"
       Exit Function
    End If
    
    Dim dblElapsed As Single
    dblElapsed = Timer
    
    Const dLeft As Single = 2#
    Const dTop As Single = 27#
    Const dLineIncr As Single = 0.4
    Dim bStartPage As Boolean
    Dim bEndPage As Boolean
    Dim dLinePosition As Single
    Dim sName As String
    Dim nPosi As Integer
    
    nPosi = InStrRev(m_strTargetFile, "\")
    
    If nPosi > 0 Then
        sName = Mid(m_strTargetFile, nPosi + 1)
    Else
        sName = m_strTargetFile
    End If
    
    Title = sName
    ScaleMode = pdfCentimeter
    PaperSize = pdfA4
    Margin = 0
    Orientation = pdfPortrait
    
    'open the pdf file
    InitPDFFile m_strTargetFile
    
    LoadFont "Fnt3", m_strFontname
    LoadFontStandard "Fnt4", m_strFontname, pdfBoldItalic

    bStartPage = True
    bEndPage = False
    dLinePosition = dTop + dLineIncr
    i = 0
    
    Do While Not objFile.EndOfFile
        strRow = objFile.GetLine
        i = i + 1
        
        '#Special code. Remove the first 3 empty lines
        If m_PageBreakBefore <> "" And i = 1 Then
            m = 0
            Do While m <= 3
                If Trim(strRow) = "" Then
                    If Not objFile.EndOfFile Then
                        strRow = objFile.GetLine
                        i = i + 1
                    End If
                Else
                    Exit Do
                End If
                m = m + 1
            Loop
        End If
        
        If m_PageBreakBefore <> "" Then
            If i > 10 Then
                If InStr(strRow, m_PageBreakBefore) > 0 Then
                    bStartPage = True
                    dLinePosition = dTop + dLineIncr
                    i = 1
                End If
            End If
        End If
        
        If i > m_intRPP Then
            bStartPage = True
            dLinePosition = dTop + dLineIncr
            i = 1
        End If
        
        If bStartPage Then
            If bEndPage Then
                EndPage
            End If
            
            StartPage
            bStartPage = False
            bEndPage = True
        End If
        
        'write to pdf
        dLinePosition = dLinePosition - dLineIncr
        DrawText dLeft, dLinePosition, strRow, "Fnt3", m_intFontSize, pdfAlignLeft
    Loop
    
    If bEndPage Then
        EndPage
    End If
    
    ClosePDFFile
    objFile.CloseFile
    
    dblElapsed = Timer - dblElapsed
    CreatePDFFile = ""
    Exit Function
    
errTrap:
    CreatePDFFile = "An error has occurred while creating the PDF file. Err Code: " _
        & Err.Number & ", Err Desc: " & Err.Description
End Function

Private Function isFileExist() As Boolean
   isFileExist = IIf(Len(Dir(m_strSourceFile)) > 0, 1, 0)
End Function
