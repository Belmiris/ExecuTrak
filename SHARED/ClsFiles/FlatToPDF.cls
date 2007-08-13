VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FlatToPDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

'###################################################################
'#This class contains all the functions to create
'#PDF file from the ascii file.
'#The idea to do this if from a FREEWARE published by Dinesh Asanka.
'#This class requires FileServer.cls
'
'#Sample: Return
'Private Function fnConvertToPDF(ByVal Source_TXT_File As String, _
'                                ByVal Target_PDF_File As String) As String
'    Dim objVBPDF As New FlatToPDF
'
'    With objVBPDF
'        .SourceFile = Source_TXT_File
'        .TargetFile = Target_PDF_File
'        .RecordsPerPage = 70
'        .Width = 83
'        .PageBreakBefore = "     Page:"
'        .FontSize = 10
'        .Font = "Courier"
'        fnConvertToPDF = .CreatePDFFile
'    End If
'End Function
'# Weigong Jiang, 03/26/2004
'####################################################################
Private strSourceFile As String
Private strTargetFiler As String
Private strFontName As String
Private intWidth As Long
Private intRPP As Long
Private intFontSize As Long
Private intPGCount As Long
Private m_PageBreakBefore As String

'Internal Variables
Private strPage As String
Private intObject As Long
Private intOffSet As Long
Private strCR As String
Private strPR As String
Private strTI As String
Private intLen As Long
Private rows As Long
Private arcount As Long
'Data Arrays
Private arrXREF() As String
Private arrData() As String
Private arrTmp() As String
'Conntant value
Const END_OBJECT = "endobj"
Const BEG_OBJECT = " 0 obj"
Const AR_1 = "<<"
Const AR_2 = ">>"
Const AR_OR = " 0 R"
Const XREF_END_CHAR = " 00000 n"

Public Property Let PageBreakBefore(ByVal sPageBr As String)
   m_PageBreakBefore = sPageBr
End Property
Public Property Get PageBreakBefore() As String
   PageBreakBefore = m_PageBreakBefore
End Property

Public Property Get SourceFile() As Variant
    SourceFile = strSourceFile
End Property

Public Property Let SourceFile(ByVal vNewValue As Variant)
  strSourceFile = vNewValue
End Property

Public Property Get TargetFile() As Variant
    strTargetFiler = TargetFile
End Property

Public Property Let TargetFile(ByVal vNewValue As Variant)
    strTargetFiler = vNewValue
End Property

Public Property Get Width() As Variant
    Width = intWidth
End Property

Public Property Let Width(ByVal vNewValue As Variant)
    intWidth = vNewValue
End Property

Public Property Get RecordsPerPage() As Variant
    RecordsPerPage = intRPP
End Property

Public Property Let RecordsPerPage(ByVal vNewValue As Variant)
    intRPP = vNewValue
End Property

Public Function CreatePDFFile() As String
    CreatePDFFile = ""
    If isFileExist Then
       InitVar
       FillXREFArray
       FillDataArray
       FillTmpArrayFromTextFile
       FillData
       FillFotter
       WriteToReport
    Else
       CreatePDFFile = strSourceFile & " File does not exist!"
    End If
End Function

Private Function isFileExist() As Boolean
   isFileExist = IIf(Len(Dir(strSourceFile)) > 0, 1, 0)
End Function

Private Sub InitVar()
    strCR = Chr(67) & Chr(114) & Chr(101) & Chr(97) & Chr(116) & Chr(111) & Chr(114)
    strPR = Chr(80) & Chr(114) & Chr(111) & Chr(100) & Chr(117) & Chr(99) & Chr(101) & Chr(114)
    strTI = Chr(84) & Chr(105) & Chr(116) & Chr(108) & Chr(101)
    intObject = 6
    intOffSet = 396
End Sub

Private Sub FillXREFArray()
    ReDim Preserve arrXREF(1 To 10)
    arrXREF(1) = "xref"
    arrXREF(2) = "0 10"
    arrXREF(3) = "0000000000 65535 f"
    arrXREF(4) = "0000000017" & XREF_END_CHAR
    arrXREF(5) = "0000000790" & XREF_END_CHAR
    arrXREF(6) = "0000000869" & XREF_END_CHAR
    arrXREF(7) = "0000000144" & XREF_END_CHAR
    arrXREF(8) = "0000000247" & XREF_END_CHAR
    arrXREF(9) = "0000000321" & XREF_END_CHAR
    arrXREF(10) = "0000000396" & XREF_END_CHAR

End Sub

Private Sub FillDataArray()
    ReDim Preserve arrData(1 To 29)
    arrData(1) = "%" & Chr(80) & Chr(68) & Chr(70) & "-1.2 "
    arrData(2) = "%сссс"
    arrData(3) = "1" & BEG_OBJECT
    arrData(4) = AR_1
    'arrData(5) = "/" & strCR & "  (ssfactor " & Chr(80) & Chr(83) & Chr(79) & Chr(80) & Chr(68) & Chr(70) & ")"
    arrData(5) = "/" & strCR & "  (FLATTPDF)"
    arrData(6) = "/" & strPR & " (ACH Notice - FLAT to PDF dchai@factor.com)"
    arrData(7) = "/" & strTI & " (FlATTOPDF)"
    arrData(8) = AR_2
    arrData(9) = END_OBJECT
    arrData(10) = "4" & BEG_OBJECT
    arrData(11) = AR_1
    arrData(12) = "/Type /Font"
    arrData(13) = "/Subtype /Type1"
    arrData(14) = "/Name /F1"
    arrData(15) = "/Encoding 5" & AR_OR
    arrData(16) = "/BaseFont /" & strFontName
    arrData(17) = AR_2
    arrData(18) = END_OBJECT
    arrData(19) = "5" & BEG_OBJECT
    arrData(20) = AR_1
    arrData(21) = "/Type /Encoding"
    arrData(22) = "/BaseEncoding /WinAnsiEncoding"
    arrData(23) = AR_2
    arrData(24) = END_OBJECT
    arrData(25) = "6" & BEG_OBJECT
    arrData(26) = AR_1
    arrData(27) = "/Font " & AR_1 & " /F1 4" & AR_OR & "   " & AR_2 & "  /ProcSet [ /" & Chr(80) & Chr(68) & Chr(70) & "  /Text ]"
    arrData(28) = AR_2
    arrData(29) = END_OBJECT
End Sub

Private Sub FillTmpArrayFromTextFile()
    Dim objFile As FileServer
    Dim strRow As String
    Dim i As Long
    Dim sPgBreak As String
    Dim j As Integer
    Dim nLeft As Integer
    Dim m As Integer
    
    sPgBreak = Trim(PageBreakBefore)
    Set objFile = New FileServer

    If Not objFile.PrepareFile(strSourceFile) Then
        Exit Sub
    End If
    i = 0
    Do While Not objFile.EndOfFile
        strRow = objFile.GetLine
        
        '#Special code. Remove the first 3 empty lines
        If sPgBreak <> "" And i = 0 Then
            m = 0
            Do While m <= 3
                If Trim(strRow) = "" Then
                    If Not objFile.EndOfFile Then
                        strRow = objFile.GetLine
                    End If
                Else
                    Exit Do
                End If
                m = m + 1
            Loop
        End If
        
        If sPgBreak <> "" And i > 10 Then
            If InStr(strRow, sPgBreak) > 0 Then
                nLeft = intRPP - (i Mod intRPP)
                For j = 1 To nLeft
                    i = i + 1
                    ReDim Preserve arrTmp(1 To i)
                    arrTmp(i) = " "
                Next
            End If
        End If
        i = i + 1
        ReDim Preserve arrTmp(1 To i)
        'david 08/10/2007  #560304
        'need to put a \ in front of \, (, and ) character
        strRow = Replace(strRow, "\", "\\")
        strRow = Replace(strRow, "(", "\(")
        strRow = Replace(strRow, ")", "\)")
        ''''''''''''''''''''''''''
        arrTmp(i) = strRow
    Loop
    intPGCount = Int(i / intRPP) + 1
    objFile.CloseFile

End Sub

Private Sub FillData()
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim h As Long
    
    arcount = 29
    For i = 1 To intPGCount
       intObject = intObject + 1
       strPage = strPage & " " & intObject & AR_OR
       h = intObject + 1
       intLen = Len(CStr(intObject)) + Len(CStr(h))
       
       ReDim Preserve arrData(1 To arcount + 18)
       
       arrData(arcount + 1) = intObject & BEG_OBJECT
       arrData(arcount + 2) = AR_1
       arrData(arcount + 3) = "/Type /Page"
       arrData(arcount + 4) = "/Parent 3" & AR_OR
       arrData(arcount + 5) = "/Resources 6" & AR_OR
       intObject = intObject + 1
       arrData(arcount + 6) = "/Contents " & intObject & AR_OR
       arrData(arcount + 7) = AR_2
       arrData(arcount + 8) = END_OBJECT
       intOffSet = intLen + 86 + intOffSet
       Dim s As String
       s = "0000000000" & intOffSet
       ReDim Preserve arrXREF(1 To UBound(arrXREF) + 1)
       arrXREF(UBound(arrXREF)) = Mid(s, Len(s) - 9, Len(s)) & XREF_END_CHAR
       
       arrData(arcount + 9) = intObject & BEG_OBJECT
       arrData(arcount + 10) = AR_1
       intObject = intObject + 1
       arrData(arcount + 11) = "/Length " & intObject & AR_OR
       arrData(arcount + 12) = AR_2
       arrData(arcount + 13) = "stream"
       arrData(arcount + 14) = "BT"
       arrData(arcount + 15) = "/F1 " & intFontSize & " Tf"
       arrData(arcount + 16) = "1 0 0 1 50 802 Tm"
       arrData(arcount + 17) = "12 TL"
       arcount = arcount + 17
       k = 0
       For j = 1 To intRPP
            ReDim Preserve arrData(1 To arcount + j)
            If UBound(arrTmp) < ((i - 1) * intRPP + j) Then
                Exit For
            Else
                k = k + 1
                If Len(arrTmp((i - 1) * intRPP + j)) > intWidth - 9 Then
                   arrData(arcount + j) = "T* (" & arrTmp((i - 1) * intRPP + j) & ") Tj"
                Else
                  arrData(arcount + j) = "T* (" & arrTmp((i - 1) * intRPP + j) & Space(intWidth - Len(arrTmp((i - 1) * intRPP + j))) & ") Tj"
                End If
            End If
       Next
       arcount = arcount + k
          
       ReDim Preserve arrData(1 To arcount + 18)
       
       arrData(arcount + 1) = "ET"
       arrData(arcount + 2) = "endstream"
       arrData(arcount + 3) = END_OBJECT
       
       
       rows = k * 90 + 45
       h = intObject - 1
       intLen = Len(CStr(intObject)) + Len(CStr(h))
       intOffSet = intLen + 57 + intOffSet + rows
       s = "0000000000" & intOffSet
       
       ReDim Preserve arrXREF(1 To UBound(arrXREF) + 1)
       arrXREF(UBound(arrXREF)) = Mid(s, Len(s) - 9, Len(s)) & XREF_END_CHAR
       
       arrData(arcount + 4) = intObject & BEG_OBJECT
       arrData(arcount + 5) = rows
       arrData(arcount + 6) = END_OBJECT
        
         
        intLen = Len(CStr(intObject)) + Len(CStr(rows))
        intOffSet = intLen + 18 + intOffSet
        
          s = "0000000000" & intOffSet
       ReDim Preserve arrXREF(1 To UBound(arrXREF) + 1)
       arrXREF(UBound(arrXREF)) = Mid(s, Len(s) - 9, Len(s)) & XREF_END_CHAR
       arcount = arcount + 6
    Next
End Sub

Private Sub FillFotter()
   Dim s As String
   ReDim Preserve arrData(1 To arcount + 15)
   arrData(arcount + 1) = "2" & BEG_OBJECT
   arrData(arcount + 2) = AR_1
   arrData(arcount + 3) = "/Type /Catalog"
   arrData(arcount + 4) = "/Pages 3" & AR_OR
   arrData(arcount + 5) = "/PageLayout /OneColumn"
   arrData(arcount + 6) = AR_2
   arrData(arcount + 7) = END_OBJECT

    UpdatexresWithLastNumber
   
    arrData(arcount + 8) = "3" & BEG_OBJECT
    arrData(arcount + 9) = AR_1
    arrData(arcount + 10) = "/Type /Pages"
    arrData(arcount + 11) = "/Count " & intPGCount
    arrData(arcount + 12) = "/MediaBox [ 0 0 595 842 ]"
    arrData(arcount + 13) = "/Kids [" & strPage & " ]"
    arrData(arcount + 14) = AR_2
    arrData(arcount + 15) = END_OBJECT
    arcount = arcount + 15
    intOffSet = intOffSet + 79
   
    s = "0000000000" & intOffSet
    s = Mid(s, Len(s) - 9, Len(s)) & XREF_END_CHAR
      
    arrXREF(6) = s

    ReDim Preserve arrXREF(1 To UBound(arrXREF) + 1)
    arrXREF(UBound(arrXREF)) = "trailer"
   
    ReDim Preserve arrXREF(1 To UBound(arrXREF) + 1)
    arrXREF(UBound(arrXREF)) = AR_1
      
    intObject = intObject + 1
    arrXREF(2) = "0" & intObject
    
    ReDim Preserve arrXREF(1 To UBound(arrXREF) + 1)
    arrXREF(UBound(arrXREF)) = "/Size " & intObject
    
    ReDim Preserve arrXREF(1 To UBound(arrXREF) + 1)
    arrXREF(UBound(arrXREF)) = "/Root 2" & AR_OR
    
    ReDim Preserve arrXREF(1 To UBound(arrXREF) + 1)
    arrXREF(UBound(arrXREF)) = "/Info 1" & AR_OR
    
    ReDim Preserve arrXREF(1 To UBound(arrXREF) + 1)
    arrXREF(UBound(arrXREF)) = AR_2
    
    ReDim Preserve arrXREF(1 To UBound(arrXREF) + 1)
    arrXREF(UBound(arrXREF)) = "startxref"
       
    intLen = Len(CStr(intPGCount)) + Len(CStr(strPage))
    intOffSet = intLen + 86 + intOffSet
    
    ReDim Preserve arrXREF(1 To UBound(arrXREF) + 1)
    arrXREF(UBound(arrXREF)) = intOffSet
    
    ReDim Preserve arrXREF(1 To UBound(arrXREF) + 1)
    arrXREF(UBound(arrXREF)) = "%%" & Chr(69) & Chr(79) + Chr(70)
    
    FillDataFromXRef
    
End Sub

Private Sub UpdatexresWithLastNumber()
    arrXREF(5) = arrXREF(UBound(arrXREF))
    arrXREF(UBound(arrXREF)) = ""
End Sub

Private Sub FillDataFromXRef()
    Dim i As Long
    For i = 1 To UBound(arrXREF)
       If arrXREF(i) <> "" Then
          ReDim Preserve arrData(1 To UBound(arrData) + 1)
          arrData(UBound(arrData)) = arrXREF(i)
       End If
    Next
End Sub

Private Sub WriteToReport()
  Dim strRow As String
  Dim i As Long
  Dim nFile As Integer
 
  nFile = FreeFile
  Open strTargetFiler For Output As #nFile
  For i = 1 To UBound(arrData)
      'david 12/01/2004  call RTrim
      Print #nFile, RTrim(arrData(i))
  Next
  Close #nFile
End Sub

Public Property Get Font() As Variant
Font = strFontName
End Property

Public Property Let Font(ByVal vNewValue As Variant)
 strFontName = vNewValue
End Property

Public Property Get FontSize() As Variant

End Property

Public Property Let FontSize(ByVal vNewValue As Variant)
intFontSize = vNewValue
End Property
