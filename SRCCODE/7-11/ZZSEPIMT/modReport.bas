Attribute VB_Name = "modPrintReport"
' Copyright (c) 2001 FACTOR, A Division of W.R.Hess Company
'Programmer: Weigong Jiang

Option Explicit
    
    Public Const BOTTOM_MARGIN = 800
    Private Const TOP_MARGIN = 800
    
    Public Const STUB_FONT = "Arial"
    Public Const STUB_FONTSIZE = 12
    
    Public Const PRINT_FONTNAME = "Courier New"     '"Courier 17*1"
    
    Public sStoreFont As String
    Public nLeftMargin As Integer
    Public nTextHeight As Integer
    Public nPrevPage As Integer
     
    Public Const PRINT_FONTSIZE = 9
    Public Const CHAR_PER_LINE = 132
    
    '**********************************************
    'Uncomment out next three lines if constant.bas not enclude
    'Private Const LEFT_JUSTIFY = 0
    'Private Const RIGHT_JUSTIFY = 1
    Public Const CENTER_JUSTIFY = 2
    '***********************************************
    
    Public ary_REPORT() As String
    Public INDEX_REPORT As Integer
    Private m_RunDate As String
    Private m_RunTime As String
    



Public Sub subLineFeed(s As String)
    INDEX_REPORT = INDEX_REPORT + 1
    ReDim Preserve ary_REPORT(INDEX_REPORT) As String
    ary_REPORT(INDEX_REPORT) = s
End Sub
    
Public Function fnSendToPrinter(sModName As String, _
                                sReportName As String, _
                                sExtraLine As String, _
                                sLineHeader As String) As Boolean
    
    Dim nPageNumber As Integer
    Dim lUBound As Long
    Dim i As Long
    
    '----Set UP-----------------
    If Not fnSetupPrinter Then
        Exit Function
    End If
    '----BODY------------------------
    On Error Resume Next
    
    nPageNumber = 1
    
    
    If INDEX_REPORT < 0 Then
        Exit Function
    End If
                
    
    subSetRunDateTime
    
    Printer.CurrentX = nLeftMargin
    Printer.CurrentY = TOP_MARGIN
    
    subPrintTitle sModName, sReportName, sExtraLine, sLineHeader, nPageNumber
    
    i = 0
    
    Do While i <= INDEX_REPORT
        If Printer.CurrentY >= nTextHeight Then
           Printer.NewPage
           nPageNumber = nPageNumber + 1
           '*****************************
           Printer.FontName = STUB_FONT
           Printer.FontSize = STUB_FONTSIZE
           Printer.Print " "
            
           Printer.FontName = PRINT_FONTNAME
           Printer.FontSize = PRINT_FONTSIZE
            
           '************************
           Printer.CurrentY = TOP_MARGIN
           Printer.CurrentX = nLeftMargin

           subPrintTitle sModName, sReportName, sExtraLine, sLineHeader, nPageNumber
        Else
            Printer.CurrentX = nLeftMargin
            Printer.Print ary_REPORT(i)
            i = i + 1
        End If
        
    Loop
    '---- END -------------------------
    subPrinterEndDocument
    
    ReDim ary_REPORT(0)
    INDEX_REPORT = -1
    
    fnSendToPrinter = True
    
End Function

    
Private Sub subSetRunDateTime()
    m_RunDate = "RUN DATE: " & Format(Date, "mm/dd/yyyy")
    m_RunTime = "RUN TIME: " & Format(Time, "hh:mm:ss AMPM")
End Sub
    
    
Private Sub subPrintTitle(sModName As String, _
                           sReportName As String, _
                           sExtraLine As String, _
                           sLineHeader As String, _
                           nPageNo As Integer)
    
    Static rs_CompanyInfo As Recordset
    Dim sTempLine As String
    Dim strSQL As String
        
    Dim sLine1 As String
    Dim sLine2 As String
    Dim sLine3 As String
    
    Dim sCompanyName As String
    Dim sPage As String
    Dim nLeft As Integer 'starting point of report Name
    Dim nMax As Integer
    
    If rs_CompanyInfo Is Nothing Then
        strSQL = "SELECT * FROM co_company_name"
        If frmZZSEPIMT.GetRecordSet(rs_CompanyInfo, strSQL) <= 0 Then
            Exit Sub
        End If
    End If
  
    sPage = fnFormatPage(nPageNo)
    
    sCompanyName = fnGetField(rs_CompanyInfo!con_name)
    
    nMax = Len(sModName)
    If Len(m_RunDate) > nMax Then
        nMax = Len(m_RunDate)
    End If
    If Len(m_RunTime) > nMax Then
        nMax = Len(m_RunTime)
    End If
    nLeft = CHAR_PER_LINE - nMax - Len(sPage)
        
    sLine1 = sModName & Space(nMax - Len(sModName)) & fnTranc(sCompanyName, nLeft, CENTER_JUSTIFY) & sPage
    sLine2 = m_RunDate & Space(nMax - Len(m_RunDate)) & fnTranc(sReportName, nLeft, CENTER_JUSTIFY)
    sLine3 = m_RunTime & Space(nMax - Len(m_RunTime))
    
    Printer.CurrentX = nLeftMargin
    Printer.Print sLine1
    Printer.CurrentX = nLeftMargin
    Printer.Print sLine2
    Printer.CurrentX = nLeftMargin
    Printer.Print sLine3
    
    If Trim(sExtraLine) <> "" Then
        Printer.Print ""
        Printer.CurrentX = nLeftMargin
        Printer.Print sExtraLine
    End If
    
    If Trim(sLineHeader) <> "" Then
        Printer.Print ""
        Printer.CurrentX = nLeftMargin
        Printer.Print sLineHeader
        Printer.CurrentX = nLeftMargin
        Printer.Print String(CHAR_PER_LINE, "-")
    
    Else
        Printer.Print ""
    End If
    
    
End Sub
    

Private Function fnFormatPage(ByVal sPage As String) As String
    
    sPage = "PAGE " & Trim(sPage)
    
    If Len(sPage) < 8 Then
        fnFormatPage = sPage & Space(8 - Len(sPage))
    Else
        fnFormatPage = Left(sPage, 8)
    End If
End Function

Public Function fnTranc(ByVal str As String, ByVal nLen As Integer, nAlign As Integer) As String
    Dim nStrLen As Integer
    Dim n As Integer
    
    str = Trim(str)
    nStrLen = Len(str)
    If nStrLen > nLen Then
        str = Left(str, nLen)
        nStrLen = nLen
    End If
    
    Select Case nAlign
       Case LEFT_JUSTIFY
            str = str & Space(nLen - nStrLen)
       Case RIGHT_JUSTIFY
            str = Space(nLen - nStrLen) & str
       Case CENTER_JUSTIFY
            n = CInt((nLen - nStrLen) / 2)
           
            str = Space(n) & str
            str = str & Space(nLen - Len(str))
    End Select
    
    fnTranc = str
End Function

'Set Up printer object before each print
Private Function fnSetupPrinter() As Boolean
    Dim n As Integer
    Dim sErrMsg As String
    
    Const SUB_NAME = "subSetupPrinter"
    
    On Error GoTo errSetup
    sStoreFont = Printer.FontName
    
    Printer.Orientation = vbPRORLandscape
    
    'Next three Lines look like stupid codes, but they are necessary
    'They will help us to set the set font size
    
    Printer.FontName = STUB_FONT
    Printer.FontSize = STUB_FONTSIZE
    Printer.Print " "
    
    Printer.FontName = PRINT_FONTNAME
    Printer.FontSize = PRINT_FONTSIZE
    
    nLeftMargin = (Printer.ScaleWidth - Printer.TextWidth(Space(CHAR_PER_LINE))) / 2
    nTextHeight = Printer.ScaleHeight - BOTTOM_MARGIN
    fnSetupPrinter = True
    Exit Function
    
errSetup:
    fnSetupPrinter = False

     sErrMsg = "Due to the following error, the system Cannot print the report." & vbCrLf
     sErrMsg = sErrMsg & vbCrLf & "Error Number = " & Err.Number & vbCrLf & "Error Message = " & Err.Description
     MsgBox sErrMsg, vbExclamation
 
End Function

Private Sub subPrinterEndDocument()
    Printer.NewPage
    Printer.EndDoc
    Printer.FontName = sStoreFont
    Printer.Orientation = vbPRORPortrait
End Sub

Public Function fnGetField(x As Variant) As String
     If IsNull(x) Then
        fnGetField = szEMPTY
     Else
        fnGetField = Trim(x)
     End If
End Function
