Attribute VB_Name = "modPrint"
Option Explicit

' Copyright (c) 1999 FACTOR, A Division of W.R.Hess Company
'
'Print Rotines for CC Processor
'Programmer: Weigong Jiang

    Private HEADER_PT As String
    
    Private sHEADER_HEADER As String
    Private sPRO_COLS_HEADER As String
    Private sERR_COLS_HEADER As String
   
    Private Const BOTTOM_MARGIN = 800
    Private Const TOP_MARGIN = 800
    
    Private Const STUB_FONT = "Arial"
    Private Const STUB_FONTSIZE = 12
    
    Private Const PRINT_FONTNAME = "Courier New"     '"Courier 17*1"
    Private Const PRINT_FONTSIZE = 9
    
    Private sStoreFont As String
    Private nLeftMargin As Integer
    Private nTextHeight As Integer
    Private nPrevPage As Integer
    Private nCharPerLine As Integer
    
    Private Const CHAR_PERLINE_132 = 132
    Private Const CHAR_PERLINE_110 = 104 '110
    
    '**********************************************
    Private Const szFACTOR_INI As String = "FACTOR.INI" 'application INI filename
    Private Const MAX_STRING_LENGTH As Integer = 255 'used with fixed length strings - normally with windows api calls
    
    Private ary_PROCESS_REPORT() As String
    Private INDEX_PROCESS_REPORT As Integer
   
    'david 02/15/2001
    'allow to override the ReportID, it was always using App.Title before
    Private sReportID As String
    
Public Sub subSetReportID(sID As String)
    sReportID = Trim(sID)
End Sub

Private Function fnGetTitle(sReportType As String, _
                           Optional nPageNo, _
                           Optional vPrint) As String
    Dim sLine1 As String
    Dim sLine2 As String
    Dim sLine3 As String
    Dim sLine4 As String
    Dim sLine5 As String
    Dim nPos As Integer
    Dim sTemp As String
    
    Dim sModName As String
    Dim sCompanyName As String
    Dim sReportName As String
    Dim sRunDate As String
    Dim sRunTime As String
    Dim sPage As String
    Dim nLeft As Integer 'starting point of report Name
    Dim nMax As Integer
    
    Dim n1 As Integer
    Dim n2 As Integer
    
    If Not IsMissing(nPageNo) Then
        sPage = fnFormatPage(nPageNo)
    Else
        sPage = ""
    End If
    sModName = App.Title '"CMFARMS" 'CC_INFO.Program_ID
    If sReportID <> "" Then
        sModName = sReportID
    End If
    sCompanyName = fnGetCompany 'CC_INFO.Company_Name
    
    sReportName = sReportType '& " REPORT"
'    If IsMissing(vOldReport) Then
'        sRunDate = "RUN DATE " & CStr(Date) 'CC_INFO.RunDate '& CStr(Date)
'        sRunTime = "RUN TIME " & Format(Now, "hh:mm AMPM") 'CC_INFO.RunTime 'Format(Now, "hh:mm AMPM")
'    Else
'        sRunDate = "PRINT DATE " & CStr(Date) 'CC_INFO.OldRunDate '& CStr(Date)
'        sRunTime = "PRINT TIME " & Format(Now, "hh:mm AMPM") 'CC_INFO.OldRunTime 'Format(Now, "hh:mm AMPM")
'    End If
    sRunDate = "DATE " & CStr(Date) 'CC_INFO.RunDate '& CStr(Date)
    sRunTime = "TIME " & Format(Now, "hh:mm AMPM") 'CC_INFO.RunTime 'Format(Now, "hh:mm AMPM")
    
    nLeft = nCharPerLine - Len(sModName) - Len(sPage)
        
    sLine1 = sModName & fnTranc(sCompanyName, nLeft, vbCenter) & sPage
    nLeft = nCharPerLine
    sLine2 = fnTranc(sRunDate, nLeft, vbLeftJustify)
    sLine3 = fnTranc(sRunTime, nLeft, vbLeftJustify)
    nPos = InStr(sReportName, vbCrLf)
    If nPos > 0 Then
        sLine4 = fnTranc(Left(sReportName, nPos - 1), nLeft, vbCenter)
        sLine5 = fnTranc(Right(sReportName, Len(sReportName) - nPos - 1), nLeft, vbCenter)
    Else
        sLine4 = fnTranc(sReportName, nLeft, vbCenter)
    End If
    
    If Not IsMissing(vPrint) Then
        Printer.CurrentX = nLeftMargin
        Printer.Print sLine1
        Printer.CurrentX = nLeftMargin
        Printer.Print sLine2
        Printer.CurrentX = nLeftMargin
        Printer.Print sLine3
        
        Printer.CurrentX = nLeftMargin
        Printer.Print sLine4
        If nPos > 0 Then
            Printer.CurrentX = nLeftMargin
            Printer.Print sLine5
        End If
        Printer.CurrentX = nLeftMargin
        Printer.Print ""
        
        Printer.CurrentX = nLeftMargin
        Printer.Print String(nCharPerLine, "=")
        Printer.CurrentX = nLeftMargin
        Printer.Print ""
        
        If Trim(HEADER_PT) <> "" Then
            n1 = 1
            Do
                n2 = InStr(n1, HEADER_PT, vbCrLf)
                If n2 > n1 Then
                    Printer.CurrentX = nLeftMargin
                    Printer.Print Mid(HEADER_PT, n1, n2 - n1)
                    n1 = n2 + 2
                End If
            Loop Until n2 = 0
            If n1 <= Len(HEADER_PT) Then
                Printer.CurrentX = nLeftMargin
                Printer.Print Mid(HEADER_PT, n1, Len(HEADER_PT) - n1 + 1)
            End If
            Printer.CurrentX = nLeftMargin
            Printer.Print ""
        End If
    End If
    fnGetTitle = sLine1 & vbCrLf & sLine2 & vbCrLf & sLine3 & vbCrLf & String(nCharPerLine, "-") & vbCrLf
End Function

Private Function fnGetCompany() As String
    Const sMod = "fnGetCompany"
    
    Dim sSql As String
    Dim rsTemp As Recordset
    
    sSql = "select con_name from co_company_name"
    Set rsTemp = fnOpenRecord(sSql, sMod, "")
    If Not rsTemp Is Nothing Then
        If rsTemp.RecordCount > 0 Then
            If IsNull(rsTemp!con_name) Then
                fnGetCompany = ""
            Else
                fnGetCompany = Trim(rsTemp!con_name)
            End If
        End If
    End If
End Function

Public Function fnOpenRecord(strSQL As String, _
                              Optional vCaller As Variant, _
                              Optional vMsg As Variant, _
                              Optional vDB As Variant) As Recordset
    Const SUB_NAME = "fnOpenRecord"
    ' Get records from the given SQL statement
    Dim objDB As DataBase
    Dim rsTemp As Recordset

    If IsMissing(vDB) Then
        Set objDB = t_dbMainDatabase
    Else
        Set objDB = vDB
    End If
    On Error GoTo SQLError
    If objDB Is t_dbMainDatabase Then
        Set rsTemp = objDB.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    Else
        Set rsTemp = objDB.OpenRecordset(strSQL, dbOpenSnapshot)
    End If
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveLast
        rsTemp.MoveFirst
    End If
    Set fnOpenRecord = rsTemp

    On Error GoTo 0
    Exit Function
SQLError:
    Set fnOpenRecord = Nothing
    Dim bShow As Boolean
    bShow = Not IsMissing(vMsg)
    If IsMissing(vCaller) Then
        tfnErrHandler SUB_NAME, strSQL, bShow
    Else
        tfnErrHandler SUB_NAME & "," & CStr(vCaller), strSQL, bShow
    End If
End Function


Public Function fnExecuteSQL(strSQL As String, _
                             Optional vCaller As Variant, _
                             Optional vMsg As Variant, _
                             Optional vDB As Variant) As Boolean
    Const SUB_NAME = "fnExecuteSQL"
    Dim objDB As DataBase
    
    If IsMissing(vDB) Then
        Set objDB = t_dbMainDatabase
    Else
        Set objDB = vDB
    End If
    On Error GoTo errExecute
    If objDB Is t_dbMainDatabase Then
        objDB.ExecuteSQL strSQL
    Else
        objDB.Execute strSQL
    End If
    fnExecuteSQL = True

    On Error GoTo 0
    Exit Function

errExecute:
    fnExecuteSQL = False
    Dim bShow As Boolean
    bShow = Not IsMissing(vMsg)
    If IsMissing(vCaller) Then
        tfnErrHandler SUB_NAME, strSQL, bShow
    Else
        tfnErrHandler SUB_NAME & "," & CStr(vCaller), strSQL, bShow
    End If
End Function


Private Function fnGetUbound(AryIn() As String) As Long
    On Error GoTo ErrorTrap
    fnGetUbound = UBound(AryIn)
    Exit Function
ErrorTrap:
    fnGetUbound = -1
    err.Clear
    On Error GoTo 0
End Function
Public Function fnSendToPrinter(AryIn() As String, _
                                Print_type As String) As Boolean
    Dim nPageNumber As Integer
    Dim lUBound As Long
    Dim i As Long
    
    '----BODY------------------------
    On Error Resume Next
    
    nPageNumber = 1
    
    lUBound = fnGetUbound(AryIn)
    If lUBound < 0 Then
        Exit Function
    End If
            
    Printer.CurrentX = nLeftMargin
    Printer.CurrentY = TOP_MARGIN

    fnGetTitle Print_type, nPageNumber, "Print"
    i = 0
    
    Do While i <= lUBound
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
            
            fnGetTitle Print_type, nPageNumber, "Print"
            If nPageNumber > 1 Then
                 If InStr(Print_type, "PROC") > 0 Then
                     Printer.CurrentX = nLeftMargin
                     Printer.Print sPRO_COLS_HEADER
                 Else
                     Printer.CurrentX = nLeftMargin
                     Printer.Print sERR_COLS_HEADER
                 End If
            End If
        Else
            Printer.CurrentX = nLeftMargin
            Printer.Print AryIn(i)
            i = i + 1
        End If
        
    Loop
    '---- END -------------------------
    subPrinterEndDocument
    
    fnSendToPrinter = True
    
End Function

Private Function fnFormatPage(ByVal sPage As String) As String
    
    sPage = "PAGE " & Trim(sPage)
    
    If Len(sPage) < 8 Then
        fnFormatPage = sPage & Space(8 - Len(sPage))
    Else
        fnFormatPage = Left(sPage, 8)
    End If
End Function
Public Function fnStartPrint(Print_type As String) As Boolean

    fnStartPrint = fnSendToPrinter(ary_PROCESS_REPORT, Print_type)
    
End Function

Public Sub subAddLine(s As String)
    INDEX_PROCESS_REPORT = INDEX_PROCESS_REPORT + 1
    ReDim Preserve ary_PROCESS_REPORT(INDEX_PROCESS_REPORT) As String
    ary_PROCESS_REPORT(INDEX_PROCESS_REPORT) = s
End Sub



Public Function fnTranc(ByVal str As String, _
                        ByVal nLen As Integer, _
                        nAlign As Integer) As String
    Dim nStrLen As Integer
    Dim n As Integer
    
    nStrLen = Len(str)
    If nStrLen > nLen Then
        str = Left(str, nLen)
        nStrLen = nLen
    End If
    
    Select Case nAlign
       Case vbLeftJustify
            str = str & Space(nLen - nStrLen)
       Case vbRightJustify
            str = Space(nLen - nStrLen) & str
       Case vbCenter
            n = CInt((nLen - nStrLen) / 2)
            str = Space(n) & str
            str = str & Space(nLen - Len(str))
    End Select
    
    fnTranc = str
End Function

'Set Up printer object before each print
Public Function fnSetupPrinter(ByVal nOrientation As Integer) As Boolean
    Dim n As Integer
    Const SUB_NAME = "subSetupPrinter"
    
    On Error GoTo errSetup
    sStoreFont = Printer.FontName
    
    Printer.Orientation = nOrientation
    
    'Next three Lines look like stupid codes, but they are necessary
    'They will help us to set the set font size
    
    Printer.FontName = STUB_FONT
    Printer.FontSize = STUB_FONTSIZE
    Printer.Print " "
    
    Printer.FontName = PRINT_FONTNAME
    Printer.FontSize = PRINT_FONTSIZE
    
    
    If nOrientation = vbPRORPortrait Then
        nCharPerLine = CHAR_PERLINE_110
    Else
        nCharPerLine = CHAR_PERLINE_132
    End If
    nLeftMargin = (Printer.ScaleWidth - Printer.TextWidth(Space(nCharPerLine))) / 2
    nTextHeight = Printer.ScaleHeight - BOTTOM_MARGIN
    fnSetupPrinter = True
    Exit Function
    
errSetup:
    Dim sErrMsg As String
    fnSetupPrinter = False

     sErrMsg = "Due to the following error, the system Cannot print the report." & vbCrLf
'     sErrMsg = sErrMsg & "However the report will be write to the log file '" & sLogFile & "':" & vbCrLf
     sErrMsg = sErrMsg & vbCrLf & "Error Number = " & err.Number & vbCrLf & "Error Message = " & err.Description
     MsgBox sErrMsg, vbExclamation
 
End Function

Private Sub subPrinterEndDocument()
    Printer.NewPage
    Printer.EndDoc
    Printer.FontName = sStoreFont
    Printer.Orientation = vbPRORPortrait
End Sub

Public Sub subSetTitle(sTitle As String)
    HEADER_PT = sTitle
    If nCharPerLine <= 0 Then
        nCharPerLine = CHAR_PERLINE_110
    End If
End Sub

Public Function fnSendToFile(AryIn() As String, _
                             sReportTitle As String, _
                             sPathFileName As String, _
                             Optional nOrientation As Integer = vbPRORPortrait) As Boolean
    
    Const SUB_NAME As String = "fnSendToFile"
    Const nDefaultLineHeight As Integer = 192
    
    Dim nPageNumber As Integer
    Dim nLineCount As Integer
    Dim nLinePerPage As Integer
    Dim lUBound As Long
    Dim i As Long
    
    Dim nptrFile As Integer
    
    On Error Resume Next
    Printer.KillDoc
    Printer.EndDoc
    
    On Error GoTo errTrap
    
    nPageNumber = 1
    nLineCount = 0
    
    lUBound = fnGetUbound(AryIn)
    If lUBound < 0 Then
        Exit Function
    End If
    
    If fnSetupPrinter(nOrientation) Then
        nLinePerPage = nTextHeight \ Printer.TextHeight("q")
    Else
        nLinePerPage = 66
    End If
    
    nptrFile = FreeFile
    
    Open sPathFileName For Output As nptrFile
            
    If Not fnWriteTitleToFile(nptrFile, sReportTitle, nPageNumber, nLineCount) Then
        On Error Resume Next
        Close nptrFile
        Printer.KillDoc
        Printer.EndDoc
        Exit Function
    End If
    
    i = 0
    
    Do While i <= lUBound
        If nLineCount >= nLinePerPage Then
            'Printer.NewPage
            Print #nptrFile, ""
            Print #nptrFile, ""
            Print #nptrFile, ""
            
            nLineCount = 0
            nPageNumber = nPageNumber + 1
            
            If Not fnWriteTitleToFile(nptrFile, sReportTitle, nPageNumber, nLineCount) Then
                On Error Resume Next
                Close nptrFile
                Printer.KillDoc
                Printer.EndDoc
                Exit Function
            End If
        Else
            Print #nptrFile, AryIn(i)
            nLineCount = nLineCount + 1
            i = i + 1
        End If
        
    Loop
    
    fnSendToFile = True
    
    Close nptrFile
    
    On Error Resume Next
    Printer.KillDoc
    Printer.EndDoc
    
    Exit Function

errTrap:
    MsgBox "An error has occurred in function " + SUB_NAME + "()." + vbCrLf + vbCrLf _
        + "Error Code: " & err.Number & vbCrLf & "Error Desc: " & err.Description _
        + vbCrLf + vbCrLf + "Please report this to support."
End Function

Private Function fnWriteTitleToFile(nptrFile As Integer, sReportType As String, _
                           Optional nPageNo, Optional nLineCount As Integer) As Boolean
    
    Const SUB_NAME As String = "fnWriteTitleToFile"
    
    Dim sLine1 As String
    Dim sLine2 As String
    Dim sLine3 As String
    Dim sLine4 As String
    Dim sLine5 As String
    Dim nPos As Integer
    Dim sTemp As String
    
    Dim sModName As String
    Dim sCompanyName As String
    Dim sReportName As String
    Dim sRunDate As String
    Dim sRunTime As String
    Dim sPage As String
    Dim nLeft As Integer 'starting point of report Name
    Dim nMax As Integer
    
    Dim n1 As Integer
    Dim n2 As Integer
    
    On Error GoTo errTrap
    
    If Not IsMissing(nPageNo) Then
        sPage = fnFormatPage(nPageNo)
    Else
        sPage = ""
    End If
    sModName = App.Title
    If sReportID <> "" Then
        sModName = sReportID
    End If
    sCompanyName = fnGetCompany
    
    sReportName = sReportType
    sRunDate = "DATE " & CStr(Date)
    sRunTime = "TIME " & Format(Now, "hh:mm AMPM")
    
    nLeft = nCharPerLine - Len(sModName) - Len(sPage)
        
    sLine1 = sModName & fnTranc(sCompanyName, nLeft, vbCenter) & sPage
    nLeft = nCharPerLine
    sLine2 = fnTranc(sRunDate, nLeft, vbLeftJustify)
    sLine3 = fnTranc(sRunTime, nLeft, vbLeftJustify)
    nPos = InStr(sReportName, vbCrLf)
    If nPos > 0 Then
        sLine4 = fnTranc(Left(sReportName, nPos - 1), nLeft, vbCenter)
        sLine5 = fnTranc(Right(sReportName, Len(sReportName) - nPos - 1), nLeft, vbCenter)
    Else
        sLine4 = fnTranc(sReportName, nLeft, vbCenter)
    End If
    
    Print #nptrFile, sLine1
    Print #nptrFile, sLine2
    Print #nptrFile, sLine3
    Print #nptrFile, sLine4
    nLineCount = nLineCount + 4
    
    If nPos > 0 Then
        Print #nptrFile, sLine5
        nLineCount = nLineCount + 1
    End If
    Print #nptrFile, ""
    nLineCount = nLineCount + 1
    
    Print #nptrFile, String(nCharPerLine, "=")
    Print #nptrFile, ""
    nLineCount = nLineCount + 2
    
    If Trim(HEADER_PT) <> "" Then
        n1 = 1
        Do
            n2 = InStr(n1, HEADER_PT, vbCrLf)
            If n2 > n1 Then
                Print #nptrFile, Mid(HEADER_PT, n1, n2 - n1)
                nLineCount = nLineCount + 1
                n1 = n2 + 2
            End If
        Loop Until n2 = 0
        If n1 <= Len(HEADER_PT) Then
            Print #nptrFile, Mid(HEADER_PT, n1, Len(HEADER_PT) - n1 + 1)
            nLineCount = nLineCount + 1
        End If
        Print #nptrFile, ""
        nLineCount = nLineCount + 1
    End If
    
    fnWriteTitleToFile = True
    
    Exit Function

errTrap:
    MsgBox "An error has occurred in function " + SUB_NAME + "()." + vbCrLf + vbCrLf _
        + "Error Code: " & err.Number & vbCrLf & "Error Desc: " & err.Description _
        + vbCrLf + vbCrLf + "Please report this to support."
End Function


