Attribute VB_Name = "modZZFINVMT"
Option Explicit

Private Type RSINV_Header
    sRecId  As String
    nProfctr As Integer
    dtReportDate As Date
    nShiftNum As Integer
    lVendor As Long
    lInvNum As Long
    dtInvDate As Date
    sTerm As String
    sPayType As String
    sDraftNum As String 'may be blank
    dInvAmount As Double
End Type

Private Type RSINV_Detail
    sRecId As String
    nDetailLine As Integer
    sTypeCode   As String
    sItemCode As String
    dQuantity As Double
    sUOM As String
    dCost As Double
    sRetail As String 'may be blank
End Type

Private g_nProcessingFile As Integer
Private g_nErrorLogFile As Integer

Public dbLocal As Database
Const nDB_LOCAL As Integer = 1
Const nDB_REMOTE As Integer = 2

Private g_dTotalCost As Double
Private g_dTotalExtCost As Double
Private g_lHeaderCount As Long
Private g_bNeedValidHeader As Boolean
Private g_bNeedWriteHeader As Boolean
Private g_lDetailLineNbr As Long

Public Function fnProcessRSInvFile(sFileName As String) As Boolean
    Dim lTotalByte As Long
    Dim lReadByte As Long
    Dim bISOpen As Boolean
    Dim nFileNum As Integer
    Dim sLine As String
    Dim nLineCount As Long
    Dim sFirstLine As String
    Dim udtInvHeader As RSINV_Header
    Dim udtInvDetail As RSINV_Detail
    Dim sErrMsg As String
    
    On Error GoTo ERRORHANDLER
    
    'check how many header read
    g_lHeaderCount = 0
    lTotalByte = FileLen(sFileName)
    nFileNum = FreeFile()
    
    Open sFileName For Input As #nFileNum
    bISOpen = True
    
    subOpenAndClearLogFile
    
    If Not EOF(nFileNum) Then
        Line Input #nFileNum, sFirstLine
    Else
        sErrMsg = "The file " & sFileName & "  is empty"
        GoTo EXITHERE
    End If
    
    If Left(sFirstLine, 3) <> "HDR" Then
        sErrMsg = "The first three letter is not 'HDR' in line 1."
        GoTo EXITHERE
    Else
        
        If Not fnProcessHeaderLine(sFirstLine, udtInvHeader) Then
            sErrMsg = "The file contents is not correct in line 1"
            GoTo EXITHERE
        End If
        
        subWriteHeaderProcLog udtInvHeader
        subWriteHeaderErrLog udtInvHeader
    End If
    
    nLineCount = 1
    
    Do While Not EOF(nFileNum)
        Line Input #nFileNum, sLine
        
        nLineCount = nLineCount + 1
        
        If Left(sLine, 3) = "HDR" Then
            
            If Not fnProcessHeaderLine(sLine, udtInvHeader) Then
                sErrMsg = "Invalid contents in line " & nLineCount
                GoTo EXITHERE
            End If
            
            subWriteHeaderProcLog udtInvHeader
            subWriteHeaderErrLog udtInvHeader
            
        ElseIf Left(sLine, 3) = "DET" Then
        
            If Not fnProcessDetailLine(sLine, udtInvDetail) Then
                sErrMsg = "Invalid contents in line " & nLineCount
                GoTo EXITHERE
            End If
            
        Else
            sErrMsg = "Invalid contents in line " & nLineCount
            GoTo EXITHERE
        End If
        
        'one header have one correspond detail
        If udtInvDetail.sRecId = "DET" Then
            
            If Not fnInsertData(udtInvHeader, udtInvDetail) Then
                sErrMsg = "Error occurs when inserting data."
                GoTo EXITHERE
            End If
            
        End If
        
        'set it to empty after new header is accept
        udtInvDetail.sRecId = ""
    Loop
    
    subWriteSummary
    subCloseLogFile
    
    fnProcessRSInvFile = True
    Close #nFileNum
    Exit Function
    
EXITHERE:
    fnProcessRSInvFile = False
    Close #nFileNum
    subCloseLogFile
    Exit Function
    
ERRORHANDLER:

    If bISOpen Then
        Close #nFileNum
    End If
    
    subDisplayMsg Err.Description
End Function


Private Function fnProcessHeaderLine(sLine As String, udtInvHeader As RSINV_Header) As Boolean
    On Error GoTo EXITHERE
    
    If Len(sLine) < 70 Then
        Exit Function
    End If
    
    With udtInvHeader
        .sRecId = Mid$(sLine, 1, 3)
        .nProfctr = CInt(Mid$(sLine, 4, 5))
        .dtReportDate = CDate(Mid$(sLine, 9, 10))
        .nShiftNum = CInt(Mid$(sLine, 19, 5))
        .lVendor = CLng(Mid$(sLine, 24, 10))
        .lInvNum = CLng(Mid$(sLine, 34, 10))
        .dtInvDate = CDate(Mid$(sLine, 44, 10))
        .sTerm = Mid$(sLine, 54, 5)
        .sPayType = Mid$(sLine, 59, 1)
        .sDraftNum = Mid$(sLine, 60, 10)
        .dInvAmount = CDbl(Mid$(sLine, 70, 10))
    End With
    
    fnProcessHeaderLine = True
    Exit Function
EXITHERE:
    fnProcessHeaderLine = False
    subDisplayMsg Err.Description
End Function

Private Function fnProcessDetailLine(sLine As String, udtInvDetail As RSINV_Detail) As Boolean
    On Error GoTo EXITHERE
    
    If Len(sLine) < 34 Then
        Exit Function
    End If
    
    With udtInvDetail
        .sRecId = Mid$(sLine, 1, 3)
        .nDetailLine = CInt(Mid$(sLine, 4, 3))
        .sTypeCode = Mid$(sLine, 7, 1)
        .sItemCode = Mid$(sLine, 8, 10)
        .dQuantity = CDbl(Mid(sLine, 18, 12))
        .sUOM = Mid$(sLine, 30, 5)
        .dCost = CDbl(Mid$(sLine, 35, 12))
        .sRetail = Mid$(sLine, 47, 12)
    End With
    
    fnProcessDetailLine = True
    Exit Function
EXITHERE:
    fnProcessDetailLine = False
    subDisplayMsg Err.Description
End Function

Private Function fnInsertData(udtInvHeader As RSINV_Header, udtInvDetail As RSINV_Detail) As Boolean
    Dim strSQL As String
    
    subWriteDetailProcLog udtInvHeader, udtInvDetail
    
    If fnValidData(udtInvHeader, udtInvDetail) Then
        'insert header data
        If g_bNeedWriteHeader Then
            'delete old data in rs_b_hold_header for this vendor and invoice
            strSQL = "DELETE FROM rs_p_hold_header WHERE rsphh_vendor = " & udtInvHeader.lVendor
            strSQL = strSQL & " AND rsphh_invoice = " & udtInvHeader.lInvNum
            
            If Not fnExecuteSQL(strSQL, nDB_REMOTE, "fnInsertData") Then
                Exit Function
            End If
            
            'delete old data in rs_b_hold_detail for this vendor and invoice
            strSQL = "DELETE FROM rs_p_hold_detail WHERE rsphd_vendor = " & udtInvHeader.lVendor
            strSQL = strSQL & " AND rsphd_invoice = " & udtInvHeader.lInvNum
            
            If Not fnExecuteSQL(strSQL, nDB_REMOTE, "fnInsertData") Then
                Exit Function
            End If
            
            strSQL = "INSERT INTO rs_p_hold_header(rsphh_prft_ctr, rsphh_rpt_date, rsphh_shift, rsphh_vendor,"
            strSQL = strSQL & "rsphh_invoice, rsphh_inv_date, rsphh_std_term, rsphh_type, rsphh_draft_nbr, rsphh_invoice_amt)"
            strSQL = strSQL & " VALUES(" & udtInvHeader.nProfctr & "," & tfnDateString(udtInvHeader.dtReportDate, True) & ","
            strSQL = strSQL & udtInvHeader.nShiftNum & "," & udtInvHeader.lVendor & "," & udtInvHeader.lInvNum & ","
            strSQL = strSQL & tfnDateString(udtInvHeader.dtInvDate, True) & "," & tfnSQLString(udtInvHeader.sTerm) & ","
            strSQL = strSQL & tfnSQLString(udtInvHeader.sPayType) & "," & IIf(udtInvHeader.sDraftNum = "", "NULL", udtInvHeader.sDraftNum)
            strSQL = strSQL & tfnRound(udtInvHeader.dInvAmount, 3) & ")"
            g_bNeedWriteHeader = False
            g_lDetailLineNbr = 1
            
            If Not fnExecuteSQL(strSQL, nDB_REMOTE, "fnInsertData") Then
                Exit Function
            End If
            
        End If
        
        strSQL = "INSERT INTO rs_b_hold_detail(rsphd_vendor, rsphd_invoice, rsphd_line_nbr, rsphd_type, "
        strSQL = strSQL & " rsphd_code, rsphd_qty, rsphd_stock_unit, rsphd_cost, rsphd_retail)"
        strSQL = strSQL & " VALUES( " & udtInvHeader.lVendor & "," & udtInvHeader.lInvNum & ","
        strSQL = strSQL & g_lDetailLineNbr & "," & tfnSQLString(udtInvDetail.sTypeCode) & ","
        strSQL = strSQL & tfnSQLString(udtInvDetail.sItemCode) & "," & tfnRound(udtInvDetail.dQuantity, 3) & ","
        strSQL = strSQL & tfnSQLString(udtInvDetail.sUOM) & "," & tfnRound(udtInvDetail.dCost, 3) & ","
        strSQL = strSQL & tfnRound(udtInvDetail.sRetail, 3)
        
        If Not fnExecuteSQL(strSQL, nDB_REMOTE, "fnInsertData") Then
            Exit Function
        End If
        
        g_lDetailLineNbr = g_lDetailLineNbr + 1
        fnInsertData = True
    Else
        fnInsertData = False
        subDisplayMsg "Data is not correct in the flat file, please check error log file."
    End If
    
End Function

Private Function fnValidData(udtInvHeader As RSINV_Header, udtInvDetail As RSINV_Detail) As Boolean
    Dim sErrMsg As String
    Dim sItemDesc As String
    Dim sPayTerm As String
    
    fnValidData = True
    
    'valid this one first, because it item description is required in the log file
    sErrMsg = fnValidItemCode(udtInvHeader.lVendor, udtInvDetail.sItemCode, sItemDesc)
    
    If sErrMsg <> "" Then
        fnValidData = False
        subWriteDetailErrLog udtInvHeader, udtInvDetail, sItemDesc, sErrMsg
    End If
    
    If g_bNeedValidHeader Then
        sErrMsg = fnValidPrftCtr(udtInvHeader.nProfctr)
        
        If sErrMsg <> "" Then
            fnValidData = False
            subWriteDetailErrLog udtInvHeader, udtInvDetail, sItemDesc, sErrMsg
        End If
        
        sErrMsg = fnValidReportDate(udtInvHeader.nProfctr, udtInvHeader.dtReportDate)
        
        If sErrMsg <> "" Then
            fnValidData = False
            subWriteDetailErrLog udtInvHeader, udtInvDetail, sItemDesc, sErrMsg
        End If
        
        sErrMsg = fnValidShiftNum(udtInvHeader.nProfctr, udtInvHeader.nShiftNum, udtInvHeader.dtReportDate)
        
        If sErrMsg <> "" Then
            fnValidData = False
            subWriteDetailErrLog udtInvHeader, udtInvDetail, sItemDesc, sErrMsg
        End If
                
        sErrMsg = fnValidVendor(udtInvHeader.lVendor, sPayTerm)
        
        If sErrMsg <> "" Then
            fnValidData = False
            subWriteDetailErrLog udtInvHeader, udtInvDetail, sItemDesc, sErrMsg
        End If
          
        sErrMsg = fnValidInvNum(udtInvHeader.lVendor, udtInvHeader.lInvNum)
        
        If sErrMsg <> "" Then
            fnValidData = False
            subWriteDetailErrLog udtInvHeader, udtInvDetail, sItemDesc, sErrMsg
        End If
        
        sErrMsg = fnValidInvDate(udtInvHeader.dtInvDate, udtInvHeader.dtReportDate)
        
        If sErrMsg <> "" Then
            fnValidData = False
            subWriteDetailErrLog udtInvHeader, udtInvDetail, sItemDesc, sErrMsg
        End If
        
        sErrMsg = fnValidPayTerm(udtInvHeader.sTerm, sPayTerm)
        
        If sErrMsg <> "" Then
            fnValidData = False
            subWriteDetailErrLog udtInvHeader, udtInvDetail, sItemDesc, sErrMsg
        End If
        
        sErrMsg = fnValidPayType(udtInvHeader.sPayType)
        
        If sErrMsg <> "" Then
            fnValidData = False
            subWriteDetailErrLog udtInvHeader, udtInvDetail, sItemDesc, sErrMsg
        End If
                
        sErrMsg = fnValidDraftNum(udtInvHeader.sPayType, udtInvHeader.sDraftNum)
        
        If sErrMsg <> "" Then
            fnValidData = False
            subWriteDetailErrLog udtInvHeader, udtInvDetail, sItemDesc, sErrMsg
        End If
        
        sErrMsg = fnValidInvAmount(udtInvHeader.dInvAmount)
        
        If sErrMsg <> "" Then
            fnValidData = False
            subWriteDetailErrLog udtInvHeader, udtInvDetail, sItemDesc, sErrMsg
        End If
        
        g_bNeedValidHeader = False
    End If
    
    'valid detail
    sErrMsg = fnValidPurchaseType(udtInvDetail.sTypeCode)
        
    If sErrMsg <> "" Then
        fnValidData = False
        subWriteDetailErrLog udtInvHeader, udtInvDetail, sItemDesc, sErrMsg
    End If
    
    sErrMsg = fnValidPurchaseQty(udtInvDetail.dQuantity)
        
    If sErrMsg <> "" Then
        fnValidData = False
        subWriteDetailErrLog udtInvHeader, udtInvDetail, sItemDesc, sErrMsg
    End If
    
    sErrMsg = fnValidUOM(udtInvDetail.sUOM)
        
    If sErrMsg <> "" Then
        fnValidData = False
        subWriteDetailErrLog udtInvHeader, udtInvDetail, sItemDesc, sErrMsg
    End If
    
    sErrMsg = fnValidPurchaseCost(udtInvDetail.dCost)
        
    If sErrMsg <> "" Then
        fnValidData = False
        subWriteDetailErrLog udtInvHeader, udtInvDetail, sItemDesc, sErrMsg
    End If
    
    sErrMsg = fnValidRetailPrice(udtInvDetail.sRetail, udtInvDetail.sItemCode, udtInvHeader.lVendor, udtInvHeader.nProfctr, udtInvHeader.dtReportDate)
        
    If sErrMsg <> "" Then
        fnValidData = False
        subWriteDetailErrLog udtInvHeader, udtInvDetail, sItemDesc, sErrMsg
    End If
    
End Function

Public Sub subSetProgress(sngPercent As Single)
    
    If sngPercent > 0# Then
        
        If Not frmZZFINVMT.PbProgressBar.Visible Then
            frmZZFINVMT.PbProgressBar.Visible = True
        End If
    
    Else
        frmZZFINVMT.PbProgressBar.Visible = False
    End If
    
    frmZZFINVMT.PbProgressBar.Value = sngPercent
    frmZZFINVMT.PbProgressBar.Refresh
    
End Sub

Public Sub subDisplayMsg(sMsg As String)
    frmZZFINVMT.lstStatus.AddItem sMsg
End Sub

Private Sub subWriteHeaderErrLog(udtInvHeader As RSINV_Header)
    Dim sLine As String
    Dim sVendorName As String
    
    sVendorName = fnGetVendorName(udtInvHeader.lVendor)
    
    sLine = "Profit Center: " & CStr(udtInvHeader.nProfctr)
    Print #g_nErrorLogFile, sLine
    sLine = "Shift Number: " & CStr(udtInvHeader.nShiftNum)
    Print #g_nErrorLogFile, sLine
    sLine = "Vendor Number: " & CStr(udtInvHeader.lVendor) & Space(10) & "Vendor Name: " & sVendorName
    Print #g_nErrorLogFile, sLine
    sLine = String(100, "-")
    Print #g_nErrorLogFile, sLine
    Print #g_nErrorLogFile, vbCrLf
    sLine = "Date" & Space(5) & "Invoice #" & Space(2) & "Inv. Code" & Space(2) & "Description" & Space(10) & "Error Message"
    Print #g_nErrorLogFile, sLine
    sLine = String(100, "-")
    Print #g_nErrorLogFile, sLine
End Sub

Private Sub subWriteDetailErrLog(udtInvHeader As RSINV_Header, udtInvDetail As RSINV_Detail, sItemDesc As String, sErrMsg As String)
    Dim sLine As String
   
    sLine = CStr(Format(udtInvHeader.dtReportDate, "MM/DD/YY"))
    sLine = sLine & Space(1) & CStr(udtInvHeader.lInvNum)
    sLine = sLine & Space(11 - Len(CStr(udtInvHeader.lInvNum))) & udtInvDetail.sItemCode
    sLine = sLine & Space(11 - Len(udtInvDetail.sItemCode)) & sItemDesc
    sLine = sLine & Space(21 - Len(sItemDesc)) & sErrMsg
    Print #g_nErrorLogFile, sLine
End Sub

Private Sub subWriteHeaderProcLog(udtInvHeader As RSINV_Header)
    Dim sLine As String
    Dim sVendorName As String
    Dim sItemDesc As String
    Dim dExtCost As Double
    
    sVendorName = fnGetVendorName(udtInvHeader.lVendor)
    g_bNeedValidHeader = True 'valid header if header changed
    g_bNeedWriteHeader = True
    g_lHeaderCount = g_lHeaderCount + 1 'add header count if header changed
    
    If g_lHeaderCount > 1 Then
        subWriteSummary
    End If
    
    g_dTotalCost = 0#
    g_dTotalExtCost = 0#
    sLine = "Profit Center: " & udtInvHeader.nProfctr
    Print #g_nProcessingFile, sLine
    sLine = "Shift Number: " & udtInvHeader.nShiftNum
    Print #g_nProcessingFile, sLine
    sLine = "Vendor Number: " & udtInvHeader.lVendor & Space(10) & "Vendor Name: " & sVendorName
    Print #g_nProcessingFile, sLine
    sLine = "Date" & Space(5) & "Invoice #" & Space(2) & "Inv. Code" & Space(2) & "Description" & Space(10) & "Qty" & Space(6) & "Cost" & Space(7) & "Ext. Cost"
    Print #g_nProcessingFile, sLine
    sLine = String(100, "-")
    Print #g_nProcessingFile, sLine

End Sub

Private Sub subWriteDetailProcLog(udtInvHeader As RSINV_Header, udtInvDetail As RSINV_Detail)
    Dim sLine As String
    Dim sItemDesc As String
    Dim dExtCost As Double
    
    'write detail
    sLine = CStr(Format(udtInvHeader.dtReportDate, "MM/DD/YY"))
    sLine = sLine & Space(1) & CStr(udtInvHeader.lInvNum)
    sLine = sLine & Space(11 - Len(CStr(udtInvHeader.lInvNum))) & udtInvDetail.sItemCode
    sItemDesc = fnGetItemDesc(udtInvHeader.lVendor, udtInvDetail.sItemCode)
    sLine = sLine & Space(11 - Len(udtInvDetail.sItemCode)) & sItemDesc
    sLine = sLine & Space(21 - Len(sItemDesc)) & CStr(udtInvDetail.dQuantity)
    sLine = sLine & Space(9 - Len(CStr(udtInvDetail.dQuantity))) & CStr(udtInvDetail.dCost)
    dExtCost = udtInvDetail.dQuantity * udtInvDetail.dCost
    sLine = sLine & CStr(dExtCost)
    Print #g_nProcessingFile, sLine
    g_dTotalCost = g_dTotalCost + udtInvDetail.dCost
    g_dTotalExtCost = g_dTotalExtCost + dExtCost
    
End Sub

Private Sub subWriteSummary()
    Dim sLine As String
    
    sLine = "TOTAL" & Space(56) & CStr(g_dTotalCost) & Space(11 - Len(CStr(g_dTotalCost))) & CStr(g_dTotalExtCost)
    
    Print #g_nProcessingFile, sLine
End Sub

Private Sub subOpenAndClearLogFile()
    Dim sProcessLogFile As String
    Dim sErrorLogFile As String

    On Error Resume Next
    
    sProcessLogFile = App.Path & "\zzrinvpl.txt"
    sErrorLogFile = App.Path & "\zzrinver.txt"
    
    If fnFileExist(sProcessLogFile) Then
        Kill sProcessLogFile
    End If
    
    If fnFileExist(sErrorLogFile) Then
        Kill sErrorLogFile
    End If
    
    g_nProcessingFile = FreeFile()
    Open sProcessLogFile For Output As #g_nProcessingFile
    
    g_nErrorLogFile = FreeFile()
    Open sErrorLogFile For Output As #g_nErrorLogFile
    
End Sub

Private Function fnGetVendorName(lVendor As Long) As String
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    strSQL = "SELECT pn_name FROM p_altName WHERE pn_alt = " & lVendor
    
    If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, "fnGetVendorName") > 0 Then
        fnGetVendorName = rsTemp!pn_name & ""
    End If
    
    
End Function

Private Function fnGetItemDesc(lVendor As Long, sItemCode As String) As String
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    strSQL = "SELECT rsbi_desc FROM rs_b_item WHERE rsbi_vendor = " & lVendor
    strSQL = strSQL & " AND rsbi_code = " & tfnSQLString(sItemCode)
    
    If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, "fnGetItemDesc") > 0 Then
        fnGetItemDesc = rsTemp!rsbi_desc & ""
    End If
    
End Function

Private Sub subCloseLogFile()
    On Error Resume Next
    Close #g_nProcessingFile
    Close #g_nErrorLogFile
End Sub

Private Function fnFileExist(sFile As String) As Boolean
    On Error Resume Next
    
    fnFileExist = (Dir$(sFile) <> "")
    
End Function

' Get records from the given SQL statement
' nDB = 1 ---> Informax Database (remote)
'     = 2 ---> Access Database (local)
'This function will return a recordcount
Public Function fnGetRecord(rsTemp As Recordset, strSQL As String, Optional nDB As Integer, Optional sCalledFrom As String, Optional bShowError As Variant) As Long
    Const SUB_NAME = "fnGetRecord"

    On Error GoTo SQLError
    
    If IsMissing(nDB) Then
        nDB = nDB_REMOTE
    End If
    
    Select Case nDB
        Case nDB_LOCAL
            Set rsTemp = dbLocal.OpenRecordset(strSQL, dbOpenSnapshot)
        Case nDB_REMOTE
            Set rsTemp = t_dbMainDatabase.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    End Select
    
    If rsTemp.RecordCount > 0 Then  'the following code is used to get the correct
        rsTemp.MoveLast             'RecordCount of the RecordSet
        rsTemp.MoveFirst
    End If
    
    fnGetRecord = rsTemp.RecordCount
    Exit Function
    
SQLError:
    
    If IsMissing(sCalledFrom) Then
        sCalledFrom = ""
    End If
    
    If IsMissing(bShowError) Then
        bShowError = True
    End If

    tfnErrHandler SUB_NAME + "," + sCalledFrom, strSQL, bShowError
    fnGetRecord = -9999
End Function

Public Function fnExecuteSQL(szSQL As String, Optional nDB As Variant, _
                Optional sCalledFrom As Variant, Optional bShowError As Variant) As Boolean
                
    On Error GoTo SQLError
    
    If IsMissing(nDB) Then
        nDB = nDB_REMOTE
    End If
    
    Select Case nDB
        Case nDB_LOCAL 'local
            dbLocal.Execute szSQL
        Case nDB_REMOTE 'remote
            t_dbMainDatabase.ExecuteSQL szSQL
    End Select
    
    fnExecuteSQL = True
    Exit Function
    
SQLError:

    If IsMissing(sCalledFrom) Then
        sCalledFrom = ""
    End If
    
    If IsMissing(bShowError) Then
        bShowError = True
    End If

    tfnErrHandler "fnExecuteSQL, " & sCalledFrom, szSQL, bShowError
    On Error GoTo 0
    
End Function


Private Function fnValidItemCode(lVendor As Long, sItemCode As String, sItemDesc As String) As String
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    strSQL = "SELECT rsbi_desc FROM rs_b_item WHERE rsbi_vendor = " & lVendor
    strSQL = strSQL & " AND rsbi_code = " & tfnSQLString(sItemCode)
    
    If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, "fnGetItemDesc") < 0 Then
        fnValidItemCode = "Database Access Error."
    ElseIf rsTemp.RecordCount = 0 Then
        fnValidItemCode = "Item code does not exists."
    Else
        sItemDesc = rsTemp!rsbi_desc
    End If

End Function

Private Function fnValidPrftCtr(nPrftCtr As Integer) As String
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    strSQL = "SELECT prft_type FROM sys_prft_ctr WHERE "
    strSQL = strSQL & " prft_ctr = " & nPrftCtr
    
    If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, "fnValidPrftCtr") < 0 Then
        fnValidPrftCtr = "Database Access Error."
    ElseIf rsTemp.RecordCount = 0 Then
        fnValidPrftCtr = "Retail profit Center does not exists."
    Else
        
        If rsTemp!prft_type & "" <> "R" And rsTemp!prft_type & "" <> "B" Then
            fnValidPrftCtr = "This is not retail profit center."
        Else
            fnValidPrftCtr = ""
        End If
        
    End If

End Function

Private Function fnValidReportDate(nPrftCtr As Integer, dtReportDate As Date) As String
    Dim strSQL As String
    Dim dtLastProcDate As Date
    Dim rsTemp As Recordset
    Dim sMsg As String
    
    strSQL = "SELECT prft_posted_date from sys_prft_ctr WHERE "
    strSQL = strSQL & " prft_ctr = " & nPrftCtr
    strSQL = strSQL & " AND (prft_type = 'R' OR prft_type = 'B')"
    
    If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, "fnvalidReportDate") < 0 Then
        sMsg = "Database Access Error."
    ElseIf rsTemp.RecordCount = 0 Then
        sMsg = "Lastest process date is not available"
    Else
        dtLastProcDate = IIf(IsNull(rsTemp!prft_posted_date), Null, CDate(rsTemp!prft_posted_date))
        
        If Not IsNull(dtLastProcDate) Then
            
            If dtLastProcDate > dtReportDate Then
                sMsg = "Report date is earlier than last processed date."
            End If
            
        End If
        
    End If
    
    If sMsg = "" Then
        strSQL = "SELECT glp_status FROM gl_period WHERE " & dtReportDate
        strSQL = strSQL & " BETWEEN glp_beg_dt and glp_end_dt"
        
        If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, "fnvalidReportDate") < 0 Then
            sMsg = "Database Access Error."
        ElseIf rsTemp.RecordCount = 0 Then
            sMsg = "The report date is not in GL period."
        Else
            
            If rsTemp!glp_status <> "O" And rsTemp!glp_status <> "W" Then
                sMsg = "The Status for this period is not open."
            Else
                sMsg = ""
            End If
            
        End If
        
    End If
    
    fnValidReportDate = sMsg
    
End Function

Private Function fnValidShiftNum(nPrftCtr As Integer, nShiftNum As Integer, dtReportDate As Date) As String
    'implement later
    fnValidShiftNum = ""
End Function

Private Function fnValidVendor(lVendor As Long, sPayTerm As String) As String
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    strSQL = "SELECT pm_status, pm_std_disc_term FROM p_vendor WHERE "
    strSQL = strSQL & " pm_vendor = " & lVendor

    
    If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, "fnValidVendor") < 0 Then
        fnValidVendor = "Database Access Error."
    ElseIf rsTemp.RecordCount = 0 Then
        fnValidVendor = "The vendor does not exists."
    Else
        
        If rsTemp!pm_status = "C" Then
            fnValidVendor = "This Vendor can't be used"
        Else
            sPayTerm = Trim$(rsTemp!pm_std_disc_term & "")
            fnValidVendor = ""
        End If
        
    End If

End Function

Private Function fnValidInvNum(lVendor As Long, lInvNum As Long) As String
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    strSQL = "SELECT * FROM p_nbr WHERE pno_vendor = " & lVendor
    strSQL = strSQL & " AND pno_invoice = " & tfnSQLString(lInvNum)
    
    If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, "fnValidInvNum") < 0 Then
        fnValidInvNum = "Database Access Error."
    ElseIf rsTemp.RecordCount > 0 Then
        fnValidInvNum = "This invoice number has already been used."
    Else
        
        strSQL = "INSERT INTO p_nbr(pno_vendor, pno_invoice, pno_lnk) VALUES"
        strSQL = strSQL & "(" & lVendor & "," & tfnSQLString(lInvNum) & ",0)"
        
        If Not fnExecuteSQL(strSQL, nDB_REMOTE, "fnValidInvNum") Then
            fnValidInvNum = "unable to create invoice number."
        Else
            fnValidInvNum = ""
        End If
        
    End If

End Function

Private Function fnValidInvDate(dtInvDate As Date, dtReportDate As Date) As String

    If dtInvDate > dtReportDate Then
        fnValidInvDate = "The Invoice date is later than report date."
    Else
        fnValidInvDate = ""
    End If
    
End Function

Private Function fnValidPayTerm(sTerm As String, sPayTerm As String) As String

    If sTerm = "" Then
        sTerm = sPayTerm
    End If
    
    fnValidPayTerm = ""
End Function

Private Function fnValidPayType(sPayType As String) As String

    If sPayType <> "C" And sPayType <> "P" And sPayType <> "D" And sPayType <> "T" Then
        fnValidPayType = "Invalid pay Type"
    Else
        fnValidPayType = ""
    End If
    
End Function

Private Function fnValidDraftNum(sPayType As String, sDraftNum As String) As String

    If sPayType = "D" Then
        
        If sDraftNum = "" Then
            fnValidDraftNum = "The draft number can't empty for pay type 'D'."
        End If
        
    End If
    
End Function

'use later
Private Function fnValidInvAmount(dAmount As Double) As String
    fnValidInvAmount = ""
End Function

Private Function fnValidPurchaseType(sType As String) As String

    If sType <> "U" And sType <> "C" And sType <> "P" Then
        fnValidPurchaseType = "Invalid purchase Type"
    Else
        fnValidPurchaseType = ""
    End If
    
End Function

'maybe use later
Private Function fnValidPurchaseQty(dAmount As Double) As String
    fnValidPurchaseQty = ""
End Function

'maybe use later
Private Function fnValidUOM(sUOM As String) As String
    fnValidUOM = ""
End Function

'maybe use later
Private Function fnValidPurchaseCost(dAmount As Double) As String
    
    If dAmount < 0 Then
        fnValidPurchaseCost = "The purchase can't be less than 0."
    Else
        fnValidPurchaseCost = ""
    End If
    
End Function

Private Function fnValidRetailPrice(sRetail As String, sItemCode As String, _
                    lVendor As Long, nPrftCtr As Integer, dtReportDate As Date) As String
    Dim dRetailPrice As Double
    
    If Trim(sRetail) <> "" Then
        
        If Not IsNumeric(sRetail) Then
            fnValidRetailPrice = "Invalid retail price."
        End If
        
    Else
        If fnGetRetailPrice(sItemCode, lVendor, nPrftCtr, dtReportDate, dRetailPrice) Then
            sRetail = CStr(dRetailPrice)
            fnValidRetailPrice = ""
        Else
            fnValidRetailPrice = "No retail price for this item."
        End If
        
    End If
    
End Function

Private Function fnGetRetailPrice(sItemCode As String, lVendor As Long, _
            nPrftCtr As Integer, dtReportDate As Date, dRetailPrice As Double) As Boolean
    Dim lBook As Long
    Dim lSubBook As Long
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim lIcLink As Long
    
    strSQL = "SELECT rsbi_lnk FROM rs_b_item WHERE rsbi_vendor = " & lVendor
    strSQL = strSQL & " AND rsbi_code = " & tfnSQLString(sItemCode)
    
    If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, "fnGetRetailPrice") > 0 Then
        lIcLink = rsTemp!rsbi_lnk
    Else
        Exit Function
    End If
    
    strSQL = "SELECT rsbs_book, rsbs_subbook FROM rs_b_store"
    strSQL = strSQL & " WHERE rsbs_vendor = " & lVendor
    strSQL = strSQL & " AND  rsbs_prft_ctr = " & nPrftCtr

    If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, "fnGetRetailPrice") > 0 Then
        lBook = tfnRound(rsTemp!rsbs_book)
        lSubBook = tfnRound(rsTemp!rsbs_subbook)
    Else
        Exit Function
    End If
    
    strSQL = "SELECT rsbp_retail FROM rs_b_price WHERE rsbp_promo = 'Y'"
    strSQL = strSQL & " AND " & dtReportDate & " BETWEEN rsbp_date and rsbp_ending_date "
    strSQL = strSQL & " AND rsbp_bk_lnk = "
    strSQL = strSQL & " (SELECT rsbb_bk_lnk FROM rs_b_book WHERE rsbb_vendor = " & lVendor
    strSQL = strSQL & " AND rsbb_book = " & lBook
    strSQL = strSQL & " AND rsbb_subbook = " & lSubBook
    strSQL = strSQL & " AND rsbb_ic_lnk = " & lIcLink
    strSQL = strSQL & ")"

    If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, "fnGetRetailPrice") > 0 Then
        dRetailPrice = IIf(IsNull(rsTemp!rsbp_retail), 0, rsTemp!rsbp_retail)
        fnGetRetailPrice = True
    Else
        strSQL = "SELECT rsbp_retail FROM rs_b_price WHERE rsbp_promo = 'N'"
        strSQL = strSQL & " AND " & dtReportDate & " BETWEEN rsbp_date and rsbp_ending_date "
        strSQL = strSQL & " AND rsbp_bk_lnk = "
        strSQL = strSQL & " (SELECT rsbb_bk_lnk FROM rs_b_book WHERE rsbb_vendor = " & lVendor
        strSQL = strSQL & " AND rsbb_book = " & lBook
        strSQL = strSQL & " AND rsbb_subbook = " & lSubBook
        strSQL = strSQL & " AND rsbb_ic_lnk = " & lIcLink
        strSQL = strSQL & ")"

        If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, "fnGetRetailPrice") > 0 Then
            dRetailPrice = IIf(IsNull(rsTemp!rsbp_retail), 0, rsTemp!rsbp_retail)
            fnGetRetailPrice = True
        Else
            fnGetRetailPrice = False
        End If
            
    End If

End Function

