Attribute VB_Name = "modzzseinvp"
Option Explicit

Private Type RSINV_Header
    sRecId  As String
    nPrftCtr As Integer
    dtReportDate As Date
    nShiftNum As Integer
    lVendor As Long
    lInvNum As Long
    dtInvdate As Date
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

'variable for print
Private Const BOTTOM_MARGIN = 600
Private Const STUB_FONT = "Arial"
Private Const PRINT_FONTNAME = "Courier New"     '"Courier 17*1"
Private Const PRINT_FONTSIZE = 10
Private sStoreFont As String
Private nLeftMargin As Integer
Private nTextHeight As Integer
Private nPrevPage As Integer
Private nPageNumber As Integer
Private Const PAGELENGTH As Integer = 120

'Gloabl variable
Private g_dTotalCost As Double
Private g_dTotalExtCost As Double
Private g_lHeaderCount As Long
Private g_bNeedWriteHeader As Boolean
Private g_bNeedErrHeader As Boolean
Private g_lDetailLineNbr As Long

Public Function fnProcessRSInvFile(sFileName As String) As Boolean
    Dim lTotalByte As Long
    Dim lReadByte As Long
    Dim sngPercent As Single
    Dim bIsOpen As Boolean
    Dim nFileNum As Integer
    Dim sLine As String
    Dim lLineCount As Long
    Dim udtInvHeader As RSINV_Header
    Dim udtInvDetail As RSINV_Detail
    Dim sErrMsg As String
    Dim lHeaderAddress As Long
    Dim lHeaderLines As Long
    Dim lSuccessHeaderLines As Long
    Dim lDetailAddress As Long
    Dim lDetailLines As Long
    Dim lInsertLines As Long
    
    Const MAXLINE As Long = 2147483647
    
    On Error GoTo ErrorHandler
    
    'check how many header read
    g_lHeaderCount = 0
    lTotalByte = FileLen(sFileName)
    nFileNum = FreeFile()
    
    Open sFileName For Input As #nFileNum
    bIsOpen = True
    
    subOpenAndClearLogFile
        
    Do While Not EOF(nFileNum)
        Line Input #nFileNum, sLine
        lLineCount = lLineCount + 1
        lReadByte = lReadByte + Len(sLine)
        sngPercent = lReadByte / lTotalByte * 100
        subSetProgress sngPercent
        
        If UCase(Left(sLine, 3)) = "HDR" Then
            lHeaderLines = lHeaderLines + 1
            
            If lDetailLines <> lInsertLines Then
                subDisplayMsg "Not successfully."
                
                If lInsertLines > 0 Then
                    fnDeleteData udtInvHeader.lVendor, udtInvHeader.lInvNum
                End If
        
            ElseIf lDetailLines = lInsertLines Then
                
                If lInsertLines > 0 Then
                    lSuccessHeaderLines = lSuccessHeaderLines + 1
                    subDisplayMsg "Processed Successfully."
                ElseIf lHeaderLines > 1 And lInsertLines = 0 Then
                    subDisplayMsg "No details for this vendor, so processed is not successfully."
                End If
                
            End If
            
            lDetailLines = 0
            lInsertLines = 0
            
            If Not fnProcessHeaderLine(sLine, udtInvHeader) Then
                lHeaderAddress = MAXLINE
                subDisplayMsg "Header in line " & lLineCount & " is not valid."
            Else
                subDisplayMsg "Processing Invoice " & udtInvHeader.lInvNum & " for Vendor " & udtInvHeader.lVendor & "."
                lHeaderAddress = lLineCount
                subWriteHeaderProcLog udtInvHeader
            End If
            
        End If
        
        If UCase(Left(sLine, 3)) = "DET" Then
            lDetailAddress = lLineCount
            
            If lDetailAddress > lHeaderAddress And lHeaderLines > 0 Then
                lDetailLines = lDetailLines + 1
                
                If Not fnProcessDetailLine(sLine, udtInvDetail) Then
                    subDisplayMsg "Detail in line " & lLineCount & " is not valid."
                    lInsertLines = lInsertLines - 1
                Else
                    
                    If fnInsertData(udtInvHeader, udtInvDetail) Then
                        lInsertLines = lInsertLines + 1
                    End If
                    
                End If
                
            End If
                
        End If
        
    Loop
    
    If lDetailLines <> lInsertLines Then
        subDisplayMsg "Not successfully."
        
        If lInsertLines > 0 Then
            fnDeleteData udtInvHeader.lVendor, udtInvHeader.lInvNum
        End If

    ElseIf lDetailLines = lInsertLines Then
        
        If lInsertLines > 0 Then
            lSuccessHeaderLines = lSuccessHeaderLines + 1
            subDisplayMsg "Processed Successfully."
        Else
            subDisplayMsg "No details for this header."
        End If
        
    End If
            
    subWriteSummary
    subCloseLogFile
    
    fnProcessRSInvFile = (lHeaderLines = lSuccessHeaderLines)
    Close #nFileNum
    Exit Function
        
ErrorHandler:

    If bIsOpen Then
        Close #nFileNum
    End If
    
    subDisplayMsg Err.Description
    subCloseLogFile
End Function


Private Function fnProcessHeaderLine(sLine As String, udtInvHeader As RSINV_Header) As Boolean
    On Error GoTo EXITHERE
    
    If Len(sLine) < 70 Then
        Exit Function
    End If
    
    With udtInvHeader
        .sRecId = Trim$(Mid$(sLine, 1, 3))
        .nPrftCtr = CInt(Mid$(sLine, 4, 5))
        .dtReportDate = CDate(Mid$(sLine, 9, 10))
        .nShiftNum = CInt(Mid$(sLine, 19, 5))
        .lVendor = CLng(Mid$(sLine, 24, 10))
        .lInvNum = CLng(Mid$(sLine, 34, 10))
        .dtInvdate = CDate(Mid$(sLine, 44, 10))
        .sTerm = Trim$(Mid$(sLine, 54, 5))
        .sPayType = Trim$(Mid$(sLine, 59, 1))
        .sDraftNum = Trim$(Mid$(sLine, 60, 10))
        .dInvAmount = CDbl(Mid$(sLine, 70, 10))
    End With
    
    fnProcessHeaderLine = True
    Exit Function
EXITHERE:
    fnProcessHeaderLine = False
    Err.Clear
    'subDisplayMsg Err.Description
End Function

Private Function fnProcessDetailLine(sLine As String, udtInvDetail As RSINV_Detail) As Boolean
    On Error GoTo EXITHERE
    
    If Len(sLine) < 34 Then
        Exit Function
    End If
    
    With udtInvDetail
        .sRecId = Trim$(Mid$(sLine, 1, 3))
        .nDetailLine = CInt(Mid$(sLine, 4, 3))
        .sTypeCode = Trim$(Mid$(sLine, 7, 1))
        .sItemCode = Trim$(Mid$(sLine, 8, 10))
        .dQuantity = CDbl(Mid(sLine, 18, 12))
        .sUOM = Trim$(Mid$(sLine, 30, 5))
        .dCost = CDbl(Mid$(sLine, 35, 12))
        .sRetail = Trim$(Mid$(sLine, 47, 12))
    End With
    
    fnProcessDetailLine = True
    Exit Function
EXITHERE:
    fnProcessDetailLine = False
    Err.Clear
    'subDisplayMsg Err.Description
End Function

Private Function fnInsertData(udtInvHeader As RSINV_Header, udtInvDetail As RSINV_Detail) As Boolean
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim bSkipLine As Boolean
    
    subWriteDetailProcLog udtInvHeader, udtInvDetail
    
    If fnValidData(udtInvHeader, udtInvDetail, bSkipLine) Then
        
        If Not bSkipLine Then
            'insert header data
            If g_bNeedWriteHeader Then
            
                If Not fnDeleteData(udtInvHeader.lVendor, udtInvHeader.lInvNum) Then
                     Exit Function
                End If
                                    
                strSQL = "INSERT INTO rs_p_hold_header(rsphh_prft_ctr, rsphh_rpt_date, rsphh_shift, rsphh_vendor,"
                strSQL = strSQL & "rsphh_invoice, rsphh_inv_date, rsphh_std_term, rsphh_type, rsphh_draft_nbr, rsphh_invoice_amt, rsphh_status)"
                strSQL = strSQL & " VALUES(" & udtInvHeader.nPrftCtr & "," & tfnDateString(udtInvHeader.dtReportDate, True) & ","
                strSQL = strSQL & udtInvHeader.nShiftNum & "," & udtInvHeader.lVendor & "," & udtInvHeader.lInvNum & ","
                strSQL = strSQL & tfnDateString(udtInvHeader.dtInvdate, True) & "," & tfnSQLString(udtInvHeader.sTerm) & ","
                strSQL = strSQL & tfnSQLString(udtInvHeader.sPayType) & "," & IIf(Trim(udtInvHeader.sDraftNum) = "", "NULL", udtInvHeader.sDraftNum) & ","
                strSQL = strSQL & tfnRound(udtInvHeader.dInvAmount, 3) & ",'N')"
                
                g_bNeedWriteHeader = False
                g_lDetailLineNbr = 1
                
                If Not fnExecuteSQL(strSQL, nDB_REMOTE, "fnInsertData") Then
                    Exit Function
                End If
                
                strSQL = "INSERT INTO p_nbr(pno_vendor, pno_invoice, pno_lnk) VALUES"
                strSQL = strSQL & "(" & udtInvHeader.lVendor & "," & tfnSQLString(udtInvHeader.lInvNum) & ",0)"
                    
                If Not fnExecuteSQL(strSQL, nDB_REMOTE, "fnValidInvNum") Then
                    Exit Function
                End If
                       
            End If
            
            strSQL = "INSERT INTO rs_p_hold_detail(rsphd_vendor, rsphd_invoice, rsphd_line_nbr, rsphd_type, "
            strSQL = strSQL & " rsphd_code, rsphd_qty, rsphd_stock_unit, rsphd_cost, rsphd_retail)"
            strSQL = strSQL & " VALUES( " & udtInvHeader.lVendor & "," & udtInvHeader.lInvNum & ","
            strSQL = strSQL & g_lDetailLineNbr & "," & tfnSQLString(udtInvDetail.sTypeCode) & ","
            strSQL = strSQL & tfnSQLString(udtInvDetail.sItemCode) & "," & tfnRound(udtInvDetail.dQuantity, 3) & ","
            strSQL = strSQL & tfnSQLString(udtInvDetail.sUOM) & "," & tfnRound(udtInvDetail.dCost, 3) & ","
            strSQL = strSQL & tfnRound(udtInvDetail.sRetail, 3) & ")"
            
            If Not fnExecuteSQL(strSQL, nDB_REMOTE, "fnInsertData") Then
                Exit Function
            End If
            
            g_lDetailLineNbr = g_lDetailLineNbr + 1
        End If
        
        fnInsertData = True
    Else
        fnInsertData = False
    End If
    
End Function

Private Function fnDeleteData(lVendor As Long, lInvoice As Long) As Boolean
    Dim strSQL As String
    
    strSQL = "DELETE FROM p_nbr WHERE pno_vendor = " & lVendor
    strSQL = strSQL & " AND pno_invoice = " & lInvoice
    strSQL = strSQL & " AND EXISTS (SELECT rsphh_status FROM rs_p_hold_header WHERE "
    strSQL = strSQL & " rsphh_vendor = " & lVendor & " AND rsphh_invoice = " & lInvoice
    strSQL = strSQL & " AND rsphh_status = 'N' )"
    
    If Not fnExecuteSQL(strSQL, nDB_REMOTE, "fnDeleteData") Then
        Exit Function
    End If
    
    'delete old data in rs_b_hold_header for this vendor and invoice
    strSQL = "DELETE FROM rs_p_hold_header WHERE rsphh_vendor = " & lVendor
    strSQL = strSQL & " AND rsphh_invoice = " & lInvoice
    
    If Not fnExecuteSQL(strSQL, nDB_REMOTE, "fnDeleteData") Then
        Exit Function
    End If
    
    'delete old data in rs_b_hold_detail for this vendor and invoice
    strSQL = "DELETE FROM rs_p_hold_detail WHERE rsphd_vendor = " & lVendor
    strSQL = strSQL & " AND rsphd_invoice = " & lInvoice
    
    If Not fnExecuteSQL(strSQL, nDB_REMOTE, "fnDeleteData") Then
        Exit Function
    End If

    fnDeleteData = True
End Function

Private Function fnValidData(ByRef udtInvHeader As RSINV_Header, ByRef udtInvDetail As RSINV_Detail, bSkipThisLine As Boolean) As Boolean
    Dim sErrMsg As String
    Dim sItemDesc As String
    Dim sStockUnit As String
    Dim sPayTerm As String
    Dim bUOMMatch As Boolean
    Dim bQuantityZero As Boolean
    Dim sHoldMsg As String
    
    fnValidData = True
    bSkipThisLine = False
    
    If tfnRound(udtInvDetail.dQuantity, DEFAULT_DECIMALS) = 0 Then
        bQuantityZero = True
    End If
    
    'valid this one first, because it item description is required in the log file
    sErrMsg = fnValidItemCode(udtInvHeader.lVendor, udtInvDetail.sItemCode, sItemDesc, sStockUnit)
    
    If sErrMsg <> "" Then
        sHoldMsg = sErrMsg
        fnValidData = False
        subWriteDetailErrLog udtInvHeader, udtInvDetail, sItemDesc, sErrMsg
    End If
    
    sErrMsg = fnValidPrftCtr(udtInvHeader.nPrftCtr)
    
    If sErrMsg <> "" Then
        sHoldMsg = sErrMsg
        fnValidData = False
        subWriteDetailErrLog udtInvHeader, udtInvDetail, sItemDesc, sErrMsg
    End If
    
    sErrMsg = fnValidReportDate(udtInvHeader.nPrftCtr, udtInvHeader.dtReportDate)
    
    If sErrMsg <> "" Then
        sHoldMsg = sErrMsg
        fnValidData = False
        subWriteDetailErrLog udtInvHeader, udtInvDetail, sItemDesc, sErrMsg
    End If
    
    sErrMsg = fnValidShiftNum(udtInvHeader.nPrftCtr, udtInvHeader.nShiftNum, udtInvHeader.dtReportDate)
    
    If sErrMsg <> "" Then
        sHoldMsg = sErrMsg
        fnValidData = False
        subWriteDetailErrLog udtInvHeader, udtInvDetail, sItemDesc, sErrMsg
    End If
    
    sErrMsg = fnValidVendor(udtInvHeader.lVendor, sPayTerm)
    
    If sErrMsg <> "" Then
        sHoldMsg = sErrMsg
        fnValidData = False
        subWriteDetailErrLog udtInvHeader, udtInvDetail, sItemDesc, sErrMsg
    End If
      
    sErrMsg = fnValidInvNum(udtInvHeader.lVendor, udtInvHeader.lInvNum)
    
    If sErrMsg <> "" Then
        sHoldMsg = sErrMsg
        fnValidData = False
        subWriteDetailErrLog udtInvHeader, udtInvDetail, sItemDesc, sErrMsg
    End If
    
    sErrMsg = fnValidInvDate(udtInvHeader.dtInvdate, udtInvHeader.dtReportDate)
    
    If sErrMsg <> "" Then
        sHoldMsg = sErrMsg
        fnValidData = False
        subWriteDetailErrLog udtInvHeader, udtInvDetail, sItemDesc, sErrMsg
    End If
    
    sErrMsg = fnValidPayTerm(udtInvHeader.lVendor, udtInvHeader.sTerm, sPayTerm)
    
    If sErrMsg <> "" Then
        sHoldMsg = sErrMsg
        fnValidData = False
        subWriteDetailErrLog udtInvHeader, udtInvDetail, sItemDesc, sErrMsg
    End If
    
    sErrMsg = fnValidPayType(udtInvHeader.sPayType)
    
    If sErrMsg <> "" Then
        sHoldMsg = sErrMsg
        fnValidData = False
        subWriteDetailErrLog udtInvHeader, udtInvDetail, sItemDesc, sErrMsg
    End If
            
    sErrMsg = fnValidDraftNum(udtInvHeader.sPayType, udtInvHeader.sDraftNum)
    
    If sErrMsg <> "" Then
        sHoldMsg = sErrMsg
        fnValidData = False
        subWriteDetailErrLog udtInvHeader, udtInvDetail, sItemDesc, sErrMsg
    End If
    
    sErrMsg = fnValidInvAmount(udtInvHeader.dInvAmount)
    
    If sErrMsg <> "" Then
        sHoldMsg = sErrMsg
        fnValidData = False
        subWriteDetailErrLog udtInvHeader, udtInvDetail, sItemDesc, sErrMsg
    End If
        
    'valid detail
    sErrMsg = fnValidPurchaseType(udtInvDetail.sTypeCode)
        
    If sErrMsg <> "" Then
        sHoldMsg = sErrMsg
        fnValidData = False
        subWriteDetailErrLog udtInvHeader, udtInvDetail, sItemDesc, sErrMsg
    End If
    
    sErrMsg = fnValidPurchaseQty(udtInvDetail.dQuantity)
        
    If sErrMsg <> "" Then
        sHoldMsg = sErrMsg
        fnValidData = False
        subWriteDetailErrLog udtInvHeader, udtInvDetail, sItemDesc, sErrMsg
    End If
    
    sErrMsg = fnValidUOM(udtInvHeader, udtInvDetail, sStockUnit, bUOMMatch)
        
    If sErrMsg <> "" Then
        sHoldMsg = sErrMsg
        fnValidData = False
        subWriteDetailErrLog udtInvHeader, udtInvDetail, sItemDesc, sErrMsg
    End If
    
    If bUOMMatch Then
        sErrMsg = fnValidPurchaseCost(udtInvDetail.dCost)
            
        If sErrMsg <> "" Then
            sHoldMsg = sErrMsg
            fnValidData = False
            subWriteDetailErrLog udtInvHeader, udtInvDetail, sItemDesc, sErrMsg
        End If
        
        sErrMsg = fnValidRetailPrice(udtInvDetail.sRetail, udtInvDetail.sItemCode, udtInvHeader.lVendor, udtInvHeader.nPrftCtr, udtInvHeader.dtInvdate)
            
        If sErrMsg <> "" Then
            sHoldMsg = sErrMsg
            fnValidData = False
            subWriteDetailErrLog udtInvHeader, udtInvDetail, sItemDesc, sErrMsg
        End If
            
    End If
    
    If sHoldMsg <> "" And bQuantityZero Then
        sErrMsg = "****All errors for this item will be ignored since quantity is Zero"
        fnValidData = True
        bSkipThisLine = True
        subWriteDetailErrLog udtInvHeader, udtInvDetail, sItemDesc, sErrMsg
    End If
    

End Function

Public Sub subSetProgress(sngPercent As Single)
    
    If sngPercent > 0# Then
        
        If Not frmzzseinvp.PbProgressBar.Visible Then
            frmzzseinvp.PbProgressBar.Visible = True
        End If
    
    Else
        frmzzseinvp.PbProgressBar.Visible = False
    End If
    
    frmzzseinvp.PbProgressBar.Value = sngPercent
    frmzzseinvp.PbProgressBar.Refresh
    DoEvents
End Sub

Public Sub subDisplayMsg(sMsg As String)
    frmzzseinvp.lstStatus.AddItem sMsg
    frmzzseinvp.lstStatus.Refresh
End Sub

Private Sub subWriteHeaderErrLog(udtInvHeader As RSINV_Header)
    Dim sLine As String
    Dim sVendorName As String
    
    sVendorName = fnGetVendorName(udtInvHeader.lVendor)
    
    Print #g_nErrorLogFile, ""
    sLine = "Profit Center: " & CStr(udtInvHeader.nPrftCtr)
    Print #g_nErrorLogFile, sLine
    sLine = "Shift Number: " & CStr(udtInvHeader.nShiftNum)
    Print #g_nErrorLogFile, sLine
    sLine = "Vendor Number: " & CStr(udtInvHeader.lVendor) & Space(10) & "Vendor Name: " & sVendorName
    Print #g_nErrorLogFile, sLine
    sLine = String(100, "-")
    Print #g_nErrorLogFile, sLine
    sLine = "Date" & Space(5) & "Invoice #" & Space(2) & "Inv. Code" & Space(2) & "Description" & Space(10) & "Error Message"
    Print #g_nErrorLogFile, sLine
    sLine = String(100, "-")
    Print #g_nErrorLogFile, sLine
End Sub

Private Sub subWriteDetailErrLog(udtInvHeader As RSINV_Header, udtInvDetail As RSINV_Detail, sItemDesc As String, sErrMsg As String)
    Dim sLine As String
    Dim lSpaceLen As Long
    
    If g_bNeedErrHeader Then
        subWriteHeaderErrLog udtInvHeader
        g_bNeedErrHeader = False
    End If
    
    sLine = CStr(Format(udtInvHeader.dtReportDate, "MM/DD/YY"))
    sLine = sLine & Space(1) & CStr(udtInvHeader.lInvNum)
    lSpaceLen = 11 - Len(CStr(udtInvHeader.lInvNum))
    
    If lSpaceLen < 0 Then
        lSpaceLen = 0
    End If
    
    sLine = sLine & Space(lSpaceLen) & udtInvDetail.sItemCode
    lSpaceLen = 11 - Len(udtInvDetail.sItemCode)
    
    If lSpaceLen < 0 Then
        lSpaceLen = 0
    End If
    
    sLine = sLine & Space(lSpaceLen) & sItemDesc
    lSpaceLen = 21 - Len(sItemDesc)
    
    If lSpaceLen < 0 Then
        lSpaceLen = 0
    End If
    
    sLine = sLine & Space(lSpaceLen) & sErrMsg
    Print #g_nErrorLogFile, sLine
End Sub

Private Sub subWriteHeaderProcLog(udtInvHeader As RSINV_Header)
    Dim sLine As String
    Dim sVendorName As String
    Dim sItemDesc As String
    Dim dExtCost As Double
    
    sVendorName = fnGetVendorName(udtInvHeader.lVendor)
    g_bNeedWriteHeader = True
    g_bNeedErrHeader = True
    
    g_lHeaderCount = g_lHeaderCount + 1 'add header count if header changed
    
    If g_lHeaderCount > 1 Then
        subWriteSummary
    End If
    
    g_dTotalCost = 0#
    g_dTotalExtCost = 0#
    
    Print #g_nProcessingFile, ""
    sLine = "Profit Center: " & udtInvHeader.nPrftCtr
    Print #g_nProcessingFile, sLine
    sLine = "Shift Number: " & udtInvHeader.nShiftNum
    Print #g_nProcessingFile, sLine
    sLine = "Vendor Number: " & udtInvHeader.lVendor & Space(10) & "Vendor Name: " & sVendorName
    Print #g_nProcessingFile, sLine
    sLine = String(100, "-")
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
    Dim lSpaceLen As Long
    
    'write detail
    sLine = CStr(Format(udtInvHeader.dtReportDate, "MM/DD/YY"))
    sLine = sLine & Space(1) & CStr(udtInvHeader.lInvNum)
    
    lSpaceLen = 11 - Len(CStr(udtInvHeader.lInvNum))
    
    If lSpaceLen < 0 Then
        lSpaceLen = 0
    End If
    
    sLine = sLine & Space(lSpaceLen) & udtInvDetail.sItemCode
    sItemDesc = fnGetItemDesc(udtInvHeader.lVendor, udtInvDetail.sItemCode)
     
    lSpaceLen = 11 - Len(udtInvDetail.sItemCode)
    
    If lSpaceLen < 0 Then
        lSpaceLen = 0
    End If
    
    sLine = sLine & Space(lSpaceLen) & sItemDesc
    lSpaceLen = 21 - Len(sItemDesc)
    
    If lSpaceLen < 0 Then
        lSpaceLen = 0
    End If
    
    sLine = sLine & Space(lSpaceLen) & CStr(udtInvDetail.dQuantity)
    lSpaceLen = 9 - Len(CStr(udtInvDetail.dQuantity))
    
    If lSpaceLen < 0 Then
        lSpaceLen = 0
    End If
    
    sLine = sLine & Space(lSpaceLen) & CStr(udtInvDetail.dCost)
    dExtCost = udtInvDetail.dQuantity * udtInvDetail.dCost
    lSpaceLen = 11 - Len(CStr(udtInvDetail.dCost))
    
    If lSpaceLen < 0 Then
        lSpaceLen = 0
    End If
    
    sLine = sLine & Space(lSpaceLen) & CStr(dExtCost)
    Print #g_nProcessingFile, sLine
    g_dTotalCost = g_dTotalCost + udtInvDetail.dCost
    
    g_dTotalExtCost = g_dTotalExtCost + dExtCost
    
End Sub

Private Sub subWriteSummary()
    Dim sLine As String
    Dim lSpaceLen As Long
    Print #g_nProcessingFile, ""
    'Added tfnRound before printing totals in order to make certain the length of g_dTotalCost
    '7-8-2002 Robert Atwood for Seven-11
    g_dTotalCost = tfnRound(g_dTotalCost, 2)
    g_dTotalExtCost = tfnRound(g_dTotalExtCost, 2)
    
    lSpaceLen = 11 - Len(CStr(g_dTotalCost))
    
    If lSpaceLen < 0 Then
        lSpaceLen = 0
    End If
    
    sLine = "TOTAL" & Space(56) & CStr(g_dTotalCost) & Space(lSpaceLen) & CStr(g_dTotalExtCost)
    Print #g_nProcessingFile, sLine
    Print #g_nProcessingFile, ""
End Sub

Private Sub subOpenAndClearLogFile()
    Dim sProcessLogFile As String
    Dim sErrorLogFile As String

    On Error Resume Next
    
    sProcessLogFile = App.Path & "\zzseplog.log"
    sErrorLogFile = App.Path & "\zzseelog.log"
    
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

Public Function fnFileExist(sFile As String) As Boolean
    On Error Resume Next
    
    fnFileExist = (Dir$(sFile) <> "")
    
End Function

' Get records from the given SQL statement
' nDB = 1 ---> Informax Database (remote)
'     = 2 ---> Access Database (local)
'This function will return a recordcount
Public Function fnGetRecord(rsTemp As Recordset, strSQL As String, Optional nDB As Variant, Optional sCalledFrom As String = "", Optional bShowError As Variant) As Long
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
    
    If IsMissing(bShowError) Then
        bShowError = True
    End If

    tfnErrHandler SUB_NAME + "," + sCalledFrom, strSQL, bShowError
    fnGetRecord = -9999
End Function

Public Function fnExecuteSQL(szSQL As String, Optional nDB As Variant, _
                Optional sCalledFrom As String = "", Optional bShowError As Variant) As Boolean
                
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
    
    If IsMissing(bShowError) Then
        bShowError = True
    End If

    tfnErrHandler "fnExecuteSQL, " & sCalledFrom, szSQL, bShowError
    On Error GoTo 0
    
End Function


Private Function fnValidItemCode(lVendor As Long, sItemCode As String, sItemDesc As String, sStockUnit As String) As String
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    strSQL = "SELECT rsbi_desc,rsbi_stock_unit FROM rs_b_item WHERE rsbi_vendor = " & lVendor
    strSQL = strSQL & " AND rsbi_code = " & tfnSQLString(sItemCode)
    
    If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, "fnGetItemDesc") < 0 Then
        fnValidItemCode = "Database Access Error."
    ElseIf rsTemp.RecordCount = 0 Then
        fnValidItemCode = "Item code does not exists."
    Else
        sItemDesc = Trim(rsTemp!rsbi_desc & "")
        sStockUnit = Trim(rsTemp!rsbi_stock_unit & "")
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
        sMsg = "Lastest process date is not available."
    Else
        dtLastProcDate = IIf(IsNull(rsTemp!prft_posted_date), Null, CDate(rsTemp!prft_posted_date))
        
        If Not IsNull(dtLastProcDate) Then
            
            If dtLastProcDate > dtReportDate Then
                sMsg = "Report date is earlier than last processed date."
            End If
            
        End If
        
    End If
    
    If sMsg = "" Then
        strSQL = "SELECT glp_status FROM gl_period WHERE " & tfnDateString(dtReportDate, True)
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
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim lShiftLink As Long
    
    strSQL = "SELECT rssl_shl FROM rs_shiftlink, rs_shiftHold WHERE rssl_shl = rssh_shl AND rssl_shift = " & nShiftNum
    strSQL = strSQL & " AND rssl_prft_ctr = " & nPrftCtr
    strSQL = strSQL & " AND rssl_date = " & tfnDateString(dtReportDate, True)
    
    If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, "fnValidShiftNum") > 0 Then
        lShiftLink = tfnRound(rsTemp!rssl_shl)
        
        strSQL = "SELECT rss_status FROM rs_summary WHERE rss_shl = " & lShiftLink
        
        If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, "fnValidShiftNum") > 0 Then
            
            If rsTemp!rss_status & "" <> "R" Then
                fnValidShiftNum = "The profit center,report date and shift No. is used already."
            End If
            
        End If
        
    End If
    
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
            fnValidVendor = "This Vendor can't be used."
        Else
            sPayTerm = Trim$(rsTemp!pm_std_disc_term & "")
            fnValidVendor = ""
        End If
        
    End If

End Function

Private Function fnValidInvNum(lVendor As Long, lInvNum As Long) As String
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    'check data in holding table already
    strSQL = "SELECT rsphh_status FROM rs_p_hold_header WHERE rsphh_vendor = " & lVendor
    strSQL = strSQL & " AND rsphh_invoice = " & tfnSQLString(lInvNum)
    
    If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, "fnValidInvNum") < 0 Then
        fnValidInvNum = "Database Access Error."
    ElseIf rsTemp.RecordCount > 0 Then
        
        If UCase(rsTemp!rsphh_status & "") = "N" Then
            fnValidInvNum = ""
        ElseIf UCase(rsTemp!rsphh_status & "") = "Y" Then
            fnValidInvNum = "The invoice for this vendor has already been posted."
        End If
            
    Else
        strSQL = "SELECT * FROM p_nbr WHERE pno_vendor = " & lVendor
        strSQL = strSQL & " AND pno_invoice = " & tfnSQLString(lInvNum)
        
        If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, "fnValidInvNum") < 0 Then
            fnValidInvNum = "Database Access Error."
        ElseIf rsTemp.RecordCount > 0 Then
            fnValidInvNum = "This invoice number has already been used."
        Else
            fnValidInvNum = ""
        End If
        
    End If
    
End Function

Private Function fnValidInvDate(dtInvdate As Date, dtReportDate As Date) As String

    If dtInvdate > dtReportDate Then
        fnValidInvDate = "The Invoice date is later than report date."
    Else
        fnValidInvDate = ""
    End If
    
End Function

Private Function fnValidPayTerm(lVendor As Long, sTerm As String, sPayTerm As String) As String
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    If Trim(sTerm) = "" Then
        sTerm = sPayTerm
        fnValidPayTerm = ""
    Else
        strSQL = "SELECT pm_std_disc_term FROM p_vendor WHERE pm_vendor = " & lVendor
        strSQL = strSQL & " AND pm_std_disc_term = " & tfnSQLString(sTerm)
        
        If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, "fnValidPayTerm") < 0 Then
            fnValidPayTerm = "Database access error."
        ElseIf rsTemp.RecordCount = 0 Then
            fnValidPayTerm = "Invalid discount term for this vendor."
        Else
            fnValidPayTerm = ""
        End If
        
    End If
    
End Function

Private Function fnValidPayType(sPayType As String) As String

    If sPayType <> "C" And sPayType <> "P" And sPayType <> "D" And sPayType <> "T" Then
        fnValidPayType = "Invalid pay Type."
    Else
        fnValidPayType = ""
    End If
    
End Function

Private Function fnValidDraftNum(sPayType As String, sDraftNum As String) As String
    Dim strSQL As String
    Dim rsTemp As Recordset
    Const FUNC_NAME As String = "fnValidDraftNum"

    If sPayType = "D" Or sPayType = "M" Then
        
        If Trim(sDraftNum) = "" Then
            fnValidDraftNum = "The draft number can't empty for pay type 'D'."
            Exit Function
        End If
        
        'Check p_draft
        strSQL = "SELECT pdr_draft_nbr FROM p_draft WHERE pdr_draft_nbr = " & CLng(sDraftNum)
    
        If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, FUNC_NAME) > 0 Then
            fnValidDraftNum = "Draft number has been used by other file."
            Exit Function
        End If

        'Check rs_shifthold
        strSQL = "SELECT rssh_5 FROM rs_shifthold WHERE rssh_type = 'V' AND rssh_4 = 2 AND rssh_5 = " & CLng(sDraftNum)
        
        If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, FUNC_NAME) > 0 Then
            fnValidDraftNum = "Draft number has been used by holding file."
            Exit Function
        End If
        
    End If
    
End Function

'use later
Private Function fnValidInvAmount(dAmount As Double) As String
    fnValidInvAmount = ""
End Function

Private Function fnValidPurchaseType(sType As String) As String

    'Right now, only handle price book type.
    'If sType <> "U" And sType <> "C" And sType <> "P" Then
    If sType <> "P" Then
        fnValidPurchaseType = "Invalid purchase Type."
    Else
        fnValidPurchaseType = ""
    End If
    
End Function

'maybe use later
Private Function fnValidPurchaseQty(dAmount As Double) As String
    fnValidPurchaseQty = ""
End Function


Private Function fnValidUOM(udtInvHeader As RSINV_Header, udtInvDetail As RSINV_Detail, sStockUnit As String, bUOMMatch As Boolean) As String
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    If UCase(udtInvDetail.sUOM) = UCase(sStockUnit) Then
        bUOMMatch = True
        fnValidUOM = ""
    Else
        strSQL = "SELECT icm1.icm_cost, icm1.icm_retail FROM item_cost_maint icm1 "
        strSQL = strSQL & " WHERE icm1.icm_vendor = " & tfnRound(udtInvHeader.lVendor)
        strSQL = strSQL & " AND icm1.icm_code = " & tfnSQLString(udtInvDetail.sItemCode)
        strSQL = strSQL & " AND icm1.icm_uom = " & tfnSQLString(UCase(udtInvDetail.sUOM))
        strSQL = strSQL & " AND icm1.icm_eff_date = "
        strSQL = strSQL & " (SELECT MAX(icm_eff_date) FROM item_cost_maint "
        strSQL = strSQL & " WHERE icm_vendor = " & tfnRound(udtInvHeader.lVendor)
        strSQL = strSQL & " AND icm_code = " & tfnSQLString(udtInvDetail.sItemCode)
        strSQL = strSQL & " AND icm_eff_date <= " & tfnDateString(udtInvHeader.dtInvdate, True)
        strSQL = strSQL & ")"
        
        If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, "fnValidUOM") < 0 Then
            fnValidUOM = "Database access error."
        ElseIf rsTemp.RecordCount = 0 Then
            fnValidUOM = udtInvDetail.sUOM & " is an invalid Unit of Measure."
        Else
            'accoring to tom's idea, all cost should came from file
            'udtInvDetail.dCost = IIf(IsNull(rsTemp!icm_cost), 0, rsTemp!icm_cost)
            udtInvDetail.sRetail = IIf(IsNull(rsTemp!icm_retail), "0", rsTemp!icm_retail)
            fnValidUOM = ""
        End If
        
    End If
    
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
                    lVendor As Long, nPrftCtr As Integer, dtInvdate As Date) As String
    Dim dRetailPrice As Double
    
    'according to tom,the retail price all came from Price book 04/06/01
    
'    If Trim(sRetail) <> "" Then
'
'        If Not IsNumeric(sRetail) Then
'            fnValidRetailPrice = "Invalid retail price."
'        End If
'
'    Else
    If fnGetRetailPrice(sItemCode, lVendor, nPrftCtr, dtInvdate, dRetailPrice) Then
        sRetail = CStr(dRetailPrice)
        fnValidRetailPrice = ""
    Else
        fnValidRetailPrice = "No retail price for this item."
    End If
    
'    End If
    
End Function

Private Function fnGetRetailPrice(sItemCode As String, lVendor As Long, _
            nPrftCtr As Integer, dtInvdate As Date, dRetailPrice As Double) As Boolean
    Dim lBook As Long
    Dim lSubBook As Long
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim lIcLink As Long
    
    strSQL = "SELECT rsbi_ic_lnk FROM rs_b_item WHERE rsbi_vendor = " & lVendor
    strSQL = strSQL & " AND rsbi_code = " & tfnSQLString(sItemCode)
    
    If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, "fnGetRetailPrice") > 0 Then
        lIcLink = IIf(IsNull(rsTemp!rsbi_ic_lnk), -1, rsTemp!rsbi_ic_lnk)
    Else
        Exit Function
    End If
    
    strSQL = "SELECT rsbs_book, rsbs_subbook FROM rs_b_store"
    strSQL = strSQL & " WHERE rsbs_vendor = " & lVendor
    strSQL = strSQL & " AND  rsbs_prft_ctr = " & nPrftCtr

    If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, "fnGetRetailPrice") > 0 Then
        lBook = IIf(IsNull(rsTemp!rsbs_book), -1, rsTemp!rsbs_book)
        lSubBook = IIf(IsNull(rsTemp!rsbs_subbook), -1, rsTemp!rsbs_subbook)
    Else
        Exit Function
    End If
    
    'strSQL = "SELECT rsbp_retail FROM rs_b_price WHERE rsbp_promo = 'Y'"
    strSQL = "SELECT rsbp_ent_retail, rsbp_date FROM rs_b_price WHERE rsbp_promo = 'Y'"
    strSQL = strSQL & " AND " & tfnDateString(dtInvdate, True) & " BETWEEN rsbp_date and rsbp_ending_date "
    strSQL = strSQL & " AND rsbp_bk_lnk = "
    strSQL = strSQL & " (SELECT rsbb_bk_lnk FROM rs_b_book WHERE rsbb_vendor = " & lVendor
    strSQL = strSQL & " AND rsbb_book = " & lBook
    strSQL = strSQL & " AND rsbb_subbook = " & lSubBook
    strSQL = strSQL & " AND rsbb_ic_lnk = " & lIcLink
    strSQL = strSQL & ")"
    strSQL = strSQL & " ORDER BY rsbp_date DESC "
    
    If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, "fnGetRetailPrice") > 0 Then
        dRetailPrice = IIf(IsNull(rsTemp!rsbp_ent_retail), 0, rsTemp!rsbp_ent_retail)
        fnGetRetailPrice = True
    Else
        'strSQL = "SELECT rsbp_retail FROM rs_b_price WHERE rsbp_promo = 'N'"
        strSQL = "SELECT rsbp_ent_retail FROM rs_b_price WHERE rsbp_promo = 'N'"
        strSQL = strSQL & " AND " & tfnDateString(dtInvdate, True) & " BETWEEN rsbp_date and rsbp_ending_date "
        strSQL = strSQL & " AND rsbp_bk_lnk = "
        strSQL = strSQL & " (SELECT rsbb_bk_lnk FROM rs_b_book WHERE rsbb_vendor = " & lVendor
        strSQL = strSQL & " AND rsbb_book = " & lBook
        strSQL = strSQL & " AND rsbb_subbook = " & lSubBook
        strSQL = strSQL & " AND rsbb_ic_lnk = " & lIcLink
        strSQL = strSQL & ")"

        If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, "fnGetRetailPrice") > 0 Then
            dRetailPrice = IIf(IsNull(rsTemp!rsbp_ent_retail), 0, rsTemp!rsbp_ent_retail)
            fnGetRetailPrice = True
        Else
            fnGetRetailPrice = False
        End If
            
    End If

End Function

Public Sub subSentErrorLogToPrinter()
    Dim nFileNum As Integer
    Dim sErrorLogFile As String
    Dim sLine As String
    Dim bIsOpen As Boolean
    
    On Error GoTo EXITHERE
    
    If Not fnInitPrinter() Then
        frmzzseinvp.tfnSetStatusBarMessage "Printer not Ready"
        Exit Sub
    End If
    
    nPageNumber = 1
    nPrevPage = 1
    
    subPrintRptHeader "Error"
    nFileNum = FreeFile()
    sErrorLogFile = App.Path & "\zzseelog.log"
    
    Open sErrorLogFile For Input As #nFileNum
    bIsOpen = True
    
    Do While Not EOF(nFileNum)
        Line Input #nFileNum, sLine
        
        If nPrevPage <> nPageNumber Then
            subPrintRptHeader "ERROR"
            subOutput Space(PAGELENGTH)
            nPrevPage = nPageNumber
        End If
    
        subOutput sLine
    Loop

    Close #nFileNum
    subPrinterEndDocument
    Exit Sub
EXITHERE:

    If bIsOpen Then
        Close #nFileNum
    End If
    
    subPrinterEndDocument
    MsgBox Err.Description, vbExclamation
End Sub

Public Sub subSentProcLogToPrinter()
    Dim nFileNum As Integer
    Dim sProcLogFile As String
    Dim sLine As String
    Dim bIsOpen As Boolean
    
    On Error GoTo EXITHERE
    If Not fnInitPrinter() Then
        frmzzseinvp.tfnSetStatusBarMessage "Printer not Ready"
        Exit Sub
    End If
    
    nPageNumber = 1
    nPrevPage = 1
    
    subPrintRptHeader "PROCESSING"
    nFileNum = FreeFile()
    sProcLogFile = App.Path & "\zzseplog.log"
    
    Open sProcLogFile For Input As #nFileNum
    bIsOpen = True
    
    Do While Not EOF(nFileNum)
        Line Input #nFileNum, sLine
        
        If nPrevPage <> nPageNumber Then
            subPrintRptHeader "PROCESSING"
            subOutput Space(PAGELENGTH)
            nPrevPage = nPageNumber
        End If
    
        subOutput sLine
    Loop

    Close #nFileNum
    subPrinterEndDocument
    Exit Sub

EXITHERE:

    If bIsOpen Then
        Close #nFileNum
    End If
    
    subPrinterEndDocument
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub subPrintRptHeader(sMsg As String)
    Dim sCompanyName As String
    Dim sRundate As String
    Dim sRuntime As String
    Dim nSpc As Integer
    Dim sReportLine As String, sPageNum As String, sHeader As String
    
    sRundate = Format(Date, "MM/DD/YYYY")
    sRuntime = Format(Time, "HH:MM AMPM")
    sHeader = "Incoming Invoice File"
    
    If fnGetCompanyName(sCompanyName) Then
        sPageNum = "Page No.  " + fnRightJustified(nPageNumber, "####")
        nSpc = (PAGELENGTH - Len(Trim(sCompanyName))) / 2 - Len("Program ID: ZZRINVPL")
        sReportLine = "Program ID: ZZSEINVP" + Space(nSpc) + sCompanyName
        sReportLine = sReportLine & Space(PAGELENGTH - Len(sReportLine) - Len(sPageNum)) & sPageNum
    End If
    
    subOutput sReportLine
    
    nSpc = (PAGELENGTH - Len(Trim(sHeader))) / 2 - Len(Trim(sRundate & "Run Date: "))
    sReportLine = "Run Date: " + sRundate + Space(nSpc) + sHeader
    subOutput sReportLine
    
    If UCase(sMsg) = "PROCESSING" Then
        nSpc = (PAGELENGTH - Len("Processing Log")) / 2 - Len(Trim(sRuntime & "Run Time: "))
        sReportLine = "Run Time: " + sRuntime + Space(nSpc) + "Processing Log"
        subOutput sReportLine
    Else
        nSpc = (PAGELENGTH - Len("Error Log")) / 2 - Len(Trim(sRuntime & "Run Time: "))
        sReportLine = "Run Time: " + sRuntime + Space(nSpc) + "Error Log"
        subOutput sReportLine
    End If
    
End Sub
Public Function fnInitPrinter(Optional vNextPage As Variant) As Boolean
    Dim sErrMsg As String
    
    On Error GoTo ErrInitPrinter
    
    If IsMissing(vNextPage) Then
        sStoreFont = Printer.FontName
    End If

    Printer.Orientation = vbPRORLandscape
    Printer.FontName = STUB_FONT
    Printer.Print " "
    Printer.FontName = PRINT_FONTNAME
    Printer.FontSize = PRINT_FONTSIZE

    
    nLeftMargin = (Printer.ScaleWidth - Printer.TextWidth(Space(PAGELENGTH))) / 2
    nTextHeight = Printer.ScaleHeight - BOTTOM_MARGIN
    
    fnInitPrinter = True
    
    Exit Function
    
ErrInitPrinter:
    sErrMsg = "An error has occurred while initializing the Printer. Err Code: " & _
        Err.Number & ", Err Desc: " & Err.Description
    MsgBox "Called by fnInitPrinter, " & sErrMsg, vbExclamation
    Printer.EndDoc
End Function

Public Sub subOutput(ByVal sOut As String)

    Printer.CurrentX = nLeftMargin
    
    If sOut <> "\page\" Then
        Printer.Print sOut
    End If
    
    If Printer.CurrentY >= nTextHeight Or sOut = "\page\" Then
        Printer.Print Space(100) & "< continued >"
        Printer.NewPage
        nPageNumber = nPageNumber + 1
    End If
    
End Sub

Public Sub subPrinterEndDocument()
    On Error Resume Next
    Printer.NewPage
    Printer.EndDoc
    Printer.FontName = sStoreFont
    Printer.Orientation = vbPRORPortrait
End Sub

Private Function fnRightJustified(ByVal sIn As String, sFormatString As String) As String
    fnRightJustified = Format(Format(sIn, sFormatString), String(Len(sFormatString), "@"))
End Function

Private Function fnCenter(sIn As String, nMaxLen As Integer) As String
    Dim nDiff As String, nSpcLeft As Integer
    
    nDiff = nMaxLen - Len(sIn)
    If nDiff >= nMaxLen Then
        fnCenter = sIn
        Exit Function
    End If
    
    nSpcLeft = Int(nDiff / 2)
    fnCenter = Space(nSpcLeft) + sIn + Space(nDiff - nSpcLeft)
End Function


Public Function fnGetCompanyName(sCompanyName As String) As Boolean
    Dim sSql As String, rsTemp As Recordset

    sSql = "SELECT con_name FROM co_company_name"

    If fnGetRecord(rsTemp, sSql, nDB_REMOTE, "fnGetCompanyName") < 0 Then
        Exit Function
    End If
    
    If rsTemp.RecordCount = 0 Then
        tfnErrHandler "fnGetCompanyName", 60001, "Company has been set up.  Program will be terminated."
        Exit Function
    End If
    
    If IsNull(rsTemp!con_name) Then
        tfnErrHandler "fnGetCompanyName", 60003, "Company is NULL.  Program will be terminated."
        Exit Function
    End If
    
    sCompanyName = Trim(rsTemp!con_name)
    
    If sCompanyName = "" Then
        tfnErrHandler "fnGetCompanyName", 60003, "Company is empty.  Program will be terminated."
        Exit Function
    End If
    
    fnGetCompanyName = True
End Function

