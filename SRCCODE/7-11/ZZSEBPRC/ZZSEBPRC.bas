Attribute VB_Name = "modZZSEBPRC"
Option Explicit

Private Const PRINT_MARGIN_LEFT = 150     'Pixel
Private Const PRINT_MARGIN_RIGHT = 150    'Pixel
Private Const PRINT_MARGIN_TOP = 250      'Pixel
Private Const PRINT_MARGIN_BOTTOM = 0   'Pixel

Public t_nFormMode As Integer         'global used to track the current form operating mode
Public Const IDLE_MODE As Integer = 0 'idle mode activates the NoDrop Cursor
Public Const ADD_MODE As Integer = 1
Public Const EDIT_MODE As Integer = 2

Public sLogFilePath As String
Public tgmSales As clsTGSpreadSheet
Public tgsSales As clsTGSelector

Public tgmApprove As clsTGSpreadSheet
Public tgsApprove As clsTGSelector
Public tgcExtension As clsColumnExtension

Public tgmDetail As clsTGSpreadSheet

Public Const TabSales As Integer = 0
Public Const TabHours As Integer = 1
Public Const TabProcess As Integer = 1
Public Const TabApprove As Integer = 2
Public Const TabDetails As Integer = 3
Public Const nTabHours As Integer = 4

'Sales Grid Column Names
Public Const colSPrftCtr As Integer = 0
Public Const colSPrftName As Integer = 1
Public Const colSAmount As Integer = 2
Public Const colSFromDate As Integer = 3
Public Const colSToDate As Integer = 4
Public ColxSOldPrftCtr As Integer

'Time Card Grid Column Names
Public Const colHClockIn As Integer = 0
Public Const colHPrftCtr As Integer = 1
Public Const colHPayCode As Integer = 2
Public Const colHHrsDol As Integer = 3
Public ColHHdnSource As Integer

'Profit Center Grid Column Names
Public Const colPProfit As Integer = 0
Public Const colPTotal As Integer = 1

'Approve value
Public Const sColAppYes As String = "Y"
Public Const sColAppNo As String = "N"

'Approve Grid Column Names
Public Const colAApprove As Integer = 0
Public Const colAEmpNo As Integer = 1
Public Const colAEmpName As Integer = 2
Public Const colADate As Integer = 3
Public Const colAPrftCtr As Integer = 4
Public Const colAPayCode As Integer = 5
Public Const colAPayHours As Integer = 6
Public Const colABonusAmt As Integer = 7
Public colAHdsOverride As Integer
Public colAHdnPrftName As Integer
Public colAHdsBonusDesc As Integer
Public colAHdnSeq As Integer
Public colAHdnBAmtLvls As Integer

'Detail Grid Column Names
Public Const colDBCode As Integer = 0
Public Const colDBCDesc As Integer = 1
Public Const colDBLevel As Integer = 2
Public Const colDBType As Integer = 3
Public Const colDBFreq As Integer = 4
Public Const colDElgDate As Integer = 5
Public Const colDBAmt As Integer = 6
Public colDHdnEmpNo As Integer
Public colDHdnPrftCtr As Integer

Public arySalesDesc() As Variant
Public arySalesType() As Variant

Public sSalesTypeCode As String
Public Const sBiWeek As String = "B"
Public Const sTwoWeek As String = "P"
Public Const sOneMth As String = "M"
Public Const sGas As String = "G"
Public Const sRatio As String = "R"

Public vArrBonus() As Variant
Public objMath As clsEquation
Public objCond As clsCondition

Public bShowDetail As Boolean

Public sPayCode_RegHrs As String
Public sPayCode_OtHrs As String
'
Public bNoRecordFound As Boolean

Public Sub Main()
    Dim sCommand As String
    
    #If PROTOTYPE Then
        frmZZSEBPRC.Show
        Exit Sub
    #End If
    
    sCommand = Trim(Command)
    sLogFilePath = fnAddBkSlash(App.Path) & "ZZSEBPRC.LOG"
    
    subDeleteErrLog 'Delete the old log file if any...
    
    If sCommand = t_szHandShake Then
        frmZZSEBPRC.Show
    Else
        If Len(sCommand) = 0 Then
            frmSplash.Show
        End If
    End If
    
End Sub

Public Sub subShowMainForm()
    frmZZSEBPRC.Show
End Sub

Private Function fnExecuteSQL(szSQL As String, _
                              Optional nDB As Variant, _
                              Optional szCalledFrom As Variant, _
                              Optional bShowError As Variant) As Boolean
                
      Dim szMsg As String
      
      On Error GoTo SQLError
      
      If IsMissing(nDB) Then
       nDB = nDB_REMOTE
      End If
    
      Select Case nDB
        
        Case nDB_LOCAL
           dbLocal.Execute szSQL
        Case nDB_REMOTE
           t_dbMainDatabase.ExecuteSQL szSQL
      End Select
      
      fnExecuteSQL = True
      Exit Function
      
SQLError:
      fnExecuteSQL = False
      If IsMissing(szCalledFrom) Then
         szCalledFrom = ""
      End If
      If IsMissing(bShowError) Then
         bShowError = True
      End If
      tfnErrHandler "fnExecuteSQL, " & szCalledFrom, szSQL, bShowError
      On Error GoTo 0
End Function

Public Function fnConCat(MyPath As String, MyName As String) As String
    fnConCat = IIf(Right(MyPath, 1) = "\", MyPath, MyPath + "\") + MyName
End Function

Public Function fnAddBkSlash(ByVal sIn As String) As String
    sIn = Trim(sIn)
    If Right(sIn, 1) <> "\" Then fnAddBkSlash = sIn + "\" Else fnAddBkSlash = sIn
End Function

Public Function fnAddFwSlash(ByVal sIn As String) As String
    sIn = Trim(sIn)
    If Right(sIn, 1) <> "/" Then fnAddFwSlash = sIn + "/" Else fnAddFwSlash = sIn
End Function

Public Sub subLogErrMsg(sMsg As String, Optional bClear As Boolean = False)
    Dim nFileNumber As Integer
    Dim sLineContents As String
    Dim sTimeStamp As String
    Dim sArrMsg() As String
    Dim i As Integer
    
    Dim x As Long
    
    On Error GoTo ErrTrap
    
    If bClear Then
        frmZZSEBPRC!lstProcess.Clear
        
        'hide the scrollbar
        x = frmZZSEBPRC.TextWidth("  ")
        frmZZSEBPRC!lstProcess.Tag = "0"
        If frmZZSEBPRC.ScaleMode = vbTwips Then x = x / Screen.TwipsPerPixelX
        SendMessageByNum frmZZSEBPRC!lstProcess.hwnd, LB_SETHORIZONTALEXTENT, x, 0
        
        DoEvents
        
        Exit Sub
    End If
    
    'Put the time stamp if the sLogFilePath is empty
    On Error Resume Next
    sTimeStamp = "Error Log Created on : " & Date & " at " & Time & vbCrLf
    
    nFileNumber = FreeFile
    Open sLogFilePath For Input As #nFileNumber
    If Not EOF(nFileNumber) Then
        Line Input #nFileNumber, sLineContents
        Close nFileNumber
    End If
    
    If sLineContents = "" Then
        tfnLog sTimeStamp, sLogFilePath
    End If
    
    'Writing the log to the file...
    tfnLog sMsg, sLogFilePath
    
    sArrMsg = Split(sMsg, vbCrLf)
    For i = 0 To UBound(sArrMsg)
        frmZZSEBPRC.lstProcess.AddItem sArrMsg(i)
        
        If sArrMsg(i) <> "" Then
            x = frmZZSEBPRC.TextWidth(sMsg & "  ")
            If x > Val(frmZZSEBPRC!lstProcess.Tag) Then
                frmZZSEBPRC!lstProcess.Tag = x
                If frmZZSEBPRC.ScaleMode = vbTwips Then x = x / Screen.TwipsPerPixelX
                SendMessageByNum frmZZSEBPRC!lstProcess.hwnd, LB_SETHORIZONTALEXTENT, x, 0
            End If
        End If
            
        frmZZSEBPRC!lstProcess.ListIndex = frmZZSEBPRC!lstProcess.ListCount - 1
    Next i
    
    DoEvents
    
    Exit Sub
    
ErrTrap:
    'error
End Sub

Private Sub subDeleteErrLog()
    Dim sFileFound As String
    
    sFileFound = Dir(sLogFilePath)
    
    If sFileFound <> "" Then
        Kill sLogFilePath
    End If

End Sub

Public Function fnCreateSearchTable(szNumber As String, szName As String) As Boolean
    Dim sSql As String
    Dim sSqDrop As String

    'Drop temp table first
    sSqDrop = " DROP TABLE sTmpEmpTable"
    On Error GoTo ErrorDropTable
    t_dbMainDatabase.ExecuteSQL sSqDrop
    'Create a temp table
    sSql = "CREATE TEMP TABLE sTmpEmpTable ( " & szNumber & " INTEGER," & szName & " CHAR(60), prm_ssn CHAR(11)) "
    On Error GoTo errCreateTable
    t_dbMainDatabase.ExecuteSQL sSql
    
    On Error GoTo errInsertRecords
    sSql = "INSERT INTO sTmpEmpTable SELECT prm_empno, TRIM(prm_last_name) || ', '"
    sSql = sSql & " || TRIM(prm_first_name) || ' '  || TRIM(prm_middle_name), prm_ssn"
    sSql = sSql & " FROM pr_master WHERE  TRIM(prm_middle_name)<> ''"
    'sSql = sSql & " AND prm_security_code <= " & tfnSQLString(Security_Code)
    t_dbMainDatabase.ExecuteSQL sSql
    
    sSql = " INSERT INTO sTmpEmpTable SELECT prm_empno, TRIM(prm_last_name) || ', '"
    sSql = sSql & " || TRIM(prm_first_name), prm_ssn"
    sSql = sSql & " FROM pr_master WHERE  TRIM(prm_middle_name) = '' OR prm_middle_name IS NULL"
    'sSql = sSql & " AND prm_security_code <= " & tfnSQLString(Security_Code)
    t_dbMainDatabase.ExecuteSQL sSql
    
    fnCreateSearchTable = True
    Exit Function

ErrorDropTable:
    Resume Next

errCreateTable:
    tfnErrHandler "fnCreateSearchTable", sSql
    Exit Function

errInsertRecords:
    tfnErrHandler "fnCreateSearchTable", sSql
End Function

Public Function fnCreateSalesTable() As Boolean
    Const SUB_NAME As String = "fnCreateTempTableVar"
    
    Dim strSQL As String
    Dim i As Integer
    
    'predefined variables
    arySalesDesc = Array("Bi-weekly Sales", "One Month Sales", "Gas Sales", _
        "Inv. Shortage Ratio")
    
    arySalesType = Array(sBiWeek, sOneMth, sGas, sRatio)
    
    On Error GoTo Continue
    strSQL = "DROP TABLE tmp_sales_type"
    t_dbMainDatabase.ExecuteSQL strSQL
    
Continue:
    strSQL = "CREATE TEMP TABLE tmp_sales_type (tst_desc char(20), tst_type char(1))"
    If Not fnExecuteSQL(strSQL, , SUB_NAME) Then
        Exit Function
    End If
    
    For i = 0 To UBound(arySalesDesc)
        strSQL = "INSERT INTO tmp_sales_type VALUES(" + tfnSQLString(arySalesDesc(i))
        strSQL = strSQL + "," + tfnSQLString(arySalesType(i)) + ")"
        If Not fnExecuteSQL(strSQL, , SUB_NAME) Then
            Exit Function
        End If
    Next
    
    fnCreateSalesTable = True
End Function

Public Function fnCreateReport(Index As Integer) As Boolean
    Dim sArrReport() As String
    Dim sHeadTitle As String
    Dim sReportTitle As String
    Dim sApprove As String
    Dim nArSize As Integer
    Dim sCondition As String
    Dim sFormula As String
    Dim sFormulaLine As String
    Dim sAdjCond As String
    Dim sAdjFormula As String
    Dim sAdjFormulaLine As String
    Dim i As Integer, j As Integer
    
    Dim sReportID As String
    
    Screen.MousePointer = vbHourglass
    frmZZSEBPRC.tfnSetStatusBarMessage "Printing report, please wait..."
    
    fnCreateReport = False

    Select Case Index
        Case TabApprove
            sReportID = "ZZSEBPRA"
            
            ReDim sArrReport(tgmApprove.RowCount - 1)
            For i = 0 To tgmApprove.RowCount - 1
                sApprove = "N"
                
                If tgmApprove.CellValue(colAApprove, i) = sColAppYes Then
                    sApprove = "Y"
                End If
                
                sArrReport(i) = fnTranc(sApprove, 5, vbCenter) & Space(1) _
                    & fnTranc(tgmApprove.CellValue(colAEmpNo, i), 9, vbLeftJustify) & Space(1) _
                    & fnTranc(tgmApprove.CellValue(colAEmpName, i), 46, vbLeftJustify) & Space(1) _
                    & fnTranc(tgmApprove.CellValue(colAPrftCtr, i), 5, vbLeftJustify) & Space(2) _
                    & fnTranc(tgmApprove.CellValue(colAPayCode, i), 4, vbLeftJustify) & Space(2) _
                    & fnTranc(tgmApprove.CellValue(colAPayHours, i), 5, vbLeftJustify) & Space(2) _
                    & fnTranc(tgmApprove.CellValue(colADate, i), 10, vbLeftJustify) & Space(1) _
                    & fnTranc(tgmApprove.CellValue(colABonusAmt, i), 10, vbRightJustify)
            Next i
            
            sReportTitle = "Employee Commission Approval Report"
            sHeadTitle = fnTranc("", 5, vbLeftJustify) & Space(1)
            sHeadTitle = sHeadTitle & fnTranc("Employee", 9, vbLeftJustify) & Space(1)
            sHeadTitle = sHeadTitle & fnTranc("", 46, vbLeftJustify) & Space(1)
            sHeadTitle = sHeadTitle & fnTranc("Prft", 5, vbLeftJustify) & Space(2)
            sHeadTitle = sHeadTitle & fnTranc("Pay", 4, vbLeftJustify) & Space(2)
            sHeadTitle = sHeadTitle & fnTranc("Pay", 5, vbLeftJustify) & Space(2)
            sHeadTitle = sHeadTitle & fnTranc("Process", 10, vbLeftJustify) & Space(1)
            sHeadTitle = sHeadTitle & fnTranc("Comm.", 10, vbRightJustify)
            'second line
            sHeadTitle = sHeadTitle & vbCrLf
            sHeadTitle = sHeadTitle & fnTranc("Apprv", 5, vbLeftJustify) & Space(1)
            sHeadTitle = sHeadTitle & fnTranc("Number", 9, vbLeftJustify) & Space(1)
            sHeadTitle = sHeadTitle & fnTranc("Employee Name", 46, vbLeftJustify) & Space(1)
            sHeadTitle = sHeadTitle & fnTranc("Ctr", 5, vbLeftJustify) & Space(2)
            sHeadTitle = sHeadTitle & fnTranc("Code", 4, vbLeftJustify) & Space(2)
            sHeadTitle = sHeadTitle & fnTranc("Hours", 5, vbLeftJustify) & Space(2)
            sHeadTitle = sHeadTitle & fnTranc("Date", 10, vbLeftJustify) & Space(1)
            sHeadTitle = sHeadTitle & fnTranc("Amount", 10, vbRightJustify)
        Case TabDetails
            sReportID = "ZZSEBPRD"
            
            nArSize = tgmDetail.RowCount * 4 - 1
            j = 0
            ReDim sArrReport(nArSize)
            For i = 0 To tgmDetail.RowCount - 1
                sArrReport(j) = fnTranc(tgmDetail.CellValue(colDBCode, i), 5, vbLeftJustify) & Space(2) _
                    & fnTranc(tgmDetail.CellValue(colDBCDesc, i), 49, vbLeftJustify) & Space(2) _
                    & fnTranc(tgmDetail.CellValue(colDBLevel, i), 5, vbCenter) & Space(2) _
                    & fnTranc(tgmDetail.CellValue(colDBType, i), 4, vbCenter) & Space(2) _
                    & fnTranc(tgmDetail.CellValue(colDBFreq, i), 9, vbCenter) & Space(2) _
                    & fnTranc(tgmDetail.CellValue(colDElgDate, i), 10, vbCenter) & Space(2) _
                    & fnTranc(tgmDetail.CellValue(colDBAmt, i), 10, vbRightJustify)
                
                j = j + 1
                
                'display condition, formula, adj.condition, adj.formula
                subGetBFormula tgmDetail.CellValue(colDBCode, i), _
                   tgmDetail.CellValue(colDBLevel, i), _
                   sCondition, sFormula, sAdjCond, sAdjFormula
                
                sFormulaLine = ""
                If sCondition <> "" Then
                    If Len(sCondition) >= 4 Then
                        If Left(sCondition, 2) <> "if" And Left(sCondition, 4) <> "when" Then
                            sFormulaLine = "if "
                        End If
                    End If
                    
                    sFormulaLine = sFormulaLine + sCondition + ", "
                End If
                
                sFormulaLine = sFormulaLine + sFormula
                
                If sFormulaLine <> "" Then
                    sArrReport(j) = Space(3) & "Formula: " & sFormulaLine
                    j = j + 1
                End If
                
                sAdjFormulaLine = ""
                If sAdjCond <> "" Then
                    If Len(sAdjCond) >= 4 Then
                        If Left(sAdjCond, 2) <> "if" And Left(sAdjCond, 4) <> "when" Then
                            sAdjFormulaLine = "if "
                        End If
                    End If
                    
                    sAdjFormulaLine = sAdjFormulaLine + sAdjCond + ", "
                End If
                
                sAdjFormulaLine = sAdjFormulaLine + sAdjFormula
                
                If sAdjFormulaLine <> "" Then
                    sArrReport(j) = Space(3) & "Adj. Formula: " & sAdjFormulaLine
                    j = j + 1
                End If
                
                If sFormulaLine = "" And sAdjFormulaLine = "" Then
                    sArrReport(j) = Space(3) & "Formula not found."
                    j = j + 1
                End If
                
                sArrReport(j) = ""
                j = j + 1
            Next i
            
            If j > 0 Then
                ReDim Preserve sArrReport(j - 1)
            End If
            
            sReportTitle = "Employee Commission Details"
            If frmZZSEBPRC.txtEmployee <> "" Then
                sReportTitle = frmZZSEBPRC.txtEmployee & "-" & frmZZSEBPRC.txtEmpName & " Commission Details"
            End If
            sHeadTitle = fnTranc("Bonus", 5, vbLeftJustify) & Space(2)
            sHeadTitle = sHeadTitle & fnTranc("Commission Code Description", 49, vbCenter) & Space(2)
            sHeadTitle = sHeadTitle & fnTranc("Level", 5, vbLeftJustify) & Space(2)
            sHeadTitle = sHeadTitle & fnTranc("Type", 4, vbLeftJustify) & Space(2)
            sHeadTitle = sHeadTitle & fnTranc("Frequency", 9, vbLeftJustify) & Space(2)
            sHeadTitle = sHeadTitle & fnTranc("Eligible", 10, vbCenter) & Space(2)
            sHeadTitle = sHeadTitle & fnTranc("Comm.", 10, vbRightJustify)
            sHeadTitle = sHeadTitle & vbCrLf
            sHeadTitle = sHeadTitle & fnTranc("Code", 5, vbLeftJustify) & Space(77)
            sHeadTitle = sHeadTitle & fnTranc("Date", 10, vbCenter) & Space(2)
            sHeadTitle = sHeadTitle & fnTranc("Amount", 10, vbRightJustify)
    End Select
    
    sHeadTitle = sHeadTitle & vbCrLf & String(104, "-")
    
    If Not fnSetupPrinter(vbPRORPortrait) Then
        frmZZSEBPRC.tfnSetStatusBarError "Failed to print report"
        Exit Function
    End If
    
    subSetReportID sReportID
    subSetTitle sHeadTitle
    
    If Not fnSendToPrinter(sArrReport(), sReportTitle) Then
        frmZZSEBPRC.tfnSetStatusBarError "Failed to print report"
        Exit Function
    End If
    
    #If WRT_RPT_TO_FILE Then
        If Not fnSendToFile(sArrReport(), sReportTitle, App.Path + "\" + sReportID + ".TXT") Then
            frmZZSEBPRC.tfnSetStatusBarError "Failed to write report to file"
            Exit Function
        End If
    #End If
    
    subSetReportID ""
    
    frmZZSEBPRC.tfnSetStatusBarMessage "Report was printed successfully"
    Screen.MousePointer = vbDefault
    fnCreateReport = True
    
End Function

Public Sub subPrintProcess(lstOutput As ListBox)
    Dim i As Integer
    Dim nLeft As Integer
    Dim nTop As Integer
    Dim nBottom As Integer
    
    nLeft = PRINT_MARGIN_LEFT * Printer.TwipsPerPixelX
    nTop = PRINT_MARGIN_TOP * Printer.TwipsPerPixelY
    nBottom = Printer.Height - (nTop + PRINT_MARGIN_BOTTOM * Printer.TwipsPerPixelY)
    
    Printer.CurrentY = nTop
    For i = 0 To lstOutput.ListCount - 1
        Printer.CurrentX = nLeft
        Printer.Print lstOutput.List(i)
        If Printer.CurrentY >= nBottom Then
            Printer.NewPage
            Printer.CurrentY = nTop
        End If
    Next i
    Printer.EndDoc

End Sub

'This function will calculate bonus for only one bonus code including all its levels
Public Function fnGetBonusAmount(rsBonus As Recordset) As Double
    Const SUB_NAME As String = "fnGetBonusAmount"
    
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim lEmployeeNo As Long
    Dim nPrftCtr As Integer
    Dim sBType As String
    Dim sBCode As String
    Dim sBGrade As String
    Dim nLevel As Integer
    Dim sErrMsg As String
    
    lEmployeeNo = tfnRound(rsBonus!bm_empno)
    nPrftCtr = tfnRound(rsBonus!bm_eligible_pc)
    sBCode = fnGetField(rsBonus!bc_bonus_code)
    sBType = fnGetField(rsBonus!bc_type)
    sBGrade = fnGetField(rsBonus!bc_grade)
    nLevel = tfnRound(rsBonus!bf_level)
    
    fnGetBonusAmount = 0#
    
    If sBCode = "" Or sBType = "" Then
        subLogErrMsg Space(7) & "Commission Code is NULL"
        Exit Function
    End If
    If sBType = "" Then
        subLogErrMsg Space(7) & "Commission Code Type is NULL"
        Exit Function
    End If
    If sBGrade = "" Then
        subLogErrMsg Space(7) & "Commission Code Type is NULL"
        Exit Function
    End If
    
    sErrMsg = fnValidGetBAmountEmp(lEmployeeNo, nPrftCtr, sBGrade)
    
    If sErrMsg <> "" Then
        subLogErrMsg Space(7) & sErrMsg
        Exit Function
    End If
    
    strSQL = "SELECT * FROM bonus_formula WHERE bf_bonus_code = " & tfnSQLString(sBCode)
    strSQL = strSQL & " AND bf_level = " & tfnRound(nLevel)
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) <> 1 Then
        subLogErrMsg Space(7) & "No record found for the commission formula"
        bNoRecordFound = True
        Exit Function
    End If
    
    fnGetBonusAmount = fnCalculateBonus(lEmployeeNo, nPrftCtr, _
        tfnRound(rsTemp!bf_percent, DEFAULT_DECIMALS), _
        tfnRound(rsTemp!bf_dollar, 2), _
        tfnRound(rsTemp!bf_amount1, 2), _
        tfnRound(rsTemp!bf_amount2, 2), _
        fnGetField(rsTemp!bf_variable1), _
        fnGetField(rsTemp!bf_variable2), _
        fnGetField(rsTemp!bf_variable3), _
        tfnRound(rsTemp!bf_max_total, 2), _
        fnGetField(rsTemp!bf_formula), _
        fnGetField(rsTemp!bf_condition), _
        fnGetField(rsTemp!bf_adj_formula), _
        fnGetField(rsTemp!bf_adj_condition), _
        sBType)
                
End Function

Public Function fnCheckApprove() As Boolean
    Dim lApproved As Long
    
    lApproved = 0
    
    fnHasApprove lApproved
    tgmApprove.Rebind
    DoEvents
    
    If lApproved = 0 Then
        frmZZSEBPRC.tfnSetStatusBarError "No Approved row available to insert"
        Exit Function
    End If
    
    If lApproved < tgmApprove.RowCount Then
        If MsgBox("Rows that are not approved will not be inserted. Are you sure you want " _
           + "to continue?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Function
        End If
    End If
    
    fnCheckApprove = True
End Function

Public Function fnInsertHoldBonus() As String
    Const SUB_NAME As String = "fnInsertHoldBonus"
    
    Dim strSQLinsert As String
    Dim strSQL1 As String
    Dim lRow As Long
    Dim lEmpNo As Long
    Dim nPrftCtr As Integer
    Dim sPayCode As String
    Dim sEndDate  As String
    Dim sErrMsg As String
    
    strSQLinsert = "INSERT INTO bonus_hold (bh_empno, bh_prft_ctr, bh_pay_code, bh_check_amount,"
    strSQLinsert = strSQLinsert & " bh_hours, bh_date, bh_override, bh_chk_link) VALUES ("
    
    For lRow = 0 To tgmApprove.RowCount - 1
        If tgmApprove.CellValue(colAApprove, lRow) = sColAppYes Then
            lEmpNo = tfnRound(tgmApprove.CellValue(colAEmpNo, lRow))
            nPrftCtr = tfnRound(tgmApprove.CellValue(colAPrftCtr, lRow))
            sPayCode = fnGetField(tgmApprove.CellValue(colAPayCode, lRow))
            sEndDate = fnGetField(tgmApprove.CellValue(colADate, lRow))
            
            If fnChkLinkIsZero(lEmpNo, nPrftCtr, sPayCode, sEndDate, sErrMsg) Then
                'delete the old data in bonus_hold first
                strSQL1 = "DELETE FROM bonus_hold WHERE bh_empno = " & lEmpNo
                strSQL1 = strSQL1 & " AND bh_prft_ctr = " & nPrftCtr
                strSQL1 = strSQL1 & " AND bh_pay_code = " & tfnSQLString(sPayCode)
                strSQL1 = strSQL1 & " AND bh_date = " & tfnDateString(sEndDate, True)
                strSQL1 = strSQL1 & " AND bh_chk_link = 0"
                
                If Not fnExecuteSQL(strSQL1, , SUB_NAME) Then
                    fnInsertHoldBonus = "Failed to delete Old Commission Data"
                    Exit Function
                End If
                
                strSQL1 = lEmpNo & ", "
                strSQL1 = strSQL1 & nPrftCtr & ", "
                strSQL1 = strSQL1 & tfnSQLString(sPayCode) & ", "
                strSQL1 = strSQL1 & tfnRound(tgmApprove.CellValue(colABonusAmt, lRow), 2) & ", "
                If fnGetField(tgmApprove.CellValue(colAPayHours, lRow)) = "" Then
                    strSQL1 = strSQL1 & "NULL, "
                Else
                    strSQL1 = strSQL1 & tfnRound(tgmApprove.CellValue(colAPayHours, lRow), 2) & ", "
                End If
                strSQL1 = strSQL1 & tfnDateString(sEndDate, True) & ", "
                strSQL1 = strSQL1 & tfnSQLString(tgmApprove.CellValue(colAHdsOverride, lRow)) & ", "
                strSQL1 = strSQL1 & "0"  'per weigong insert 0 for bh_chk_link
                
                If Not fnExecuteSQL(strSQLinsert + strSQL1 + ")", , SUB_NAME) Then
                    fnInsertHoldBonus = "Failed to insert Commission Data"
                    Exit Function
                End If
            Else
                If sErrMsg <> "" Then
                    fnInsertHoldBonus = sErrMsg
                    Exit Function
                End If
            End If
        End If
    Next lRow
    
    fnInsertHoldBonus = ""
End Function

Private Function fnChkLinkIsZero(lEmpNo As Long, _
                                nPrftCtr As Integer, _
                                sPayCode As String, _
                                sEndDate As String, _
                                sErrMsg As String) As Boolean
    
    Const SUB_NAME As String = "fnChkLinkIsZero"
    
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    strSQL = "SELECT bh_chk_link FROM bonus_hold"
    strSQL = strSQL & " WHERE bh_chk_link <> 0"
    strSQL = strSQL & " AND bh_empno = " & lEmpNo
    strSQL = strSQL & " AND bh_prft_ctr = " & nPrftCtr
    strSQL = strSQL & " AND bh_pay_code = " & tfnSQLString(sPayCode)
    strSQL = strSQL & " AND bh_date = " & tfnDateString(sEndDate, True)

    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        sErrMsg = "Failed to access database"
    End If
    
    If rsTemp.RecordCount > 0 Then
        Exit Function
    End If
    
    fnChkLinkIsZero = True
End Function

'return vbYes, vbNo, or vbCancel
Public Function fnCheckBonusHold() As Integer
    Const SUB_NAME As String = "fnCheckBonusHold"
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim sMsg As String
    Dim nBonusMasterCount As Long
    Dim nChkLinkIsZero As Long
    Dim nChkLinkNotZero As Long
    
    strSQL = "SELECT bm_empno, bm_eligible_pc FROM bonus_master, bonus_formula"
    strSQL = strSQL & " WHERE " & tfnDateString(frmZZSEBPRC!txtEndDate, True)
    strSQL = strSQL & " BETWEEN bm_eligible_date AND bm_stop_date"
    
    If frmZZSEBPRC!txtPrftCtr <> "" Then
        strSQL = strSQL & " AND bm_eligible_pc = " & tfnRound(frmZZSEBPRC!txtPrftCtr)
    End If
    
    If frmZZSEBPRC!txtEmpProcess <> "" Then
        strSQL = strSQL & " AND bm_empno = " & tfnRound(frmZZSEBPRC!txtEmpProcess)
    End If
    
    strSQL = strSQL & " AND bm_bonus_code = bf_bonus_code"
    'strSQL = strSQL & " GROUP BY bm_empno, bm_eligible_pc"
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        MsgBox "Failed to access database", vbExclamation
        Exit Function
    End If
    
    nBonusMasterCount = rsTemp.RecordCount
    
    If nBonusMasterCount = 0 Then
        fnCheckBonusHold = vbYes
        Exit Function
    End If
    
    strSQL = "SELECT bh_chk_link FROM bonus_hold"
    strSQL = strSQL & " WHERE bh_chk_link <> 0"
    
    If frmZZSEBPRC!txtPrftCtr <> "" Then
        strSQL = strSQL & " AND bh_prft_ctr = " & tfnRound(frmZZSEBPRC!txtPrftCtr)
    End If
    
    If frmZZSEBPRC!txtEmpProcess <> "" Then
        strSQL = strSQL & " AND bh_empno = " & tfnRound(frmZZSEBPRC!txtEmpProcess)
    End If
    
    strSQL = strSQL & " AND bh_date = " & tfnDateString(frmZZSEBPRC!txtEndDate, True)
    'do not include pay code is hoursly
    strSQL = strSQL & " AND bh_pay_code NOT IN (SELECT prpa_pay_code FROM pr_pay "
    strSQL = strSQL & " WHERE (prpa_type = 'P' AND prpa_calc_method = 'H') "
    strSQL = strSQL & " OR (prpa_type = 'N' AND prpa_calc_method = 'D'))"
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        MsgBox "Failed to access database", vbExclamation
        Exit Function
    End If
    
    nChkLinkNotZero = rsTemp.RecordCount

    strSQL = "SELECT bh_chk_link FROM bonus_hold"
    strSQL = strSQL & " WHERE bh_chk_link = 0"
    'do not include pay code is hoursly
    strSQL = strSQL & " AND bh_pay_code NOT IN (SELECT prpa_pay_code FROM pr_pay "
    strSQL = strSQL & " WHERE (prpa_type = 'P' AND prpa_calc_method = 'H') "
    strSQL = strSQL & " OR (prpa_type = 'N' AND prpa_calc_method = 'D'))"
    
    
    If frmZZSEBPRC!txtPrftCtr <> "" Then
        strSQL = strSQL & " AND bh_prft_ctr = " & tfnRound(frmZZSEBPRC!txtPrftCtr)
    End If
    
    If frmZZSEBPRC!txtEmpProcess <> "" Then
        strSQL = strSQL & " AND bh_empno = " & tfnRound(frmZZSEBPRC!txtEmpProcess)
    End If
    
    strSQL = strSQL & " AND bh_date = " & tfnDateString(frmZZSEBPRC!txtEndDate, True)

    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        MsgBox "Failed to access database", vbExclamation
        Exit Function
    End If
    
    nChkLinkIsZero = rsTemp.RecordCount

    If nChkLinkNotZero >= nBonusMasterCount Then
        
        If frmZZSEBPRC!txtPrftCtr = "" And frmZZSEBPRC!txtEmpProcess = "" Then
            sMsg = "All the Commission Records have been selected into Payroll Print Group, " _
                + "and cannot be processed."
        ElseIf frmZZSEBPRC!txtPrftCtr = "" Then
            sMsg = "All the Commission Records for the Employee Number have been " _
                + "selected into Payroll Print Group, and cannot not be processed."
        Else
            sMsg = "All the Commission Records for the Profit Center have been selected " _
                + "into Payroll Print Group, and cannot not be processed."
        End If
        
        MsgBox sMsg, vbExclamation

        fnCheckBonusHold = vbCancel
        
        subLogErrMsg sMsg
        subLogErrMsg " "

        Exit Function
    Else
        
        If nChkLinkNotZero > 0 Then
            
            If frmZZSEBPRC!txtPrftCtr = "" And frmZZSEBPRC!txtEmpProcess = "" Then
                sMsg = "Some Commission Records have been selected into Payroll Print " _
                    + "Group, and cannot be processed. "
            ElseIf frmZZSEBPRC!txtPrftCtr = "" Then
                sMsg = "Some Commission Records for the Employee Number have been " _
                    + "selected into Payroll Print Group, and cannot not be processed. "
            Else
                sMsg = "Some Commission Records for the Profit Center have been selected " _
                    + "into Payroll Print Group, and cannot not be processed. "
            End If
            
            fnCheckBonusHold = MsgBox(sMsg + "Are you sure you want to continue?", vbQuestion _
                + vbYesNo + vbDefaultButton2)
            
            subLogErrMsg sMsg
            subLogErrMsg " "
            
            Exit Function
        End If
        
    End If
    
    
    sMsg = ""
    
    If nChkLinkIsZero >= nBonusMasterCount Then
        
        If frmZZSEBPRC!txtPrftCtr = "" And frmZZSEBPRC!txtEmpProcess = "" Then
            sMsg = "All the Commission Records"
        ElseIf frmZZSEBPRC!txtPrftCtr = "" Then
            sMsg = "All the Commission Records for the Employee Number"
        Else
            sMsg = "All the Commission Records for the Profit Center"
        End If
    
    Else
        
        If nChkLinkIsZero > 0 Then
            
            If frmZZSEBPRC!txtPrftCtr = "" And frmZZSEBPRC!txtEmpProcess = "" Then
                sMsg = "Some Commission Records"
            ElseIf frmZZSEBPRC!txtPrftCtr = "" Then
                sMsg = "Some Commission Records for the Employee Number"
            Else
                sMsg = "Some Commission Records for the Profit Center"
            End If
        
        End If
    
    End If
    
    If sMsg <> "" Then
        sMsg = sMsg + " have been processed, and will be replaced! "

        fnCheckBonusHold = MsgBox(sMsg + "Are you sure you want to continue?", vbQuestion _
            + vbYesNo + vbDefaultButton2)
        
        subLogErrMsg sMsg
        subLogErrMsg " "
        
        Exit Function
    End If
    
    fnCheckBonusHold = vbYes
End Function

'This function will calculate the amount for 1 Employee, 1 BCode and 1 Level at a time
Private Function fnCalculateBonus(lEmpNo As Long, _
                                  nPrftCtr As Integer, _
                                  PCT As Double, _
                                  DOL As Double, _
                                  AMT1 As Double, _
                                  AMT2 As Double, _
                                  sV1 As String, _
                                  sV2 As String, _
                                  sV3 As String, _
                                  MXT As Double, _
                                  sFmla As String, _
                                  sCond As String, _
                                  sAFmla As String, _
                                  sACond As String, _
                                  sBType As String) As Double
    
    Const SUB_NAME As String = "fnCalculateBonus"
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim sErrMsg As String
    Dim i As Integer
    Dim V1 As Double, V2 As Double, V3 As Double
    Dim bConditionOK As Boolean
    Dim dBonusAmt As Double
    
    fnCalculateBonus = 0#
    
    sFmla = fnGetField(sFmla)
    sCond = fnGetField(sCond)
    sACond = fnGetField(sACond)
    sAFmla = fnGetField(sAFmla)
    sV1 = Trim(sV1)
    sV2 = Trim(sV2)
    sV3 = Trim(sV3)
    PCT = tfnRound(PCT, DEFAULT_DECIMALS)
    DOL = tfnRound(DOL, DEFAULT_DECIMALS)
    
    'check formula
    If sFmla = "" Then
        sErrMsg = "Formula not supply"
    Else
        sErrMsg = fnCheckFormula(sFmla, sBType)
    End If
    
    If sErrMsg <> "" Then
        subLogErrMsg sErrMsg
        Exit Function
    End If
    
    'check condition
    If sCond = "" Then
        bConditionOK = True
    Else
        sErrMsg = fnCheckCondition(sCond, sBType)
    End If
    
    If sErrMsg <> "" Then
        subLogErrMsg sErrMsg
        Exit Function
    End If
    
    'check adj formula
    If sAFmla <> "" Then
        sErrMsg = fnCheckFormula(sAFmla, sBType)
    End If
    
    If sErrMsg <> "" Then
        subLogErrMsg sErrMsg
        Exit Function
    End If
    
    'check adj. condition
    If sACond <> "" Then
        sErrMsg = fnCheckCondition(sACond, sBType)
    End If
    
    If sErrMsg <> "" Then
        subLogErrMsg sErrMsg
        Exit Function
    End If
    
    If bShowDetail Then
        subLogErrMsg " "
    End If
    
    'Get real values...
    If sV1 <> "" Then
        V1 = fnGetVarValue(lEmpNo, nPrftCtr, "v1", sV1, sErrMsg)
        
        If sErrMsg <> "" Then
            subLogErrMsg sErrMsg
        Else
            
            If bShowDetail Then
                
                If sV1 = "check_amount" Then
                    subLogErrMsg "**v1 (" + sV1 + ") = the result of the formula"
                Else
                    subLogErrMsg "**v1 (" + sV1 + ") = " & V1
                End If
                
            End If
            
        End If
        
    End If
    
    If sV2 <> "" Then
        V2 = fnGetVarValue(lEmpNo, nPrftCtr, "v2", sV2, sErrMsg)
        
        If sErrMsg <> "" Then
            subLogErrMsg sErrMsg
        Else
            If bShowDetail Then
                
                If sV2 = "check_amount" Then
                    subLogErrMsg "**v2 (" + sV2 + ") = the result of the formula"
                Else
                    subLogErrMsg "**v2 (" + sV2 + ") = " & V2
                End If
            
            End If
        
        End If
    
    End If
    
    If sV3 <> "" Then
        V3 = fnGetVarValue(lEmpNo, nPrftCtr, "v3", sV3, sErrMsg)
        
        If sErrMsg <> "" Then
            subLogErrMsg sErrMsg
        Else
        
            If bShowDetail Then
                
                If sV3 = "check_amount" Then
                    subLogErrMsg "**v3 (" + sV3 + ") = the result of the formula"
                Else
                    subLogErrMsg "**v3 (" + sV3 + ") = " & V3
                End If
                
            End If
            
        End If
        
    End If
    
    'set the variables value for condition
    If bShowDetail Then
        subLogErrMsg "**pct=" & PCT & ", dol=" & DOL & ", amt1=" & AMT1 _
            & ", amt2=" & AMT2 & ", mxt=" & MXT
    
        subLogErrMsg " "
    End If
    
    objCond.Var("pct") = PCT
    objCond.Var("dol") = DOL
    objCond.Var("amt1") = AMT1
    objCond.Var("amt2") = AMT2
    objCond.Var("mxt") = MXT
    
    If sV1 <> "" Then
        objCond.Var("v1") = V1
    End If
    
    If sV2 <> "" Then
        objCond.Var("v2") = V2
    End If
    
    If sV3 <> "" Then
        objCond.Var("v3") = V3
    End If
    
    'set the variables value for formula
    objMath.Var("pct") = PCT
    objMath.Var("dol") = DOL
    objMath.Var("amt1") = AMT1
    objMath.Var("amt2") = AMT2
    objMath.Var("mxt") = MXT
    
    If sV1 <> "" Then
        objMath.Var("v1") = V1
    End If
    
    If sV2 <> "" Then
        objMath.Var("v2") = V2
    End If
    
    If sV3 <> "" Then
        objMath.Var("v3") = V3
    End If
    
    If sCond <> "" Then
        bConditionOK = objCond.CheckCondition(sCond, sErrMsg)
        
        If sErrMsg <> "" Then
            subLogErrMsg sErrMsg & ", Invalid Condition Clause (" & sCond & ")"
            Exit Function
        Else
            
            If bShowDetail Then
                subLogErrMsg "Condition = " & sCond
                subLogErrMsg "Result = " & IIf(bConditionOK, "True", "False")
            End If
            
        End If
        
    End If
    
    dBonusAmt = 0#
    
    If bConditionOK Then
        dBonusAmt = tfnRound(objMath.Calculate(sFmla, sErrMsg), DEFAULT_DECIMALS)
        
        If sErrMsg <> "" Then
            subLogErrMsg sErrMsg & ", Invalid Formula (" & sFmla & ")"
            Exit Function
        Else
            
            If bShowDetail Then
                subLogErrMsg "Formula = " & sFmla
                subLogErrMsg "Result = " & dBonusAmt
            End If
            
        End If
        
    End If
    
    'reset the v1, v2, or v3 if they are "check_amount"
    If bShowDetail Then
    
        If sV1 = "check_amount" Or sV2 = "check_amount" Or sV3 = "check_amount" Then
            subLogErrMsg "check_amount = " & dBonusAmt
        End If
    
    End If
    
    If sV1 = "check_amount" Then
        V1 = dBonusAmt
        objMath.Var("v1") = V1
    End If
    
    If sV2 = "check_amount" Then
        V2 = dBonusAmt
        objMath.Var("v2") = V2
    End If
    
    If sV3 = "check_amount" Then
        V3 = dBonusAmt
        objMath.Var("v3") = V3
    End If
    
    'Adj. condition and formula
    If sAFmla <> "" Then
        
        If sACond = "" Then
            bConditionOK = True
        Else
            bConditionOK = objCond.CheckCondition(sACond, sErrMsg)
            
            If sErrMsg <> "" Then
                subLogErrMsg sErrMsg & ", Invalid Condition Clause (" & sACond & ")"
                Exit Function
            Else
                
                If bShowDetail Then
                    subLogErrMsg "Adj. Condition = " & sACond
                    subLogErrMsg "Result = " & IIf(bConditionOK, "True", "False")
                End If
                
            End If
            
        End If
    
        If bConditionOK Then
            dBonusAmt = tfnRound(objMath.Calculate(sAFmla, sErrMsg), DEFAULT_DECIMALS)
            
            If sErrMsg <> "" Then
                subLogErrMsg sErrMsg & ", Invalid Formula (" & sAFmla & ")"
                Exit Function
            Else
                
                If bShowDetail Then
                    subLogErrMsg "Formula = " & sAFmla
                    subLogErrMsg "Result = " & dBonusAmt
                End If
            
            End If
        
        End If
        
    End If
    
    fnCalculateBonus = dBonusAmt
End Function

Private Function fnGetVarValue(lEmpNo As Long, _
                               nPrftCtr As Integer, _
                               sV As String, _
                               ByVal sVariable As String, _
                               sErrMsg As String) As Double
                               
    Const SUB_NAME As String = "fnGetVarValue"
    
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim sVinV As String
    
    sVinV = sVariable + " in " + sV
    
    fnGetVarValue = 0#
    sErrMsg = ""
    
    sVariable = LCase(Trim(sVariable))
    
    If sVariable = "" Then
        'sErrMsg = "Variable is not defined in " + sV
        Exit Function
    End If
    
    Select Case sVariable
        Case "3_mo_shortage_avg"
            fnGetVarValue = fn3MonthsAverage(sVariable, sVinV, sErrMsg, lEmpNo, _
                tfnFormatDate(frmZZSEBPRC.txtStartDate))
            Exit Function
            
        Case "3_month_sales_avg"
            fnGetVarValue = fn3MonthsAverage(sVariable, sVinV, sErrMsg, lEmpNo, _
                tfnFormatDate(frmZZSEBPRC.txtStartDate))
            Exit Function

        Case "day_off_slip_day"
            strSQL = "SELECT COUNT (bd_prft_ctr) AS var_value "
            strSQL = strSQL & " FROM bonus_day_off_slip "
            strSQL = strSQL & " WHERE bd_empno = " & lEmpNo
            strSQL = strSQL & " AND bd_prft_ctr = " & nPrftCtr
            strSQL = strSQL & " AND bd_slip_date BETWEEN " & tfnDateString(frmZZSEBPRC.txtStartDate, True)
            strSQL = strSQL & " AND " & tfnDateString(frmZZSEBPRC.txtEndDate, True)
        
        Case "days_employed"
            fnGetVarValue = fnMonthsEmployed(sVinV, sErrMsg, lEmpNo, True)
            Exit Function
        
        Case "gallons_sold"
            fnGetVarValue = fnGallonsSold(sVinV, sErrMsg, lEmpNo)
            Exit Function
        
        Case "inside_sales"
            fnGetVarValue = fnInsideSales(sVinV, sErrMsg, lEmpNo)
            Exit Function
        
        Case "inv_record_months"
            fnGetVarValue = fnInvRecordMonths(sVinV, sErrMsg, lEmpNo)
            Exit Function
            
        Case "months_as_manager"
            fnGetVarValue = fnMonthsInGrade(sVinV, sErrMsg, lEmpNo, "M")
            Exit Function
        
        Case "months_employed"
            fnGetVarValue = fnMonthsEmployed(sVinV, sErrMsg, lEmpNo)
            Exit Function
            
        Case "ot_hours"
            If sPayCode_OtHrs = "" Then
                sErrMsg = "Pay Code for Overtime Hour not set up for " + sVinV
                Exit Function
            End If
            
            strSQL = "SELECT SUM(bh_hours) AS var_value  FROM bonus_hold "
            strSQL = strSQL & " WHERE bh_chk_link = 0"
            'david 05/10/2002 #368969
            'strSQL = strSQL & " AND bh_pay_code = " & tfnSQLString(sPayCode_OtHrs)
            strSQL = strSQL & " AND bh_pay_code IN (" & sPayCode_OtHrs + ")"
            strSQL = strSQL & " AND bh_empno = " & lEmpNo
            
        Case "regular_hours"
            If sPayCode_RegHrs = "" Then
                sErrMsg = "Pay Code for Overtime Hour not set up for " + sVinV
                Exit Function
            End If
            
            strSQL = "SELECT SUM(bh_hours) AS var_value  FROM bonus_hold "
            strSQL = strSQL & " WHERE bh_chk_link = 0"
            'david 05/10/2002 #368969
            'strSQL = strSQL & " AND bh_pay_code = " & tfnSQLString(sPayCode_RegHrs)
            strSQL = strSQL & " AND bh_pay_code IN (" & sPayCode_RegHrs + ")"
            strSQL = strSQL & " AND bh_empno = " & lEmpNo
                        
        Case "two_week_sales"
            fnGetVarValue = fnTwoWeekSales(sVinV, sErrMsg, lEmpNo)
            Exit Function
            
        Case "yrs_at_lvl_jan_1"
            fnGetVarValue = fnYearAtLevelJan1(sVinV, sErrMsg, lEmpNo)
            Exit Function
            
        Case "check_amount"
            'the value will be obtained from the formula evaluation
            'return 0
            Exit Function
        Case "asst_mgr_3_m_sales"
            fnGetVarValue = fnGetAsstMgr3MonthSales(sVinV, sErrMsg, lEmpNo)
            Exit Function
        Case "asst_mgr_gals_sold"
            strSQL = "SELECT bs_sales_amount as var_value FROM bonus_sales, pr_master, bonus_grades"
            strSQL = strSQL & " WHERE prm_empno = " & lEmpNo
            strSQL = strSQL & " AND bs_prft_ctr = prm_prft_ctr1"
            strSQL = strSQL & " AND prm_emp_level = bg_emp_level"
            strSQL = strSQL & " AND bg_grade = 'A' "
            strSQL = strSQL & " AND bs_sales_type = " & tfnSQLString(sGas)
            strSQL = strSQL & " AND bs_from_date = " & tfnDateString(frmZZSEBPRC!txtStartDate, True)

        Case "not used"
            'return 0
            Exit Function
                        
        Case Else
            sErrMsg = "Variable " + sVinV + " is not valid"
            Exit Function
    
    End Select
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        sErrMsg = "Failed to access the database to get " & tfnSQLString(sVariable)
        Exit Function
    End If
    
    If rsTemp.RecordCount = 0 Then
        sErrMsg = "No record found for " & sVinV
        bNoRecordFound = True
        Exit Function
    End If
    
    If rsTemp.RecordCount > 0 Then
        fnGetVarValue = tfnRound(rsTemp!var_value, DEFAULT_DECIMALS)
    End If
    
End Function

Private Function fn3MonthsAverage(sVariable As String, _
                                  sVinV As String, _
                                  sErrMsg As String, _
                                  lEmpNo As Long, _
                                  sDateStart As String) As Double

    Const SUB_NAME As String = "fn3MonthsAverage"
    
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim sTemp As String
    Dim sDateHired  As String
    Dim sDatePrev As String
    Dim sFirstDate As String
    Dim nPrftCtr As Integer
    Dim nPrevPrftCtr As Integer
    Dim i As Long
    Dim nMonthCount As Integer
    Dim dTmpAmt As Double
    Dim dAmount As Double
    
    'get hired date
    strSQL = "SELECT prm_date_hired, prm_prft_ctr1"
    strSQL = strSQL & " FROM pr_master"
    strSQL = strSQL & " WHERE prm_empno = " & lEmpNo
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        sErrMsg = "Failed to access the database to get " & sVinV
        Exit Function
    End If
    
    If rsTemp.RecordCount = 0 Then
        sErrMsg = "Employee record not found for " & sVinV
        Exit Function
    End If
    
    sDateHired = fnGetField(rsTemp!prm_date_hired)
    nPrftCtr = tfnRound(rsTemp!prm_prft_ctr1)
            
    subLogErrMsg "Current Profit Center " & nPrftCtr & "."
    
    dAmount = 0#
    nMonthCount = 0
    
    sFirstDate = Year(CDate(sDateStart)) & "/" & Month(CDate(sDateStart)) & "/" & "01"
    
    For i = 1 To 3 '3 months
        
        Select Case i
            Case 1
                sTemp = "Previous Month"
            Case 2
                sTemp = "2 Months ago"
            Case 3
                sTemp = "3 Months ago"
        End Select
        
        'get the profit center of the employee worked for
        Select Case sVariable
            Case "3_mo_shortage_avg"
                sDatePrev = DateAdd("m", -i, CDate(sFirstDate))
            ' the first month is current month
            Case "3_month_sales_avg"
                sDatePrev = DateAdd("m", -i + 1, CDate(sFirstDate))
        End Select
        
        strSQL = "SELECT prhs_effective_dt, prhs_prft_ctr1"
        strSQL = strSQL & " FROM pr_history, bonus_grades"
        strSQL = strSQL & " WHERE prhs_empno = " & lEmpNo
        strSQL = strSQL & " AND prhs_effective_dt <= " & tfnDateString(sDatePrev, True)
        
        Select Case sVariable
            Case "3_mo_shortage_avg"
                strSQL = strSQL & " AND prhs_emp_level =bg_emp_level AND bg_grade IN ('M', 'A', 'N')"
            Case "3_month_sales_avg"
                strSQL = strSQL & " AND prhs_emp_level =bg_emp_level AND bg_grade IN ('M', 'A')"
        End Select
        
        strSQL = strSQL & " ORDER BY prhs_effective_dt DESC"
        
        If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
            sErrMsg = "Failed to access the database to get " & sVinV
            Exit Function
        End If
        
        If rsTemp.RecordCount = 0 Then
            'subLogErrMsg "History not found for " + tfnDateString(sDatePrev, True) + ", use Date Hired."
            Select Case sVariable
                Case "3_mo_shortage_avg"
                    subLogErrMsg "No history found for " & lEmpNo & " as manager, assistant manager or night clerk before " & tfnDateString(sDatePrev, True)
                Case "3_month_sales_avg"
                    subLogErrMsg "No history found for " & lEmpNo & " as manager or assistant manager before " & tfnDateString(sDatePrev, True)
            End Select
            
            
            
            Exit For
            
'            If IsValidDate(sDateHired) Then
'
'                If CDate(sDateHired) <= CDate(sDatePrev) Then
'                    nPrevPrftCtr = nPrftCtr
'                Else
'                    'changed by junsong, if employee only hired one month ago, we need its value
'                    'sErrMsg = "Date Hired is later than " + tfnDateString(sDatePrev, True) + " to get " & sVinV
'                    subLogErrMsg "Date Hired is later than " + tfnDateString(sDatePrev, True)
'                    Exit For
'                    'Exit Function
'                End If
'
'            Else
'                sErrMsg = "Date Hired is not valid"
'                Exit Function
'            End If
            
        Else
            If bShowDetail Then
                subLogErrMsg "Effective Date " + tfnDateString(rsTemp!prhs_effective_dt, True) + " in History found."
            End If
            
            If CDate(tfnDateString(rsTemp!prhs_effective_dt)) <= CDate(sDatePrev) Then
                nPrevPrftCtr = tfnRound(rsTemp!prhs_prft_ctr1)
            Else
                subLogErrMsg "Effective Date is later than " + tfnDateString(sDatePrev, True)
                Exit For
            End If
            
        End If
        
        If nPrftCtr <> nPrevPrftCtr Then
            subLogErrMsg sTemp + " Profit Center " & nPrevPrftCtr & "."
        End If
        
        Select Case sVariable
        Case "3_mo_shortage_avg"
            strSQL = "SELECT bs_sales_amount AS var_value "
            strSQL = strSQL & " FROM bonus_sales"
            strSQL = strSQL & " WHERE bs_prft_ctr = " & nPrevPrftCtr
            strSQL = strSQL & " AND bs_from_date <= " & tfnDateString(sDatePrev, True)
            strSQL = strSQL & " AND bs_to_date >= " & tfnDateString(sDatePrev, True)
            strSQL = strSQL & " AND bs_sales_type = " & tfnSQLString(sRatio)
        Case "3_month_sales_avg"
            strSQL = "SELECT bs_sales_amount AS var_value "
            strSQL = strSQL & " FROM bonus_sales"
            strSQL = strSQL & " WHERE bs_prft_ctr = " & nPrevPrftCtr
            strSQL = strSQL & " AND bs_from_date <= " & tfnDateString(sDatePrev, True)
            strSQL = strSQL & " AND bs_to_date >= " & tfnDateString(sDatePrev, True)
            strSQL = strSQL & " AND bs_sales_type = " & tfnSQLString(sOneMth)
        End Select
        
        If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
            sErrMsg = "Failed to access the database to get " & sVinV
            Exit Function
        End If
        
        If rsTemp.RecordCount = 0 Then
            subLogErrMsg "No record found as of " + tfnDateString(sDatePrev) + " for " & sVinV & "."
            bNoRecordFound = True
        Else
            Select Case sVariable
                Case "3_mo_shortage_avg"
                    dTmpAmt = tfnRound(rsTemp!var_value, 2)
                    
                    If bShowDetail Then
                        subLogErrMsg "Shortage Ratio as of " + tfnDateString(sDatePrev) + " = " & dTmpAmt
                    End If
                    
                Case "3_month_sales_avg"
                    dTmpAmt = tfnRound(rsTemp!var_value, 2)
                    
                    If bShowDetail Then
                        subLogErrMsg "Inside Sales as of " + tfnDateString(sDatePrev) + " = " & dTmpAmt
                    End If
                    
            End Select
            
            dAmount = dAmount + dTmpAmt
            nMonthCount = nMonthCount + 1
        End If
    
    Next i
    
    If nMonthCount > 0 Then
        fn3MonthsAverage = tfnRound(dAmount / nMonthCount, 2)
    Else
        fn3MonthsAverage = 0#
    End If
    
End Function

Private Function fnInsideSales(sVinV As String, _
                                 sErrMsg As String, _
                                 lEmpNo As Long) As Double
    Dim strSQL As String
    Dim rsTemp As Recordset
    Const FUNC_NAME As String = "fnInsideSales"
    
'    strSQL = "SELECT prhs_effective_dt, prhs_prft_ctr1 FROM pr_history, bonus_grades, bonus_sales"
'    strSQL = strSQL & " WHERE prhs_effective_dt <= bs_from_date "
'    strSQL = strSQL & " AND bs_sales_type = " & tfnSQLString(sOneMth)
'    strSQL = strSQL & " AND bs_from_date <= " & tfnDateString(DateAdd("M", -1, CDate(frmZZSEBPRC.txtStartDate)), True)
'    strSQL = strSQL & " AND bs_to_date >= " & tfnDateString(DateAdd("M", -1, CDate(frmZZSEBPRC.txtStartDate)), True)
'    strSQL = strSQL & " AND prhs_emp_level = bg_emp_level AND bg_grade = 'A' "
'    strSQL = strSQL & " AND prhs_empno = " & lEmpNo
'    strSQL = strSQL & " ORDER BY prhs_effective_dt DESC"
'
'    If GetRecordSet(rsTemp, strSQL, , FUNC_NAME) < 0 Then
'        sErrMsg = "Failed to access the database to get " & sVinV
'        Exit Function
'    End If
'
'    If rsTemp.RecordCount = 0 Then
'        subLogErrMsg "No History found for " & lEmpNo & " to get inside sales."
'        fnInsideSales = 0
'        Exit Function
'    End If
'
'    strSQL = "SELECT bs_sales_amount FROM bonus_sales "
'    strSQL = strSQL & " WHERE bs_prft_ctr = " & tfnRound(rsTemp!prhs_prft_ctr1)
'    strSQL = strSQL & " AND bs_from_date <= " & tfnDateString(DateAdd("M", -1, CDate(frmZZSEBPRC.txtStartDate)), True)
'    strSQL = strSQL & " AND bs_to_date >= " & tfnDateString(DateAdd("M", -1, CDate(frmZZSEBPRC.txtStartDate)), True)
'    strSQL = strSQL & " AND bs_sales_type = " & tfnSQLString(sOneMth)
'
    strSQL = "SELECT bs_sales_amount FROM bonus_sales, pr_master, bonus_grades"
    strSQL = strSQL & " WHERE prm_empno = " & lEmpNo
    strSQL = strSQL & " AND bs_prft_ctr = prm_prft_ctr1 "
    strSQL = strSQL & " AND prm_emp_level = bg_emp_level"
    strSQL = strSQL & " AND bg_grade = 'A' "
    strSQL = strSQL & " AND bs_from_date <= " & tfnDateString(DateAdd("M", -1, CDate(frmZZSEBPRC.txtStartDate)), True)
    strSQL = strSQL & " AND bs_to_date >= " & tfnDateString(DateAdd("M", -1, CDate(frmZZSEBPRC.txtStartDate)), True)
    strSQL = strSQL & " AND bs_sales_type = " & tfnSQLString(sOneMth)

    If GetRecordSet(rsTemp, strSQL, , FUNC_NAME) < 0 Then
        sErrMsg = "Failed to access the database to get " & sVinV
        Exit Function
    End If
    
    If rsTemp.RecordCount = 0 Then
        sErrMsg = "No record found for " & sVinV
        bNoRecordFound = True
        fnInsideSales = 0#
        Exit Function
    Else
        fnInsideSales = tfnRound(rsTemp!bs_sales_amount, 2)
    End If
    
End Function


Private Function fnTwoWeekSales(sVinV As String, _
                                 sErrMsg As String, _
                                 lEmpNo As Long) As Double
    Dim strSQL As String
    Dim rsTemp As Recordset
    Const FUNC_NAME As String = "fnTwoWeekSales"
    
    strSQL = "SELECT prhs_effective_dt, prhs_prft_ctr1 FROM pr_history, bonus_grades, bonus_sales"
    strSQL = strSQL & " WHERE prhs_effective_dt <= bs_from_date "
    strSQL = strSQL & " AND bs_from_date = " & tfnDateString(frmZZSEBPRC.txtStartDate, True)
    strSQL = strSQL & " AND bs_sales_type = " & tfnSQLString(sBiWeek)
    strSQL = strSQL & " AND prhs_emp_level = bg_emp_level AND bg_grade = 'M' "
    strSQL = strSQL & " AND prhs_empno = " & lEmpNo
    strSQL = strSQL & " ORDER BY prhs_effective_dt DESC"

    If GetRecordSet(rsTemp, strSQL, , FUNC_NAME) < 0 Then
        sErrMsg = "Failed to access the database to get " & sVinV
        Exit Function
    End If
    
    If rsTemp.RecordCount = 0 Then
        subLogErrMsg "No History found for " & lEmpNo & " to get two weeks sales."
        fnTwoWeekSales = 0#
        Exit Function
    End If
            
    strSQL = "SELECT bs_sales_amount FROM bonus_sales "
    strSQL = strSQL & " WHERE bs_prft_ctr = " & tfnRound(rsTemp!prhs_prft_ctr1)
    strSQL = strSQL & " AND bs_from_date = " & tfnDateString(frmZZSEBPRC.txtStartDate, True)
    strSQL = strSQL & " AND bs_sales_type = " & tfnSQLString(sBiWeek)
    
    If GetRecordSet(rsTemp, strSQL, , FUNC_NAME) < 0 Then
        sErrMsg = "Failed to access the database to get " & sVinV
        Exit Function
    End If
    
    If rsTemp.RecordCount = 0 Then
        sErrMsg = "No record found for " & sVinV
        bNoRecordFound = True
        fnTwoWeekSales = 0#
        Exit Function
    Else
        fnTwoWeekSales = tfnRound(rsTemp!bs_sales_amount, 2)
    End If
    
End Function

Private Function fnGallonsSold(sVinV As String, _
                                 sErrMsg As String, _
                                 lEmpNo As Long) As Double
    Dim strSQL As String
    Dim rsTemp As Recordset
    Const FUNC_NAME As String = "fnGallonsSold"
    
    strSQL = "SELECT prhs_effective_dt, prhs_prft_ctr1 FROM pr_history, bonus_grades, bonus_sales"
    strSQL = strSQL & " WHERE prhs_effective_dt <= bs_from_date "
    strSQL = strSQL & " AND bs_from_date = " & tfnDateString(frmZZSEBPRC.txtStartDate, True)
    strSQL = strSQL & " AND bs_sales_type = " & tfnSQLString(sGas)
    strSQL = strSQL & " AND prhs_emp_level = bg_emp_level AND bg_grade IN ('M', 'A') "
    strSQL = strSQL & " AND prhs_empno = " & lEmpNo
    strSQL = strSQL & " ORDER BY prhs_effective_dt DESC"

    If GetRecordSet(rsTemp, strSQL, , FUNC_NAME) < 0 Then
        sErrMsg = "Failed to access the database to get " & sVinV
        Exit Function
    End If
    
    If rsTemp.RecordCount = 0 Then
        subLogErrMsg "No History found for " & lEmpNo & " to get Gallons Sold."
        fnGallonsSold = 0#
        Exit Function
    End If
            
    strSQL = "SELECT bs_sales_amount FROM bonus_sales "
    strSQL = strSQL & " WHERE bs_prft_ctr = " & tfnRound(rsTemp!prhs_prft_ctr1)
    strSQL = strSQL & " AND bs_from_date = " & tfnDateString(frmZZSEBPRC.txtStartDate, True)
    strSQL = strSQL & " AND bs_sales_type = " & tfnSQLString(sGas)
    
    If GetRecordSet(rsTemp, strSQL, , FUNC_NAME) < 0 Then
        sErrMsg = "Failed to access the database to get " & sVinV
        Exit Function
    End If
    
    If rsTemp.RecordCount = 0 Then
        sErrMsg = "No record found for " & sVinV
        bNoRecordFound = True
        fnGallonsSold = 0#
        Exit Function
    Else
        fnGallonsSold = tfnRound(rsTemp!bs_sales_amount, 2)
    End If
    
End Function

Private Function fnInvRecordMonths(sVinV As String, _
                                 sErrMsg As String, _
                                 lEmpNo As Long) As Double
                                 
    Const SUB_NAME As String = "fnInvRecordMonths"
    Const sGradeList As String = "('M', 'A', 'N')"
    
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim nEmpLevel As Integer
    Dim aryEmpLevelList()
    Dim sDateHired As String
    Dim sDateTerminated As String
    Dim sDateStart As String
    Dim sDateEnd As String
    Dim dDiff As Double
    Dim i As Long
    Dim nEmpLevelCount As Integer
    
    fnInvRecordMonths = 0#
    
    'get the employee level list for the Grade
    nEmpLevelCount = -1
    ReDim aryEmpLevelList(0)
    
    strSQL = "SELECT bg_emp_level"
    strSQL = strSQL & " FROM bonus_grades"
    strSQL = strSQL & " WHERE bg_grade IN " + sGradeList
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        sErrMsg = "Failed to access the database to get " & sVinV
        Exit Function
    End If
    
    If rsTemp.RecordCount = 0 Then
        sErrMsg = "Grade record not found for " & sVinV
        Exit Function
    End If
    
    For i = 1 To rsTemp.RecordCount
        
        If Not IsNull(rsTemp!bg_emp_level) Then
            nEmpLevelCount = nEmpLevelCount + 1
            ReDim Preserve aryEmpLevelList(nEmpLevelCount)
            aryEmpLevelList(nEmpLevelCount) = tfnRound(rsTemp!bg_emp_level)
        End If
        
        rsTemp.MoveNext
    Next i
    
    strSQL = "SELECT prm_emp_level, prm_date_hired, prm_date_termed"
    strSQL = strSQL & " FROM pr_master"
    strSQL = strSQL & " WHERE prm_empno = " & lEmpNo
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        sErrMsg = "Failed to access the database to get " & sVinV
        Exit Function
    End If
    
    If rsTemp.RecordCount = 0 Then
        sErrMsg = "Employee record not found for " & sVinV
        Exit Function
    End If
    
    If IsNull(rsTemp!prm_emp_level) Then
        sErrMsg = "Employee level is NULL for " & sVinV
        Exit Function
    End If
    
    nEmpLevel = tfnRound(rsTemp!prm_emp_level)
    sDateHired = fnGetField(rsTemp!prm_date_hired)
    sDateTerminated = fnGetField(rsTemp!prm_date_termed)
    
    strSQL = "SELECT prhs_effective_dt, prhs_emp_level, prhs_date_hired, prhs_date_termed"
    strSQL = strSQL & " FROM pr_history"
    strSQL = strSQL & " WHERE prhs_empno = " & lEmpNo
    strSQL = strSQL & " AND prhs_effective_dt <= " & tfnDateString(frmZZSEBPRC!txtStartDate, True)
    strSQL = strSQL & " ORDER BY prhs_effective_dt"
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        sErrMsg = "Failed to access the database to get " & sVinV
        Exit Function
    End If
    
    If rsTemp.RecordCount = 0 Then
        sErrMsg = "Employee history record not found for " & sVinV
        
        'use date hired and/or date terminated for calculation
        If Not IsValidDate(sDateHired) Then
            sErrMsg = "Date Hired is not valid for " & sVinV
            Exit Function
        End If
        
        If nEmpLevelCount < 0 Or fnFindInList(nEmpLevel, aryEmpLevelList, nEmpLevelCount) Then
            
            If Not IsValidDate(sDateTerminated) Then
                'changed by junsong using the first day of the month
                sDateTerminated = Year(CDate(frmZZSEBPRC!txtStartDate)) & "/" & Month(CDate(frmZZSEBPRC!txtStartDate)) & "/" & "01"
                
                'sDateTerminated = frmZZSEBPRC!txtStartDate
            End If
            
            dDiff = Int(fnDateDiff("m", CDate(sDateHired), CDate(sDateTerminated)))
        End If
            
        fnInvRecordMonths = dDiff
        
        If dDiff < 0 Then
            dDiff = 0
        End If
        
        Exit Function
    End If
    
    If rsTemp.RecordCount = 1 Then
    
        If (nEmpLevelCount < 0 And tfnRound(rsTemp!prhs_emp_level) = nEmpLevel) Or _
           (fnFindInList(tfnRound(rsTemp!prhs_emp_level), aryEmpLevelList, nEmpLevelCount)) Then
            
            sDateStart = fnGetField(rsTemp!prhs_effective_dt)
            
            If Not IsValidDate(sDateStart) Then
                sErrMsg = "Effective Date is not valid for " & sVinV
                Exit Function
            End If
            
            'sDateEnd = frmZZSEBPRC!txtStartDate
            sDateEnd = Year(CDate(frmZZSEBPRC!txtStartDate)) & "/" & Month(CDate(frmZZSEBPRC!txtStartDate)) & "/" & "01"
            dDiff = fnDateDiff("m", CDate(sDateStart), CDate(sDateEnd))
        End If
    
    Else
        dDiff = 0
        i = 1
        
        Do
            If (nEmpLevelCount < 0 And tfnRound(rsTemp!prhs_emp_level) = nEmpLevel) Or _
               (fnFindInList(tfnRound(rsTemp!prhs_emp_level), aryEmpLevelList, nEmpLevelCount)) Then
                sDateStart = fnGetField(rsTemp!prhs_effective_dt)
                
                If IsValidDate(sDateStart) Then
                    
                    If i <= rsTemp.RecordCount - 1 Then
                        rsTemp.MoveNext
                        i = i + 1
                        sDateEnd = fnGetField(rsTemp!prhs_effective_dt)
                        
                        If Not IsValidDate(sDateEnd) Then
                            sErrMsg = "Effective Date is not valid for " & sVinV
                            Exit Function
                        End If
                        
                        dDiff = dDiff + fnDateDiff("m", CDate(sDateStart), CDate(sDateEnd))
                    Else
                        sDateStart = fnGetField(rsTemp!prhs_effective_dt)
                        Exit Do
                    End If
                
                Else
                    sDateStart = ""
                    sErrMsg = "Effective Date is not valid for " & sVinV
                    Exit Function
                End If
                
            Else
                sDateStart = ""
                rsTemp.MoveNext
                i = i + 1
            End If
            
        Loop Until rsTemp.EOF
        
        'last record - from last effective date until now
        If sDateStart <> "" Then
            
            If (nEmpLevelCount < 0 And tfnRound(rsTemp!prhs_emp_level) = nEmpLevel) Or _
               (fnFindInList(tfnRound(rsTemp!prhs_emp_level), aryEmpLevelList, nEmpLevelCount)) Then
                'sDateEnd = frmZZSEBPRC!txtStartDate
                sDateEnd = Year(CDate(frmZZSEBPRC!txtStartDate)) & "/" & Month(CDate(frmZZSEBPRC!txtStartDate)) & "/" & "01"
                dDiff = dDiff + fnDateDiff("m", CDate(sDateStart), CDate(sDateEnd))
            End If
            
        End If
        
    End If
    
    fnInvRecordMonths = Int(dDiff)
End Function

Private Function fnYearAtLevelJan1(sVinV As String, _
                                 sErrMsg As String, _
                                 lEmpNo As Long) As Double
                                 
    Const SUB_NAME As String = "fnYearAtLevelJan1"
    Const sGradeManager As String = "('M')"
    Const sGradeAsstManager As String = "('A', 'N')"
    
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim bIsManager As Boolean
    Dim sDateHired As String
    Dim sDateStart As String
    Dim dDiff As Double
    
    fnYearAtLevelJan1 = 0#
    
    'get employee's position based on the current employee level
    strSQL = "SELECT prm_emp_level, prm_date_hired"
    strSQL = strSQL & " FROM pr_master"
    strSQL = strSQL & " WHERE prm_empno = " & lEmpNo
    strSQL = strSQL & " AND prm_emp_level IN ("
    strSQL = strSQL & "SELECT bg_emp_level"
    strSQL = strSQL & " FROM bonus_grades"
    strSQL = strSQL & " WHERE bg_grade IN " + sGradeManager + ")"
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        sErrMsg = "Failed to access the database to get " & sVinV
        Exit Function
    End If
    
    If rsTemp.RecordCount = 0 Then
        bIsManager = False
        
        strSQL = "SELECT prm_emp_level, prm_date_hired"
        strSQL = strSQL & " FROM pr_master"
        strSQL = strSQL & " WHERE prm_empno = " & lEmpNo
        strSQL = strSQL & " AND prm_emp_level IN ("
        strSQL = strSQL & "SELECT bg_emp_level"
        strSQL = strSQL & " FROM bonus_grades"
        strSQL = strSQL & " WHERE bg_grade IN " + sGradeAsstManager + ")"
        
        If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
            sErrMsg = "Failed to access the database to get " & sVinV
            Exit Function
        End If
    
        If rsTemp.RecordCount = 0 Then
            sErrMsg = "Employee is not Manager, Assistant Manager, or Night Manager"
            Exit Function
        End If
    Else
        bIsManager = True
    End If
    
    sDateHired = fnGetField(rsTemp!prm_date_hired)
    sDateStart = frmZZSEBPRC!txtStartDate
    
    strSQL = "SELECT prhs_effective_dt, prhs_emp_level, prhs_date_hired, prhs_date_termed"
    strSQL = strSQL & " FROM pr_history"
    strSQL = strSQL & " WHERE prhs_empno = " & lEmpNo
    
    If bIsManager Then
        strSQL = strSQL & " AND prhs_emp_level IN ("
        strSQL = strSQL & " SELECT bg_emp_level"
        strSQL = strSQL & " FROM bonus_grades"
        strSQL = strSQL & " WHERE bg_grade IN " & sGradeManager & ")"
    Else
        strSQL = strSQL & " AND prhs_emp_level IN ("
        strSQL = strSQL & " SELECT bg_emp_level"
        strSQL = strSQL & " FROM bonus_grades"
        strSQL = strSQL & " WHERE bg_grade IN " & sGradeAsstManager & ")"
    End If
    
    strSQL = strSQL & " AND prhs_effective_dt <= " & tfnDateString(sDateStart, True)
    strSQL = strSQL & " ORDER BY prhs_effective_dt"
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        sErrMsg = "Failed to access the database to get " & sVinV
        Exit Function
    End If
    
    If rsTemp.RecordCount = 0 Then
        subLogErrMsg "History not found for " + tfnDateString(sDateStart, True) + ", use Date Hired."
        
        If IsValidDate(sDateHired) Then
            
            If CDate(sDateHired) <= CDate(sDateStart) Then
                dDiff = Int(fnDateDiff("yyyy", CDate(sDateHired), CDate(sDateStart), _
                    vbFirstJan1))
            Else
                sErrMsg = "Date Hired is later than " + tfnDateString(sDateStart, True) + " to get " & sVinV
                Exit Function
            End If
        
        Else
            sErrMsg = "Date Hired is not valid"
            Exit Function
        End If
    
    Else
        
        If bShowDetail Then
            subLogErrMsg "Effective Date " + tfnDateString(rsTemp!prhs_effective_dt, True) + " in History found."
        End If
        
        If IsValidDate(rsTemp!prhs_effective_dt) Then
            dDiff = Int(fnDateDiff("yyyy", CDate(rsTemp!prhs_effective_dt), CDate(sDateStart), _
                vbFirstJan1))
        Else
            sErrMsg = "Effective Hired is not valid"
            Exit Function
        End If
    
    End If
    
    fnYearAtLevelJan1 = dDiff
End Function

Private Function fnMonthsInGrade(sVinV As String, _
                                 sErrMsg As String, _
                                 lEmpNo As Long, _
                                 Optional sGrade As String = "", _
                                 Optional bConvertToYear As Boolean = False) As Double
                                 
    Const SUB_NAME As String = "fnMonthsInGrade"
    
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim nEmpLevel As Integer
    Dim aryEmpLevelList()
    Dim sDateHired As String
    Dim sDateTerminated As String
    Dim sDateStart As String
    Dim sDateEnd As String
    Dim dDiff As Double
    Dim i As Long
    Dim nEmpLevelCount As Integer
    
    fnMonthsInGrade = 0#
    
    'get the employee level list for the Grade
    nEmpLevelCount = -1
    ReDim aryEmpLevelList(0)
    
    If sGrade <> "" Then
        strSQL = "SELECT bg_emp_level"
        strSQL = strSQL & " FROM bonus_grades"
        strSQL = strSQL & " WHERE bg_grade = " + tfnSQLString(sGrade)
        
        If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
            sErrMsg = "Failed to access the database to get " & sVinV
            Exit Function
        End If
        
        If rsTemp.RecordCount = 0 Then
            sErrMsg = "Grade record not found for " & sVinV
            Exit Function
        End If
        
        For i = 1 To rsTemp.RecordCount
            
            If Not IsNull(rsTemp!bg_emp_level) Then
                nEmpLevelCount = nEmpLevelCount + 1
                ReDim Preserve aryEmpLevelList(nEmpLevelCount)
                aryEmpLevelList(nEmpLevelCount) = tfnRound(rsTemp!bg_emp_level)
            End If
            
            rsTemp.MoveNext
        Next i
    
    End If
    
    strSQL = "SELECT prm_emp_level, prm_date_hired, prm_date_termed"
    strSQL = strSQL & " FROM pr_master"
    strSQL = strSQL & " WHERE prm_empno = " & lEmpNo
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        sErrMsg = "Failed to access the database to get " & sVinV
        Exit Function
    End If
    
    If rsTemp.RecordCount = 0 Then
        sErrMsg = "Employee record not found for " & sVinV
        Exit Function
    End If
    
    If IsNull(rsTemp!prm_emp_level) Then
        sErrMsg = "Employee level is NULL for " & sVinV
        Exit Function
    End If
    
    nEmpLevel = tfnRound(rsTemp!prm_emp_level)
    sDateHired = fnGetField(rsTemp!prm_date_hired)
    sDateTerminated = fnGetField(rsTemp!prm_date_termed)
    
    strSQL = "SELECT prhs_effective_dt, prhs_emp_level, prhs_date_hired, prhs_date_termed"
    strSQL = strSQL & " FROM pr_history"
    strSQL = strSQL & " WHERE prhs_empno = " & lEmpNo
    strSQL = strSQL & " AND prhs_effective_dt <= " & tfnDateString(frmZZSEBPRC!txtStartDate, True)
    strSQL = strSQL & " ORDER BY prhs_effective_dt"
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        sErrMsg = "Failed to access the database to get " & sVinV
        Exit Function
    End If
    
    If rsTemp.RecordCount = 0 Then
        sErrMsg = "Employee history record not found for " & sVinV
        
        'use date hired and/or date terminated for calculation
        If Not IsValidDate(sDateHired) Then
            sErrMsg = "Date Hired is not valid for " & sVinV
            Exit Function
        End If
        
        If nEmpLevelCount < 0 Or fnFindInList(nEmpLevel, aryEmpLevelList, nEmpLevelCount) Then
            
            If Not IsValidDate(sDateTerminated) Then
                sDateTerminated = frmZZSEBPRC!txtStartDate
            End If
            
            dDiff = fnDateDiff("m", CDate(sDateHired), CDate(sDateTerminated))
        End If
        
        If bConvertToYear Then
            fnMonthsInGrade = dDiff / 12
        Else
            fnMonthsInGrade = dDiff
        End If
        
        Exit Function
    End If
    
    If rsTemp.RecordCount = 1 Then
        
        If (nEmpLevelCount < 0 And tfnRound(rsTemp!prhs_emp_level) = nEmpLevel) Or _
           (fnFindInList(tfnRound(rsTemp!prhs_emp_level), aryEmpLevelList, nEmpLevelCount)) Then
            sDateStart = fnGetField(rsTemp!prhs_effective_dt)
            
            If Not IsValidDate(sDateStart) Then
                sErrMsg = "Effective Date is not valid for " & sVinV
                Exit Function
            End If
            
            sDateEnd = frmZZSEBPRC!txtStartDate
            dDiff = fnDateDiff("m", CDate(sDateStart), CDate(sDateEnd))
        End If
    
    Else
        dDiff = 0
        i = 1
        Do
            If (nEmpLevelCount < 0 And tfnRound(rsTemp!prhs_emp_level) = nEmpLevel) Or _
               (fnFindInList(tfnRound(rsTemp!prhs_emp_level), aryEmpLevelList, nEmpLevelCount)) Then
                sDateStart = fnGetField(rsTemp!prhs_effective_dt)
                
                If IsValidDate(sDateStart) Then
                    
                    If i <= rsTemp.RecordCount - 1 Then
                        rsTemp.MoveNext
                        i = i + 1
                        sDateEnd = fnGetField(rsTemp!prhs_effective_dt)
                        
                        If Not IsValidDate(sDateEnd) Then
                            sErrMsg = "Effective Date is not valid for " & sVinV
                            Exit Function
                        End If
                        
                        dDiff = dDiff + fnDateDiff("m", CDate(sDateStart), CDate(sDateEnd))
                    Else
                        sDateStart = fnGetField(rsTemp!prhs_effective_dt)
                        Exit Do
                    End If
                
                Else
                    sDateStart = ""
                    sErrMsg = "Effective Date is not valid for " & sVinV
                    Exit Function
                End If
            
            Else
                sDateStart = ""
                rsTemp.MoveNext
                i = i + 1
            End If
        
        Loop Until rsTemp.EOF
        
        'last record - from last effective date until now
        If sDateStart <> "" Then
            
            If (nEmpLevelCount < 0 And tfnRound(rsTemp!prhs_emp_level) = nEmpLevel) Or _
               (fnFindInList(tfnRound(rsTemp!prhs_emp_level), aryEmpLevelList, nEmpLevelCount)) Then
                sDateEnd = frmZZSEBPRC!txtStartDate
                dDiff = dDiff + fnDateDiff("m", CDate(sDateStart), CDate(sDateEnd))
            End If
        
        End If
    
    End If
    
    If bConvertToYear Then
        fnMonthsInGrade = dDiff / 12
    Else
        fnMonthsInGrade = dDiff
    End If
    
End Function

Private Function fnGetAsstMgr3MonthSales(sVinV As String, sErrMsg As String, _
                                        lEmpNo As Long) As Double

    Dim strSQL As String
    Dim strSubSql As String
    Dim rsTemp As Recordset
    Dim dTotalSales As Double
    Dim nCount As Double
    Dim sStartDate As String
    Const FUNC_NAME As String = "fnGetAsstMgr3MonthSales"
    
    sStartDate = frmZZSEBPRC!txtStartDate
    
    strSubSql = "SELECT bs_sales_amount FROM bonus_sales, pr_master, bonus_grades"
    strSubSql = strSubSql & " WHERE prm_empno = " & lEmpNo
    strSubSql = strSubSql & " AND bs_prft_ctr = prm_prft_ctr1"
    strSubSql = strSubSql & " AND prm_emp_level = bg_emp_level"
    strSubSql = strSubSql & " AND bg_grade = 'A' "
    strSubSql = strSubSql & " AND bs_sales_type = " & tfnSQLString(sOneMth)
    
    'preious month
    strSQL = strSubSql & " AND bs_from_date <= " & tfnDateString(sStartDate, True)
    strSQL = strSQL & " AND bs_to_date >= " & tfnDateString(sStartDate, True)
    
    If GetRecordSet(rsTemp, strSQL, nDB_REMOTE, FUNC_NAME) < 0 Then
        sErrMsg = "Failed to access the database to get " & sVinV
        Exit Function
    ElseIf rsTemp.RecordCount = 0 Then
        dTotalSales = dTotalSales + 0#
    Else
        dTotalSales = dTotalSales + tfnRound(rsTemp!bs_sales_amount, 2)
        nCount = nCount + 1#
    End If
    
    '2 months ago
    sStartDate = CStr(DateAdd("M", -1, CDate(sStartDate)))
    strSQL = strSubSql & " AND bs_from_date <= " & tfnDateString(sStartDate, True)
    strSQL = strSQL & " AND bs_to_date >= " & tfnDateString(sStartDate, True)
    
    If GetRecordSet(rsTemp, strSQL, nDB_REMOTE, FUNC_NAME) < 0 Then
        sErrMsg = "Failed to access the database to get " & sVinV
        Exit Function
    ElseIf rsTemp.RecordCount = 0 Then
        dTotalSales = dTotalSales + 0#
    Else
        dTotalSales = dTotalSales + tfnRound(rsTemp!bs_sales_amount, 2)
        nCount = nCount + 1#
    End If
    

    '3 months ago
    sStartDate = CStr(DateAdd("M", -1, CDate(sStartDate)))
    strSQL = strSubSql & " AND bs_from_date <= " & tfnDateString(sStartDate, True)
    strSQL = strSQL & " AND bs_to_date >= " & tfnDateString(sStartDate, True)
    
    If GetRecordSet(rsTemp, strSQL, nDB_REMOTE, FUNC_NAME) < 0 Then
        sErrMsg = "Failed to access the database to get " & sVinV
        Exit Function
    ElseIf rsTemp.RecordCount = 0 Then
        dTotalSales = dTotalSales + 0#
    Else
        dTotalSales = dTotalSales + tfnRound(rsTemp!bs_sales_amount, 2)
        nCount = nCount + 1#
    End If
    
    If nCount = 0 Then
        fnGetAsstMgr3MonthSales = 0#
    Else
        fnGetAsstMgr3MonthSales = tfnRound(dTotalSales / nCount, 2)
    End If
    
End Function
Private Function fnFindInList(vItemToFind As Variant, _
                                aryList() As Variant, _
                                nListCount As Integer) As Boolean
    Dim i As Integer
    Dim sItemToFind As String
    Dim lItemToFind As Long
    Dim bIsStringType As Boolean
    
    If VarType(vItemToFind) = vbString Then
        bIsStringType = True
    End If
    
    For i = 0 To nListCount
        
        If bIsStringType Then
            
            If CStr(aryList(i)) = CStr(vItemToFind) Then
                fnFindInList = True
                Exit Function
            End If
        
        Else
            
            If Val(aryList(i)) = Val(vItemToFind) Then
                fnFindInList = True
                Exit Function
            End If
        
        End If
    
    Next i
    
End Function

Private Function fnMonthsEmployed(sVinV As String, _
                                  sErrMsg As String, _
                                  lEmpNo As Long, _
                                  Optional bReturnDays = False) As Double
    
    Const SUB_NAME As String = "fnMonthsEmployed"
    
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim sDateHired As String
    Dim sDateTerminated As String
    Dim sDateStart As String
    Dim sDateEnd As String
    Dim dDiff As Double

    fnMonthsEmployed = 0#
            
    strSQL = "SELECT prm_date_hired, prm_date_termed"
    strSQL = strSQL & " FROM pr_master"
    strSQL = strSQL & " WHERE prm_empno = " & lEmpNo
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        sErrMsg = "Failed to access the database to get " & sVinV
        Exit Function
    End If
    
    If rsTemp.RecordCount = 0 Then
        sErrMsg = "Employee record not found for " & sVinV
        Exit Function
    End If
    
    sDateHired = fnGetField(rsTemp!prm_date_hired)
    sDateTerminated = fnGetField(rsTemp!prm_date_termed)
    
    strSQL = "SELECT prhs_date_hired, prhs_date_termed"
    strSQL = strSQL & " FROM pr_history"
    strSQL = strSQL & " WHERE prhs_empno = " & lEmpNo
    strSQL = strSQL & " AND prhs_date_hired <> " + tfnDateString(sDateHired, True)
    strSQL = strSQL & " AND prhs_date_termed IS NOT NULL"
    strSQL = strSQL & " ORDER BY prhs_date_hired"
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        sErrMsg = "Failed to access the database to get " & sVinV
        Exit Function
    End If
    
    If rsTemp.RecordCount = 0 Then
        
        'sErrMsg = "Employee is not rehired for " & sVinV
        'use date hired and/or date terminated for calculation
        If Not IsValidDate(sDateHired) Then
            sErrMsg = "Date Hired is not valid for " & sVinV
            Exit Function
        End If
            
        If Not IsValidDate(sDateTerminated) Then
            sDateTerminated = frmZZSEBPRC!txtStartDate
        End If
        
        If bReturnDays Then
            dDiff = fnDateDiff("d", CDate(sDateHired), CDate(sDateTerminated))
        Else
            dDiff = fnDateDiff("m", CDate(sDateHired), CDate(sDateTerminated))
        End If
        
        fnMonthsEmployed = dDiff
        Exit Function
    End If
    
    If rsTemp.RecordCount = 1 Then
        sDateStart = fnGetField(rsTemp!prhs_date_hired)
        sDateEnd = fnGetField(frmZZSEBPRC!prhs_date_termed)
        
        If IsValidDate(sDateStart) And IsValidDate(sDateEnd) Then
            
            If CDate(sDateStart) < CDate(sDateHired) Then
                
                If bReturnDays Then
                    dDiff = fnDateDiff("d", CDate(sDateStart), CDate(sDateEnd))
                Else
                    dDiff = fnDateDiff("m", CDate(sDateStart), CDate(sDateEnd))
                End If
            
            End If
        
        End If
    
    Else
        dDiff = 0
        
        Do
            sDateStart = fnGetField(rsTemp!prhs_date_hired)
            sDateEnd = fnGetField(frmZZSEBPRC!prhs_date_termed)
            
            If IsValidDate(sDateStart) And IsValidDate(sDateEnd) Then
                
                If CDate(sDateStart) < CDate(sDateHired) Then
                    
                    If bReturnDays Then
                        dDiff = dDiff + fnDateDiff("d", CDate(sDateStart), CDate(sDateEnd))
                    Else
                        dDiff = dDiff + fnDateDiff("m", CDate(sDateStart), CDate(sDateEnd))
                    End If
                
                End If
            
            End If
            
            rsTemp.MoveNext
        Loop Until rsTemp.EOF
        
        'from date hired until now (ending date)
        If bReturnDays Then
            dDiff = dDiff + fnDateDiff("d", CDate(sDateHired), CDate(frmZZSEBPRC!txtStartDate))
        Else
            dDiff = dDiff + fnDateDiff("m", CDate(sDateHired), CDate(frmZZSEBPRC!txtStartDate))
        End If
    
    End If
    
    fnMonthsEmployed = dDiff
End Function

Private Function fnShortageAmount(sVinV As String, _
                                  sErrMsg As String, _
                                  nPrftCtr As Integer) As Double
    
    Const SUB_NAME As String = "fnShortageAmount"
    
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim lSysParm3004  As Long
    Dim dDebitAmt As Double
    Dim dCreditAmt As Double
    Dim i As Long

    fnShortageAmount = 0#
            
    strSQL = "SELECT parm_field FROM sys_parm WHERE parm_nbr = 3004"
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        sErrMsg = "Failed to access the database to get " & tfnSQLString(sVinV)
        Exit Function
    End If
    
    If rsTemp.RecordCount = 0 Then
        sErrMsg = "SysParm#3004 not found for " & sVinV
        Exit Function
    End If
        
    lSysParm3004 = tfnRound(rsTemp!parm_field)
    
    strSQL = "SELECT SUM (gljrs_amount) AS var_value, gljrs_flag "
    strSQL = strSQL & " FROM gl_jrnl_rs, rs_shiftlink"
    strSQL = strSQL & " WHERE gljrs_shl = rssl_shl"
    strSQL = strSQL & " AND gljrs_account = " & lSysParm3004
    strSQL = strSQL & " AND rssl_prft_ctr = " & nPrftCtr
    strSQL = strSQL & " AND rssl_date BETWEEN " & tfnDateString(frmZZSEBPRC.txtStartDate, True)
    strSQL = strSQL & " AND " & tfnDateString(frmZZSEBPRC.txtEndDate, True)
    strSQL = strSQL & " GROUP BY gljrs_flag"
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        sErrMsg = "Failed to access the database to get " & tfnSQLString(sVinV)
        Exit Function
    End If
    
    If rsTemp.RecordCount = 0 Then
        sErrMsg = "No record found for " & sVinV
        bNoRecordFound = True
        Exit Function
    End If
    
    rsTemp.MoveFirst
    
    For i = 1 To rsTemp.RecordCount
        
        If fnGetField(rsTemp!gljrs_flag) = "D" Then
            dDebitAmt = tfnRound(rsTemp!var_value, DEFAULT_DECIMALS)
        Else
            dCreditAmt = tfnRound(rsTemp!var_value, DEFAULT_DECIMALS)
        End If
    
    Next i
    
    fnShortageAmount = dDebitAmt - dCreditAmt

End Function

'return error message if any
Private Function fnCheckFormula(ByVal sFormula As String, ByVal sBonusType As String) As String
    Dim i As Integer
    Dim sErrMsg As String
    Dim aryVariables As Variant
    Dim aryValues As Variant
    Dim objEvaluate As clsEquation
    
    On Error GoTo ErrTrap
    
    sFormula = LCase(Trim(sFormula))
    
    'check formula using bonus type
    sErrMsg = fnCheckVarAllowed(sFormula, sBonusType)
    
    If sErrMsg <> "" Then
        fnCheckFormula = sErrMsg
        Exit Function
    End If
    
    'start formula evaluation
    aryVariables = Array("pct", "dol", "amt1", "amt2", "mxt", "v1", "v2", "v3")
    aryValues = Array(1.23, 4.56, 7.89, 2.34, 3.45, 5.67, 6.78, 8.91)

    
    Set objEvaluate = New clsEquation
    
    For i = 0 To UBound(aryVariables)
        objEvaluate.Var(CStr(aryVariables(i))) = aryValues(i)
    Next i
    
    objEvaluate.Equation = sFormula
    
    fnCheckFormula = objEvaluate.Solve()
    
    Set objEvaluate = Nothing
    
    Exit Function
    
ErrTrap:
    tfnErrHandler "fnCheckFormula"
    fnCheckFormula = "Failed to validate Formula"

End Function

'return error message if any
Private Function fnCheckCondition(ByVal sCond As String, ByVal sBonusType As String) As String
    Dim i As Integer
    Dim sErrMsg As String
    Dim aryVariables As Variant
    Dim aryValues As Variant
    Dim objCondition As clsCondition
    
    On Error GoTo ErrTrap
    
    sCond = LCase(Trim(sCond))
    
    'check condition using bonus type
    sErrMsg = fnCheckVarAllowed(sCond, sBonusType)
    
    If sErrMsg <> "" Then
        fnCheckCondition = sErrMsg
        Exit Function
    End If
    
    'start formula evaluation
    aryVariables = Array("pct", "dol", "amt1", "amt2", "mxt", "v1", "v2", "v3")
    aryValues = Array(1.23, 4.56, 7.89, 2.34, 3.45, 5.67, 6.78, 8.91)

    
    Set objCondition = New clsCondition
    
    For i = 0 To UBound(aryVariables)
        objCondition.Var(CStr(aryVariables(i))) = aryValues(i)
    Next i
    
    objCondition.Equation = sCond
    
    fnCheckCondition = objCondition.Solve()
    
    Set objCondition = Nothing
    
    Exit Function
    
ErrTrap:
    tfnErrHandler "fnCheckCondition"
    fnCheckCondition = "Failed to validate Condition"

End Function

Private Function fnCheckVarAllowed(sFormula As String, sBonusType As String) As String
    Dim sInvalidVar As String
    Dim aryInvalidVar() As String
    Dim i As Integer
    
    'check formula using bonus type
    'vaid bonus type format: T[123][ECX]
    If Len(sBonusType) = 3 Then
        
        Select Case tfnRound(Mid(sBonusType, 2, 1))
            Case 1
                sInvalidVar = sInvalidVar + "v2,v3"
            Case 2
                sInvalidVar = sInvalidVar + "v3"
        End Select
    
        If UCase(Right(sBonusType, 1)) <> "E" Then
            sInvalidVar = sInvalidVar + ",mxt"
        End If
        
    End If
    
    aryInvalidVar = Split(sInvalidVar, ",")
    
    For i = 0 To UBound(aryInvalidVar)
        
        If aryInvalidVar(i) <> "" Then
            
            If InStr(sFormula, aryInvalidVar(i)) > 0 Then
                fnCheckVarAllowed = tfnSQLString(aryInvalidVar(i)) + _
                    " is not valid for Commission Type " + tfnSQLString(sBonusType)
                Exit Function
            End If
        
        End If
        
    Next i

    fnCheckVarAllowed = ""
End Function

Private Sub subGetBFormula(sBCode As String, _
                          nBLevel As Integer, _
                          sCondition As String, _
                          sFormula As String, _
                          sAdjCond As String, _
                          sAdjFormula As String)
    
    Const SUB_NAME As String = "subGetBFormula"
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    strSQL = "SELECT bf_condition, bf_formula, bf_adj_condition, bf_adj_formula"
    strSQL = strSQL + " FROM bonus_formula"
    strSQL = strSQL + " WHERE bf_bonus_code = " & tfnSQLString(sBCode)
    strSQL = strSQL & " AND bf_level = " & tfnRound(nBLevel)
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) > 0 Then
        sCondition = LCase(fnGetField(rsTemp!bf_condition))
        sFormula = LCase(fnGetField(rsTemp!bf_formula))
        sAdjCond = LCase(fnGetField(rsTemp!bf_adj_condition))
        sAdjFormula = LCase(fnGetField(rsTemp!bf_adj_formula))
    End If
    
End Sub

Public Function fnDeleteSalesRecord() As Boolean
    Const SUB_NAME As String = "fnDeleteSalesRecord"
    
    Dim strSQL As String
    
    strSQL = "DELETE FROM bonus_sales"
    strSQL = strSQL & " WHERE bs_sales_type = " & tfnSQLString(sSalesTypeCode)
    strSQL = strSQL & " AND " & tfnDateString(frmZZSEBPRC.txtFromDate, True)
    strSQL = strSQL & " BETWEEN bs_from_date AND bs_to_date"
    
    If Not fnExecuteSQL(strSQL, , SUB_NAME) Then
        Exit Function
    End If

    strSQL = "DELETE FROM bonus_sales"
    strSQL = strSQL & " WHERE bs_sales_type = " & tfnSQLString(sSalesTypeCode)
    strSQL = strSQL & " AND " & tfnDateString(frmZZSEBPRC.txtToDate, True)
    strSQL = strSQL & " BETWEEN bs_from_date AND bs_to_date"
    
    If Not fnExecuteSQL(strSQL, , SUB_NAME) Then
        Exit Function
    End If

    strSQL = "DELETE FROM bonus_sales"
    strSQL = strSQL & " WHERE bs_sales_type = " & tfnSQLString(sSalesTypeCode)
    strSQL = strSQL & " AND bs_from_date BETWEEN " & tfnDateString(frmZZSEBPRC.txtFromDate, True)
    strSQL = strSQL & " AND " & tfnDateString(frmZZSEBPRC.txtToDate, True)
    
    If Not fnExecuteSQL(strSQL, , SUB_NAME) Then
        Exit Function
    End If

    strSQL = "DELETE FROM bonus_sales"
    strSQL = strSQL & " WHERE bs_sales_type = " & tfnSQLString(sSalesTypeCode)
    strSQL = strSQL & " AND bs_to_date BETWEEN " & tfnDateString(frmZZSEBPRC.txtFromDate, True)
    strSQL = strSQL & " AND " & tfnDateString(frmZZSEBPRC.txtToDate, True)
    
    fnDeleteSalesRecord = fnExecuteSQL(strSQL, , SUB_NAME)
End Function

Public Function fnInsertUpdateSales() As Boolean
    Const SUB_NAME As String = "fnInsertUpdateSales"
    
    Dim i As Integer
    Dim strSQL As String
    Dim nPrftCtr As Integer
    Dim sOldPrftCtr As String
    Dim sFrmDt As String
    Dim sToDt As String
    Dim dSlsAmt As Double
    Dim sSType As String
    
    sSType = tfnSQLString(sSalesTypeCode)
    
    For i = 0 To tgmSales.RowCount - 1
        
        If tgmSales.ValidCell(colSPrftCtr, i) Then
            nPrftCtr = tfnRound(tgmSales.CellValue(colSPrftCtr, i))
'            sFrmDt = tfnDateString(tgmSales.CellValue(colSFromDate, i), True)
'            sToDt = tfnDateString(tgmSales.CellValue(colSToDate, i), True)
            sFrmDt = tfnDateString(frmZZSEBPRC!txtFromDate, True)
            sToDt = tfnDateString(frmZZSEBPRC!txtToDate, True)
            dSlsAmt = tfnRound(tgmSales.CellValue(colSAmount, i), 2)
            
            sOldPrftCtr = fnGetField(tgmSales.CellValue(ColxSOldPrftCtr, i))
            
            If sOldPrftCtr = "" Then
                strSQL = "INSERT INTO bonus_sales (bs_prft_ctr, bs_from_date, bs_to_date,"
                strSQL = strSQL & " bs_sales_amount, bs_sales_type) VALUES ("
                strSQL = strSQL & nPrftCtr & ","
                strSQL = strSQL & sFrmDt & ","
                strSQL = strSQL & sToDt & ","
                strSQL = strSQL & dSlsAmt & ","
                strSQL = strSQL & sSType & ")"
            
                If Not fnExecuteSQL(strSQL, , SUB_NAME) Then
                    Exit Function
                End If
            
            Else
                strSQL = "UPDATE bonus_sales SET"
                
                If nPrftCtr <> tfnRound(sOldPrftCtr) Then
                    strSQL = strSQL & " bs_prft_ctr = " & nPrftCtr * -1 & ","
                End If
                
                'strSQL = strSQL & " bs_from_date = " & sFrmDt & ","
                'strSQL = strSQL & " bs_to_date = " & sToDt & ","
                strSQL = strSQL & " bs_sales_amount = " & dSlsAmt
                strSQL = strSQL & " WHERE bs_sales_type = " & sSType
                strSQL = strSQL & " AND bs_prft_ctr = " & tfnRound(sOldPrftCtr)
                strSQL = strSQL & " AND bs_from_date = " & sFrmDt
                strSQL = strSQL & " AND bs_to_date = " & sToDt
            
                If Not fnExecuteSQL(strSQL, , SUB_NAME) Then
                    Exit Function
                End If
            
            End If
        
        End If
    
    Next i
    
    If t_nFormMode = EDIT_MODE Then
        'change bs_prft_ctr back to positive
        strSQL = "UPDATE bonus_sales SET"
        strSQL = strSQL & " bs_prft_ctr = bs_prft_ctr * -1"
        strSQL = strSQL & " WHERE bs_sales_type = " & sSType
        strSQL = strSQL & " AND bs_from_date = " & sFrmDt
        strSQL = strSQL & " AND bs_to_date = " & sToDt
        strSQL = strSQL & " AND bs_prft_ctr < 0"
    
        If Not fnExecuteSQL(strSQL, , SUB_NAME) Then
            Exit Function
        End If
    
    End If

    fnInsertUpdateSales = True
End Function

Public Function fnDeleteSales(sSType As String, sOldPrftCtr As String, sToDt As String, sFrmDt As String) As Boolean
    Const SUB_NAME As String = "fnDeleteSales"
    Dim strSQL As String
    
    If sOldPrftCtr = "" Then
        fnDeleteSales = True
        Exit Function
    End If
    
    strSQL = "DELETE FROM bonus_sales WHERE bs_sales_type = " & tfnSQLString(Trim(sSType))
    strSQL = strSQL & " AND bs_prft_ctr = " & tfnRound(sOldPrftCtr)
    strSQL = strSQL & " AND bs_from_date = " & tfnDateString(Trim(sFrmDt), True)
    strSQL = strSQL & " AND bs_to_date = " & tfnDateString(Trim(sToDt), True)
    
    fnDeleteSales = fnExecuteSQL(strSQL, , SUB_NAME)
End Function

Private Function fnValidGetBAmountEmp(lEmpNo As Long, _
                                 nPrftCtr As Integer, _
                                 sBGrade As String) As String

    Const SUB_NAME As String = "fnValidGetBAmountEmp"
    
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim i As Long
    Dim nEmpLevel As Integer
    Dim sDateHired As String
    Dim sDateTermed As String
    
    fnValidGetBAmountEmp = False
    
    strSQL = "SELECT *"
    strSQL = strSQL + " FROM pr_master WHERE prm_empno = " & lEmpNo
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        fnValidGetBAmountEmp = "Failed to access Database"
        Exit Function
    End If
    
    If rsTemp.RecordCount = 0 Then
        fnValidGetBAmountEmp = "Employee Number does not exist"
        Exit Function
    End If
    
    'checking prft_ctr
    For i = 1 To 5
        
        If fnGetField(rsTemp.Fields("prm_prft_ctr" & i)) <> "" Then
            
            If tfnRound(rsTemp.Fields("prm_prft_ctr" & i)) = nPrftCtr Then
                Exit For
            End If
            
        End If
        
    Next i
    
    If i > 5 Then
        fnValidGetBAmountEmp = "Profit Center " & nPrftCtr & " is not connected to the employee"
        bNoRecordFound = True
        Exit Function
    End If
    
    If fnGetField(rsTemp!prm_emp_level) = "" Then
        fnValidGetBAmountEmp = "Employee Level is NULL for the employee"
        Exit Function
    End If
    
    nEmpLevel = tfnRound(rsTemp!prm_emp_level)
    sDateHired = tfnFormatDate(rsTemp!prm_date_hired)
    sDateTermed = tfnFormatDate(rsTemp!prm_date_termed)
    
    'checking grade and employee level
    strSQL = "SELECT bg_emp_level"
    strSQL = strSQL + " FROM bonus_grades"
    strSQL = strSQL + " WHERE bg_grade = " & tfnSQLString(sBGrade)
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        fnValidGetBAmountEmp = "Failed to access Database"
        Exit Function
    End If
    
    If rsTemp.RecordCount = 0 Then
        fnValidGetBAmountEmp = "Commission Grade " + tfnSQLString(sBGrade) + " does not exist"
        Exit Function
    End If
    
    For i = 1 To rsTemp.RecordCount
        
        If tfnRound(rsTemp!bg_emp_level) = nEmpLevel Then
            Exit For
        End If
        
        rsTemp.MoveNext
    Next i
    
    If rsTemp.EOF Then
        fnValidGetBAmountEmp = "Employee Level " & nEmpLevel & " is not valid for Commission Grade " + tfnSQLString(sBGrade)
        Exit Function
    End If
    
    If sDateTermed <> "" Then
        
        If CDate(tfnDateString(sDateTermed)) < CDate(tfnDateString(frmZZSEBPRC!txtEndDate)) Then
            fnValidGetBAmountEmp = "Employee was terminated on " + sDateTermed
            Exit Function
        End If
        
    End If
    
    fnValidGetBAmountEmp = ""
End Function

Public Function IsValidDate(ByVal sDate As String) As Boolean
    If sDate = "" Then
        Exit Function
    End If
    
    sDate = tfnFormatDate(sDate)
    
    If SRegExpMatch(szDatePattern, sDate) <> 0 Then
        Exit Function
    End If
    
    If Not IsDate(tfnDateString(sDate)) Then
        Exit Function
    End If
    
    IsValidDate = True
End Function

Public Function fnBuildList(tgmEditor As clsTGSpreadSheet, _
                             nColIndex As Integer, _
                             sColType As Integer, _
                             Optional bCheckValid As Boolean = True, _
                             Optional bIncludeCurrentRow As Boolean = False, _
                             Optional bUnique As Boolean = False, _
                             Optional nWhereCol As Integer = -1, _
                             Optional sWhereItem As String = "") As String
    
    Const ColType_NUMERIC As Integer = 1
    Const ColType_STRING As Integer = 2
    
    Dim sTemp As String
    Dim i As Long
    Dim bAdd As Boolean
    Dim aryList() As Variant
    Dim nListCount As Integer
    Dim j As Integer
    
    If tgmEditor.RowCount < 1 Then
        Exit Function
    End If
    
    If tgmEditor.RowCount = 1 And Not bIncludeCurrentRow And tgmEditor.GetCurrentRowNumber = 0 Then
        Exit Function
    End If
    
    sTemp = ""
    
    nListCount = -1
    
    For i = 0 To tgmEditor.RowCount - 1
        
        If fnGetField(tgmEditor.CellValue(nColIndex, i)) <> "" Then
            
            If nWhereCol < 0 Then
                bAdd = True
            Else
                bAdd = fnGetField(tgmEditor.CellValue(nWhereCol, i)) = sWhereItem
            End If
            
            If bAdd Then
                
                If bUnique Then
                    
                    If sColType = ColType_NUMERIC Then
                        bAdd = Not fnFindInList(tfnRound(tgmEditor.CellValue(nColIndex, i)), aryList, nListCount)
                    Else
                        bAdd = Not fnFindInList(fnGetField(tgmEditor.CellValue(nColIndex, i)), aryList, nListCount)
                    End If
                
                Else
                    bAdd = True
                End If
            
            End If
            
            If bAdd Then
                bAdd = bIncludeCurrentRow Or ((Not bIncludeCurrentRow) And _
                    (i <> tgmEditor.GetCurrentRowNumber))
            End If
            
            If bAdd Then
                
                If bCheckValid Then
                    bAdd = tgmEditor.ValidCell(nColIndex, i)
                End If
            
            End If
            
            If bAdd Then
                
                If sColType = ColType_NUMERIC Then
                    sTemp = sTemp & tfnRound(tgmEditor.CellValue(nColIndex, i)) & ","
                Else
                    sTemp = sTemp & tfnSQLString(tgmEditor.CellValue(nColIndex, i)) & ","
                End If
                
                If bUnique Then
                    nListCount = nListCount + 1
                    ReDim Preserve aryList(nListCount)
                    
                    If sColType = ColType_NUMERIC Then
                        aryList(nListCount) = tfnRound(tgmEditor.CellValue(nColIndex, i))
                    Else
                        aryList(nListCount) = CStr(tgmEditor.CellValue(nColIndex, i))
                    End If
                
                End If
            
            End If
        
        End If
        
    Next i
    
    If sTemp <> "" Then
        fnBuildList = Left(sTemp, Len(sTemp) - 1)
    End If
    
End Function

Public Function fnLoadBonusDetails(lEmpNo As Long, _
                                   nPrftCtr As Integer, _
                                   sBonusCode As String) As Boolean
    
    Const SUB_NAME As String = "fnLoadBonusDetails"
    
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim sApproveEmpList As String
    Dim sApprovePrftCtrList As String
    
    strSQL = "SELECT bm_empno, bm_eligible_pc, bm_bonus_code, bm_eligible_date, bc_type,"
    strSQL = strSQL & " bc_frequency, bm_sequence, bf_level"
    strSQL = strSQL & " FROM bonus_master, bonus_codes, bonus_formula"
    strSQL = strSQL & " WHERE bm_bonus_code = bc_bonus_code"
    strSQL = strSQL & " AND bm_bonus_code = bf_bonus_code"
    
    If sBonusCode <> "" Then
        strSQL = strSQL & " AND bm_bonus_code = " & tfnSQLString(sBonusCode)
    End If
    
    If lEmpNo < 0 And nPrftCtr < 0 Then
        sApproveEmpList = fnBuildEmpPrftCtrList()
        
        If sApproveEmpList <> "" Then
            strSQL = strSQL & " AND (bm_empno || bm_eligible_pc) IN (" + sApproveEmpList + ")"
        End If
    
    Else
        
        If lEmpNo < 0 Then
            sApproveEmpList = fnBuildList(tgmApprove, colAEmpNo, 1, False, True, True, colAPrftCtr, fnGetField(nPrftCtr))
            
            If sApproveEmpList <> "" Then
                strSQL = strSQL & " AND bm_empno IN (" + sApproveEmpList + ")"
            End If
        
        Else
            strSQL = strSQL & " AND bm_empno = " & tfnRound(lEmpNo)
        End If
        
        If nPrftCtr < 0 Then
            sApprovePrftCtrList = fnBuildList(tgmApprove, colAPrftCtr, 1, False, True, True, colAEmpNo, fnGetField(lEmpNo))
            
            If sApprovePrftCtrList <> "" Then
                strSQL = strSQL & " AND bm_eligible_pc IN (" + sApprovePrftCtrList + ")"
            End If
        
        Else
            strSQL = strSQL & " AND bm_eligible_pc = " & nPrftCtr
        End If
    
    End If
    
    strSQL = strSQL & " AND bc_frequency = " & tfnSQLString(frmZZSEBPRC!txtFrequency)
    strSQL = strSQL & " AND " & tfnDateString(frmZZSEBPRC!txtEndDate, True)
    strSQL = strSQL & " BETWEEN bm_eligible_date AND bm_stop_date"
    strSQL = strSQL & " ORDER BY bm_empno, bm_eligible_pc, bm_sequence, bm_bonus_code, bf_level"
        
    tgmDetail.FillWithSQL t_dbMainDatabase, strSQL
    
    If tgmDetail.RowCount <= 0 Then
        MsgBox "No record found for the selection criteria", vbExclamation
        Exit Function
    End If
    
    fnLoadBonusDetails = True
End Function

Public Function fnBuildEmpPrftCtrList() As String
    
    Dim sTemp As String
    Dim i As Long
    
    If tgmApprove.RowCount < 1 Then
        Exit Function
    End If
    
    sTemp = ""
    
    For i = 0 To tgmApprove.RowCount - 1
        sTemp = sTemp + tfnSQLString(fnGetField(tgmApprove.CellValue(colAEmpNo, i)) + "," _
            + fnGetField(tgmApprove.CellValue(colAPrftCtr, i))) + ","
    Next i
    
    If sTemp <> "" Then
        fnBuildEmpPrftCtrList = Left(sTemp, Len(sTemp) - 1)
    End If
    
End Function

Public Function fnGetProposedEndDate(ByVal sStartDate As String, sFreq As String) As String
    Dim sEndDate As String
    
    sStartDate = tfnFormatDate(sStartDate)
    
    If Not IsValidDate(sStartDate) Then
        Exit Function
    End If
    
    Select Case sFreq
        Case sOneMth, sGas, sRatio
            sEndDate = DateAdd("d", -1, DateAdd("m", 1, CDate(sStartDate)))
        Case sBiWeek, sTwoWeek
            sEndDate = DateAdd("d", 13, CDate(sStartDate))
        Case "D"
            sEndDate = DateAdd("d", -1, DateAdd("d", 1, CDate(sStartDate)))
        Case "W"
            sEndDate = DateAdd("d", -1, DateAdd("ww", 1, CDate(sStartDate)))
        Case "Q"
            sEndDate = DateAdd("d", -1, DateAdd("q", 1, CDate(sStartDate)))
        Case "Y", "A"
            sEndDate = DateAdd("d", -1, DateAdd("yyyy", 1, CDate(sStartDate)))
        Case Else 'all others default to biweek
            sEndDate = DateAdd("d", 13, CDate(sStartDate))
    End Select
    
    fnGetProposedEndDate = tfnFormatDate(sEndDate)
End Function

Public Function fnHasApprove(Optional vApproveCount) As Boolean
    Dim i As Long
    
    For i = 0 To tgmApprove.RowCount - 1
        
        If tgmApprove.CellValue(colAApprove, i) = sColAppYes Then
            fnHasApprove = True
            
            If IsMissing(vApproveCount) Then
                Exit Function
            End If
            
            vApproveCount = vApproveCount + 1
        End If
        
    Next i
    
End Function

Public Function fnDateDiff(sInterval As String, _
                            sDate1 As String, _
                            sDate2 As String, _
                            Optional FirstDayOfYear As Integer = 0) As Double

    Dim sDateStart As String
    Dim sDateEnd As String
    Dim lYears As Long
    Dim lDaysInYears As Long
    Dim lMonths As Long
    Dim lDaysInMonths As Long
    Dim lDiff As Long
    
    On Error GoTo ErrTrap
    
    sInterval = LCase(sInterval)
    
    Select Case sInterval
        Case "d", "y", "w", "ww"
            sDateStart = sDate1
            sDateEnd = sDate2
            fnDateDiff = tfnRound(Abs(DateDiff(sInterval, CDate(sDate1), CDate(sDate2))))
        Case "m"
            lMonths = Abs(DateDiff("m", CDate(sDate1), CDate(sDate2)))
            
            If lMonths = 0 Then
                lMonths = 1
            End If
            
            sDateStart = sDate1
            sDateEnd = DateAdd("m", lMonths, sDateStart)
            lDaysInMonths = Abs(DateDiff("d", CDate(sDateStart), CDate(sDateEnd)))
            
            lDiff = Abs(DateDiff("d", CDate(sDate1), CDate(sDate2)))
            
            fnDateDiff = tfnRound(lDiff / lDaysInMonths, DEFAULT_DECIMALS) * lMonths
        Case "yyyy"
            lYears = Abs(DateDiff("yyyy", CDate(sDate1), CDate(sDate2)))
            
            'change by junsong 05/16/02, don't make so complicated and it doesn't work!
            'see employee # promotion date  pay period              years(correct)  years(original logic)
            '   7110009     01/02/1997      02/07/01-02/20/2001     3               3
            '   7110010     12/31/1995      02/07/01-02/20/2001     5               4
             
            If Left(sDate1, 5) <> "01/01" Then
                fnDateDiff = lYears - 1
            Else
                fnDateDiff = lYears
            End If
            
            If lYears < 0 Then
                lYears = 0
            End If
            
'            If lYears = 0 Then
'                lYears = 1
'            End If
'
'            If FirstDayOfYear = vbFirstJan1 Then
'                sDateStart = "01/01/" + Right(sDate1, 2)
'                sDateEnd = DateAdd("yyyy", lYears, sDateStart)
'            Else
'                sDateStart = sDate1
'                sDateEnd = DateAdd("yyyy", lYears, sDateStart)
'            End If
'
'            lDaysInYears = Abs(DateDiff("y", CDate(sDateStart), CDate(sDateEnd)))
'
'            lDiff = Abs(DateDiff("y", CDate(sDate1), CDate(sDate2)))
'
'            If FirstDayOfYear = vbFirstJan1 Then
'
'                If Left(sDate1, 5) <> "01/01" Then
'                    fnDateDiff = tfnRound(lDiff / lDaysInYears, DEFAULT_DECIMALS) * lYears - 1
'
'                    If fnDateDiff < 0 Then
'                        fnDateDiff = 0
'                    End If
'
'                Else
'                    fnDateDiff = tfnRound(lDiff / lDaysInYears, DEFAULT_DECIMALS) * lYears
'                End If
'
'            Else
'                fnDateDiff = tfnRound(lDiff / lDaysInYears, DEFAULT_DECIMALS) * lYears
'            End If
    
    End Select
    
    Exit Function
    
ErrTrap:
    tfnErrHandler "fnDateDiff"
End Function

Public Function fnSetRegularOtHoursPayCode() As Boolean
    Const SUB_NAME As String = "fnSetRegularOtHoursPayCode"
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    Dim i As Integer
    Dim sTemp As String
    Dim bError As Boolean
    
'david 05/10/2002 #368969
'commented out old code
'
'    Dim sSysParm30854 As String
'
'    strSQL = "SELECT parm_field FROM sys_parm WHERE parm_nbr = 30854"
'
'    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
'        subLogErrMsg "Failed to access the database."
'        subLogErrMsg "Processing terminates."
'        Exit Function
'    End If
'
'    If rsTemp.RecordCount > 0 Then
'        If IsNull(rsTemp!parm_field) Then
'            If bShowDetail Then
'                subLogErrMsg "SysParm#30854 is NULL"
'            End If
'        Else
'            sSysParm30854 = UCase(rsTemp!parm_field)
'
'            If bShowDetail Then
'                subLogErrMsg "SysParm#30854 = " + tfnSQLString(sSysParm30854)
'            End If
'
'            If Len(sSysParm30854) >= 4 Then
'                sPayCode_RegHrs = Left(sSysParm30854, 4)
'            End If
'
'            If Len(sSysParm30854) >= 9 Then
'                sPayCode_OtHrs = Trim(Mid(sSysParm30854, 6, 4))
'            End If
'
'            If bShowDetail Then
'                subLogErrMsg "Pay Code for Regular Hour = " + tfnSQLString(sPayCode_RegHrs)
'                subLogErrMsg "Pay Code for Overtime Hour = " + tfnSQLString(sPayCode_OtHrs)
'            End If
'
'        End If
'
'    Else
'        subLogErrMsg "SysParm#30854 not found"
'    End If
    
    sPayCode_OtHrs = ""
    
    strSQL = "SELECT zzsep_code FROM zzse_pay_code WHERE zzsep_factor > 1"

    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        subLogErrMsg "Failed to access the database."
        subLogErrMsg "Processing terminates."
        Exit Function
    End If

    If rsTemp.RecordCount > 0 Then
        For i = 1 To rsTemp.RecordCount
            sTemp = fnGetField(rsTemp!zzsep_code)
            
            If sTemp <> "" Then
                If sPayCode_OtHrs = "" Then
                    sPayCode_OtHrs = tfnSQLString(sTemp)
                Else
                    sPayCode_OtHrs = sPayCode_OtHrs + ", " + tfnSQLString(sTemp)
                End If
            End If
            
            rsTemp.MoveNext
        Next i
    
        If sPayCode_OtHrs = "" Then
            subLogErrMsg "Pay Code for Overtime Hour was not set up"
            bError = True
        Else
            If bShowDetail Then
                subLogErrMsg "Pay Code for Overtime Hour = " + sPayCode_OtHrs
            End If
        End If
    Else
        subLogErrMsg "Pay Code for Overtime Hours was not set up"
        bError = True
    End If
    
    sPayCode_RegHrs = ""
    
    strSQL = "SELECT zzsep_code FROM zzse_pay_code WHERE zzsep_factor = 1"

    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        subLogErrMsg "Failed to access the database."
        subLogErrMsg "Processing terminates."
        Exit Function
    End If

    If rsTemp.RecordCount > 0 Then
        For i = 1 To rsTemp.RecordCount
            sTemp = fnGetField(rsTemp!zzsep_code)
            
            If sTemp <> "" Then
                If sPayCode_RegHrs = "" Then
                    sPayCode_RegHrs = tfnSQLString(sTemp)
                Else
                    sPayCode_RegHrs = sPayCode_RegHrs + ", " + tfnSQLString(sTemp)
                End If
            End If
            
            rsTemp.MoveNext
        Next i
    
        If sPayCode_RegHrs = "" Then
            subLogErrMsg "Pay Code for Regular Hour was not set up"
            bError = True
        Else
            If bShowDetail Then
                    subLogErrMsg "Pay Code for Regular Hour = " + sPayCode_RegHrs
            End If
        End If
    Else
        subLogErrMsg "Pay Code for Regular Hours was not set up"
        bError = True
    End If
    
    If bError Then
        subLogErrMsg "Processing terminates."
        Exit Function
    End If
    
    fnSetRegularOtHoursPayCode = True
End Function
