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
Public tgmHours As clsTGSpreadSheet
Public tgmPrftCtr As clsTGSpreadSheet
Public tgmApprove As clsTGSpreadSheet
Public tgsApprove As clsTGSelector
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
'Public Const colSFromDate As Integer = 3
'Public Const colSToDate As Integer = 4

'Time Card Grid Column Names
Public Const colHClockIn As Integer = 0
Public Const colHPrftCtr As Integer = 1
Public Const colHPayCode As Integer = 2
Public Const colHPayType As Integer = 3
Public Const colHHrsDol As Integer = 4
Public ColHHdnSource As Integer

'Profit Center Grid Column Names
Public Const colPProfit As Integer = 0
Public Const colPTotal As Integer = 1

'Approve Grid Column Names
Public Const colAApprove As Integer = 0
Public Const colAppYes As Integer = 0
Public Const colAppNo As Integer = 1
Public Const colAPrftCtr As Integer = 1
Public Const colAEmpNo As Integer = 2
Public Const colAEmpName As Integer = 3
Public Const colAPayCode As Integer = 4
Public Const colAPayHours As Integer = 5
Public Const colADate As Integer = 6
Public Const colABonusAmt As Integer = 7
Public colAHdnBAmtLvls As Integer

'Detail Grid Column Names
Public Const colDBCode As Integer = 0
Public Const colDBCDesc As Integer = 1
Public Const colDBLevel As Integer = 2
Public Const colDBType As Integer = 3
Public Const colDBFreq As Integer = 4
Public Const colDElgDate As Integer = 5
Public Const colDBAmt As Integer = 6

Public vArrBonus() As Variant
Public clsMath As clsEquation
'

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

Public Sub subLogErrMsg(sErrorMessage As String)
    Dim nFileNumber As Integer
    Dim sLineContents As String
    Dim sTimeStamp As String
    Dim sArrMsg() As String
    Dim i As Integer
    
    'Put the time stamp if the sLogFilePath is empty
    On Error Resume Next
    sTimeStamp = "Error Log Created on : " & Date & " at " & Time & vbCrLf
    nFileNumber = FreeFile
    Open sLogFilePath For Input As #nFileNumber
    If Not EOF(nFileNumber) Then Line Input #nFileNumber, sLineContents: Close nFileNumber
    If sLineContents = "" Then tfnLog sTimeStamp, sLogFilePath
    
    'Writing the log to the file...
    tfnLog sErrorMessage, sLogFilePath
    
    sArrMsg = Split(sErrorMessage, vbCrLf)
    For i = 0 To UBound(sArrMsg)
        frmZZSEBPRC.lstProcess.AddItem sArrMsg(i)
    Next i
    DoEvents
    
End Sub

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

Public Function fnCreateReport(Index As Integer) As Boolean
    Dim sArrReport() As String
    Dim sHeadTitle As String
    Dim sReportTitle As String
    Dim sApprove As String
    Dim nArSize As Integer
    Dim sFormula As String
    Dim i As Integer, j As Integer
    
    Screen.MousePointer = vbHourglass
    frmZZSEBPRC.tfnSetStatusBarMessage "Printing report, please wait..."
    
    fnCreateReport = False

    Select Case Index
        Case TabApprove
            ReDim sArrReport(tgmApprove.RowCount - 1)
            For i = 0 To tgmApprove.RowCount - 1
                sApprove = "N"
                If tgmApprove.CellValue(colAApprove, i) = colAppYes Then
                    sApprove = "Y"
                End If
                sArrReport(i) = fnTranc(sApprove, 8, vbCenter) & Space(1) _
                    & fnTranc(tgmApprove.CellValue(colAEmpNo, i), 9, vbLeftJustify) & Space(1) _
                    & fnTranc(tgmApprove.CellValue(colAEmpName, i), 46, vbLeftJustify) & Space(1) _
                    & fnTranc(tgmApprove.CellValue(colAPayCode, i), 7, vbLeftJustify) & Space(2) _
                    & fnTranc(tgmApprove.CellValue(colAPayHours, i), 6, vbLeftJustify) & Space(2) _
                    & fnTranc(tgmApprove.CellValue(colADate, i), 10, vbLeftJustify) & Space(1) _
                    & fnTranc(tgmApprove.CellValue(colABonusAmt, i), 10, vbRightJustify)
            Next i
            sReportTitle = "Employee Bonus Approval Report"
            sHeadTitle = fnTranc("Approval", 8, vbLeftJustify) & Space(1)
            sHeadTitle = sHeadTitle & fnTranc("Employee", 9, vbLeftJustify) & Space(1)
            sHeadTitle = sHeadTitle & fnTranc("Employee Name", 46, vbLeftJustify) & Space(1)
            sHeadTitle = sHeadTitle & fnTranc("PayCode", 7, vbLeftJustify) & Space(2)
            sHeadTitle = sHeadTitle & fnTranc("PayHrs", 6, vbLeftJustify) & Space(2)
            sHeadTitle = sHeadTitle & fnTranc("Date", 10, vbLeftJustify) & Space(1)
            sHeadTitle = sHeadTitle & fnTranc("Amount", 10, vbRightJustify)
        Case TabDetails
            nArSize = (tgmDetail.RowCount * 2) - 1
            i = 0
            ReDim sArrReport(nArSize)
            For j = 0 To UBound(sArrReport)
                If j Mod 2 = 0 Then
                    sArrReport(j) = fnTranc(tgmDetail.CellValue(colDBCode, i), 5, vbLeftJustify) & Space(2) _
                     & fnTranc(tgmDetail.CellValue(colDBCDesc, i), 49, vbLeftJustify) & Space(2) _
                     & fnTranc(tgmDetail.CellValue(colDBLevel, i), 5, vbCenter) & Space(2) _
                     & fnTranc(tgmDetail.CellValue(colDBType, i), 4, vbCenter) & Space(2) _
                     & fnTranc(tgmDetail.CellValue(colDBFreq, i), 9, vbCenter) & Space(2) _
                     & fnTranc(tgmDetail.CellValue(colDElgDate, i), 10, vbCenter) & Space(2) _
                     & fnTranc(tgmDetail.CellValue(colDBAmt, i), 10, vbRightJustify)
                     sFormula = fnGetBFormula(tgmDetail.CellValue(colDBCode, i), tgmDetail.CellValue(colDBLevel, i))
                    i = i + 1
                Else
                    sArrReport(j) = Space(7) & "Formula: (" & sFormula & ")" & vbCrLf
                End If
            Next j
            sReportTitle = "Employee Bonus Details"
            If frmZZSEBPRC.txtEmployee <> "" Then
                sReportTitle = frmZZSEBPRC.txtEmployee & "-" & frmZZSEBPRC.txtEmpName & " Bonus Details"
            End If
            sHeadTitle = fnTranc("Bonus", 5, vbLeftJustify) & Space(2)
            sHeadTitle = sHeadTitle & fnTranc("Bonus Code Description", 49, vbCenter) & Space(2)
            sHeadTitle = sHeadTitle & fnTranc("Level", 5, vbLeftJustify) & Space(2)
            sHeadTitle = sHeadTitle & fnTranc("Type", 4, vbLeftJustify) & Space(2)
            sHeadTitle = sHeadTitle & fnTranc("Frequency", 9, vbLeftJustify) & Space(2)
            sHeadTitle = sHeadTitle & fnTranc("Eligible", 10, vbCenter) & Space(2)
            sHeadTitle = sHeadTitle & fnTranc("Bonus", 10, vbRightJustify)
            sHeadTitle = sHeadTitle & vbCrLf
            sHeadTitle = sHeadTitle & fnTranc("Code", 5, vbLeftJustify) & Space(77)
            sHeadTitle = sHeadTitle & fnTranc("Date", 10, vbCenter) & Space(2)
            sHeadTitle = sHeadTitle & fnTranc("Amount", 10, vbRightJustify)
    End Select
    
    sHeadTitle = sHeadTitle & vbCrLf & String(104, "-")
    
    If Not fnSetupPrinter(vbPRORPortrait) Then
        Exit Function
    End If
    
    subSetTitle sHeadTitle
    
    If Not fnSendToPrinter(sArrReport(), sReportTitle) Then
        Exit Function
    End If
    
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
Public Function fnGetBonusAmount(lEmployeeNo As Long, sBType As String, _
                                    sBCode As String, nLevel As Integer) As Double
    Const SUB_NAME As String = "fnGetBonusAmount"
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim i As Integer
    
    sBCode = Trim(sBCode)
    sBType = Trim(sBType)
    fnGetBonusAmount = 0#
    
    If sBCode = "" Or sBType = "" Then
        Exit Function
    End If
    
    strSQL = "SELECT * FROM bonus_formula WHERE bf_bonus_code = " & tfnSQLString(sBCode)
    strSQL = strSQL & " AND bf_level = " & tfnRound(nLevel)
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) <> 1 Then
        subLogErrMsg Space(7) & "No record found for the bonus formula"
        Exit Function
    End If
    
    fnGetBonusAmount = fnCalculateBonus(rsTemp!bf_percent, rsTemp!bf_dollar, _
                        rsTemp!bf_amount1, rsTemp!bf_amount2, rsTemp!bf_variable1, _
                        rsTemp!bf_variable2, rsTemp!bf_variable3, rsTemp!bf_max_total, _
                        rsTemp!bf_formula, fnGetField(rsTemp!bf_condition), _
                        fnGetField(rsTemp!bf_adj_formula), fnGetField(rsTemp!bf_adj_condition), sBType)
                
End Function

Public Function fnInsertHoldBonus(lEmpNo As Long, sPayCode As String, dChkAmt As Double, _
                                  lHours As Long, sDate As String) As Boolean
    Const SUB_NAME As String = "fnInsertHoldBonus"
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    fnInsertHoldBonus = False
    
    strSQL = "INSERT INTO bonus_hold(bh_empno, bh_pay_code, bh_check_amount, bh_hours, bh_date)"
    strSQL = strSQL & " VALUES(" & tfnRound(lEmpNo) & ","
    strSQL = strSQL & tfnSQLString(Trim(sPayCode)) & ", "
    strSQL = strSQL & tfnRound(dChkAmt, DEFAULT_DECIMALS) & ", "
    strSQL = strSQL & tfnRound(lHours) & ", "
    strSQL = strSQL & tfnDateString(Trim(sDate), True) & ")"
    
    If Not fnExecuteSQL(strSQL, , SUB_NAME) Then
        MsgBox "Failed to insert the record", vbExclamation
        Exit Function
    End If
    
    fnInsertHoldBonus = True
    
End Function

Private Function fnExecuteSQL(szSQL As String, Optional nDB As Variant, _
                Optional szCalledFrom As Variant, Optional bShowError As Variant) As Boolean
                
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

'This function will calculate the amount for 1 Employee, 1 BCode and 1 Level at a time
Private Function fnCalculateBonus(PCT As Double, DOL As Double, AMT1 As Double, AMT2 As Double, _
                                  sV1 As String, sV2 As String, sV3 As String, MXT As Double, _
                                  sFmla As String, sCond As String, sAFmla As String, _
                                  sACond As String, sBType As String) As Double
    Const SUB_NAME As String = "fnCalculateBonus"
    Dim strSQL As String, rsTemp As Recordset
    Dim sErrMsg As String, i As Integer
    Dim V1 As Double, V2 As Double, V3 As Double
    Dim p1 As Integer, sActualVal As String, arVal() As String
    Dim dBonusAmt As Double
    
    fnCalculateBonus = 0#
    sFmla = Trim(sFmla): sCond = Trim(sCond): sACond = Trim(sACond): sAFmla = Trim(sAFmla)
    sV1 = Trim(sV1): sV2 = Trim(sV2): sV3 = Trim(sV3)
    PCT = tfnRound(PCT, DEFAULT_DECIMALS): DOL = tfnRound(DOL, DEFAULT_DECIMALS)
    
    'Get real values...
    V1 = fnGetVarValue(sV1, sErrMsg)
    V2 = fnGetVarValue(sV2, sErrMsg)
    V3 = fnGetVarValue(sV3, sErrMsg)
    
    If sErrMsg <> "" Then
        subLogErrMsg sErrMsg
        Exit Function
    End If
    
    If sFmla = "" Then Exit Function
    arVal = Split(sFmla, " ")
    For i = 0 To UBound(arVal)
        Select Case arVal(i)
            Case "v1"
                sActualVal = sActualVal & CStr(V1)
            Case "v2"
                sActualVal = sActualVal & CStr(V2)
            Case "v3"
                sActualVal = sActualVal & CStr(V3)
            Case "amt1"
                sActualVal = sActualVal & CStr(AMT1)
            Case "amt2"
                sActualVal = sActualVal & CStr(AMT2)
            Case "pct"
                sActualVal = sActualVal & CStr(PCT / 100)
            Case "dol"
                sActualVal = sActualVal & CStr(DOL)
            Case "mxt"
                sActualVal = sActualVal & CStr(MXT)
            Case Else
                sActualVal = sActualVal & arVal(i)
        End Select
    Next i
    
    'Get the bonus amount based on the formula...
    dBonusAmt = tfnRound(clsMath.Calculate(sActualVal, sErrMsg), DEFAULT_DECIMALS)
    If sErrMsg <> "" Then
        subLogErrMsg sErrMsg & ", Invalid Formula (" & sFmla & ")"
        Exit Function
    End If
    
    'Apply the condition Now...
    If sCond <> "" Then
        
    End If
    
    fnCalculateBonus = dBonusAmt

End Function

Private Function fnGetVarValue(sVariable As String, sErrMsg As String) As Double
    Const SUB_NAME As String = "fnGetVarValue"
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim szAltSQL As String
    Dim dDebitAmt As Double, dCreditAmt As Double
    Dim nVarUsed As Integer
    Dim i As Integer
    
    fnGetVarValue = 0: sErrMsg = ""
    
    If frmZZSEBPRC.txtPrftCtr <> "" Then
        szAltSQL = " AND rssl_prft_ctr = " & frmZZSEBPRC.txtPrftCtr
        If frmZZSEBPRC.txtDate <> "" Then
            szAltSQL = szAltSQL & " AND rssl_date = " & tfnDateString(frmZZSEBPRC.txtDate, True)
        Else
            szAltSQL = szAltSQL & " AND rssl_date = " & tfnDateString(Date, True)
        End If
    End If
    
    Select Case LCase(sVariable)
        Case "inside_sales"
            strSQL = "SELECT SUM (rsc_retail) as var_value FROM rs_scat, rs_shiftlink, rs_cat"
            strSQL = strSQL & " WHERE rsc_shl = rssl_shl"
            strSQL = strSQL & " AND rsc_catagory = rsct_catagory"
            strSQL = strSQL & " AND rsct_catagory IN('M','N','D')" & szAltSQL
        Case "gallons_gas"
            strSQL = "SELECT SUM (rsd_gal) as var_value FROM rs_daily, rs_shiftlink"
            strSQL = strSQL & " WHERE rsd_shl = rssl_shl" & szAltSQL
        Case "day_off_slip_days"
            strSQL = "SELECT COUNT (bd_shl) as var_value FROM bonus_day_off_slip, rs_shiftlink"
            strSQL = strSQL & " WHERE bd_shl = rssl_shl" & szAltSQL
        Case "total_pay"
            strSQL = "SELECT SUM (prci_total) as var_value FROM pr_check_item, pr_check, pr_pay"
            strSQL = strSQL & " WHERE prc_lnk = prci_lnk"
            strSQL = strSQL & " AND prc_check_paid <> 'V'"
            strSQL = strSQL & " AND prci_pay_code = prpa_pay_code"
            strSQL = strSQL & " AND prpa_pay_code = 'P'"
            'strSQL = strSQL & " AND prc_chk_date IN()"
        Case "months_in_grade"
        Case "years_as_manager"
        Case "months_employed"
            strSQL = "SELECT prm_date_hired, prhs_date_termed FROM pr_master, pr_history"
            strSQL = strSQL & " WHERE prm_empno = prhs_empno"
            If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
                sErrMsg = "Failed to access the database to get " & sVariable
                Exit Function
            End If
            If rsTemp.RecordCount = 0 Then
                sErrMsg = "No record found for " & sVariable
                Exit Function
            End If
            If fnGetField(rsTemp!prhs_date_termed) <> "" Then
                fnGetVarValue = DateDiff("m", rsTemp!prm_date_hired, rsTemp!prhs_date_termed)
            End If
            Exit Function
        Case "shortage_amount"
            strSQL = "SELECT SUM (gljrs_amount) as var_value, gljrs_flag "
            strSQL = strSQL & " FROM gl_jrnl_rs, rs_shiftlink"
            strSQL = strSQL & " WHERE gljrs_shl = rssl_shl" & szAltSQL
            strSQL = strSQL & " AND gljrs_account IN (SELECT parm_field FROM sys_parm WHERE parm_nbr = 3004)"
            strSQL = strSQL & " GROUP BY gljrs_flag"
            If GetRecordSet(rsTemp, strSQL, , SUB_NAME) > 0 Then
                rsTemp.MoveFirst
                For i = 1 To rsTemp.RecordCount
                    If fnGetField(rsTemp!gljrs_flag) = "D" Then
                        dDebitAmt = tfnRound(rsTemp!var_value, DEFAULT_DECIMALS)
                    Else
                        dCreditAmt = tfnRound(rsTemp!var_value, DEFAULT_DECIMALS)
                    End If
                Next i
                fnGetVarValue = dDebitAmt - dCreditAmt
                Exit Function
            End If
        Case "check_amount"
        Case "pay_hours"
            strSQL = "SELECT SUM (prci_input_amt) as var_value "
            strSQL = strSQL & " FROM pr_check_item, pr_check, pr_pay"
            strSQL = strSQL & " WHERE prc_lnk = prci_lnk"
            strSQL = strSQL & " AND prc_check_paid <> 'V' "
            strSQL = strSQL & " AND prci_pay_code = prpa_pay_code"
            strSQL = strSQL & " AND prpa_pay_code = 'P'"
            strSQL = strSQL & " AND prpa_calc_method = 'H'"
            'strSQL = strSQL & " AND prc_chk_date IN()"
        Case "min_pay"
            strSQL = ""
            strSQL = strSQL & ""
        Case Else
            Exit Function
    End Select
    
    If strSQL = "" Then Exit Function
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        sErrMsg = "Failed to access the database to get " & sVariable
        Exit Function
    End If
    
    If rsTemp.RecordCount = 0 Then
        sErrMsg = "No record found for " & sVariable
        Exit Function
    End If
    
    If rsTemp.RecordCount = 1 Then
        fnGetVarValue = tfnRound(rsTemp!var_value, DEFAULT_DECIMALS)
    End If
    
End Function

Private Function fnGetBFormula(sBCode As String, nBLevel As Integer) As String
    Const SUB_NAME As String = "fnGetBFormula"
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    fnGetBFormula = ""
    
    strSQL = "SELECT bf_formula FROM bonus_formula WHERE bf_bonus_code = " & tfnSQLString(sBCode)
    strSQL = strSQL & " AND bf_level = " & tfnRound(nBLevel)
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) = 1 Then
        fnGetBFormula = fnGetField(rsTemp!bf_formula)
    End If

End Function

Public Function fnDeleteSalesRecord() As Boolean
    Const SUB_NAME As String = "fnDeleteSalesRecord"
    
    Dim strSQL As String
    
    fnDeleteSalesRecord = False
    
    strSQL = "DELETE FROM bonus_sales"
    strSQL = strSQL & " WHERE bs_sales_type = " & tfnSQLString(frmZZSEBPRC.fnGetSalesType())
    strSQL = strSQL & " AND " & tfnDateString(frmZZSEBPRC.txtFromDate, True)
    strSQL = strSQL & " BETWEEN bs_from_date AND bs_to_date"
    
    If Not fnExecuteSQL(strSQL, , SUB_NAME) Then
        Exit Function
    End If

    strSQL = "DELETE FROM bonus_sales"
    strSQL = strSQL & " WHERE bs_sales_type = " & tfnSQLString(frmZZSEBPRC.fnGetSalesType())
    strSQL = strSQL & " AND " & tfnDateString(frmZZSEBPRC.txtToDate, True)
    strSQL = strSQL & " BETWEEN bs_from_date AND bs_to_date"
    
    If Not fnExecuteSQL(strSQL, , SUB_NAME) Then
        Exit Function
    End If

    strSQL = "DELETE FROM bonus_sales"
    strSQL = strSQL & " WHERE bs_sales_type = " & tfnSQLString(frmZZSEBPRC.fnGetSalesType())
    strSQL = strSQL & " AND bs_from_date BETWEEN " & tfnDateString(frmZZSEBPRC.txtFromDate, True)
    strSQL = strSQL & " AND " & tfnDateString(frmZZSEBPRC.txtToDate, True)
    
    If Not fnExecuteSQL(strSQL, , SUB_NAME) Then
        Exit Function
    End If

    strSQL = "DELETE FROM bonus_sales"
    strSQL = strSQL & " WHERE bs_sales_type = " & tfnSQLString(frmZZSEBPRC.fnGetSalesType())
    strSQL = strSQL & " AND bs_to_date BETWEEN " & tfnDateString(frmZZSEBPRC.txtFromDate, True)
    strSQL = strSQL & " AND " & tfnDateString(frmZZSEBPRC.txtToDate, True)
    
    If Not fnExecuteSQL(strSQL, , SUB_NAME) Then
        Exit Function
    End If

    fnDeleteSalesRecord = True
End Function

Public Function fnInsertUpdateSales() As Boolean
    Const SUB_NAME As String = "fnInsertUpdateSales"
    Dim i As Integer
    Dim strSQL As String
    
    Dim nPrftCtr As Integer
    Dim sFrmDt As String
    Dim sToDt As String
    Dim dSlsAmt As Double
    Dim sSType As String
    
    fnInsertUpdateSales = False
    
    sSType = tfnSQLString(frmZZSEBPRC.fnGetSalesType())
    
    For i = 0 To tgmSales.RowCount - 1
        If tgmSales.ValidCell(colSPrftCtr, i) And fnGetField(tgmSales.CellValue(colSPrftCtr, i)) <> "" Then
            nPrftCtr = tfnRound(tgmSales.CellValue(colSPrftCtr, i))
    '        sFrmDt = tfnDateString(tgmSales.CellValue(colSFromDate, i), True)
    '        sToDt = tfnDateString(tgmSales.CellValue(colSToDate, i), True)
            sFrmDt = tfnDateString(frmZZSEBPRC!txtFromDate, True)
            sToDt = tfnDateString(frmZZSEBPRC!txtToDate, True)
            dSlsAmt = tfnRound(tgmSales.CellValue(colSAmount, i), 2)
            
            If t_nFormMode = ADD_MODE Then
                strSQL = "INSERT INTO bonus_sales (bs_prft_ctr, bs_from_date, bs_to_date,"
                strSQL = strSQL & " bs_sales_amount, bs_sales_type) VALUES ("
                strSQL = strSQL & nPrftCtr & ","
                strSQL = strSQL & sFrmDt & ","
                strSQL = strSQL & sToDt & ","
                strSQL = strSQL & dSlsAmt & ","
                strSQL = strSQL & sSType & ")"
            Else
                strSQL = "UPDATE bonus_sales SET"
                'strSQL = strSQL & " bs_prft_ctr = " & nPrftCtr & ","
                'strSQL = strSQL & " bs_from_date = " & sFrmDt & ","
                'strSQL = strSQL & " bs_to_date = " & sToDt & ","
                strSQL = strSQL & " bs_sales_amount = " & dSlsAmt
                strSQL = strSQL & " WHERE bs_sales_type = " & sSType
                strSQL = strSQL & " AND bs_prft_ctr = " & nPrftCtr
                strSQL = strSQL & " AND bs_from_date = " & sFrmDt
                strSQL = strSQL & " AND bs_to_date = " & sToDt
            End If
        
            If Not fnExecuteSQL(strSQL, , SUB_NAME) Then
                Exit Function
            End If
        End If
    Next i
    
    fnInsertUpdateSales = True

End Function

Public Function fnDeleteSales(sSType As String, nPrftCtr As Integer, sToDt As String, sFrmDt As String) As Boolean
    Const SUB_NAME As String = "fnDeleteSales"
    Dim strSQL As String
    
    fnDeleteSales = False
    
    strSQL = "DELETE FROM bonus_sales WHERE bs_sales_type = " & tfnSQLString(Trim(sSType))
    strSQL = strSQL & " AND bs_prft_ctr = " & nPrftCtr
    strSQL = strSQL & " AND bs_from_date = " & tfnDateString(Trim(sFrmDt), True)
    strSQL = strSQL & " AND bs_to_date = " & tfnDateString(Trim(sToDt), True)
    
    If fnExecuteSQL(strSQL, , SUB_NAME) Then
        fnDeleteSales = True
    End If

End Function
