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
Public Const colHPayType As Integer = 3
Public Const colHHrsDol As Integer = 4
Public ColHHdnSource As Integer

'Profit Center Grid Column Names
Public Const colPProfit As Integer = 0
Public Const colPTotal As Integer = 1

'Approve value
Public Const colAppYes As Integer = 0
Public Const colAppNo As Integer = 1

'Approve Grid Column Names
Public Const colAApprove As Integer = 0
Public Const colAEmpNo As Integer = 1
Public Const colAEmpName As Integer = 2
Public Const colADate As Integer = 3
Public Const colAPrftCtr As Integer = 4
Public Const colAPayCode As Integer = 5
Public Const colAPayHours As Integer = 6
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
Public objMath As clsEquation
Public objCond As clsCondition
'

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
    If Not EOF(nFileNumber) Then
        Line Input #nFileNumber, sLineContents
        Close nFileNumber
    End If
    
    If sLineContents = "" Then
        tfnLog sTimeStamp, sLogFilePath
    End If
    
    'Writing the log to the file...
    tfnLog sErrorMessage, sLogFilePath
    
    sArrMsg = Split(sErrorMessage, vbCrLf)
    For i = 0 To UBound(sArrMsg)
        frmZZSEBPRC.lstProcess.AddItem sArrMsg(i)
    Next i
    
    DoEvents
    
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
    
    sErrMsg = fnValidEmployee(lEmployeeNo, nPrftCtr, sBGrade)
    
    If sErrMsg <> "" Then
        subLogErrMsg Space(7) & sErrMsg
        Exit Function
    End If
    
    strSQL = "SELECT * FROM bonus_formula WHERE bf_bonus_code = " & tfnSQLString(sBCode)
    strSQL = strSQL & " AND bf_level = " & tfnRound(nLevel)
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) <> 1 Then
        subLogErrMsg Space(7) & "No record found for the bonus formula"
        Exit Function
    End If
    
    fnGetBonusAmount = fnCalculateBonus(nPrftCtr, _
        tfnRound(rsTemp!bf_percent, DEFAULT_DECIMALS), _
        tfnRound(rsTemp!bf_dollar, 2), _
        tfnRound(rsTemp!bf_amount1, 2), _
        tfnRound(rsTemp!bf_amount2, 2), _
        fnGetField(rsTemp!bf_variable1), _
        fnGetField(rsTemp!bf_variable2), _
        fnGetField(rsTemp!bf_variable3), _
        tfnRound(rsTemp!bf_max_total), _
        fnGetField(rsTemp!bf_formula), _
        fnGetField(rsTemp!bf_condition), _
        fnGetField(rsTemp!bf_adj_formula), _
        fnGetField(rsTemp!bf_adj_condition), _
        sBType)
                
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

'This function will calculate the amount for 1 Employee, 1 BCode and 1 Level at a time
Private Function fnCalculateBonus(nPrftCtr As Integer, _
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
    
    'Get real values...
    V1 = fnGetVarValue(nPrftCtr, sV1, sErrMsg)
    V2 = fnGetVarValue(sV2, sErrMsg)
    V3 = fnGetVarValue(sV3, sErrMsg)
    
    If sErrMsg <> "" Then
        subLogErrMsg sErrMsg
        Exit Function
    End If
    
    'set the variables value for condition
    objCond.Var("pct") = PCT
    objCond.Var("dol") = DOL
    objCond.Var("amt1") = AMT1
    objCond.Var("amt2") = AMT2
    objCond.Var("mxt") = MXT
    objCond.Var("v1") = V1
    objCond.Var("v2") = V2
    objCond.Var("v3") = V3
    
    'set the variables value for formula
    objMath.Var("pct") = PCT
    objMath.Var("dol") = DOL
    objMath.Var("amt1") = AMT1
    objMath.Var("amt2") = AMT2
    objMath.Var("mxt") = MXT
    objMath.Var("v1") = V1
    objMath.Var("v2") = V2
    objMath.Var("v3") = V3
    
    If sCond <> "" Then
        bConditionOK = objCond.CheckCondition(sCond, sErrMsg)
        If sErrMsg <> "" Then
            subLogErrMsg sErrMsg & ", Invalid Condition Clause (" & sCond & ")"
            Exit Function
        End If
    End If
    
    If bConditionOK Then
        dBonusAmt = tfnRound(objMath.Calculate(sFmla, sErrMsg), DEFAULT_DECIMALS)
        If sErrMsg <> "" Then
            subLogErrMsg sErrMsg & ", Invalid Formula (" & sFmla & ")"
            Exit Function
        End If
    Else
        dBonusAmt = 0#
    End If
    
    'reset the v1, v2, or v3 if they are "check_amount"
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
        If sCond = "" Then
            bConditionOK = True
        Else
            bConditionOK = objCond.CheckCondition(sACond, sErrMsg)
            If sErrMsg <> "" Then
                subLogErrMsg sErrMsg & ", Invalid Condition Clause (" & sACond & ")"
                Exit Function
            End If
        End If
    
        If bConditionOK Then
            dBonusAmt = tfnRound(objMath.Calculate(sAFmla, sErrMsg), DEFAULT_DECIMALS)
            If sErrMsg <> "" Then
                subLogErrMsg sErrMsg & ", Invalid Formula (" & sAFmla & ")"
                Exit Function
            End If
        End If
    End If
    
    fnCalculateBonus = dBonusAmt
End Function

Private Function fnGetVarValue(lEmpNo As Long, _
                               nPrftCtr As Integer, _
                               ByVal sVariable As String, _
                               sErrMsg As String) As Double
                               
    Const SUB_NAME As String = "fnGetVarValue"
    
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim lSysParm3004  As Long
    Dim nEmpLevel As Integer
    Dim sDateHired As String
    Dim sDateTerminated As String
    Dim sDateStart As String
    Dim sDateEnd As String
    Dim lDiff As Long
    Dim dDebitAmt As Double
    Dim dCreditAmt As Double
    Dim i As Long
    Dim sarrVariable()
    
    'predefined variables - SHOULD BE THE SAME AS THE DEFINITION IN ZZSEBFMT
    sarrVariable = Array("inside_sales", _
                         "gallons_sold", _
                         "day_off_slip_day", _
                         "total_pay", _
                         "months_in_grade", _
                         "months_as_manager", _
                         "years_as_manager", _
                         "months_employed", _
                         "shortage_amount", _
                         "check_amount", _
                         "pay_hours", _
                         "not used")
    
    fnGetVarValue = 0#
    sErrMsg = ""
    
    sVariable = LCase(sVariable)
    
    Select Case sVariable
        Case sarrVariable(0)  'inside sales
            strSQL = "SELECT bs_sales_amount AS var_value "
            strSQL = strSQL & " FROM bonus_sales"
            strSQL = strSQL & " WHERE bs_prft_ctr = " & nPrftCtr
            strSQL = strSQL & " AND bs_from_date = " & tfnDateString(frmZZSEBPRC.txtStartDate, True)
            strSQL = strSQL & " AND bs_to_date = " & tfnDateString(frmZZSEBPRC.txtEndDate, True)
            strSQL = strSQL & " AND bs_sales_type = " & tfnSQLString(frmZZSEBPRC.txtFrequency)
        
        Case sarrVariable(1)  'gallons sold
            strSQL = "SELECT bs_sales_amount AS var_value "
            strSQL = strSQL & " FROM bonus_sales"
            strSQL = strSQL & " WHERE bs_prft_ctr = " & nPrftCtr
            strSQL = strSQL & " AND bs_from_date = " & tfnDateString(frmZZSEBPRC.txtStartDate, True)
            strSQL = strSQL & " AND bs_to_date = " & tfnDateString(frmZZSEBPRC.txtEndDate, True)
            strSQL = strSQL & " AND bs_sales_type = " & tfnSQLString(sGas)
        
        Case sarrVariable(2)  'day off slip days
            strSQL = "SELECT COUNT (bd_shl) AS var_value "
            strSQL = strSQL & " FROM bonus_day_off_slip, rs_shiftlink"
            strSQL = strSQL & " WHERE bd_empno = " & lEmpNo
            strSQL = strSQL & " AND bd_prft_ctr = " & nPrftCtr
            strSQL = strSQL & " AND bd_slip_date BETWEEN " & tfnDateString(frmZZSEBPRC.txtStartDate, True)
            strSQL = strSQL & " AND " & tfnDateString(frmZZSEBPRC.txtEndDate, True)
            strSQL = strSQL & " AND bd_prft_ctr = rssl_prft_ctr"
            strSQL = strSQL & " AND bd_slip_date = rssl_date"
            strSQL = strSQL & " AND bd_shift = rssl_shift"
        
        Case sarrVariable(3)  'total pay
            strSQL = "SELECT SUM (prci_total) AS var_value"
            strSQL = strSQL & " FROM pr_check_item, pr_check, pr_pay"
            strSQL = strSQL & " WHERE prc_empno = " & lEmpNo
            strSQL = strSQL & " AND prc_lnk = prci_lnk"
            strSQL = strSQL & " AND prc_check_paid <> 'V'"
            strSQL = strSQL & " AND prci_pay_code = prpa_pay_code"
            strSQL = strSQL & " AND prpa_pay_type = 'P'"
            strSQL = strSQL & " AND prc_chk_date BETWEEN " & tfnDateString(frmZZSEBPRC.txtStartDate, True)
            strSQL = strSQL & " AND " & tfnDateString(frmZZSEBPRC.txtEndDate, True)
            
        Case sarrVariable(4)  'months in grade
            strSQL = "SELECT prm_emp_level, prm_date_hired, prm_date_termed"
            strSQL = strSQL & " FROM pr_master"
            strSQL = strSQL & " WHERE prm_empno = " & lEmpNo
            If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
                sErrMsg = "Failed to access the database to get " & sVariable
                Exit Function
            End If
            If rsTemp.RecordCount = 0 Then
                sErrMsg = "Employee record not found for " & sVariable
                Exit Function
            End If
            If IsNull(rsTemp!prm_emp_level) Then
                sErrMsg = "Employee level is NULL for " & sVariable
                Exit Function
            End If
            
            nEmpLevel = tfnRound(rsTemp!prm_emp_level)
            sDateHired = tfnFormatDate(fnGetField(rsTemp!prm_date_hired))
            sDateTerminated = tfnFormatDate(fnGetField(rsTemp!prm_date_termed))
            
            strSQL = "SELECT prhs_effect_dt, prhs_emp_level, prhs_date_hired, prhs_date_termed"
            strSQL = strSQL & " FROM pr_history"
            strSQL = strSQL & " WHERE prhs_empno = " & lEmpNo
            strSQL = strSQL & " ORDER BY prhs_effect_dt"
            If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
                sErrMsg = "Failed to access the database to get " & sVariable
                Exit Function
            End If
            If rsTemp.RecordCount = 0 Then
                sErrMsg = "Employee history record not found for " & sVariable
                
                If IsValidDate(sDateHired) Then
                    If IsValidDate(sDateTerminated) Then
                        lDiff = Abs(DateDiff("m", CDate(sDateHired), CDate(sDateTerminated)))
                    Else
                        lDiff = Abs(DateDiff("m", CDate(sDateHired), CDate(Date)))
                    End If
                End If
                
                fnGetVarValue = lDiff
                Exit Function
            End If
            
            If rsTemp.RecordCount = 1 Then
                If tfnRound(rsTemp!prhs_emp_level) = nEmpLevel Then
                    sDateStart = tfnFormatDate(rsTemp!prhs_effective_dt)
                    sDateEnd = tfnFormatDate(Date)
                    lDiff = Abs(DateDiff("m", CDate(sDateStart), CDate(sDateEnd)))
                End If
            Else
                For i = 1 To rsTemp.RecordCount
                    If tfnRound(rsTemp!prhs_emp_level) = nEmpLevel Then
                        sDateStart = tfnFormatDate(rsTemp!prhs_effective_dt)
                        If i <= rsTemp.RecordCount - 1 Then
                            
                        End If
                    Else
                        sDateStart = ""
                    End If
                    rsTemp.MoveNext
                Next i
            End If
            
            Exit Function
            
        Case sarrVariable(5)  'months as manager
        Case sarrVariable(6)  'years as manager
        Case sarrVariable(7)  'months employed
            strSQL = "SELECT prm_date_hired, prhs_date_termed"
            strSQL = strSQL & " FROM pr_master, pr_history"
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
            
        Case sarrVariable(8)  'shortage amount
            strSQL = "SELECT parm_field FROM sys_parm WHERE parm_nbr = 3004"
            If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
                sErrMsg = "Failed to access the database to get " & tfnSQLString(sVariable)
                Exit Function
            End If
            
            If rsTemp.RecordCount = 0 Then
                sErrMsg = "SysParm#3004 not found for " & sVariable
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
                sErrMsg = "Failed to access the database to get " & tfnSQLString(sVariable)
                Exit Function
            End If
            
            If rsTemp.RecordCount = 0 Then
                sErrMsg = "No record found for " & sVariable
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
            
            fnGetVarValue = dDebitAmt - dCreditAmt
            
            Exit Function
            
        Case sarrVariable(9)  'check_amount
            'the value will be obtained from the formula evaluation
            Exit Function
            
        Case sarrVariable(10)  'pay hours
            strSQL = "SELECT SUM (prci_input_amt) AS var_value"
            strSQL = strSQL & " FROM pr_check_item, pr_check, pr_pay"
            strSQL = strSQL & " WHERE prc_empno = " & lEmpNo
            strSQL = strSQL & " AND prc_lnk = prci_lnk"
            strSQL = strSQL & " AND prc_check_paid <> 'V'"
            strSQL = strSQL & " AND prci_pay_code = prpa_pay_code"
            strSQL = strSQL & " AND prpa_pay_type = 'P'"
            strSQL = strSQL & " AND prpa_calc_method = 'H'"
            strSQL = strSQL & " AND prc_chk_date BETWEEN " & tfnDateString(frmZZSEBPRC.txtStartDate, True)
            strSQL = strSQL & " AND " & tfnDateString(frmZZSEBPRC.txtEndDate, True)
            
        Case sarrVariable(11)  'not used
            Exit Function
        
        Case Else
            sErrMsg = "Variable " + tfnSQLString(sVariable) + " is not defined"
            Exit Function
    
    End Select
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        sErrMsg = "Failed to access the database to get " & tfnSQLString(sVariable)
        Exit Function
    End If
    
    If rsTemp.RecordCount = 0 Then
        sErrMsg = "No record found for " & sVariable
        Exit Function
    End If
    
    If rsTemp.RecordCount > 0 Then
        fnGetVarValue = tfnRound(rsTemp!var_value, DEFAULT_DECIMALS)
    End If
    
End Function

'return error message if any
Private Function fnCheckFormula(ByVal sFormula As String, ByVal sBonusType As String) As String
    Dim i As Integer
    Dim sErrMsg As String
    Dim aryVariables As Variant
    Dim aryValues As Variant
    Dim objEvaluate As clsEquation
    
    On Error GoTo errTrap
    
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
    
errTrap:
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
    
    On Error GoTo errTrap
    
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
    
errTrap:
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
                    " is not valid for Bonus Type " + tfnSQLString(sBonusType)
                Exit Function
            End If
        End If
    Next i

    fnCheckVarAllowed = ""
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
    
    sSType = tfnSQLString(frmZZSEBPRC.fnGetSalesType())
    
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

Private Function fnValidEmployee(lEmployeeNo As Long, _
                                 nPrftCtr As Integer, _
                                 sBGrade As String) As String


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

