Attribute VB_Name = "modzzeininv"
Option Explicit

Public dbLocal As DataBase
Public Const nDB_LOCAL As Integer = 1
Public Const nDB_REMOTE As Integer = 2

Public Const ColHeaderPrftctr As Integer = 0
Public Const ColHeaderPrftDesc As Integer = 1
Public Const ColHeaderRptDate As Integer = 2
Public Const ColHeaderVendor As Integer = 3
Public Const ColHeaderVendorName As Integer = 4
Public Const ColHeaderInvoice As Integer = 5
Public Const ColHeaderInvAmount As Integer = 6
Public Const ColHeaderStatus As Integer = 7
Public ColHdnHeaderShift As Integer
Public ColHdnHeaderTerm As Integer
Public ColHdnHeaderType As Integer
Public ColHdnHeaderDraft As Integer
Public colHdnHeaderInvDate As Integer
Public Const ColDetailLine As Integer = 0
Public Const ColDetailItemCode As Integer = 1
Public Const ColDetailItemDesc As Integer = 2
Public Const ColDetailUOM As Integer = 3
Public Const ColDetailQty As Integer = 4
Public Const ColDetailCost As Integer = 5
Public Const ColDetailExtCost As Integer = 6
Public Const ColDetailPBCost As Integer = 7
Public Const ColDetailExtPBCost As Integer = 8
Public Const ColDetailRetail As Integer = 9
Public Const ColDetailExtRetail As Integer = 10
'use to check the cost is equal to price book or not. 'Y' is Equal, 'N' is not equal
Public colHdnDetailFlag As Integer
Public Const nMaxDetailCol As Integer = 11

Public Const DBASEDATE As Date = "12/31/1899"

Private Const STATUS_INSERT = "I"
Private sStatusChar As String

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

Public Function fnCheckLocked(nProfitCenter As Integer) As String
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    fnCheckLocked = ""
    strSQL = "SELECT lock_user FROM sys_lock WHERE lock_name = 'RSE" & CStr(nProfitCenter) & "'"
    
    If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, "fnCheckLocked") > 0 Then
        fnCheckLocked = rsTemp!lock_user
    End If

End Function

Public Function fnLockShiftReport(nProfitCenter As Integer, lShift As Long, dDate As Date, sLockUser As String) As Long

    Dim strLock As String
    Dim lShiftLink As Long
    
    strLock = "RSE" & CStr(nProfitCenter)
    
    If fnLockProcess(strLock, sLockUser) Then
        '// Someone already has this profit center in use
        fnLockShiftReport = -1
        Exit Function
    End If
    
    lShiftLink = fnGetShiftLink(nProfitCenter, lShift, tfnDateString(dDate))
    
    If lShiftLink > 0 Then
        
        If fnLockShiftLink(lShiftLink) Then
            fnLockShiftReport = lShiftLink
        Else
            fnLockShiftReport = -3
        End If
        
    Else
        fnLockShiftReport = -2
    End If

End Function

Private Function fnGetShiftLink(nProfitCenter As Integer, lShift As Long, dDate As Date) As Long
    Dim strSQL As String
    Dim rsShiftLink As Recordset
    Dim strInsSQL As String
    Const FUNC_NAME As String = "fnGetShiftLink"

    strSQL = "SELECT rssl_shl FROM rs_shiftlink WHERE rssl_prft_ctr = " & nProfitCenter
    strSQL = strSQL & " AND rssl_shift = " & lShift
    strSQL = strSQL & " AND rssl_date = " & tfnDateString(dDate, True)
                
    If fnGetRecord(rsShiftLink, strSQL, nDB_REMOTE, FUNC_NAME) > 0 Then
        fnGetShiftLink = rsShiftLink!rssl_shl
    ElseIf rsShiftLink.RecordCount = 0 Then
        strInsSQL = "INSERT INTO rs_shiftlink VALUES(" & nProfitCenter & ", " & lShift & "," & tfnDateString(dDate, True) & ", 0)"
        
        If fnExecuteSQL(strInsSQL, nDB_REMOTE, FUNC_NAME) Then
            
            If fnGetRecord(rsShiftLink, strSQL, nDB_REMOTE, FUNC_NAME) > 0 Then
                fnGetShiftLink = rsShiftLink!rssl_shl
            Else
                fnGetShiftLink = 0
            End If
            
        Else
            fnGetShiftLink = 0
        End If
        
    Else
        fnGetShiftLink = 0
    End If
            
End Function
    
Private Function fnLockShiftLink(lShiftLink As Long, Optional bSale As Boolean) As Boolean
    Dim rsSummary As Recordset
    Dim strSQL As String
    Const FUNC_NAME As String = "fnLockShiftLink"
    
    If lShiftLink <= 0 Then
        fnLockShiftLink = False
        Exit Function
    End If
    
    If IsMissing(bSale) Then bSale = False
    
    strSQL = "SELECT rss_status FROM rs_summary WHERE rss_shl = " & lShiftLink
    
    If fnGetRecord(rsSummary, strSQL, nDB_REMOTE, FUNC_NAME) > 0 Then
        sStatusChar = rsSummary!rss_status
        
        If sStatusChar <> "R" Then
            '// Can't change anything that's not an R
            fnLockShiftLink = False
            Exit Function
        Else
            '// If the status is R then we're OK
            If bSale Then
                strSQL = "UPDATE rs_summary SET rss_status = 'E' WHERE rss_shl = " & CStr(lShiftLink)
            Else
                strSQL = "UPDATE rs_summary SET rss_ap_entry = 'Y' WHERE rss_shl = " & CStr(lShiftLink)
            End If
            
            If Not fnExecuteSQL(strSQL, nDB_REMOTE, FUNC_NAME) Then
                fnLockShiftLink = False
                Exit Function
            End If
            
        End If
        
    Else
        If bSale Then
            strSQL = "INSERT INTO rs_summary VALUES ( " & lShiftLink & " , 0,0,0,0,0,0,0,0,0,0,0,'E','N')"
        Else
            strSQL = "INSERT INTO rs_summary VALUES ( " & lShiftLink & " , 0,0,0,0,0,0,0,0,0,0,0,'R','Y')"
        End If
        
        If Not fnExecuteSQL(strSQL, nDB_REMOTE, FUNC_NAME) Then
            fnLockShiftLink = False
            Exit Function
        End If
        
        sStatusChar = STATUS_INSERT
    End If
    
    fnLockShiftLink = True
End Function

Public Function fnUnlockShiftReport(nProfitCenter As Integer, lShiftLink As Long) As Boolean
    Dim strLock As String
    Dim strSQL As String
    
    strLock = "RSE" & CStr(nProfitCenter)
    
    If Not fnUnlockProcess(strLock) Then
        Exit Function
    End If
    
    If lShiftLink > 0 Then
    
        If sStatusChar = STATUS_INSERT Then
            strSQL = "DELETE FROM rs_summary WHERE rss_shl = " & lShiftLink
        ElseIf sStatusChar <> "" Then
            strSQL = "UPDATE rs_summary SET rss_status = " & tfnSQLString(sStatusChar) & " WHERE rss_shl = " & lShiftLink
        End If
        
        If fnExecuteSQL(strSQL, nDB_REMOTE, "fnUnlockShiftReport") Then
            fnUnlockShiftReport = True
        End If
        
    End If
    
End Function


Public Function fnLockProcess(sPID As String, Optional sLockUser As String) As Boolean
    Const FUNC_NAME = "fnLockProcess"
    Dim sUser As String
    Dim rsTemp As Recordset
    Dim strSQL As String
    
    If Not IsMissing(sLockUser) Then
        sLockUser = ""
    End If
    
    strSQL = "SELECT lock_user FROM sys_lock WHERE lock_name = " & tfnSQLString(sPID)
    
    If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, FUNC_NAME) < 0 Then
        fnLockProcess = True
        Exit Function
    ElseIf rsTemp.RecordCount = 0 Then
        sUser = tfnGetUserName()
        strSQL = "INSERT INTO sys_lock VALUES (" & tfnSQLString(sPID) & "," & tfnSQLString(sUser) & ")"
        
        If fnExecuteSQL(strSQL, nDB_REMOTE, FUNC_NAME) Then
            fnLockProcess = False
        Else
            fnLockProcess = True
        End If
        
    Else
        fnLockProcess = True
            
        If IsNull(rsTemp!lock_user) Or (rsTemp!lock_user & "" = "") Then
            sLockUser = "UNKNOWN"
        Else
            sLockUser = Trim(rsTemp!lock_user & "")
        End If
        
    End If
    
End Function

Public Function fnUnlockProcess(sPID As String) As Boolean
    Dim strSQL As String
    Const FUNC_NAME As String = "fnUnlockProcess"
    
    strSQL = "DELETE FROM sys_lock WHERE lock_name = " & tfnSQLString(sPID)
    
    If fnExecuteSQL(strSQL, nDB_REMOTE, FUNC_NAME) Then
        fnUnlockProcess = True
    Else
        fnUnlockProcess = False
    End If
    
End Function

Public Function fnGetPayGLAccount(sPayType As String) As Long
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim sParm As String
    
    Select Case sPayType
        Case "C":
            sParm = "3004"
        Case "D"
            sParm = "3007"
        Case "P"
            sParm = "3008"
        'usually no transfer type
        Case "T"
            sParm = "3101"
        Case "M"
            sParm = "3017"
    End Select
    
    strSQL = "SELECT parm_nbr, parm_field FROM sys_parm" _
           & " WHERE parm_nbr = " & tfnSQLString(sParm)
    If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, "fnGetGLAccount") > 0 Then
        fnGetPayGLAccount = tfnRound(rsTemp!parm_field)
    Else
        Exit Function
    End If

End Function

Public Function fnCheckGLStatus(dtRptDate As Date) As String
    Dim strSQL As String
    Dim rsTemp As Recordset
    Const FUNC_NAME As String = "fnCheckGLStatus"
    
    strSQL = "SELECT glp_status FROM gl_period WHERE " & tfnDateString(dtRptDate, True) & " BETWEEN glp_beg_dt AND glp_end_dt"
    
    If fnGetRecord(rsTemp, strSQL, nDB_REMOTE, FUNC_NAME) > 0 Then
        If rsTemp!glp_status = "C" Or rsTemp!glp_status = "G" Then
            fnCheckGLStatus = "G/L period is closed"
        Else
            fnCheckGLStatus = ""
        End If

    ElseIf rsTemp.RecordCount <= 0 Then
        fnCheckGLStatus = "G/L period is not setup for the date"
    End If

End Function

Public Function fnDateToSQLIntegerString(sDate As Date) As String
    'This function converts a date to an integer string
    'whose value is the number of days elapsed from 12/31/1899 (DBASEDATE)
     
    Dim nDays As Long
    
    nDays = DateDiff("d", DBASEDATE, sDate)
    
    fnDateToSQLIntegerString = CStr(nDays)

End Function

