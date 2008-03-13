Attribute VB_Name = "modAPChk"
Option Explicit

'###################################################################
'#Ticket #: 417974. WJ 10/27/2003
'#In what follows, we implement a routine to lock/unlock p_checks.
'#It also generate check nbr or valid check #
'#The stored precedure get_chcke_nbr did not work correctly since it did
'#not ask check account. This module will be used in all VB AP check #
'#validations.
'#
'# How to use this module
'#(1)Call fnLockP_Checks(Prt_grp,User,lstart,bSingleCheck) in check validation
'#   if only one check will be printed , pass bSingleCheck= true
'#   if want to generate start check, user lStart = 0 , a valid start check
'#   will be returned(byref). If check is locked sucessfully, fnLockP_Checks=SZEMPTY
'#   otherwise, error message will be returned.
'#(2)call subUnlockP_checks after AP check is done or Canceled
'###################################################################
'# Modification History
'   Programmer  :   Vijaya B Alla
'   Magic No    :   468962
'   Description :   Bank Reconciliation
'   Changes     :   Bank Account Changes from Integer to Char(17) in
'                   p_checks.pv_account Changes madef long to string
'                   but only enters numbers in sys_bnk_acct_hdr table
'                   no characters
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
Private m_FirstLockedCheck As String
Private m_LastLockedCheck As String

Private Function fnFormatPCheck(ByVal sAcct As String, ByVal lChk As Long) As String
    fnFormatPCheck = Format(sAcct, String(17, "0")) & Format(lChk, String(10, "0"))
End Function

Public Function fnLockP_Checks(ByVal sPGrp As String, _
                                ByVal sUser As String, _
                                ByRef lStart As Long, _
                                Optional bSingleCheck As Boolean = False) As String
    Const QUERY_FAILED = "SQL query failed. Contact Factor"
    Const TableName = "p_checks"
    Dim strSQL As String
    Dim sTemp As String
    Dim lEnd As Long
    Dim i As Long
    Dim j As Integer
    Dim sChkAcct As String
    Dim rsTemp As Recordset
    Dim lNumberChecks As Long
    Dim lRecCount As Long
    Dim bGenerateStartChk As Boolean
    'david 03/12/2008  #583727
    'fixed duplicate check number (after entering 0)
    'Dim sLockData As String * 20
    Dim sLockData As String
    Dim lReservedNbr As Long
    
    On Error GoTo errTrap
    
    '#Make sure the Print Group is Valid,
    '#then get acct from p group
    sPGrp = Trim(sPGrp)
    sUser = Trim(sUser)
    
    strSQL = "SELECT pg_chk_acct FROM p_print WHERE pg_group = " & tfnSQLString(sPGrp)
    
    lRecCount = apc_GetRecordSet(rsTemp, strSQL)
    If lRecCount < 0 Then
        fnLockP_Checks = QUERY_FAILED
        Exit Function
    ElseIf lRecCount = 0 Then
        fnLockP_Checks = "Invalid Print Group"
        Exit Function
    End If
    
    sChkAcct = fnCStr(rsTemp!pg_chk_acct)
    
    '#Figure out how many checks to write
    'build SQL to verify number of check to be printed
    If bSingleCheck Then
        lNumberChecks = 1
    Else
        strSQL = "SELECT COUNT(DISTINCT ps_vendor) FROM p_selection" + _
                 " WHERE ps_print_group=" & tfnSQLString(sPGrp)
    
        If sUser <> "" Then
            strSQL = strSQL + " AND ps_user_id=" & tfnSQLString(Trim(sUser))
        End If
                    
        lRecCount = apc_GetRecordSet(rsTemp, strSQL)
        If lRecCount < 0 Then
             fnLockP_Checks = QUERY_FAILED
             Exit Function
        ElseIf lRecCount = 0 Then
            fnLockP_Checks = "No records found for payment"
            Exit Function
        End If
        lNumberChecks = tfnRound(rsTemp.Fields(0))
        If lNumberChecks = 0 Then
            fnLockP_Checks = "No records found for payment"
            Exit Function
        End If
     End If
      
     
     '#Unlock any checks done preciously
     subUnlockP_checks
    
     '#If we need to generate check, then do it
     If lStart = 0 Then
        bGenerateStartChk = True
     Else
        bGenerateStartChk = False
     End If
     If bGenerateStartChk Then
        strSQL = "SELECT max(pv_check_nbr) max_check_nbr from p_checks WHERE pv_account = " & tfnSQLString(sChkAcct)
        lRecCount = apc_GetRecordSet(rsTemp, strSQL)
        If lRecCount < 0 Then
             fnLockP_Checks = QUERY_FAILED
             Exit Function
        End If
        lStart = tfnRound(rsTemp!max_check_nbr) + 1
        
        '#Check if any chk # are reserved
        strSQL = "SELECT srl_criteria FROM sys_row_lock" _
                & " WHERE srl_table = '" & TableName & "'" _
                & " AND srl_criteria[1,17] = '" & Format(sChkAcct, String(17, "0")) & "'"
        strSQL = strSQL & " ORDER BY srl_criteria DESC"
        If apc_GetRecordSet(rsTemp, strSQL) > 0 Then
            sLockData = rsTemp!srl_criteria & ""
            lReservedNbr = Val(Mid(sLockData, 18, 10))
            If lStart <= lReservedNbr Then
                lStart = lReservedNbr + 1
            End If
        End If
        lEnd = lStart + lNumberChecks - 1
     Else
        '#Validate check #s against p_checks!
        lEnd = lStart + lNumberChecks - 1
        strSQL = "SELECT pv_check_nbr FROM p_checks WHERE pv_account = " & tfnSQLString(sChkAcct) _
               & " AND pv_check_nbr BETWEEN " & CStr(lStart) & " AND " & CStr(lEnd)
        
        lRecCount = apc_GetRecordSet(rsTemp, strSQL)
        If lRecCount < 0 Then
             fnLockP_Checks = QUERY_FAILED
             Exit Function
        End If
        If lRecCount > 0 Then
            fnLockP_Checks = "Check # " & tfnRound(rsTemp!pv_check_nbr) & " is used; it falls between start and end numbers"
            Exit Function
        End If
    
        '#Check the if any of the check is locked
        strSQL = "SELECT srl_criteria FROM sys_row_lock" _
                & " WHERE srl_table = '" & TableName & "'" _
                & " AND srl_criteria[1,17] = '" & Format(sChkAcct, String(17, "0")) & "'"
        strSQL = strSQL & " AND srl_criteria BETWEEN '" & fnFormatPCheck(sChkAcct, lStart) & "'"
        strSQL = strSQL & " AND '" & fnFormatPCheck(sChkAcct, lEnd) & "'"
        If apc_GetRecordSet(rsTemp, strSQL) > 0 Then
            fnLockP_Checks = "Check # " & tfnRound(Right(Trim(rsTemp!srl_criteria & ""), 10)) & " is locked; it falls between start and end numbers"
            Exit Function
        End If
    End If
    '#Lock the Checks
    strSQL = ""
    j = 0
    For i = lStart To lEnd
        j = j + 1
        strSQL = strSQL & "INSERT INTO sys_row_lock VALUES('" & TableName & "'," & tfnSQLString(LCase(App.EXEName)) & ",'" & tfnGetUserName & "','" & fnFormatPCheck(sChkAcct, i) & "',0);"
        If i = lEnd Or j = 300 Then
            If Not apc_ExecuteSQL(strSQL) Then
                m_FirstLockedCheck = fnFormatPCheck(sChkAcct, lStart)
                m_LastLockedCheck = fnFormatPCheck(sChkAcct, lEnd)
                fnLockP_Checks = QUERY_FAILED
                Exit Function
            End If
            strSQL = ""
            j = 0
        End If
    Next i
    
    m_FirstLockedCheck = fnFormatPCheck(sChkAcct, lStart)
    m_LastLockedCheck = fnFormatPCheck(sChkAcct, lEnd)
    
    fnLockP_Checks = szEMPTY
    Exit Function
errTrap:
    fnLockP_Checks = "Error: " & Err.Description
End Function

Public Sub subUnlockP_checks()
    Const SUB_NAME = "subUnlockP_checks"
    Dim sSql As String
    Dim rsTemp As Recordset
    If m_FirstLockedCheck = "" Then
       Exit Sub
    End If
    sSql = "DELETE FROM sys_row_lock" _
         & " WHERE srl_table = 'p_checks'" _
         & " AND srl_prog_id = " & tfnSQLString(LCase(App.EXEName)) _
         & " AND srl_user_id = '" & tfnGetUserName & "'" _
         & " AND srl_criteria BETWEEN " & tfnSQLString(m_FirstLockedCheck) _
         & " AND " & tfnSQLString(m_LastLockedCheck)
    apc_ExecuteSQL sSql
    m_FirstLockedCheck = ""
    m_LastLockedCheck = ""
    
End Sub
Private Function apc_GetRecordSet(rsTemp As Recordset, szSQL As String) As Long
    On Error GoTo SQLError
    
    Set rsTemp = t_dbMainDatabase.OpenRecordset(szSQL, dbOpenSnapshot, dbSQLPassThrough)
    If rsTemp.RecordCount > 0 Then
       rsTemp.MoveLast
       rsTemp.MoveFirst
    End If
    
    apc_GetRecordSet = rsTemp.RecordCount
    Exit Function
SQLError:
    apc_GetRecordSet = -1
    tfnErrHandler "apc_GetRecordSet", szSQL, True
    On Error GoTo 0
End Function

Private Function apc_ExecuteSQL(szSQL As String, _
                Optional szCalledFrom As Variant, Optional bShowError As Variant) As Boolean
                
      Dim szMsg As String
      
      On Error GoTo SQLError
      
      t_dbMainDatabase.ExecuteSQL szSQL
      apc_ExecuteSQL = True
      Exit Function
      
SQLError:
      apc_ExecuteSQL = False
      If IsMissing(szCalledFrom) Then
         szCalledFrom = ""
      End If
      If IsMissing(bShowError) Then
         bShowError = True
      End If
      tfnErrHandler "Apc_ExecuteSQL, " & szCalledFrom, szSQL, bShowError
      On Error GoTo 0
End Function

Private Function fnCStr(vIn) As String
    fnCStr = Trim(vIn & "")
End Function
