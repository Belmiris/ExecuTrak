Attribute VB_Name = "modAPChk"
Option Explicit

'###################################################################
'#Ticket #: 417974. WJ 10/27/2003
'#In what follows, we implement a routine to lock/unlock p_checks.
'#It also generate check nbr or valid check #
'#The stored precedure get_chcke_nbr did not work correctly.
'#
'###################################################################
Private m_FirstLockedCheck As String
Private m_LastLockedCheck As String

Private Function fnFormatPCheck(ByVal lAcct As Long, ByVal lChk As Long)
    fnFormatPCheck = Format(lAcct, String(10, "0")) & Format(lChk, String(10, "0"))
End Function

Private Function fnLockP_Checks(ByVal sPGrp As String, _
                                ByVal sUser As String, _
                                ByRef lStart As Long, _
                                Optional bSingleCheck As Boolean = False) As String
    Const QUERY_FAILED = "SQL query failed. Contact Factor"
    Const TABLENAME = "p_checks"
    Dim strSQL As String
    Dim sTemp As String
    Dim lEnd As Long
    Dim i As Long
    Dim j As Integer
    Dim lChkAcct As Long
    Dim rsTemp As Recordset
    Dim lNumberChecks As Long
    Dim lRecCount As Long
    
    '#Make sure the Print Group is Valid,
    '#then get acct from p group
    sPGrp = Trim(sPGrp)
    
    strSQL = "SELECT pg_chk_acct FROM p_print WHERE pg_group = " & tfnSQLString(sPGrp)
    
    lRecCount = apc_GetRecordSet(rsTemp, strSQL)
    If lRecCount < 0 Then
        fnLockP_Checks = QUERY_FAILED
        Exit Function
    ElseIf lRecCount = 0 Then
        fnLockP_Checks = "Invalid Print Group"
        Exit Function
    End If
    
    lChkAcct = tfnRound(rsTemp!pg_chk_acct)
    
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
     
     '#If we need to generate check, then do it
     If lStart = 0 Then
        strSQL = "SELECT max(pv_check_nbr) max_check_nbr from p_checks WHERE pv_account = " & lChkAcct
        lRecCount = apc_GetRecordSet(rsTemp, strSQL)
        If lRecCount < 0 Then
             fnLockP_Checks = QUERY_FAILED
             Exit Function
        End If
        lStart = tfnRound(rsTemp!max_check_nbr) + 1
     End If
    
    'Validate check #s against p_checks!
    
    lEnd = lStart + lNumberChecks - 1
    strSQL = "SELECT pv_check_nbr FROM p_checks WHERE pv_account = " & lChkAcct _
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
    
    '#Unlock any checks done preciously
    subUnlockP_checks
    
    '#Check the if any of the check is locked
    sTemp = "'" & fnFormatPCheck(lChkAcct, lStart) & "'"
    j = 0
    For i = lStart + 1 To lEnd
        j = j + 1
        sTemp = sTemp & ",'" & fnFormatPCheck(lChkAcct, i) & "'"
        If i = lEnd Or j = 500 Then
            strSQL = "SELECT srl_criteria FROM sys_row_lock" _
                   & " WHERE srl_table = '" & TABLENAME & "'" _
                   & " AND srl_criteria IN (" & sTemp & ")"
            lRecCount = apc_GetRecordSet(rsTemp, strSQL)
            If lRecCount < 0 Then
                 fnLockP_Checks = QUERY_FAILED
                 Exit Function
            ElseIf lRecCount > 0 Then
                fnLockP_Checks = "Check # " & tfnRound(Right(rsTemp!srl_criteria & "", 10)) & " is locked; it falls between start and end numbers"
                Exit Function
            End If
            j = 0
            sTemp = "'" & fnFormatPCheck(lChkAcct, lStart) & "'"
        End If
    Next i
    
    '#Lock the Checks
    strSQL = ""
    j = 0
    For i = lStart To lEnd
        j = j + 1
        strSQL = strSQL & "INSERT INTO sys_row_lock VALUES('" & TABLENAME & "'," & tfnSQLString(LCase(App.EXEName)) & ",'" & tfnGetUserName & "'," & fnFormatPCheck(lChkAcct, i) & ",0);"
        If i = lEnd Or j = 300 Then
            If Not apc_ExecuteSQL(strSQL) Then
                fnLockP_Checks = QUERY_FAILED
                Exit Function
            End If
            strSQL = ""
            j = 0
        End If
    Next i
    
    m_FirstLockedCheck = fnFormatPCheck(lChkAcct, lStart)
    m_LastLockedCheck = fnFormatPCheck(lChkAcct, lEnd)
End Function

Private Sub subUnlockP_checks()
    Const SUB_NAME = "subUnlockP_checks"
    Dim sSql As String
    Dim rsTemp As Recordset
    
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


