Attribute VB_Name = "modShiftReport"
Option Explicit
    
'when new record is inserted into rs_summary, set sStatusChar = STATUS_INSERT.
'otherwise set sStatusChar = rs_summary.rss_status

Private Const STATUS_INSERT = "I"

Private sStatusChar As String

Public Function fnCheckLocked(lProfitCenter As Long) As String
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    fnCheckLocked = ""
    strSQL = "SELECT lock_user FROM sys_lock WHERE lock_name = 'RSE" & CStr(lProfitCenter) & "'"
    On Error GoTo errCheckLocked
    Set rsTemp = t_dbMainDatabase.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    If rsTemp.RecordCount > 0 Then
        fnCheckLocked = rsTemp!lock_user
    End If

extCheckLocked:
    Exit Function

errCheckLocked:
    tfnErrHandler "fnCheckLocked", strSQL
    Resume extCheckLocked
End Function

Public Function fnLockShiftReport(lProfitCenter As Long, _
                                                            Optional lShift As Variant, _
                                                            Optional dDate As Variant, _
                                                            Optional Sale As Variant) As Long

    Dim strLock As String
    Dim lShiftLink As Long
    Dim bSale As Boolean
    
    strLock = "RSE" & CStr(lProfitCenter)
    
    sStatusChar = ""
    
    '// Not sure that we need this but, maybe
    If (process_locked(strLock)) Then
        '// Someone already has this profit center in use
        fnLockShiftReport = -1
        Exit Function
    End If
    
    If IsMissing(lShift) Then
        Exit Function
    End If
    
    If IsMissing(Sale) Then
        bSale = False
    Else
        bSale = Sale
    End If
    
    lShiftLink = fnGetShiftLink(lProfitCenter, CLng(lShift), tfnDateString(dDate))
    Debug.Print "shift link = "; lShiftLink
    If lShiftLink > 0 Then
        If fnLockShiftLink(CLng(lShiftLink), bSale) Then
            fnLockShiftReport = CLng(lShiftLink)
        Else
            fnLockShiftReport = -3
        End If
    Else
        fnLockShiftReport = -2
    End If

End Function

Private Function fnGetShiftLink(lProfitCenter As Long, lShift As Long, dDate As Date) As Long

Dim strSQL As String
Dim rsShiftLink As Recordset
Dim strInsSQL As String

strSQL = "SELECT rssl_shl FROM rs_shiftlink WHERE rssl_prft_ctr = " & CStr(lProfitCenter) & _
                " AND rssl_shift = " & CStr(lShift) & " AND rssl_date = '" & CStr(Format(dDate, "mm/dd/yyyy")) & "'"
                
On Error GoTo GET_SHIFT_LINK_ERROR
    Set rsShiftLink = t_dbMainDatabase.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)

If rsShiftLink.RecordCount > 0 Then
    '//If we have a valid shift link then use it
    fnGetShiftLink = rsShiftLink!rssl_shl
Else
    '// Really not sure what to do here
    strInsSQL = "INSERT INTO rs_shiftlink VALUES(" & CStr(lProfitCenter) & ", " & CStr(lShift) & ", '" & CStr(Format(dDate, "mm/dd/yyyy")) & "', 0)"
    t_dbMainDatabase.ExecuteSQL strInsSQL
    Set rsShiftLink = t_dbMainDatabase.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    fnGetShiftLink = rsShiftLink!rssl_shl
End If

extShiftLink:
On Error GoTo 0
Exit Function
'// Handle errors
GET_SHIFT_LINK_ERROR:
    fnGetShiftLink = 0
    tfnErrHandler "fnGetShiftLink", strSQL
    Resume extShiftLink

End Function
    
Private Function fnLockShiftLink(lShiftLink As Long, Optional Sale As Variant) As Boolean

    Dim rsSummary As Recordset
    Dim strSQL As String
    
    If Not (lShiftLink > 0) Then
        #If DEVELOP Then
            MsgBox "THIS MESSAGE WILL BE DISPLAY ONLY ON DEVELOPMENT." & vbCrLf & vbCrLf & _
                "Shift link: " & lShiftLink & " is not valid.  fnLockShiftLink() do nothing."
        #End If
        fnLockShiftLink = False
        Exit Function
    End If
    
    If IsMissing(Sale) Then Sale = False
    
    strSQL = "SELECT rss_status FROM rs_summary WHERE rss_shl = " & CStr(lShiftLink)
    
    On Error GoTo LOCK_SHIFT_ERROR
        Set rsSummary = t_dbMainDatabase.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    On Error GoTo 0
    
    If rsSummary.RecordCount > 0 Then
        sStatusChar = rsSummary!rss_status
        If sStatusChar <> "R" Then
            '// Can't change anything that's not an R
            fnLockShiftLink = False
            Exit Function
        
        Else
            '// If the status is R then we're OK
            If Sale Then
                strSQL = "UPDATE rs_summary SET rss_status = 'E' WHERE rss_shl = " & CStr(lShiftLink)
            Else
                strSQL = "UPDATE rs_summary SET rss_ap_entry = 'Y' WHERE rss_shl = " & CStr(lShiftLink)
            End If
            
            On Error GoTo LOCK_SHIFT_ERROR
            t_dbMainDatabase.ExecuteSQL strSQL
            On Error GoTo 0
            
            'sStatusChar has already been set to rsSummary!rss_status, see 18 lines above.
'            sStatusChar = STATUS_UPDATE
        End If
    Else
        '// It doesn't exist then add it in
        If Sale Then
            strSQL = "INSERT INTO rs_summary VALUES ( " & CStr(lShiftLink) & " , 0,0,0,0,0,0,0,0,0,0,0,'E','N')"
        Else
            strSQL = "INSERT INTO rs_summary VALUES ( " & CStr(lShiftLink) & " , 0,0,0,0,0,0,0,0,0,0,0,'R','Y')"
        End If
        
        On Error GoTo LOCK_SHIFT_ERROR
        t_dbMainDatabase.ExecuteSQL strSQL
        On Error GoTo 0
        
        sStatusChar = STATUS_INSERT
    End If
    
    fnLockShiftLink = True
    
Exit Function

LOCK_SHIFT_ERROR:
    fnLockShiftLink = False
    tfnErrHandler "fnLockShiftLink", strSQL

End Function

'david - 7-25-97
'call this function with lShiftLink parameter ONLY WHEN the user CANCEL a shift report
'upon UPDATING shift report DO NOT call this function with lShiftLink parameter!!!
Public Sub fnUnlockShiftReport(lProfitCenter As Long, Optional lShiftLink As Variant)

    Dim strLock As String
    Dim strSQL As String
    
    strLock = "RSE" & CStr(lProfitCenter)
    
    unlock_process (strLock)
    
    If sStatusChar < "A" Or sStatusChar > "Z" Then
        sStatusChar = ""
    End If
    
    If Not IsMissing(lShiftLink) Then
        If lShiftLink > 0 Then
            If sStatusChar = STATUS_INSERT Then
                strSQL = "DELETE FROM rs_summary WHERE rss_shl = " & CStr(lShiftLink)
                
                On Error GoTo errUnlock
                t_dbMainDatabase.ExecuteSQL strSQL
                On Error GoTo 0
            ElseIf sStatusChar <> "" Then
                strSQL = "UPDATE rs_summary SET rss_status = " & tfnSQLString(sStatusChar) & " WHERE rss_shl = " & CStr(lShiftLink)
                
                On Error GoTo errUnlock
                    t_dbMainDatabase.ExecuteSQL strSQL
                On Error GoTo 0
            End If
        Else
            #If DEVELOP Then
                MsgBox "THIS MESSAGE WILL BE DISPLAY ONLY ON DEVELOPMENT." & vbCrLf & vbCrLf & _
                    "Shift link: " & lShiftLink & " is not valid.  fnUnlockShiftReport() do nothing."
            #End If
        End If
    End If

extUnlock:
    Exit Sub

errUnlock:
    tfnErrHandler "fnUnlockShiftReport", strSQL
    Resume extUnlock
End Sub

