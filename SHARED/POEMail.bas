Attribute VB_Name = "modPOEMail"
Option Explicit
'Programmer : Rajneesh Aggarwal(20 April 00)
'Created for Purchase Order Module only.
'



Public Function fnCheckApprovalAuthority(sProgramID As String, _
                                          vPurchaseNumber As Variant, _
                                          nPrftCtr As Integer, _
                                          dPurchaseTotal As Double) As Boolean
    Const SUB_NAME As String = "fnCheckApprovalAuthority"
    Dim sUserName As String
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim sMsg As String
    Dim sName As String 'Concatenated First and Last Names
    Dim sSubject As String
    Dim sEMailMsg As String
    Dim sEMailAdd As String
    Dim sSuperID As String
    Dim sProgram As String
        
    fnCheckApprovalAuthority = False
    Screen.MousePointer = vbHourglass
    sUserName = Trim(tfnGetUserName())
    
    If Not fnGetName(sUserName, sName, sEMailAdd) Then
        sName = sUserName
    End If
    
    Select Case LCase(sProgramID)
        Case "poerentr"
            sProgram = " Request"
        Case "poeoentr"
            sProgram = " Order"
        Case "pofbrspr"
            sProgram = " Request"
        Case "pofbrspo"
            sProgram = " Order"
        Case "poeselct"
            sProgram = " Selection"
    End Select
    
    sSubject = "A Purchase" & sProgram & " requires your approval"
    sEMailMsg = sName & " has attempted to approve Purchase" & sProgram & " number '"
    sEMailMsg = sEMailMsg & vPurchaseNumber & "', but lacked sufficient approval authority. "
    sEMailMsg = sEMailMsg & "Please review this Purchase" & sProgram & ", and approve or cancel it."
    
    strSQL = "SELECT pos_prft_ctr, pos_user_id, pos_approv_level, pos_super_userid"
    strSQL = strSQL & ", pos_pr_approv, pos_po_approv, polv_prft_ctr"
    strSQL = strSQL & ", polv_level, polv_level_desc, polv_auth_amount, sum_user_id "
    strSQL = strSQL & ", sum_first_name, sum_last_name"
    strSQL = strSQL & " FROM po_security, po_levels, sys_user_master"
    strSQL = strSQL & " WHERE pos_prft_ctr = " & nPrftCtr
    strSQL = strSQL & " AND pos_user_id = sum_user_id"
    strSQL = strSQL & " AND pos_prft_ctr = polv_prft_ctr"
    strSQL = strSQL & " AND pos_approv_level = polv_level"
    strSQL = strSQL & " AND pos_user_id = " & tfnSQLString(sUserName)
    
    If GetRecordSet(rsTemp, strSQL, , "SUB_NAME") < 0 Then
        MsgBox "Failed to access database to check user's authority.", vbCritical
        Exit Function
    End If
    
    If rsTemp.RecordCount = 0 Then
        MsgBox "Authorization record not found for the user.", vbInformation
        Exit Function
    End If
    
    'Check the approval flag in the PO Security table for the profit center.
    If sProgramID = "POFBRSPR" Or sProgramID = "POERENTR" Or sProgramID = "POESELCT" Then
        If fnCstr(rsTemp!pos_pr_approv) <> "Y" Then
            MsgBox "Authorization Failed, You are not authorized to approve this Purchase" _
                    & sProgram, vbInformation
            Exit Function
        End If
    ElseIf sProgramID = "POFBRSPO" Or sProgramID = "POEOENTR" Then
        If fnCstr(rsTemp!pos_po_approv) <> "Y" Then
            MsgBox "Authorization Failed, You are not authorized to approve this Purchase" _
                    & sProgram, vbInformation
            Exit Function
        End If
    End If
            
    sMsg = "Authorization failed, Purchase" & sProgram & " value exceeds sanctioned limit of $"
    sMsg = sMsg & tfnRound(rsTemp!polv_auth_amount, DEFAULT_DECIMALS) & "."
    sSuperID = fnCstr(rsTemp!pos_super_userid)
    
    Screen.MousePointer = vbHourglass
    If tfnRound(dPurchaseTotal, DEFAULT_DECIMALS) > tfnRound(rsTemp!polv_auth_amount, 6) Then
        If sSuperID <> "" Then 'Get Supervisor Name and his EMail Address
            If Not fnGetName(sSuperID, sName, sEMailAdd) Then 'Show Supervisor's UserID
                sMsg = sMsg & " Please ask '" & sSuperID
            Else 'Show Supervisor's Name
                sMsg = sMsg & " Please ask '" & sName
            End If
            sMsg = sMsg & "' to approve this Purchase" & sProgram & "."
        End If
        'Send an E-Mail Message to user's supervisor...
        If fnSendEmail(sProgramID, sEMailAdd, sSubject, sEMailMsg) Then
            Screen.MousePointer = vbHourglass
            tfnWaitSeconds 4
        End If
        MsgBox sMsg, vbInformation
        Exit Function
    End If
    
    fnCheckApprovalAuthority = True
            
End Function

Private Function fnSendEmail(sProgramID As String, sE_MailAddress As String, _
                             sE_MailSubject As String, sE_MailMessage As String) As Boolean
    Const sDQ = """"
    Dim sParm As String
    
    fnSendEmail = False
    
    If sE_MailAddress = "" Or sE_MailMessage = "" Then
        Exit Function
    End If
    
    'This function will return the UserID and Password if the sys_parm is found
    If Not fnCheckSysParam() Then
        Exit Function
    End If
    
    sParm = "sendmail " & sDQ & "HIDE" & sDQ & " "
    sParm = sParm & sDQ & Trim(sProgramID) & sDQ & " "
    sParm = sParm & sDQ & Trim(sProgramID) & sDQ & " "
    sParm = sParm & sDQ & Trim(sE_MailAddress) & sDQ & " "
    sParm = sParm & sDQ & Trim(sE_MailSubject) & sDQ & " "
    sParm = sParm & sDQ & Trim(sE_MailMessage) & sDQ
    
    If tfnRun(sParm) Then
        fnSendEmail = True
    End If
    
End Function
    
Private Function fnCheckSysParam() As Boolean
    Const SUB_NAME As String = "fnCheckSysParam"
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    fnCheckSysParam = False
    
    strSQL = "SELECT parm_field, parm_desc FROM sys_parm WHERE parm_nbr = 13007"
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) <= 0 Then
        Exit Function
    End If
    
    'System parameter found
    If fnCstr(rsTemp.parm_field) = "Y" Then
        fnCheckSysParam = True
    End If
    
'    sUserID = fnCstr(rsTemp!parm_field)
'    sPassword = fnCstr(rsTemp!parm_desc)

End Function

Public Function fnCstr(v) As String
    If Not IsNull(v) Then
        fnCstr = Trim(v)
    End If
End Function

Private Function fnGetName(sID As String, sName As String, sEMailAdd As String) As Boolean
    Const SUB_NAME As String = ""
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    fnGetName = False
    
    strSQL = "SELECT sum_employee_flag, sum_empno, sum_e_mail_addy FROM sys_user_master WHERE "
    strSQL = strSQL & " sum_user_id = " & tfnSQLString(Trim(sID))
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) <= 0 Then
        Exit Function
    End If
    
    If rsTemp!sum_employee_flag = "Y" Then
        strSQL = "SELECT prm_first_name AS sum_first_name, prm_last_name AS sum_last_name "
        strSQL = strSQL & " FROM pr_master WHERE prm_empno = " & rsTemp!sum_empno
    Else 'Not an Employee
        strSQL = "SELECT sum_first_name, sum_last_name FROM sys_user_master WHERE"
        strSQL = strSQL & " sum_user_id = " & tfnSQLString(Trim(sID))
    End If
    
    sEMailAdd = fnCstr(rsTemp!sum_e_mail_addy)
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) <= 0 Then
        Exit Function
    End If

    sName = Trim(rsTemp!sum_first_name) & " " & Trim(rsTemp!sum_last_name)
    
    fnGetName = True
    
End Function

Private Function tfnRun(szExeName As String, Optional vWindowStyle) As Boolean
    Const SHELL_OK As Long = 32
    Dim szCmd As String
    Dim hTempInstance As Long
    
    #If FACTOR_MENU < 0 Then
        Const gszBINROOT As String = ".\bin\"
    #Else
        Const gszBINROOT As String = "g:\program\factmenu\bin\"
    #End If

    On Error GoTo ErrorRun
    
    If IsMissing(vWindowStyle) Then
        vWindowStyle = SW_SHOWNORMAL
    End If
    
    szCmd = gszBINROOT & szExeName
    
    hTempInstance = shell(szCmd, vWindowStyle)
    
    'if hInstance greater than 32 application is running
    If hTempInstance > SHELL_OK Or hTempInstance < 0 Then
        tfnRun = True 'application running
    Else
        tfnRun = False 'application failed to launch
    End If

    Exit Function

ErrorRun:
    #If NO_ERROR_HANDLER Then
        MsgBox "Cannot execute program" & vbCrLf & Err.Description
    #Else
        tfnErrHandler "tfnRun"
    #End If
End Function

