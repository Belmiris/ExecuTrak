Attribute VB_Name = "modPOEMail"
Option Explicit

Public Function fnSendEmail(sE_MailAddress As String, sE_MailSubject As String, _
                            sE_MailMessage As String) As Boolean
    Dim sUserID As String
    Dim sPassword As String
    Const sDQ = """"
    Dim sParm As String
    
    fnSendEmail = False
    
    If sE_MailAddress = "" Or sE_MailMessage = "" Then
        Exit Function
    End If
    
    'This function will return the UserID and Password if the sys_parm is found
    If Not fnCheckSysParam(sUserID, sPassword) Then
        Exit Function
    End If
    
    sParm = "sendmail " & sDQ & "SHOW" & sDQ & " "
    sParm = sParm & sDQ & Trim(sUserID) & sDQ & " "
    sParm = sParm & sDQ & Trim(sPassword) & sDQ & " "
    sParm = sParm & sDQ & Trim(sE_MailAddress) & sDQ & " "
    sParm = sParm & sDQ & Trim(sE_MailSubject) & sDQ & " "
    sParm = sParm & sDQ & Trim(sE_MailMessage) & sDQ
    
    If tfnRun(sParm) Then
        fnSendEmail = True
    End If
    
End Function
    
Private Function fnCheckSysParam(sUserID As String, sPassword As String) As Boolean
    Const SUB_NAME As String = "fnCheckSysParam"
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    fnCheckSysParam = False
    
    strSQL = "SELECT parm_field, parm_desc FROM sys_parm WHERE parm_nbr = 13007"
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) <= 0 Then
        Exit Function
    End If
    
    'System parameter found
    sUserID = fnCstr(rsTemp!parm_field)
    sPassword = fnCstr(rsTemp!parm_desc)
    
    If sUserID <> "" Then
        fnCheckSysParam = True
    End If

End Function

Private Function fnCstr(v) As String
    If Not IsNull(v) Then
        fnCstr = Trim(v)
    End If
End Function

Public Function fnGetName(sID As String, sName As String, sEMailAdd As String) As Boolean
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

