Attribute VB_Name = "modRCmd"
Option Explicit

Private Const ERR_LOGIN = -1
Private Const MAX_MSG_LEN = 4096
Private Const WINSOCK_PORT = 512

Private Const RUN_TIME_RCMD = -234
Private Const RUN_TIME_PROC = -334
Private Const ERR_RCMD_MISSING = -430
Private Const ERR_MSG_RETURNED = -431
Private Const ERR_MSG_RUN4GE = -432

Private Const CONNECT_HOST = ";HOST"
Private Const CONNECT_DBPATH1 = ";DB"
Private Const CONNECT_DBPATH2 = "DATABASE"

Private Const CONNECT_USERID = ";UID"
Private Const CONNECT_PSWD = ";PWD"

Public Const FLAG_IDX_APVPRINT As Integer = 1
Private Const FLAG_CAP_APVPRINT = "APVPRINT"
Public Const FLAG_IDX_APVOID As Integer = 2
Private Const FLAG_CAP_APVOID = "APVOID"
Public Const FLAG_IDX_PRVPRINT As Integer = 3
Private Const FLAG_CAP_PRVPRINT = "PRVPRINT"
Public Const FLAG_IDX_OEEINVCE As Integer = 4
Private Const FLAG_CAP_OEEINVCE = "OEEINVCE"
Public Const FLAG_IDX_FMESTORE As Integer = 5
Private Const FLAG_CAP_FMESTORE = "FMESTORE"
Private Const FLAG_COUNT = 5

Private Const PARM_RUN_4GE = 500
Public Const USE_STORED_PROC = 0
Public Const USE_RCMD = 1
'Private nWhatToUse As Integer
Public nWhatToUse As Integer

#If Win32 Then
    Declare Function WinsockRCmd Lib "RCMD32.DLL" (ByVal RHost As String, ByVal RPort As Long, ByVal LocalUser As String, ByVal RemoteUser As String, ByVal Cmd As String, ByVal ErrorMsg As String, ByVal ErrLen As Long) As Long
    Declare Function RCmdRead Lib "RCMD32.DLL" (ByVal hRCmd As Long, ByVal RData As String, ByVal RCount As Long) As Long
    Declare Function RCmdReadByte Lib "RCMD32.DLL" (ByVal hRCmd As Long) As Long
    Declare Function RCmdClose Lib "RCMD32.DLL" (ByVal hRCmd As Long) As Long
    Declare Function RCmdHandle Lib "RCMD32.DLL" (ByVal hRCmd As Long) As Long
    Declare Function RCmdSend Lib "RCMD32.DLL" (ByVal hRCmd As Long, ByVal RData As String, ByVal RCount As Long) As Long
#Else
    Private Declare Function WinsockRCmd Lib "RCMD.DLL" (ByVal RHost As String, ByVal RPort As Integer, ByVal LocalUser As String, ByVal RemoteUser As String, ByVal Cmd As String, ByVal ErrorMsg As String, ByVal ErrLen As Integer) As Integer
    Private Declare Function RCmdRead Lib "RCMD.DLL" (ByVal hRCmd As Integer, ByVal RData As String, ByVal RCount As Integer) As Integer
    Private Declare Function RCmdReadByte Lib "RCMD.DLL" (ByVal hRCmd As Integer) As Integer
    Private Declare Function RCmdClose Lib "RCMD.DLL" (ByVal hRCmd As Integer) As Integer
    Private Declare Function RCmdHandle Lib "RCMD.DLL" (ByVal hRCmd As Integer) As Integer
    Private Declare Function RCmdCancel Lib "RCMD.DLL" () As Integer
    Private Declare Function RCmdNoEvents Lib "RCMD.DLL" () As Integer
    Private Declare Function RCmdEvents Lib "RCMD.DLL" () As Integer
#End If

Private Function fnCStr(vTemp As Variant) As String
    Dim nPos As Integer
    
    If IsNull(vTemp) Then
        fnCStr = ""
    Else
        fnCStr = Trim(vTemp)
        Do
            nPos = InStr(fnCStr, Chr(0))
            If nPos > 0 Then
                Mid(fnCStr, nPos, 1) = " "
            End If
        Loop Until nPos = 0
        fnCStr = Trim(fnCStr)
    End If

End Function

Private Function fnDBPath() As String
    Dim sDBPath As String
    Dim sStatus As String
    Dim i As Integer
    
    sDBPath = tfnGetNamedString(t_dbMainDatabase.Connect, CONNECT_DBPATH1)
    If Trim(sDBPath) = "" Then
        sDBPath = tfnGetNamedString(t_dbMainDatabase.Connect, CONNECT_DBPATH2)
    End If
    i = Len(sDBPath)
    sStatus = " "
    While i > 0 And sStatus <> "/"
        sStatus = Mid(sDBPath, i, 1)
        i = i - 1
    Wend
    If i > 0 Then
        fnDBPath = Left(sDBPath, i)
    Else
        fnDBPath = sDBPath
    End If

End Function

Public Function fnExecute4GE(sCmdLine As String, _
                             Optional vEnviron As Variant) As Boolean
    Const SUB_NAME = "fnExecute4GE"
    
    Dim sHost As String
    Dim sUserID As String
    Dim sPassWD As String
    Dim sDBPath As String
    Dim nCode As Integer
    Dim sCmd As String
    Dim sTemp As String
    Dim sEnviron As String
    
    If t_dbMainDatabase Is Nothing Then
        #If DEVELOP Then
            sHost = "ether5"
            sUserID = "ssfactor"
            sPassWD = "menus"
            sDBPath = "/factor/retail"
        #Else
            fnExecute4GE = False
            Exit Function
        #End If
    Else
        sHost = tfnGetNamedString(t_dbMainDatabase.Connect, CONNECT_HOST)
        sDBPath = fnDBPath
        #If FACTOR_MENU Then
            sUserID = tfnGetNamedString(t_dbMainDatabase.Connect, CONNECT_USERID)
            sPassWD = tfnGetNamedString(t_dbMainDatabase.Connect, CONNECT_PSWD)
        #Else
            sUserID = "ssfactor"
            sPassWD = "menus"
        #End If
    End If
    If Trim(sHost) = "" Then
        If Not t_oleObject Is Nothing Then
            sHost = t_oleObject.ConnectHost
        End If
    End If
    
    If nWhatToUse = USE_RCMD Then
        If IsMissing(vEnviron) Then
            sEnviron = ""
        Else
            sEnviron = Trim(vEnviron)
            If Right(sEnviron, 1) <> ";" Then
                sEnviron = sEnviron & ";"
            End If
        End If
        sCmd = fnVariables(sHost) & "DBPATH=" & sDBPath & ":$PROGPATH; export DBPATH;" & sEnviron & "cd $HOME;" _
             & "$PROGPATH/" & sCmdLine
        
        sTemp = fnRunRCmd(sHost, sUserID, sPassWD, sCmd)
        If sTemp = "" Then
            fnExecute4GE = True
        Else
            fnExecute4GE = False
            tfnErrHandler SUB_NAME, ERR_MSG_RUN4GE, "Failed to execute 4ge program" & vbCrLf & sTemp & vbCrLf & "Command line string:" & sCmd
        End If
    Else
        sCmd = "DBPATH=" & sDBPath & ":$PROGPATH; export DBPATH;cd " & sDBPath & ";" _
             & "$PROGPATH/" & sCmdLine
        
        Dim strSQL As String
        Dim rsTemp As Recordset
        
        strSQL = "EXECUTE PROCEDURE execute_unix_cmd (" & tfnSQLString(sCmd) & ")"
        On Error GoTo errExecuteProcedure
        Set rsTemp = t_dbMainDatabase.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
        If rsTemp.RecordCount > 0 Then
            If tfnRound(rsTemp.Fields(0)) = 0 Then
                sTemp = fnCStr(rsTemp.Fields(1))
                If rsTemp.Fields.Count > 2 Then
                    sTemp = sTemp & vbCrLf & "System command: " & fnCStr(rsTemp.Fields(2))
                End If
                tfnErrHandler SUB_NAME, -1, sTemp
            Else
                fnExecute4GE = True
            End If
        End If
    End If
    Exit Function
errExecuteProcedure:
    tfnErrHandler SUB_NAME, RUN_TIME_PROC, strSQL
End Function

Private Function fnParmIndex(vTemp As Variant) As Integer

    Dim sTemp As String
    
    fnParmIndex = 0
    If Not IsNull(vTemp) Then
        If TypeOf vTemp Is Form Then
            On Error Resume Next
            sTemp = UCase(Trim(vTemp.efraToolBar.FMName))
        ElseIf VarType(vTemp) = vbString Then
            sTemp = UCase(fnCStr(vTemp))
        Else 'Assume integer
            fnParmIndex = Val(vTemp)
            Exit Function
        End If
        Select Case sTemp
            Case FLAG_CAP_APVPRINT
                fnParmIndex = FLAG_IDX_APVPRINT
            Case FLAG_CAP_APVOID
                fnParmIndex = FLAG_IDX_APVOID
            Case FLAG_CAP_PRVPRINT
                fnParmIndex = FLAG_IDX_PRVPRINT
            Case FLAG_CAP_OEEINVCE
                fnParmIndex = FLAG_IDX_OEEINVCE
            Case FLAG_CAP_FMESTORE
                fnParmIndex = FLAG_IDX_FMESTORE
        End Select
    End If
    
End Function

'This function is modified by Weigong on 11/25/98
'The changes are: Add one more parameter (the 1st one)
'sCDFlag -- "C" or "D" (originally it is always "P")
'Now it can print both checks (when they put "C" in)
'and Direct Deposit Slip (when they put "D" in)
Public Function fnPRPrintCheck(sCDFlag As String, _
                               sPrinter As String, _
                               lStart As Long, _
                               sCheckDate As String, _
                               sEffcDate As String, _
                               sPrintGroup As String, _
                               nSortBy As Integer) As Boolean
    Dim sHost As String
    Dim sUserID As String
    Dim sPassWD As String
    Dim sDBPath As String
    Dim nCode As Integer
    Dim sCmd As String
    
    #If DEVELOP Then
        sHost = "ether5"
        sUserID = "ssfactor"
        sPassWD = "menus"
        sDBPath = "/factor/retail"
    #Else
        sHost = tfnGetNamedString(t_dbMainDatabase.Connect, CONNECT_HOST)
        sUserID = tfnGetNamedString(t_dbMainDatabase.Connect, CONNECT_USERID)
        sPassWD = tfnGetNamedString(t_dbMainDatabase.Connect, CONNECT_PSWD)
        sDBPath = fnDBPath
    #End If
    
    sCmd = "prvprint.4ge " & sPrinter & " " & sUserID & " " & sCDFlag & " " & CStr(lStart) & " " _
         & sCheckDate & " " & sEffcDate & " " & sPrintGroup & " " & CStr(nSortBy)
    
    fnPRPrintCheck = fnExecute4GE(sCmd)

End Function

Private Function fnRunRCmd(sHost As String, _
                           sLocalUID As String, _
                           sRemoteUID As String, _
                           sCmd As String) As String
    Const SUB_NAME = "fnRunRCmd"
    
    Dim sErrMsg As String
    Dim nCode As Integer
    Dim nMsgLen As Integer
    Dim nOutput As Integer
    
    On Error GoTo errRunShell
    
    fnRunRCmd = ""
    nCode = ERR_LOGIN
    sErrMsg = Space(MAX_MSG_LEN + 1)
    
'    MsgBox "Logon: Host = " & sHost & vbCrLf _
'           & "UID = " & sLocalUID & vbCrLf _
'           & "PWD = " & sRemoteUID & vbCrLf _
'           & "Command line = " & sCmd

    nCode = WinsockRCmd(sHost, WINSOCK_PORT, sLocalUID, sRemoteUID, sCmd, sErrMsg, MAX_MSG_LEN)
    
    If nCode < 0 Then
        nMsgLen = InStr(sErrMsg, Chr(0))
        If nMsgLen > 0 Then
            fnRunRCmd = Left(sErrMsg, nMsgLen)
'            MsgBox "Failed here: " & fnRunRCmd
        Else
            fnRunRCmd = "Cannot logon to the server to execute server program"
        End If
    Else
        sErrMsg = ""
        Do
            nOutput = RCmdReadByte(nCode)
            If nOutput > 1 Then
                sErrMsg = sErrMsg & Chr(nOutput)
            End If
            DoEvents
        Loop Until nOutput <= 1
        RCmdClose nCode
        If sErrMsg <> "" Then
            tfnErrHandler SUB_NAME, ERR_MSG_RETURNED, "A message has been returned from the server:" & vbCrLf & sErrMsg & vbCrLf & vbCrLf & "Command sent to server '" & sHost & "' by user '" & sLocalUID & "':" & vbCrLf & sCmd
        End If
    End If
    
    Exit Function
    
errRunShell:
    If Err.Number = 48 Then
        tfnErrHandler SUB_NAME, ERR_RCMD_MISSING, "Cannot find file 'RCMD32.DLL'"
    Else
        tfnErrHandler SUB_NAME, RUN_TIME_RCMD, Err.Description
    End If
End Function

Public Function fnPRTestPrint(sPrinter As String, _
                              lStart As Long, _
                              sDate As String) As Boolean
    Dim sHost As String
    Dim sUserID As String
    Dim sPassWD As String
    Dim sDBPath As String
    Dim nCode As Integer
    Dim sCmd As String
    
    #If DEVELOP Then
        sHost = "ether5"
        sUserID = "ssfactor"
        sPassWD = "menus"
        sDBPath = "/factor/retail"
    #Else
        sHost = tfnGetNamedString(t_dbMainDatabase.Connect, CONNECT_HOST)
        sUserID = tfnGetNamedString(t_dbMainDatabase.Connect, CONNECT_USERID)
        sPassWD = tfnGetNamedString(t_dbMainDatabase.Connect, CONNECT_PSWD)
        sDBPath = fnDBPath
    #End If
    
    sCmd = "prvprint.4ge " & sPrinter & " " & sUserID & " T " & CStr(lStart) & " " & sDate

    fnPRTestPrint = fnExecute4GE(sCmd)

End Function


Private Function fnVariables(sHost As String) As String
    Dim sTemp As String
    sTemp = "DBPRINT=/usr/factor/isqlflt; export DBPRINT;"
    sTemp = sTemp & "DBTEMP=/usr/tmp; export DBTEMP;"
    sTemp = sTemp & "INFORMIXDIR=/usr/informix; export INFORMIXDIR;"
    sTemp = sTemp & "INFORMIXSERVER=" & sHost & "; export INFORMIXSERVER;"
    sTemp = sTemp & "PROGPATH=/usr/factor; export PROGPATH;"
    sTemp = sTemp & "SQLEXEC=/usr/informix/lib/sqlrm;export SQLEXEC;"
    'Took out because these may cause problem on different systems. Ma. 2/5/99
    sTemp = sTemp & "TERMCAP=/usr/informix/etc/Termcap;export TERMCAP;"
    sTemp = sTemp & "PATH=/bin:/usr/bin::/usr/informix/breakaway:/usr/informix/bin:/usr/factor; export PATH;"
    fnVariables = sTemp
End Function

Public Function fnSetParmForUnixCmd(vFlag As Variant, _
                                    Optional vDefault As Variant) As Boolean

    Const SUB_NAME = "fnSetParmForUnixCmd"
    
    Dim sTemp As String
    Dim rsTemp As Recordset
    Dim nParmIdx As Integer
    
    fnSetParmForUnixCmd = False
    If IsMissing(vDefault) Then
        nWhatToUse = USE_STORED_PROC
    Else
        nWhatToUse = Val(vDefault)
    End If
    nParmIdx = fnParmIndex(vFlag)
    If nParmIdx > 0 And nParmIdx <= FLAG_COUNT Then
        sTemp = "SELECT parm_field FROM sys_parm" _
               & " WHERE parm_nbr = " & PARM_RUN_4GE
    
        On Error GoTo errGetParm
        Set rsTemp = t_dbMainDatabase.OpenRecordset(sTemp, dbOpenSnapshot, dbSQLPassThrough)
        If rsTemp.RecordCount > 0 Then
            sTemp = fnCStr(rsTemp!parm_field)
            If nParmIdx <= Len(sTemp) Then
                If UCase(Mid(sTemp, nParmIdx, 1)) = "D" Then
                    nWhatToUse = USE_RCMD
                End If
            End If
        End If
    End If
    fnSetParmForUnixCmd = True
    Exit Function
errGetParm:
    tfnErrHandler SUB_NAME, sTemp, False
    tfnErrHandler SUB_NAME, -1, "Read database failed. Stored procedure will be used to run 4ge programs", vbExclamation + vbOK
End Function

Public Function ExecUnixCmd(sHost As String, _
                           sLocalUID As String, _
                           sRemoteUID As String, _
                           sCmd As String) As String

    Const SUB_NAME = "ExecUnixCmd"
    
    Dim sErrMsg As String
    Dim nCode As Integer
    Dim nMsgLen As Integer
    Dim nOutput As Integer
    
    On Error GoTo errRunShell
    
    ExecUnixCmd = ""
    nCode = ERR_LOGIN
    sErrMsg = Space(MAX_MSG_LEN + 1)
    
    nCode = WinsockRCmd(sHost, WINSOCK_PORT, sLocalUID, sRemoteUID, sCmd, sErrMsg, MAX_MSG_LEN)
    
    If nCode < 0 Then
        nMsgLen = InStr(sErrMsg, Chr(0))
        If nMsgLen > 0 Then
            ExecUnixCmd = Left(sErrMsg, nMsgLen)
'            MsgBox "Failed here: " & fnRunRCmd
        Else
            ExecUnixCmd = "Cannot logon to the server to execute server program"
        End If
    Else
        sErrMsg = ""
        Do
            nOutput = RCmdReadByte(nCode)
            If nOutput > 0 Then
                sErrMsg = sErrMsg & Chr(nOutput)
            End If
            DoEvents
        Loop Until nOutput = 0
        RCmdClose nCode
        ExecUnixCmd = sErrMsg
    End If
    
    Exit Function
    
errRunShell:
    If Err.Number = 48 Then
        tfnErrHandler SUB_NAME, ERR_RCMD_MISSING, "Cannot find file 'RCMD32.DLL'"
    Else
        tfnErrHandler SUB_NAME, RUN_TIME_RCMD, Err.Description
    End If
    
    
End Function


