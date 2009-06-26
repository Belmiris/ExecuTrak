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
'david 10/23/00
Private Const CONNECT_HOST2 = ";SRVR"
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

Private Const SEC_SETUP_4GE = "SETUP OF 4GL PROGRAMS"
Private Const KEY_PROGPATH_4GE = "PROG PATH"
Private Const DEFAULT_PROGPATH_4GE = "/usr/factor"
    
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
                             Optional vEnviron As Variant, _
                             Optional bShowMsgBox As Boolean = True, _
                             Optional sErrMsg As String = "") As Boolean
    
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
'MsgBox "t_dbMainDatabase.Connect=" + tfnSQLString(t_dbMainDatabase.Connect)
        'sHost = tfnGetNamedString(t_dbMainDatabase.Connect, CONNECT_HOST)
        sHost = tfnGetHostName()
        sDBPath = fnDBPath()
        
        'david 11/16/00
        'sUserID = tfnGetNamedString(t_dbMainDatabase.Connect, CONNECT_USERID)
        'sPassWD = tfnGetNamedString(t_dbMainDatabase.Connect, CONNECT_PSWD)
        sUserID = tfnGetUserName()
        sPassWD = tfnGetPassword()
    End If
    
    'david 10/23/00
    If Trim(sHost) = "" Then
        If Not t_dbMainDatabase Is Nothing Then
            sHost = tfnGetNamedString(t_dbMainDatabase.Connect, CONNECT_HOST2)
        End If
    End If
    
    If Trim(sHost) = "" Then
        If Not t_oleObject Is Nothing Then
            sHost = t_oleObject.ConnectHost
        End If
    End If
    
    If Trim(sPassWD) = "" Then
        If Not t_oleObject Is Nothing Then
            sPassWD = t_oleObject.Password
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
        sCmd = fnVariables(sHost, sDBPath) & sEnviron & "cd /home/" & sUserID & ";" & "$PROGPATH/" & sCmdLine
        
        sTemp = tfnRunRCmd(sHost, sUserID, sPassWD, sCmd)
        If sTemp = "" Then
            fnExecute4GE = True
        Else
            fnExecute4GE = False
            'Vijaya on 10/25/02 Magic#387222 we need to show the password as **** characters.
            sTemp = Replace(sTemp, "PWD=" & tfnGetNamedString(sTemp, "PWD"), "PWD=" & String(Len(tfnGetNamedString(sTemp, "PWD")), "*"))
            'end of Vijaya Code
            'david 01/18/2002
            If bShowMsgBox Then
                tfnErrHandler SUB_NAME, ERR_MSG_RUN4GE, sTemp
            Else
                tfnErrHandler SUB_NAME, ERR_MSG_RUN4GE, sTemp, False
            End If
            
            sErrMsg = ERR_MSG_RUN4GE & " - " & sTemp
            '''''''''''''''''''
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
                
                'david 01/18/2002
                If bShowMsgBox Then
                    tfnErrHandler SUB_NAME, -1, sTemp
                Else
                    tfnErrHandler SUB_NAME, -1, sTemp, False
                End If
                
                sErrMsg = sTemp
                '''''''''''''''''''
            Else
                fnExecute4GE = True
            End If
        End If
    End If
    Exit Function
errExecuteProcedure:
    'david 01/18/2002
    If bShowMsgBox Then
        tfnErrHandler SUB_NAME, RUN_TIME_PROC, strSQL
    Else
        tfnErrHandler SUB_NAME, RUN_TIME_PROC, strSQL, False
    End If
    
    sErrMsg = RUN_TIME_PROC & " - " & strSQL
    '''''''''''''''''''
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
            fnParmIndex = val(vTemp)
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
    
    'david 08/30/2002  #369533
    'Added for input pay period or other message which will show on the 4th line of pay stub
    If sCDFlag <> "U" Then
        Dim sMsg As String
        Dim sTmp As String
        Dim i As Integer
        Dim bDone As Boolean
        
        '##Provide a message box for user input the pay period or other message
        bDone = False
        
        Do While Not bDone
            sMsg = InputBox("Optional - You may enter the pay period or any other message, up to 40 characters," _
                + " that will be printed on the Check Stub: ", "Message for Check Stub", sMsg)
            sMsg = Trim(sMsg)
            
            bDone = True
            
            If Len(sMsg) > 40 Then
                MsgBox "The length of the message cannot be over 40 characters.", vbExclamation
                bDone = False
            ElseIf Len(sMsg) > 0 Then
                For i = 1 To Len(sMsg)
                    sTmp = Mid(sMsg, i, 1)
                    
                    If sTmp = Chr(10) Or sTmp = Chr(13) Or sTmp = Chr(34) Or sTmp = Chr(39) Then
                        If MsgBox("Single quote ('), double quote ("") or 'Line Feed' are not allowed in the message," _
                           + " and will be replaced with SPACE. Are you sure you want to continue?" + vbCrLf + vbCrLf _
                           + "Choose No to re-enter the message.", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                            'go back to enter message again
                            bDone = False
                        End If
                    
                        Exit For
                    End If
                Next
            End If
        Loop
        
        '##replace quote marks and carriage return with space in the message
        sTmp = Chr(13)  'vbCr
        sMsg = Replace(sMsg, sTmp, " ")
        sTmp = Chr(10)  'vbLf
        sMsg = Replace(sMsg, sTmp, " ")
        sTmp = Chr(34)  'double quote
        sMsg = Replace(sMsg, sTmp, " ")
        sTmp = Chr(39)  'single quote
        sMsg = Replace(sMsg, sTmp, " ")
    End If
    'david 08/30/2002  #369533
    '##if sMsg is not empty, add it to the last arguments when passing to prvprint.4ge
    If sMsg <> "" Then
        sCmd = "prvprint.4ge " & sPrinter & " " & sUserID & " " & sCDFlag & " " & CStr(lStart) & " " _
             & sCheckDate & " " & sEffcDate & " " & sPrintGroup & " " & CStr(nSortBy) & " " & "'" & sMsg & "'"
    Else
        sCmd = "prvprint.4ge " & sPrinter & " " & sUserID & " " & sCDFlag & " " & CStr(lStart) & " " _
             & sCheckDate & " " & sEffcDate & " " & sPrintGroup & " " & CStr(nSortBy)
    End If
    
    fnPRPrintCheck = fnExecute4GE(sCmd)
End Function

'david 03/17/2004
'make doevents optional, and default to no doevents
'because it will mess up the keystroke in the grid
Public Function tfnRunRCmd(sHost As String, _
                           sLocalUID As String, _
                           sRemoteUID As String, _
                           sCmd As String, _
                           Optional rtn_ori_str As Boolean = False, _
                           Optional bDoEvents As Boolean = False) As String
    
    Const SUB_NAME = "tfnRunRCmd"
    
    'maximum time to execute a unix command
    Const MAX_CALL_TIME As Long = 1800  '3 hours
    
    Dim sErrMsg As String
    Dim nCode As Integer
    Dim nMsgLen As Integer
    Dim nOutput As Integer
    Dim sConnect_Used As String
    
    On Error GoTo errRunShell
    
    If Not t_dbMainDatabase Is Nothing Then
        sConnect_Used = t_dbMainDatabase.Connect
    End If
    
    tfnRunRCmd = ""
    nCode = ERR_LOGIN
    sErrMsg = Space(MAX_MSG_LEN + 1)
    
'MsgBox "Before calling WinsockRCmd()" + vbCrLf _
    + "sHost=" & tfnSQLString(sHost) + vbCrLf _
    + "WINSOCK_PORT=" & tfnSQLString(WINSOCK_PORT) + vbCrLf _
    + "sLocalUID=" & tfnSQLString(sLocalUID) + vbCrLf _
    + "sRemoteUID=" & tfnSQLString(sRemoteUID) + vbCrLf _
    + "sCmd=" & tfnSQLString(sCmd)
    
    nCode = WinsockRCmd(sHost, WINSOCK_PORT, sLocalUID, sRemoteUID, sCmd, sErrMsg, MAX_MSG_LEN)
    
    If nCode < 0 Then
        nMsgLen = InStr(sErrMsg, Chr(0))
              
        If nMsgLen > 0 Then
            tfnRunRCmd = tfnStripNULL(Left(sErrMsg, nMsgLen)) & vbCrLf & "Connection String: " & sConnect_Used
        Else
            tfnRunRCmd = "Cannot logon to the server to execute server program" & vbCrLf & "Connection String: " & sConnect_Used
        End If
    Else
        'david 03/17/2004
        'put a time out here
        Dim sngTimer As Single
        sngTimer = Timer
        
        sErrMsg = ""
        Do
            nOutput = RCmdReadByte(nCode)
            If nOutput > 1 Then
                sErrMsg = sErrMsg & Chr(nOutput)
            End If
            
            If bDoEvents Then
                DoEvents
            End If
            
            If Timer - sngTimer > MAX_CALL_TIME Then
                If MsgBox("RCMD time out. Do you want to continue to wait (30 minutes)" _
                   + " for the program to finished?", vbQuestion + vbYesNo) = vbYes Then
                    sngTimer = Timer
                Else
                    tfnRunRCmd = "RCMD time out"
                    Exit Function
                End If
            End If
        Loop Until nOutput <= 1
        ''''''''''''''''''''
        
        RCmdClose nCode
        
        If sErrMsg <> "" Then
            If Not rtn_ori_str Then
                sErrMsg = "A message has been returned from the server:" & vbCrLf & sErrMsg & vbCrLf & vbCrLf & "Command sent to server '" & sHost & "' by user '" & sLocalUID & "':" & vbCrLf & sCmd
            End If
            tfnRunRCmd = sErrMsg
        End If
    End If
    
    Exit Function
    
errRunShell:
    If Err.number = 48 Then
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


Private Function fnVariables(sHost As String, sDBPath As String) As String
    Dim sTemp As String
    
    If InStr(sDBPath, "/") > 0 Then
        'standard server
        sTemp = "DBPRINT=/usr/factor/isqlflt; export DBPRINT;"
        sTemp = sTemp + "DBTEMP=/usr/tmp; export DBTEMP;"
        sTemp = sTemp + "INFORMIXDIR=/usr/informix; export INFORMIXDIR;"
        sTemp = sTemp + "INFORMIXSERVER=" + sHost + "; export INFORMIXSERVER;"
        sTemp = sTemp + "PROGPATH=" + fnGetProgPath + "; export PROGPATH;"
        sTemp = sTemp + "SQLEXEC=/usr/informix/lib/sqlrm;export SQLEXEC;"
        sTemp = sTemp + "TERMCAP=/usr/informix/etc/Termcap;export TERMCAP;"
        sTemp = sTemp + "PATH=/bin:/usr/bin::/usr/informix/breakaway:/usr/informix/bin:/usr/factor; export PATH;"
        sTemp = sTemp + "DBPATH=" + sDBPath + ":$PROGPATH; export DBPATH;"
    Else
        'ids - dynamic server
        sTemp = ""
        'sTemp = sTemp + "DBPRINT=/usr/factor/isqlflt; export DBPRINT;"
        'sTemp = sTemp + "DBTEMP=/usr/tmp; export DBTEMP;"
        sTemp = sTemp + "INFORMIXDIR=/usr/ids; export INFORMIXDIR;"
        sTemp = sTemp + "ONCONFIG=onconfig; export ONCONFIG;"
        sTemp = sTemp + "PATH=/bin:/usr/bin:$INFORMIXDIR/bin:/usr/cmds:/usr/factor; export PATH;"
        sTemp = sTemp + "INFORMIXSERVER=" + sHost + "; export INFORMIXSERVER;"
        sTemp = sTemp + "INFORMIXSQLHOSTS=$INFORMIXDIR/etc/sqlhosts; export INFORMIXSQLHOSTS;"
        sTemp = sTemp + "LIBPATH=$INFORMIXDIR/lib:$LIBPATH; export LIBPATH;"
        sTemp = sTemp + "PROGPATH=" + fnGetProgPath + "; export PROGPATH;"
        sTemp = sTemp + "DATADIR=/factor; export DATADIR;"
        sTemp = sTemp + "TERMCAP=/usr/informix/etc/Termcap;export TERMCAP;"
        sTemp = sTemp + "DBPATH=$DATADIR:$PROGPATH; export DBPATH;"
        sTemp = sTemp + "DBNAME=" + sDBPath + "; export DBNAME;"
    End If
    
    fnVariables = sTemp
End Function

Private Function fnDefaultParm(sSECTION As String, _
                              sKey As String, _
                              sDefault As String) As String
    Dim sIniFileName As String
    
    Dim nLength As Long 'length of the value returned for api call
    Dim sBuffer As String
    Dim bStatus As Boolean 'status returned from api call

    sIniFileName = tfnGetWindowsDir(True) & szFACTOR_INI
    
    sBuffer = Space(MAX_STRING_LENGTH) 'clear and make the string fixed length
    
    'get the [value] for the [section], [key], and ini file sent
    nLength = GetPrivateProfileString(sSECTION, sKey, szEMPTY, sBuffer, MAX_STRING_LENGTH, sIniFileName)
    
    If nLength <> 0 Then 'if length positive [value] has been found
        fnDefaultParm = Left(sBuffer, nLength) 'make it a basic string
    Else
        'write the [value] for the [section], [key], and ini file sent
        WritePrivateProfileString sSECTION, sKey, sDefault, sIniFileName
        fnDefaultParm = sDefault
    End If

End Function

Private Function fnGetProgPath() As String
    fnGetProgPath = Trim(fnDefaultParm(SEC_SETUP_4GE, KEY_PROGPATH_4GE, DEFAULT_PROGPATH_4GE))
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
        nWhatToUse = val(vDefault)
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
    
    fnGetProgPath ' WJ 01/09/01, this will set the program path in the factor.ini
    
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
    If Err.number = 48 Then
        tfnErrHandler SUB_NAME, ERR_RCMD_MISSING, "Cannot find file 'RCMD32.DLL'"
    Else
        tfnErrHandler SUB_NAME, RUN_TIME_RCMD, Err.Description
    End If
    
    
End Function

Public Function fnRun4GLPricing(sCmdLine As String, _
                                    Optional bAlwaysLunch As Boolean = False) As String
    
    Dim sHost As String
    Dim sUserID As String
    Dim sPassWD As String
    Dim sDBPath As String
    Dim nCode As Integer
    Dim sCmd As String
    Dim sTemp As String
    Dim sEnviron As String
    Static staCmd As String
    Static staRtn As String
    
    'Vijaya on 07/09/03 needs always lunch the 4ge by default it will check
    
    If staCmd = sCmdLine And Not bAlwaysLunch Then
        fnRun4GLPricing = staRtn
        Exit Function
    Else
        staCmd = sCmdLine
    End If
    
    If Not t_dbMainDatabase Is Nothing Then
        sHost = tfnGetHostName
        sDBPath = fnDBPath
        sUserID = tfnGetUserName
        sPassWD = tfnGetPassword
    End If
    If Trim(sHost) = "" Then
        If Not t_dbMainDatabase Is Nothing Then
            sHost = tfnGetNamedString(t_dbMainDatabase.Connect, CONNECT_HOST2)
        End If
    End If
    
    If Trim(sHost) = "" Then
        If Not t_oleObject Is Nothing Then
            sHost = t_oleObject.ConnectHost
        End If
    End If
    
    If Trim(sPassWD) = "" Then
        If Not t_oleObject Is Nothing Then
            sPassWD = t_oleObject.Password
        End If
    End If
    
    sCmd = fnVariables(sHost, sDBPath) & sEnviron & "cd /home/" & sUserID & ";" & "$PROGPATH/" & sCmdLine
    staRtn = tfnRunRCmd(sHost, sUserID, sPassWD, sCmd, True)
    fnRun4GLPricing = staRtn
End Function

Public Function tfnRemoteFileExists(sFilename As String, ByRef sErrMsg As String) As Boolean
    sErrMsg = tfnRunRCmd(tfnGetHostName(), tfnGetUserName(), tfnGetPassword(), "ls " + sFilename & " > /dev/null")
    
    If sErrMsg = "" Then
        'file found
        sErrMsg = ""
        tfnRemoteFileExists = True
    ElseIf InStr(sErrMsg, "ls:") > 0 Then
        'file not found
        sErrMsg = ""
        tfnRemoteFileExists = False
    Else
        'error executing ls command
    End If
End Function
