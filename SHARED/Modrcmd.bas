Attribute VB_Name = "modRCmd"
' 8829 - tthompson

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

'**********************************************************
'* USED TO CALL PLINK - 8696
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Const SYNCHRONIZE = &H100000
Const INFINITE = &HFFFFFFFF

Private Declare Function OpenProcess Lib "kernel32" _
    (ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As Long
    
Private Declare Function WaitForSingleObject Lib "kernel32" _
    (ByVal hHandle As Long, _
    ByVal dwMilliseconds As Long) As Long
    
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Function CoCreateGuid Lib "OLE32.DLL" (pGuid As GUID) As Long

Public Enum RemoteApps
    rapPlink = 0
    rapPscp = 1
    rapPsftp = 2
End Enum

Private m_sSSH_KEY As String
Private m_bSSH_KEY As Boolean
Private m_DontDeleteRemoteFiles As Boolean
Private m_sLastHost As String
'**********************************************************

'
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
        fnExecute4GE = False
        Exit Function
    End If
    
    sHost = tfnGetHostName()
    sDBPath = fnDBPath()
    sUserID = tfnGetUserName()
    sPassWD = tfnGetPassword()
    
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
            sPassWD = t_oleObject.password
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
        
        sCmd = fnVariables(sHost, sDBPath, sUserID, sPassWD) & sEnviron & "cd /home/" & sUserID & ";" & "$PROGPATH/" & sCmdLine
        
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
                
                If rsTemp.Fields.count > 2 Then
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
    
    sHost = tfnGetNamedString(t_dbMainDatabase.Connect, CONNECT_HOST)
    sUserID = tfnGetNamedString(t_dbMainDatabase.Connect, CONNECT_USERID)
    sPassWD = tfnGetNamedString(t_dbMainDatabase.Connect, CONNECT_PSWD)
    sDBPath = fnDBPath
    
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
    Const MAX_CALL_TIME As Long = 10800  '3 hours
    
    Dim sErrMsg As String
    Dim nCode As Integer
    Dim nMsgLen As Integer
    Dim nOutput As Integer
    Dim sConnect_Used As String
    Dim sHostKey As String          '8696
    Dim sPLinkResult As String      '8696
    Dim sPLinkError As String       '8696
    
    On Error GoTo errRunShell
    
    If Not t_dbMainDatabase Is Nothing Then
        sConnect_Used = t_dbMainDatabase.Connect
    End If
    
    tfnRunRCmd = ""
    nCode = ERR_LOGIN
    sErrMsg = Space(MAX_MSG_LEN + 1)
    
    sHostKey = Trim(fnGetServerHostKey(sHost))
    If sHostKey <> "" Then GoTo RUN_PLINK
    
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
    
RUN_PLINK:
    If tfnRunRemoteCmd(sHostKey, _
                       rapPlink, _
                       sCmd, _
                       sPLinkResult, _
                       sPLinkError, _
                       sHost, _
                       sLocalUID, _
                       sRemoteUID, _
                       MAX_CALL_TIME, _
                       True) Then
        tfnRunRCmd = sPLinkResult
    Else
        If Not rtn_ori_str Then
            sErrMsg = "A message has been returned from the server:" & vbCrLf & sPLinkError & vbCrLf & vbCrLf & "Command sent to server '" & sHost & "' by user '" & sLocalUID & "':" & vbCrLf & sCmd
        Else
            sErrMsg = sPLinkError
        End If
        tfnRunRCmd = sErrMsg
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

Private Function fnVariables(sHost As String, sDBPath As String, sUserID As String, sPassWD As String) As String
    Static bTestServerScript As Boolean
    Static bUseServerScript As Boolean
    Dim sCmd As String
    Dim sTemp As String
    
        If Not bTestServerScript Then
            sCmd = fnVariables_UseScript(sHost, sDBPath)
            sTemp = tfnRunRCmd(sHost, sUserID, sPassWD, sCmd)
            If Trim$(sTemp) = "" Then
                bUseServerScript = True
            End If
            bTestServerScript = True
        End If
        
    
    If bUseServerScript Then
        fnVariables = fnVariables_UseScript(sHost, sDBPath)
    Else
        fnVariables = fnVariables_UseCommands(sHost, sDBPath)
    End If
End Function

Private Function fnVariables_UseScript(sHost As String, sDBPath As String) As String
    Dim sTemp As String
    
    If InStr(sDBPath, "/") > 0 Then
        'standard server
        sTemp = ". " & fnGetProgPath & "/syesevar.sh " & sDBPath & " " & fnGetProgPath & ";"
    Else
        'ids - dynamic server
        sTemp = ". " & fnGetProgPath & "/syedsvar.sh " & sDBPath & " " & fnGetProgPath & ";"
    End If
    
    fnVariables_UseScript = sTemp
End Function

Private Function fnVariables_UseCommands(sHost As String, sDBPath As String) As String
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
        sTemp = sTemp + "LD_LIBRARY_PATH='/usr/informix/lib:/usr/informix/lib/esql:/usr/informix/lib/tools'; export LD_LIBRARY_PATH;"
    
    Else
        'ids - dynamic server
        sTemp = ""
        'sTemp = sTemp + "DBPRINT=/usr/factor/isqlflt; export DBPRINT;"
        'sTemp = sTemp + "DBTEMP=/usr/tmp; export DBTEMP;"
        sTemp = sTemp + "INFORMIXDIR=/usr/informix; export INFORMIXDIR;"
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
        sTemp = sTemp + "DBNAME=" + sDBPath + "; export DBNAME;"
    End If
    
    fnVariables_UseCommands = sTemp
End Function
Private Function fnDefaultParm(sSection As String, _
                              sKey As String, _
                              sDefault As String) As String
    Dim sIniFileName As String
    
    Dim nLength As Long 'length of the value returned for api call
    Dim sBuffer As String
    Dim bStatus As Boolean 'status returned from api call

    sIniFileName = tfnGetWindowsDir(True) & szFACTOR_INI
    
    sBuffer = Space(MAX_STRING_LENGTH) 'clear and make the string fixed length
    
    'get the [value] for the [section], [key], and ini file sent
    nLength = GetPrivateProfileString(sSection, sKey, szEMPTY, sBuffer, MAX_STRING_LENGTH, sIniFileName)
    
    If nLength <> 0 Then 'if length positive [value] has been found
        fnDefaultParm = Left(sBuffer, nLength) 'make it a basic string
    Else
        'write the [value] for the [section], [key], and ini file sent
        WritePrivateProfileString sSection, sKey, sDefault, sIniFileName
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
    Dim sHostKey As String          '8696
    Dim sPLinkResult As String      '8696
    Dim sPLinkError As String       '8696
    
    On Error GoTo errRunShell
    
    ExecUnixCmd = ""
    nCode = ERR_LOGIN
    sErrMsg = Space(MAX_MSG_LEN + 1)
    
    sHostKey = Trim(fnGetServerHostKey())
    If sHostKey <> "" Then GoTo RUN_PLINK
    
    nCode = WinsockRCmd(sHost, WINSOCK_PORT, sLocalUID, sRemoteUID, sCmd, sErrMsg, MAX_MSG_LEN)
    
    If nCode < 0 Then
        nMsgLen = InStr(sErrMsg, Chr(0))
        If nMsgLen > 0 Then
            ExecUnixCmd = Left(sErrMsg, nMsgLen)
            'MsgBox "Failed here: " & fnRunRCmd
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
    
RUN_PLINK:
    '8696 - chmod 777 /factor/fuel_mvmt_upld/ht090821105135.sav
    If tfnRunRemoteCmd(sHostKey, rapPlink, sCmd, sPLinkResult, sPLinkError, sHost, sLocalUID) Then
        ExecUnixCmd = sPLinkResult
    Else
        ExecUnixCmd = sPLinkError & " " & sPLinkResult
    End If
    
    Exit Function
    
errRunShell:
    If Err.Number = 48 Then
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
    Dim sHostKey As String
    Dim sPLinkResult As String
    Dim sPLinkError As String
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
            sPassWD = t_oleObject.password
        End If
    End If
    
    sHostKey = Trim(fnGetServerHostKey())
    
    sCmd = fnVariables(sHost, sDBPath, sUserID, sPassWD) & sEnviron & "cd /home/" & sUserID & ";" & "$PROGPATH/" & sCmdLine
    If sHostKey <> "" Then GoTo RUN_PLINK
    staRtn = tfnRunRCmd(sHost, sUserID, sPassWD, sCmd, True)
    
    fnRun4GLPricing = staRtn
    
    Exit Function
    
RUN_PLINK:
    If tfnRunRemoteCmd(sHostKey, rapPlink, sCmd, staRtn, sPLinkError, sHost, sUserID, sPassWD) Then
        fnRun4GLPricing = staRtn
    Else
        fnRun4GLPricing = sPLinkError
        If Len(staRtn) > 0 Then fnRun4GLPricing = Trim(fnRun4GLPricing & " " & staRtn)
    End If
End Function

Public Function tfnRemoteFileExists(sFileName As String, ByRef sErrMsg As String) As Boolean
    Dim sHostKey As String
    Dim sPLinkResult As String
    Dim sCmdLine As String
    
    sHostKey = Trim(fnGetServerHostKey())
    If sHostKey <> "" Then GoTo RUN_PLINK
    
    sErrMsg = tfnRunRCmd(tfnGetHostName(), tfnGetUserName(), tfnGetPassword(), "ls " + sFileName & " > /dev/null")
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
        'replace the password with *'s
        sErrMsg = Replace(sErrMsg, "PWD=" & tfnGetNamedString(sErrMsg, "PWD"), "PWD=" & String(Len(tfnGetNamedString(sErrMsg, "PWD")), "*"))
    End If
        
    Exit Function
    
RUN_PLINK:    ' 8696
    If tfnRunRemoteCmd(sHostKey, _
                       rapPlink, _
                       "ls " & sFileName, _
                       sPLinkResult, _
                       sErrMsg, _
                       tfnGetHostName(), _
                       tfnGetUserName(), _
                       tfnGetPassword()) Then
        If InStr(sPLinkResult, "ls:") > 0 Then
            tfnRemoteFileExists = False
        Else
            tfnRemoteFileExists = True
        End If
    Else
        tfnRemoteFileExists = False
    End If
    
End Function

'**********************************************************
'* USE PLINK - 8696

Public Function fnGetServerHostKey(Optional sHost As String = "") As String
    Dim rsTemp As Recordset
    Dim strSQL As String
    'Dim sHost As String
    Dim sAppPath As String
    Dim sName As String
    Dim sValue As String
    
    If Trim(sHost) = "" Then
        sHost = tfnGetHostName()
    End If
    
    If m_bSSH_KEY Then
        If sHost = m_sLastHost Then
            fnGetServerHostKey = m_sSSH_KEY
            Exit Function
        End If
    End If
    
    m_sSSH_KEY = ""
    m_bSSH_KEY = True
    fnGetServerHostKey = ""
    m_DontDeleteRemoteFiles = False
    
    sAppPath = App.Path
    If Right$(sAppPath, 1) <> "\" Then sAppPath = sAppPath & "\"
    If Not io.FileExists(sAppPath & "plink.exe") Then Exit Function
    If Not io.FileExists(sAppPath & "pscp.exe") Then Exit Function
    If Not io.FileExists(sAppPath & "psftp.exe") Then Exit Function
    
    strSQL = "select ini_field_name, ini_value " _
             & "from sys_ini " _
            & "where ini_file_name = 'HOST_KEY' " _
              & "and ini_section = '" & sHost & "' " _
              & "and ini_field_name in ('SSH_KEY', 'DONT_DELETE')"

    With t_dbMainDatabase.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
        While .EOF = False
            sName = UCase(Trim$(.Fields(0).value & ""))
            sValue = UCase(Trim$(.Fields(1).value & ""))
            Select Case sName
                Case "SSH_KEY"
                    m_sSSH_KEY = sValue
                Case "DONT_DELETE"
                    m_DontDeleteRemoteFiles = (sValue = "Y")
            End Select
            .MoveNext
        Wend
       .Close
    End With
    
    m_sLastHost = sHost
    fnGetServerHostKey = m_sSSH_KEY
    
End Function

Public Function tfnRunRemoteCmd(ByVal sHostKey As String, _
                                ByVal eRemoteApp As RemoteApps, _
                                ByVal sCmd As String, _
                                ByRef sResult As String, _
                                ByRef sErrMsg As String, _
                                Optional ByVal sHost As String = "", _
                                Optional ByVal sUser As String = "", _
                                Optional ByVal sPWD As String = "", _
                                Optional ByVal lTimeoutSecs As Long = 2100, _
                                Optional ByVal bPromptContinueWait As Boolean = False) As Boolean
    On Error GoTo FINISHED
    Dim sGuid As String
    Dim sArgs As String
    Dim sAppName As String
    Dim sAppPath As String
    Dim sCommand As String
    Dim sBatFile As String
    Dim sOutFile As String
    Dim sErrFile As String
    Dim hFile As Integer
    Dim lTemp As Long
    
    sErrMsg = ""
    If Trim(sCmd) = "" Then Err.Raise -1, "tfnRunRemoteCmd", "Empty command passed to function."
    If Trim(sHostKey) = "" Then Err.Raise -2, "tfnRunRemoteCmd", "Empty sHostKey passed to function."
    
    sGuid = Trim(GetGUID())
    sHost = Trim(sHost)
    sUser = Trim(sUser)
    sPWD = Trim(sPWD)
    
    If sHost = "" Or sUser = "" Or sPWD = "" Then
        If t_dbMainDatabase Is Nothing Then
            Err.Raise -2, "tfnRunRemoteCmd", "No FactMenu connection found."
        End If
    End If
    
    If sHost = "" Then sHost = tfnGetHostName()
    If sUser = "" Then sUser = tfnGetUserName()
    If sPWD = "" Then sPWD = tfnGetPassword()
    If sGuid = "" Then sGuid = Format$(Now, "yyyy-mm-dd") & "-" & Replace(Format$(Now, "hh:mm:ss"), ":", "-")
    
    ' TEST SPACES IN PATHS
    'sBatFile = "C:\temp\PUT FILES HERE\" & "rap" & sGuid & ".bat"
   'sOutFile = "C:\temp\PUT FILES HERE\" & "rap" & sGuid & ".txt"
    'sErrFile = "C:\temp\PUT FILES HERE\" & "rap" & sGuid & ".err"
    'sAppPath = "C:\temp\PUT PLINK HERE\"
    
    sBatFile = io.LocalAppPath & "rap" & sGuid & ".bat"
    sOutFile = io.LocalAppPath & "rap" & sGuid & ".txt"
    sErrFile = io.LocalAppPath & "rap" & sGuid & ".err"
    
    sAppPath = App.Path
    If Right$(sAppPath, 1) <> "\" Then sAppPath = sAppPath & "\"
    
    If eRemoteApp = rapPlink Then
        If io.FileExists(sBatFile) Then Kill sBatFile
        hFile = FreeFile()
        Open sBatFile For Output As #hFile
        Print #hFile, ". /etc/profile > /dev/null 2> /dev/null;"
        Print #hFile, ". $(pwd)/.profile > /dev/null 2> /dev/null;"
        Print #hFile, sCmd
        Close #hFile
        hFile = 0

        sAppName = "plink.exe"
        sAppPath = sAppPath & sAppName
        sArgs = " -ssh -pw " & sPWD & " -hostkey " & sHostKey & " " & sUser & "@" & sHost & " -m " & Chr$(34) & sBatFile & Chr$(34)
    ElseIf eRemoteApp = rapPscp Then
        ' pscp will set the errorlevel while psftp will not
        ' sCmd must contain user@host:fullfilename and local\fullfilename in the preferred order for get/put
        ' -q hides progress
        ' -batch means to avoid interactive prompts.  if something goes wrong at connection time, the batch job will fail rather than hang.
        ' -r means recursive when passing wildcards
        '
        ' example     get> pscp fred@example.com:/etc/hosts         c:\temp\example-hosts.txt
        ' wildcard ex get> pscp fred@example.com:/etc/hosts/*.txt   c:\temp
        '
        ' example     put> pscp c:\documents\foo.txt                fred@example.com:/tmp/foo
        ' wildcard ex put> pscp c:\documents\*.txt                  fred@example.com:/tmp/foo
        '
        ' -sftp is the newest protocol which attempts to use SSH-2 (if use -scp then warning message occurs, see documentaion)
        '  we added this to allow for case insensitive GET
        '  insens.ex  get> pscp fred@example.com:/etc/hosts/[Tt][Ee][Ss][Tt].txt c:\temp
        '
        sAppName = "pscp.exe"
        sAppPath = sAppPath & sAppName
        sArgs = " -q -batch -sftp -pw " & sPWD & " -hostkey " & sHostKey & " " & sCmd
    ElseIf eRemoteApp = rapPsftp Then
        ' psftp works best for sending delete file commands
        If io.FileExists(sBatFile) Then Kill sBatFile
        hFile = FreeFile()
        Open sBatFile For Output As #hFile
        Print #hFile, sCmd
        Close #hFile
        hFile = 0
        
        sAppName = "psftp.exe"
        sAppPath = sAppPath & sAppName
        sArgs = " -be -batch -pw " & sPWD & " -hostkey " & sHostKey & " " & sUser & "@" & sHost & " -b " & Chr$(34) & sBatFile & Chr$(34)
    Else
        Err.Raise -3, "", "Invalid remote app passed to tfnRunRemoteCmd - " & eRemoteApp
    End If
    
    If Not io.FileExists(sAppPath) Then
        Err.Raise -1, "tfnRunRemoteCmd", sAppName & " program not found. " & vbCrLf & sAppPath
    End If
        
    sCommand = "cmd.exe /c " & Chr(34) & Chr(34) & sAppPath & Chr(34) & sArgs & " 1>" & Chr(34) & sOutFile & Chr(34) & " 2>" & Chr(34) & sErrFile & Chr(34) & Chr(34)
    If Not ShellRemoteAppAndWait(eRemoteApp, sCommand, vbHide, sErrMsg) Then
        tfnRunRemoteCmd = False
        GoTo FINISHED
    End If
    
    If io.FileExists(sOutFile) Then
        lTemp = tfnReadWholeFile(sOutFile, sResult)
    End If
    
    If io.FileExists(sErrFile) Then
        lTemp = tfnReadWholeFile(sErrFile, sErrMsg)
    End If
    
    ' Get rid of trash errors we don't care about.
    If sErrMsg <> "" Then
        If eRemoteApp = rapPscp Then
            sErrMsg = fnCullPscpErrors(sErrMsg)
        End If
    End If
    
    tfnRunRemoteCmd = Len(Trim(sErrMsg)) < 1
    
    Err.Clear
FINISHED:
    If Err.Number <> 0 Then
        tfnRunRemoteCmd = False
        sErrMsg = Err.Description
        Err.Clear
    End If
    On Error Resume Next
    If hFile <> 0 Then Close #hFile
    If m_DontDeleteRemoteFiles = False Then
        If io.FileExists(sBatFile) Then Kill sBatFile
        If io.FileExists(sOutFile) Then Kill sOutFile
        If io.FileExists(sErrFile) Then Kill sErrFile
    End If
    Err.Clear
End Function

Private Function ShellRemoteAppAndWait( _
        ByVal eRemoteApp As RemoteApps, _
        ByVal sCmd As String, _
        ByVal windowstyle As VbAppWinStyle, _
        ByRef sError As String, _
        Optional lTimeoutSecs As Long = 2100, _
        Optional bPromptContinueWait As Boolean = False) As Boolean
        
    Dim ProcessID As Long
    Dim ProcessHandle As Long
    Dim wfsoReply As Long
    Dim nWaitStart As Long
    Dim sAppName As String
    
    On Error GoTo FINISHED
        
    sAppName = fnGetRemoteAppName(eRemoteApp)
    
    ' Shell the program, get its handle,
    ' and wait for it to terminate
    nWaitStart = Timer
    ProcessID = Shell(sCmd, windowstyle)
    ProcessHandle = OpenProcess(SYNCHRONIZE, True, ProcessID)
    wfsoReply = WaitForSingleObject(ProcessHandle, 500)
    Do While wfsoReply = 258
        DoEvents
        If (Timer - nWaitStart) > lTimeoutSecs Then
            If bPromptContinueWait Then
                If MsgBox(sAppName & " time out. Do you want to continue to wait (30 minutes) " _
                        & "for the program to finished?", vbQuestion + vbYesNo) = vbYes Then
                    nWaitStart = Timer
                Else
                    sError = sAppName & " timed out"
                    ShellRemoteAppAndWait = False
                    GoTo FINISHED
                End If
            Else
                ShellRemoteAppAndWait = False
                sError = "Timed out waiting for " & sAppName & " to finish."
                GoTo FINISHED
            End If
        End If
        wfsoReply = WaitForSingleObject(ProcessHandle, 500)
    Loop
    
    ShellRemoteAppAndWait = True
    CloseHandle ProcessHandle
    
    Err.Clear
FINISHED:
    If Err.Number <> 0 Then
        sError = Err.Description
        ShellRemoteAppAndWait = False
        Err.Clear
    End If
End Function

Private Function ShellRemoteAppAndWait_WSH_RUN( _
        ByVal eRemoteApp As RemoteApps, _
        ByVal sCmd As String, _
        ByVal windowstyle As VbAppWinStyle, _
        ByRef sError As String, _
        Optional lTimeoutSecs As Long = 2100, _
        Optional bPromptContinueWait As Boolean = False) As Boolean
        
    Dim ProcessID As Long
    Dim ProcessHandle As Long
    Dim wfsoReply As Long
    Dim nWaitStart As Long
    Dim sAppName As String
    Dim Wscript, objShell, objExecObject, strLine As String, strip As Long
    On Error GoTo FINISHED
        
    sAppName = fnGetRemoteAppName(eRemoteApp)
    
    ' Shell the program, get its handle,
    ' and wait for it to terminate
    nWaitStart = Timer
    
    Set objShell = CreateObject("WScript.Shell")
    Set objExecObject = objShell.Exec("%comspec% /c " & sCmd)
    
    Do Until objExecObject.StdOut.AtEndOfStream
        strLine = objExecObject.StdOut.ReadLine()
        strip = InStr(strLine, "Address")
        If strip <> 0 Then
            Wscript.Echo strLine
        End If
    Loop
    
    Do Until objExecObject.StdErr.AtEndOfStream
        strLine = objExecObject.StdErr.ReadLine()
        strip = InStr(strLine, "Address")
        If strip <> 0 Then
            sError = sError & vbCrLf & strLine
        End If
    Loop
        
    ShellRemoteAppAndWait_WSH_RUN = True
    
    Err.Clear
FINISHED:
    If Err.Number <> 0 Then
        sError = Err.Description
        ShellRemoteAppAndWait_WSH_RUN = False
        Err.Clear
    End If
End Function

Public Function tfnReadWholeFile(ByVal sFile As String, ByRef sBuffer As String) As Long
    Dim hFile As Integer
    
    sBuffer = ""
    
    If io.FileExists(sFile) Then
        hFile = FreeFile()
        Open sFile For Input As #hFile
        sBuffer = Input$(LOF(hFile), hFile)
        Close #hFile
    End If
    
    tfnReadWholeFile = Len(sBuffer)
    
End Function

Public Function GetGUID() As String
    '(c) 2000 Gus Molina
    Dim udtGUID As GUID
     
    If (CoCreateGuid(udtGUID) = 0) Then
        GetGUID = _
            String(8 - Len(Hex$(udtGUID.Data1)), "0") & Hex$(udtGUID.Data1) & _
            String(4 - Len(Hex$(udtGUID.Data2)), "0") & Hex$(udtGUID.Data2) & _
            String(4 - Len(Hex$(udtGUID.Data3)), "0") & Hex$(udtGUID.Data3) & _
            IIf((udtGUID.Data4(0) < &H10), "0", "") & Hex$(udtGUID.Data4(0)) & _
            IIf((udtGUID.Data4(1) < &H10), "0", "") & Hex$(udtGUID.Data4(1)) & _
            IIf((udtGUID.Data4(2) < &H10), "0", "") & Hex$(udtGUID.Data4(2)) & _
            IIf((udtGUID.Data4(3) < &H10), "0", "") & Hex$(udtGUID.Data4(3)) & _
            IIf((udtGUID.Data4(4) < &H10), "0", "") & Hex$(udtGUID.Data4(4)) & _
            IIf((udtGUID.Data4(5) < &H10), "0", "") & Hex$(udtGUID.Data4(5)) & _
            IIf((udtGUID.Data4(6) < &H10), "0", "") & Hex$(udtGUID.Data4(6)) & _
            IIf((udtGUID.Data4(7) < &H10), "0", "") & Hex$(udtGUID.Data4(7))
    End If

End Function

Private Function fnGetRemoteAppName(eRemoteApp As RemoteApps)
    fnGetRemoteAppName = "remote application"
    If eRemoteApp = rapPlink Then
        fnGetRemoteAppName = "plink.exe"
    ElseIf eRemoteApp = rapPscp Then
        fnGetRemoteAppName = "pscp.exe"
    End If
End Function

Private Function fnCullPscpErrors(ByVal sErrMsg As String) As String
    Dim sTemp As String
    Dim ary() As String
    Dim i As Long
    Dim x As Long
    Dim bChanged As Boolean
    Dim cnt As Long
    
    On Error GoTo FINISHED
    
    fnCullPscpErrors = sErrMsg
    If InStr(LCase(sErrMsg), " is a directory") < 1 Then
        Exit Function
    End If
    
    sErrMsg = Trim(sErrMsg) & vbLf
    sErrMsg = Replace(sErrMsg, vbCrLf, vbLf)
    
    ary = Split(sErrMsg, vbLf)
    cnt = UBound(ary)
    For i = 0 To cnt
        If Trim(ary(i)) <> "" Then
            If InStr(LCase(ary(i)), " is a directory") < 1 Then
                If sTemp <> "" Then
                    sTemp = sTemp & vbLf & ary(i)
                Else
                    sTemp = ary(i)
                End If
            End If
        End If
    Next
    
    sTemp = Trim(sTemp)
    fnCullPscpErrors = sTemp

FINISHED:
    If Err.Number <> 0 Then
        Debug.Print Err.Description
        Err.Clear
    End If
End Function
