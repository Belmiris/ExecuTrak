Attribute VB_Name = "modDatabaseVersion"
Option Explicit

Private Const CONNECT_HOST = ";HOST"
Private Const CONNECT_DBPATH = ";DB"
Private Const CONNECT_USERID = ";UID"
Private Const CONNECT_PSWD = ";PWD"
Private Const CMD_CHMOD = "chmod 666 "
Private Const WINSOCK_PORT = 512
Private Const MAX_MSG_LEN = 4096
Private Const DATABASE_EXTN = ".dbs"
Private Const ERR_LOGIN = -1


Private Declare Function WinsockRCmd Lib "RCMD.DLL" (ByVal RHost As String, ByVal RPort As Integer, ByVal LocalUser As String, ByVal RemoteUser As String, ByVal Cmd As String, ByVal ErrorMsg As String, ByVal ErrLen As Integer) As Integer
Private Declare Function RCmdRead Lib "RCMD.DLL" (ByVal hRCmd As Integer, ByVal RData As String, ByVal RCount As Integer) As Integer
Private Declare Function RCmdReadByte Lib "RCMD.DLL" (ByVal hRCmd As Integer) As Integer
Private Declare Function RCmdClose Lib "RCMD.DLL" (ByVal hRCmd As Integer) As Integer
Private Declare Function RCmdHandle Lib "RCMD.DLL" (ByVal hRCmd As Integer) As Integer
Private Declare Function RCmdCancel Lib "RCMD.DLL" () As Integer
Private Declare Function RCmdNoEvents Lib "RCMD.DLL" () As Integer
Private Declare Function RCmdEvents Lib "RCMD.DLL" () As Integer

Private Const sINIFile As String = "C:\Factor\dbverson.ini"

Private Function fnChangeFileAttr(sTable As String) As Boolean
    Dim sHost As String
    Dim sLocalUID As String
    Dim sRemoteUID As String
    Dim sDBPath As String
    Dim sErrMsg As String
    Dim nMsgLen As Integer
    Dim nCode As Integer
    Dim sCmd As String
    Dim nOutput As Integer
    
    sHost = tfnGetNamedString(t_dbMainDatabase.Connect, CONNECT_HOST)
    sLocalUID = tfnGetNamedString(t_dbMainDatabase.Connect, CONNECT_USERID)
    sRemoteUID = tfnGetNamedString(t_dbMainDatabase.Connect, CONNECT_PSWD)

    sDBPath = tfnGetNamedString(t_dbMainDatabase.Connect, CONNECT_DBPATH)

    fnChangeFileAttr = False
    
    sCmd = "cd " & sDBPath & DATABASE_EXTN & szSEMICOLON & " " & CMD_CHMOD & Left(sTable, 4) & "*"
    
    nCode = ERR_LOGIN
    sErrMsg = Space(MAX_MSG_LEN)
    nCode = WinsockRCmd(sHost, WINSOCK_PORT, sLocalUID, sRemoteUID, sCmd, sErrMsg, MAX_MSG_LEN)
    
    If nCode >= 0 Then
        fnChangeFileAttr = True
    Else
        MsgBox Trim(sErrMsg), vbExclamation
    End If
    Exit Function
    
errRunShell:
    Err.Clear
End Function


Public Function tfnCheckDatabaseVersion() As Boolean
    Const gszSYSTEM_USERNAME = "ssfactor"
    
    Const sID_INDEX As String = "INDEX"
    Const sKEY_RELEASE As String = "RELEASE"
    Const sID_SECTION As String = "SECTION_"
    Const sKEY_VERSION As String = "VERSION_"
    Const sKEY_SQL As String = "SQLSTMT_"
    Const sKEY_QRT_FILE As String = "QRTFILE_"
    Const sKEY_EXE_FILE As String = "EXEFILE_"
    Const sKEY_PERMISSION As String = "PERMISSION_"
    
    Dim sDBReleaseVersion As String, sDBBetaVersion As String, sDBProfileVersion As String
    Dim sDBLatestBetaVersion As String, bRunSuccessful As Boolean
    Dim bCheckPrivledges As Boolean, szUserName As String
    
    Dim sArySection() As String, sAryVersion() As String, sArySqlStmt() As String
    Dim sAryQrtFile() As String, sAryExeFile() As String, sAryPermission() As String
    
    Dim nIndexCount As Integer, nSqlCount As Integer, nQrtCount As Integer, i As Integer
    Dim nExeCount As Integer, nPermiCount As Integer, nTemp As Integer, bTemp As Boolean
    
    'get database version from remote database
    If Not fnGetDatabaseVersion(sDBReleaseVersion, sDBBetaVersion) Then Exit Function
    
    sDBProfileVersion = tfnReadINI(sID_INDEX, sKEY_RELEASE, sINIFile)
    
    If sDBProfileVersion = "" Then
        tfnCheckDatabaseVersion = True
        Exit Function
    End If
    
    If sDBReleaseVersion >= sDBProfileVersion Then  'no table altered needed
        tfnCheckDatabaseVersion = True
        Exit Function
    End If
    
    'QUIT IF THE DATABASE RELEASE VERSION IS NOT 3.19
    MsgBox "Factor Database Version 3.19 or greater is required for this version of the C-Store" & _
        " Auditing System." & vbCrLf & vbCrLf & "Contact technical support for further information."
    tfnCheckDatabaseVersion = False
    Exit Function
    
    'ALL CODES BELOW WILL BE SKIPPED
    
    'Read Profile  [INDEX], KEY = SECTION_?
    nIndexCount = fnReadINIContents(sID_INDEX, sID_SECTION, sArySection)
    
    If nIndexCount = 0 Then
        tfnCheckDatabaseVersion = True
        Exit Function
    End If
    
    'Read Profile  [INDEX], KEY = VERSION_?
    nTemp = fnReadINIContents(sID_INDEX, sKEY_VERSION, sAryVersion)
    
    If nTemp <> nIndexCount Then  'SECTION count and VERSION count are not match
        Exit Function
    End If
    
    bRunSuccessful = True
    sDBLatestBetaVersion = sDBBetaVersion
    
    'loop through all sections for 4 keys
    For i = 0 To nIndexCount - 1
        If sAryVersion(i) > sDBBetaVersion Then  'for each section with version greater than beta version
            'check privledges
            If Not bCheckPrivledges Then
                szUserName = tfnGetUserName()
                If szUserName <> gszSYSTEM_USERNAME Then
                    MsgBox "User Name: " & szUserName & " does not have the Privledges to Update the database." & _
                        vbCrLf & vbCrLf & "Program will be terminated.", vbExclamation
                        bRunSuccessful = False
                        Exit For
                    End
                End If
                bCheckPrivledges = True
            End If
            
            'Read Profile  [SECTION_?], KEY = SQLSTMT_?
            nSqlCount = fnReadINIContents(sArySection(i), sKEY_SQL, sArySqlStmt)
            If nSqlCount > 0 Then
                #If Not SKIP_CHECK Then
                    bRunSuccessful = fnDoSQL(sArySqlStmt, nSqlCount)
                #End If
            End If
            
            'Read Profile  [SECTION_?], KEY = QRTFILE_?
            nQrtCount = fnReadINIContents(sArySection(i), sKEY_QRT_FILE, sAryQrtFile)
            If nQrtCount > 0 Then
                #If Not SKIP_CHECK Then
                    bTemp = fnDoQRT(sAryQrtFile, nQrtCount)
                    If bRunSuccessful Then bRunSuccessful = bTemp
                #End If
            End If
            
            'Read Profile  [SECTION_?], KEY = EXEFILE_?
            nExeCount = fnReadINIContents(sArySection(i), sKEY_EXE_FILE, sAryExeFile)
            If nExeCount > 0 Then
                #If Not SKIP_CHECK Then
                    bTemp = fnDoEXE(sAryExeFile, nExeCount)
                    If bRunSuccessful Then bRunSuccessful = bTemp
                #End If
            End If
        
            'Read Profile  [SECTION_?], KEY = PERMISSION_?
            nPermiCount = fnReadINIContents(sArySection(i), sKEY_PERMISSION, sAryPermission)
            If nPermiCount > 0 Then
                #If Not SKIP_CHECK Then
                    bTemp = fnDoPERMISSION(sAryPermission, nPermiCount)
                    If bRunSuccessful Then bRunSuccessful = bTemp
                #End If
            End If
            
            'if everything was successful then update the database beta version
            If bRunSuccessful Then
                If sAryVersion(i) > sDBLatestBetaVersion Then sDBLatestBetaVersion = sAryVersion(i)
            End If
        End If
    Next i
        
    'if everything was successful then update the database beta version
    If bRunSuccessful Then
        If sDBLatestBetaVersion <> sDBBetaVersion Then
            If Not fnUpdateDBBetaVersion(sDBLatestBetaVersion) Then
                Exit Function
            End If
        End If
        tfnCheckDatabaseVersion = True
    Else
        tfnCheckDatabaseVersion = False
    End If
End Function

Private Function fnDoSQL(sAryTemp() As String, nCount As Integer) As Boolean
    Dim i As Integer, bError As Boolean
    
    On Error GoTo ErrorExecSQL
    
    For i = 0 To nCount - 1
        t_dbMainDatabase.ExecuteSQL sAryTemp(i)
    Next i
    
    If Not bError Then fnDoSQL = True
    
    Exit Function

ErrorExecSQL:
    If InStr(UCase(sAryTemp(i)), "DROP") > 0 Then  'DROP keyword found, do not show error message
'        tfnErrHandler "fnDoSQL", sAryTemp(i), False
        Resume Next
    Else
        bError = True
        tfnErrHandler "fnDoSQL", sAryTemp(i)
        If MsgBox("An error has occurred in function 'fnDoSQL'.  Do you want to continue?", vbYesNo + vbQuestion) = vbYes Then
            Resume Next
        End If
    End If
End Function

Private Function fnDoPERMISSION(sAryTemp() As String, nCount As Integer) As Boolean
    Dim i As Integer, bError As Boolean
    
    For i = 0 To nCount - 1
        If Not fnChangeFileAttr(sAryTemp(i)) Then
            bError = True
            If MsgBox("An error has occurred in function 'fnDoPERMISSION'.  Do you want to continue?", vbYesNo + vbQuestion) = vbNo Then
                Exit Function
            End If
        End If
    Next i
    fnDoPERMISSION = Not bError
End Function

Private Function fnDoQRT(sAryTemp() As String, nCount As Integer) As Boolean
    'not implemented
    fnDoQRT = True
End Function

Private Function fnDoEXE(sAryTemp() As String, nCount As Integer) As Boolean
    'not implemented
    fnDoEXE = True
End Function

Private Function fnReadINIContents(sSectionID As String, sKeyId As String, sAryTemp() As String) As Integer
    Dim nProfileCount As Integer, sTemp As String, i As Integer
    
    i = 1
    nProfileCount = -1
    
    ReDim sAryTemp(0)
    
    Do
        sTemp = tfnReadINI(sSectionID, sKeyId & CStr(i), sINIFile)
        
        If sTemp <> "" Then
            Debug.Print "SECTION="; sSectionID, "KEY="; sKeyId & CStr(i), "DATA="; sTemp
            nProfileCount = nProfileCount + 1
            On Error GoTo ErrorRedim
            ReDim Preserve sAryTemp(nProfileCount)
            sAryTemp(nProfileCount) = sTemp
            i = i + 1
        End If
    Loop Until sTemp = ""
    
    fnReadINIContents = nProfileCount + 1
    
    Exit Function
    
ErrorRedim:
    'do nothing
    fnReadINIContents = 0
End Function

'if sys_parm 3 and/or 6 does not exists, create them
Private Function fnGetDatabaseVersion(sDBReleaseVersion As String, sDBBetaVersion As String) As Boolean
    Dim sSql As String, rsTemp As Recordset

    sDBReleaseVersion = ""
    sDBBetaVersion = ""
    
    On Error GoTo ErrorReadData
    sSql = "SELECT parm_nbr, parm_field FROM sys_parm WHERE parm_nbr = 3 OR parm_nbr = 6"
    Set rsTemp = t_dbMainDatabase.OpenRecordset(sSql, dbOpenSnapshot, dbSQLPassThrough)
    
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveLast
        rsTemp.MoveFirst
    End If
    
    'sys_parm 3 and 6 not found, create them
    If rsTemp.RecordCount = 0 Then
        sSql = "INSERT INTO sys_parm (parm_nbr, parm_field, parm_desc) VALUES "
        sSql = sSql & "(3, '','Database Release Version')"
        t_dbMainDatabase.ExecuteSQL sSql
        
        sSql = "INSERT INTO sys_parm (parm_nbr, parm_field, parm_desc) VALUES "
        sSql = sSql & "(6, '','Database Beta Version')"
        t_dbMainDatabase.ExecuteSQL sSql
    ElseIf rsTemp.RecordCount = 1 Then
        If rsTemp!parm_nbr = 3 Then
            sDBReleaseVersion = IIf(IsNull(rsTemp!parm_field), "", Trim(rsTemp!parm_field))
            
            sSql = "INSERT INTO sys_parm (parm_nbr, parm_field, parm_desc) VALUES "
            sSql = sSql & "(6, '','Database Beta Version')"
            t_dbMainDatabase.ExecuteSQL sSql
        Else  '= 6
            sDBBetaVersion = IIf(IsNull(rsTemp!parm_field), "", Trim(rsTemp!parm_field))
            
            sSql = "INSERT INTO sys_parm (parm_nbr, parm_field, parm_desc) VALUES "
            sSql = sSql & "(3, '','Database Release Version')"
            t_dbMainDatabase.ExecuteSQL sSql
        End If
    Else
        While Not rsTemp.EOF
            If rsTemp!parm_nbr = 3 Then  'database release version
                sDBReleaseVersion = IIf(IsNull(rsTemp!parm_field), "", Trim(rsTemp!parm_field))
            Else  'database beta version
                sDBBetaVersion = IIf(IsNull(rsTemp!parm_field), "", Trim(rsTemp!parm_field))
            End If
            rsTemp.MoveNext
        Wend
    End If
    
    fnGetDatabaseVersion = True
    
    Exit Function

ErrorReadData:
    tfnErrHandler "fnGetDatabaseVersion", sSql
End Function

Private Function fnUpdateDBBetaVersion(sDBBetaVersion As String) As Boolean
    Dim sSql As String
    
    sSql = "UPDATE sys_parm SET parm_field = " & tfnSQLString(sDBBetaVersion)
    sSql = sSql & "WHERE parm_nbr = 6"  'database beta version
    
    On Error GoTo ErrorExecSQL
    t_dbMainDatabase.ExecuteSQL sSql
    
    fnUpdateDBBetaVersion = True
    
    Exit Function

ErrorExecSQL:
    tfnErrHandler "fnUpdateDBBetaVersion", sSql
End Function
