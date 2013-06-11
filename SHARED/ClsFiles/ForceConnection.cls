VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ForceConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'********************************************************************
' Designed to allow nasty old command line programs to work properly
' in a hosted environment.
'********************************************************************
Private Declare Function RegCloseKey Lib "advapi32.dll" _
    (ByVal hKey As Long) As Long

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
    (ByVal hKey As Long, _
     ByVal lpSubKey As String, _
     ByVal ulOptions As Long, _
     ByVal samDesired As Long, _
     phkResult As Long) As Long

Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" _
    (ByVal hKey As Long, _
     ByVal lpValueName As String, _
     ByVal lpReserved As Long, _
     lpType As Long, _
     ByVal lpData As Long, _
     lpcbData As Long) As Long

Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" _
    (ByVal hKey As Long, _
     ByVal lpValueName As String, _
     ByVal lpReserved As Long, _
     lpType As Long, _
     ByVal lpData As String, _
     lpcbData As Long) As Long

Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" _
    (ByVal hKey As Long, _
     ByVal lpValueName As String, _
     ByVal lpReserved As Long, _
     lpType As Long, _
     lpData As Long, _
     lpcbData As Long) As Long

Private Const szODBC_REG_KEY1 = ".Default\Software\ODBC\ODBC.INI\"
Private Const szODBC_REG_KEY2 = "Software\ODBC\ODBC.INI\"
Private Const szODBC_REG_KEY3 = "SOFTWARE\Informix\SqlHosts\"

Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003

Private Const REG_OPTION_NON_VOLATILE = 0
Private Const REG_SZ As Long = 1
Private Const REG_DWORD As Long = 4
Private Const REG_OPTION_VOLATILE = 1           ' Key is not preserved when system is rebooted

Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const SYNCHRONIZE = &H100000
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))

Private Const ERROR_NONE = 0
Private Const ERROR_BADDB = 1
Private Const ERROR_BADKEY = 2
Private Const ERROR_CANTOPEN = 3
Private Const ERROR_CANTREAD = 4
Private Const ERROR_CANTWRITE = 5
Private Const ERROR_OUTOFMEMORY = 6
Private Const ERROR_INVALID_PARAMETER = 7
Private Const ERROR_ACCESS_DENIED = 8
Private Const ERROR_INVALID_PARAMETERS = 87
Private Const ERROR_NO_MORE_ITEMS = 259

Private Const szODBC_DATABASE = "Database"
Private Const szODBC_CLIENT_LOCALE = "CLIENT_LOCALE"
Private Const szODBC_DB_LOCALE = "DB_LOCALE"
Private Const szODBC_HOST = "HostName"
Private Const szODBC_HOST2 = "Host"
Private Const szODBC_SERVERNAME = "ServerName"
Private Const szODBC_SERVERNAME2 = "Server"
Private Const szODBC_SERVICE = "Service"
Private Const szODBC_YIELDPROC = "YieldProc"
Private Const szODBC_CB = "CursorBehavior"
Private Const szODBC_PROTOCOL = "Protocol"
Private Const szODBC_DRIVER = "Driver"

Private m_sODBC_INI_Path As String
Private m_lODBC_INI_Key As Long
Private m_sServer As String              'Server name which may be differnt than host
Private m_sUID As String
Private m_sPWD As String
Private m_sHost As String
Private m_sDriver As String

Private m_sDSN As String
Private m_connect As String
Private io As CommonIO

Option Explicit

Public Property Get DSN() As String
    DSN = m_sDSN
End Property

Public Property Let DSN(ByVal value As String)
    m_sDSN = value
End Property

Public Property Get connectionString() As String
    connectionString = GetConnectionString()
End Property

Public Property Let connectionString(ByVal value As String)
    m_connect = value
End Property

Public Property Get CommonIO() As CommonIO
    GetCommonIO
    Set CommonIO = io
End Property

Private Function GetCommonIO() As CommonIO
    Dim connect$
    Dim engine As DBEngine
    Dim ws As Workspace
    Dim db As Database
    
    Set io = Nothing
    
    connect = GetConnectionString()
    
    Set engine = New DBEngine
    Set ws = engine.Workspaces(0)
    Set db = ws.OpenDatabase("", False, False, connect)
    
    dsnName = m_sDSN
    Set io = New CommonIO
    
    io.HostedFolder = GetServiceTrakHostedFolder(db)
    io.DatabaseName = GetDatabaseName(connect)
    io.DatabasePath = GetDatabasePath(connect)
    
    Set GetCommonIO = io
    
    db.Close
    Set db = Nothing
    Set ws = Nothing
    Set engine = Nothing
    
End Function

Private Function GetConnectionString() As String
    Dim sTemp As String
    Dim sODBCKey As String
    Dim sDatabase As String
    Dim sDSN As String
    
    If Len(m_connect) > 0 Then
        GetConnectionString = m_connect
        Exit Function
    End If
    
    If Len(m_sDSN) < 1 Then Err.Raise -1, "GetConnectionString", "DSN not set"
    sDSN = m_sDSN
    
    Call SetODBCINIPath(m_sDSN)
    
    m_sUID = "ssfactor"
    m_sPWD = "menus"
    
    sODBCKey = m_sODBC_INI_Path & sDSN
    m_sServer = QueryValue(m_lODBC_INI_Key, sODBCKey, szODBC_SERVERNAME2)
    m_sHost = QueryValue(m_lODBC_INI_Key, sODBCKey, szODBC_SERVERNAME)
    If Trim(m_sHost) = "" Then
        m_sHost = QueryValue(m_lODBC_INI_Key, sODBCKey, szODBC_HOST)
    End If
    If Trim(m_sHost) = "" Then
        m_sHost = QueryValue(m_lODBC_INI_Key, szODBC_REG_KEY3 & m_sServer, szODBC_HOST2)
    End If
    If Trim(m_sHost) = "" Then
        m_sHost = QueryValue(HKEY_LOCAL_MACHINE, szODBC_REG_KEY3 & m_sServer, szODBC_HOST2)
    End If
    If Trim(m_sHost) = "" Then
        m_sHost = QueryValue(m_lODBC_INI_Key, sODBCKey, szODBC_SERVERNAME2)
    End If
    GetConnectionString = "ODBC;DSN=" & sDSN & ";UID=" & m_sUID _
            & ";PWD=" & m_sPWD
    sDatabase = QueryValue(m_lODBC_INI_Key, sODBCKey, szODBC_DATABASE)
    GetConnectionString = GetConnectionString & ";DB=" & sDatabase & ";HOST=" & m_sHost
    sTemp = QueryValue(m_lODBC_INI_Key, sODBCKey, szODBC_SERVICE)
    If Trim(sTemp) = "" Then
        sTemp = QueryValue(m_lODBC_INI_Key, szODBC_REG_KEY3 & m_sServer, szODBC_SERVICE)
    End If
    If Trim(sTemp) = "" Then
        sTemp = QueryValue(HKEY_LOCAL_MACHINE, szODBC_REG_KEY3 & m_sServer, szODBC_SERVICE)
    End If
    GetConnectionString = GetConnectionString & ";SERV=" & sTemp
    sTemp = QueryValue(m_lODBC_INI_Key, sODBCKey, szODBC_YIELDPROC)
    GetConnectionString = GetConnectionString & ";YLD=" & sTemp
    sTemp = QueryValue(m_lODBC_INI_Key, sODBCKey, szODBC_CB)
    GetConnectionString = GetConnectionString & ";CB=" & sTemp
    sTemp = QueryValue(m_lODBC_INI_Key, sODBCKey, szODBC_PROTOCOL)
    If Trim(sTemp) = "" Then
        sTemp = QueryValue(m_lODBC_INI_Key, szODBC_REG_KEY3 & m_sServer, szODBC_PROTOCOL)
    End If
    If Trim(sTemp) = "" Then
        sTemp = QueryValue(HKEY_LOCAL_MACHINE, szODBC_REG_KEY3 & m_sServer, szODBC_PROTOCOL)
    End If
    GetConnectionString = GetConnectionString & ";PRO=" & sTemp
    sTemp = QueryValue(m_lODBC_INI_Key, sODBCKey, szODBC_DB_LOCALE)
    GetConnectionString = GetConnectionString & ";DB_LOCALE=" & sTemp
    sTemp = QueryValue(m_lODBC_INI_Key, sODBCKey, szODBC_CLIENT_LOCALE)
    GetConnectionString = GetConnectionString & ";CLIENT_LOCALE=" & sTemp
    If Trim(m_sServer) <> "" Then
        GetConnectionString = GetConnectionString & ";SERVER=" & m_sServer
    End If
    
    m_connect = GetConnectionString
    
End Function

Private Function SetODBCINIPath(sDSN As String) As Boolean
    Dim sTemp As String
    
    m_sODBC_INI_Path = szODBC_REG_KEY2
    m_lODBC_INI_Key = HKEY_CURRENT_USER
    sTemp = QueryValue(m_lODBC_INI_Key, m_sODBC_INI_Path & sDSN, szODBC_DATABASE)
    If sTemp = "" Then
        m_sODBC_INI_Path = szODBC_REG_KEY2
        m_lODBC_INI_Key = HKEY_LOCAL_MACHINE
        sTemp = QueryValue(m_lODBC_INI_Key, m_sODBC_INI_Path & sDSN, szODBC_DATABASE)
    End If
    If sTemp = "" Then
        m_sODBC_INI_Path = szODBC_REG_KEY1
        m_lODBC_INI_Key = HKEY_USERS
        sTemp = QueryValue(m_lODBC_INI_Key, m_sODBC_INI_Path & sDSN, szODBC_DATABASE)
    End If
    If sTemp = "" Then
        SetODBCINIPath = False
    Else
        SetODBCINIPath = True
    End If

End Function

Private Function QueryValue(ByVal lKey As Long, _
                           sKeyName As String, _
                           sValueName As String) As String
    Dim lRetVal As Long         'result of the API functions
    Dim hKey As Long            'handle of opened key
    Dim vValue As Variant       'setting of queried value

    QueryValue = ""
    lRetVal = RegOpenKeyEx(lKey, sKeyName, 0, KEY_READ, hKey)
    If lRetVal = 0 Then
        lRetVal = QueryValueEx(hKey, sValueName, vValue)
        If lRetVal = 0 Then
            QueryValue = vValue
        End If
        RegCloseKey (hKey)
    End If

End Function

Private Function QueryValueEx(ByVal lhKey As Long, _
                      ByVal szValueName As String, _
                      vValue As Variant) As Long
    Dim cch As Long
    Dim lrc As Long
    Dim lType As Long
    Dim lValue As Long
    Dim sValue As String

    On Error GoTo QueryValueExError

    ' Determine the size and type of data to be read
    lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
    If lrc = ERROR_NONE Then
        Select Case lType
            ' For strings
            Case REG_SZ:
                sValue = String(cch, 0)
                lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
                If lrc = ERROR_NONE Then
                    vValue = RemoveNull(sValue)
                Else
                    vValue = Empty
                End If
            ' For DWORDS
            Case REG_DWORD:
                lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
                If lrc = ERROR_NONE Then
                    vValue = lValue
                End If
            Case Else
                'all other data types not supported
                lrc = -1
        End Select
    End If
QueryValueExExit:
    QueryValueEx = lrc
    Exit Function
QueryValueExError:
    Resume QueryValueExExit

End Function

Private Function RemoveNull(szStr As String) As String
    Dim nPos As Integer
    
    nPos = InStr(szStr, Chr(0))
    If nPos > 0 Then
        RemoveNull = Left(szStr, nPos - 1)
    Else
        RemoveNull = szStr
    End If

End Function

Public Function GetServiceTrakHostedFolder(db As Database) As String
    Dim SQL$
    Dim parm15$
    
    SQL = "select parm_field from sys_parm where parm_nbr = 15"
    With db.OpenRecordset(SQL, dbOpenSnapshot, dbSQLPassThrough)
        If Not .EOF Then
            parm15 = UCase$(Trim$(.Fields(0).value & ""))
        End If
        .Close
    End With
    
    SQL = "select ini_value from sys_ini where ini_section = 'WorkingDirectory' AND ini_file_name = 'HOSTED'"
    If parm15 = "Y" Then
        With db.OpenRecordset(SQL, dbOpenSnapshot, dbSQLPassThrough)
            If Not .EOF Then
                GetServiceTrakHostedFolder = UCase$(Trim$(.Fields(0).value & ""))
            End If
            .Close
        End With
    End If
    
End Function

Public Function GetNamedString(sSource As String, sName As String) As String
    GetNamedString = ""
    If sSource = "" Or sName = "" Then
        Exit Function
    End If
    
    Dim nPos1 As Integer
    Dim nPos2 As Integer
    Dim sUcaseSource As String
    Dim sUcaseName As String
    
    sUcaseSource = UCase(sSource)
    sUcaseName = UCase(sName)
    nPos1 = InStr(sUcaseSource, sUcaseName)

    If nPos1 > 0 Then
        nPos1 = InStr(nPos1, sUcaseSource, "=")
        If nPos1 > 0 Then
            nPos2 = InStr(nPos1, sUcaseSource, ";")
            If nPos2 = 0 Then
                nPos2 = Len(sUcaseSource) + 1
            End If
            nPos1 = nPos1 + 1
            If nPos2 > nPos1 Then
                GetNamedString = Trim(Mid(sSource, nPos1, nPos2 - nPos1))
            End If
        End If
    End If
End Function

Public Function GetDatabaseName(connectionString As String) As String
    Dim sDBPath As String
    Dim sDBName As String
    Dim i As Integer
    
    sDBPath = GetNamedString(connectionString, ";DB")
    If Trim(sDBPath) = "" Then
        sDBPath = GetNamedString(connectionString, "DATABASE")
    End If
    
    i = InStrRev(sDBPath, "/")
    
    If i > 1 Then
        sDBPath = Left(sDBPath, i - 1)
    
        i = InStrRev(sDBPath, "/")
    
        If i > 1 Then
            sDBName = Mid(sDBPath, i + 1)
        End If
    End If

    GetDatabaseName = sDBName

End Function

Public Function GetDatabasePath(connectionString As String) As String
    
    GetDatabasePath = GetNamedString(connectionString, ";DB")
    If Trim(GetDatabasePath) = "" Then
        GetDatabasePath = GetNamedString(connectionString, "DATABASE")
    End If

End Function

