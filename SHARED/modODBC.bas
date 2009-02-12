Attribute VB_Name = "modODBC"
Option Explicit

Private Declare Function SQLAllocConnect Lib "odbc32.dll" (ByVal henv&, phdbc&) As Integer
Private Declare Function SQLAllocEnv Lib "odbc32.dll" (phenv&) As Integer
Private Declare Function SQLAllocStmt Lib "odbc32.dll" (ByVal hdbc&, phstmt&) As Integer
Private Declare Function SQLFetch Lib "odbc32.dll" (ByVal hstmt&) As Integer
Private Declare Function SQLFreeConnect Lib "odbc32.dll" (ByVal hdbc&) As Integer
Private Declare Function SQLFreeEnv Lib "odbc32.dll" (ByVal henv&) As Integer
Private Declare Function SQLFreeStmt Lib "odbc32.dll" (ByVal hstmt&, ByVal fOption%) As Integer
Private Declare Function SQLDrivers Lib "odbc32.dll" (ByVal henv&, ByVal fDirection%, ByVal szDriverDesc$, ByVal cbDriverDescMax%, pcbDriverDesc%, ByVal szDriverAttr$, ByVal cbDrvrAttrMax%, pcbDrvrAttr%) As Integer
Private Declare Function SQLDataSources Lib "odbc32.dll" (ByVal henv&, ByVal fDirection%, ByVal szDSN$, ByVal cbDSNMax%, pcbDSN%, ByVal szDescription$, ByVal cbDescriptionMax%, pcbDescription%) As Integer

Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

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

Private Declare Function RegCloseKey Lib "advapi32.dll" _
    (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
    (ByVal hKey As Long, _
     ByVal lpSubKey As String, _
     ByVal ulOptions As Long, _
     ByVal samDesired As Long, _
     phkResult As Long) As Long
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
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" _
    (ByVal hKey As Long, _
     ByVal lpValueName As String, _
     ByVal lpReserved As Long, _
     lpType As Long, _
     ByVal lpData As Long, _
     lpcbData As Long) As Long
     
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" _
    (ByVal hKey As Long, _
     ByVal lpSubKey As String, _
     phkResult As Long) As Long
     
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" _
    (ByVal hKey As Long, _
     ByVal lpSubKey As String, _
     ByVal Reserved As Long, _
     ByVal lpClass As String, _
     ByVal dwOptions As Long, _
     ByVal samDesired As Long, _
     lpSecurityAttributes As SECURITY_ATTRIBUTES, _
     phkResult As Long, _
     lpdwDisposition As Long) As Long
     
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
    (ByVal hKey As Long, _
     ByVal lpValueName As String, _
     ByVal Reserved As Long, _
     ByVal dwType As Long, _
     ByVal lpData As String, _
     ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.

Private Const szDRIVER_DESCRIPTION As String = "INFORMIX"
Private Const szSECURITY As String = "XTrackSecurity"
Private Const szSECURITY_DATABASE_NETWORK As String = "/factor/factor"
Private Const szSECURITY_DATABASE_LOCAL As String = "security"
Private Const szHENV_ERROR As String = "Cannot Allocate Environment Handle"

Private Const SQL_ERROR = -1
Private Const SQL_INVALID_HANDLE = -2
Private Const SQL_NO_DATA_FOUND = 100
Private Const SQL_SUCCESS = 0
Private Const SQL_SUCCESS_WITH_INFO = 1

Private Const SQL_FETCH_NEXT = 1
Private Const SQL_FETCH_FIRST = 2
Private Const SQL_FETCH_LAST = 3
Private Const SQL_FETCH_PRIOR = 4
Private Const SQL_FETCH_ABSOLUTE = 5
Private Const SQL_FETCH_RELATIVE = 6
Private Const SQL_FETCH_BOOKMARK = 8

Private Const szODBC_DATA_SOURCES As String = "ODBC Data Sources"
Private Const szODBC_DATA_SOURCES32 As String = "ODBC 32 bit Data Sources"
Private Const szDRIVER As String = "Driver"
Private Const szDRIVER32 As String = "Driver32"
Private Const szODBC_ERROR As String = "<Error>"
Private Const szODBC_REG_KEY1 = ".Default\Software\ODBC\ODBC.INI\"
Private Const szODBC_REG_KEY2 = "Software\ODBC\ODBC.INI\"
Private Const szODBC_DATABASE = "Database"
Private Const szODBC_HOST = "HostName"
Private Const szODBC_SERVERNAME = "ServerName"
Private Const szODBC_SERVICE = "Service"
Private Const szODBC_YIELDPROC = "YieldProc"
Private Const szODBC_CB = "CursorBehavior"
Private Const szODBC_PROTOCOL = "Protocol"
Private Const szODBC_DRIVER = "Driver"

Private Const szODBC_SERVERNAME2 = "Server"
Private Const szODBC_CLIENT_LOCALE = "CLIENT_LOCALE"
Private Const szODBC_DB_LOCALE = "DB_LOCALE"
Private Const szODBC_VMBCHARLENEXACT = "VMBCHARLENEXACT"
Private Const szODBC_ENABLESCROLLABLECURSORS = "ENABLESCROLLABLECURSORS"
Private Const szODBC_ENABLEINSERTCURSORS = "ENABLEINSERTCURSORS"
Private Const szODBC_OPTIMIZEAUTOCOMMIT = "OPTIMIZEAUTOCOMMIT"
Private Const szODBC_OPTOFC = "OPTOFC"
Private Const szODBC_REPORTKEYSETCURSORS = "REPORTKEYSETCURSORS"
Private Const szODBC_NEEDODBCTYPESONLY = "NEEDODBCTYPESONLY"
Private Const szODBC_FETCHBUFFERSIZE = "FETCHBUFFERSIZE"

Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003
Private Const KEY_ALL_ACCESS = &H3F
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const SYNCHRONIZE = &H100000
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))

Private Const REG_OPTION_NON_VOLATILE = 0
Private Const REG_SZ As Long = 1
Private Const REG_DWORD As Long = 4
Private Const REG_OPTION_VOLATILE = 1           ' Key is not preserved when system is rebooted

Private m_sODBC_INI_Path As String
Private m_lODBC_INI_Key As Long
'

Public Function tfnGetDataSources(colDSN As Collection) As Integer
    Dim szDataSourceName As String       'Data Source Name returned from SQL function call
    Dim szDriverDescription As String    'Driver Description returned from SQL function call
    Dim nDataSourceNameLen As Integer    'DSN string length returned from SQL function call
    Dim nDriverDescriptionLen As Integer 'Driver Description string length returned from SQL function call
    Dim nReturn As Integer               'return code from SQL call
    Dim lHenv As Long                    'handle to the environment

    Const DSN_LENGTH = 32          'Data Source Name (fixed) length passed to SQL function
    Const DRIVER_DESC_LENGTH = 255 'Driver Description (fixed) length passed to SQL function

    If Not colDSN Is Nothing Then
        While colDSN.Count > 0
            colDSN.Remove 1
        Wend
        
        Set colDSN = Nothing
    End If
    
    Set colDSN = New Collection 'clear drivers collection
    
    If SQLAllocEnv(lHenv) <> SQL_ERROR Then 'get valid environment handle
        szDataSourceName = String(DSN_LENGTH, " ")
        szDriverDescription = String(DRIVER_DESC_LENGTH, " ") 'set fixed length strings, pad with spaces

        nReturn = SQLDataSources(lHenv, SQL_FETCH_FIRST, szDataSourceName, DSN_LENGTH, nDataSourceNameLen, _
            szDriverDescription, DRIVER_DESC_LENGTH, nDriverDescriptionLen) 'get first DSN/Driver Description values

        While nReturn = SQL_SUCCESS Or nReturn = SQL_SUCCESS_WITH_INFO 'process if SQL function call OK
            szDriverDescription = Left(szDriverDescription, nDriverDescriptionLen) 'strip any spaces and --> terminating NULL
            szDataSourceName = Left(szDataSourceName, nDataSourceNameLen)

            If InStr(1, szDriverDescription, szDRIVER_DESCRIPTION) > 0 Then 'check for application Driver
                If Not szDataSourceName = szSECURITY Then 'don'y display security entry it its exists
                    colDSN.Add Item:=szDataSourceName, Key:=szDataSourceName 'save driver using DataSourceName as Key
                End If
            End If

            szDataSourceName = String(DSN_LENGTH, " ")
            szDriverDescription = String(DRIVER_DESC_LENGTH, " ") 're-initialized fixed length strings for next fetch

            nReturn = SQLDataSources(lHenv, SQL_FETCH_NEXT, szDataSourceName, DSN_LENGTH, nDataSourceNameLen, _
                szDriverDescription, DRIVER_DESC_LENGTH, nDriverDescriptionLen) 'get next DSN/Driver Description values
        Wend

        SQLFreeEnv (lHenv) 'free the environment handle
    Else
        subCriticalMsg szHENV_ERROR & ", TXFCOMBN", "Module Error"
    End If
    
    tfnGetDataSources = colDSN.Count 'return the number of valid DSN's configured in ODBC.INI
End Function

Public Function tfnConnectString(sDSN As String, m_sUID As String, m_sPWD As String) As String
    Dim sTemp As String
    Dim sODBCKey As String
    Dim sDatabase As String
    Dim m_sHost As String
    
    sODBCKey = m_sODBC_INI_Path & sDSN
    m_sHost = QueryValue(m_lODBC_INI_Key, sODBCKey, szODBC_SERVERNAME)
    If Trim(m_sHost) = "" Then
        m_sHost = QueryValue(m_lODBC_INI_Key, sODBCKey, szODBC_HOST)
    End If
    If Trim(m_sHost) = "" Then
        m_sHost = QueryValue(m_lODBC_INI_Key, sODBCKey, szODBC_SERVERNAME2)
    End If
    
    tfnConnectString = "ODBC;DSN=" & sDSN & ";UID=" & m_sUID _
            & ";PWD=" & m_sPWD
    sDatabase = QueryValue(m_lODBC_INI_Key, sODBCKey, szODBC_DATABASE)
    tfnConnectString = tfnConnectString & ";DB=" & sDatabase & ";HOST=" & m_sHost
    sTemp = QueryValue(m_lODBC_INI_Key, sODBCKey, szODBC_SERVICE)
    tfnConnectString = tfnConnectString & ";SERV=" & sTemp
    sTemp = QueryValue(m_lODBC_INI_Key, sODBCKey, szODBC_YIELDPROC)
    tfnConnectString = tfnConnectString & ";YLD=" & sTemp
    sTemp = QueryValue(m_lODBC_INI_Key, sODBCKey, szODBC_CB)
    tfnConnectString = tfnConnectString & ";CB=" & sTemp
    sTemp = QueryValue(m_lODBC_INI_Key, sODBCKey, szODBC_PROTOCOL)
    tfnConnectString = tfnConnectString & ";PRO=" & sTemp
End Function

Public Function tfnSetODBCINIPath(sDSN As String) As Boolean
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
        tfnSetODBCINIPath = False
    Else
        tfnSetODBCINIPath = True
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
                    vValue = fnSZStr2Str(sValue)
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

Private Function fnSZStr2Str(szStr As String) As String
    Dim nPos As Integer
    
    nPos = InStr(szStr, Chr(0))
    If nPos > 0 Then
        fnSZStr2Str = Left(szStr, nPos - 1)
    Else
        fnSZStr2Str = szStr
    End If
End Function

Private Function QueryValue(ByVal lKey As Long, _
                           sKeyName As String, _
                           sValueName As String) As String
    Dim lRetVal As Long         'result of the API functions
    Dim hKey As Long         'handle of opened key
    Dim vValue As Variant      'setting of queried value

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

Private Function RegKeyExist(ByVal lKey As Long, _
                            sKeyName As String) As Boolean
    Dim lRetVal As Long         'result of the API functions
    Dim hKey As Long         'handle of opened key

    lRetVal = RegOpenKeyEx(lKey, sKeyName, 0, KEY_READ, hKey)
    If lRetVal = 0 Then
        RegCloseKey (hKey)
        RegKeyExist = True
    Else
        RegKeyExist = False
    End If
End Function

Private Function RegSetValue(ByVal lKey As Long, _
                            sKeyName As String, _
                            sValueName As String, _
                            ByVal lType As Long, _
                            ByVal sValue As String) As Boolean
    Dim lRetVal As Long         'result of the API functions
    Dim hKey As Long         'handle of opened key
    Dim vValue As String      'setting of queried value
    Dim cbData As Long
    
    Dim Sec_Att As SECURITY_ATTRIBUTES
    Sec_Att.nLength = 12&
    Sec_Att.lpSecurityDescriptor = 0&
    Sec_Att.bInheritHandle = False
    
    RegSetValue = False

    'lRetVal = RegCreateKeyEx(lKey, sKeyName, 0, 0, REG_OPTION_VOLATILE, KEY_ALL_ACCESS, 0, hKey, cbData)
    lRetVal = RegCreateKeyEx(lKey, sKeyName, 0&, "", REG_OPTION_VOLATILE, KEY_ALL_ACCESS, Sec_Att, hKey, cbData)
    If lRetVal = 0 Then
        If lType >= 0 Then
            If lType = REG_DWORD Then
                vValue = Chr(sValue)
                cbData = 4
            Else
                vValue = sValue & Chr(0)
                cbData = Len(vValue)
            End If
            lRetVal = RegSetValueEx(hKey, sValueName, 0, lType, vValue, cbData)
            If lRetVal = 0 Then
                RegSetValue = True
            End If
        End If
        RegCloseKey (hKey)
    End If
End Function

Private Sub subCriticalMsg(sMsg As String, _
                          sCaption As String)
    MsgBox sMsg, vbOKOnly + vbCritical, sCaption
End Sub

