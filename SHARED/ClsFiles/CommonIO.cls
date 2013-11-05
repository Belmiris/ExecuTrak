VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CommonIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_WorkingFolder$
Private m_ApplicationPath$
Private m_StandardLogPath$
Private m_StandardLogFile$
Private m_LocalAppPath$
Private m_LocalLogPath$
Private m_LocalLogFile$
Private m_LocalTempPath$
Private m_WindowsTempPath$
Private m_ExecuTrakEdiDataPath$
Private m_ExecuTrakPath$
Private m_FuelMovementUploadPath$
Private m_XMLDataPath$
Private m_StandardEmailPath$
Private m_LocalEmailPath$
Private m_FactorConfigurationFile$
Private m_FactorEDIDirectory$
Private m_DTNLogFilePath$
Private m_DTNInputPath$
Private m_TLCInboundDirectory$
Private m_User$
Private m_DatabaseName$
Private m_UnixDataDirectory$
Private m_DatabasePath$
Private m_EDIXLATERoot$
Private m_EDIXLATEPath$
Private m_EDIXLATEFile$
Private m_MultiDatabase As Boolean

Private m_sODBC_INI_Path As String
Private m_lODBC_INI_Key As Long
Private m_sServer As String              'Server name which may be different than host
Private m_sUID As String
Private m_sPWD As String
Private m_sHost As String
Private m_sDriver As String
Private m_sDSN As String
Private m_connect As String

Private m_loaded As Boolean

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function GetFileTime Lib "kernel32.dll" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32.dll" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, ByRef TokenHandle As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long

Private Declare Function GetUserProfileDirectory Lib "userenv.dll" Alias "GetUserProfileDirectoryA" (ByVal hToken As Long, ByVal lpProfileDir As String, lpcchSize As Long) As Boolean

Private Const LOCAL_FACTOR_PATH = "C:\FACTOR\"
Private Const szODBC_REG_KEY1 = ".Default\Software\ODBC\ODBC.INI\"
Private Const szODBC_REG_KEY2 = "Software\ODBC\ODBC.INI\"
Private Const szODBC_REG_KEY3 = "SOFTWARE\Informix\SqlHosts\"
Private Const ERROR_NOTLOADED = "CommonIO was used before being initialized! Please contact Factor support."

Private Const MAX_PATH = 260
Private Const MAXDWORD = &HFFFF
Private Const INVALID_HANDLE_VALUE = -1

Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100

Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const CREATE_ALWAYS = 2
Private Const CREATE_NEW = 1
Private Const OPEN_ALWAYS = 4
Private Const OPEN_EXISTING = 3
Private Const TRUNCATE_EXISTING = 5

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

Private Const TOKEN_QUERY = (&H8)

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
Private Const szODBC_UID = "UID"
Private Const szODBC_PWD = "PWD"

Private Const ENCRYPT_KEY As String = "T0Rn4do!"

Private Const TICKS_PER_SECOND = 10000000   'FILETIME units are 100s of nanoseconds.

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

' *******************************************************************
' INITIALIZATION METHODS - SINCE WE DON'T HAVE CONSTRUCTORS         *
' *******************************************************************

Public Sub LoadViaDSN(DSN As String, Optional USER As String = "", Optional pwd = "")
    Dim Connect$
    Dim engine As DBEngine
    Dim ws As Workspace
    Dim db As DataBase
    
    m_sDSN = DSN
    If Len(Trim(USER)) > 0 Then m_sUID = Trim(USER)
    If Len(Trim(pwd)) > 0 Then m_sPWD = Trim(pwd)
    
    Connect = GetConnectionString()
    
    Set engine = New DBEngine
    Set ws = engine.Workspaces(0)
    Set db = ws.OpenDatabase("", False, False, Connect)
    
    Call Me.LoadViaDatabase(DSN, db)
    
    db.Close
    Set db = Nothing
    Set ws = Nothing
    Set engine = Nothing

End Sub

Public Sub LoadViaDatabase(DSN As String, db As DataBase)
    
    m_sDSN = Trim(DSN)
    If Len(m_sDSN) < 1 Then Call Err.Raise(-1, , "CommonIO.LoadViaDatabase(): dsn parameter cannot be empty!")
    
    m_WorkingFolder = GetWorkingFolder(db)
    Me.DatabaseName = GetDatabaseName(db.Connect)
    Me.DatabasePath = Me.GetDatabasePath(db.Connect)
    
    m_loaded = True
    
End Sub

' *******************************************************************
' PROPERTIES                                                        *
' *******************************************************************

Public Property Get MultiDatabase() As Boolean
    'If Not m_loaded Then Err.Raise -2, , "MultiDatabase property(): " & ERROR_NOTLOADED
    MultiDatabase = m_MultiDatabase
End Property

Public Property Get WorkingFolder() As String
    'If Not m_loaded Then Err.Raise -2, , "WorkingFolder property(): " & ERROR_NOTLOADED
    WorkingFolder = m_WorkingFolder
End Property

Public Property Get dsnName() As String
    'If Not m_loaded Then Err.Raise -2, , "DsnName property(): " & ERROR_NOTLOADED
    dsnName = m_sDSN
End Property

Public Property Get ApplicationPath() As String
    GetApplicationPath
    ApplicationPath = m_ApplicationPath
End Property

Public Property Let ApplicationPath(ByVal value As String)
    ApplicationPath = value
End Property

Public Property Get StandardLogPath() As String
    GetStandardLogPath
    StandardLogPath = m_StandardLogPath
End Property

Public Property Let StandardLogPath(ByVal value As String)
    m_StandardLogPath = value
End Property

Public Property Get StandardLogFile() As String
    GetStandardLogFile
    StandardLogFile = m_StandardLogFile
End Property

Public Property Let StandardLogFile(ByVal value As String)
    m_StandardLogFile = value
End Property

Public Property Get LocalAppPath() As String
    GetLocalAppPath
    LocalAppPath = m_LocalAppPath
End Property

Public Property Let LocalAppPath(value As String)
    m_LocalAppPath = value
End Property

Public Property Get LocalLogPath() As String
    GetLocalLogPath
    LocalLogPath = m_LocalLogPath
End Property

Public Property Let LocalLogPath(ByVal value As String)
    m_LocalLogPath = value
End Property

Public Property Get LocalLogFile() As String
    GetLocalLogFile
    LocalLogFile = m_LocalLogFile
End Property

Public Property Let LocalLogFile(ByVal value As String)
    m_LocalLogFile = value
End Property

Public Property Get LocalTempPath() As String
    GetLocalTempPath
    LocalTempPath = m_LocalTempPath
End Property

Public Property Let LocalTempPath(ByVal value As String)
    m_LocalTempPath = value
End Property

Public Property Get WindowsTempPath() As String
    GetWindowsTempPath
    WindowsTempPath = m_WindowsTempPath
End Property

Public Property Let WindowsTempPath(ByVal value As String)
    m_WindowsTempPath = value
End Property

Public Property Get ExecuTrakEdiDataPath() As String
    GetExecuTrakEdiDataPath
    ExecuTrakEdiDataPath = m_ExecuTrakEdiDataPath
End Property

Public Property Let ExecuTrakEdiDataPath(ByVal value As String)
    m_ExecuTrakEdiDataPath = value
End Property

Public Property Get FuelMovementUploadPath() As String
    GetFuelMovementUploadPath
    FuelMovementUploadPath = m_FuelMovementUploadPath
End Property

Public Property Let FuelMovementUploadPath(ByVal value As String)
    m_FuelMovementUploadPath = value
End Property

Public Property Get XMLDataPath() As String
    GetXMLDataPath
    XMLDataPath = m_XMLDataPath
End Property

Public Property Let XMLDataPath(ByVal value As String)
    m_XMLDataPath = value
End Property

Public Property Get StandardEmailPath() As String
    GetStandardEmailPath
    StandardEmailPath = m_StandardEmailPath
End Property

Public Property Let StandardEmailPath(ByVal value As String)
    m_StandardEmailPath = value
End Property

Public Property Get LocalEmailPath() As String
    GetLocalEmailPath
    LocalEmailPath = m_LocalEmailPath
End Property

Public Property Let LocalEmailPath(ByVal value As String)
    m_LocalEmailPath = value
End Property

Public Property Get FactorConfigurationFile() As String
    GetFactorConfigurationFile
    FactorConfigurationFile = m_FactorConfigurationFile
End Property

Public Property Let FactorConfigurationFile(ByVal value As String)
    m_FactorConfigurationFile = value
End Property

Public Property Get FactorEDIDirectory() As String
    GetFactorEDIDirectory
    FactorEDIDirectory = m_FactorEDIDirectory
End Property

Public Property Let FactorEDIDirectory(ByVal value As String)
    m_FactorEDIDirectory = value
End Property

Public Property Get DTNLogFilePath() As String
    GetDTNLogFilePath
    DTNLogFilePath = m_DTNLogFilePath
End Property

Public Property Let DTNLogFilePath(ByVal value As String)
    m_DTNLogFilePath = value
End Property

Public Property Get DTNInputPath() As String
    GetDTNInputPath
    DTNInputPath = m_DTNInputPath
End Property

Public Property Let DTNInputPath(ByVal value As String)
    m_DTNInputPath = value
End Property

Public Property Get TLCInboundDirectory() As String
    GetTLCInboundDirectory
    TLCInboundDirectory = m_TLCInboundDirectory
End Property

Public Property Let TLCInboundDirectory(ByVal value As String)
    m_TLCInboundDirectory = value
End Property

Public Property Get ExecuTrakPath() As String
    GetExecuTrakPath
    ExecuTrakPath = m_ExecuTrakPath
End Property

Public Property Let ExecuTrakPath(value As String)
    m_ExecuTrakPath = value
End Property

Public Property Get DatabaseName() As String
    DatabaseName = m_DatabaseName
End Property

Public Property Let DatabaseName(ByVal value As String)
    m_DatabaseName = value
End Property

Public Property Get UnixDataDirectory() As String
    GetUnixDataDirectory
    UnixDataDirectory = m_UnixDataDirectory
End Property

Public Property Let UnixDataDirectory(ByVal value As String)
    m_UnixDataDirectory = value
End Property

Public Property Get DatabasePath() As String
    DatabasePath = m_DatabasePath
End Property

Public Property Let DatabasePath(ByVal value As String)
    m_DatabasePath = value
End Property

Public Property Get EDIXLATERoot() As String
    GetEDIXLATERoot
    EDIXLATERoot = m_EDIXLATERoot
End Property

Public Property Let EDIXLATERoot(ByVal value As String)
    m_EDIXLATERoot = value
End Property

Public Property Get EDIXLATEFile() As String
    GetEDIXLATEFile
    EDIXLATEFile = m_EDIXLATEFile
End Property

Public Property Let EDIXLATEFile(ByVal value As String)
    m_EDIXLATEFile = value
End Property

Public Property Get EDIXLATEPath() As String
    GetEDIXLATEPath
    EDIXLATEPath = m_EDIXLATEPath
End Property

Public Property Let EDIXLATEPath(ByVal value As String)
    m_EDIXLATEPath = value
End Property

' FUNCTIONS / SUBROUTINES *******************************************

Private Sub GetApplicationPath()
        
    'If Not m_loaded Then Err.Raise -2, , "GetApplicationPath(): " & ERROR_NOTLOADED
    
    If Len(m_ApplicationPath) < 1 Then
        If Len(m_WorkingFolder) > 0 Then
            m_ApplicationPath = UCase(AppendDSN(m_WorkingFolder))
        Else
            m_ApplicationPath = UCase(AppendDSN(App.Path))
        End If
        CreateFolder m_ApplicationPath
    End If
    
End Sub

Private Sub GetExecuTrakPath()
    Dim folder$
        
    'If Not m_loaded Then Err.Raise -2, , "GetExecuTrakPath(): " & ERROR_NOTLOADED
    
    If Len(m_ExecuTrakPath) < 1 Then
        folder = GetExecTrakPath()
    
        m_ExecuTrakPath = AppendDSN(folder)
        
        Call CreateFolder(m_ExecuTrakPath)
    End If
    
End Sub

Private Sub GetExecuTrakEdiDataPath()
    Dim folder$
        
    'If Not m_loaded Then Err.Raise -2, , "GetExecuTrakEdiDataPath(): " & ERROR_NOTLOADED
    
    If Len(m_ExecuTrakEdiDataPath) < 1 Then
        folder = GetExecTrakPath()
    
        folder = RemoveEndSlash(folder) + "\EDI_DATA"
        m_ExecuTrakEdiDataPath = AppendDSN(folder)
        
        Call CreateFolder(m_ExecuTrakEdiDataPath)
    End If
    
End Sub

Private Sub GetWindowsTempPath()
    Dim folder$
    Dim Ret As Long
        
    'If Not m_loaded Then Err.Raise -2, , "GetWindowsTempPath(): " & ERROR_NOTLOADED
        
    If Len(m_WindowsTempPath) < 1 Then
        If Len(m_WorkingFolder) > 0 Then
            m_WindowsTempPath = m_WorkingFolder & "TEMP"
            CreateFolder m_WindowsTempPath
        Else
            folder = String(MAX_PATH, 0)
            Ret = GetTempPath(MAX_PATH, folder)
            If Ret <> 0 Then
                m_WindowsTempPath = AppendSlash(Left(folder, InStr(folder, Chr(0)) - 1))
            Else
                m_WindowsTempPath = Me.LocalTempPath
            End If
        End If
    End If
    
End Sub

Private Sub GetStandardLogPath()
    Dim folder$
        
    'If Not m_loaded Then Err.Raise -2, , "GetStandardLogPath(): " & ERROR_NOTLOADED
    
    If Len(m_StandardLogPath) < 1 Then
        folder = GetExecTrakPath()
    
        folder = RemoveEndSlash(folder) + "\LOG"
        m_StandardLogPath = AppendDSN(folder)
        
        If Not CreateFolder(m_StandardLogPath) Then
            If Len(m_WorkingFolder) < 1 Then
                m_StandardLogPath = Me.LocalLogPath
            End If
        End If
    End If
    
End Sub

Private Sub GetStandardLogFile()
    Dim file$
    Dim hFile%
        
    'If Not m_loaded Then Err.Raise -2, , "GetStandardLogFile(): " & ERROR_NOTLOADED
    
    If Len(m_StandardLogFile) < 1 Then
        GetStandardLogPath
        file = AppendSlash(m_StandardLogPath) & App.Title & ".LOG"
        
        On Error GoTo FINISHED
        hFile = FreeFile()
        
        If FileExists(file) Then
            Open file For Append As #hFile
        Else
            Open file For Output As #hFile
        End If
        
        Close #hFile
        
        m_StandardLogFile = file
    End If
    
    Err.Clear
FINISHED:
    If Err.number <> 0 Then
        m_StandardLogFile = Me.LocalLogFile
        Err.Clear
    End If
End Sub

Private Sub GetLocalAppPath()
    Dim Path$
        
    'If Not m_loaded Then Err.Raise -2, , "GetLocalAppPath(): " & ERROR_NOTLOADED
    
    If Len(m_LocalAppPath) < 1 Then
        If Len(m_WorkingFolder) > 0 Then
            Path = Me.ApplicationPath
            CreateFolder Path
        Else
            Path = AppendDSN("C:\FACTOR\" & App.Title)
            CreateFolder Path
        End If
        m_LocalAppPath = Path
    End If
    
End Sub

Private Sub GetLocalLogPath()
    Dim folder$
        
    'If Not m_loaded Then Err.Raise -2, , "GetLocalLogPath(): " & ERROR_NOTLOADED
    
    If Len(m_LocalLogPath) < 1 Then
        If Len(m_WorkingFolder) > 0 Then
            folder = Me.StandardLogPath
        Else
            folder = RemoveEndSlash(LOCAL_FACTOR_PATH) & "\EXECTRAK\LOG"
            folder = AppendDSN(folder)
        End If
        m_LocalLogPath = folder
        CreateFolder folder
    End If
    
End Sub

Private Sub GetLocalLogFile()
    Dim hFile%
        
    'If Not m_loaded Then Err.Raise -2, , "GetLocalLogFile(): " & ERROR_NOTLOADED
    
    If Len(m_LocalLogFile) < 1 Then
        GetLocalLogPath
        m_LocalLogFile = AppendSlash(LocalLogPath) & App.Title & ".LOG"
        If Not FileExists(m_LocalLogFile) Then
            hFile = FreeFile()
            Open m_LocalLogFile For Output As #hFile
            Close hFile
        End If
    End If
    
End Sub

Private Sub GetLocalTempPath()
    Dim folder$
        
    'If Not m_loaded Then Err.Raise -2, , "GetLocalTempPath(): " & ERROR_NOTLOADED
    
    If Len(m_LocalTempPath) < 1 Then
        If Len(m_WorkingFolder) > 0 Then
            folder = m_WorkingFolder & "TEMP"
        Else
            folder = AppendSlash(LOCAL_FACTOR_PATH) & "TEMP\"
        End If
        m_LocalTempPath = AppendDSN(folder)
        CreateFolder m_LocalTempPath
    End If
    
End Sub

Private Sub GetFuelMovementUploadPath()
    Dim folder$
        
    'If Not m_loaded Then Err.Raise -2, , "GetFuelMovementUploadPath(): " & ERROR_NOTLOADED
    
    If Len(m_FuelMovementUploadPath) < 1 Then
        If Len(m_WorkingFolder) > 0 Then
            folder = m_WorkingFolder & "FUEL-MVMNT-UPLD"
        Else
            folder = AppendSlash(LOCAL_FACTOR_PATH) & "FUEL-MVMNT-UPLD"
        End If
        m_FuelMovementUploadPath = AppendDSN(folder)
    End If

End Sub

Private Sub GetXMLDataPath()
    Dim folder$
        
    'If Not m_loaded Then Err.Raise -2, , "GetXMLDataPath(): " & ERROR_NOTLOADED
    
    If Len(m_XMLDataPath) < 1 Then
        folder = GetExecTrakPath()
        folder = folder & "XML_DATA"
        m_XMLDataPath = AppendDSN(folder)
    End If
    
End Sub

Private Sub GetStandardEmailPath()
    Dim folder$
        
    'If Not m_loaded Then Err.Raise -2, , "GetStandardEmailPath(): " & ERROR_NOTLOADED
    
    If Len(m_StandardEmailPath) < 1 Then
        folder = GetExecTrakPath()
        folder = RemoveEndSlash(folder) & "\EMAIL\"
        folder = AppendDSN(folder)
        If Not CreateFolder(folder) Then folder = Me.LocalEmailPath
        m_StandardEmailPath = folder
    End If
    
End Sub

Private Sub GetLocalEmailPath()
    Dim folder$
        
    'If Not m_loaded Then Err.Raise -2, , "GetLocalEmailPath(): " & ERROR_NOTLOADED
    
    If Len(m_LocalEmailPath) < 1 Then
        If Len(m_WorkingFolder) > 0 Then
            folder = m_WorkingFolder & "EMAIL"
        Else
            folder = AppendSlash(LOCAL_FACTOR_PATH) & "EMAIL"
        End If
        m_LocalEmailPath = AppendDSN(folder)
        
        CreateFolder m_LocalEmailPath
    End If
    
End Sub

Private Sub GetFactorConfigurationFile()
        
    'If Not m_loaded Then Err.Raise -2, , "GetFactorConfigurationFile(): " & ERROR_NOTLOADED
    
    If Len(m_FactorConfigurationFile) < 1 Then
        m_FactorConfigurationFile = NewIniPath("C:\FACTOR\FACTOR.INI")
    End If
    
End Sub

Private Sub GetFactorEDIDirectory()
    Dim folder$
        
    'If Not m_loaded Then Err.Raise -2, , "GetFactorEDIDirectory(): " & ERROR_NOTLOADED
    
    If Len(m_FactorEDIDirectory) < 1 Then
        If Len(m_WorkingFolder) > 0 Then
            folder = Me.GetFactorINIValue("EDI", "Factor EDI Directory", m_WorkingFolder & "EDI")
        Else
            folder = Me.GetFactorINIValue("EDI", "Factor EDI Directory", App.Path)
            'folder = Me.GetFactorINIValue("EDI", "Factor EDI Directory", "C:\FACTOR\EXECTRAK\BIN")
        End If
        m_FactorEDIDirectory = AppendDSN(folder)
    End If
    
End Sub

Private Sub GetDTNLogFilePath()
    Dim folder$
        
    'If Not m_loaded Then Err.Raise -2, , "GetDTNLogFilePath(): " & ERROR_NOTLOADED
    
    If Len(m_DTNLogFilePath) < 1 Then
        If Len(m_WorkingFolder) > 0 Then
            folder = Me.GetFactorINIValue("DTN_MSG", "Log File Path", m_WorkingFolder & "DTN\LOG")
        Else
            folder = Me.GetFactorINIValue("DTN_MSG", "Log File Path", AppendSlash(LOCAL_FACTOR_PATH) & "LOG")
        End If
        m_DTNLogFilePath = folder
    End If
    
End Sub

Private Sub GetDTNInputPath()
    Dim folder$
        
    'If Not m_loaded Then Err.Raise -2, , "GetDTNInputPath(): " & ERROR_NOTLOADED
    
    If Len(m_DTNInputPath) < 1 Then
        If Len(m_WorkingFolder) > 0 Then
            folder = Me.GetFactorINIValue("DTN_MSG", "Input Path", m_WorkingFolder & "DTN")
        Else
            folder = Me.GetFactorINIValue("DTN_MSG", "Input Path", "C:\FACTOR\DTN")
        End If
        m_DTNInputPath = folder
    End If
    
End Sub

Private Sub GetTLCInboundDirectory()
    Dim folder$
        
    'If Not m_loaded Then Err.Raise -2, , "GetTLCInboundDirectory(): " & ERROR_NOTLOADED
    
    If Len(m_TLCInboundDirectory) < 1 Then
        folder = Me.GetFactorINIValue("EDI", "TLC Inbound Directory")
        'TODO
    End If
    
End Sub

Private Sub GetUnixDataDirectory()
    Dim arr$()
    Dim x%
        
    'If Not m_loaded Then Err.Raise -2, , "GetUnixDataDirectory(): " & ERROR_NOTLOADED
    
    If Len(m_UnixDataDirectory) < 1 Then
        
        m_UnixDataDirectory = "/factor/"
        
        If Len(m_DatabasePath) > 0 And Len(WorkingFolder) > 0 Then
            arr = Split(m_DatabasePath, "/")
            x = UBound(arr)
            If x > 0 Then m_UnixDataDirectory = "/" & arr(1) & "/"
            If x > 1 Then m_UnixDataDirectory = m_UnixDataDirectory & arr(2) & "/"
        End If
        
    End If
    
End Sub

Private Sub GetEDIXLATERoot()
    Dim folder$
        
    'If Not m_loaded Then Err.Raise -2, , "GetEDIXLATERoot(): " & ERROR_NOTLOADED
    
    If Len(m_EDIXLATERoot) < 1 Then
        If Len(WorkingFolder) < 1 Then
            folder = UCase(AppendSlash(App.Path))
        Else
            folder = UCase(WorkingFolder & "EDIXLATE")
        End If
        Call CreateFolder(folder)
        m_EDIXLATERoot = AppendSlash(folder)
    End If
    
End Sub

Private Sub GetEDIXLATEPath()
    Dim folder$
        
    'If Not m_loaded Then Err.Raise -2, , "GetEDIXLATEPath(): " & ERROR_NOTLOADED
    
    If Len(m_EDIXLATEPath) < 1 Then
        folder = EDIXLATERoot
        If Len(WorkingFolder) > 0 Then
            If Len(DatabaseName) > 0 Then
                folder = UCase(folder & DatabaseName)
            Else
                folder = UCase(folder & "UNKNOWNDB")
            End If
        End If
        Call CreateFolder(folder)
        m_EDIXLATEPath = AppendSlash(folder)
    End If

End Sub

Private Sub GetEDIXLATEFile()
    Dim folder$
        
    'If Not m_loaded Then Err.Raise -2, , "GetEDIXLATEFile(): " & ERROR_NOTLOADED
    
    If Len(m_EDIXLATEFile) < 1 Then
        folder = UCase(AppendSlash(EDIXLATEPath))
        Call CreateFolder(folder)
        m_EDIXLATEFile = AppendSlash(folder) & "EDIXLATE.MDB"
    End If

End Sub

'********************************************************************
'*                     ***  UTILITIES ***                           *
'********************************************************************

Private Function AppendDSN(ByVal Path$) As String
    
    If Len(WorkingFolder) > 0 Then
        AppendDSN = AppendDSNToPath(Path)
    Else
        AppendDSN = AppendSlash(Path)
    End If
    
End Function

Public Function AppendDSNToPath(ByVal Path$) As String
    
    Path = AppendSlash(Path) + UCase$(m_sDSN)
    AppendDSNToPath = AppendSlash(Path)
    
End Function

Public Function CreateFolder(ByVal Path$) As Boolean
    Dim wfd As WIN32_FIND_DATA
    Dim hSearch As Long
    Dim arr() As String
    Dim folder$
    Dim i%
    
    On Error GoTo FINISHED
        
    If Len(Path) < 1 Then GoTo FINISHED
    
    If Not DirectoryExists(Path) Then
        ' Network shares are a pain
        If Left$(Path, 2) = "\\" Then
            CreateFolder = CreateNetworkFolder(Path)
            Exit Function
        End If
        
        ' Nice normal letter drive
        Path = RemoveEndSlash(Path)
        arr = Split(Path, "\")
        For i = 0 To UBound(arr)
            folder = IIf(i > 0, folder & "\" & arr(i), arr(i))
            If Len(arr(i)) > 0 Then
                If Not DirectoryExists(folder) Then
                    MkDir folder
                End If
            End If
        Next
        CreateFolder = True
    Else
        CreateFolder = True
    End If
    
FINISHED:
    Err.Clear
    
End Function

Private Function CreateNetworkFolder(ByVal Path$) As Boolean
    Dim share$
    Dim folder$
    Dim arr() As String
    Dim i%
    
    On Error GoTo FINISHED
    
    Path = RemoveEndSlash(Path)
    
    share = FindNetworkShare(Path)
    If Len(share) > 0 Then
        If share = Path Then
            CreateNetworkFolder = True
            Exit Function
        End If
        
        folder = share
        Path = Mid(Path, Len(share) + 1)
        If Left(Path, 1) = "\" Then Path = Mid(Path, 2)
        
        arr = Split(Path, "\")
        For i = 0 To UBound(arr)
            folder = folder & "\" & arr(i)
            If Len(arr(i)) > 0 Then
                If Not DirectoryExists(folder) Then
                    MkDir folder
                End If
            End If
        Next
        CreateNetworkFolder = True
    End If
    
FINISHED:
    Err.Clear
    
End Function

Public Function FileExists(ByVal file$) As Boolean
    Dim hSearch As Long
    Dim wfd As WIN32_FIND_DATA
    
    If Len(file) > 0 Then
        hSearch = FindFirstFile(file, wfd)
        If hSearch <> INVALID_HANDLE_VALUE Then
            If Not wfd.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
                FileExists = True
            End If
            FindClose hSearch
        End If
    End If
    
End Function

Public Function DirectoryExists(ByVal folder$) As Boolean
    Dim attr As Long
    
    If Len(folder) > 0 Then
        folder = RemoveEndSlash(folder)
        attr = GetFileAttributes(folder)
        If (attr <> INVALID_HANDLE_VALUE) Then
            If (attr And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
                DirectoryExists = True
            End If
        End If
    End If
    
End Function

Private Function FindNetworkShare(ByVal Path$) As String
    Dim arr() As String
    Dim folder$
    Dim i%
    Dim hSearch As Long
    Dim wfd As WIN32_FIND_DATA
    
    If Left$(Path, 2) = "\\" Then
        Path = Mid(Path, 3)
        arr = Split(Path, "\")
        folder = "\"
        For i = 0 To UBound(arr)
            folder = folder + "\" + arr(i)
            hSearch = FindFirstFile(folder + "\*", wfd)
            If hSearch <> INVALID_HANDLE_VALUE Then
                FindNetworkShare = folder
                FindClose hSearch
                Exit Function
            End If
        Next
    End If
    
End Function

Private Function AppendSlash(ByVal Path$) As String
        
    AppendSlash = IIf(Right$(Path, 1) <> "\", Path + "\", Path)
    
End Function

Public Function RemoveEndSlash(ByVal Path$) As String

    RemoveEndSlash = IIf(Right$(Path, 1) = "\", Mid(Path, 1, Len(Path) - 1), Path)

End Function

Private Function TrimNull(ByVal value As String) As String
    Dim i As Long
    
    i = InStr(1, value, vbNullChar)
    If i < 1 Then
        TrimNull = value
    Else
        TrimNull = Left$(value, i - 1)
    End If
    
End Function

Private Function GetExecTrakPath() As String
    Dim Path$
    Dim folder$
    Dim x%
    
    If Len(m_WorkingFolder) > 0 Then
        folder = m_WorkingFolder
    Else
        Path = App.Path
        'path = "C:\FACTOR\EXECTRAK\BIN"
        
        x = InStr(1, Path, "\EXECTRAK\", vbTextCompare)
        If x > 0 Then
            folder = Mid(Path, 1, x - 1)
        Else
            folder = Path
        End If
    End If
        
    GetExecTrakPath = AppendSlash(folder) + "EXECTRAK\"
    CreateFolder GetExecTrakPath

End Function

Public Function NewIniPath(iniFile$)
    Dim arr$()
    Dim Name$
    Dim x%
    
    NewIniPath = iniFile
    
    If Len(m_WorkingFolder) > 0 Then
        If Len(iniFile) > 0 Then
            If UCase(Right$(iniFile, 4)) = ".INI" Then
                arr = Split(iniFile, "\")
                x = UBound(arr)
                Name = UCase(arr(x))
                Select Case Name
                    Case "FACTOR.INI", "IQFACT.INI"
                        NewIniPath = m_WorkingFolder & UCase(m_sDSN) & "\" & Name
                    Case "ETETFSETUP.INI"
                        If Len(DatabaseName) > 0 Then
                            NewIniPath = EDIXLATEPath & Name
                        End If
                    Case Else
                        NewIniPath = m_WorkingFolder & Name
                End Select
            End If
        End If
    End If
    
End Function

Public Function GetFactorINIValue(section$, key$, Optional defaultValue$ = "")
    
    GetFactorINIValue = ReadINIString(Me.FactorConfigurationFile, section, key, defaultValue)
    
End Function

Public Function GetDirectory(filePath As String) As String
    Dim parts$()
    Dim Path$
    Dim x&
    Dim i&
    Dim network As Boolean
    
    Path = filePath
    If Len(Path) > 0 Then
        If Left(Path, 2) = "\\" Then
            network = True
            Path = Mid(Path, 3)
        End If
        
        parts = Split(Path, "\")
        x = UBound(parts)
        If x > 0 Then
            GetDirectory = parts(0)
            For i = 1 To x - 1
                GetDirectory = GetDirectory + "\" + parts(i)
            Next
        Else
            GetDirectory = Path
        End If
        
        If network Then GetDirectory = "\\" + GetDirectory
    End If
    
End Function

Public Function ReadINIString(configFile$, section$, key$, Optional defaultValue$ = "")
    Dim Ret As Long
    Dim value$
    Dim file$
    
    file = NewIniPath(configFile)
    
    value = Space$(256)
    Ret = GetPrivateProfileString(section, key, defaultValue, value, Len(value), file)
    If Ret > 0 Then
        ReadINIString = Left$(value, Ret)
    Else
        ReadINIString = defaultValue
    End If
    
End Function

Public Function WriteINIString(configFile$, section$, key$, value$) As Long
    Dim file$
    
    file = NewIniPath(configFile)
    
    WriteINIString = WritePrivateProfileString(section, key, value, file)
    
End Function

Public Function GetFileCreationDate(file As String, fileDate As Date, Optional toLocalTime As Boolean = True) As Boolean
    Dim createDate As FILETIME
    Dim lastRead As FILETIME
    Dim lastWrite As FILETIME
    
    If GetFileDates(file, createDate, lastRead, lastWrite, toLocalTime) Then
        fileDate = FileTimeToDate(createDate)
        GetFileCreationDate = True
    End If
    
End Function

'********************************************************************
' http://www.vb-helper.com/howto_file_times.html
'********************************************************************
Private Function GetFileDates(file As String, createDate As FILETIME, lastRead As FILETIME, lastWrite As FILETIME, Optional toLocalTime As Boolean = True) As Boolean
    Dim hFile As Long
    Dim file_time As FILETIME
    Dim sa As SECURITY_ATTRIBUTES
    
    hFile = CreateFile(file, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, sa, OPEN_EXISTING, 0&, 0&)
    If hFile = 0 Then Exit Function

    If GetFileTime(hFile, createDate, lastRead, lastWrite) = 0 Then Exit Function
    
    Call CloseHandle(hFile)
    
    If toLocalTime Then
        ' Convert to local file system time.
        FileTimeToLocalFileTime createDate, file_time
        createDate = file_time

        FileTimeToLocalFileTime lastRead, file_time
        lastRead = file_time

        FileTimeToLocalFileTime lastWrite, file_time
        lastWrite = file_time
    End If
    
    GetFileDates = True
    
End Function

'********************************************************************
' http://www.vb-helper.com/howto_file_times.html
'********************************************************************
Private Function FileTimeToDate(ft As FILETIME) As Date
    Dim lo_time As Double
    Dim hi_time As Double
    Dim seconds As Double
    Dim hours As Double
    Dim the_date As Date

    ' Get the low order data.
    If ft.dwLowDateTime < 0 Then
        lo_time = 2 ^ 31 + (ft.dwLowDateTime And &H7FFFFFFF)
    Else
        lo_time = ft.dwLowDateTime
    End If

    ' Get the high order data.
    If ft.dwHighDateTime < 0 Then
        hi_time = 2 ^ 31 + (ft.dwHighDateTime And &H7FFFFFFF)
    Else
        hi_time = ft.dwHighDateTime
    End If

    ' Combine them and turn the result into hours.
    seconds = (lo_time + 2 ^ 32 * hi_time) / TICKS_PER_SECOND
    hours = CLng(seconds / 3600)
    seconds = seconds - hours * 3600

    ' Make the date.
    the_date = DateAdd("h", hours, "1/1/1601 0:00 AM")
    the_date = DateAdd("s", seconds, the_date)
    FileTimeToDate = the_date

End Function

'********************************************************************
' Connecting to database
'********************************************************************

Public Function GetConnectionString() As String
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
    
    sODBCKey = m_sODBC_INI_Path & sDSN
    
    If Len(m_sUID) < 1 Or Len(m_sPWD) < 1 Then
        Call LoadDSNSettings(m_sDSN, m_sUID, m_sPWD)
    End If
    
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

Private Function GetWorkingFolder(db As DataBase) As String
    Dim SQL$
    Dim parm15$
    Dim folder$
    
    SQL = "select parm_field from sys_parm where parm_nbr = 15"
    With db.OpenRecordset(SQL, dbOpenSnapshot, dbSQLPassThrough)
        If Not .EOF Then
            parm15 = UCase$(Trim$(.Fields(0).value & ""))
        End If
        .Close
    End With
    
    m_MultiDatabase = IIf(parm15 = "Y", True, False)
    
    SQL = "select ini_value from sys_ini where ini_section = 'WorkingDirectory' AND ini_file_name = 'WORKPATH'"
    If m_MultiDatabase Then
        With db.OpenRecordset(SQL, dbOpenSnapshot, dbSQLPassThrough)
            If Not .EOF Then
                folder = UCase$(Trim$(.Fields(0).value & ""))
            End If
            .Close
        End With
                
        If Len(folder) > 0 Then
            If DirectoryExists(folder) Then
                m_WorkingFolder = AppendSlash(folder)
            End If
        End If
        
        If Len(m_WorkingFolder) < 1 Then m_MultiDatabase = False
    End If
    
    GetWorkingFolder = m_WorkingFolder
    
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

Private Function StripDatabaseName(sDBPath As String) As String
    Dim sDBName As String
    Dim i As Integer
    
    i = InStrRev(sDBPath, "/")
    
    If i > 1 Then
        sDBPath = Left(sDBPath, i - 1)
    
        i = InStrRev(sDBPath, "/")
    
        If i > 1 Then
            sDBName = Mid(sDBPath, i + 1)
        End If
    Else
        sDBName = sDBPath
    End If

    StripDatabaseName = sDBName

End Function

Public Function GetDatabasePath(connectionString As String) As String
    
    GetDatabasePath = GetNamedString(connectionString, ";DB")
    If Trim(GetDatabasePath) = "" Then
        GetDatabasePath = GetNamedString(connectionString, "DATABASE")
    End If

End Function

Public Sub LoadDSNSettings(ByVal dsnName As String, ByRef userName As String, ByRef password As String)
    Dim sBuffer As String
    Dim Ret As Long
    Dim hToken As Long
    Dim file As String
    Dim encrypted$
    
    userName = ""
    password = ""
    
    sBuffer = String(MAX_PATH, 0)
    OpenProcessToken GetCurrentProcess, TOKEN_QUERY, hToken
    GetUserProfileDirectory hToken, sBuffer, MAX_PATH
    file = TrimNull(sBuffer)
    
    If Len(file) > 0 Then
        file = AppendSlash(file) & "factorsettings.ini"
        Ret = GetPrivateProfileString(dsnName, "right", "", sBuffer, MAX_PATH, file)
        If Ret > 0 Then
            userName = TrimNull(sBuffer)
        End If
        Ret = GetPrivateProfileString(dsnName, "left", "", sBuffer, MAX_PATH, file)
        password = DecryptText(TrimNull(sBuffer), ENCRYPT_KEY)
    End If
    
    If Len(userName) < 1 Then userName = "nousername"
    If Len(password) < 1 Then password = "password"

End Sub

Public Sub SaveDSNSettings(ByVal dsnName As String, ByVal userName As String, ByVal password As String)
    Dim sBuffer As String
    Dim Ret As Long
    Dim hToken As Long
    Dim file As String
    Dim encrypted$
    
    sBuffer = String(MAX_PATH, 0)
    OpenProcessToken GetCurrentProcess, TOKEN_QUERY, hToken
    GetUserProfileDirectory hToken, sBuffer, MAX_PATH
    file = TrimNull(sBuffer)
    
    If Len(file) > 0 Then
        file = AppendSlash(file) & "factorsettings.ini"
        encrypted = EncryptText(password, ENCRYPT_KEY)
        WritePrivateProfileString dsnName, "right", userName, file
        WritePrivateProfileString dsnName, "left", encrypted, file
    End If
    
End Sub

'********************************************************************
' ENCRYPTION ROUTINES
' http://www.developerfusion.com/code/157/encryptdecrypt/
'********************************************************************

Public Function EncryptText(strText As String, ByVal strPwd As String)
    Dim i As Integer, c As Integer
    Dim strBuff As String

    'Encrypt string
    If Len(strPwd) Then
        For i = 1 To Len(strText)
            c = Asc(Mid$(strText, i, 1))
            c = c + Asc(Mid$(strPwd, (i Mod Len(strPwd)) + 1, 1))
            strBuff = strBuff & Chr$(c And &HFF)
        Next i
    Else
        strBuff = strText
    End If
    
    EncryptText = strBuff

End Function

Public Function DecryptText(strText As String, ByVal strPwd As String)
    Dim i As Integer, c As Integer
    Dim strBuff As String

    'Decrypt string
    If Len(strPwd) Then
        For i = 1 To Len(strText)
            c = Asc(Mid$(strText, i, 1))
            c = c - Asc(Mid$(strPwd, (i Mod Len(strPwd)) + 1, 1))
            strBuff = strBuff & Chr$(c And &HFF)
        Next i
    Else
        strBuff = strText
    End If
    
    DecryptText = strBuff

End Function
