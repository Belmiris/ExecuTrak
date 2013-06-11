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

Private m_HostedFolder$
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

Private Const LOCAL_FACTOR_PATH = "C:\FACTOR\"

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

' PROPERTIES ********************************************************

Public Property Get HostedFolder() As String
    HostedFolder = m_HostedFolder
End Property

Public Property Let HostedFolder(ByVal value As String)
    If DirectoryExists(value) Then
        m_HostedFolder = AppendSlash(value)
    End If
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
    
    If Len(m_ApplicationPath) < 1 Then
        If Len(m_HostedFolder) > 0 Then
            m_ApplicationPath = UCase(AppendDSN(m_HostedFolder))
        Else
            m_ApplicationPath = UCase(AppendDSN(App.path))
        End If
        CreateFolder m_ApplicationPath
    End If
    
End Sub

Private Sub GetExecuTrakPath()
    Dim folder$
    
    If Len(m_ExecuTrakPath) < 1 Then
        folder = GetExecTrakPath()
    
        m_ExecuTrakPath = AppendDSN(folder)
        
        Call CreateFolder(m_ExecuTrakPath)
    End If
    
End Sub

Private Sub GetExecuTrakEdiDataPath()
    Dim folder$
    
    If Len(m_ExecuTrakEdiDataPath) < 1 Then
        folder = GetExecTrakPath()
    
        folder = RemoveEndSlash(folder) + "\EDI_DATA"
        m_ExecuTrakEdiDataPath = AppendDSN(folder)
        
        Call CreateFolder(m_ExecuTrakEdiDataPath)
    End If
    
End Sub

Private Sub GetWindowsTempPath()
    Dim folder$
    Dim ret As Long
        
    If Len(m_WindowsTempPath) < 1 Then
        If Len(m_HostedFolder) > 0 Then
            m_WindowsTempPath = m_HostedFolder & "TEMP"
            CreateFolder m_WindowsTempPath
        Else
            folder = String(MAX_PATH, 0)
            ret = GetTempPath(MAX_PATH, folder)
            If ret <> 0 Then
                m_WindowsTempPath = AppendSlash(Left(folder, InStr(folder, Chr(0)) - 1))
            Else
                m_WindowsTempPath = Me.LocalTempPath
            End If
        End If
    End If
    
End Sub

Private Sub GetStandardLogPath()
    Dim folder$
    
    If Len(m_StandardLogPath) < 1 Then
        folder = GetExecTrakPath()
    
        folder = RemoveEndSlash(folder) + "\LOG"
        m_StandardLogPath = AppendDSN(folder)
        
        If Not CreateFolder(m_StandardLogPath) Then
            If Len(m_HostedFolder) < 1 Then
                m_StandardLogPath = Me.LocalLogPath
            End If
        End If
    End If
    
End Sub

Private Sub GetStandardLogFile()
    Dim file$
    Dim hFile%
    
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
    If Err.Number <> 0 Then
        m_StandardLogFile = Me.LocalLogFile
        Err.Clear
    End If
End Sub

Private Sub GetLocalAppPath()
    Dim path$
    
    If Len(m_LocalAppPath) < 1 Then
        If Len(m_HostedFolder) > 0 Then
            path = Me.ApplicationPath
            CreateFolder path
        Else
            path = AppendDSN("C:\FACTOR\" & App.Title)
            CreateFolder path
        End If
        m_LocalAppPath = path
    End If
    
End Sub

Private Sub GetLocalLogPath()
    Dim folder$
    
    If Len(m_LocalLogPath) < 1 Then
        If Len(m_HostedFolder) > 0 Then
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
    
    If Len(m_LocalTempPath) < 1 Then
        If Len(m_HostedFolder) > 0 Then
            folder = m_HostedFolder & "TEMP"
        Else
            folder = AppendSlash(LOCAL_FACTOR_PATH) & "TEMP\"
        End If
        m_LocalTempPath = AppendDSN(folder)
        CreateFolder m_LocalTempPath
    End If
    
End Sub

Private Sub GetFuelMovementUploadPath()
    Dim folder$
    
    If Len(m_FuelMovementUploadPath) < 1 Then
        If Len(m_HostedFolder) > 0 Then
            folder = m_HostedFolder & "FUEL-MVMNT-UPLD"
        Else
            folder = AppendSlash(LOCAL_FACTOR_PATH) & "FUEL-MVMNT-UPLD"
        End If
        m_FuelMovementUploadPath = AppendDSN(folder)
    End If

End Sub

Private Sub GetXMLDataPath()
    Dim folder$
    
    If Len(m_XMLDataPath) < 1 Then
        folder = GetExecTrakPath()
        folder = folder & "XML_DATA"
        m_XMLDataPath = AppendDSN(folder)
    End If
    
End Sub

Private Sub GetStandardEmailPath()
    Dim folder$
    
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
    
    If Len(m_LocalEmailPath) < 1 Then
        If Len(m_HostedFolder) > 0 Then
            folder = m_HostedFolder & "EMAIL"
        Else
            folder = AppendSlash(LOCAL_FACTOR_PATH) & "EMAIL"
        End If
        m_LocalEmailPath = AppendDSN(folder)
        
        CreateFolder m_LocalEmailPath
    End If
    
End Sub

Private Sub GetFactorConfigurationFile()
    
    If Len(m_FactorConfigurationFile) < 1 Then
        m_FactorConfigurationFile = NewIniPath("C:\FACTOR\FACTOR.INI")
    End If
    
End Sub

Private Sub GetFactorEDIDirectory()
    Dim folder$
    
    If Len(m_FactorEDIDirectory) < 1 Then
        If Len(m_HostedFolder) > 0 Then
            folder = Me.GetFactorINIValue("EDI", "Factor EDI Directory", m_HostedFolder & "EDI")
        Else
            folder = Me.GetFactorINIValue("EDI", "Factor EDI Directory", App.path)
            'folder = Me.GetFactorINIValue("EDI", "Factor EDI Directory", "C:\FACTOR\EXECTRAK\BIN")
        End If
        m_FactorEDIDirectory = AppendDSN(folder)
    End If
    
End Sub

Private Sub GetDTNLogFilePath()
    Dim folder$
    
    If Len(m_DTNLogFilePath) < 1 Then
        If Len(m_HostedFolder) > 0 Then
            folder = Me.GetFactorINIValue("DTN_MSG", "Log File Path", m_HostedFolder & "DTN\LOG")
        Else
            folder = Me.GetFactorINIValue("DTN_MSG", "Log File Path", AppendSlash(LOCAL_FACTOR_PATH) & "LOG")
        End If
        m_DTNLogFilePath = folder
    End If
    
End Sub

Private Sub GetDTNInputPath()
    Dim folder$
    
    If Len(m_DTNInputPath) < 1 Then
        If Len(m_HostedFolder) > 0 Then
            folder = Me.GetFactorINIValue("DTN_MSG", "Input Path", m_HostedFolder & "DTN")
        Else
            folder = Me.GetFactorINIValue("DTN_MSG", "Input Path", "C:\FACTOR\DTN")
        End If
        m_DTNInputPath = folder
    End If
    
End Sub

Private Sub GetTLCInboundDirectory()
    Dim folder$
    
    If Len(m_TLCInboundDirectory) < 1 Then
        folder = Me.GetFactorINIValue("EDI", "TLC Inbound Directory")
        'TODO
    End If
    
End Sub

Private Sub GetUnixDataDirectory()
    Dim arr$()
    Dim x%
    
    If Len(m_UnixDataDirectory) < 1 Then
        
        m_UnixDataDirectory = "/factor/"
        
        If Len(m_DatabasePath) > 0 Then
            arr = Split(m_DatabasePath, "/")
            x = UBound(arr)
            If x > 0 Then m_UnixDataDirectory = "/" & arr(1) & "/"
            If x > 1 Then m_UnixDataDirectory = m_UnixDataDirectory & arr(2) & "/"
        End If
        
    End If
    
End Sub

Private Sub GetEDIXLATERoot()
    Dim folder$
    
    If Len(m_EDIXLATERoot) < 1 Then
        If Len(HostedFolder) < 1 Then
            folder = UCase(AppendSlash(App.path) & "EDIXLATE")
        Else
            folder = UCase(HostedFolder & "EDIXLATE")
        End If
        Call CreateFolder(folder)
        m_EDIXLATERoot = AppendSlash(folder)
    End If
    
End Sub

Private Sub GetEDIXLATEPath()
    Dim folder$
    
    If Len(m_EDIXLATEPath) < 1 Then
        folder = EDIXLATERoot
        If Len(DatabaseName) > 0 Then
            folder = UCase(folder & DatabaseName)
        Else
            folder = UCase(folder & "UNKNOWNDB")
        End If
        Call CreateFolder(folder)
        m_EDIXLATEPath = AppendSlash(folder)
    End If

End Sub

Private Sub GetEDIXLATEFile()
    Dim folder$
    
    If Len(m_EDIXLATEFile) < 1 Then
        folder = UCase(AppendSlash(EDIXLATEPath))
        Call CreateFolder(folder)
        m_EDIXLATEFile = AppendSlash(folder) & "EDIXLATE.MDB"
    End If

End Sub

'********************************************************************
'*                     ***  UTILITIES ***                           *
'********************************************************************

Public Function AppendDSN(ByVal path$) As String
    
    path = AppendSlash(path) + UCase$(dsnName)
    AppendDSN = AppendSlash(path)
    
End Function

Public Function CreateFolder(ByVal path$) As Boolean
    Dim wfd As WIN32_FIND_DATA
    Dim hSearch As Long
    Dim arr() As String
    Dim folder$
    Dim i%
    
    On Error GoTo FINISHED
        
    If Len(path) < 1 Then GoTo FINISHED
    
    If Not DirectoryExists(path) Then
        ' Network shares are a pain
        If Left$(path, 2) = "\\" Then
            CreateFolder = CreateNetworkFolder(path)
            Exit Function
        End If
        
        ' Nice normal letter drive
        path = RemoveEndSlash(path)
        arr = Split(path, "\")
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

Private Function CreateNetworkFolder(ByVal path$) As Boolean
    Dim share$
    Dim folder$
    Dim arr() As String
    Dim i%
    
    On Error GoTo FINISHED
    
    path = RemoveEndSlash(path)
    
    share = FindNetworkShare(path)
    If Len(share) > 0 Then
        If share = path Then
            CreateNetworkFolder = True
            Exit Function
        End If
        
        folder = share
        path = Mid(path, Len(share) + 1)
        If Left(path, 1) = "\" Then path = Mid(path, 2)
        
        arr = Split(path, "\")
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

Private Function FindNetworkShare(ByVal path$) As String
    Dim arr() As String
    Dim folder$
    Dim i%
    Dim hSearch As Long
    Dim wfd As WIN32_FIND_DATA
    
    If Left$(path, 2) = "\\" Then
        path = Mid(path, 3)
        arr = Split(path, "\")
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

Private Function AppendSlash(ByVal path$) As String
        
    AppendSlash = IIf(Right$(path, 1) <> "\", path + "\", path)
    
End Function

Public Function RemoveEndSlash(ByVal path$) As String

    RemoveEndSlash = IIf(Right$(path, 1) = "\", Mid(path, 1, Len(path) - 1), path)

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
    Dim path$
    Dim folder$
    Dim x%
    
    If Len(m_HostedFolder) > 0 Then
        folder = m_HostedFolder
    Else
        path = App.path
        'path = "C:\FACTOR\EXECTRAK\BIN"
        
        x = InStr(1, path, "\EXECTRAK\", vbTextCompare)
        If x > 0 Then
            folder = Mid(path, 1, x - 1)
        Else
            folder = path
        End If
    End If
        
    GetExecTrakPath = AppendSlash(folder) + "EXECTRAK\"
    CreateFolder GetExecTrakPath

End Function

Public Function NewIniPath(iniFile$)
    Dim arr$()
    Dim name$
    Dim x%
    
    NewIniPath = iniFile
    
    If Len(m_HostedFolder) > 0 Then
        If Len(iniFile) > 0 Then
            arr = Split(iniFile, "\")
            x = UBound(arr)
            name = UCase(arr(x))
            Select Case name
                Case "FACTOR.INI", "IQFACT.INI"
                    NewIniPath = m_HostedFolder & UCase(dsnName) & "\" & name
                Case "ETETFSETUP.INI"
                    If Len(DatabaseName) > 0 Then
                        NewIniPath = EDIXLATEPath & name
                    End If
                Case Else
                    NewIniPath = m_HostedFolder & name
            End Select
        End If
    End If
    
End Function

Public Function GetFactorINIValue(section$, key$, Optional defaultValue$ = "")
    
    GetFactorINIValue = ReadINIString(Me.FactorConfigurationFile, section, key, defaultValue)
    
End Function

Public Function GetDirectory(filePath As String) As String
    Dim parts$()
    Dim path$
    Dim x&
    Dim i&
    Dim network As Boolean
    
    path = filePath
    If Len(path) > 0 Then
        If Left(path, 2) = "\\" Then
            network = True
            path = Mid(path, 3)
        End If
        
        parts = Split(path, "\")
        x = UBound(parts)
        If x > 0 Then
            GetDirectory = parts(0)
            For i = 1 To x - 1
                GetDirectory = GetDirectory + "\" + parts(i)
            Next
        Else
            GetDirectory = path
        End If
        
        If network Then GetDirectory = "\\" + GetDirectory
    End If
    
End Function

Public Function ReadINIString(configFile$, section$, key$, Optional defaultValue$ = "")
    Dim ret As Long
    Dim value$
    Dim file$
    
    file = NewIniPath(configFile)
    
    value = Space$(256)
    ret = GetPrivateProfileString(section, key, defaultValue, value, Len(value), file)
    If ret > 0 Then
        ReadINIString = Left$(value, ret)
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

