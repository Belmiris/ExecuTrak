Attribute VB_Name = "FTP"
Option Explicit

'*******************************************************************************************
'Created    : 26th September 2000                                                          *
'Created By : Rajneesh Aggarwal                                                            *
'DLL Used   : WININET.DLL                                                                  *
'Usage      : fnTransferFile("Server", "UserID", "PassWord", True, True, _                 *
'           :                "RemoteFilePath&Name", "LocalFilePath", "ErrorMessage", True) *
'*******************************************************************************************
'G8696 - Tthompson
'G8829 - tthompson

Declare Function GetProcessHeap Lib "kernel32" () As Long
Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Declare Function IsCharAlphaW Lib "user32" (ByVal cChar As Integer) As Long

Private Const HEAP_ZERO_MEMORY = &H8
Private Const HEAP_GENERATE_EXCEPTIONS = &H4

Declare Sub CopyMemory1 Lib "kernel32" Alias "RtlMoveMemory" ( _
         hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Declare Sub CopyMemory2 Lib "kernel32" Alias "RtlMoveMemory" ( _
         hpvDest As Long, hpvSource As Any, ByVal cbCopy As Long)

Private Const MAX_PATH = 260
Private Const NO_ERROR = 0
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_OFFLINE = &H1000


Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Type WIN32_FIND_DATA
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


Private Const ERROR_NO_MORE_FILES = 18

Private Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" _
    (ByVal hFind As Long, lpvFindData As WIN32_FIND_DATA) As Long
    
Private Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" _
(ByVal hFtpSession As Long, ByVal lpszSearchFile As String, _
      lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContent As Long) As Long

Private Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" _
(ByVal hFtpSession As Long, ByVal lpszRemoteFile As String, _
      ByVal lpszNewFile As String, ByVal fFailIfExists As Boolean, ByVal dwFlagsAndAttributes As Long, _
      ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean

Private Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" _
(ByVal hFtpSession As Long, ByVal lpszLocalFile As String, _
      ByVal lpszRemoteFile As String, _
      ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean


Public Declare Function FtpRenameFile Lib "wininet.dll" Alias "FtpRenameFileA" _
    (ByVal hFtpSession As Long, ByVal lpszOldFileName As String, ByVal lpszNewFileName As String) As Boolean


Private Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" _
    (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
    
Private Declare Function FtpGetCurrentDirectory Lib "wininet.dll" Alias "FtpGetCurrentDirectoryA" _
   (ByVal hFtpSession As Long, ByVal lpszCurrentDirectory As String, lpdwCurrentDirectory As Long) As Boolean
    
'paul - added a few more useful features
Private Declare Function FtpCreateDirectory Lib "wininet.dll" Alias "FtpCreateDirectoryA" _
    (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean

Private Declare Function FtpRemoveDirectory Lib "wininet.dll" Alias "FtpRemoveDirectoryA" _
    (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean

    
' Initializes an application's use of the Win32 Internet functions
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" _
(ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, _
ByVal sProxyBypass As String, ByVal lFlags As Long) As Long

' User agent constant.
Private Const SCUSERAGENT = "vb wininet"

' Use registry access settings.
Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Private Const INTERNET_OPEN_TYPE_DIRECT = 1
Private Const INTERNET_OPEN_TYPE_PROXY = 3
Private Const INTERNET_INVALID_PORT_NUMBER = 0

Private Const FTP_TRANSFER_TYPE_ASCII = &H1
Private Const FTP_TRANSFER_TYPE_BINARY = &H0
Private Const INTERNET_FLAG_PASSIVE = &H8000000
Private Const INTERNET_FLAG_EXISTING_CONNECT = &H20000000
Private Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000
' Opens a HTTP session for a given site.
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" _
(ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Long, _
ByVal sUserName As String, ByVal sPassword As String, ByVal lService As Long, _
ByVal lFlags As Long, ByVal lContext As Long) As Long
                
Private Const ERROR_INTERNET_EXTENDED_ERROR = 12003
Private Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" ( _
    lpdwError As Long, _
    ByVal lpszBuffer As String, _
    lpdwBufferLength As Long) As Boolean

' Number of the TCP/IP port on the server to connect to.
Private Const INTERNET_DEFAULT_FTP_PORT = 21
Private Const INTERNET_DEFAULT_GOPHER_PORT = 70
Private Const INTERNET_DEFAULT_HTTP_PORT = 80
Private Const INTERNET_DEFAULT_HTTPS_PORT = 443
Private Const INTERNET_DEFAULT_SOCKS_PORT = 1080

Private Const INTERNET_OPTION_CONNECT_TIMEOUT = 2
Private Const INTERNET_OPTION_RECEIVE_TIMEOUT = 6
Private Const INTERNET_OPTION_SEND_TIMEOUT = 5

Private Const INTERNET_OPTION_USERNAME = 28
Private Const INTERNET_OPTION_PASSWORD = 29
Private Const INTERNET_OPTION_PROXY_USERNAME = 43
Private Const INTERNET_OPTION_PROXY_PASSWORD = 44

' Type of service to access.
Private Const INTERNET_SERVICE_FTP = 1
Private Const INTERNET_SERVICE_GOPHER = 2
Private Const INTERNET_SERVICE_HTTP = 3

' Opens an HTTP request handle.
Private Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" _
(ByVal hHttpSession As Long, ByVal sVerb As String, ByVal sObjectName As String, ByVal sVersion As String, _
ByVal sReferer As String, ByVal something As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long

' Brings the data across the wire even if it locally cached.
Private Const INTERNET_FLAG_RELOAD = &H80000000
Private Const INTERNET_FLAG_KEEP_CONNECTION = &H400000
Private Const INTERNET_FLAG_MULTIPART = &H200000

Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000

' Sends the specified request to the HTTP server.
Private Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" (ByVal _
hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal sOptional As _
String, ByVal lOptionalLength As Long) As Integer


' Queries for information about an HTTP request.
Private Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" _
(ByVal hHttpRequest As Long, ByVal lInfoLevel As Long, ByRef sBuffer As Any, _
ByRef lBufferLength As Long, ByRef lIndex As Long) As Integer

' The possible values for the lInfoLevel parameter include:
Private Const HTTP_QUERY_CONTENT_TYPE = 1
Private Const HTTP_QUERY_CONTENT_LENGTH = 5
Private Const HTTP_QUERY_EXPIRES = 10
Private Const HTTP_QUERY_LAST_MODIFIED = 11
Private Const HTTP_QUERY_PRAGMA = 17
Private Const HTTP_QUERY_VERSION = 18
Private Const HTTP_QUERY_STATUS_CODE = 19
Private Const HTTP_QUERY_STATUS_TEXT = 20
Private Const HTTP_QUERY_RAW_HEADERS = 21
Private Const HTTP_QUERY_RAW_HEADERS_CRLF = 22
Private Const HTTP_QUERY_FORWARDED = 30
Private Const HTTP_QUERY_SERVER = 37
Private Const HTTP_QUERY_USER_AGENT = 39
Private Const HTTP_QUERY_SET_COOKIE = 43
Private Const HTTP_QUERY_REQUEST_METHOD = 45
Private Const HTTP_STATUS_DENIED = 401
Private Const HTTP_STATUS_PROXY_AUTH_REQ = 407

' Add this flag to the about flags to get request header.
Private Const HTTP_QUERY_FLAG_REQUEST_HEADERS = &H80000000
Private Const HTTP_QUERY_FLAG_NUMBER = &H20000000
' Reads data from a handle opened by the HttpOpenRequest function.
Private Declare Function InternetReadFile Lib "wininet.dll" _
(ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, _
lNumberOfBytesRead As Long) As Integer

Private Declare Function InternetWriteFile Lib "wininet.dll" _
        (ByVal hFile As Long, ByVal sBuffer As String, _
        ByVal lNumberOfBytesToRead As Long, _
        lNumberOfBytesRead As Long) As Integer

Private Declare Function FtpOpenFile Lib "wininet.dll" Alias _
        "FtpOpenFileA" (ByVal hFtpSession As Long, _
        ByVal sFileName As String, ByVal lAccess As Long, _
        ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function FtpDeleteFile Lib "wininet.dll" _
    Alias "FtpDeleteFileA" (ByVal hFtpSession As Long, _
    ByVal lpszFileName As String) As Boolean
Private Declare Function InternetSetOption Lib "wininet.dll" Alias "InternetSetOptionA" _
(ByVal hInternet As Long, ByVal lOption As Long, ByRef sBuffer As Any, ByVal lBufferLength As Long) As Integer
Private Declare Function InternetSetOptionStr Lib "wininet.dll" Alias "InternetSetOptionA" _
(ByVal hInternet As Long, ByVal lOption As Long, ByVal sBuffer As String, ByVal lBufferLength As Long) As Integer

' Closes a single Internet handle or a subtree of Internet handles.
Private Declare Function InternetCloseHandle Lib "wininet.dll" _
(ByVal hInet As Long) As Integer

' Queries an Internet option on the specified handle
Private Declare Function InternetQueryOption Lib "wininet.dll" Alias "InternetQueryOptionA" _
(ByVal hInternet As Long, ByVal lOption As Long, ByRef sBuffer As Any, ByRef lBufferLength As Long) As Integer

' Returns the version number of Wininet.dll.
Private Const INTERNET_OPTION_VERSION = 40

' Contains the version number of the DLL that contains the Windows Internet
' functions (Wininet.dll). This structure is used when passing the
' INTERNET_OPTION_VERSION flag to the InternetQueryOption function.
Private Type tWinInetDLLVersion
    lMajorVersion As Long
    lMinorVersion As Long
End Type

' Adds one or more HTTP request headers to the HTTP request handle.
Private Declare Function HttpAddRequestHeaders Lib "wininet.dll" Alias "HttpAddRequestHeadersA" _
(ByVal hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, _
ByVal lModifiers As Long) As Integer

' Flags to modify the semantics of this function. Can be a combination of these values:

' Adds the header only if it does not already exist; otherwise, an error is returned.
Private Const HTTP_ADDREQ_FLAG_ADD_IF_NEW = &H10000000

' Adds the header if it does not exist. Used with REPLACE.
Private Const HTTP_ADDREQ_FLAG_ADD = &H20000000

' Replaces or removes a header. If the header value is empty and the header is found,
' it is removed. If not empty, the header value is replaced
Private Const HTTP_ADDREQ_FLAG_REPLACE = &H80000000
'


Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, ByVal lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpLibFileName As String) As Long
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800


Public Function fnTransferFile(sHost As String, sUser As String, sPWD As String, _
                                sSourcePathAndFile As String, _
                                Optional bFromUnixToLocal As Boolean = True, _
                                Optional sDestPath As String, _
                                Optional sDestFilename As String, _
                                Optional sRetErrMsg As String, _
                                Optional bShowError As Boolean = True, _
                                Optional bDeleteSourceFile As Boolean = False, _
                                Optional bPassive As Boolean = True, _
                                Optional bBinary As Boolean = True) As Boolean
    
    Dim hOpen As Long, hConnection As Long, hFind As Long
    Dim lBinAscii As Long, lActPassive As Long
    Dim bFlag, bRenameFlag As Boolean, i As Integer
    Dim sSourcePath As String, sSourceFile As String, sItemName As String
    Dim pData As WIN32_FIND_DATA
    Dim lLastDLLError As Long
    Dim sArrFile() As String
    Dim sArrOldname() As String
    Dim sHostKey As String          '8696
    Dim sPLinkResult As String      '8696
    Dim sPLinkError As String       '8696
    
    fnTransferFile = False
    sDestPath = Trim(sDestPath)
    sSourcePathAndFile = Trim(sSourcePathAndFile)
    
    If sDestFilename <> "" Then
        If InStr(1, sDestFilename, "*") > 0 Or InStr(1, sDestFilename, "?") > 0 Then
            sRetErrMsg = "Invalid destination path specified, cannot continue!"
            subShowError bShowError, sRetErrMsg, vbExclamation
            GoTo CloseHandle
        End If
    End If
        
    sHostKey = Trim(fnGetServerHostKey())
    If sHostKey <> "" Then GoTo RUN_SFTP
        
    hOpen = InternetOpen(SCUSERAGENT, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, INTERNET_FLAG_NO_CACHE_WRITE)
    If hOpen = 0 Then
        fnErrorOut Err.LastDllError, "InternetOpen", bShowError, sRetErrMsg
        GoTo CloseHandle
    End If
    
    'IIf bPassive, lActPassive = INTERNET_FLAG_PASSIVE, lActPassive = 0
    hConnection = InternetConnect(hOpen, sHost, INTERNET_DEFAULT_FTP_PORT, sUser, sPWD, INTERNET_SERVICE_FTP, INTERNET_FLAG_EXISTING_CONNECT Or INTERNET_FLAG_PASSIVE, 0)
    If hConnection = 0 Then
        fnErrorOut Err.LastDllError, "InternetConnect", bShowError, sRetErrMsg
        GoTo CloseHandle
    End If
    
    If sSourcePathAndFile = "" Then
        sRetErrMsg = "Source file path is not specified, cannot continue!"
        subShowError bShowError, sRetErrMsg, vbExclamation
        GoTo CloseHandle
    End If
    
    If sDestPath = "" Then
        Dim szDir As String
        If bFromUnixToLocal Then
            'szDir = App.path
            szDir = io.ApplicationPath
        Else
            'Get the destination directory path from unix server...
            szDir = String(1024, Chr$(0))
            If FtpGetCurrentDirectory(hConnection, szDir, 1024) = False Then
                fnErrorOut Err.LastDllError, "FtpGetCurrentDirectory", bShowError, sRetErrMsg
                GoTo CloseHandle
            End If
        End If
        sDestPath = Trim(szDir)
        If InStr(1, Trim(szDir), Chr(0)) > 0 Then
            sDestPath = Mid(Trim(szDir), 1, InStr(1, Trim(szDir), Chr(0)) - 1)
        End If
    End If
    
    On Error GoTo errTrap
    If bFromUnixToLocal Then
        'Separate the file path and file name...
        subSeparateFileAndPath sSourcePathAndFile, sSourceFile, sSourcePath
        'Validate and set the source path...
        bFlag = FtpSetCurrentDirectory(hConnection, sSourcePath)
        If bFlag = False Then
            fnErrorOut Err.LastDllError, "FtpSetCurrentDirectory", bShowError, sRetErrMsg
            GoTo CloseHandle
        End If
        'Validate the destination path...
        If Dir(sDestPath, vbDirectory) = "" Then
            sRetErrMsg = "Invalid destination path specified, cannot continue!"
            subShowError bShowError, sRetErrMsg, vbExclamation
            GoTo CloseHandle
        End If
        subAddSlash sDestPath, "\"
        pData.cFileName = String(MAX_PATH, 0)
        'Get all the matching files...
        hFind = FtpFindFirstFile(hConnection, sSourceFile, pData, INTERNET_FLAG_RELOAD Or _
                               INTERNET_FLAG_NO_CACHE_WRITE, 0)
        If hFind = 0 Then
            lLastDLLError = Err.LastDllError
            If lLastDLLError = ERROR_NO_MORE_FILES Then
                sRetErrMsg = "No file(s) found to transfer!"
                subShowError bShowError, sRetErrMsg, vbExclamation
            Else
                fnErrorOut Err.LastDllError, "FtpFindFirstFile", bShowError, sRetErrMsg
            End If
            GoTo CloseHandle
        End If
        'Get the filenames, More than one file possible...
        
        
        sItemName = Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
        lLastDLLError = NO_ERROR
        Do
            pData.cFileName = String(MAX_PATH, 0)
            bFlag = InternetFindNextFile(hFind, pData)
            If bFlag = False Then
                lLastDLLError = Err.LastDllError
                If lLastDLLError = ERROR_NO_MORE_FILES Then
                    Exit Do
                Else
                    fnErrorOut Err.LastDllError, "InternetFindNextFile", bShowError, sRetErrMsg
                    GoTo CloseHandle
                End If
            Else
                sItemName = sItemName & "," & Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
            End If
        Loop
        
        If Trim(sItemName) = "" Then GoTo CloseHandle
        'Fix for 421631 Robert C Atwood 8/29/03 6:46 PM just before a long weekend. :P
        'Old code never worked, bBinary not initialized, etc.  Someone was being way too
        '"clever" for words here, so binary/ascii has been broken forever.
        'IIf bBinary, lBinAscii = FTP_TRANSFER_TYPE_BINARY, lBinAscii = FTP_TRANSFER_TYPE_ASCII
        lBinAscii = 0 'Default Binary
        If Not (bBinary) Then ' If we're in NOT in binary, set binascii to 1.
            lBinAscii = 1
        End If
        sArrFile = Split(sItemName, ",")
        'sArrOldname = Split(sItemName, ",")
        
        'Transfer the file to the local path specified...
        Dim sDest As String
        Dim nNum As Integer, sNum As String
        For i = 0 To UBound(sArrFile)
           'Commented out by Junsong 08/21/2006 ticket 529160
            'I changed the parameter for FTPOPE, so there is no cache will be written, so we don't need to rename the file
            'otherwise, some error occur, the program will generate some strange files.
            'sArrFile(i) = fnGenerateNewFile10()
            
            If sDestFilename <> "" Then
                nNum = InStr(1, sDestFilename, ".")
                If i = 0 Then
                    sNum = ""
                Else
                    sNum = CStr(i)
                End If
                If nNum > 0 Then
                    sDest = sDestPath & Mid(sDestFilename, 1, nNum - 1) & sNum & "." & Mid(sDestFilename, nNum + 1)
                Else
                    sDest = sDestPath & sDestFilename & sNum
                End If
            Else
                sDest = sDestPath & sArrFile(i)
            End If
            
'            bFlag = FtpRenameFile(hConnection, sArrOldname(i), sArrFile(i))
'            If bFlag = False Then
'                fnErrorOut Err.LastDllError, "FtpGetFile", bShowError, sRetErrMsg
'                GoTo CloseHandle
'            End If
'
            bFlag = FtpGetFile(hConnection, sArrFile(i), sDest, False, INTERNET_FLAG_RELOAD, lBinAscii, 0)
           If bFlag = False Then
                'bRenameFlag = FtpRenameFile(hConnection, sArrFile(i), sArrOldname(i))
                fnErrorOut Err.LastDllError, "FtpGetFile", bShowError, sRetErrMsg
                GoTo CloseHandle
            End If
'            On Error Resume Next
'            Name sDestPath & sArrFile(i) As sDestPath & sArrOldname(i)
'            On Error GoTo errTrap
            If bDeleteSourceFile Then
                bFlag = FtpDeleteFile(hConnection, sArrFile(i))
                If bFlag = False Then
'                    bRenameFlag = FtpRenameFile(hConnection, sArrFile(i), sArrOldname(i))
                    fnErrorOut Err.LastDllError, "FtpDeleteFile", bShowError, sRetErrMsg
                    GoTo CloseHandle
                End If
            End If
        Next i
    Else 'Transfer from local to unix server...
        '#wj120104 - Initialize lBinAscii
        lBinAscii = 0 'Default Binary
        If Not (bBinary) Then ' If we're in NOT in binary, set binascii to 1.
            lBinAscii = 1
        End If
        
        Dim nNumFiles As Integer
        'Separate the file path and file name...
        subSeparateFileAndPath sSourcePathAndFile, sSourceFile, sSourcePath, False
        
        'Validate the source path...
        If Dir(sSourcePath, vbDirectory) = "" Then
            sRetErrMsg = "Invalid Source path specified, cannot continue!"
            subShowError bShowError, sRetErrMsg, vbExclamation
            GoTo CloseHandle
        End If
        'Validate the destination path...
        bFlag = FtpSetCurrentDirectory(hConnection, sDestPath)
        If bFlag = False Then
            fnErrorOut Err.LastDllError, "FtpSetCurrentDirectory", bShowError, sRetErrMsg
            GoTo CloseHandle
        End If
        'Get all the matching files if more than one...
        nNumFiles = fnBuildFileArray(sSourcePath, sSourceFile, sArrFile)
        If nNumFiles >= 0 Then
            For i = 0 To nNumFiles
                If sDestFilename <> "" Then
                    nNum = InStr(1, sDestFilename, ".")
                    If i = 0 Then
                        sNum = ""
                    Else
                        sNum = CStr(i)
                    End If
                    If nNum > 0 Then
                        sDest = Mid(sDestFilename, 1, nNum - 1) & sNum & "." & Mid(sDestFilename, nNum + 1)
                    Else
                        sDest = sDestFilename & sNum
                    End If
                Else
                    sDest = sArrFile(i)
                End If
                subAddSlash sSourcePath, "\"
                sSourceFile = sSourcePath & sArrFile(i)
                bFlag = FtpPutFile(hConnection, sSourceFile, sDest, lBinAscii, 0)
                If bFlag = False Then
                    fnErrorOut Err.LastDllError, "FtpPutFile", bShowError, sRetErrMsg
                    GoTo CloseHandle
                End If
                If bDeleteSourceFile Then
                    Kill (sSourceFile)
                End If
            Next i
        End If
    End If
    
    
    fnTransferFile = True
    
CloseHandle:
    On Error Resume Next
    InternetCloseHandle hFind
    InternetCloseHandle hConnection
    InternetCloseHandle hOpen
    Exit Function
           
RUN_SFTP: 'G8696
    fnTransferFile = fnSTransferFile(sHostKey, sHost, sUser, sPWD, _
                                sSourcePathAndFile, _
                                 bFromUnixToLocal, _
                                 sDestPath, _
                                 sDestFilename, _
                                 sRetErrMsg, _
                                 bShowError, _
                                 bDeleteSourceFile, _
                                 bPassive, _
                                 bBinary)
    Exit Function
    
errTrap:
    sRetErrMsg = Err.Number & ": " & Err.Description
    subShowError bShowError, sRetErrMsg, vbCritical
    GoTo CloseHandle
End Function

Private Function fnErrorOut(dError As Long, szCallFunction As String, _
                            bShowError As Boolean, sErrorMsg As String) As Boolean
    Const WRONGPASSWORD As Integer = 12014
    Dim dwIntError As Long
    Dim dwLength As Long
    Dim strBuffer As String
    
    fnErrorOut = True
    sErrorMsg = ""
    
    If dError = ERROR_INTERNET_EXTENDED_ERROR Then
        InternetGetLastResponseInfo dwIntError, vbNullString, dwLength
        strBuffer = String(dwLength + 1, 0)
        InternetGetLastResponseInfo dwIntError, strBuffer, dwLength
        sErrorMsg = szCallFunction & " Err: " & dwIntError & " " & strBuffer
        subShowError bShowError, sErrorMsg, vbCritical
        Exit Function
    End If
    
    Select Case szCallFunction
        Case "InternetOpen"
            sErrorMsg = "Failed to access the DLL"
        Case "InternetConnect"
            sErrorMsg = "Failed to connect to the host server"
            If dError = WRONGPASSWORD Then
                sErrorMsg = "Invalid login information provided, cannot continue!"
            End If
        Case "FtpSetCurrentDirectory"
            sErrorMsg = "Invalid path, Failed to set the current directory"
        Case "FtpGetCurrentDirectory"
            sErrorMsg = "Failed to get the current directory path"
        Case "FtpFindFirstFile"
            sErrorMsg = "Failed to get the file from the server"
        Case "InternetFindNextFile"
            sErrorMsg = "Failed to get the file from the server"
        Case "FtpGetFile"
            sErrorMsg = "Failed to transfer the file"
        Case "FtpPutFile"
            sErrorMsg = "Failed to transfer the file"
        Case "FtpDeleteFile"
            sErrorMsg = "Failed to delete the source file"
    End Select
    
    sErrorMsg = sErrorMsg & ". (" & fnGetErrorMessage(dError) & ")"
    subShowError bShowError, sErrorMsg, vbCritical
    fnErrorOut = False
    
End Function

Private Function fnGetErrorMessage(ByVal dwError As Long, Optional sModuleName As String = "wininet.dll") As String
    Dim dwLength As Long
    Dim szString As String * 2048
    
    If (dwError < 12000) Then
        fnGetErrorMessage = "Error code: " & dwError
        Exit Function
    End If
        
    dwLength = FormatMessage(FORMAT_MESSAGE_FROM_HMODULE, GetModuleHandle(sModuleName), dwError, 0, szString, 256, 0)
    
    If dwLength > 0 Then
        fnGetErrorMessage = "Error code: " & dwError & " Message: " & Left(szString, dwLength - 2)
    Else
        fnGetErrorMessage = "Error code: " & dwError
    End If
    
End Function

Private Function fnIsWholePath(sPath As String) As Boolean
    If InStr(sPath, ":") > 0 Or Left(sPath, 1) = "\" Then
        fnIsWholePath = True
    Else
        fnIsWholePath = False
    End If
End Function

Private Sub subAddSlash(sStr As String, sSlash As String)
    If sStr <> "" Then
        If Right(sStr, 1) <> sSlash Then
            sStr = sStr & sSlash
        End If
    End If
End Sub

Private Sub subShowError(bShow As Boolean, sErrMsg As String, vbType As VbMsgBoxStyle)
    If bShow Then
        MsgBox sErrMsg, vbType
    End If
End Sub

Private Sub subSeparateFileAndPath(sPath As String, sFile As String, sFilePath As String, _
                                   Optional bRemote As Boolean = True)
    Dim nPos As Integer
    Dim sSlash As String
    
    If sPath = "" Then
        Exit Sub
    End If
    
    If bRemote Then
        sSlash = "/"
    Else
        sSlash = "\"
    End If
    
    nPos = InStrRev(sPath, sSlash)
    If nPos > 1 Then
        sFile = Mid(sPath, nPos + 1)
        sFilePath = Left(sPath, nPos - 1)
    End If
    
End Sub

Public Function fnValidUnixPath(sHost As String, sUser As String, sPWD As String, _
                                sPath As String, sRetErrMsg As String, _
                                Optional bShowError As Boolean = False) As Boolean
    Dim sHostKey As String
    Dim hOpen As Long, hConnection As Long, hFind As Long
    Dim bFlag As Boolean
    
    fnValidUnixPath = False
    
    If Trim(sPath) = "" Then Exit Function
    
    sHostKey = Trim(fnGetServerHostKey())
    If sHostKey <> "" Then GoTo RUN_fnSValidUnixPath
    
    hOpen = InternetOpen(SCUSERAGENT, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
    If hOpen = 0 Then
        fnErrorOut Err.LastDllError, "InternetOpen", bShowError, sRetErrMsg
        GoTo CloseHandle
    End If
    
    hConnection = InternetConnect(hOpen, sHost, INTERNET_DEFAULT_FTP_PORT, sUser, sPWD, INTERNET_SERVICE_FTP, INTERNET_FLAG_EXISTING_CONNECT Or INTERNET_FLAG_PASSIVE, 0)
    If hConnection = 0 Then
        fnErrorOut Err.LastDllError, "InternetConnect", bShowError, sRetErrMsg
        GoTo CloseHandle
    End If
    
    'Validate the unix server path...
    sPath = Trim(sPath)
    bFlag = FtpSetCurrentDirectory(hConnection, sPath)
    If bFlag = False Then
        fnErrorOut Err.LastDllError, "FtpSetCurrentDirectory", bShowError, sRetErrMsg
        GoTo CloseHandle
    End If
    
    fnValidUnixPath = True

CloseHandle:
    On Error Resume Next
    InternetCloseHandle hConnection
    InternetCloseHandle hOpen
    Exit Function
    
RUN_fnSValidUnixPath:
    fnValidUnixPath = fnSValidUnixPath(sHostKey, _
                                sHost, sUser, sPWD, _
                                sPath, sRetErrMsg, _
                                bShowError)
End Function

Private Function fnBuildFileArray(sPath As String, sFileName As String, _
                                  aryFiles() As String, Optional vSortFilesByDate) As Integer
    
    Const SUB_NAME = "fnBuildFileArray"
    Dim nMaxArraySize As Integer, nNumOfFiles As Integer
    Dim sFileFound As String, AryDates() As String

    If IsMissing(vSortFilesByDate) Then
        vSortFilesByDate = False
    Else
        vSortFilesByDate = True
    End If
    
    nNumOfFiles = -1
    On Error GoTo errBuildList
    
    sFileFound = Dir(fnConCat(sPath, sFileName))
    If sFileFound <> "" Then
        nMaxArraySize = 50
        ReDim aryFiles(nMaxArraySize)
        
        While sFileFound <> ""
            nNumOfFiles = nNumOfFiles + 1
            If nNumOfFiles > nMaxArraySize Then
                nMaxArraySize = nMaxArraySize + 10
                ReDim Preserve aryFiles(nMaxArraySize)
            End If
            aryFiles(nNumOfFiles) = sFileFound
            ReDim Preserve AryDates(nNumOfFiles)
            AryDates(nNumOfFiles) = Format(FileDateTime(fnConCat(sPath, sFileFound)), "mm/dd/yyyy  hh:mm:ss AMPM")
            sFileFound = Dir()
        Wend
        
        If vSortFilesByDate Then
            'sort aryFiles by Date ascending
            subSortFileDate aryFiles, nNumOfFiles, AryDates
        End If
    End If
    
    fnBuildFileArray = nNumOfFiles
    Exit Function

errBuildList:
    If Err.Number = 76 Then
        MsgBox " Path (" & fnConCat(sPath, sFileName) & ") not found in Function '" & SUB_NAME & "'"
    End If
    fnBuildFileArray = -1
End Function

Private Function fnConCat(MyPath As String, MyName As String) As String
    fnConCat = IIf(Right(MyPath, 1) = "\", MyPath, MyPath + "\") + MyName
End Function

Private Sub subSortFileDate(aryFiles() As String, nNumOfFiles As Integer, AryDates() As String)
    Dim lRow As Long, nCol As Integer, j As Long, k As Long, i As Integer, sTemp As String
    
    On Error GoTo ErrorOutofRange
    lRow = UBound(AryDates)
    
    If nNumOfFiles <> lRow Then
        tfnErrHandler "subSortFileDate", -1, "arydates Size mismatch", False
        Exit Sub
    End If
    
    On Error GoTo 0
    For j = 0 To lRow
        For k = 0 To lRow
            If AryDates(j) < AryDates(k) Then
                'swap key array
                sTemp = AryDates(j)
                AryDates(j) = AryDates(k)
                AryDates(k) = sTemp
                
                'swap data array
                sTemp = aryFiles(j)
                aryFiles(j) = aryFiles(k)
                aryFiles(k) = sTemp
            End If
        Next k
    Next j
    
    Exit Sub
    
ErrorOutofRange:
    tfnErrHandler "subSortFileDate", False
    On Error GoTo 0
End Sub

Private Function fnGenerateNewFile10() As String
    Dim i As Integer
    Dim sTemp As String
    Dim sChar As String
    
    Randomize
    
    sTemp = ""
    For i = 1 To 10
        sChar = Chr((Int(Rnd() * 25)) + 97)
        sTemp = sTemp + sChar
    Next i
    
    fnGenerateNewFile10 = sTemp & ".eml"
End Function

Public Function fnCreateUnixDirectory(sHost As String, sUser As String, sPWD As String, _
                                sPath As String, sRetErrMsg As String, _
                                Optional bShowError As Boolean = False) As Boolean
    Dim sHostKey As String
    Dim hOpen As Long, hConnection As Long, hFind As Long
    Dim bFlag As Boolean
    
    fnCreateUnixDirectory = False
    
    If Trim(sPath) = "" Then Exit Function
        
    sHostKey = Trim(fnGetServerHostKey())
    If sHostKey <> "" Then GoTo RUN_fnSCreateUnixDirectory
    
    hOpen = InternetOpen(SCUSERAGENT, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
    If hOpen = 0 Then
        fnErrorOut Err.LastDllError, "InternetOpen", bShowError, sRetErrMsg
        GoTo CloseHandle
    End If
    
    hConnection = InternetConnect(hOpen, sHost, INTERNET_DEFAULT_FTP_PORT, sUser, sPWD, INTERNET_SERVICE_FTP, INTERNET_FLAG_EXISTING_CONNECT Or INTERNET_FLAG_PASSIVE, 0)
    If hConnection = 0 Then
        fnErrorOut Err.LastDllError, "InternetConnect", bShowError, sRetErrMsg
        GoTo CloseHandle
    End If
    
    'Validate the unix server path...
    sPath = Trim(sPath)
    bFlag = FtpCreateDirectory(hConnection, sPath)
    If bFlag = False Then
        fnErrorOut Err.LastDllError, "FtpCreateDirectory", bShowError, sRetErrMsg
        GoTo CloseHandle
    End If
    
    fnCreateUnixDirectory = True

CloseHandle:
    On Error Resume Next
    InternetCloseHandle hConnection
    InternetCloseHandle hOpen
    Exit Function
    
RUN_fnSCreateUnixDirectory:
    fnCreateUnixDirectory = fnSCreateUnixDirectory(sHostKey, sHost, sUser, sPWD, _
                                sPath, sRetErrMsg, _
                                bShowError)
End Function

Public Function fnConCatUnixPath(MyPath As String, MyName As String) As String
    fnConCatUnixPath = IIf(Right(MyPath, 1) = "/", MyPath, MyPath + "/") + MyName
End Function

Public Function fnRemoveInvalidUnixCharacters(MyName As String) As String
    'removes these characters as they will cause problems on unix
    ' ' " ( ) < > \ / ? * | :
    Const strip As String = "'""()<>/\?*|"
    Dim i As Integer
    
    fnRemoveInvalidUnixCharacters = MyName
    For i = 1 To Len(strip)
        fnRemoveInvalidUnixCharacters = Replace(fnRemoveInvalidUnixCharacters, Mid$(strip, i, 1), "")
    Next i

End Function

'**********************************************************
'* USE PSCP - 8696

Public Function fnSCreateUnixDirectory(ByVal sHostKey As String, _
                                sHost As String, sUser As String, sPWD As String, _
                                sPath As String, sRetErrMsg As String, _
                                Optional bShowError As Boolean = False) As Boolean
    'G8696 - Tthompson
  ' use plink via tfnRunRemoteCmd to "mkdir /home/ssfactor/asdf"
        ' STDOUT success results: mkdir /home/ssfactor/asdf: OK
    ' use plink via tfnRunRemoteCmd to "mkdir /home/ssfactor/asdf"
        ' STDOUT fail results: mkdir /home/ssfactor/asdf: failure
    ' use plink via tfnRunRemoteCmd to "mkdir /asdf/asdf"
        ' STDOUT fail results: mkdir /asdf/asdf: no such file or directory
            
    'sPath is passed through Q_Str() for commandline processing
    
    Dim sCmd As String, sResult As String
    
    sCmd = "mkdir " & Q_Str(sPath)
    fnSCreateUnixDirectory = tfnRunRemoteCmd(sHostKey, _
                       rapPlink, _
                       sCmd, _
                       sResult, _
                       sRetErrMsg, _
                       sHost, _
                        sUser, _
                        sPWD, _
                        , _
                       True)
                                                
End Function

Public Function fnSValidUnixPath(ByVal sHostKey As String, _
                                sHost As String, sUser As String, sPWD As String, _
                                sPath As String, sRetErrMsg As String, _
                                Optional bShowError As Boolean = False) As Boolean
    'G8696 - Tthompson
    ' use psftp via tfnRunRemoteCmd to "cd /home/ssfactor"
        ' STDOUT success results: Remote directory is now /home/ssfactor
    ' use psftp via tfnRunRemoteCmd to "cd /home/asdf"
        ' STDOUT fail results: Directory /asdf: no such file or directory
            
    'sPath is passed through Q_Str() for commandline processing
    
    Dim sCmd As String, sResult As String
    
    sCmd = "cd " & Q_Str(sPath)
    fnSValidUnixPath = tfnRunRemoteCmd(sHostKey, _
                       rapPlink, _
                       sCmd, _
                       sResult, _
                       sRetErrMsg, _
                       sHost, _
                        sUser, _
                        sPWD, _
                        , _
                       True)
    
End Function

Public Function fnSTransferFile(ByVal sHostKey As String, _
                                sHost As String, sUser As String, sPWD As String, _
                                sSourcePathAndFile As String, _
                                Optional bFromUnixToLocal As Boolean = True, _
                                Optional sDestPath As String, _
                                Optional sDestFilename As String, _
                                Optional sRetErrMsg As String, _
                                Optional bShowError As Boolean = True, _
                                Optional bDeleteSourceFile As Boolean = False, _
                                Optional bPassive As Boolean = True, _
                                Optional bBinary As Boolean = True) As Boolean
    Dim bFlag, bRenameFlag As Boolean, i As Long
    Dim sSourcePath As String, sSourceFile As String, sItemName As String
    Dim sCmd As String, sResult As String
    Dim sArrFile() As String
    Dim sDest As String
    Dim nNum As Integer, sNum As String
    Dim sWorkDir As String
    Dim sFileList As String
    Dim sInsSrcSpec As String
    Dim aryFiles() As String
    Dim aryCmds() As String
    Dim colFiles As New Collection
    Dim fCnt As Long
    Dim cCnt As Long
    Dim sFile As String
    Dim sRemoteFile As String
    Dim sWorkFile As String
    Dim sDestFile As String
    Dim sFindName As String
    Dim sFindPath As String
    Dim sHold As String
    Dim lTime As Long
    Dim sModule As String
    Dim sTmpPath As String
    Dim sTmpFile As String
    
    fnSTransferFile = False
    sDestPath = Trim(sDestPath)
    sDestFilename = Trim(sDestFilename)
    sSourcePathAndFile = Trim(sSourcePathAndFile)
    sModule = fnGetProgramName()
    
    If sDestFilename <> "" Then
        If InStr(1, sDestFilename, "*") > 0 Or InStr(1, sDestFilename, "?") > 0 Then
            sRetErrMsg = "Invalid destination path specified, cannot continue!"
            subShowError bShowError, sRetErrMsg, vbExclamation
            GoTo CloseHandle
        End If
    End If
    
    If sSourcePathAndFile = "" Then
        sRetErrMsg = "Source file path is not specified, cannot continue!"
        subShowError bShowError, sRetErrMsg, vbExclamation
        GoTo CloseHandle
    End If
    
    If sDestPath = "" Then
        Dim szDir As String
        If bFromUnixToLocal Then
            szDir = io.ApplicationPath
        Else
            'Get the destination directory path from unix server...
            If tfnRunRemoteCmd(sHostKey, rapPlink, "pwd", szDir, sRetErrMsg, sHost, sUser, sPWD, , True) = False Then
                sRetErrMsg = "Error Getting destination directory from unix path: " & vbCrLf _
                           & sRetErrMsg
                subShowError bShowError, sRetErrMsg, vbExclamation
                GoTo CloseHandle
            End If
        End If
        sDestPath = Trim(szDir)
        If InStr(1, Trim(szDir), Chr(0)) > 0 Then
            sDestPath = Mid(Trim(szDir), 1, InStr(1, Trim(szDir), Chr(0)) - 1)
        End If
    End If
    
    ' Create ftpwork folder
    sWorkDir = fnCombinePath(io.ApplicationPath, "ftpwork\" & sModule)
    
    If Not io.CreateFolder(sWorkDir) Then
        sRetErrMsg = "Failed to create ftp work folder '" & sWorkDir & "'"
        subShowError bShowError, sRetErrMsg, vbExclamation
        Exit Function
    End If
    
    If Not fnClearDirectory(sWorkDir, sRetErrMsg) Then
        subShowError bShowError, sRetErrMsg, vbExclamation
        Exit Function
    End If
    
    On Error GoTo errTrap
    If bFromUnixToLocal Then
    
        'potential for sDestFilename to contain a full path
        ' check if sDestFilename contains a path or drive separator
        ' if a path is given then dont use sDestPath
        If InStr(1, sDestFilename, "\") + InStr(1, sDestFilename, ":") > 0 Then
            sDest = sDestFilename
        Else
            subAddSlash sDestPath, "\"
            sDest = sDestPath & sDestFilename
        End If
        If Right(sDest, 1) = "\" Then sDest = Left(sDest, Len(sDest) - 1)
        
        'Separate the file path and file name...
        subSeparateFileAndPath sSourcePathAndFile, sSourceFile, sSourcePath
        subAddSlash sSourcePath, "/"
        
        ' Test unix folder
        If Not fnSValidUnixPath(sHostKey, sHost, sUser, sPWD, sSourcePath, sRetErrMsg, bShowError) Then
            subShowError bShowError, sRetErrMsg, vbExclamation
            GoTo CloseHandle
        End If
        
        ' Get remote files
        ' This will get us a snap-shot of files on the server.
        ' We must anticipate that more files may be added and some files
        ' may be removed before the file transfer is complete.
        sInsSrcSpec = fnInsenseString(sSourceFile)
        sFindName = Replace(sInsSrcSpec, " ", "' '")
        sFindName = Replace(sFindName, "(", "'('")
        sFindName = Replace(sFindName, ")", "')'")
        
        sFindPath = Replace(sSourcePath, " ", "' '")
        
        fCnt = fnGetRemoteFileList(sHostKey, sFindPath, sFindName, aryFiles, sRetErrMsg, _
                                   sHost, sUser, sPWD)
        If fCnt < 1 Then
            sRetErrMsg = "No file(s) found to transfer!"
            subShowError bShowError, sRetErrMsg, vbExclamation
            GoTo CloseHandle
        ElseIf sRetErrMsg <> "" Then
            If InStr(sRetErrMsg, "ls:") > 0 Then
                sRetErrMsg = "No file(s) found to transfer!"
                subShowError bShowError, sRetErrMsg, vbExclamation
                GoTo CloseHandle
            Else
                sRetErrMsg = "Error getting files from UNIX: " & vbCrLf _
                           & sRetErrMsg
                subShowError bShowError, sRetErrMsg, vbExclamation
                GoTo CloseHandle
            End If
        End If
        
        sInsSrcSpec = fnCombinePath(sSourcePath, sInsSrcSpec, True)
        
        lTime = Timer
        
        sCmd = sUser & "@" & sHost & ":" & Q_Str(sInsSrcSpec) & " " & Q_Str(sWorkDir)
        fnSTransferFile = tfnRunRemoteCmd(sHostKey, _
                                          rapPscp, _
                                          sCmd, _
                                          sResult, _
                                          sRetErrMsg, _
                                          sHost, _
                                          sUser, _
                                          sPWD, _
                                          , _
                                          True)
        If Not fnSTransferFile Then
            sRetErrMsg = "Error transfering files: " & vbCrLf _
                       & sCmd & vbCrLf _
                       & sRetErrMsg
            subShowError bShowError, sRetErrMsg, vbExclamation
            GoTo CloseHandle
        End If
        
        lTime = Timer - lTime
        Debug.Print "Download time = " & lTime
        
        If sDestFilename <> "" Then
            sTmpFile = fnGetFileName(sDestFilename)
        End If
        
        For i = 0 To (fCnt - 1)
            sDestFile = ""
            
            sFile = fnGetFileName(aryFiles(i))
            If sTmpFile <> "" Then
                sDestFile = fnNewDestFilename(sTmpFile, i)
            End If
            sDestFile = IIf(sDestFile <> "", fnCombinePath(sDestPath, sDestFile), fnCombinePath(sDestPath, sFile))
            sWorkFile = fnCombinePath(sWorkDir, sFile)
            
            ' Don't trip over any files that may have been removed from the server
            ' after we got our 'snap-shot' and were not downloaded.
            If io.FileExists(sWorkFile) Then
                
                ' We do NOT want to delete files from the server that were
                ' not downloaded to the PC. This collection will be used to
                ' build the delete commands.
                fnAddKeyToCollection colFiles, sFile, ""
            
                If Not bBinary Then
                    fnSTransferFile = fnConvertToWindowsEOL(sWorkFile, sRetErrMsg, True, sDestFile)
                    If Not fnSTransferFile Then
                        sRetErrMsg = "Failed to copy or convert ftp work file: '" & sWorkFile & "' " & vbCrLf _
                                   & "To: '" & sDestFile & "'" & vbCrLf _
                                   & sRetErrMsg
                    End If
                Else
                    fnSTransferFile = io.CopyFile(sWorkFile, sDestFile)
                    If Not fnSTransferFile Then
                        sRetErrMsg = "Failed to copy ftp work file: '" & sWorkFile & "' " & vbCrLf _
                                   & "To: '" & sDestFile & "'"
                    End If
                End If
    
                If fnSTransferFile Then
                    fnSTransferFile = io.DeleteFile(sWorkFile)
                    If Not fnSTransferFile Then
                        sRetErrMsg = "Failed to delete ftp work file: '" & sWorkFile & "'"
                    End If
                End If
            
                If Not fnSTransferFile Then
                    If sRetErrMsg = "" Then sRetErrMsg = "Failed to copy file from ftp work folder"
                    subShowError bShowError, sRetErrMsg, vbExclamation
                    GoTo CloseHandle
                End If
                
            End If
            
        Next
        
        If bDeleteSourceFile And fnSTransferFile Then
            
            cCnt = fnBuildRemoveFileCmdArray(aryFiles, colFiles, aryCmds, sRetErrMsg)
            If sRetErrMsg <> "" Then
                fnErrorOut Err.Number, "SFtpGetCurrentDirectory", bShowError, sRetErrMsg
                GoTo CloseHandle
            End If
            
            For i = 0 To (cCnt - 1)
                fnSTransferFile = tfnRunRemoteCmd(sHostKey, _
                                                  rapPlink, _
                                                  aryCmds(i), _
                                                  sResult, _
                                                  sRetErrMsg, _
                                                  sHost, _
                                                  sUser, _
                                                  sPWD, _
                                                  , _
                                                  True)
                If Not fnSTransferFile Then
                    sHold = "Failed to delete some or all of the server files below: " & vbCrLf _
                          & Replace(aryCmds(i), ";", vbCrLf) _
                          & vbCrLf _
                          & sRetErrMsg
                    sRetErrMsg = sHold
                    subShowError bShowError, sRetErrMsg, vbExclamation
                    GoTo CloseHandle
                End If
            Next
            
        End If
        
        If Not fnSTransferFile Then
            ' set error msg to ?
            Exit Function
        End If
        
    Else
        'Transfer from local(source) to unix server(dest)...
        ' example     put> pscp c:\documents\foo.txt    fred@example.com:/tmp/foo
        ' wildcard ex put> pscp c:\documents\*.txt      fred@example.com:/tmp/foo
        
        'Separate the file path and file name...
        'subSeparateFileAndPath sSourcePathAndFile, sSourceFile, sSourcePath, False
        
        'potential for sDestFilename to contain a full path
        ' check if sDestFilename contains a path or drive separator
        ' if a path is given then dont use sDestPath
        If InStr(1, sDestFilename, "/") + InStr(1, sDestFilename, ":") > 0 Then
            sDest = sDestFilename
        Else
            subAddSlash sDestPath, "/"
            sDest = sDestPath & sDestFilename
        End If
        If Right(sDest, 1) = "/" Then sDest = Left(sDest, Len(sDest) - 1)
        
        'Separate the file path and file name...
        subSeparateFileAndPath sSourcePathAndFile, sSourceFile, sSourcePath, False
        subAddSlash sSourcePath, "\"
        
        ' if sDestFilename contains a path, it overrides
        ' Get list of local files
        fCnt = fnGetLocalFileList(sSourceFile, sSourcePath, aryFiles, sRetErrMsg)
        If fCnt = 0 Then
            sRetErrMsg = "No file(s) found to transfer!"
            subShowError bShowError, sRetErrMsg, vbExclamation
            GoTo CloseHandle
        End If
        
        cCnt = 0
        ReDim aryCmds(fCnt)
        
        If sDestFilename <> "" Then
            sTmpFile = fnGetFileName(sDestFilename)
        End If
        
        ' Copy files to ftpwork folder. Convert EOL if not binary.
        For i = 0 To (fCnt - 1)
            If io.FileExists(aryFiles(i)) Then
                sDestFile = ""
                sFile = fnGetFileName(aryFiles(i))
                If sTmpFile <> "" Then
                    sDestFile = fnNewDestFilename(sTmpFile, i)
                End If
                
                sDestFile = IIf(sDestFile <> "", _
                            fnCombinePath(sWorkDir, sDestFile), _
                            fnCombinePath(sWorkDir, sFile))
                                
                ' So we know what files to delete from ftpwork folder
                aryCmds(cCnt) = sDestFile
                cCnt = cCnt + 1
                
                If bBinary Then
                    If Not io.CopyFile(aryFiles(i), sDestFile, True, bShowError) Then
                        sRetErrMsg = "Failed to copy file: " & vbCrLf & _
                                     "From '" & aryFiles(i) & "' " & vbCrLf & _
                                     "To: '" & sDestFile & "'"
                        subShowError bShowError, sRetErrMsg, vbExclamation
                        GoTo CloseHandle
                    End If
                Else
                    If Not fnConvertToUnixEOL(aryFiles(i), sRetErrMsg, bShowError, sDestFile) Then
                        sRetErrMsg = "Failed to convert EOL format for file: " & vbCrLf & _
                                     "From '" & aryFiles(i) & "' " & vbCrLf & _
                                     "To: '" & sDestFile & "'" & vbCrLf & _
                                     sRetErrMsg
                        subShowError bShowError, sRetErrMsg, vbExclamation
                        GoTo CloseHandle
                    End If
                End If
            End If
        Next
        
        ' Test unix folder
        If Not fnSValidUnixPath(sHostKey, sHost, sUser, sPWD, sDestPath, sRetErrMsg, bShowError) Then
            subShowError bShowError, sRetErrMsg, vbExclamation
            GoTo CloseHandle
        End If
        
        ' We want to FTP from the ftpwork folder now so we know what is going
        ' across.
        If sTmpFile = "" Then
            sHold = fnCombinePath(sWorkDir, sSourceFile)
            sCmd = Q_Str(sHold) & " " & sUser & "@" & sHost & ":" & Q_Str(sDestPath)
        Else
            sHold = fnNewDestFileWildcard(sTmpFile)
            sHold = fnCombinePath(sWorkDir, sHold)
            sCmd = Q_Str(sHold) & " " & sUser & "@" & sHost & ":" & Q_Str(sDestPath)
        End If
        
        fnSTransferFile = tfnRunRemoteCmd(sHostKey, _
                           rapPscp, _
                           sCmd, _
                           sResult, _
                           sRetErrMsg, _
                           sHost, _
                           sUser, _
                           sPWD, _
                           , _
                           True)
                           
        If Not fnSTransferFile Then
            sRetErrMsg = "Error transfering files: " & vbCrLf _
                       & sCmd & vbCrLf _
                       & sRetErrMsg
            subShowError bShowError, sRetErrMsg, vbExclamation
            GoTo CloseHandle
        End If
        
        sHold = fnCombinePath(sDestPath, "*", True)
        
        If fnSTransferFile Then
            tfnRunRemoteCmd sHostKey, _
                   rapPlink, _
                   "chmod 777 " & sHold, _
                   sResult, _
                   "", _
                   sHost, _
                   sUser, _
                   sPWD, _
                   , _
                   True
        End If
        
        fnDeleteFtpWorkFiles aryCmds
        
        If bDeleteSourceFile And fnSTransferFile Then
            
            For i = 0 To (fCnt - 1)
                If io.FileExists(aryFiles(i)) Then
                    If Not io.DeleteFile(aryFiles(i), bShowError) Then
                        sRetErrMsg = "Failed to delete file: " & vbCrLf & _
                                     "'" & aryFiles(i) & "' "
                        subShowError bShowError, sRetErrMsg, vbExclamation
                        'GoTo CloseHandle ' keep going?
                    End If
                End If
            Next
        End If
    
    End If
    
CloseHandle:
    Exit Function
    
errTrap:
    sRetErrMsg = Err.Number & ": " & Err.Description
    subShowError bShowError, sRetErrMsg, vbCritical
    GoTo CloseHandle
End Function

' Intended for UNIX OS.
' Assumes '\' path seperator.
' Assumes case sensitivity -> [Zz][Zz].sql lists zz.sql, Zz.sql, ZZ.sql, etc..
Public Function fnGetRemoteFileList(ByVal sHostKey As String, _
                                    ByVal sRemotePath As String, _
                                    ByVal sFileNameOrSpec As String, _
                                    ByRef aryFiles() As String, _
                                    ByRef sErrMsg As String, _
                                    Optional ByVal sHost As String = "", _
                                    Optional ByVal sUser As String = "", _
                                    Optional ByVal sPWD As String = "") As Long
    On Error GoTo FINISHED
    Dim count As Long
    Dim files() As String
    Dim sCmd As String
    Dim sResult As String
    Dim i As Long
    Dim j As Long
    
    ReDim aryFiles(0)
    
    If sRemotePath = "" Then
        sErrMsg = "Empty Remote Path passed to fnGetRemoteFileList"
        GoTo FINISHED
    End If
    
    If sFileNameOrSpec = "" Then
        sErrMsg = "Empty File Name Or Spec passed to fnGetRemoteFileList"
        GoTo FINISHED
    End If
    
    'sCmd = "cd " & sRemotePath & "; "
    If Right(sRemotePath, 1) = "/" Then
        sCmd = "ls -1 -d -p " & sRemotePath & sFileNameOrSpec & " | grep -v '.*\/$'"
    Else
        sCmd = "ls -1 -d -p " & sRemotePath & "/" & sFileNameOrSpec & " | grep -v '.*\/$'"
    End If
    
    If tfnRunRemoteCmd(sHostKey, rapPlink, sCmd, sResult, sErrMsg, sHost, sUser, sPWD) Then
        If Trim(sResult) = "" Then GoTo FINISHED
        If InStr(sErrMsg, "ls:") > 0 Then GoTo FINISHED
        files = Split(sResult, vbLf)
        count = UBound(files)
        ReDim aryFiles(count)
        For i = 0 To count
            If Trim(files(i)) <> "" Then
                aryFiles(j) = files(i)
                j = j + 1
            End If
        Next
        fnGetRemoteFileList = j
    Else
        fnGetRemoteFileList = -1
    End If
    
    Err.Clear
FINISHED:
    If Err.Number <> 0 Then
        sErrMsg = Err.Number & ": " & Err.Description
        Err.Clear
    End If
End Function

Public Function fnConvertToWindowsEOL(ByVal sFile As String, _
                                      sErrMsg As String, _
                                      ByVal bShowError As Boolean, _
                                      Optional ByVal sTargetPath As String = "") As Boolean
    On Error GoTo FINISHED
    Dim hFile As Integer
    Dim sBuffer As String
    Dim nError As Long
    
    If sTargetPath = "" Then
        sTargetPath = sFile
    End If
    
    If io.FileExists(sFile) Then
        hFile = FreeFile()
        Open sFile For Input As #hFile
        sBuffer = Input(LOF(hFile), hFile)
        Close #hFile
        hFile = 0
                
        If io.DeleteFile(sTargetPath) Then
            sBuffer = Replace(sBuffer, vbCrLf, vbLf)
            sBuffer = Replace(sBuffer, vbLf, vbCrLf)
            
            hFile = FreeFile()
            Open sTargetPath For Output As #hFile
            Print #hFile, sBuffer
            Close #hFile
            hFile = 0
        End If
        
        fnConvertToWindowsEOL = True
    Else
        Err.Raise -99, "fnConvertToWindowsEOL", "File not found."
    End If
    
    Err.Clear
FINISHED:
    If Err.Number <> 0 And Err.Number <> 62 Then
        sErrMsg = Err.Number & ": " & Err.Description
        If bShowError Then MsgBox "An error occurred converting end of line characters to CRLF for file: " & vbCrLf & vbCrLf _
                                 & sFile & vbCrLf & vbCrLf _
                                 & Err.Number & ": " & Err.Description
    End If
    nError = Err.Number
    On Error Resume Next
    If hFile <> 0 Then Close #hFile
    If nError = 62 Then
        If sFile <> sTargetPath Then io.CopyFile sFile, sTargetPath
        fnConvertToWindowsEOL = True
    End If
    Err.Clear
End Function

Public Function fnConvertToUnixEOL(ByVal sFile As String, _
                                   sErrMsg As String, _
                                   ByVal bShowError As Boolean, _
                                   Optional ByVal sTargetPath As String = "") As Boolean
    On Error GoTo FINISHED
    Dim hFile As Integer
    Dim sBuffer As String
    Dim nError As Long
    
    If sTargetPath = "" Then
        sTargetPath = sFile
    End If
    
    If io.FileExists(sFile) Then
        hFile = FreeFile()
        Open sFile For Input As #hFile
        sBuffer = Input(LOF(hFile), hFile)
        Close #hFile
        hFile = 0
                
        If io.DeleteFile(sTargetPath) Then
            sBuffer = Replace(sBuffer, vbCrLf, vbLf)
            
            hFile = FreeFile()
            Open sTargetPath For Output As #hFile
            Print #hFile, sBuffer
            Close #hFile
            hFile = 0
        End If
        
        fnConvertToUnixEOL = True
    Else
        Err.Raise -99, "fnConvertToUnixEOL", "File not found."
    End If
    
    Err.Clear
FINISHED:
    If Err.Number <> 0 And Err.Number <> 62 Then
        sErrMsg = Err.Number & ": " & Err.Description
        If bShowError Then MsgBox "An error occurred converting end of line characters to LF for file: " & vbCrLf & vbCrLf _
                                 & sFile & vbCrLf & vbCrLf _
                                 & Err.Number & ": " & Err.Description
    End If
    nError = Err.Number
    On Error Resume Next
    If Err.Number = 62 Then
        If sFile <> sTargetPath Then io.CopyFile sFile, sTargetPath
        fnConvertToUnixEOL = True
    End If
    If hFile <> 0 Then Close #hFile
    Err.Clear
End Function

Private Function Q_Str(ByVal Str As String, Optional ByVal Quote As String = """") As String
    Q_Str = Quote & Str & Quote
End Function

'Does not require the file to exist
Private Function fnFileNameFromPath(ByVal strFullPath As String, Optional ByVal sDriveSeparator As String = ":", Optional ByVal sFolderSeparator As String = "\") As String
    Dim sPathNoDrive As String
    
    'strFullPath has the potential to not contain a folder separator
    ' eg: "C:tmp.txt" refers to a file named "tmp.txt" in the current directory on drive C.
    ' see Microsoft's explanation "Fully Qualified vs. Relative Paths"
    '  https://msdn.microsoft.com/en-us/library/windows/desktop/aa365247(v=vs.85).aspx
    
    'Remove the disk designator
    If sDriveSeparator <> "" Then
        sPathNoDrive = Right(strFullPath, Len(strFullPath) - InStrRev(strFullPath, sDriveSeparator))
    Else
        sPathNoDrive = strFullPath
    End If
    
    'Return the filename
    If sFolderSeparator <> "" Then
        fnFileNameFromPath = Right(sPathNoDrive, Len(sPathNoDrive) - InStrRev(sPathNoDrive, sFolderSeparator))
    Else
        fnFileNameFromPath = sPathNoDrive
    End If

End Function

Private Function fnInsenseString(ByVal vTemp As Variant) As String
    Dim nIdx As Integer
    Dim szParameter As String
    Dim sTemp As String
    
    sTemp = CStr(vTemp)
    
    For nIdx = 1 To Len(sTemp)
        szParameter = szParameter & fnInsenseChar(Mid(sTemp, nIdx, 1))
    Next
    
    fnInsenseString = szParameter
End Function

Private Function fnInsenseChar(ByVal character As String) As String
    If IsLetter(character) Then
        fnInsenseChar = "[" & UCase(character) & LCase(character) & "]"
    Else
        fnInsenseChar = character
    End If
End Function

Private Function IsLetter(character As String) As Boolean
    IsLetter = IsCharAlphaW(AscW(character))
End Function

Private Function fnCombinePath(ByVal sDir As String, _
                               ByVal sFile As String, _
                               Optional ByVal bUnixOS As Boolean = False) As String
    If sDir <> "" And sFile <> "" Then
        If Not bUnixOS Then
            If Left(sFile, 1) = "\" Then sFile = Mid(sFile, 2)
            If Right$(sDir, 1) <> "\" Then
                fnCombinePath = sDir & "\" & sFile
            Else
                fnCombinePath = sDir & sFile
            End If
        Else
            If Left(sFile, 1) = "/" Then sFile = Mid(sFile, 2)
            If Right$(sDir, 1) <> "/" Then
                fnCombinePath = sDir & "/" & sFile
            Else
                fnCombinePath = sDir & sFile
            End If
        End If
    Else
        fnCombinePath = sDir & sFile
    End If
    
End Function

Private Function fnGetFileName(ByVal sFilePath As String) As String
    Dim arr() As String
    
    sFilePath = Trim(sFilePath)
    If sFilePath = "" Then
        Exit Function
    End If
    
    If InStr(1, sFilePath, "\") > 0 Then
        arr = Split(sFilePath, "\")
    Else
        arr = Split(sFilePath, "/")
    End If
    
    If UBound(arr) > -1 Then
        fnGetFileName = arr(UBound(arr))
    End If
    
End Function

Private Function fnNewDestFilename(ByVal sDestFilename As String, ByVal fCnt As Long) As String
    Dim idx As Long
    Dim sExtn As String
        
    fnNewDestFilename = sDestFilename
    If sDestFilename = "" Or fCnt = 0 Then Exit Function
    
    idx = InStrRev(sDestFilename, ".")
    If idx > 0 Then
        sExtn = Mid(sDestFilename, idx)
        sDestFilename = Mid(sDestFilename, 1, idx - 1)
    End If
        
    fnNewDestFilename = sDestFilename & fCnt & sExtn
    
End Function

Private Function fnNewDestFileWildcard(ByVal sDestFilename As String) As String
    Dim idx As Long
    Dim sExtn As String
    
    fnNewDestFileWildcard = sDestFilename
    If sDestFilename = "" Then Exit Function
    
    idx = InStrRev(sDestFilename, ".")
    If idx > 0 Then
        sExtn = Mid(sDestFilename, idx)
        sDestFilename = Mid(sDestFilename, 1, idx - 1)
    End If
        
    fnNewDestFileWildcard = sDestFilename & "*" & sExtn
    
End Function

Private Function fnBuildRemoveFileCmdArray(aryFiles() As String, _
                                           colFiles As Collection, _
                                           aryCmds() As String, _
                                           sErrMsg As String) As Long
    On Error GoTo FINISHED
    Dim cnt As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim sCmd As String
    Dim sFile As String
    
    cnt = UBound(aryFiles)
    ReDim aryCmds(cnt)
    
    For i = 0 To cnt
        sFile = fnGetFileName(aryFiles(i))
        If sFile <> "" Then
            If fnIsKeyInCollection(colFiles, sFile) Then
                sCmd = sCmd & "rm " & Q_Str(aryFiles(i)) & "; "
                k = k + 1
                If k > 19 Then
                    aryCmds(j) = sCmd
                    j = j + 1
                    sCmd = ""
                    k = 0
                End If
            End If
        End If
    Next
        
    If sCmd <> "" Then
        aryCmds(j) = sCmd
        j = j + 1
        sCmd = ""
        k = 0
    End If
    
    fnBuildRemoveFileCmdArray = j
    
    Err.Clear
FINISHED:
    If Err.Number <> 0 Then
        sErrMsg = "An error occurred building the server file removal list: " & _
                  Err.Number & ": " & Err.Description
        fnBuildRemoveFileCmdArray = 0
    End If
End Function

Private Function fnAddKeyToCollection(col As Collection, ByVal sKey As String, ByVal sValue As String) As Boolean
    
    If col Is Nothing Then Set col = New Collection
    
    sKey = "k" & sKey
    
    On Error Resume Next
    Err.Clear
    col.Add sValue, sKey
    fnAddKeyToCollection = Err.Number = 0
    Err.Clear
    
End Function

Private Function fnIsKeyInCollection(col As Collection, ByVal sKey As String) As Boolean
    Dim sValue As String
    
    If col Is Nothing Then Set col = New Collection
    
    sKey = "k" & sKey
    
    On Error Resume Next
    Err.Clear
    sValue = col(sKey)
    fnIsKeyInCollection = Err.Number = 0
    Err.Clear
    
End Function

'Private Sub CleanupFtpWorkFolder(ByVal sFolder As String)
'    On Error GoTo FINISHED
'    Dim sFile As String
'    Dim dtMin As Date
'    Dim dtFile As Date
'    Dim lDiff As Long
'
'    If Not io.DirectoryExists(sFolder) Then GoTo FINISHED
'
'    dtMin = Now
'    dtMin = DateAdd("h", -1, dtMin)
'
'    If Right(sFolder, 1) = "\" Then
'        sFile = Dir(sFolder & "*.*")
'    Else
'        sFile = Dir(sFolder & "\*.*")
'    End If
'
'    On Error Resume Next
'    While sFile <> ""
'        sFile = fnCombinePath(sFolder, sFile)
'        dtFile = FileDateTime(sFile)
'        lDiff = Abs(DateDiff("h", dtFile, dtMin))
'        If lDiff > 1 Then
'            io.DeleteFile sFile, False
'        End If
'        sFile = Dir()
'    Wend
'
'FINISHED:
'    Err.Clear
'
'End Sub

' *********************************************************
' Using Dir so don't call this if you are using Dir as well.
'
Public Function fnGetLocalFileList(ByVal sFileNameOrSpec As String, _
                                   ByVal sLocalPath As String, _
                                   ByRef aryFiles() As String, _
                                   ByRef sErrMsg As String) As Long
    On Error GoTo FINISHED
    Dim count As Long
    Dim files() As String
    Dim sCmd As String
    Dim sFile As String
    Dim sBuffer As String
    Dim cnt As Long
    Dim i As Long
    Dim j As Long
    
    ReDim aryFiles(0)
    
    If sLocalPath = "" Then
        sErrMsg = "Empty Local Path passed to fnGetLocalFileList"
        GoTo FINISHED
    End If
    
    If sFileNameOrSpec = "" Then
        sErrMsg = "Empty File Name Or Spec passed to fnGetLocalFileList"
        GoTo FINISHED
    End If
    
    If Right(sLocalPath, 1) <> "\" Then sLocalPath = sLocalPath & "\"
    'sCmd = Q_Str(sLocalPath & sFileNameOrSpec)
    sCmd = sLocalPath & sFileNameOrSpec
    
    sFile = Dir(sCmd)
    While sFile <> ""
        If sBuffer <> "" Then
            sBuffer = sBuffer & "|" & sFile
        Else
            sBuffer = sFile
        End If
        sFile = Dir()
    Wend
    
    If sBuffer = "" Then Exit Function
    files = Split(sBuffer, "|")
        
    cnt = UBound(files)
    If cnt < 0 Then Exit Function
        
    ReDim aryFiles(cnt)
    
    For i = 0 To cnt
        sFile = files(i)
        If sFile <> "" Then
            aryFiles(j) = sLocalPath & sFile
            j = j + 1
        End If
    Next
    
    fnGetLocalFileList = j
    
    Err.Clear
FINISHED:
    If Err.Number <> 0 Then
        sErrMsg = Err.Number & ": " & Err.Description
        Err.Clear
    End If
End Function

Private Sub fnDeleteFtpWorkFiles(aryCmds() As String)
    On Error Resume Next
    Dim cnt As Long
    Dim i As Long
    
    cnt = UBound(aryCmds)
    For i = 0 To cnt
        io.DeleteFile aryCmds(i), False
    Next
    
    Err.Clear
End Sub

Private Function fnGetProgramName() As String
    Dim sAppName As String
        
    On Error GoTo FINISHED
    fnGetProgramName = UCase(App.EXEName)
    
FINISHED:
    Err.Clear
    On Error Resume Next
    If fnGetProgramName = "" Then fnGetProgramName = "UNKNOWN"
    Err.Clear
End Function

'Private Function fnRemoveDirectory(ByVal sFolder As String, sErrMsg As String) As Boolean
'    On Error GoTo FINISHED
'
'    If Right(sFolder, 1) = "\" Then sFolder = Left(sFolder, Len(sFolder) - 1)
'
'    If io.DirectoryExists(sFolder) Then
'        If fnClearDirectory(sFolder, sErrMsg) Then
'            RmDir sFolder
'        Else
'            GoTo FINISHED
'        End If
'    End If
'
'    fnRemoveDirectory = True
'
'    Err.Clear
'FINISHED:
'    If Err.Number <> 0 Then
'        sErrMsg = "Error removing directory '" & sFolder & "': " & vbCrLf & _
'                  Err.Number & ": " & Err.Description
'        Err.Clear
'    End If
'End Function

' WARNING! Uses Dir(). This will mess up any other open Dir() searches.
Private Function fnClearDirectory(ByVal sFolder As String, sErrMsg As String) As Boolean
    On Error GoTo FINISHED
    Dim sFile As String
    
    If Trim(sFolder) <> "" Then
        If Right(sFolder, 1) = "\" Then sFolder = Left(sFolder, Len(sFolder) - 1)
        If io.DirectoryExists(sFolder) Then
            sFile = Dir(sFolder & "\*")
            If sFile <> "" Then
                Kill sFolder & "\*"
            End If
        End If
    End If
    
    fnClearDirectory = True
    
    Err.Clear
FINISHED:
    If Err.Number = 53 Then
        Err.Clear
        fnClearDirectory = True
    ElseIf Err.Number <> 0 Then
        sErrMsg = "Error clearing directory '" & sFolder & "': " & vbCrLf & _
                  Err.Number & ": " & Err.Description
        Err.Clear
    End If
End Function
