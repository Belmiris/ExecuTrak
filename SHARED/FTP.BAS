Attribute VB_Name = "FTP"
Option Explicit

'*******************************************************************************************
'Created    : 26th September 2000                                                          *
'Created By : Rajneesh Aggarwal                                                            *
'DLL Used   : WININET.DLL                                                                  *
'Usage      : fnTransferFile("Server", "UserID", "PassWord", True, True, _                 *
'           :                "RemoteFilePath&Name", "LocalFilePath", "ErrorMessage", True) *
'*******************************************************************************************

Declare Function GetProcessHeap Lib "kernel32" () As Long
Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
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

Private Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" _
    (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
' Initializes an application's use of the Win32 Internet functions
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" _
(ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, _
ByVal sProxyBypass As String, ByVal lFlags As Long) As Long

' User agent constant.
Private Const scUserAgent = "vb wininet"

' Use registry access settings.
Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Private Const INTERNET_OPEN_TYPE_DIRECT = 1
Private Const INTERNET_OPEN_TYPE_PROXY = 3
Private Const INTERNET_INVALID_PORT_NUMBER = 0

Private Const FTP_TRANSFER_TYPE_ASCII = &H1
Private Const FTP_TRANSFER_TYPE_BINARY = &H1
Private Const INTERNET_FLAG_PASSIVE = &H8000000

' Opens a HTTP session for a given site.
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" _
(ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, _
ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, _
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

Public Function fnTransferFile(sHost As String, sUser As String, sPwd As String, _
                                bPassive As Boolean, bBinary As Boolean, _
                                sRemotePathAndFile As String, _
                                sLocalPath As String, sRetErrMsg As String, _
                                Optional bShowError As Boolean = True) As Boolean
    
    Dim hOpen As Long, hConnection As Long, hFind As Long
    Dim lBinAscii As Long, lActPassive As Long
    Dim bFlag As Boolean, i As Integer
    Dim sRemotePath As String, sRemoteFile As String, sItemName As String
    Dim pData As WIN32_FIND_DATA
    Dim lLastDLLError As Long
    Dim sArrFile() As String
    
    fnTransferFile = False
    sLocalPath = Trim(sLocalPath)
    
    hOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
    If hOpen = 0 Then
        fnErrorOut Err.LastDllError, "InternetOpen", bShowError, sRetErrMsg
        GoTo CloseHandle
    End If
    
    IIf bPassive, lActPassive = INTERNET_FLAG_PASSIVE, lActPassive = 0
    
    hConnection = InternetConnect(hOpen, sHost, INTERNET_INVALID_PORT_NUMBER, sUser, sPwd, _
                                                          INTERNET_SERVICE_FTP, lActPassive, 0)
    If hConnection = 0 Then
        fnErrorOut Err.LastDllError, "InternetConnect", bShowError, sRetErrMsg
        GoTo CloseHandle
    End If
    
    IIf bBinary, lBinAscii = FTP_TRANSFER_TYPE_BINARY, lBinAscii = FTP_TRANSFER_TYPE_ASCII
    
    'The remote file specified should be valid....
    If Trim(sRemotePathAndFile) = "" Then
        sRetErrMsg = "Remote file path is not specified, cannot continue!"
        subShowError bShowError, sRetErrMsg, vbExclamation
        GoTo CloseHandle
    End If
    
    'Separate the file path and file name...
    sRemoteFile = fnGetRemoteFile(sRemotePathAndFile)
    sRemotePath = fnGetRemotePath(sRemotePathAndFile)
    
    'Validate the localpath...
    On Error GoTo errTrap
    If Dir(sLocalPath, vbDirectory) = "" Or sLocalPath = "" Then
        sRetErrMsg = "Invalid Local path specified!"
        subShowError bShowError, sRetErrMsg, vbExclamation
        GoTo CloseHandle
    End If
    
    subAddSlash sLocalPath
    
    'Set the current directory path...
    bFlag = FtpSetCurrentDirectory(hConnection, sRemotePath)
    If bFlag = False Then
        fnErrorOut Err.LastDllError, "FtpSetCurrentDirectory", bShowError, sRetErrMsg
        GoTo CloseHandle
    End If
    
    pData.cFileName = String(MAX_PATH, 0)
    
    hFind = FtpFindFirstFile(hConnection, sRemoteFile, pData, 0, 0)
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
    
    If Trim(sItemName) = "" Then
        GoTo CloseHandle
    End If
    
    sArrFile = Split(sItemName, ",")
    
    'Transfer the file to the local path specified...
    For i = 0 To UBound(sArrFile)
        bFlag = FtpGetFile(hConnection, sArrFile(i), sLocalPath & sArrFile(i), False, INTERNET_FLAG_RELOAD, lBinAscii, 0)
        If bFlag = False Then
            fnErrorOut Err.LastDllError, "FtpGetFile", bShowError, sRetErrMsg
            GoTo CloseHandle
        End If
    Next i
    
    InternetCloseHandle hConnection
    InternetCloseHandle hOpen
    InternetCloseHandle hFind
    
    fnTransferFile = True
    Exit Function
    
CloseHandle:
    If hConnection > 0 Then InternetCloseHandle hConnection
    If hOpen > 0 Then InternetCloseHandle hOpen
    If hFind > 0 Then InternetCloseHandle hFind
    Exit Function
errTrap:
    sRetErrMsg = Err.Number & ": " & Err.Description
    subShowError bShowError, sRetErrMsg, vbCritical
    GoTo CloseHandle
End Function

Private Function fnErrorOut(dError As Long, szCallFunction As String, _
                            bShowError As Boolean, sErrorMsg As String) As Boolean
    Dim dwIntError As Long
    Dim dwLength As Long
    Dim strBuffer As String
    
    fnErrorOut = True
    sErrorMsg = ""
    
    If dError = ERROR_INTERNET_EXTENDED_ERROR Then
        InternetGetLastResponseInfo dwIntError, vbNullString, dwLength
        strBuffer = String(dwLength + 1, 0)
        InternetGetLastResponseInfo dwIntError, strBuffer, dwLength
        sErrorMsg = szCallFunction & vbCrLf & "Err: " & dwIntError & " " & strBuffer
        subShowError bShowError, sErrorMsg, vbCritical
        Exit Function
    End If
    
    Select Case szCallFunction
        Case "InternetOpen"
            sErrorMsg = "Failed to access the DLL"
        Case "InternetConnect"
            sErrorMsg = "Failed to connect to the host server"
        Case "FtpSetCurrentDirectory"
            sErrorMsg = "Invalid path, Failed to set the current directory"
        Case "FtpFindFirstFile"
            sErrorMsg = "Failed to get the file from the server"
        Case "InternetFindNextFile"
            sErrorMsg = "Failed to get the file from the server"
        Case "FtpGetFile"
            sErrorMsg = "Failed to transfer the file"
    End Select
    
    subShowError bShowError, sErrorMsg, vbCritical
    fnErrorOut = False
    
End Function

Private Function fnIsWholePath(sPath As String) As Boolean
    If InStr(sPath, ":") > 0 Or Left(sPath, 1) = "\" Then
        fnIsWholePath = True
    Else
        fnIsWholePath = False
    End If
End Function

Private Function fnGetRemoteFile(sPath As String) As String
    Dim nPos As Integer
    
    fnGetRemoteFile = ""
    nPos = InStrRev(sPath, "/")
    fnGetRemoteFile = Mid(sPath, nPos + 1)
    
End Function

Private Sub subAddSlash(sStr As String)
    If sStr <> "" Then
        If Right(sStr, 1) <> "\" Then
            sStr = sStr & "\"
        End If
    End If
End Sub

Private Function fnGetRemotePath(sPath As String) As String
    Dim nPos As Integer
    
    fnGetRemotePath = ""
    nPos = InStrRev(sPath, "/")
    fnGetRemotePath = Mid(sPath, 1, nPos)

End Function

Private Sub subShowError(bShow As Boolean, sErrMsg As String, vbType As VbMsgBoxStyle)
    If bShow Then
        MsgBox sErrMsg, vbType
    End If
End Sub
