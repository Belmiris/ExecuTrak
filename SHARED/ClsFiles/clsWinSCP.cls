VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWinSCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' This class handles secure FTP using WinSCP
'   Most auxillary functions are called from FTP.bas
'
'WinSCP homepage is https://winscp.net/
'
'This is the README file for automation package of WinSCP.
'
'For use of WinSCP automation package see
'https://winscp.net/eng/docs/library
'https://winscp.net/eng/docs/library_install
'
'To use the WinSCP assembly via COM interop, register it using:
'%WINDIR%\Microsoft.NET\Framework\<version>\RegAsm.exe WinSCPnet.dll /codebase /tlb
'where <version> is typically either v4.0.30319 or v2.0.50727.
'https://winscp.net/eng/docs/library_install#registering
'
'See the file 'license-dotnet.txt' for the license conditions.

Private WithEvents ftpSession As WinSCPnet.Session
Attribute ftpSession.VB_VarHelpID = -1
Private m_Error As String

Public Property Get LastError() As String
    LastError = m_Error
End Property

Private Sub Init()
    
    ' Enable custom error handling
    On Error Resume Next
        
    If ftpSession Is Nothing Then
        Set ftpSession = New Session
    End If
    
    ' Query for errors
    If Err.Number <> 0 Then
        m_Error = "Error: " & Err.Description
 
        ' Clear the error
        Err.Clear
    End If
          
    ' Restore default error handling
    On Error GoTo 0
    
End Sub

Public Sub Dispose()
    On Error Resume Next
    
    ' Disconnect, clean up
    ftpSession.Dispose
End Sub

Public Function TestRun(ByVal sHostKey As String, _
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
    On Error Resume Next
    Init

    sHostKey = "ssh-rsa 4096 IlZoz8HtMoIaudrq3H7muzyMOF+gtFkKT/17wFzRHbI="
    
    ' Setup session options
    Dim ftpSessionOptions As New SessionOptions
    With ftpSessionOptions
        .Protocol = Protocol_Sftp
        .HostName = "upload.factor.com"
        .userName = "tthompson"
        .Password = "Factor10"
        .SshHostKeyFingerprint = sHostKey
    End With

    ' Now, with custom error handling enabled,
    ' macro does not abort, when opening a session fails
    ftpSession.Open ftpSessionOptions

    ' Query for errors
    If Err.Number <> 0 Then
        m_Error = "Error opening session: " & Err.Description
        GoTo clean_up
    End If

    ' Restore default error handling
    On Error GoTo 0

    ' Now, with default error handling restored,
    ' macro aborts, if a reading remote directory fails
    Dim directoryInfo As RemoteDirectoryInfo
    Set directoryInfo = ftpSession.ListDirectory("/210")


    ' Upload files
    Dim ftpTransferOptions As New TransferOptions
    ftpTransferOptions.TransferMode = TransferMode_Binary
    ftpTransferOptions.PreserveTimestamp = False

    Dim transferResult As TransferOperationResult
    Set transferResult = _
        ftpSession.PutFiles(sSourcePathAndFile, sDestPath & "/" & sDestFilename, False, ftpTransferOptions)

    ' Throw on any error
    transferResult.Check

    ' Display results
    Dim transfer As TransferEventArgs
    For Each transfer In transferResult.Transfers
        MsgBox "Upload of " & transfer.FileName & " succeeded"
    Next

clean_up:
    ' Disconnect, clean up
    ftpSession.Close
End Function

' same signature as old methods
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
'            subShowError bShowError, sRetErrMsg, vbExclamation
            m_Error = sRetErrMsg
            GoTo CloseHandle
        End If
    End If
    
    If sSourcePathAndFile = "" Then
        sRetErrMsg = "Source file path is not specified, cannot continue!"
'        subShowError bShowError, sRetErrMsg, vbExclamation
        m_Error = sRetErrMsg
        GoTo CloseHandle
    End If
    
    If sDestPath = "" Then
        Dim szDir As String
        If bFromUnixToLocal Then
            szDir = io.ApplicationPath
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
        m_Error = sRetErrMsg
'        subShowError bShowError, sRetErrMsg, vbExclamation
        Exit Function
    End If
    
    If Not fnClearDirectory(sWorkDir, sRetErrMsg) Then
'        subShowError bShowError, sRetErrMsg, vbExclamation
        m_Error = sRetErrMsg
        Exit Function
    End If
    
    On Error Resume Next
    Init
    
    ' Setup session options
    Dim ftpSessionOptions As New SessionOptions
    With ftpSessionOptions
        .Protocol = Protocol_Sftp
        .HostName = sHost
        .userName = sUser
        .Password = sPWD
        .SshHostKeyFingerprint = sHostKey
    End With

    ftpSession.Open ftpSessionOptions

    If Err.Number <> 0 Then
        sRetErrMsg = "Error opening session: " & Err.Description
        m_Error = sRetErrMsg
        Err.Clear
        GoTo CloseHandle
    End If
    
    On Error GoTo errTrap
        
    Dim ftpTransferOptions As New TransferOptions
    
    ftpTransferOptions.TransferMode = TransferMode_Binary
    ftpTransferOptions.PreserveTimestamp = False

    Dim transfer As TransferEventArgs
    Dim transferResult As TransferOperationResult
    Dim directoryInfo As RemoteDirectoryInfo
    Dim fileInfo As RemoteFileInfo
    
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
        
        ' Get remote files
        ' This will get us a snap-shot of files on the server.
        ' We must anticipate that more files may be added and some files
        ' may be removed before the file transfer is complete.
        sInsSrcSpec = fnInsenseString(sSourceFile)
        sFindName = Replace(sInsSrcSpec, " ", "' '")
        sFindName = Replace(sFindName, "(", "'('")
        sFindName = Replace(sFindName, ")", "')'")
        
        sFindPath = Replace(sSourcePath, " ", "' '")
        
        Set directoryInfo = ftpSession.ListDirectory(sFindPath)
        
        For Each fileInfo In directoryInfo.files
            If fileInfo.Name Like sFindName Then fCnt = fCnt + 1
        Next
        
        If fCnt < 1 Then
            sRetErrMsg = "No file(s) found to transfer!"
'            subShowError bShowError, sRetErrMsg, vbExclamation
            m_Error = sRetErrMsg
            GoTo CloseHandle
        ElseIf sRetErrMsg <> "" Then
            If InStr(sRetErrMsg, "ls:") > 0 Then
                sRetErrMsg = "No file(s) found to transfer!"
'                subShowError bShowError, sRetErrMsg, vbExclamation
                m_Error = sRetErrMsg
                GoTo CloseHandle
            Else
                sRetErrMsg = "Error getting files from UNIX: " & vbCrLf _
                           & sRetErrMsg
'                subShowError bShowError, sRetErrMsg, vbExclamation
                m_Error = sRetErrMsg
                GoTo CloseHandle
            End If
        End If
        
        sInsSrcSpec = fnCombinePath(sSourcePath, sInsSrcSpec, True)
        
        lTime = Timer
            
        ' download files
        Set transferResult = _
            ftpSession.GetFiles(sInsSrcSpec, sWorkDir, bDeleteSourceFile, ftpTransferOptions)
    
        On Error Resume Next
        ' Throw on any error
        transferResult.Check
        
        If Err.Number <> 0 Then
            sRetErrMsg = "Error transfering files: " & vbCrLf _
                       & sCmd & vbCrLf _
                       & sRetErrMsg
'            subShowError bShowError, sRetErrMsg, vbExclamation
            m_Error = sRetErrMsg
            GoTo CloseHandle
        End If
        
        On Error GoTo errTrap
        
        ReDim aryFiles(transferResult.Transfers.count)
        
        For Each transfer In transferResult.Transfers
            aryFiles(i) = transfer.FileName
        Next
        
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
                        m_Error = sRetErrMsg
                    End If
                Else
                    fnSTransferFile = io.CopyFile(sWorkFile, sDestFile)
                    If Not fnSTransferFile Then
                        sRetErrMsg = "Failed to copy ftp work file: '" & sWorkFile & "' " & vbCrLf _
                                   & "To: '" & sDestFile & "'"
                        m_Error = sRetErrMsg
                    End If
                End If
    
                If fnSTransferFile Then
                    fnSTransferFile = io.DeleteFile(sWorkFile)
                    If Not fnSTransferFile Then
                        sRetErrMsg = "Failed to delete ftp work file: '" & sWorkFile & "'"
                        m_Error = sRetErrMsg
                    End If
                End If
            
                If Not fnSTransferFile Then
                    If sRetErrMsg = "" Then sRetErrMsg = "Failed to copy file from ftp work folder"
'                    subShowError bShowError, sRetErrMsg, vbExclamation
                    m_Error = sRetErrMsg
                    GoTo CloseHandle
                End If
                
            End If
            
        Next
        
        If Not fnSTransferFile Then
            ' set error msg to ?
            Exit Function
        End If
        
    Else
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
'            subShowError bShowError, sRetErrMsg, vbExclamation
            m_Error = sRetErrMsg
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
'                        subShowError bShowError, sRetErrMsg, vbExclamation
                        m_Error = sRetErrMsg
                        GoTo CloseHandle
                    End If
                Else
                    If Not fnConvertToUnixEOL(aryFiles(i), sRetErrMsg, bShowError, sDestFile) Then
                        sRetErrMsg = "Failed to convert EOL format for file: " & vbCrLf & _
                                     "From '" & aryFiles(i) & "' " & vbCrLf & _
                                     "To: '" & sDestFile & "'" & vbCrLf & _
                                     sRetErrMsg
'                        subShowError bShowError, sRetErrMsg, vbExclamation
                        m_Error = sRetErrMsg
                        GoTo CloseHandle
                    End If
                End If
            End If
        Next
        
        ' We want to FTP from the ftpwork folder now so we know what is going
        ' across.
        If sTmpFile = "" Then
            sHold = fnCombinePath(sWorkDir, sSourceFile)
            'sCmd = Q_Str(sHold) & " " & sUser & "@" & sHost & ":" & Q_Str(sDestPath)
        Else
            sHold = fnNewDestFileWildcard(sTmpFile)
            sHold = fnCombinePath(sWorkDir, sHold)
            'sCmd = Q_Str(sHold) & " " & sUser & "@" & sHost & ":" & Q_Str(sDestPath)
        End If
        
        
        
        ' upload files
        Set transferResult = _
            ftpSession.PutFiles(sHold, sDestPath, bDeleteSourceFile, ftpTransferOptions)
    
        On Error Resume Next
        ' Throw on any error
        transferResult.Check
        
        If Err.Number <> 0 Then
            sRetErrMsg = "Error transfering files: " & vbCrLf _
                       & Err.Description & vbCrLf _
                       & sRetErrMsg
'            subShowError bShowError, sRetErrMsg, vbExclamation
            m_Error = sRetErrMsg
            Err.Clear
            GoTo CloseHandle
        End If
        
        On Error GoTo errTrap
        
        ReDim aryFiles(transferResult.Transfers.count)
        
        For Each transfer In transferResult.Transfers
            aryFiles(i) = transfer.FileName
        Next
        
        sHold = fnCombinePath(sDestPath, "*", True)
        
        fnDeleteFtpWorkFiles aryCmds
        
        If bDeleteSourceFile And fnSTransferFile Then
            
            For i = 0 To (fCnt - 1)
                If io.FileExists(aryFiles(i)) Then
                    If Not io.DeleteFile(aryFiles(i), bShowError) Then
                        sRetErrMsg = "Failed to delete file: " & vbCrLf & _
                                     "'" & aryFiles(i) & "' "
'                        subShowError bShowError, sRetErrMsg, vbExclamation
                        m_Error = sRetErrMsg
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
    'subShowError bShowError, sRetErrMsg, vbCritical
    m_Error = sRetErrMsg
    GoTo CloseHandle
End Function
