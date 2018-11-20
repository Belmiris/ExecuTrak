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
'G9222 - tthompson
'This class handles secure FTP using WinSCP
'   Most auxillary functions are called from FTP.bas
'
' WinSCP automation uses WinSCPnet.tlb references
    'Set objWinSCP_Session = CreateObject("WinSCP.Session")
    
' MINIMUM REQUIREMENTS:
'   .Net Framework 3.5
'   .Net Framework 4.0
'   WinSCP Automation in Local Folder
'
'WinSCP homepage is https://winscp.net/
'
'This is the README file for automation package of WinSCP.
'
'For use of WinSCP automation package see
'https://winscp.net/eng/docs/library
'https://winscp.net/eng/docs/library_install
'
'DEVELOPER... To use the WinSCP assembly via COM interop, register it using:
'%WINDIR%\Microsoft.NET\Framework\<version>\RegAsm.exe WinSCPnet.dll /codebase /tlb
'where <version> is typically either v4.0.30319 or v2.0.50727.
'https://winscp.net/eng/docs/library_install#registering
'
'USER?... To use the WinSCP assembly, register it using:
'%WINDIR%\Microsoft.NET\Framework\<version>\RegAsm.exe WinSCPnet.dll /tlb
'where <version> is typically either v4.0.30319 or v2.0.50727.
'
'See the file 'license-dotnet.txt' for the license conditions.

Private objWinSCP_Session As Object

Private m_Error As String

'perhaps not necessary since sRetErrMsg is passed as ByRef to fnSTransferFile()
Public Property Get LastError() As String
    LastError = m_Error
End Property

Private Sub Class_Initialize()
    
    If Not TestInstallation Then
        Err.Raise -1, "WinSCP Automation ", m_Error
    End If
    m_Error = ""
End Sub

'Perhaps not neccessary since fnSTransferFile() is only typically called once
Private Sub Class_Terminate()
    On Error Resume Next
    
    ' Disconnect, clean up
    If objWinSCP_Session.Opened Then objWinSCP_Session.Close
    Set objWinSCP_Session = Nothing
End Sub

' if this succeeds then we have a WinSCP object ready to use
Private Function TestInstallation() As Boolean
    
    TestInstallation = InstantiateWinSCP
    If Not TestInstallation Then
        TestInstallation = CheckMinimumRequirements
        If TestInstallation Then
            TestInstallation = InstantiateWinSCP
        End If
    End If
    
End Function

' if this succeeds then we have a WinSCP object ready to use
Private Function InstantiateWinSCP() As Boolean

    ' Attempt to instantiate WinSCP Automation
    On Error Resume Next
    Set objWinSCP_Session = CreateObject("WinSCP.Session")
    If Err.Number <> 0 Then
        m_Error = Err.Description
        Err.Clear
        Exit Function
    End If
    
    On Error GoTo 0
    InstantiateWinSCP = True
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
    
    
    'Going to use Late-Binding to hopefully circumvent any 3rd Party version change issues that might occur
    '    this has occured for other controls in the past and Late-Binding was chosen as the best method to reduce chance for error.
    
    'Late VS Early Binding
    '    Early binding is the method of referencing an ActiveX control in Project > References...
    '       This method provides Intellisense
    '       Properties and Methods are called via the v-table
    '    Late binding is the method of calling CreateObject()
    '       This method does not provide Intellisense
    '       Properties and Methods are called via IDispatch
    
    Dim objWinSCP_SessionOptions As Object 'SessionOptions
    Dim objWinSCP_TransferOptions As Object 'TransferOptions
    Dim objWinSCP_Transfer As Object 'TransferEventArgs
    Dim objWinSCP_TransferResult As Object 'TransferOperationResult
    Dim objWinSCP_DirectoryInfo As Object 'RemoteDirectoryInfo
    Dim objWinSCP_FileInfo As Object 'RemoteFileInfo
    
    fnSTransferFile = False
    
    ' first check the paths are ok
    sDestPath = Trim(sDestPath)
    sDestFilename = Trim(sDestFilename)
    sSourcePathAndFile = Trim(sSourcePathAndFile)
    sModule = fnGetProgramName()
    
    If sDestFilename <> "" Then
        If InStr(1, sDestFilename, "*") > 0 Or InStr(1, sDestFilename, "?") > 0 Then
            sRetErrMsg = "Invalid destination path specified, cannot continue!"
            GoTo CloseHandle
        End If
    End If
    
    If sSourcePathAndFile = "" Then
        sRetErrMsg = "Source file path is not specified, cannot continue!"
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
    End If
    
    If Not fnClearDirectory(sWorkDir, sRetErrMsg) Then
        GoTo CloseHandle
    End If
    
    On Error Resume Next
    
    ' Setup session options
    Set objWinSCP_SessionOptions = CreateObject("WinSCP.SessionOptions")
    With objWinSCP_SessionOptions
        .Protocol = Protocol_Sftp
        .HostName = sHost
        .userName = sUser
        .password = sPWD
        .SshHostKeyFingerprint = sHostKey
    End With

    If objWinSCP_Session Is Nothing Then InstantiateWinSCP
    objWinSCP_Session.Open objWinSCP_SessionOptions

    If Err.Number <> 0 Then
        sRetErrMsg = "Error opening session: " & Err.Description
        Err.Clear
        GoTo CloseHandle
    End If
    
    On Error GoTo ErrTrap
    
    'START WINSCP CODE
    Set objWinSCP_TransferOptions = CreateObject("WinSCP.TransferOptions")
        objWinSCP_TransferOptions.TransferMode = TransferMode_Binary
        objWinSCP_TransferOptions.PreserveTimestamp = False
    
    If bFromUnixToLocal Then
    
        ' ******* NOT IMPLEMENTED ***********
        '
        ' G9222 - This section was not tested or fully resolved for WinSCP because G9222 was only implemented for WSTRNEDI UPLOADING
        '
        Err.Raise 999, "WinSCP", "Downloading via WinSCP (SFTP) has not yet been Implemented"
        
        ' the download logic below was written with WinSCP in mind but not yet tested:
        '   not sure about the following:
        '       objWinSCP_Session.ListDirectory() loop
        '       delete work/source files (see implemented WinSCP upload logic for reference)
        ' ******* NOT IMPLEMENTED ***********
            
            
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
        
        Set directoryInfo = objWinSCP_Session.ListDirectory(sFindPath)
        
        For Each fileInfo In directoryInfo.files
            If fileInfo.Name Like sFindName Then fCnt = fCnt + 1
        Next
        
        If fCnt < 1 Then
            sRetErrMsg = "No file(s) found to transfer!"
            GoTo CloseHandle
        ElseIf sRetErrMsg <> "" Then
            If InStr(sRetErrMsg, "ls:") > 0 Then
                sRetErrMsg = "No file(s) found to transfer!"
                GoTo CloseHandle
            Else
                sRetErrMsg = "Error getting files from UNIX: " & vbCrLf _
                           & sRetErrMsg
                GoTo CloseHandle
            End If
        End If
        
        sInsSrcSpec = fnCombinePath(sSourcePath, sInsSrcSpec, True)
        
        lTime = Timer
            
        ' download files
        Set transferResult = _
            objWinSCP_Session.GetFiles(sInsSrcSpec, sWorkDir, bDeleteSourceFile, ftpTransferOptions)
    
        On Error Resume Next
        ' Throw on any error
        transferResult.Check
        
        If Err.Number <> 0 Then
            sRetErrMsg = "Error transfering files: " & vbCrLf _
                       & sCmd & vbCrLf _
                       & sRetErrMsg
            Err.Clear
            GoTo CloseHandle
        End If
        
        On Error GoTo ErrTrap
        
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
                    GoTo CloseHandle
                End If
                
            End If
            
        Next
        
        If Not fnSTransferFile Then
            ' set error msg to ?
            sRetErrMsg = "Error transfering files." & vbCrLf _
                & sRetErrMsg
            GoTo CloseHandle
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
            GoTo CloseHandle
        End If
        
        cCnt = 0
        ReDim aryCmds(fCnt - 1)
        
        If sDestFilename <> "" Then
            sTmpFile = fnGetFileName(sDestFilename)
        End If
        
        ' Copy files to ftpwork folder. Convert EOL if not binary.
        '  this will also match aryCmds() [the new filename] to aryFiles() [the source name]
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
                        GoTo CloseHandle
                    End If
                Else
                    If Not fnConvertToUnixEOL(aryFiles(i), sRetErrMsg, bShowError, sDestFile) Then
                        sRetErrMsg = "Failed to convert EOL format for file: " & vbCrLf & _
                                     "From '" & aryFiles(i) & "' " & vbCrLf & _
                                     "To: '" & sDestFile & "'" & vbCrLf & _
                                     sRetErrMsg
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
        
                
        ' winscp upload files
        Set objWinSCP_TransferResult = _
            objWinSCP_Session.PutFiles(sHold, sDestPath, True, objWinSCP_TransferOptions)
    
        On Error Resume Next
        ' Throw on any error
        objWinSCP_TransferResult.Check
        
        If Err.Number <> 0 Then
            sRetErrMsg = "Error transfering files: " & vbCrLf _
                       & Err.Description & vbCrLf _
                       & sRetErrMsg
            Err.Clear
            GoTo CloseHandle
        Else
            fnSTransferFile = True
        End If
        
        On Error GoTo ErrTrap
                
        'i = objWinSCP_TransferResult.Transfers.count
                            
        'is bDeleteSourceFile = TRUE?
        'at this point we have:
        '   an array containing the source files - aryFiles
        '   an array containing the respective work files - aryCmds
        '   a collection and count of the transferred work files
        '     if a work file was transferred then it was also deleted during the transfer
        '     we could check if the file in aryCmds exists (succeeded) and then delete the source if bDeleteSourceFile is TRUE
        If bDeleteSourceFile And fnSTransferFile Then
            
            For i = 0 To (fCnt - 1)
                If Not io.FileExists(aryCmds(i)) Then
                    If Not io.DeleteFile(aryFiles(i), bShowError) Then
                        sRetErrMsg = "Failed to delete file: " & vbCrLf & _
                                     "'" & aryFiles(i) & "' "
                        'GoTo CloseHandle ' keep going?
                    End If
                End If
            Next
        End If
    End If
    
CloseHandle:
    If Not objWinSCP_Session Is Nothing Then
        If objWinSCP_Session.Opened Then objWinSCP_Session.Close
        Set objWinSCP_Session = Nothing
    End If
    
    If sRetErrMsg <> "" Then subShowError bShowError, sRetErrMsg, vbExclamation
    Exit Function
    
ErrTrap:
    sRetErrMsg = Err.Number & ": " & Err.Description
    Err.Clear
    GoTo CloseHandle
End Function

' There are 2 minimum requirements before we can register the DLL
'  1) WinScp Automation installed (files exist on client because cannot register over network)
'  2) .Net Framework v2 installed (installing v3.5 also installs v2)
'
'  Business Logic for CheckMinimumRequirements():
'   1) Check for existence of "WinSCP Automation" DLL using OS's "ProgramFiles" variable...  the system will return the expected path of "[drive]:\Program Files" or "[drive]:\Program Files (x86)"
'       If Success, then move to next step
'       If Fail, then attempt to copy files from server (".\WinSCP Automation\") which is also "[drive]:\Factor\ExecTrak\Bin\WinSCP Automation\"
'           If Success, then move to next step
'           If Fail, then something is wrong, return message should indicate something useful for support
'   2) Check for .Net Framework v2 install in Registry
'       If Success, then we meet the minimum requirements
'       If Fail, then something is wrong, return message should indicate something useful for support
'   3) If we meet the minimum requirements, then register the DLL
'       If Success, then all is good
'       If Fail, then something is wrong, return message should indicate something useful for support
Private Function CheckMinimumRequirements() As Boolean
    Const NET20 As String = "2.0.50727"
    Const WinSCPFolderName As String = "WinSCP Automation"
    Dim sWinSCPSERVERPath As String
    Dim sWinSCPClientPath As String
    Dim sWinSCPClientDLL As String
    Dim bWinSCPDLLFound As Boolean
    Dim bNet20Found As Boolean
    Dim vVersion As Variant
    
    m_Error = ""
    
    'Step 1: Check for the existence of WinSCP Automation DLL
    sWinSCPSERVERPath = io.ApplicationPath & WinSCPFolderName & "\" 'WARNING - App.Path is borked for UNC paths
    sWinSCPClientPath = subBuildPath(False, "\", Environ$("ProgramFiles"), WinSCPFolderName)
    sWinSCPClientDLL = sWinSCPClientPath & "WinSCPnet.dll"
    
    bWinSCPDLLFound = io.FileExists(sWinSCPClientDLL)
    If Not bWinSCPDLLFound Then
        If Not io.DirectoryExists(sWinSCPSERVERPath) Then
            m_Error = m_Error & "WinSCP Automation Installation not found in: """ & sWinSCPSERVERPath & """!"
            GoTo Exit_Function
        End If
        
        If Not io.DirectoryExists(sWinSCPClientPath) Then MkDir sWinSCPClientPath
        vFiles = FileSystem.Dir(sWinSCPSERVERPath)
        Do While vFiles <> ""
            io.CopyFile sWinSCPSERVERPath & CStr(vFiles), sWinSCPClientPath & CStr(vFiles), True, False
            vFiles = FileSystem.Dir()
        Loop
        Err.Clear
        On Error GoTo 0
        
        bWinSCPDLLFound = io.FileExists(sWinSCPClientDLL)
        If Not bWinSCPDLLFound Then
            m_Error = m_Error & "WinSCP Automation Installation not found in: """ & sWinSCPClientDLL & """!"
            GoTo Exit_Function
        End If
    End If
    
    'Step 2: Enumerate the registry entries and locate .Net Framework v2
    For Each vVersion In GetDotNetFrameworkVersions
        If vVersion = NET20 Then
            bNet20Found = True
            Exit For
        End If
    Next
    If Not bNet20Found Then
        m_Error = m_Error & "Microsoft .Net v2 Installation not found!"
        GoTo Exit_Function
    End If

Exit_Function:
    CheckMinimumRequirements = bNet20Found And bWinSCPDLLFound
    
    If CheckMinimumRequirements Then CheckMinimumRequirements = RegisterWinSCPDLL(sWinSCPClientDLL)
    
End Function

Private Function RegisterWinSCPDLL(sPath As String) As Boolean
    Dim sCommand As String
    Dim sResult As String
    
    sCmd = """" & sPath & """ /codebase /tlb"
        
    RegisterWinSCPDLL = tfnRunRemoteCmd("DontNeedOne", _
                    rapRegAsm, _
                       sCmd, _
                        sResult, _
                       m_Error, _
                       bPromptContinueWait:=True)
End Function

' Return a collection of strings that identify the installed versions of the
' .NET Framework
' Note: this function require the EnumRegistryKeys and EnumRegistryValues
' functions
Function GetDotNetFrameworkVersions() As Collection
    Const HKEY_LOCAL_MACHINE = &H80000002
    Dim KeyName As String
    KeyName = "Software\Microsoft\.NETFramework\policy"

    Set GetDotNetFrameworkVersions = New Collection

    Dim versions As Collection, version As Variant
    ' get the collection of major.minor versions
    Set versions = EnumRegistryKeys(HKEY_LOCAL_MACHINE, KeyName)
    
    Dim builds As Collection, build As Variant
    ' for each version, get the installed builds, and finally get the complete
    ' version string
    For Each version In versions
        Set builds = EnumRegistryValues(HKEY_LOCAL_MACHINE, KeyName & "\" & version)
        For Each build In builds
            GetDotNetFrameworkVersions.Add (Mid(version, 2) & "." & build(0))
        Next
    Next
End Function

' I got this function working sufficiently fr testing if WinSCP would work but it does not follow the old logic implementation:
'   source and dest folder handling
'   use a "work" folder to handle EOL charaters (windows/unix)
'
'Public Function TestRun(ByVal sHostKey As String, _
'                                sHost As String, sUser As String, sPWD As String, _
'                                sSourcePathAndFile As String, _
'                                Optional bFromUnixToLocal As Boolean = True, _
'                                Optional sDestPath As String, _
'                                Optional sDestFilename As String, _
'                                Optional sRetErrMsg As String, _
'                                Optional bShowError As Boolean = True, _
'                                Optional bDeleteSourceFile As Boolean = False, _
'                                Optional bPassive As Boolean = True, _
'                                Optional bBinary As Boolean = True) As Boolean
'    On Error Resume Next
'    Init
'
'    sHostKey = "ssh-rsa 4096 IlZoz8HtMoIaudrq3H7muzyMOF+gtFkKT/17wFzRHbI="
'
'    ' Setup session options
'    Dim objWinSCP_SessionOptions As New SessionOptions
'    With objWinSCP_SessionOptions
'        .Protocol = Protocol_Sftp
'        .HostName = "upload.factor.com"
'        .userName = "tthompson"
'        .password = "Factor10"
'        .SshHostKeyFingerprint = sHostKey
'    End With
'
'    ' Now, with custom error handling enabled,
'    ' macro does not abort, when opening a session fails
'    objWinSCP_Session.Open objWinSCP_SessionOptions
'
'    ' Query for errors
'    If Err.Number <> 0 Then
'        m_Error = "Error opening session: " & Err.Description
'        GoTo clean_up
'    End If
'
'    ' Restore default error handling
'    On Error GoTo 0
'
'    ' Now, with default error handling restored,
'    ' macro aborts, if a reading remote directory fails
'    Dim directoryInfo As RemoteDirectoryInfo
'    Set directoryInfo = objWinSCP_Session.ListDirectory("/210")
'
'
'    ' Upload files
'    Dim ftpTransferOptions As New TransferOptions
'    ftpTransferOptions.TransferMode = TransferMode_Binary
'    ftpTransferOptions.PreserveTimestamp = False
'
'    Dim transferResult As TransferOperationResult
'    Set transferResult = _
'        objWinSCP_Session.PutFiles(sSourcePathAndFile, sDestPath & "/" & sDestFilename, False, ftpTransferOptions)
'
'    ' Throw on any error
'    transferResult.Check
'
'    ' Display results
'    Dim transfer As TransferEventArgs
'    For Each transfer In transferResult.Transfers
'        MsgBox "Upload of " & transfer.FileName & " succeeded"
'    Next
'
'clean_up:
'    ' Disconnect, clean up
'    Dispose
'End Function
