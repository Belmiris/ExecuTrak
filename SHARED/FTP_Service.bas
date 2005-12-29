Attribute VB_Name = "modFTP_Service"
Option Explicit

Public Const SQL_SEL_PARTNER_PROFILE  As String = "SELECT *" _
                                                & "  FROM sys_partnr_tran" _
                                                & " WHERE (ptrn_trnsfr_id='@TransferID')" _
                                                & "   AND (ptrn_partner_id='@PartnerID')"

Public Const SQL_SEL_TRANSFER_PROFILE As String = "SELECT *" _
                                                & "  FROM sys_trnsfr_prof" _
                                                & " WHERE (tprf_trnsfr_id='@TransferID')"
Public Type TransferActivityRecord
    TransferID      As String
    PartnerID       As String
    IP_Addr         As Variant
    IP_Name         As Variant
    PathName        As Variant
    FileName        As Variant
    Status          As Variant
    Retries         As Variant
    FunctionName    As Variant
    ReturnCode      As Variant
    BytesTransfered As Variant
End Type

Private Const INFINITE = &HFFFF&
Private Const SYNCHRONIZE = &H100000

Private Declare Function WaitForSingleObject Lib "kernel32" ( _
    ByVal hHandle As Long, _
    ByVal dwMilliseconds As Long _
) As Long

Private Declare Function OpenProcess Lib "kernel32" ( _
    ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long _
) As Long

Private Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As Long _
) As Long
Public Sub FTP_CreateBatchFile(ByVal BatchName As String, ByVal ScriptName As String, ProfileInfo As Collection)
    Const ProcName = "FTP_CreateBatchFile"
    Dim IsPut     As Boolean
    Dim hFile     As Integer
    Dim LocalFile As String
    Dim LocalPath As String
    
    '------------------------------------------------------------------------------------
    'Initialize
    '------------------------------------------------------------------------------------
    On Error GoTo errHandler
    IsPut = (ProfileInfo("tprf_getput_flag") = "P")
    LocalPath = FixPath(ProfileInfo("ptrn_int_path"))
    If LenB(Nz(ProfileInfo("ptrn_int_file"))) Then
        LocalFile = ProfileInfo("ptrn_int_file")
    Else
        LocalFile = ProfileInfo("ptrn_ext_file")
    End If
    
    '------------------------------------------------------------------------------------
    'Open Batch File
    '------------------------------------------------------------------------------------
    hFile = FreeFile()
    Open BatchName For Output As #hFile
    Print #hFile, "@Echo Off"
    
    'Change Directory
    Print #hFile, Left$(ProfileInfo("ptrn_int_path"), 2) 'Change Drive
    Print #hFile, "cd "; ProfileInfo("ptrn_int_path")    'Change Directory
    
    'Launch FTP with Script File
    Print #hFile, "ftp -s:"; Q_Str(ScriptName); " "; ProfileInfo("ptrn_ext_ip_addr")
    
    'Backup File
    If Nz(ProfileInfo("ptrn_int_bkup_flg")) = "Y" Then
        Print #hFile, "copy "; Q_Str(LocalFile); " ";
        Print #hFile, Q_Str(BackupFilename(LocalFile, ProfileInfo("ptrn_int_bkup_loc")))
    End If
    
    'Append File
    If Nz(ProfileInfo("ptrn_int_app_flg")) = "Y" Then
        Print #hFile, "copy /b "; Q_Str(ProfileInfo("ptrn_int_app_file"));
        Print #hFile, "+"; Q_Str(LocalFile); " "; Q_Str(ProfileInfo("ptrn_int_app_file"))
    End If
    
    'Delete/Move/Rename File
    If IsPut And (ProfileInfo("ptrn_delete_flg") = "Y") Then
        'Delete File - only applies if we are PUTting the file onto the Remote System
        Print #hFile, "del "; Q_Str(LocalFile)
    ElseIf (ProfileInfo("ptrn_int_mv_flg") = "Y") Then
        'Move File
        Print #hFile, "move "; Q_Str(LocalFile); " ";
        Print #hFile, Q_Str(ProfileInfo("ptrn_int_mov_loc"))
    ElseIf (ProfileInfo("ptrn_int_ren_flg") = "Y") Then
        'Rename File - MOVE can move across drives, paths, and even rename the file
        Print #hFile, "move "; Q_Str(LocalFile); " ";
        Print #hFile, Q_Str(ProfileInfo("ptrn_int_ren_file"))
    End If
    
    '------------------------------------------------------------------------------------
    'Finish up Batch File
    '------------------------------------------------------------------------------------
    Print #hFile, "del "; Q_Str(ScriptName) 'Delete the FTP Script File
    Close #hFile
    Exit Sub
    
errHandler:
    tfnErrHandler ProcName, False
    Err.Clear
End Sub
Public Sub FTP_CreateScriptFile(ByVal ScriptName As String, ProfileInfo As Collection)
    Const ProcName = "FTP_CreateScriptFile"
    Dim hFile      As Integer
    Dim LocalFile  As String
    Dim RemoteFile As String
    Dim RemotePath As String
    Dim IsPut      As Boolean
    
    '------------------------------------------------------------------------------------
    'Initialize
    '------------------------------------------------------------------------------------
    On Error GoTo errHandler
    IsPut = (ProfileInfo("tprf_getput_flag") = "P")
    If LenB(Nz(ProfileInfo("ptrn_ext_file"))) Then
        RemoteFile = Nz(ProfileInfo("ptrn_ext_file"))
    Else
        RemoteFile = Nz(ProfileInfo("ptrn_int_file"))
    End If
    If LenB(Nz(ProfileInfo("ptrn_int_file"))) Then
        LocalFile = Nz(ProfileInfo("ptrn_int_file"))
    Else
        LocalFile = Nz(ProfileInfo("ptrn_ext_file"))
    End If
    
    '------------------------------------------------------------------------------------
    'Open Script File
    '------------------------------------------------------------------------------------
    hFile = FreeFile()
    Open ScriptName For Output As #hFile
    
    'Sign on with UID & PWD
    Print #hFile, ProfileInfo("ptrn_ext_userid")
    Print #hFile, Decrypt(ProfileInfo("ptrn_ext_pwd"), True)
    
    'Set Transfer Mode to ASCII/Binary
    If ProfileInfo("ptrn_bin_asc_flg") = "B" Then
        Print #hFile, "binary"
    Else
        Print #hFile, "ascii"
    End If
    
    'Change Remote Directory
    If LenB(ProfileInfo("ptrn_ext_path")) Then
        Print #hFile, "cd "; ProfileInfo("ptrn_ext_path")
    End If
    
    'Get/Put File
    If IsPut Then
        RemotePath = ProfileInfo("ptrn_ext_path")
        If Right$(RemotePath, 1) <> "/" Then
            RemotePath = RemotePath & "/"
        End If
        'Delete existing remote file (just let it fail if it doesn't exist)
        Print #hFile, "delete "; RemotePath; RemoteFile
        'Put remote file
        Print #hFile, "put "; LocalFile; " "; RemoteFile
        'Change file permissions
        'If not sending to a UNIX box, then this should simply fail, but no harm done
        Print #hFile, "literal site chmod 777 "; RemotePath; RemoteFile
    Else
        Print #hFile, "get "; RemoteFile; " "; LocalFile
    End If
    
    'Delete Remote File
    If (Not IsPut) And (ProfileInfo("ptrn_delete_flg") = "Y") Then
        'Only applies if we are GETting file
        Print #hFile, "delete "; RemoteFile
    End If
    
    '------------------------------------------------------------------------------------
    'Finish up Script File
    '------------------------------------------------------------------------------------
    Print #hFile, "quit"
    Close #hFile
    Exit Sub
    
errHandler:
    tfnErrHandler ProcName, False
    Err.Clear
End Sub
Public Function FTP_GetTempBatchName(ByVal Path As String) As String
    Const ProcName = "FTP_GetTempBatchName"
    'Returns a unique filename for a batch file that is to conduct an FTP session.
    'Filename is in the format of "FTPnnnnn.BAT", where nnnnn is a 5-digit counter.
    Dim CurFile  As String
    Dim CurIndex As Long
    Dim NewIndex As Long
    
    On Error GoTo errHandler
    '------------------------------------------------------------------------------------
    'Get the next available 5-digit counter value
    '------------------------------------------------------------------------------------
    Path = FixPath(Path)
    CurFile = UCase$(Dir(Path & "FTP?????.BAT"))
    Do While LenB(CurFile)
        If CurFile Like "FTP#####.BAT" Then
            CurIndex = CLng(Mid$(CurFile, 4, 5))
            If CurIndex > NewIndex Then
                NewIndex = CurIndex
            End If
        End If
        
        CurFile = UCase$(Dir())
    Loop
    NewIndex = NewIndex + 1
    
    '------------------------------------------------------------------------------------
    'Generate the filename
    '------------------------------------------------------------------------------------
    FTP_GetTempBatchName = Path & "FTP" & Format$(NewIndex, "00000") & ".BAT"
    Exit Function
    
errHandler:
    FTP_GetTempBatchName = vbNullString
    tfnErrHandler ProcName, False
    Err.Clear
End Function
Public Function Decrypt(sSource As String, Optional ByVal SuppressError As Boolean = False) As String
    'This routine was taken from the SYFTRNPR & BREFTPBA programs, each of
    'which has their own distinct copy. It's been adapted and placed into
    'this standard code module so that these encrypt/decrypt routines can
    'be used in multiple projects and still maintain only one copy of these
    'routines.
    Const ProcName = "Decrypt"
    Dim i As Integer
    Dim nLen As Integer
    Dim sTemp As String
    Dim nAsc As Integer
    Dim sAry() As String
    
    sTemp = vbNullString
    i = 0
On Error GoTo ERROR_HANDLER
    
    sAry = Split(sSource, "-")
    For i = 0 To UBound(sAry)
        sAry(i) = Chr(sAry(i))
    Next
    sSource = Join(sAry, "")
    
    nLen = Len(sSource)
    For i = 3 To nLen
        nAsc = Asc(Mid(sSource, nLen - i + 3, 1))
        sTemp = sTemp & Chr(nAsc - 2 * nLen + i + 1)
    Next i
    
    Decrypt = sTemp
    Exit Function
    
ERROR_HANDLER:
    tfnErrHandler ProcName, False
    Err.Clear
    If Not SuppressError Then
        MsgBox "The password can not be decrypted. Please report to Factor Suppport", vbCritical
    End If
    Decrypt = vbNullString
End Function
Public Function Encrypt(sSource As String) As String
    'This routine was taken from the SYFTRNPR & BREFTPBA programs, each of
    'which has their own distinct copy. It's been adapted and placed into
    'this standard code module so that these encrypt/decrypt routines can
    'be used in multiple projects and still maintain only one copy of these
    'routines.
    Const ProcName = "Encrypt"
    Dim i As Integer
    Dim nLen As Integer
    Dim sTemp As String
    Dim nAsc As Integer
    Dim sEncrypt As String
    
    On Error GoTo errHandler
    sTemp = ""
    nLen = Len(sSource)
    If nLen < 2 Then
        sTemp = "a3"
    Else
        nAsc = Asc(Left(sSource, 1)) + nLen
        sTemp = Chr(nAsc \ 2)
        nAsc = Asc(Right(sSource, 1)) + nLen
        sTemp = sTemp & Chr(nAsc / 1.5)
    End If
    For i = 1 To nLen
        nAsc = Asc(Mid(sSource, nLen - i + 1, 1))
        sTemp = sTemp & Chr(nAsc + nLen + i)
    Next i
    
    sEncrypt = ""
    For i = 1 To Len(sTemp)
        If i = 1 Then
            sEncrypt = Asc(Mid(sTemp, i, 1))
        Else
            sEncrypt = sEncrypt & "-" & Asc(Mid(sTemp, i, 1))
        End If
    Next
    
    Encrypt = sEncrypt
    Exit Function
    
errHandler:
    tfnErrHandler ProcName, False
    Err.Clear
End Function
Public Sub FTP_LogActivity(XferRec As TransferActivityRecord)
    Const ProcName = "FTP_LogActivity"
    
    Dim SQL As String
    
    On Error GoTo errHandler
    With XferRec
        SQL = "INSERT INTO sys_trnsfr_acty VALUES(TODAY,CURRENT HOUR TO MINUTE," _
            & SQL_FieldValue(.TransferID, dbChar) & "," _
            & SQL_FieldValue(.PartnerID, dbChar) & "," _
            & SQL_FieldValue(.IP_Addr, dbChar) & "," _
            & SQL_FieldValue(.IP_Name, dbChar) & "," _
            & SQL_FieldValue(.PathName, dbChar) & "," _
            & SQL_FieldValue(.FileName, dbChar) & "," _
            & SQL_FieldValue(.Status, dbChar) & "," _
            & SQL_FieldValue(.Retries, dbNumeric) & "," _
            & SQL_FieldValue(.FunctionName, dbChar) & "," _
            & SQL_FieldValue(.ReturnCode, dbChar) & "," _
            & SQL_FieldValue(.BytesTransfered, dbNumeric) & ")"
        fnExecSQL SQL
    End With
    Exit Sub
    
errHandler:
    tfnErrHandler ProcName, SQL, False
    Err.Clear
End Sub
Public Sub FTP_ProcessRequest(ByVal TransferID As String, ByVal PartnerID As String)
    Const ProcName = "FTP_ProcessRequest"
    Dim ProfileInfo As Collection
    Dim BatchFile   As String
    Dim ScriptFile  As String
    Dim XferLogRec  As TransferActivityRecord
    
    On Error GoTo errHandler
    '------------------------------------------------------------------------------------
    'Get Partner/Transfer Profile Information
    '------------------------------------------------------------------------------------
    Set ProfileInfo = GetTransferProfileInfo(TransferID, PartnerID)
    
    '------------------------------------------------------------------------------------
    'Create FTP Batch and Script Files, and Execute Batch File
    '------------------------------------------------------------------------------------
    If ProfileInfo.Count Then
        '------------------------------------------------------------------------------------
        'Execute the PreProcess
        '------------------------------------------------------------------------------------
        If LenB(Nz(ProfileInfo("tprf_preproc"))) Then
            'Execute the Win/4GL Program
            FTP_ExecuteOutsideProcess ProfileInfo("tprf_preproc")
        End If
        
        '------------------------------------------------------------------------------------
        'Create and Execute Batch and Script Files
        '------------------------------------------------------------------------------------
        BatchFile = FTP_GetTempBatchName(AppPath())
        ScriptFile = Replace(BatchFile, ".bat", ".scr", , , vbTextCompare)
        FTP_CreateBatchFile BatchFile, ScriptFile, ProfileInfo
        FTP_CreateScriptFile ScriptFile, ProfileInfo
        
        'Execute Newly-Created Batch File
        FTP_ExecuteOutsideProcess BatchFile 'Does not return until execution has completed
        
        'Delete Batch File and Script File
        Kill BatchFile
        'Kill ScriptFile 'ScriptFile already being deleted from the Batch File
        
        '------------------------------------------------------------------------------------
        'Follow-Up Activities for Success
        '------------------------------------------------------------------------------------
        'Success is assumed since we're invoking FTP.EXE and we cannot tell whether the
        'actual file transfer was successful or not.
        
        'EMail Notification
        If LenB(Nz(ProfileInfo("tprf_notify_ok"))) Then
            FTP_EMailNotify ProfileInfo("tprf_notify_ok"), _
                            "File Transfer - " & PartnerID & "-" & TransferID, _
                            "File Transfer was completed successfully."
        End If
        
        'Execute the PostProcess for Success
        If LenB(Nz(ProfileInfo("tprf_postproc_ok"))) Then
            'Execute the Win/4GL Program
            FTP_ExecuteOutsideProcess ProfileInfo("tprf_postproc_ok")
        End If
        
        'Set FTP Activity Log Data for Success
        With XferLogRec
            .TransferID = TransferID
            .PartnerID = PartnerID
            .IP_Addr = ProfileInfo("ptrn_ext_ip_addr")
            .IP_Name = Null
            .PathName = ProfileInfo("ptrn_int_path")
            .FileName = ProfileInfo("ptrn_ext_file")
            .Status = "S"
            .Retries = 0
            .FunctionName = IIf(ProfileInfo("tprf_getput_flag") = "P", "PUT", "GET")
            .ReturnCode = "OK"
            .BytesTransfered = Null
        End With
    Else 'ProfileInfo.Count
        '------------------------------------------------------------------------------------
        'Follow-Up Activities for Failure
        '------------------------------------------------------------------------------------
        'FAILURE - No Partner/Transfer Profile Record
        
        'Set FTP Activity Log Data - Failed
        With XferLogRec
            .TransferID = TransferID
            .PartnerID = PartnerID
            .IP_Addr = Null
            .IP_Name = Null
            .PathName = Null
            .FileName = TransferID
            .Status = "F"
            .Retries = 0
            .FunctionName = Null
            .ReturnCode = "ZZZ"
            .BytesTransfered = 0
        End With
    End If 'ProfileInfo.Count
    
    '------------------------------------------------------------------------------------
    'Log the Results
    '------------------------------------------------------------------------------------
    FTP_LogActivity XferLogRec
    
    Set ProfileInfo = Nothing
    Exit Sub
    
errHandler:
    tfnErrHandler ProcName, False
    Err.Clear
End Sub
Public Function GetTransferProfileInfo(ByVal TransferID As String, ByVal PartnerID As String) As Collection
    Const ProcName = "GetTransferProfileInfo"
    Dim field  As DAO.field
    Dim Fields As Collection
    Dim rs     As DAO.Recordset
    Dim SQL    As String
    Dim Value  As Variant
    
    '------------------------------------------------------------------------------------
    'Initialize
    '------------------------------------------------------------------------------------
    On Error GoTo errHandler
    Set Fields = New Collection
    
    '------------------------------------------------------------------------------------
    'Get Partner Transfer Profile Record
    '------------------------------------------------------------------------------------
    SQL = SQLParm(SQL_SEL_PARTNER_PROFILE, "@TransferID", TransferID, _
                                           "@PartnerID", PartnerID)
    If fnRecordset(rs, SQL) >= 0 Then
        If Not rs.EOF Then
            For Each field In rs.Fields
                With field
                    Value = Nz(.Value)
                    If VarType(Value) = vbString Then
                        Value = Trim$(Value)
                    End If
                    Fields.Add Value, .Name 'Add field's value with its name as the Key
                End With
            Next 'Field
            Set field = Nothing
        End If
        rs.Close
    End If
    Set rs = Nothing
    
    '------------------------------------------------------------------------------------
    'Get Transfer Profile Record
    '------------------------------------------------------------------------------------
    SQL = SQLParm(SQL_SEL_TRANSFER_PROFILE, "@TransferID", TransferID)
    If fnRecordset(rs, SQL) >= 0 Then
        If Not rs.EOF Then
            For Each field In rs.Fields
                With field
                    Value = Nz(.Value)
                    If VarType(Value) = vbString Then
                        Value = Trim$(Value)
                    End If
                    Fields.Add Value, .Name 'Add field's value with its name as the Key
                End With
            Next 'Field
            Set field = Nothing
        End If
        rs.Close
    End If
    Set rs = Nothing
    
    '------------------------------------------------------------------------------------
    'Return the Transfer Profile Information
    '------------------------------------------------------------------------------------
    Set GetTransferProfileInfo = Fields
    Exit Function
    
errHandler:
    tfnErrHandler ProcName, False
    Set GetTransferProfileInfo = New Collection
    Err.Clear
End Function
Public Sub FTP_ExecuteOutsideProcess(ByVal ProgName As String)
    Const ProcName = "FTP_ExecuteOutsideProcess"
    'Executes a Windows or 4GL program and waits for it to terminate before returning
    Dim ProcessID     As Long
    Dim ProcessHandle As Long
    
    On Error GoTo errHandler
    
    If InStr(UCase$(ProgName), ".4GE") Then
        nWhatToUse = USE_RCMD
        fnExecute4GE ProgName, , False
    Else
        ProcessID = Shell(ProgName)
        If ProcessID Then
            ProcessHandle = OpenProcess(SYNCHRONIZE, 0&, ProcessID)
            WaitForSingleObject ProcessHandle, INFINITE
            CloseHandle ProcessHandle
        End If
    End If
    Exit Sub
    
errHandler:
    'This is a background process ... cannot display message box
    tfnErrHandler ProcName, False
    Err.Clear
End Sub
Public Sub FTP_EMailNotify(ByVal SendTo As String, ByVal Subj As String, ByVal Msg As String)
    Const ProcName = "FTP_EMailNotify"
    Dim objMail As clsSendMail
    
    On Error GoTo errHandler
    Set objMail = New clsSendMail
    With objMail
        .UserName = "abc" 'does not matter what this is set to
        .Password = "abc" 'does not matter what this is set to
        .SendTo = SendTo
        .Subject = Subj
        .message = Msg
        .LaunchSENDMAIL
    End With
    Exit Sub
    
errHandler:
    tfnErrHandler ProcName, False
    Err.Clear
End Sub
