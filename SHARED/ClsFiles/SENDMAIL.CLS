VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSendMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Implementation:
'1. Include this Class SENDMAIL.CLS (clsSendMail) in
'   shared\clsfiles directory into your project.
'2. Include TASK.BAS (modTASK) in shared directory into your project.

'used with fixed length strings - normally with windows api calls
Const MAX_STRING_LENGTH As Integer = 255
Const szEMPTY As String = ""
'david 11/14/2001 commented out, use the template definition
'Const szFACTOR_INI As String = "FACTOR.INI" 'application INI filename

Const sSECTION As String = "SendMail_Info"
Const sSECTION_OLD As String = "Old_SendMail"
Const sKEY_emailaddress As String = "emailaddress"
Const sKEY_subject As String = "subject"
Const sKEY_message As String = "message"
Const sKEY_attachment As String = "attachment"
Const sKEY_delete_after_sent As String = "delete_after_sent"
Const sKEY_rename_after_sent As String = "rename_after_sent"
Const sKEY_rename_to As String = "rename_to"
Const sKEY_logid As String = "logid"
Const sKEY_lasterror As String = "lasterror"

Private m_lstLog As ListBox
Private m_ProgramPath As String

Private m_UserName As String
Private m_Password As String
Private m_SendTo As String
Private m_Subject As String
Private m_Message As String
Private m_Attachment As String
Private m_DeleteAftSent As String
Private m_RenameAftSent As String
Private m_RenameTo As String
Private m_LogID As String
Private m_ErrorMessage As String
'

'interface properties
Property Set LogListBox(lstList As ListBox)
    Set m_lstLog = lstList
End Property

Property Let UserName(sParm As String)
    m_UserName = sParm
End Property
Property Let Password(sParm As String)
    m_Password = sParm
End Property
Property Let SendTo(sParm As String)
    m_SendTo = sParm
End Property
Property Let Subject(sParm As String)
    m_Subject = sParm
End Property
Property Let message(sParm As String)
    m_Message = sParm
End Property
Property Let Attachment(sParm As String)
    m_Attachment = sParm
End Property
Property Let DeleteAfterSent(sParm As String)
    m_DeleteAftSent = sParm
End Property
Property Let RenameAfterSent(sParm As String)
    m_RenameAftSent = sParm
End Property
Property Let RenameTo(sParm As String)
    m_RenameTo = sParm
End Property

Property Get ErrorMessage() As String
    ErrorMessage = m_ErrorMessage
End Property

'return an error message if failed
Public Function LaunchSENDMAIL(Optional nWait As Integer = 20) As String
    Const SUB_NAME As String = "LaunchSENDMAIL"
    
    Const sSENDMAIL_EXE As String = "SENDMAIL.EXE"
    Const sLocalBin As String = "C:\FACTOR\EXECTRAK\BIN"
    Const sRemoteBin As String = "J:\FACTOR\EXECTRAK\BIN"
    
    
    Dim dfTimer As Single, sExeFile As String, sMsg As String
    Dim dfBeginRun As Single, sParm As String
    Dim sFile As String, bDone As Boolean
    Dim bMoursePointerIsSet As Boolean
    
    bMoursePointerIsSet = Screen.MousePointer <> vbDefault
    
    m_ProgramPath = ""
    m_ErrorMessage = ""
    
    'check current directory
    m_ProgramPath = App.path
    sExeFile = fnAddBkSlash(m_ProgramPath) + sSENDMAIL_EXE
    On Error Resume Next
    sFile = ""
    sFile = Dir(sExeFile)
    On Error GoTo 0
    
    If sFile = "" Then
        'search local bin
        m_ProgramPath = sLocalBin
        sExeFile = fnAddBkSlash(m_ProgramPath) + sSENDMAIL_EXE
        On Error Resume Next
        sFile = ""
        sFile = Dir(sExeFile)
        On Error GoTo 0
        
        If sFile = "" Then
            'search remote bin
            m_ProgramPath = sRemoteBin
            sExeFile = fnAddBkSlash(m_ProgramPath) + sSENDMAIL_EXE
            On Error Resume Next
            sFile = ""
            sFile = Dir(sExeFile)
            On Error GoTo 0
        End If
    End If
    
    If sFile = "" Then
        m_ErrorMessage = "SENDMAIL program not found."
        tfnErrHandler SUB_NAME, -1, m_ErrorMessage, False
        LaunchSENDMAIL = m_ErrorMessage
        Exit Function
    End If
    
    sParm = fnBuildParm(sMsg)
    
    If sMsg <> "" Then
        GoTo errSendMailError
    End If
    
    'Launch SENDMAIL program
    dfBeginRun = Timer
    If Not fnRunExe(sExeFile + " " + sParm, , , False) Then
        m_ErrorMessage = "Failed to run SENDMAIL program."
        tfnErrHandler SUB_NAME, -1, m_ErrorMessage, False
        LaunchSENDMAIL = m_ErrorMessage
        Exit Function
    End If
        
    sMsg = "Start Launching Program: " + sExeFile + "  " + sParm
    subAddLog sMsg
    
    Screen.MousePointer = vbHourglass
    
    'busy waiting for EDI Translation Processor to finish
    dfTimer = Timer
    Do
        Screen.MousePointer = vbHourglass
        DoEvents
        If Timer - dfTimer > 2100 Then
            m_ErrorMessage = "SENDMAIL program time out."
            tfnErrHandler SUB_NAME, -1, m_ErrorMessage, False
            LaunchSENDMAIL = m_ErrorMessage
            Exit Function
        End If
        
        Screen.MousePointer = vbHourglass
        
        Wait nWait
    Loop While fnExeIsRunning(sExeFile)
    
    If Not bMoursePointerIsSet Then
        Screen.MousePointer = vbDefault
    End If
    
    sMsg = "Program Finished. Running Time: " & Timer - dfBeginRun & " seconds."
    subAddLog sMsg
    
    Wait 1
    
    sMsg = fnCheckSendMailError()
    If sMsg <> "" Then
        GoTo errSendMailError
    End If
    
    Wait 1
    
    LaunchSENDMAIL = ""
    
    Exit Function

errSendMailError:
    
    If sMsg = "" Then
        If Err.Number > 0 Then
            sMsg = "Error(s) has occurred while sending an E-mail." _
                + " Err Code: " & Err.Number & " Err Desc: " & Err.Description + "."
        End If
    End If
    
    m_ErrorMessage = sMsg

    LaunchSENDMAIL = m_ErrorMessage
End Function

Private Function fnCheckSendMailError() As String
    
    Dim sLogID As String
    
    fnCheckSendMailError = ""
    
    'retrieve logid
    sLogID = Trim(tfnReadINI(sSECTION_OLD, sKEY_logid, tfnGetWindowsDir() + szFACTOR_INI))
    
    If sLogID = m_LogID Then
        fnCheckSendMailError = Trim(tfnReadINI(sSECTION_OLD, sKEY_lasterror, tfnGetWindowsDir() + szFACTOR_INI))
    End If
End Function

Private Function fnAddBkSlash(ByVal sIn As String) As String
    sIn = Trim(sIn)
    If Right(sIn, 1) <> "\" Then fnAddBkSlash = sIn + "\" Else fnAddBkSlash = sIn
End Function

Private Sub Form_Initialize()
    Set m_lstLog = Nothing
End Sub

Private Sub Form_Terminate()
    On Error Resume Next
    Set m_lstLog = Nothing
End Sub

Private Function fnBuildParm(sErrMsg As String) As String
    Const sUseProfile As String = "USE_PROFILE"

    Dim sParm As String
    Dim sTemp As String
    
    '"SHOW|HIDE","User Name","Password","USE_PROFILE"[,LogID=xxxxxxxxxxxx][,iniFile=x:\xxx]
    
    sTemp = ""
    
    If m_UserName = "" Then
        If sTemp <> "" Then
            sTemp = sTemp + ", "
        End If
        sTemp = sTemp + "'User Name'"
    End If
    If m_Password = "" Then
        If sTemp <> "" Then
            sTemp = sTemp + ", "
        End If
        sTemp = sTemp + "'Password'"
    End If
    If m_SendTo = "" Then
        If sTemp <> "" Then
            sTemp = sTemp + ", "
        End If
        sTemp = sTemp + "'Send To'"
    End If
    If m_Subject = "" Then
        If sTemp <> "" Then
            sTemp = sTemp + ", "
        End If
        sTemp = sTemp + "'Subject'"
    End If
    If m_Message = "" Then
        If sTemp <> "" Then
            sTemp = sTemp + ", "
        End If
        sTemp = sTemp + "'Message'"
    End If
    
    If sTemp <> "" Then
        sErrMsg = sTemp + " cannot be empty."
        Exit Function
    End If
    
    sParm = """HIDE""," + """" + m_UserName + """," + """" + m_Password + ""","
    sParm = sParm + """" + sUseProfile + """"
    
    If (Len(io.HostedFolder) > 0) Then
        sTemp = io.NewIniPath("C:\FACTOR\FACTOR.INI")
        sParm = sParm + "," + """" + "iniFile=" + sTemp + """"
    End If
    
    m_LogID = fnCreateLogID()
    
    subWriteSendMailInfo
    fnBuildParm = sParm
End Function

Private Sub subAddLog(sMsg As String, Optional vPipeAsCrLf)
    Dim lLine As Long, sTemp As String, nPosi As Integer
    
    If m_lstLog Is Nothing Then
        Exit Sub
    End If
    
    On Error Resume Next
    
    sTemp = sMsg
    nPosi = InStr(sTemp, vbCr)
     
     While nPosi > 0
        m_lstLog.AddItem Mid(sTemp, 1, nPosi - 1)
        m_lstLog.ListIndex = m_lstLog.ListCount - 1
        If Mid(sTemp, nPosi + 1) = vbLf Then
            sTemp = Mid(sTemp, nPosi + 2)
        Else
            sTemp = Mid(sTemp, nPosi + 1)
        End If
        nPosi = InStr(sTemp, vbCr)
    Wend
    
    If sTemp <> "" Then
        m_lstLog.AddItem sTemp
        m_lstLog.ListIndex = m_lstLog.ListCount - 1
    End If
    
    m_lstLog.Refresh
End Sub

Private Sub Wait(dfSeconds As Integer)
    Dim dfTimer As Single
    
    dfTimer = Timer
    While Timer - dfTimer <= dfSeconds
        DoEvents
    Wend
End Sub

'Log ID
Private Function fnCreateLogID() As String
    Dim sID_Prefix1 As String
    Dim sID_Prefix2 As String
    Dim sDateTime As String
    
    Randomize Timer
    sID_Prefix1 = Format(Int((98 * Rnd) + 1), "00")
    Randomize Timer
    sID_Prefix2 = Format(Int((98 * Rnd) + 1), "00")
    
    sDateTime = Format(Date, "mmdd") + Format(Time, "mmss")
    
    fnCreateLogID = sID_Prefix1 + sID_Prefix2 + sDateTime
End Function

Private Sub subWriteSendMailInfo()
    'store the key values in [SendMail_Info]
    tfnWriteINI sSECTION, sKEY_emailaddress, m_SendTo, tfnGetWindowsDir() + szFACTOR_INI
    tfnWriteINI sSECTION, sKEY_subject, m_Subject, tfnGetWindowsDir() + szFACTOR_INI
    tfnWriteINI sSECTION, sKEY_message, m_Message, tfnGetWindowsDir() + szFACTOR_INI
    tfnWriteINI sSECTION, sKEY_attachment, m_Attachment, tfnGetWindowsDir() + szFACTOR_INI
    If m_DeleteAftSent = "" Then
        m_DeleteAftSent = "false"
    End If
    tfnWriteINI sSECTION, sKEY_delete_after_sent, m_DeleteAftSent, tfnGetWindowsDir() + szFACTOR_INI
    If m_RenameAftSent = "" Then
        m_RenameAftSent = "false"
    End If
    tfnWriteINI sSECTION, sKEY_rename_after_sent, m_RenameAftSent, tfnGetWindowsDir() + szFACTOR_INI
    tfnWriteINI sSECTION, sKEY_rename_to, m_RenameTo, tfnGetWindowsDir() + szFACTOR_INI
    tfnWriteINI sSECTION, sKEY_logid, m_LogID, tfnGetWindowsDir() + szFACTOR_INI
End Sub

Private Function tfnReadINI(szSection As String, szKey As String, szINIFile As String) As String
    Dim nLength As Long     'length of the value returned for api call
    Dim szINI As String     'string to hold the value retrieved
    Dim iniFile$
    
    iniFile = io.NewIniPath(szINIFile)
    
    szINI = Space(MAX_STRING_LENGTH) 'clear and make the string fixed length
    
    'get the [value] for the [section], [key], and ini file sent
    nLength = GetPrivateProfileString(szSection, szKey, szEMPTY, szINI, MAX_STRING_LENGTH, iniFile)
    
    If nLength <> 0 Then 'if length positive [value] has been found
        szINI = Left(szINI, nLength) 'make it a basic string
    Else
        szINI = ""
    End If
    
    tfnReadINI = szINI 'return the value

End Function

Private Function tfnWriteINI(szSection As String, szKey As String, szValue As String, szINIFile As String) As Boolean
    Dim bStatus As Boolean 'status returned from api call
    Dim iniFile$
    
    iniFile = io.NewIniPath(szINIFile)
    
    'write the [value] for the [section], [key], and ini file sent
    bStatus = WritePrivateProfileString(szSection, szKey, szValue, iniFile)
    
    tfnWriteINI = bStatus

End Function

Private Sub Class_Initialize()
    Set m_lstLog = Nothing
    
    m_ProgramPath = ""
    m_UserName = ""
    m_Password = ""
    m_SendTo = ""
    m_Subject = ""
    m_Message = ""
    m_Attachment = ""
    m_DeleteAftSent = ""
    m_RenameAftSent = ""
    m_RenameTo = ""
    m_LogID = ""
    m_ErrorMessage = ""
End Sub

Private Sub Class_Terminate()
    Set m_lstLog = Nothing
    
    m_ProgramPath = ""
    m_UserName = ""
    m_Password = ""
    m_SendTo = ""
    m_Subject = ""
    m_Message = ""
    m_Attachment = ""
    m_DeleteAftSent = ""
    m_RenameAftSent = ""
    m_RenameTo = ""
    m_LogID = ""
    m_ErrorMessage = ""
End Sub

Public Sub ResetVariables()
    m_UserName = ""
    m_Password = ""
    m_SendTo = ""
    m_Subject = ""
    m_Message = ""
    m_Attachment = ""
    m_DeleteAftSent = ""
    m_RenameAftSent = ""
    m_RenameTo = ""
    m_LogID = ""
    m_ErrorMessage = ""
End Sub
