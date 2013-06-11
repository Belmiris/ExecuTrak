VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form objEDIPROC 
   Caption         =   "objEDIPROC"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2130
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1380
   ScaleWidth      =   2130
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdlgOpen 
      Left            =   276
      Top             =   192
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Select EDI Translator Program"
   End
End
Attribute VB_Name = "objEDIPROC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Implementation:
'0. Include the component Microsoft Common Dialog Control 6.0 (SP3) in your project.
'1. Include this Form EDIPROC.FRM (objEDIPROC) into your project.
'2. Include TASK.BAS (modTASK) in shared directory into your project.

Private m_AppPath As String
Private m_frmParent As Form
Private m_lstLog As ListBox

Private m_ProgramPath As String

'GENERAL
Private m_MessageType As String
Private m_WorkPath As String
Private m_lFilePath As String
Private m_WriteLogFlag As String
Private m_LFileName As String
Private m_LFileType As String

'INCOMING and OUTGOING
Private m_IFilePath As String
Private m_IFileType As String
Private m_IFileName As String
Private m_OFilePath As String
Private m_OFileType As String
Private m_OFileName As String
Private m_WriteMode As String
Private m_BackupFileFlag As String
Private m_RemoveOriginal As String
Private m_BackupPath As String
Private m_BackupFileType As String
Private m_BackupFileName As String
Private m_RunMode As String
Private m_LogID As String
Private m_ErrorMessage As String

'Vijaya on 05/20/03 Magic#401028-7 hide error popup form
'by default it will show the popup.
Private m_SummarySection As String
Private m_bShowError As Boolean
'

'interface properties
Property Set ParentForm(frmForm As Form)
    Set m_frmParent = frmForm
    m_AppPath = fnAddBkSlash(UCase(App.path))
End Property

Property Set LogListBox(lstList As ListBox)
    Set m_lstLog = lstList
End Property

Property Let MessageType(sParm As String)
    m_MessageType = sParm
End Property
Property Let WorkPath(sParm As String)
    m_WorkPath = sParm
End Property
Property Let LFilePath(sParm As String)
    m_lFilePath = sParm
End Property
Property Let WriteLogFlag(sParm As String)
    m_WriteLogFlag = sParm
End Property
Property Let LFileName(sParm As String)
    m_lFilePath = sParm
End Property
Property Let LFileType(sParm As String)
    m_LFileType = sParm
End Property
Property Let IFilePath(sParm As String)
    m_IFilePath = sParm
End Property
Property Let IFileType(sParm As String)
    m_IFileType = sParm
End Property
Property Let IFileName(sParm As String)
    m_IFileName = sParm
End Property
Property Let OFilePath(sParm As String)
    m_OFilePath = sParm
End Property
Property Let OFileType(sParm As String)
    m_OFileType = sParm
End Property
Property Let OFileName(sParm As String)
    m_OFileName = sParm
End Property
Property Let WriteMode(sParm As String)
    m_WriteMode = sParm
End Property
Property Let BackupFileFlag(sParm As String)
    m_BackupFileFlag = sParm
End Property
Property Let RemoveOriginal(sParm As String)
    m_RemoveOriginal = sParm
End Property
Property Let BackupPath(sParm As String)
    m_BackupPath = sParm
End Property
Property Let BackupFileType(sParm As String)
    m_BackupFileType = sParm
End Property
Property Let BackupFileName(sParm As String)
    m_BackupFileName = sParm
End Property
Property Let RunMode(sParm As String)
    m_RunMode = sParm
End Property

Property Get ErrorMessage() As String
    ErrorMessage = m_ErrorMessage
End Property

'Vijaya on 05/20/03 Magic#401028-7
Property Let SummarySection(sParm As String)
    m_SummarySection = sParm
End Property

Property Let ShowErrorPopup(bParm As Boolean)
    m_bShowError = bParm
End Property
'''''''''''''''''''''''''''''''''

'return an error message if failed
Public Function LaunchEDIPROC(Optional nWait As Integer = 20) As String
    Const SUB_NAME As String = "LaunchEDIPROC"
    Const sEDIPROC_EXE As String = "EDIPROC.EXE"
    Const sEDIPROC_INI As String = "EDIPROC.INI"
    Const sEDIPROC_MDB As String = "EDIXLATE.MDB"
    Const sSECTION_EDIXLATOR As String = "EDI XLATOR"
    Const sKEY_PROGRAM As String = "Program"
    Const sLocalBin As String = "C:\FACTOR\EXECTRAK\BIN"
    Const sRemoteBin As String = "J:\FACTOR\EXECTRAK\BIN"
    
    
    Dim dfTimer As Single, sExeFile As String, sMsg As String
    Dim dfBeginRun As Single, sParm As String
    Dim sFile As String, bDone As Boolean

    'david 06/19/2008  #596302
    CheckXlatorProgramPath
    ''''''''''''''''''''''''''

    m_ProgramPath = ""
    m_ErrorMessage = ""
    
    'get the EDIPROC program path from factor.ini
    m_ProgramPath = Trim(tfnReadINI(sSECTION_EDIXLATOR, sKEY_PROGRAM, tfnGetWindowsDir + szFACTOR_INI))
    
    If m_ProgramPath <> "" Then
        sExeFile = fnAddBkSlash(m_ProgramPath) + sEDIPROC_EXE
        On Error Resume Next
        sFile = ""
        sFile = Dir(sExeFile)
        On Error GoTo 0
    End If
    
    If sFile = "" Then
        'check current directory
        m_ProgramPath = App.path
        sExeFile = fnAddBkSlash(m_ProgramPath) + sEDIPROC_EXE
        On Error Resume Next
        sFile = ""
        sFile = Dir(sExeFile)
        On Error GoTo 0
        
        If sFile = "" Then
            'search local bin
            m_ProgramPath = sLocalBin
            sExeFile = fnAddBkSlash(m_ProgramPath) + sEDIPROC_EXE
            On Error Resume Next
            sFile = ""
            sFile = Dir(sExeFile)
            On Error GoTo 0
            
            If sFile = "" Then
                'search remote bin
                m_ProgramPath = sRemoteBin
                sExeFile = fnAddBkSlash(m_ProgramPath) + sEDIPROC_EXE
                On Error Resume Next
                sFile = ""
                sFile = Dir(sExeFile)
                On Error GoTo 0
            End If
        End If
    End If
    
    If sFile = "" Then
        cdlgOpen.Filter = "EDI Translator Program (" + sEDIPROC_EXE + ")|" + sEDIPROC_EXE
        On Error Resume Next
        cdlgOpen.InitDir = m_ProgramPath
        cdlgOpen.Filename = ""
        On Error GoTo 0
        
        Do
            If fnOpenFile() Then
                sExeFile = UCase(cdlgOpen.Filename)
                On Error Resume Next
                sFile = ""
                sFile = Dir(sExeFile)
                On Error GoTo 0
                If Dir(sFile) <> "" Then
                    If UCase(Right(sFile, Len(sEDIPROC_EXE))) = sEDIPROC_EXE Then
                        bDone = True
                    End If
                End If
            Else
                m_ErrorMessage = "EDI Translation Processor program not found."
                tfnErrHandler SUB_NAME, -1, m_ErrorMessage, False
                LaunchEDIPROC = m_ErrorMessage
                Exit Function
            End If
        Loop Until bDone
        
        m_ProgramPath = fnGetPath(sExeFile)
        
        'update the EDIPROC program path to factor.ini
        tfnWriteINI sSECTION_EDIXLATOR, sKEY_PROGRAM, m_ProgramPath, szFACTOR_INI
    End If
    
    'validate EDIPROC files
    'Database
    On Error Resume Next
    sFile = ""
    'sFile = Dir(fnAddBkSlash(m_ProgramPath) + sEDIPROC_MDB)
    sFile = io.EDIXLATEFile
    On Error GoTo 0
    
    If sFile = "" Then
        m_ErrorMessage = "EDI Translation Processor database not found [" + sFile + "]."
        tfnErrHandler SUB_NAME, -1, m_ErrorMessage, False
        LaunchEDIPROC = m_ErrorMessage
        Exit Function
    End If
    
    'INI
    On Error Resume Next
    sFile = ""
    sFile = io.NewIniPath(Dir(fnAddBkSlash(m_ProgramPath) + sEDIPROC_INI))
    On Error GoTo 0
    
    If sFile = "" Then
        m_ErrorMessage = "EDI Translation Processor INI file not found [" + sFile + "]."
        tfnErrHandler SUB_NAME, -1, m_ErrorMessage, False
        LaunchEDIPROC = m_ErrorMessage
        Exit Function
    End If
    
    If sMsg <> "" Then
        m_ErrorMessage = sMsg
        tfnErrHandler SUB_NAME, -1, m_ErrorMessage, False
        LaunchEDIPROC = m_ErrorMessage
        Exit Function
    End If
    
    sParm = fnBuildParm()
    
    'Launch EDI Translation Processor
    dfBeginRun = Timer
    If Not fnRunExe(sExeFile + " " + sParm, , , False) Then
        m_ErrorMessage = "Failed to run EDI Translation Processor."
        tfnErrHandler SUB_NAME, -1, m_ErrorMessage, False
        LaunchEDIPROC = m_ErrorMessage
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
        'change the program, so it takes more than 2 hours, considers it time out
        If Timer - dfTimer > 12000 Then
            m_ErrorMessage = "EDI Translation Processor time out."
            tfnErrHandler SUB_NAME, -1, m_ErrorMessage, False
            LaunchEDIPROC = m_ErrorMessage
            Exit Function
        End If
        
        Screen.MousePointer = vbHourglass
        
        Wait nWait
    Loop While fnExeIsRunning(sExeFile)
    
    Screen.MousePointer = vbDefault
    
    sMsg = "Program Finished. Running Time: " & Timer - dfBeginRun & " seconds."
    subAddLog sMsg
    
    Wait 1
    
    sMsg = fnCheckEdiProcError()
    If sMsg <> "" Then
        GoTo errEdiProcError
    End If
    
    Wait 1
    
    LaunchEDIPROC = ""
    
    Exit Function

errEdiProcError:
    
    m_ErrorMessage = "Error(s) has occurred during EDI Translation." _
        + vbCrLf + vbCrLf + sMsg

    LaunchEDIPROC = m_ErrorMessage
End Function

Private Function fnOpenFile() As Boolean
    On Error GoTo errHandler
    
    cdlgOpen.Action = 1
    
    fnOpenFile = True
    
    Exit Function

errHandler:
    'do nothing
End Function

Private Function fnCheckEdiProcError() As String
    Const SUB_NAME As String = "fnCheckEdiProcError"
    
    Dim dbEDI As Database
    Dim sSql As String
    Dim rsTemp As Recordset
    Dim sFileName As String
    Dim sErrMsg As String
    Dim sTemp As String
    Dim bShowForm As Boolean
    
    fnCheckEdiProcError = ""
    
    sFileName = fnAddBkSlash(m_ProgramPath) + "EDIXLATE.MDB"
    
    Set dbEDI = tfnOpenEdiXlateDatabase(sFileName)
    
    If dbEDI Is Nothing Then
        fnCheckEdiProcError = "Unable to open EDI Translation Database"
        Exit Function
    End If
    
    sSql = "SELECT el_desc, el_number FROM err_log"
    sSql = sSql + " WHERE el_log_id = " + tfnSQLString(m_LogID)
    sSql = sSql + " ORDER BY el_id, el_seq"
    
    If fnGetEdiXlateRecord(rsTemp, sSql, dbEDI, SUB_NAME, , False) < 0 Then
        fnCheckEdiProcError = "Failed to access EDI Translation Database"
        Exit Function
    End If
    
    bShowForm = True
    
    While Not rsTemp.EOF
        If Not IsNull(rsTemp!el_desc) Then
            sErrMsg = sErrMsg + Trim(rsTemp!el_desc)
            sTemp = sTemp & Trim(rsTemp!el_desc) & vbCrLf
            
            If tfnRound(rsTemp!el_number <> -2410) Then
                bShowForm = False
            End If
                
        End If
        rsTemp.MoveNext
    Wend
    
    'Vijaya on 05/20/03 Magic#401028-7
    If Not m_bShowError Then
        bShowForm = False
    End If
    ''''''''''''''''''''''''''''''''''
    
    'Add by junsong 04/05/02 to skip some error
    If bShowForm And sTemp <> "" Then
        With EDIErrLog
            .txtMsg = sTemp
            .Show vbModal
            
            If .Action = vbIgnore Then
                sErrMsg = ""
            End If
            
        End With
        
    End If

    rsTemp.Close
    Set rsTemp = Nothing
    dbEDI.Close
    Set dbEDI = Nothing
    
    fnCheckEdiProcError = sErrMsg
End Function

Private Function tfnOpenEdiXlateDatabase(sDatabaseNamePath As String) As Database
    
    On Error GoTo ERROR_CONNECTING 'set the runtime error handler for database connection

    If t_engFactor Is Nothing Then
        Set t_engFactor = New DBEngine 'create a new dDBEngine
    End If
    
    If t_wsWorkSpace Is Nothing Then
        Set t_wsWorkSpace = t_engFactor.Workspaces(0) 'set the default workspace handle
    End If
    
    Set tfnOpenEdiXlateDatabase = t_wsWorkSpace.OpenDatabase(sDatabaseNamePath)
    
    Exit Function
    
ERROR_CONNECTING:
    
    tfnErrHandler "tfnOpenEdiXlateDatabase"
    Set tfnOpenEdiXlateDatabase = Nothing
End Function

Private Function fnGetEdiXlateRecord(rsTemp As Recordset, strSQL As String, _
                            dbLocal As Database, sCalledFrom As String, _
                            Optional bShowError, _
                            Optional vMoveLast) As Long

    Dim bMoveLast As Boolean
    
    If IsMissing(vMoveLast) Then
        bMoveLast = True
    Else
        bMoveLast = vMoveLast
    End If

    On Error GoTo SQLError
    Set rsTemp = dbLocal.OpenRecordset(strSQL, dbOpenSnapshot)
    
    If bMoveLast Then
        If rsTemp.RecordCount > 0 Then
            rsTemp.MoveLast
            rsTemp.MoveFirst
        End If
    End If
    
    fnGetEdiXlateRecord = rsTemp.RecordCount
    Exit Function

SQLError:
    tfnErrHandler "fnGetEdiXlateRecord," & sCalledFrom, strSQL, bShowError
    fnGetEdiXlateRecord = -9999
End Function

Private Function fnAddBkSlash(ByVal sIn As String) As String
    sIn = Trim(sIn)
    If Right(sIn, 1) <> "\" Then fnAddBkSlash = sIn + "\" Else fnAddBkSlash = sIn
End Function

Private Sub Form_Initialize()
    Set m_frmParent = Nothing
    Set m_lstLog = Nothing
    
    'Vijaya on 05/20/03 Magic#401028-7
    'Default Value
    m_bShowError = True
End Sub

Private Sub Form_Terminate()
    On Error Resume Next
    Set m_lstLog = Nothing
    Set m_frmParent = Nothing
End Sub

Private Function fnBuildParm() As String
    Dim sParm As String
    
    sParm = "DSN=" + dsnName + ","
    
    If m_MessageType <> "" Then
        sParm = sParm + "MessageType=" + m_MessageType + ","
    End If
    If m_WorkPath <> "" Then
        sParm = sParm + "WorkPath=" + m_WorkPath + ","
    End If
    If m_lFilePath <> "" Then
        sParm = sParm + "LFilePath=" + m_lFilePath + ","
    End If
    If m_WriteLogFlag <> "" Then
        sParm = sParm + "WriteLogFlag=" + m_WriteLogFlag + ","
    End If
    If m_LFileName <> "" Then
        sParm = sParm + "LFileName=" + m_LFileName + ","
    End If
    If m_LFileType <> "" Then
        sParm = sParm + "LFileType=" + m_LFileType + ","
    End If
    If m_IFilePath <> "" Then
        sParm = sParm + "IFilePath=" + m_IFilePath + ","
    End If
    If m_IFileType <> "" Then
        sParm = sParm + "IFileType=" + m_IFileType + ","
    End If
    If m_IFileName <> "" Then
        sParm = sParm + "IFileName=" + m_IFileName + ","
    End If
    If m_OFilePath <> "" Then
        fnCheckAndCreateDirectory m_OFilePath, False
        sParm = sParm + "OFilePath=" + m_OFilePath + ","
    End If
    If m_OFileType <> "" Then
        sParm = sParm + "OFileType=" + m_OFileType + ","
    End If
    If m_OFileName <> "" Then
        sParm = sParm + "OFileName=" + m_OFileName + ","
    End If
    If m_WriteMode <> "" Then
        sParm = sParm + "WriteMode=" + m_WriteMode + ","
    End If
    If m_BackupFileFlag <> "" Then
        sParm = sParm + "BackupFileFlag=" + m_BackupFileFlag + ","
    End If
    If m_RemoveOriginal <> "" Then
        sParm = sParm + "RemoveOriginal=" + m_RemoveOriginal + ","
    End If
    If m_BackupPath <> "" Then
        sParm = sParm + "BackupPath=" + m_BackupPath + ","
    End If
    If m_BackupFileType <> "" Then
        sParm = sParm + "BackupFileType=" + m_BackupFileType + ","
    End If
    If m_BackupFileName <> "" Then
        sParm = sParm + "BackupFileName=" + m_BackupFileName + ","
    End If
    If m_RunMode <> "" Then
        sParm = sParm + "RunMode=" + m_RunMode + ","
    End If
    
    'Vijaya on 05/22/03 Magic#401028-7
    If m_SummarySection <> "" Then
        sParm = sParm + "SummarySection=" + m_SummarySection + ","
    End If
    'end of Vijaya
    
    If sParm <> "" Then
        sParm = Left(sParm, Len(sParm) - 1)
    End If

    m_LogID = fnCreateLogID()
    
    If sParm <> "" Then
        sParm = sParm + ", LogID=" + m_LogID
    Else
        sParm = "LogID=" + m_LogID
    End If
    
    fnBuildParm = sParm
End Function

Private Sub subAddLog(sMsg As String, Optional vPipeAsCrLf)
    Dim lLine As Long, sTemp As String, nPosi As Integer
    Dim sDelimeter As String
    Dim x As Long
    
    If m_lstLog Is Nothing Then
        Exit Sub
    End If
    
    On Error Resume Next
    
    If IsMissing(vPipeAsCrLf) Or sMsg = "" Then
        sDelimeter = vbCr
    Else
        sDelimeter = "|"
    End If
    
    sTemp = sMsg
    nPosi = InStr(sTemp, sDelimeter)
    
     While nPosi > 0
        sTemp = Mid(sTemp, 1, nPosi - 1)
        m_lstLog.AddItem sTemp
        
        x = m_frmParent.TextWidth(sTemp & "  ")
        If x > Val(m_lstLog.Tag) Then
            m_lstLog.Tag = x
            If m_frmParent.ScaleMode = vbTwips Then x = x / Screen.TwipsPerPixelX
            SendMessageByNum m_lstLog.hwnd, LB_SETHORIZONTALEXTENT, x, 0
        End If
    
        m_lstLog.ListIndex = m_lstLog.ListCount - 1
        If Mid(sTemp, nPosi + 1) = vbLf Then
            sTemp = Mid(sTemp, nPosi + 2)
        Else
            sTemp = Mid(sTemp, nPosi + 1)
        End If
        nPosi = InStr(sTemp, sDelimeter)
    Wend
    
    If sTemp <> "" Then
        m_lstLog.AddItem sTemp
        
        x = m_frmParent.TextWidth(sTemp & "  ")
        If x > Val(m_lstLog.Tag) Then
            m_lstLog.Tag = x
            If m_frmParent.ScaleMode = vbTwips Then x = x / Screen.TwipsPerPixelX
            SendMessageByNum m_lstLog.hwnd, LB_SETHORIZONTALEXTENT, x, 0
        End If
    
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

Private Function fnGetPath(sPathFile As String) As String
    Dim i As Integer
    
    For i = Len(sPathFile) To 1 Step -1
        If Mid(sPathFile, i, 1) = "\" Then
            Exit For
        End If
    Next i
    
    If i > 0 Then
        fnGetPath = Left(sPathFile, i - 1)
    End If
End Function

'david 05/06/00
'Log ID: max 15 characters
Private Function fnCreateLogID() As String
    Dim sID_Prefix1_4 As String
    Dim sID_Prefix2_3 As String
    Dim sDateTime As String
    
    Randomize Timer
    sID_Prefix1_4 = Format(Int((9999 * Rnd) + 1), "000")
    Randomize Timer
    sID_Prefix2_3 = Format(Int((999 * Rnd) + 1), "000")
    
    sDateTime = Format(Date, "mmdd") + Format(Time, "mmss")
    
    fnCreateLogID = sID_Prefix1_4 + sID_Prefix2_3 + sDateTime
End Function

Public Function fnCheckAndCreateDirectory(ByVal sDir As String, bPromptForCreate As Boolean) As String
    Dim sTmpDir As String, sDrive, nPosi As String, sParentDir As String, sNewDir As String
    
    fnCheckAndCreateDirectory = ""
    
    On Error Resume Next
    sTmpDir = Dir(sDir, vbDirectory)
    
    If sTmpDir <> "" Then
        Exit Function
    End If
    
    If bPromptForCreate Then
        If MsgBox("Directory '" + sDir + "' does not exist.  Do you want to create it?", vbQuestion + vbYesNo) = vbNo Then
            fnCheckAndCreateDirectory = "Directory '" + sDir + "' does not exist"
            Exit Function
        End If
    End If
    
    If Right(sDir, 1) <> "\" Then sDir = sDir + "\"
    
    nPosi = InStr(sDir, ":\")
    If nPosi > 0 Then
        sDrive = Left(sDir, nPosi)
        sDir = Mid(sDir, nPosi + 2)
    End If
    
    sParentDir = sDrive
    
    On Error Resume Next
    nPosi = InStr(sDir, "\")
    
    While nPosi > 0
        sParentDir = sParentDir + "\" + Left(sDir, nPosi - 1)
        
        On Error Resume Next
        sTmpDir = Dir(sParentDir, vbDirectory)
        
        If sTmpDir = "" Then  'directory does not exist, create it
            On Error GoTo errTrap
            MkDir sParentDir
        End If
        
        sTmpDir = Mid(sDir, nPosi + 1)
        
        sDir = sTmpDir
        
        nPosi = 0
        nPosi = InStr(sDir, "\")
    Wend
    
    Exit Function

errTrap:
    tfnErrHandler "fnCheckAndCreateDirectory", bPromptForCreate
    fnCheckAndCreateDirectory = "Failed to create Directory '" + sDir + "'"
End Function

'david 06/19/2008  #596302
'return error message if any
Public Sub CheckXlatorProgramPath()

    Const sEDIPROC_EXE As String = "EDIPROC.EXE"
    Const sEDIPROC_INI As String = "EDIPROC.INI"
    Const sEDIPROC_MDB As String = "EDIXLATE.MDB"
    Const sSECTION_EDIXLATOR As String = "EDI XLATOR"
    Const sKEY_PROGRAM As String = "Program"
    Const sLocalBin As String = "C:\FACTOR\EXECTRAK\BIN"
    Const sRemoteBin As String = "J:\FACTOR\EXECTRAK\BIN"
    
    
    Dim sExeFile As String, sMsg As String
    Dim sFile As String, sFileCurr As String

    'check current directory
    m_ProgramPath = App.path
    sExeFile = fnAddBkSlash(m_ProgramPath) + sEDIPROC_EXE
    On Error Resume Next
    sFileCurr = Dir(sExeFile)
    On Error GoTo 0

    'get the EDIPROC program path from factor.ini
    m_ProgramPath = Trim(tfnReadINI(sSECTION_EDIXLATOR, sKEY_PROGRAM, tfnGetWindowsDir + szFACTOR_INI))
    
    If m_ProgramPath <> "" Then
        sExeFile = fnAddBkSlash(m_ProgramPath) + sEDIPROC_EXE
        On Error Resume Next
        sFile = Dir(sExeFile)
        On Error GoTo 0
    End If
    
    If sFile = "" Then
        'check current directory
        m_ProgramPath = App.path
        sExeFile = fnAddBkSlash(m_ProgramPath) + sEDIPROC_EXE
        On Error Resume Next
        sFile = ""
        sFile = Dir(sExeFile)
        On Error GoTo 0
    
        If sFile = "" Then
            If sFile = "" Then
                'search local bin
                m_ProgramPath = sLocalBin
                sExeFile = fnAddBkSlash(m_ProgramPath) + sEDIPROC_EXE
                On Error Resume Next
                sFile = Dir(sExeFile)
                On Error GoTo 0
                
                If sFile = "" Then
                    'search remote bin
                    m_ProgramPath = sRemoteBin
                    sExeFile = fnAddBkSlash(m_ProgramPath) + sEDIPROC_EXE
                    On Error Resume Next
                    sFile = Dir(sExeFile)
                    On Error GoTo 0
                End If
            End If
        End If
    End If
    
    If sFileCurr <> "" And fnAddBkSlash(UCase(m_ProgramPath)) <> fnAddBkSlash(UCase(App.path)) Then
        MsgBox "WARNING:" + vbCrLf + vbCrLf + "Multiple copies of FACTOR EDI Translator found in the system," _
            + " and the FACTOR EDI Translator is launching from a directory different from the running program." _
            + " You may proceed, but you may not be using the latest EDI Maps." + vbCrLf + vbCrLf _
            + "Contact FACTOR Support if you have questions.", vbInformation
        subAddLog "WARNING: FACTOR EDI Translator is launching from a directory different from the running program."
    End If
    
End Sub
''''''''''''''''''''''''''

