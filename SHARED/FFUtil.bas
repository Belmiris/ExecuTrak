Attribute VB_Name = "FlatFileUtil"
Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'On 05/09/2002, we did the following two things:
'(1)Add output file information to the INI file
'(2)The INI File Can be set up by the calling program
' If you do not need output file setup, you still work the same way as
' the template. If you do need output file setup, See subGetInfo and subinitilize
' for detail. (the program ZZPSPRI is the first program with output set up
' ---- Weigong Jiang
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    Private Const FF_PROC_INIFILE = "FACTOR.INI"

    Private FF_PROC_INIFILE As String  '= "FACTOR.INI"
    Private Const szEMPTY = ""

    Private Const ERR_AUTO_FAILED = "Auto processing failed."
    
    Private Const DATA_BACKUP_PATH = "BAK"
    Private Const DATA_BACKUP_EXTN = "SAV"
    Private Const LOG_NAME_EXTN As String = "LOG"

    Private Const PRINT_MARGIN_LEFT = 150     'Pixel
    Private Const PRINT_MARGIN_RIGHT = 150    'Pixel
    Private Const PRINT_MARGIN_TOP = 250      'Pixel
    Private Const PRINT_MARGIN_BOTTOM = 0   'Pixel

    Private Const MAX_STRING_LENGTH = 512

    Public Const CLP_IDX_IPATH = 0
    Public Const CLP_ID_IPATH = "IFilePath"
    Public Const CLP_IDX_ITYPE = 1
    Public Const CLP_ID_ITYPE = "IFileType"
    Public Const CLP_IDX_INAME = 2
    Public Const CLP_ID_INAME = "IFileName"
    Public Const CLP_IDX_MODE = 3
    Public Const CLP_ID_MODE = "RunMode"
    Public Const CLP_IDX_BKFILE = 4
    Public Const CLP_ID_BKFILE = "BackupFileFlag"
    Public Const CLP_IDX_RMFILE = 5
    Public Const CLP_ID_RMFILE = "RemoveOriginal"
    Public Const CLP_IDX_BKPATH = 6
    Public Const CLP_ID_BKPATH = "BackupPath"
    Public Const CLP_IDX_BKNAME = 7
    Public Const CLP_ID_BKNAME = "BackupFileName"
    Public Const CLP_IDX_BKTYPE = 8
    Public Const CLP_ID_BKTYPE = "BackupFileType"
    Public Const CLP_IDX_WPATH = 9
    Public Const CLP_ID_WPATH = "WorkPath"
    Public Const CLP_IDX_LPATH = 10
    Public Const CLP_ID_LPATH = "LFilePath"
    Public Const CLP_IDX_LNAME = 11
    Public Const CLP_ID_LNAME = "LFileName"
    Public Const CLP_IDX_LTYPE = 12
    Public Const CLP_ID_LTYPE = "LFileType"
    Public Const CLP_IDX_WTLOG = 13
    Public Const CLP_ID_WTLOG = "WriteLogFlag"
    Public Const CLP_IDX_OPATH = 14
    Public Const CLP_ID_OPATH = "OFilePath"
    Public Const CLP_IDX_OTYPE = 15
    Public Const CLP_ID_OTYPE = "OFileType"
    Public Const CLP_IDX_ONAME = 16
    Public Const CLP_ID_ONAME = "OFileName"
    
    Public Const CLP_PARM_COUNT = 17
    Private aryCmdLineParms(CLP_PARM_COUNT - 1) As String

'status bar colors
    
    Public Const RP_MASK_AUTO = 1
    Public Const RP_MASK_BACK = 2
    Public Const RP_MASK_WTLOG = 4
    Public Const RP_MASK_BKFILE = 8
    Public Const RP_MASK_RMFILE = 16
    Public Const RP_MASK_FILEWRITE = 32
    Public Const PV_AUTO = "AUTO"
    Public Const PV_BACK = "BACK"
    Public Const PV_TRUE = "TRUE"
    Public Const PV_FALSE = "FALSE"
    Public Const FILE_MODE_READ = 1
    Public Const FILE_MODE_WRITE = 2
    
    Public dbLocal As Database
    
    Public Type tpFileInfo
        m_sPath As String
        m_sFile As String
        m_sType As String
    End Type
    
    Public udtInputInfo As tpFileInfo
    Public udtLogInfo As tpFileInfo
    Public udtBackupInfo As tpFileInfo
    Public udtOutputInfo As tpFileInfo
    Public m_sIniFile As String
    
    Public m_sWorkPath As String

    Private m_nLogFile As Integer
    Private m_sLFFullName As String
    Private m_bWriteLogFile As Boolean
    Private m_sINIParmSetion As String
    Public m_nRunParm As Integer
    Public bBatchMode As Boolean
    Private lFileSize As Long
    Private lFileCursor As Long
    Private m_nFileHandle As Integer
    Private m_sOutputFName As String
    
    Private Const INI_BUFFRER_SIZE = 512
    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
        (ByVal lpApplicationName As String, _
         ByVal lpKeyName As Any, _
         ByVal lpDefault As String, _
         ByVal lpReturnedString As String, _
         ByVal nSize As Long, _
         ByVal lpFileName As String) As Long

    Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
        ByVal lpApplicationName As String, _
        ByVal lpKeyName As Any, _
        ByVal lpString As Any, _
        ByVal lpFileName As String) As Long
    
    Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" ( _
        ByVal lpBuffer As String, _
        ByVal nSize As Long) As Long

Public Function fnCreatePath(sPath As String) As Boolean
    Dim aryDirs() As String
    Dim sOldPath As String
    Dim i As Integer
    Dim i1 As Integer
    
    sOldPath = CurDir
    subParseString aryDirs, sPath, "\"
    i1 = 0
    If Mid(aryDirs(0), 2, 1) = ":" Then
        ChDrive aryDirs(0) & "\"
        ChDir "\"
        i1 = 1
    End If
    For i = i1 To UBound(aryDirs)
        If Trim(aryDirs(i)) <> "" Then
            If Not fnIsFilePath(aryDirs(i)) Then
                MkDir UCase(aryDirs(i))
            End If
            ChDir aryDirs(i)
        End If
    Next i
    ChDir sOldPath
    fnCreatePath = True
    Exit Function
errCreateDir:
    fnCreatePath = False
End Function

Public Function fnEOF() As Boolean
    fnEOF = EOF(m_nFileHandle)
End Function

Public Function fnFileType(sName As String) As String

    Dim nPos As Integer
    
    fnFileType = ""
    nPos = Len(sName)
    If nPos > 0 Then
        Do
            If Mid(sName, nPos, 1) = "." Then
                fnFileType = Right(sName, Len(sName) - nPos)
                Exit Function
            End If
            nPos = nPos - 1
        Loop Until nPos <= 0
    End If
End Function




Public Function fnGetSubPath(sMain As String, _
                             sPath As String) As String
                             
    Dim nTemp As Integer
    Dim sTemp1 As String
    
    fnGetSubPath = sPath
    sTemp1 = UCase(sMain)
    subAddSlash sTemp1
    nTemp = Len(sTemp1)
    If nTemp > 0 Then
        If Left(UCase(sPath), nTemp) = sTemp1 Then
            nTemp = Len(sPath) - nTemp
            If nTemp > 0 Then
                fnGetSubPath = Right(sPath, nTemp)
            Else
                fnGetSubPath = ""
            End If
        End If
    End If
End Function

Public Function fnIsWholePath(sPath As String) As Boolean
    If InStr(sPath, ":") > 0 Or Left(sPath, 1) = "\" Then
        fnIsWholePath = True
    Else
        fnIsWholePath = False
    End If
End Function

Public Function fnMakeFName(sName As String, _
                             sType As String) As String
    If InStr(sName, ".") > 0 Then
        fnMakeFName = sName
    Else
        If sName = "" And sType = "" Then
            fnMakeFName = ""
        Else
            fnMakeFName = sName & "." & sType
        End If
    End If
End Function


Private Function fnMatchPart(sSrc As String, _
                             sDest As String) As Boolean
    fnMatchPart = False
    If sSrc <> "" Then
        If sSrc = Left(sDest, Len(sSrc)) Then
            fnMatchPart = True
        End If
    End If
    
End Function

Public Function fnNeedBackupFile() As Boolean
    fnNeedBackupFile = fnTestFlag(m_nRunParm, RP_MASK_BKFILE)
End Function

Public Function fnOutputFName() As String
    fnOutputFName = m_sOutputFName
End Function

Public Function fnReadINI(szSection As String, szKey As String, szINIFile As String) As String

    Dim nLength As Integer 'length of the value returned for api call
    Dim szINI As String    'string to hold the value retrieved

    szINI = Space(INI_BUFFRER_SIZE) 'clear and make the string fixed length
    
    'get the [value] for the [section], [key], and ini file sent
    nLength = GetPrivateProfileString(szSection, szKey, szEMPTY, szINI, INI_BUFFRER_SIZE, szINIFile)
    
    If nLength <> 0 Then 'if length positive [value] has been found
        szINI = Left(szINI, nLength) 'make it a basic string
    Else
        szINI = ""
    End If
    
    fnReadINI = szINI 'return the value

End Function
Public Function fnIsFile(ByVal szFilename As String) As Boolean
    
    On Error GoTo errNotFile

    fnIsFile = False
    If InStr(szFilename, "?") > 0 Then
        Exit Function
    End If
    If InStr(szFilename, "*") > 0 Then
        Exit Function
    End If
    If Trim(szFilename) <> "" Then
        Open szFilename For Input As #29
        Close #29
        fnIsFile = True
    End If
    Exit Function
errNotFile:
    #If DEVELOP Then
        MsgBox "Error # " & Err.Number & vbCrLf & "Error Message: " & Err.Description & " - " & szFilename
    #End If
End Function

Public Function fnGetSystemDir(Optional vAddSlash As Variant) As String
    
    Dim nLength As Integer     'length returned from API call
    Dim szSystemDir As String  'temp string to hold system directory
    
    szSystemDir = Space(MAX_STRING_LENGTH) 'set the string to a fixed length for API call, pad with spaces
    
    nLength = GetSystemDirectory(szSystemDir, MAX_STRING_LENGTH) 'call the API function to get the system directory
  
    szSystemDir = Left(szSystemDir, nLength) 'trim off the excess spaces
    
    If Not IsMissing(vAddSlash) Then
        If Right(szSystemDir, 1) <> "\" And vAddSlash Then 'add a slash if it needs one
            szSystemDir = szSystemDir + "\"
        End If
    End If
    
    fnGetSystemDir = szSystemDir 'return system directory back to the caller

End Function

Public Function fnCStr(vTemp As Variant) As String

    If IsNull(vTemp) Then
        fnCStr = ""
    Else
        fnCStr = Trim(vTemp)
    End If
    
End Function

Private Function fnStr2BoolTrue(sTemp As String) As Boolean
    fnStr2BoolTrue = False
    If IsNumeric(sTemp) Then
        If Val(sTemp) <> 0 Then
            fnStr2BoolTrue = True
        End If
    ElseIf UCase(sTemp) = PV_TRUE Then
        fnStr2BoolTrue = True
    End If
End Function

Public Function fnStripFileName(sFile As String) As String

    Dim nPos As Integer
    nPos = InStr(sFile, ".")
    If nPos > 0 Then
        fnStripFileName = Left(sFile, nPos - 1)
    Else
        fnStripFileName = sFile
    End If
End Function

Public Function fnGetOutputFile(sPath As String, _
                                ByVal sPrefix As String, _
                                sType As String, _
                                bFullName As Boolean) As String

    Dim nSeq2 As Integer
    Dim sName As String
    
    On Error GoTo errName
    sName = sPath & sPrefix & "." & sType
    If Dir(sName) = "" Then
        If bFullName Then
            fnGetOutputFile = sName
        Else
            fnGetOutputFile = sPrefix & "." & sType
        End If
    Else
        If Len(sPrefix) >= 8 Then
            sPrefix = Left(sPrefix, 7)
        End If
        nSeq2 = vbKeyA - 1
        Do
            nSeq2 = nSeq2 + 1
            sName = sPath & sPrefix & Chr(nSeq2) & "." & sType
        Loop Until Dir(sName) = "" Or nSeq2 >= vbKeyZ - 1
        If bFullName Then
            fnGetOutputFile = sName
        Else
            fnGetOutputFile = sPrefix & Chr(nSeq2) & "." & sType
        End If
    End If
    Exit Function
errName:
    fnGetOutputFile = ""
End Function

Public Function fnTestFlag(ByVal nParm As Integer, _
                           ByVal nFunc As Integer) As Boolean

    If nParm And nFunc Then
        fnTestFlag = True
    Else
        fnTestFlag = False
    End If
    
End Function

Public Sub subAppendType(sFile As String, _
                         sType As String)
    Dim sChar As String
    Dim i As Integer
    
    i = Len(sFile) + 1
    While i > 1 And sChar <> "." And sChar <> "\"
        i = i - 1
        sChar = Mid(sFile, i, 1)
    Wend
    If sChar <> "." Then
        sFile = sFile & "." & sType
    End If
    
End Sub

Private Sub subAutoRun()
    Const SUB_NAME = "subAutoRun"

    Dim sTemp As String
    Dim sNewFile As String
    Dim sOutput As String
    Dim i As Integer
    Dim aryFiles() As String
    Dim sPath  As String
    Dim sFile As String
    
    On Error GoTo errAutoRun
    
    subPrepareLog
    sTemp = udtInputInfo.m_sPath
    sNewFile = udtInputInfo.m_sFile
    subAppendType sNewFile, udtInputInfo.m_sType
    subAddSlash sTemp
    sTemp = sTemp & sNewFile
    If fnTestFlag(m_nRunParm, RP_MASK_FILEWRITE) Then
        bBatchMode = False
        fnProcessFile sTemp
        subCloseFile
        subEnablePrint True
        LogForm.subEnableCancel True
    Else
        If fnIsFile(sTemp) Then
            'Single file is going to be processed
            bBatchMode = False
            fnProcessFile sTemp
            subCloseFile
            subBackupFile sTemp
            subRemoveFile sTemp
            subEnablePrint True
            LogForm.subEnableCancel True
        Else
            'Multiple files
            LogForm.lstFile.Path = udtInputInfo.m_sPath
            If InStr(udtInputInfo.m_sType, ".") > 0 Then
                If InStr(udtInputInfo.m_sType, "*") > 0 Then
                    sTemp = udtInputInfo.m_sType
                Else
                    sTemp = "*" & udtInputInfo.m_sType
                End If
            Else
                sTemp = "*." & udtInputInfo.m_sType
            End If
            LogForm.lstFile.Pattern = sTemp
            If LogForm.lstFile.ListCount > 0 Then
                LogForm.ShowProgress 0
                LogForm.ShowProgressbar True
                lFileCursor = 0
                subSetFileSize
                ReDim aryFiles(LogForm.lstFile.ListCount - 1)
                sPath = udtInputInfo.m_sPath
                subAddSlash sPath
                For i = 0 To LogForm.lstFile.ListCount - 1
                    aryFiles(i) = LogForm.lstFile.List(i)
                    sFile = sPath & aryFiles(i)
                    'Single file is going to be processed
                    fnProcessFile sFile
                    subCloseFile
                    subBackupFile sFile
                    subRemoveFile sFile
                    subEnablePrint True
                    LogForm.subEnableCancel True
                Next i
            Else
                subWriteLog "No files (" & sTemp & ") found in the input directory (" & udtInputInfo.m_sPath & ")"
                subWriteLog "Auto processing stopped."
            End If
        End If
    End If
    Exit Sub
errAutoRun:
    Dim sMsg As String
    sMsg = "The following error ccurred in function: " & SUB_NAME & vbCrLf
    sMsg = sMsg & "Error # " & Err.Number & ", " & Err.Description & vbCrLf & ERR_AUTO_FAILED
    subWriteLog sMsg
    subSetButtonStatus
End Sub

Public Function fnConfirmed(sMsg As String) As Integer
    
    fnConfirmed = MsgBox(sMsg, vbQuestion + vbYesNo)
    
End Function


Public Sub subBackupFile(sFile As String)

    Dim sTemp As String
    Dim nErr As Integer
    Dim sDataPath As String
    Dim sFullName As String
    
    If fnTestFlag(m_nRunParm, RP_MASK_BKFILE) Then
        On Error GoTo errCopy
        sDataPath = udtBackupInfo.m_sPath
        sTemp = Dir(sDataPath, vbDirectory)
        If sTemp = "" Then
            MkDir sDataPath
        End If
        subAddSlash sDataPath
        If udtBackupInfo.m_sFile = "" Then
            sFullName = fnFileNameForToday("", sDataPath, udtBackupInfo.m_sType)
        Else
            sFullName = sDataPath & fnMakeFName(udtBackupInfo.m_sFile, udtBackupInfo.m_sType)
        End If
        FileCopy sFile, sFullName
    End If
    Exit Sub
errCopy:
    sTemp = Err.Description
    nErr = Err.Number
'    subShowMessage "Cannot backup data file because of the following error:" & vbCrLf _
'            & "VB Error # " & nErr & vbCrLf & sTemp
    subWriteLog "Cannot backup data file. VB Error #: " & nErr & "Description: " & sTemp
    subWriteLog "Input file name " & sFile
    subWriteLog "Output file name " & sFullName

End Sub

Public Sub subPutLine(sLine As String)
    
    On Error GoTo errWriteLine
    Open m_sOutputFName For Append As #m_nFileHandle
    Print #m_nFileHandle, sLine
    subCloseFile
    lFileCursor = lFileCursor + Len(sLine) + 2
    Exit Sub
    
errWriteLine:
    subWriteLog "Unable to write to output file: " & m_sOutputFName
    subWriteLog "Error " & Err.Number & ", " & Err.Description
End Sub

Public Sub subRemoveFile(sFile As String)
    If fnTestFlag(m_nRunParm, RP_MASK_RMFILE) Then
        Kill sFile
    End If
End Sub


Public Sub subCenterForm(frmCurrent As Form, Optional vParentForm As Variant)
  
    If IsMissing(vParentForm) Then
        frmCurrent.Left = (Screen.Width - frmCurrent.Width) \ 2
        frmCurrent.Top = (Screen.Height - frmCurrent.Height) \ 2
    Else
        
        If vParentForm.Width > frmCurrent.Width Then
            frmCurrent.Left = vParentForm.Left + (vParentForm.Width - frmCurrent.Width) \ 2
        Else
            frmCurrent.Left = (Screen.Width - frmCurrent.Width) \ 2
        End If

        If vParentForm.Height > frmCurrent.Height Then
            frmCurrent.Top = vParentForm.Top + (vParentForm.Height - frmCurrent.Height) \ 2
        Else
            frmCurrent.Top = (Screen.Height - frmCurrent.Height) \ 2
        End If
    End If
    
End Sub


Public Sub subCloseFile()
    Close #m_nFileHandle
End Sub

Public Sub subParseFile(sPath As String, _
                        sFile As String, _
                        sFullName As String)

    Dim i As Integer
    Dim nLen As Integer
    
    nLen = Len(sFullName)
    i = nLen
    Do While i > 0
        If Mid(sFullName, i, 1) = "\" Then
            Exit Do
        End If
        i = i - 1
    Loop
    If i > 0 Then
        sPath = Left(sFullName, i - 1)
    End If
    i = nLen - i
    If i > 0 Then
        sFile = Right(sFullName, i)
    End If
    
End Sub

Public Sub subGoLevelUp(sPath As String)

    Dim nPos As Integer
    Dim nLen As String
    
    nPos = Len(sPath)
    nLen = nPos
    Do While nPos > 0
        If Mid(sPath, nPos, 1) = "\" Then
            If nPos > 0 And nPos < nLen Then
                sPath = Left(sPath, nPos - 1)
                Exit Do
            End If
        End If
        nPos = nPos - 1
    Loop
End Sub

Public Sub subAddSlash(sStr As String)
    If sStr <> "" Then
        If Right(sStr, 1) <> "\" Then
            sStr = sStr & "\"
        End If
    End If
End Sub

Public Sub subPrepareLog()
    
    If fnTestFlag(m_nRunParm, RP_MASK_WTLOG) Then
        subAddSlash udtLogInfo.m_sPath
        If udtLogInfo.m_sFile = "" Then
            m_sLFFullName = fnFileNameForToday("", udtLogInfo.m_sPath, udtLogInfo.m_sType)
        Else
            m_sLFFullName = udtLogInfo.m_sPath & fnMakeFName(udtLogInfo.m_sFile, udtLogInfo.m_sType)
        End If
        m_nLogFile = FreeFile
        On Error GoTo errCreateLog
        Open m_sLFFullName For Output As #m_nLogFile
        Close #m_nLogFile
        m_bWriteLogFile = True
    Else
        m_bWriteLogFile = False
    End If
    Exit Sub
errCreateLog:
    Dim sError As String
    sError = "Error: " & Err.Number & ", " & Err.Description
    LogForm.ShowLog "Cannot create file for log(" & m_sLFFullName & ") because of the following error:"
    LogForm.ShowLog sError
    LogForm.ShowLog "Log will be shown on the screen only"
    m_bWriteLogFile = False
End Sub

Public Function fnFileNameForToday(sPref As String, _
                                   sPath As String, _
                                   sType As String) As String

    Dim sPrefix As String
    Dim nSeq1 As Integer
    Dim nSeq2 As Integer
    Dim sName As String
    
    sPrefix = sPref & Format(Date, "MMDDYY")
    nSeq1 = 0
    nSeq2 = vbKeyA
    Do
        sName = sPath & sPrefix & Chr(nSeq2) & CStr(nSeq1) & "." & sType
        nSeq1 = nSeq1 + 1
        If nSeq1 > 9 Then
            nSeq1 = 0
            nSeq2 = nSeq2 + 1
        End If
    Loop Until Dir(sName) = "" Or nSeq2 >= vbKeyZ
    
    fnFileNameForToday = sName
    
End Function


Public Function fnPrepareFile(sFName As String) As Boolean
    
    fnPrepareFile = False
    On Error GoTo errOpenFile
    m_nFileHandle = FreeFile

    If fnTestFlag(m_nRunParm, RP_MASK_FILEWRITE) Then
        If fnIsFile(sFName) Then
            If vbNo = fnConfirmed("File: " & sFName & " already exists." & vbCrLf & "Do you want to over write it?") Then
                Exit Function
            End If
        End If
        Open sFName For Output As #m_nFileHandle
        subCloseFile
        lFileCursor = 0
        fnPrepareFile = True
        LogForm.ShowProgress 0
        LogForm.ShowProgressbar True
        m_sOutputFName = sFName
    Else
        If bBatchMode Then
            Open sFName For Input As #m_nFileHandle
        Else
            lFileSize = FileLen(sFName)
            lFileCursor = 0
            If lFileSize = 0 Then
                subWriteLog "File name and path: " & sFName
            Else
                Open sFName For Input As #m_nFileHandle
            End If
            LogForm.ShowProgress 0
            LogForm.ShowProgressbar True
        End If
        fnPrepareFile = True
    End If
    Exit Function
errOpenFile:
    subWriteLog "Error " & Err.Number & ", " & Err.Description
End Function


Public Function fnGetLine() As String
    
    Line Input #m_nFileHandle, fnGetLine
    lFileCursor = lFileCursor + Len(fnGetLine) + 2
    LogForm.ShowProgress fnProgress
    
End Function


Private Function fnProgress() As Single
    If lFileSize <> 0 Then
        fnProgress = lFileCursor / lFileSize
        If fnProgress > 1 Then
            fnProgress = 1
        End If
    End If
End Function

Public Sub subPrint(lstOutput As ListBox)
    Dim i As Integer
    Dim nLeft As Integer
    Dim nTop As Integer
    Dim nBottom As Integer
    
    nLeft = PRINT_MARGIN_LEFT * Printer.TwipsPerPixelX
    nTop = PRINT_MARGIN_TOP * Printer.TwipsPerPixelY
    nBottom = Printer.Height - (nTop + PRINT_MARGIN_BOTTOM * Printer.TwipsPerPixelY)
    
    Printer.CurrentY = nTop
    For i = 0 To lstOutput.ListCount - 1
        Printer.CurrentX = nLeft
        Printer.Print lstOutput.List(i)
        If Printer.CurrentY >= nBottom Then
            Printer.NewPage
            Printer.CurrentY = nTop
        End If
    Next i
    Printer.EndDoc

End Sub

Public Sub Main()
    Dim szCommand As String
    
    szCommand = Command
    
    subInitialize szCommand
    
    subParseCmdLine szCommand
    subAddSlash m_sWorkPath
    subProcessPath udtLogInfo.m_sPath
    subProcessPath udtBackupInfo.m_sPath
    subProcessPath udtInputInfo.m_sPath
    subProcessPath udtOutputInfo.m_sPath
    
    LogForm.subSetFileInfo
    
    If fnTestFlag(m_nRunParm, RP_MASK_BACK) Then
        subAutoRun
        End
    Else
        If fnTestFlag(m_nRunParm, RP_MASK_AUTO) Then
            LogForm.Show
            LogForm.EnableButtons False
            DoEvents
            subAutoRun
            LogForm.EnableButtons True
            LogForm.subEnableProcess False
        Else
            LogForm.Show
        End If
    End If
End Sub

Private Sub subInitialize(sCommand As String)
    
    Dim aryInfo() As String
    Dim i As Integer
    Dim nUbound As Integer
    
    #If PROTOTYPE Then
        LogForm.mnuPrint1.Visible = False
        LogForm.mnuPrint.Visible = True
        Exit Sub
    #End If
    
    subGetInfo aryInfo
    LogForm.efraToolBar.FMName = aryInfo(0)
    LogForm.Caption = aryInfo(1)
    
    'Set Default Ini name and Section
    m_sINIParmSetion = aryInfo(0)
    FF_PROC_INIFILE = "FACTOR.INI"
    'We can overwite it
    nUbound = UBound(aryInfo)
    If nUbound >= 2 Then
        If Trim(aryInfo(2)) <> szEMPTY Then
            FF_PROC_INIFILE = Trim(aryInfo(2)) & ".INI"
        End If
    End If
    If nUbound >= 3 Then
        If Trim(aryInfo(3)) <> szEMPTY Then
            m_sINIParmSetion = Trim(aryInfo(3))
        End If
    End If
    subGetPrintMenu aryInfo
    If UBound(aryInfo) = 0 Then
        LogForm.mnuPrint1.Visible = False
        LogForm.mnuPrint.Visible = True
        LogForm.mnuPrint.Caption = aryInfo(0)
    Else
        With LogForm
            .mnuPrint1.Visible = True
            .mnuPrint.Visible = False
            .mnuSubPrint(0).Caption = aryInfo(0)
            For i = 1 To UBound(aryInfo)
                Load .mnuSubPrint(i)
                .mnuSubPrint(i).Caption = aryInfo(i)
            Next i
        End With
    End If
    
    m_nRunParm = 0
    
    aryCmdLineParms(CLP_IDX_IPATH) = UCase(CLP_ID_IPATH)
    aryCmdLineParms(CLP_IDX_ITYPE) = UCase(CLP_ID_ITYPE)
    aryCmdLineParms(CLP_IDX_INAME) = UCase(CLP_ID_INAME)
    aryCmdLineParms(CLP_IDX_LPATH) = UCase(CLP_ID_LPATH)
    aryCmdLineParms(CLP_IDX_LNAME) = UCase(CLP_ID_LNAME)
    aryCmdLineParms(CLP_IDX_LTYPE) = UCase(CLP_ID_LTYPE)
    aryCmdLineParms(CLP_IDX_MODE) = UCase(CLP_ID_MODE)
    aryCmdLineParms(CLP_IDX_WTLOG) = UCase(CLP_ID_WTLOG)
    aryCmdLineParms(CLP_IDX_BKFILE) = UCase(CLP_ID_BKFILE)
    aryCmdLineParms(CLP_IDX_RMFILE) = UCase(CLP_ID_RMFILE)
    aryCmdLineParms(CLP_IDX_WPATH) = UCase(CLP_ID_WPATH)
    aryCmdLineParms(CLP_IDX_BKPATH) = UCase(CLP_ID_BKPATH)
    aryCmdLineParms(CLP_IDX_BKNAME) = UCase(CLP_ID_BKNAME)
    aryCmdLineParms(CLP_IDX_BKTYPE) = UCase(CLP_ID_BKTYPE)

    aryCmdLineParms(CLP_IDX_OPATH) = UCase(CLP_ID_OPATH)
    aryCmdLineParms(CLP_IDX_OTYPE) = UCase(CLP_ID_OTYPE)
    aryCmdLineParms(CLP_IDX_ONAME) = UCase(CLP_ID_ONAME)
    
    udtLogInfo.m_sType = LOG_NAME_EXTN
    udtBackupInfo.m_sType = DATA_BACKUP_EXTN
    subReadINIParms
    subSetFileMode FILE_MODE_READ
    
    If fnAllowStandalone Then
        If Not fnCheckRunMethod(sCommand) Then
            End
        End If
    End If
    
    If Not tfnAuthorizeExecute(sCommand) Then 'Check for handshake if not in the development mode
        End
    End If
    
    If tfnOpenDatabase Then
        Set dbLocal = tfnOpenLocalDatabase
        subInitErrorHandler
        tfnUpdateVersion
    Else
        Unload LogForm
        End
    End If
    
End Sub

Private Function fnCheckRunMethod(sCommand As String) As Boolean
    Dim sErrMsg As String
    
    If sCommand = t_szHandShake Then
        Load LogForm
    Else
        If Not fnRunWithCommandLine(sCommand, sErrMsg) Then
            If sErrMsg = "" Then
                frmSplash.Caption = "Select Data Sources"
                frmSplash.Show vbModal
            Else
                MsgBox sErrMsg, vbCritical
                Exit Function
            End If
        End If
    End If

    fnCheckRunMethod = True
End Function


Private Sub subInitErrorHandler()
    If objErrHandler Is Nothing Then
        Set objErrHandler = New clsErrorHandler
        With objErrHandler
            Set .FormParent = LogForm
            Set .DatabaseEngine = t_engFactor
            Set .LocalDatabase = dbLocal
        End With
    End If
End Sub

Private Sub subParseCmdLine(sCmdL As String)

    Dim aryParms() As String
    
    If Trim(sCmdL) <> "" Then
        subParseString aryParms, sCmdL, ","
        subSetTheValues aryParms
    End If

End Sub

Public Sub subParseLeftRight(sLeft As String, _
                              sRight As String, _
                              sSource As String, _
                              sDel As String)
    Dim nPos As Integer
    nPos = InStr(UCase(sSource), sDel)
    If nPos > 0 Then
        sLeft = Trim(Left(sSource, nPos - 1))
        nPos = Len(sSource) - nPos - Len(sDel) + 1
        If nPos > 0 Then
            sRight = Trim(Right(sSource, nPos))
        End If
    Else
        sLeft = Trim(sSource)
    End If
End Sub


Public Sub subEnablePrint(ByVal bFlag As Boolean)
    frmContext.ButtonEnabled(PRINT_UP) = bFlag
    LogForm.mnuPrint.Enabled = bFlag
    LogForm.mnuPrint1.Enabled = bFlag
End Sub


Public Sub subEnableSubPrint(ByVal nIndex As Integer, _
                             ByVal bFlag As Boolean)
    
    LogForm.mnuSubPrint(nIndex).Enabled = bFlag
    
End Sub


Public Sub subProcessPath(sPath As String)

    If sPath = "" Then
        sPath = m_sWorkPath
    Else
        If Not fnIsWholePath(sPath) Then
            subAddSlash m_sWorkPath
            sPath = m_sWorkPath & sPath
        End If
    End If
End Sub

Public Function fnIsFilePath(ByVal sPath As String) As Boolean

    On Error GoTo errChDir
    subAddSlash sPath
    If Dir(sPath & "*.*", vbNormal + vbDirectory) = "" Then
        fnIsFilePath = False
    Else
        fnIsFilePath = True
    End If
    Exit Function
    
errChDir:
    fnIsFilePath = False
End Function


Private Sub subReadINIParms()

    Dim j As Integer
    Dim sValue As String
    
    m_sIniFile = tfnGetWindowsDir(True) & FF_PROC_INIFILE
    For j = 0 To CLP_PARM_COUNT - 1
        sValue = fnReadINI(m_sINIParmSetion, aryCmdLineParms(j), m_sIniFile)
        subSetParmValue j, sValue
    Next j
    If m_sWorkPath = "" Then
        m_sWorkPath = App.Path
    End If
    subAddSlash m_sWorkPath
'    subProcessPath udtBackupInfo.m_sPath
'    subProcessPath udtInputInfo.m_sPath

End Sub

'
'Function : tfnWriteINI - writes a value to a windows INI file
'Variables: [section], [key], [value], and ini file name
'Return   : status of api call
'
Public Function fnWriteINI(szSection As String, szKey As String, szValue As String, szINIFile As String) As Boolean

    Dim bStatus As Boolean 'status returned from api call
    
    'write the [value] for the [section], [key], and ini file sent
    bStatus = WritePrivateProfileString(szSection, szKey, szValue, szINIFile)
    
    fnWriteINI = bStatus

End Function

Public Sub subSetFileMode(ByVal nFlag As Integer)
    If nFlag = FILE_MODE_WRITE Then
        subSetFlag m_nRunParm, RP_MASK_FILEWRITE, True
    Else
        subSetFlag m_nRunParm, RP_MASK_FILEWRITE, False
    End If
End Sub

Private Sub subSetFileSize()
    Dim sPath As String
    Dim i As Integer
    
    sPath = udtInputInfo.m_sPath
    subAddSlash sPath
    For i = 0 To LogForm.lstFile.ListCount - 1
        lFileSize = lFileSize + FileLen(sPath & LogForm.lstFile.List(i))
    Next i
    bBatchMode = True

End Sub

Public Sub subSetFlag(nParm As Integer, _
                      ByVal nMask As Integer, _
                      ByVal bFlag As Boolean)

    If bFlag Then
        nParm = nParm Or nMask
    Else
        nParm = nParm And Not nMask
    End If

End Sub


Private Sub subSetTheValues(aryParms() As String)
    
    Dim i As Integer
    Dim j As Integer
    Dim sName As String
    Dim sValue As String
    For i = 0 To UBound(aryParms)
        subParseLeftRight sName, sValue, aryParms(i), "="
        sName = UCase(sName)
        For j = 0 To CLP_PARM_COUNT - 1
            If sName = aryCmdLineParms(j) Then
                subSetParmValue j, sValue
                Exit For
            End If
        Next j
    Next i

End Sub

Private Sub subSetButtonStatus()
    
    LogForm.EnableButtons True
    LogForm.subEnableCancel True
    LogForm.subEnableProcess True

End Sub

Private Sub subSetParmValue(ByVal nIdx As Integer, _
                            sValue As String)
    If Trim(sValue) = "" Then
        Exit Sub
    End If
    Select Case nIdx
        Case CLP_IDX_IPATH
            udtInputInfo.m_sPath = sValue
        Case CLP_IDX_INAME
            udtInputInfo.m_sFile = sValue
        Case CLP_IDX_ITYPE
            udtInputInfo.m_sType = sValue
        Case CLP_IDX_OPATH
            udtOutputInfo.m_sPath = sValue
        Case CLP_IDX_ONAME
            udtOutputInfo.m_sFile = sValue
        Case CLP_IDX_OTYPE
            udtOutputInfo.m_sType = sValue
        Case CLP_IDX_LPATH
            udtLogInfo.m_sPath = sValue
        Case CLP_IDX_LNAME
            udtLogInfo.m_sFile = sValue
        Case CLP_IDX_LTYPE
            udtLogInfo.m_sType = sValue
        Case CLP_IDX_MODE
            sValue = UCase(sValue)
            If fnMatchPart(sValue, PV_AUTO) Then
                subSetFlag m_nRunParm, RP_MASK_AUTO, True
            Else
                subSetFlag m_nRunParm, RP_MASK_AUTO, False
            End If
            If fnMatchPart(sValue, PV_BACK) Then
                subSetFlag m_nRunParm, RP_MASK_BACK, True
            Else
                subSetFlag m_nRunParm, RP_MASK_BACK, False
            End If
        Case CLP_IDX_WTLOG
            If fnStr2BoolTrue(sValue) Then
                subSetFlag m_nRunParm, RP_MASK_WTLOG, True
            Else
                subSetFlag m_nRunParm, RP_MASK_WTLOG, False
            End If
        Case CLP_IDX_BKFILE
            If fnStr2BoolTrue(sValue) Then
                subSetFlag m_nRunParm, RP_MASK_BKFILE, True
            Else
                subSetFlag m_nRunParm, RP_MASK_BKFILE, False
            End If
        Case CLP_IDX_RMFILE
            If fnStr2BoolTrue(sValue) Then
                subSetFlag m_nRunParm, RP_MASK_RMFILE, True
            Else
                subSetFlag m_nRunParm, RP_MASK_RMFILE, False
            End If
        Case CLP_IDX_WPATH
            m_sWorkPath = sValue
        Case CLP_IDX_BKPATH
            udtBackupInfo.m_sPath = sValue
        Case CLP_IDX_BKNAME
            udtBackupInfo.m_sFile = sValue
        Case CLP_IDX_BKTYPE
            udtBackupInfo.m_sType = sValue
    End Select

End Sub

Public Sub subParseString(sParam() As String, _
                           sSrc As String, _
                           sDelim As String, _
                           Optional vStart As Variant, _
                           Optional vEnd As Variant)
                          
    If Trim(sSrc) = "" Then
        Exit Sub
    End If

    Const nArrayInc As Integer = 5
    Dim i1 As Integer
    Dim i2 As Integer
    Dim k As Integer
    Dim nEnd As Integer
    Dim sTemp As String
    Dim nDLen As Integer
    
    If IsMissing(vStart) Then
        i1 = 1
    Else
        i1 = vStart
    End If
    If IsMissing(vEnd) Then
        nEnd = Len(sSrc)
    Else
        nEnd = vEnd
    End If
    If i1 < 1 Then i1 = 1
    i2 = 1
    k = 0
    nDLen = Len(sDelim)
    ReDim sParam(nArrayInc)
    While i1 <= nEnd And i2 > 0 And i2 <= nEnd
        i2 = InStr(i1, sSrc, sDelim)
        If i2 >= i1 And i2 <= nEnd Then
            If k > UBound(sParam) Then
                ReDim Preserve sParam(k + nArrayInc)
            End If
            sTemp = Mid$(sSrc, i1, i2 - i1)
            If sTemp <> "" Or sDelim <> " " Then
                sParam(k) = sTemp
                k = k + 1
            End If
            i1 = i2 + nDLen
        End If
    Wend
    If i2 <= nEnd Then
        If k > UBound(sParam) Then
            ReDim Preserve sParam(k + nArrayInc)
        End If
        sParam(k) = Mid$(sSrc, i1, nEnd - i1 + 1)
        ReDim Preserve sParam(k)
    Else
        If k > 0 Then
            sParam(k - 1) = Mid$(sSrc, i1, nEnd - i1 + 1)
            ReDim Preserve sParam(k - 1)
        End If
    End If
End Sub

Public Sub subShowMainForm()
    LogForm.Show
    Screen.MousePointer = vbDefault
End Sub

Public Sub subWriteInOut()
    Dim sTemp As String
    
    fnWriteINI m_sINIParmSetion, CLP_ID_WPATH, m_sWorkPath, m_sIniFile
    fnWriteINI m_sINIParmSetion, CLP_ID_LPATH, fnGetSubPath(m_sWorkPath, udtLogInfo.m_sPath), m_sIniFile
    fnWriteINI m_sINIParmSetion, CLP_ID_LNAME, udtLogInfo.m_sFile, m_sIniFile
    fnWriteINI m_sINIParmSetion, CLP_ID_LTYPE, udtLogInfo.m_sType, m_sIniFile

    If fnTestFlag(m_nRunParm, RP_MASK_WTLOG) Then
        sTemp = PV_TRUE
    Else
        sTemp = PV_FALSE
    End If
    fnWriteINI m_sINIParmSetion, CLP_ID_WTLOG, sTemp, m_sIniFile
    If fnTestFlag(m_nRunParm, RP_MASK_AUTO) Then
        sTemp = PV_AUTO
    Else
        sTemp = ""
    End If
    fnWriteINI m_sINIParmSetion, CLP_ID_MODE, sTemp, m_sIniFile
    fnWriteINI m_sINIParmSetion, CLP_ID_IPATH, fnGetSubPath(m_sWorkPath, udtInputInfo.m_sPath), m_sIniFile
    fnWriteINI m_sINIParmSetion, CLP_ID_INAME, udtInputInfo.m_sFile, m_sIniFile
    fnWriteINI m_sINIParmSetion, CLP_ID_ITYPE, udtInputInfo.m_sType, m_sIniFile
    
    fnWriteINI m_sINIParmSetion, CLP_ID_OPATH, fnGetSubPath(m_sWorkPath, udtOutputInfo.m_sPath), m_sIniFile
    fnWriteINI m_sINIParmSetion, CLP_ID_ONAME, udtOutputInfo.m_sFile, m_sIniFile
    fnWriteINI m_sINIParmSetion, CLP_ID_OTYPE, udtOutputInfo.m_sType, m_sIniFile
    
    If Not fnTestFlag(m_nRunParm, RP_MASK_FILEWRITE) Then
        If fnTestFlag(m_nRunParm, RP_MASK_BKFILE) Then
            sTemp = PV_TRUE
        Else
            sTemp = PV_FALSE
        End If
        fnWriteINI m_sINIParmSetion, CLP_ID_BKFILE, sTemp, m_sIniFile
        If fnTestFlag(m_nRunParm, RP_MASK_RMFILE) Then
            sTemp = PV_TRUE
        Else
            sTemp = PV_FALSE
        End If
        fnWriteINI m_sINIParmSetion, CLP_ID_RMFILE, sTemp, m_sIniFile
        fnWriteINI m_sINIParmSetion, CLP_ID_BKPATH, fnGetSubPath(m_sWorkPath, udtBackupInfo.m_sPath), m_sIniFile
        fnWriteINI m_sINIParmSetion, CLP_ID_BKNAME, udtBackupInfo.m_sFile, m_sIniFile
        fnWriteINI m_sINIParmSetion, CLP_ID_BKTYPE, udtBackupInfo.m_sType, m_sIniFile
    End If
End Sub

Public Sub subWriteLog(sLog As String)
    LogForm.ShowLog sLog
    If m_bWriteLogFile Then
        m_nLogFile = FreeFile
        On Error GoTo errWriteLog
        Open m_sLFFullName For Append As #m_nLogFile
        Print #m_nLogFile, sLog
        Close #m_nLogFile
    End If
    Exit Sub
errWriteLog:
    LogForm.ShowLog "Cannot write to log file."
    LogForm.ShowLog "Error # " & Err.Number & ", Error: " & Err.Description
End Sub

'david 04/04/2001
'functions to handler the program when it is launching from the EC Scheduler
Private Function fnRunWithCommandLine(sCommand As String, sErrMsg As String) As Boolean
    Const CMD_LINE_DELIMITER = "ï"  '239   '"º"  '186
    
    Const PARM_MODE As Integer = 0
    Const PARM_DSN As Integer = 1
    Const PARM_USERID As Integer = 2
    Const PARM_PASSWORD As Integer = 3
    
    Dim aryParm() As String
    
    If sCommand = "" Then
        Exit Function
    End If
    
    If InStr(sCommand, CMD_LINE_DELIMITER) <= 0 Then
        Exit Function
    End If
    
    aryParm = Split(sCommand, CMD_LINE_DELIMITER)
    
    If UBound(aryParm) < PARM_PASSWORD Then
        Exit Function
    End If
    
    If UCase(aryParm(PARM_MODE)) <> "AUTO" Or _
       aryParm(PARM_DSN) = "" Or _
       aryParm(PARM_USERID) = "" Or _
       aryParm(PARM_PASSWORD) = "" Then
        Exit Function
    End If
    
    'connect automatically
    If Not frmSplash.Connect(aryParm(PARM_DSN), aryParm(PARM_USERID), _
       fnUncrypt(aryParm(PARM_PASSWORD)), sErrMsg) Then
        Exit Function
    End If
    
    're-construct command line
    sCommand = ""
    If UBound(aryParm) > PARM_PASSWORD Then
        Dim i As Integer
        sCommand = aryParm(PARM_PASSWORD + 1)
        For i = PARM_PASSWORD + 2 To UBound(aryParm)
            sCommand = sCommand + CMD_LINE_DELIMITER + aryParm(i)
        Next i
    End If
    
    fnRunWithCommandLine = True
End Function

Private Function fnUncrypt(sSource As String) As String

    Dim i As Integer
    Dim nLen As Integer
    Dim sTemp As String
    Dim nAsc As Integer
    
    sTemp = ""
    nLen = Len(sSource)
    For i = 3 To nLen
        nAsc = Asc(Mid(sSource, nLen - i + 3, 1))
        sTemp = sTemp & Chr(nAsc - 2 * nLen + i + 1)
    Next i
    fnUncrypt = sTemp
    
End Function
