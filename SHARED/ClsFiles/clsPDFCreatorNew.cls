VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPDFCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Boolean, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hSnapShot As Long) As Boolean
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Private Const SYNCHRONIZE = &H100000

' Enum
Public Enum pdfPathOptions
 Stroked = &H1
 Closed = &H2
 Filled = &H4
 Nil = &H0
End Enum

Public Enum pdfColorSpace
 pdfRGB = 0
 pdfGrayScale = 1
End Enum

Public Enum pdfFontStyle
 pdfNormal = 1
 pdfBold = 2
 pdfItalic = 3
 pdfBoldItalic = 4
End Enum

Public Enum pdfTextAlign
 pdfAlignLeft = 1
 pdfAlignRight = 2
 pdfCenter = 3
End Enum

Public Enum pdfPaperSize
  pdfA4 = 0
  pdfA3 = 1
  pdf8_5x11 = 2
  pdf8_5x12 = 3
  pdfUser = &HFF
End Enum

Public Enum pdfScaleMode
  pdf72PxInch = 1
  pdfInch = 2
  pdfMillimeter = 3
  pdfCentimeter = 4
End Enum

Public Enum pdfPrintOrientation
  pdfPortrait = 1
  pdfLandscape = 2
End Enum

Public Enum pdfObjectType
  pdfNull = &H0
  pdfFirstPage = &H1
  pdfEvenPages = &H2
  pdfOddPages = &H4
  pdfNotFirstPage = &H8
  pdfAllPages = pdfEvenPages Or pdfOddPages Or pdfFirstPage
End Enum

Private m_strSourceFile As String
Private m_strTargetFile As String
Private m_numLinesPerPage As Long
Private m_strFontName As String
Private m_sngFontSize As Single
'

'
Private Sub Class_Terminate()

End Sub

''''''''''''''''''''  INTERFACE PROPERTIES AND FUNCTIONS  '''''''''''''''''''''''''''''''''''
Public Property Let PageBreakBefore(ByVal sPageBr As String)
    
End Property

Public Property Get PageBreakBefore() As String
    
End Property

Public Property Get sourceFile() As String
    sourceFile = m_strSourceFile
End Property

Public Property Let sourceFile(ByVal vNewValue As String)
    m_strSourceFile = vNewValue
   
End Property

Public Property Get targetFile() As String
    targetFile = m_strTargetFile
End Property

Public Property Let targetFile(ByVal vNewValue As String)
    m_strTargetFile = vNewValue
End Property

Public Property Get RecordsPerPage() As Long
    RecordsPerPage = m_numLinesPerPage
End Property

Public Property Let RecordsPerPage(ByVal vNewValue As Long)
    m_numLinesPerPage = vNewValue
End Property

Public Property Get Font() As String
    Font = m_strFontName
End Property

Public Property Let Font(ByVal vNewValue As String)
    m_strFontName = vNewValue
End Property

Public Property Get FontSize() As Single
    FontSize = m_sngFontSize
End Property

Public Property Let FontSize(ByVal vNewValue As Single)
    m_sngFontSize = vNewValue
End Property

Public Function CreatePDFFile(Optional bRaiseError As Boolean = False) As String
    Dim strConvertApp As String
    Dim strExeAndCmd As String
    Dim sMsg As String
    
    strConvertApp = IIf(Right(App.Path, 1) = "\", App.Path & "txt2pdf.exe", App.Path & "\txt2pdf.exe")
    If Dir(strConvertApp) = "" Then
        CreatePDFFile = "Cannot CreatePDFFile(); '" & strConvertApp & "' was not found"
        If bRaiseError Then Err.Raise -1, "clsPDFCreator.CreatePDFFile", CreatePDFFile
        Exit Function
    End If
    
    If m_strSourceFile = "" Then
        CreatePDFFile = "Cannot CreatePDFFile(); No Source File was specified"
        If bRaiseError Then Err.Raise -1, "clsPDFCreator.CreatePDFFile", CreatePDFFile
        Exit Function
    End If
    
    If Dir(m_strSourceFile) = "" Then
        CreatePDFFile = "Cannot CreatePDFFile(); Source File was not found '" & m_strSourceFile & "'"
        If bRaiseError Then Err.Raise -1, "clsPDFCreator.CreatePDFFile", CreatePDFFile
        Exit Function
    End If
    
    If m_strTargetFile = "" Then
        CreatePDFFile = "Cannot CreatePDFFile(); No Target File was specified"
        If bRaiseError Then Err.Raise -1, "clsPDFCreator.CreatePDFFile", CreatePDFFile
        Exit Function
    End If
    
    strExeAndCmd = Chr$(34) & strConvertApp & Chr$(34)
    strExeAndCmd = strExeAndCmd & " " & Chr$(34) & m_strSourceFile & Chr$(34)
    strExeAndCmd = strExeAndCmd & " " & Chr$(34) & m_strTargetFile & Chr$(34)
    strExeAndCmd = strExeAndCmd & " -pfs9 -pfn200 -lpp66"
    
    If Not ShellProgramAndWait(strExeAndCmd, vbMinimizedNoFocus, sMsg) Then
        CreatePDFFile = "CreatePDFFile() failed; ShellProgramAndWait returned an error: " & sMsg
        If bRaiseError Then Err.Raise -1, "clsPDFCreator.CreatePDFFile", CreatePDFFile
    End If
    
    If Dir(m_strTargetFile) = "" Then
        Sleep 1000
        If Dir(m_strTargetFile) = "" Then
            CreatePDFFile = "CreatePDFFile() failed; Target file '" & m_strTargetFile & "' not created"
            If bRaiseError Then Err.Raise -1, "clsPDFCreator.CreatePDFFile", CreatePDFFile
            Exit Function
        End If
    End If
    
End Function

'****************************
' Include the command line in strExePathCmd
' Example:
' sExeFile + " " + sParm
'****************************
Private Function ShellProgramAndWait(ByVal strExePathCmd As String, windowstyle As VbAppWinStyle, sMsg As String, Optional timeoutSeconds As Long = 2100) As Boolean
    Dim ProcessID As Long
    Dim ProcessHandle As Long
    Dim wfsoReply As Long
    Dim nWaitStart As Long
    
    ' In VB4, an error occurs if Shell
    ' fails to start the program
    On Error GoTo FINISHED
    sMsg = ""
    
    ' Shell the program, get its handle,
    ' and wait for it to terminate
    nWaitStart = Timer
    ProcessID = Shell(strExePathCmd, windowstyle)
    ProcessHandle = OpenProcess(SYNCHRONIZE, True, ProcessID)
    wfsoReply = WaitForSingleObject(ProcessHandle, 500)
    Do While wfsoReply = 258
        DoEvents
        If (Timer - nWaitStart) > timeoutSeconds Then       ' 35 minutes
            ShellProgramAndWait = False
            sMsg = "Timed out waiting for " & vbCrLf & "'" & strExePathCmd & "' to finish"
            GoTo FINISHED
        End If
        wfsoReply = WaitForSingleObject(ProcessHandle, 500)
    Loop
    
    ShellProgramAndWait = True
    CloseHandle ProcessHandle
    
    Err.Clear
FINISHED:
    If Err.Number <> 0 Then
        sMsg = Err.Description
        ShellProgramAndWait = False
    End If
End Function


