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

' For help use the link
' http://www.verypdf.com/txt2pdf/help.htm

' based on output from fmvcpntf.4ge
' 136 max horizontal chars Landscape    ' Courier Font max size = 8
' 49 max vertical chars

' 83 max horizontal for Portrait        ' Courier Font max size = 10 - Use 9 though
' 66 max vertical for Portrait

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
Private m_numLinesPerPagePortrait As Long
Private m_numLinesPerPageLandscape As Long
Private m_strFontName As String
Private m_sngFontSize As Single
Private m_sngFontSizePortrait As Single
Private m_sngFontSizeLandscape As Single
Private m_Orientation As pdfPrintOrientation
Private m_bAutoAdjustOrientation As Boolean
Private m_strPageBreakBefore As String
'

Private Sub Class_Initialize()
    Me.Orientation = pdfPortrait
End Sub

''''''''''''''''''''  INTERFACE PROPERTIES AND FUNCTIONS  '''''''''''''''''''''''''''''''''''

'************************************************
'
'************************************************
Public Property Get AutoAdjustOrientation() As Boolean
    AutoAdjustOrientation = m_bAutoAdjustOrientation
End Property

Public Property Let AutoAdjustOrientation(ByVal vNewValue As Boolean)
    m_bAutoAdjustOrientation = vNewValue
End Property

'************************************************
'
'************************************************
Public Property Get PageBreakBefore() As String
    PageBreakBefore = m_strPageBreakBefore
End Property

Public Property Let PageBreakBefore(ByVal vNewValue As String)
    m_strPageBreakBefore = vNewValue
End Property

'************************************************
'
'************************************************
Public Property Get sourceFile() As String
    sourceFile = m_strSourceFile
End Property

Public Property Let sourceFile(ByVal vNewValue As String)
    m_strSourceFile = vNewValue
End Property

'************************************************
'
'************************************************
Public Property Get targetFile() As String
    targetFile = m_strTargetFile
End Property

Public Property Let targetFile(ByVal vNewValue As String)
    m_strTargetFile = vNewValue
End Property

'************************************************
'
'************************************************
Public Property Get Font() As String
    Font = IIf(m_strFontName <> "", m_strFontName, "Courier New")
End Property

Public Property Let Font(ByVal vNewValue As String)
    m_strFontName = vNewValue
End Property

'************************************************
'
'************************************************
Public Property Get FontSize() As Single
    FontSize = IIf(m_sngFontSize > 0, m_sngFontSize, 6.5)
End Property

Public Property Let FontSize(ByVal vNewValue As Single)
    m_sngFontSize = vNewValue
End Property

'************************************************
'
'************************************************
Public Property Get FontSizePortrait() As Single
    FontSizePortrait = m_sngFontSizePortrait
End Property

Public Property Let FontSizePortrait(ByVal vNewValue As Single)
    m_sngFontSizePortrait = vNewValue
End Property

'************************************************
'
'************************************************
Public Property Get FontSizeLandscape() As Single
    FontSizeLandscape = m_sngFontSizeLandscape
End Property

Public Property Let FontSizeLandscape(ByVal vNewValue As Single)
    m_sngFontSizeLandscape = vNewValue
End Property

'************************************************
'
'************************************************
Public Property Get Orientation() As pdfPrintOrientation
    Orientation = m_Orientation
End Property

Public Property Let Orientation(ByVal vNewValue As pdfPrintOrientation)
    m_Orientation = vNewValue
End Property

'************************************************
'
'************************************************
Public Property Get RecordsPerPage() As Long
    RecordsPerPage = IIf(m_numLinesPerPage > 0, m_numLinesPerPage, 66)
End Property

Public Property Let RecordsPerPage(ByVal vNewValue As Long)
    m_numLinesPerPage = vNewValue
End Property

'************************************************
'
'************************************************
Public Property Get RecordsPerPageLandscape() As Long
    RecordsPerPageLandscape = m_numLinesPerPageLandscape
End Property

Public Property Let RecordsPerPageLandscape(ByVal vNewValue As Long)
    m_numLinesPerPageLandscape = vNewValue
End Property

'************************************************
'
'************************************************
Public Property Get RecordsPerPagePortrait() As Long
    RecordsPerPagePortrait = m_numLinesPerPagePortrait
End Property

Public Property Let RecordsPerPagePortrait(ByVal vNewValue As Long)
    m_numLinesPerPagePortrait = vNewValue
End Property

'************************************************
'
'************************************************
Public Function CreatePDFFile(Optional bRaiseError As Boolean = False) As String
    Dim strConvertApp As String
    Dim strExeAndCmd As String
    Dim parameters As String
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
    
    If Me.AutoAdjustOrientation Then
        CreatePDFFile = fnCheckSetOrientation
        If CreatePDFFile <> "" Then Exit Function
    End If
    
    parameters = GetParameters()
    
    strExeAndCmd = Chr$(34) & strConvertApp & Chr$(34)
    strExeAndCmd = strExeAndCmd & " " & Chr$(34) & m_strSourceFile & Chr$(34)
    strExeAndCmd = strExeAndCmd & " " & Chr$(34) & m_strTargetFile & Chr$(34)
    strExeAndCmd = strExeAndCmd & " " & parameters
    
'    ' Testing
'    strExeAndCmd = Chr$(34) & strConvertApp & Chr$(34)
'    strExeAndCmd = strExeAndCmd & " " & Chr$(34) & m_strSourceFile & Chr$(34)
'    strExeAndCmd = strExeAndCmd & " " & Chr$(34) & m_strTargetFile & Chr$(34)
'    strExeAndCmd = strExeAndCmd & " -pfs9 -pfn200 -lpp66"
    
'    Debug.Print strExeAndCmd
    
    If Not ShellProgramAndWait(strExeAndCmd, vbMinimizedNoFocus, sMsg) Then
        CreatePDFFile = "CreatePDFFile() failed; ShellProgramAndWait returned an error: " & sMsg
        If bRaiseError Then Err.Raise -1, "clsPDFCreator.CreatePDFFile", CreatePDFFile
    End If
    
    If Dir(m_strTargetFile) = "" Then
        Sleep 2000
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

'************************************************
'
'************************************************
Public Function GetParameters() As String
    Dim lpp As String
    Dim pfn As String
    Dim pfs As String
    Dim pot As String
    Dim parameters As String
    
    pfs = GetFontSizeParameter()
    pfn = GetFontNameParameter()
    lpp = GetLinesPerPageParameter()
    pot = GetOrientationParameter()
    
    parameters = pfs & " " & pfn & " " & lpp
    If pot <> "" Then
        parameters = parameters + " " + pot
    End If
    
    GetParameters = parameters
    
End Function

'************************************************
'100  Times-Roman
'101  Times-Bold
'102  Times-Italic
'103  Times-Bold-Italic
'200  Courier
'201  Courier-Bold
'202  Courier-Oblique
'203  Courier-Bold-Oblique
'300  Helvetica
'301  Helvetica-Bold
'302  Helvetica-Oblique
'303  Helvetica-Bold-Oblique
'400  Symbol
'************************************************
Private Function GetFontNameParameter() As String
    Dim fName As String
    Dim fCode As String
    
    fName = IIf(Me.Font <> "", UCase(Me.Font), "COURIER NEW")
    
    Select Case fName
        Case "TIMES-ROMAN"
            fCode = "100"
            
        Case "TIMES-BOLD"
            fCode = "101"
            
        Case "TIMES-ITALIC"
            fCode = "102"
            
        Case "TIMES-BOLD-ITALIC"
            fCode = "103"
            
        Case "COURIER"
        Case "COURIER NEW"
            fCode = "200"
            
        Case "COURIER-BOLD"
            fCode = "201"
            
        Case "COURIER-OBLIQUE"
            fCode = "202"
            
        Case "COURIER-BOLD-OBLIQUE"
            fCode = "203"
            
        Case "HELVETICA"
            fCode = "300"
            
        Case "HELVETICA-BOLD"
            fCode = "301"
            
        Case "HELVETICA-OBLIQUE"
            fCode = "302"
            
        Case "HELVETICA-BOLD-OBLIQUE"
            fCode = "303"
            
        Case "Symbol"
            fCode = "400"
    End Select
        
    If fCode <> "" Then
        GetFontNameParameter = "-pfn" & fCode
    ElseIf Me.Font <> "" Then
        GetFontNameParameter = "-pfn:" & Me.Font
    Else
        GetFontNameParameter = "-pfn200"
    End If
    
End Function

Private Function GetFontSizeParameter() As String
    
    Dim font_size As Single
    
    font_size = Me.FontSize
    
    If Me.Orientation <> pdfPortrait Then
        If Me.FontSizeLandscape > 0 Then
            font_size = Me.FontSizeLandscape
        End If
    Else
        If Me.FontSizePortrait > 0 Then
            font_size = Me.FontSizePortrait
        End If
    End If
    
    If font_size > 0 Then
        GetFontSizeParameter = "-pfs" & CStr(font_size)
    Else
        GetFontSizeParameter = "-pfs9"
    End If
    
End Function

Private Function GetLinesPerPageParameter() As String
    Dim linesPerPage As Long
    
    linesPerPage = Me.RecordsPerPage
    
    If Me.Orientation <> pdfPortrait Then
        If Me.RecordsPerPageLandscape > 0 Then
            linesPerPage = Me.RecordsPerPageLandscape
        End If
    Else
        If Me.RecordsPerPagePortrait > 0 Then
            linesPerPage = Me.RecordsPerPagePortrait
        End If
    End If
    
    If linesPerPage > 0 Then
        GetLinesPerPageParameter = "-lpp" & CStr(linesPerPage)
    Else
        GetLinesPerPageParameter = "-lpp66"
    End If
    
End Function

Private Function GetOrientationParameter() As String
    
    If Me.Orientation <> pdfPortrait Then
        GetOrientationParameter = "-pot"
    End If
    
End Function

Private Function fnCheckSetOrientation() As String
    Dim objFile As FileServer
    Dim strRow As String
    Dim i As Long
    Dim m As Integer
    
    On Error GoTo errTrap
    
    If Not m_bAutoAdjustOrientation Then
        'okay, don't need to check
        Exit Function
    End If
    
    'note: the portrait page length must be > than the landscape page length + 5
    If Trim(m_strPageBreakBefore) = "" Or Me.FontSizeLandscape = Me.FontSizePortrait Or Me.RecordsPerPagePortrait <= Me.RecordsPerPageLandscape + 5 Then
        'okay, don't need to check
        Exit Function
    End If
    
    If Not isFileExist() Then
       fnCheckSetOrientation = m_strSourceFile & " File does not exist!"
       Exit Function
    End If
    
    Set objFile = New FileServer

    If Not objFile.PrepareFile(m_strSourceFile) Then
       fnCheckSetOrientation = "Failed to open " & m_strSourceFile & "!"
       Exit Function
    End If
    
    i = 0
    
    Do While Not objFile.EndOfFile
        strRow = objFile.GetLine
        i = i + 1
        
        '#Special code. Remove the first 3 empty lines
        If m_strPageBreakBefore <> "" And i = 1 Then
            m = 0
            Do While m <= 3
                If Trim(strRow) = "" Then
                    If Not objFile.EndOfFile Then
                        strRow = objFile.GetLine
                        i = i + 1
                    End If
                Else
                    Exit Do
                End If
                m = m + 1
            Loop
        End If
        
        If m_strPageBreakBefore <> "" Then
            If i > 10 Then
                If InStr(strRow, m_strPageBreakBefore) > 0 Then
                    i = i - 1
                    Exit Do
                End If
            End If
        End If
        
        If i >= Me.RecordsPerPagePortrait Then
            Exit Do
        End If
    Loop
    
    If i + 5 >= m_numLinesPerPagePortrait Then
        Me.Orientation = pdfPortrait
    Else
        Me.Orientation = pdfLandscape
    End If
    
    objFile.CloseFile
    fnCheckSetOrientation = ""
    Exit Function
    
errTrap:
    fnCheckSetOrientation = "An error has occurred while checking the report file. Err Code: " _
        & Err.Number & ", Err Desc: " & Err.Description
End Function

Private Function isFileExist() As Boolean
   isFileExist = IIf(Len(Dir(m_strSourceFile)) > 0, 1, 0)
End Function




