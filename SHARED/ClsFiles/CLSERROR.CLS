VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsErrorHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ODBC_CALL_FAIL As Integer = 3146
Private Const DEFINED_ERROR As Integer = 0
Private Const CHECK_EXEC_PROCEDURE As Integer = 1

Private Const REPORT_TO As String = "Report this message to Support."

'member variables
Private m_engFactor As DBEngine
Private m_frmParent As Form

Private m_sAppName As String
Private m_sProcName As String
Private m_nErrorCode As Variant
Private m_sErrorDesc As String
Private m_sOdbcErrorDesc As String
Private m_sSqlStatement As String
Private m_sOccurDateTime As String
Private m_UseDBEngin As Boolean

Private m_bAlwaysShowMessage As Boolean

Private m_sErrScrnPath As String
Private m_sLogFilePath As String
Private m_dbLocal As Database
Private m_nLogTo As Integer
Private m_bCirCular As Boolean
Private m_nMaxErrors As Integer
'

Private Sub Class_Initialize()
    Set m_engFactor = Nothing
    Set m_dbLocal = Nothing
    Set m_frmParent = Nothing
    
    m_sAppName = ""
    m_sProcName = ""
    m_nErrorCode = 0
    m_sErrorDesc = ""
    m_sOdbcErrorDesc = ""
    m_sSqlStatement = ""
    m_sOccurDateTime = ""
    m_UseDBEngin = True
    
    m_bAlwaysShowMessage = True

    m_sErrScrnPath = ""
    m_sLogFilePath = ""
    m_nLogTo = LOG_TO_FILE
    m_nMaxErrors = 10
    m_bCirCular = True
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    m_dbLocal.Close
    Set m_dbLocal = Nothing
    Set m_engFactor = Nothing
    Set m_frmParent = Nothing
End Sub

Public Function ErrorHandler(sProcName As String, _
                        Optional vSQL As Variant, _
                        Optional vExecProcMessage As Variant, _
                        Optional vShowMessage As Variant) As Boolean
    
    ErrorHandler = True
    
    Dim bShowMessage As Boolean, sErrorMsg As String, bDefinedError As Boolean
    Dim bExecProcError As Boolean
    
    If m_frmParent Is Nothing Then
        MsgBox "Form not set.  Error Handler function calls failed.", vbCritical
        Exit Function
    End If
    
    #If FACTOR_MENU <> 1 Then
    If m_UseDBEngin Then
        If m_engFactor Is Nothing Then
            MsgBox "DatabaseEngine not set.  Error Handler function calls failed.", vbCritical
            Exit Function
        End If
    End If
    #End If
'    #If DEVELOP Then
'        If m_dbLocal Is Nothing Then
'                MsgBox "Local Database not set.  Error Handler function calls failed.", vbCritical
'                Exit Function
'        End If
'    #End If
    
    sProcName = Trim(sProcName)
    If sProcName = "" Then sProcName = "unknown"
    m_sProcName = sProcName
    m_nErrorCode = Err.Number
    m_sErrorDesc = Err.Description
    
    If Not IsMissing(vExecProcMessage) Then
        If VarType(vExecProcMessage) = vbString Then
            If InStr(UCase(vExecProcMessage), "EXCEPTION") > 0 Then
                #If Not DEVELOP Then
                    'store the snapshot of the App screen
                    Set m_frmParent.Picture = CaptureForm(m_frmParent)
                    Clipboard.SetData m_frmParent.Picture, vbCFBitmap
                    SavePicture m_frmParent.Picture, m_sErrScrnPath
                    Set m_frmParent.Picture = Nothing
                #End If
            End If
        End If
    Else
        #If Not DEVELOP Then
            'store the snapshot of the App screen when vSQL does not contain "drop table"
            If Not IsMissing(vSQL) Then
                If InStr(LCase(vSQL), "drop table") = 0 Then
                    Set m_frmParent.Picture = CaptureForm(m_frmParent)
                    Clipboard.SetData m_frmParent.Picture, vbCFBitmap
                    SavePicture m_frmParent.Picture, m_sErrScrnPath
                    Set m_frmParent.Picture = Nothing
                End If
            Else
                Set m_frmParent.Picture = CaptureForm(m_frmParent)
                Clipboard.SetData m_frmParent.Picture, vbCFBitmap
                SavePicture m_frmParent.Picture, m_sErrScrnPath
                Set m_frmParent.Picture = Nothing
            End If
        #End If
    End If
    
    m_sOccurDateTime = Format$(Now)
    
    bDefinedError = False
    bExecProcError = False
    
    If IsMissing(vSQL) Then
        m_sSqlStatement = ""
        bShowMessage = m_bAlwaysShowMessage
    Else
        If VarType(vSQL) = vbBoolean Then
            m_sSqlStatement = ""
            bShowMessage = vSQL
        Else  'either a SQL string or an integer or long Error Code
            If VarType(vSQL) = vbInteger Or VarType(vSQL) = vbLong Then
                'override error code
                m_sErrorDesc = ""
                bDefinedError = True
                m_nErrorCode = vSQL
                bShowMessage = m_bAlwaysShowMessage
                If Not IsMissing(vExecProcMessage) Then
                    If VarType(vExecProcMessage) = vbString Then
                        m_sErrorDesc = vExecProcMessage
                        If Not IsMissing(vShowMessage) Then
                            If VarType(vShowMessage) = vbBoolean Then bShowMessage = vShowMessage
                        End If
                    ElseIf VarType(vExecProcMessage) = vbBoolean Then
                        bShowMessage = vShowMessage
                    End If
                End If
            ElseIf VarType(vSQL) = vbString Then 'SQL string
                If m_nErrorCode = ODBC_CALL_FAIL Or fnIsAccessDatabaseError(m_nErrorCode) Or VarType(vExecProcMessage) = vbString Then
                    m_sSqlStatement = vSQL
                Else
                    m_sSqlStatement = ""
                End If
                bShowMessage = m_bAlwaysShowMessage
                If Not IsMissing(vExecProcMessage) Then
                    If VarType(vExecProcMessage) = vbBoolean Then
                        bShowMessage = vExecProcMessage
                    Else  'Message from Execute Procedure
                        'check for exception in ExecProcMessage
                        If InStr(UCase(vExecProcMessage), "EXCEPTION") = 0 Then
                            ErrorHandler = False
                            Exit Function
                        End If
                        bExecProcError = True
                        m_sErrorDesc = vExecProcMessage
                        If Not IsMissing(vShowMessage) Then
                            If VarType(vShowMessage) = vbBoolean Then bShowMessage = vShowMessage
                        End If
                    End If
                End If
            Else
                MsgBox "Invalid Parameter.  ErrorHandler() calls failed.", vbCritical
                Exit Function
            End If
        End If
    End If
    
    
    'determine error
    If m_nErrorCode = ODBC_CALL_FAIL Then
        sErrorMsg = fnGetErrorMessage(ODBC_CALL_FAIL)
    ElseIf bDefinedError Then
        sErrorMsg = fnGetErrorMessage(DEFINED_ERROR)
    ElseIf bExecProcError Then
        sErrorMsg = fnGetErrorMessage(CHECK_EXEC_PROCEDURE)
    Else
        sErrorMsg = fnGetErrorMessage()
    End If
    
    If bShowMessage Then
        MsgBox sErrorMsg, vbOKOnly + vbCritical, m_sAppName
    Else
        #If DEVELOP Then
            MsgBox "THIS MESSAGE BOX WILL BE SHOWN ON DEVELOP MODE ONLY." & vbCrLf & vbCrLf & sErrorMsg, vbOKOnly + vbCritical, m_sAppName
        #End If
    End If
    
    Select Case m_nLogTo
        Case LOG_TO_FILE
            WriteToFile sErrorMsg
        Case LOG_TO_DB
            WriteToFile sErrorMsg
            If Not m_dbLocal Is Nothing Then
                #If Not DEVELOP Then
                    WriteToDB sErrorMsg
                #End If
            End If
    End Select
    Err.Clear

    Exit Function

ErrorTrap:
    'MsgBox "erro code2:" & Err.Number & " error desc2: " & Err.Description
    Resume Next
End Function

'valid nWhere is 0 - NO_LOG, 1 - LOG_TO_FILE, 2 - LOG_TO_DB
Public Property Let WriteLogTo(nWhere As Integer)
    m_nLogTo = nWhere
End Property

Public Property Let Circular(bOnOff As Boolean)
    m_bCirCular = bOnOff
End Property

Public Property Let MaxErrors(nMax As Integer)
    m_nMaxErrors = nMax
End Property

Public Property Set FormParent(frmForm As Object)
    Dim sVerInfo As String
    
    Set m_frmParent = frmForm
    m_sLogFilePath = tfnGetWindowsDir() & "\temp\errorlog.txt"
    m_sErrScrnPath = tfnGetWindowsDir() & "\temp\errscrn.bmp"
    sVerInfo = " (Version " & App.Major & "." & App.Minor & "." & App.Revision & ")"
    On Error GoTo ErrorSetName
    m_sAppName = App.EXEName & sVerInfo
    Exit Property
    
ErrorSetName:
    'MsgBox "Exe File not found, use form name instead", vbInformation
    m_sAppName = Trim(frmForm.Name) & sVerInfo
End Property

Public Property Set DatabaseEngine(dbEngTmp As DBEngine)
    Set m_engFactor = dbEngTmp
End Property

Public Property Set LocalDatabase(dbLocal As Database)
    Set m_dbLocal = dbLocal
End Property

Public Property Let LogFilePath(sPath As String)
    sPath = Trim(sPath)
    m_sLogFilePath = sPath & IIf(Right(sPath, 1) = "\", "", "\") & "ErrorLog.txt"
End Property

Private Function fnGetErrorMessage(Optional nErrorType As Variant) As String
    Dim sMsg As String, sTmp As String, nPosi As Integer, sProcName(9) As String
    Dim sTable As String, i As Integer, rsTemp As Recordset, sSql As String
    
    'Parse ProcNames if needed
    nPosi = 0
    sTmp = m_sProcName
    i = 0
    
    On Error Resume Next
    Do
        nPosi = InStr(sTmp, ",")
        If nPosi = 0 Then
            nPosi = InStr(sTmp, " ")
            If nPosi = 0 Then 'called from is missing or last parent function
                sProcName(i) = Trim(sTmp)
                Exit Do
            End If
        End If
        sProcName(i) = Left(sTmp, nPosi - 1)
        sTmp = Trim(Mid(sTmp, nPosi + 1))
        i = i + 1
        If i > 9 Then Exit Do
    Loop While nPosi > 0
    On Error GoTo 0
    
    sMsg = "An error has occurred in Module '" & m_sAppName & "', function '" & sProcName(0) & "'."
    If sProcName(1) <> "" Then
        sMsg = sMsg & "  Called From '" & sProcName(1) & "'." & vbCrLf
    End If
    
    If sProcName(2) <> "" Then
        sMsg = sMsg & "Parent function(s): " & sProcName(2)
        For i = 3 To UBound(sProcName)
            If sProcName(i) = "" Then Exit For
            If i = 6 Then
                sMsg = sMsg & ", " & vbCrLf & vbTab & vbTab & sProcName(i)
            Else
                sMsg = sMsg & ", " & sProcName(i)
            End If
        Next i
        sMsg = sMsg & "." & vbCrLf
    End If
    
    'take care defined error first
    If Not IsMissing(nErrorType) Then
        If nErrorType = DEFINED_ERROR Then
            If m_sErrorDesc = "" Then
                sTmp = "unknown error"
            Else
                sTmp = ""
            End If
            
            If Not m_dbLocal Is Nothing Then
                sSql = "SELECT * FROM ErrMessage WHERE ErrorCode = " & m_nErrorCode
                On Error Resume Next
                Set rsTemp = m_dbLocal.OpenRecordset(sSql, dbOpenSnapshot)
                sTmp = rsTemp!ErrorDesc
            End If
            
            m_sErrorDesc = sTmp & IIf(m_sErrorDesc <> "", " " & Trim(m_sErrorDesc), "")
            If Right(m_sErrorDesc, 1) <> "." Then m_sErrorDesc = m_sErrorDesc & "."
            sMsg = sMsg & vbCrLf & vbCrLf & "Error Code = " & m_nErrorCode & vbCrLf & "Error Desc: " & m_sErrorDesc & vbCrLf
            sMsg = sMsg & vbCrLf & REPORT_TO & vbCrLf
            fnGetErrorMessage = sMsg
            Exit Function
        End If
    End If
    
    'append SQL message
    If m_sSqlStatement <> "" Then
        On Error Resume Next
        sTable = "UNKNOWN"
        sTmp = UCase(Trim(m_sSqlStatement))
        sMsg = sMsg & vbCrLf & "Failed to "
        
        nPosi = 0
        
        Select Case True
            Case Left(sTmp, 11) = "INSERT INTO"
                sTmp = Trim(Mid(sTmp, 12))
                nPosi = InStr(sTmp, " ")
                If nPosi = 0 Then
                    nPosi = InStr(sTmp, "(")
                End If
                If nPosi > 0 Then
                    sTable = Trim(Left(sTmp, nPosi - 1))
                End If
                sMsg = sMsg & "insert data into the table '"
            Case Left(sTmp, 6) = "UPDATE"
                sTmp = Trim(Mid(sTmp, nPosi + 7))
                nPosi = InStr(sTmp, " ")
                If nPosi = 0 Then
                    nPosi = InStr(sTmp, "SET")
                End If
                If nPosi > 0 Then
                    sTable = Trim(Left(sTmp, nPosi - 1))
                End If
                sMsg = sMsg & "update data into the table '"
            Case Left(sTmp, 6) = "DELETE"
                nPosi = InStr(sTmp, "FROM")
                If nPosi > 0 Then
                    sTmp = Trim(Mid(sTmp, nPosi + 4))
                    nPosi = InStr(sTmp, " ")
                    If nPosi > 0 Then
                        sTable = Left(sTmp, nPosi - 1)
                    End If
                End If
                sMsg = sMsg & "delete data from the table '"
            Case Left(sTmp, 12) = "CREATE TABLE"
                sTmp = Trim(Mid(sTmp, 13))
                nPosi = InStr(sTmp, "(")
                If nPosi > 0 Then
                    sTable = Trim(Left(sTmp, nPosi - 1))
                End If
                sMsg = sMsg & "create the table '"
            Case Left(sTmp, 17) = "CREATE TEMP TABLE"
                sTmp = Trim(Mid(sTmp, 18))
                nPosi = InStr(sTmp, "(")
                If nPosi > 0 Then
                    sTable = Trim(Left(sTmp, nPosi - 1))
                End If
                sMsg = sMsg & "create the temp table '"
            Case InStr(sTmp, "INTO TEMP") > 0
                nPosi = InStr(sTmp, "INTO TEMP")
                If nPosi > 0 Then
                    sTable = Trim(Mid(sTmp, nPosi + 9))
                End If
                    sMsg = sMsg & "insert data into the temp table '"
            Case Left(sTmp, 11) = "CREATE PROC"
                sTmp = Trim(Mid(sTmp, 17))
                nPosi = InStr(sTmp, "(")
                If nPosi > 0 Then
                    sTable = Trim(Left(sTmp, nPosi - 1))
                End If
                sMsg = sMsg & "create the stored procedure '"
            Case Left(sTmp, 8) = "CREATE I" Or Left(sTmp, 8) = "CREATE U"
                nPosi = InStr(sTmp, " ON")
                If nPosi > 0 Then
                    sTable = Trim(Mid(sTmp, IIf(Left(sTmp, 8) = "CREATE I", 13, 20), nPosi - IIf(Left(sTmp, 8) = "CREATE I", 13, 20)))
                End If
                sMsg = sMsg & "create the index '"
            Case Left(sTmp, 10) = "DROP TABLE"
                sTable = Trim(Mid(sTmp, 11))
                sMsg = sMsg & "drop the table '"
            Case Left(sTmp, 9) = "DROP PROC"
                sTable = Trim(Mid(sTmp, 15))
                sMsg = sMsg & "drop the stored procedure '"
            Case Left(sTmp, 10) = "DROP INDEX"
                sTable = Trim(Mid(sTmp, 11))
                sMsg = sMsg & "drop the index '"
            Case Left(sTmp, 6) = "SELECT"
                nPosi = InStr(sTmp, "FROM")
                If nPosi > 0 Then
                    sTmp = Trim(Mid(sTmp, nPosi + 4))
                    nPosi = InStr(sTmp, "WHERE")
                    If nPosi > 0 Then
                        sTable = Trim(Left(sTmp, nPosi - 1))
                    Else
                        sTable = sTmp
                    End If
                End If
                sMsg = sMsg & "read data from the table '"
            Case Left(sTmp, 17) = "EXECUTE PROCEDURE"
                sTmp = Trim(Mid(sTmp, nPosi + 18))
                nPosi = InStr(sTmp, " ")
                If nPosi = 0 Then
                    nPosi = InStr(sTmp, "(")
                End If
                If nPosi > 0 Then
                    sTable = Trim(Left(sTmp, nPosi - 1))
                End If
                sMsg = sMsg & "execute the stored procedure '"
            Case Else  'keyword is not valid
                sTable = "'Keyword is not valid"
                sMsg = sMsg & "run the SQL statement, "
        End Select
        
        sMsg = sMsg & sTable & "'." & vbCrLf & vbCrLf & "SQL statement: " & m_sSqlStatement & "." & vbCrLf
        
        If Not IsMissing(nErrorType) Then
            If nErrorType = CHECK_EXEC_PROCEDURE Then
                sMsg = sMsg & vbCrLf & "Error Code = 0" & vbCrLf & "Error Desc: " & m_sErrorDesc & vbCrLf
                fnGetErrorMessage = sMsg
                Exit Function
            End If
        End If
        On Error GoTo 0
    End If  'm_sSqlStatement <> ""
    
    'append odbcErrorDesc
    If IsMissing(nErrorType) Then  'vb run-time error
        sMsg = sMsg & vbCrLf & vbCrLf & "Error Code = " & m_nErrorCode & vbCrLf & "Error Desc: " & m_sErrorDesc & vbCrLf
    Else
        If nErrorType = ODBC_CALL_FAIL Then
            m_sOdbcErrorDesc = fnGetOdbcErrorDesc()
            sMsg = sMsg & vbCrLf & "Error Code = " & m_nErrorCode & vbCrLf & "Error Desc: " & m_sOdbcErrorDesc & vbCrLf
        End If
    End If
    
    sMsg = sMsg & vbCrLf & REPORT_TO & vbCrLf
    
    fnGetErrorMessage = sMsg
End Function

Public Function fnGetOdbcErrorDesc() As String
    Dim n As Integer
    
    If Not m_UseDBEngin Then
        Exit Function
    End If
    
    fnGetOdbcErrorDesc = ""
    For n = 0 To m_engFactor.Errors.Count - 1
        If n = 0 Then
            fnGetOdbcErrorDesc = m_engFactor.Errors(n)
        Else
            fnGetOdbcErrorDesc = fnGetOdbcErrorDesc & vbCrLf & m_engFactor.Errors(n)
        End If
    Next
End Function

Private Function WriteToFile(sMsg As String) As Boolean
    Dim nFilePtr As Integer, aryErrors() As String, nTotalLines As Integer
    Dim sErrLine As String, nErrCount As Integer, i As Integer, nEndOfFirstError As Integer
    
    On Error GoTo ErrorFatal
    ReDim aryErrors(500)
    On Error GoTo 0
    
    sMsg = sMsg & vbCrLf & "Occurred at: " & m_sOccurDateTime & vbCrLf
    
    'delete oldest log
    
    nFilePtr = FreeFile
    On Error Resume Next
    Open m_sLogFilePath For Input As #nFilePtr
    If Err.Number = 53 Then
        Open m_sLogFilePath For Output As #nFilePtr
    ElseIf Err.Number = 0 Then
        nErrCount = 0
        nTotalLines = 0
        While Not EOF(nFilePtr)
            sErrLine = ""
            Line Input #nFilePtr, sErrLine
            On Error GoTo ErrorSubscriptOutOfRange
            aryErrors(nTotalLines) = sErrLine
    '        On Error Resume Next
            If InStr(sErrLine, "Occurred at:") > 0 Then
                nErrCount = nErrCount + 1
                If nErrCount = 1 Then nEndOfFirstError = nTotalLines + 2
            End If
            nTotalLines = nTotalLines + 1
        Wend
    End If
    Close #nFilePtr
        
    If nErrCount >= m_nMaxErrors Then
        nFilePtr = FreeFile
        On Error GoTo ErrorOpenFile
        Open m_sLogFilePath For Output As #nFilePtr
        For i = nEndOfFirstError To nTotalLines
            Print #nFilePtr, aryErrors(i)
        Next i
        Close #nFilePtr
    End If
    
    nFilePtr = FreeFile
    On Error GoTo ErrorOpenFile
    Open m_sLogFilePath For Append Access Write As #nFilePtr
    Print #nFilePtr, sMsg
    Close #nFilePtr

    Exit Function

ErrorOpenFile:
'    MsgBox "Failed to open Error Log File: " & m_sLogFilePath, vbCritical
    Exit Function

ErrorSubscriptOutOfRange:
    On Error GoTo ErrorFatal
    ReDim Preserve aryErrors(UBound(aryErrors) + 100)
    Resume
    
ErrorFatal:
'    MsgBox "Out of memory in function WriteToFile.", vbCritical
End Function

Private Function WriteToDB(sMsg As String) As Boolean
    Dim rsTemp As Recordset, sErrMsg As String, i As Integer, sSql As String
    Dim iFileNumber As Integer, fdObject As Field
    
    If m_dbLocal Is Nothing Then
        Exit Function
    End If
    
    On Error GoTo ErrorAccessErrorLog
    
    Set rsTemp = m_dbLocal.OpenRecordset("select * from ErrorLog order by OccurDateTime asc")
    
    If rsTemp.EOF Then
        rsTemp.AddNew
    Else
        rsTemp.MoveLast
        rsTemp.MoveFirst
        
        If rsTemp.RecordCount >= m_nMaxErrors Then
            If m_bCirCular Then
                rsTemp.Delete
                rsTemp.MoveLast
            Else
                For i = 1 To rsTemp.RecordCount
                    rsTemp.Delete
                    rsTemp.MoveNext
                Next i
            End If
        End If
        rsTemp.AddNew
    End If
   
    rsTemp!AppName = m_sAppName
    rsTemp!ProcName = m_sProcName
    rsTemp!ErrorCode = Str$(m_nErrorCode)
    If Len(m_sErrorDesc) > 255 Then
        rsTemp!ErrorDesc = Left(m_sErrorDesc, 255)
    Else
        rsTemp!ErrorDesc = m_sErrorDesc
    End If
    rsTemp!OdbcErrorDesc = m_sOdbcErrorDesc
    rsTemp!SqlStatement = m_sSqlStatement
    rsTemp!OccurDateTime = m_sOccurDateTime
    
    'store the bitmap when error message does not contain "drop the table"
    If InStr(m_sErrorDesc, "drop the table") < 0 Then
        iFileNumber = FreeFile
        Open m_sErrScrnPath For Binary As iFileNumber 'Create temp
        Set fdObject = rsTemp("ScreenImage")
        FileStreamToField iFileNumber, fdObject 'Put data
        Close iFileNumber                'Close temp file
    End If
    
    rsTemp.Update
    rsTemp.Close
    Set rsTemp = Nothing
    
    Exit Function

ErrorAccessErrorLog:
    If Err.Number = 3078 Then 'Couldn't find input table or query
        sSql = "CREATE TABLE ErrorLog (AppName char(30), " & _
                                      "ProcName char(255), " & _
                                      "ErrorCode char(10), " & _
                                      "ErrorDesc char(255), " & _
                                      "OdbcErrorDesc LONGTEXT, " & _
                                      "SqlStatement LONGTEXT, " & _
                                      "ScreenImage LONGBINARY, " & _
                                      "OccurDateTime char(30))"
        On Error GoTo ErrorDatabase
        m_dbLocal.Execute sSql
        Resume
    End If
    
ErrorDatabase:
    sErrMsg = "An error has occurred in function WriteToDB." & vbCrLf & vbCrLf
    sErrMsg = sErrMsg & "Failed to access Error Log File."
    sErrMsg = sErrMsg & "Error Log will be written into " & m_sLogFilePath
'    MsgBox sErrMsg, vbCritical
    WriteToFile sMsg
End Function

Private Function fnIsAccessDatabaseError(lErrorCode As Variant) As Boolean
    If lErrorCode >= 2420 And lErrorCode <= 3622 Then
        fnIsAccessDatabaseError = True
    Else
        fnIsAccessDatabaseError = False
    End If
End Function

Property Let UseFactorDBEngin(bFlag As Boolean)
    m_UseDBEngin = bFlag
End Property
