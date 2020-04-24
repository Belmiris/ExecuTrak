VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMultiFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_nSeq As Long
Private m_bQuiet As Boolean
Private m_fileLength As Long
Private m_lineLength As Long
'
Private Sub Class_Initialize()
    m_nSeq = 0
    m_bQuiet = False
    m_fileLength = 300
    m_lineLength = 1000
End Sub

Public Sub CreateTempTable(Optional fileLength As Long = 300, _
                           Optional lineLength As Long = 1000)
    On Error GoTo FINISHED
    Dim sSql As String
    Dim nCnt As Long
    
    m_fileLength = IIf(fileLength > 0, fileLength, 300)
    m_lineLength = IIf(lineLength > 0, lineLength, 1000)
    
    sSql = _
        "create temp table tmp_multi_file " & vbCrLf & _
        "( " & vbCrLf & _
        "    tmf_file       char(" & m_fileLength & "), " & vbCrLf & _
        "    tmf_seq        int, " & vbCrLf & _
        "    tmf_length     int, " & vbCrLf & _
        "    tmf_newline    char(1), " & vbCrLf & _
        "    tmf_contents   char(" & m_lineLength & ") " & vbCrLf & _
        ") " & vbCrLf
    nCnt = t_dbMainDatabase.ExecuteSQL(sSql)
    
    Err.Clear
FINISHED:
    If Err.Number <> 0 Then
        subHandleError "CreateTempTable", Err.Number, Erl(), Err.Description, sSql
    End If
End Sub

Public Property Get QuietMode() As Boolean
    QuietMode = m_bQuiet
End Property

Public Property Let QuietMode(ByVal vNewValue As Boolean)
    m_bQuiet = vNewValue
End Property

Public Function AddLine(sFile As String, sLine As String, Optional bWithCRLF As Boolean = True)
    On Error GoTo FINISHED
    Dim sSql As String
    Dim nLen As Long
    Dim sNewLine As String
    Dim nCnt As Long
    
    If Len(sFile) > m_fileLength Then
        Err.Raise -1234, "", "The file path '" & sFile & "' is too long for the field tmf_file "
    End If
    
    m_nSeq = m_nSeq + 1
    nLen = Len(sLine)
    If nLen > m_lineLength Then
        subHandleError "AddLine", 0, 0, "The line " & vbCrLf & sLine & vbCrLf & " in file " & vbCrLf & sFile & vbCrLf & "is too long for tmf_contents and will be truncated"
        nLen = m_lineLength
    End If
    
    If bWithCRLF Then sNewLine = "Y" Else sNewLine = "N"
    
    sSql = _
        "insert into tmp_multi_file " & vbCrLf & _
        "values " & vbCrLf & _
        "( " & vbCrLf & _
        "    " & tfnSQLString(UCase(sFile)) & ", " & vbCrLf & _
        "    " & m_nSeq & ", " & vbCrLf & _
        "    " & nLen & ", " & vbCrLf & _
        "    " & tfnSQLString(sNewLine) & ", " & vbCrLf & _
        "    " & tfnSQLString(sLine) & " " & vbCrLf & _
        ") " & vbCrLf
        
    nCnt = t_dbMainDatabase.ExecuteSQL(sSql)
    
    Err.Clear
FINISHED:
    If Err.Number <> 0 Then
        subHandleError "AddLine", Err.Number, Erl(), Err.Description, sSql
    End If
End Function

Public Function WriteFile(sFile As String, Optional bAppend = True) As Long
    On Error GoTo FINISHED
    Dim rsTemp As Recordset
    Dim sSql As String
    Dim hFile As Integer
    Dim nLen As Long
    Dim bWithCRLF As Boolean
    Dim sLine As String
    Dim nLines As Long
    
    sSql = _
        "select * " & vbCrLf & _
        "  from tmp_multi_file " & vbCrLf & _
        " where tmf_file = " & tfnSQLString(UCase(sFile)) & " " & vbCrLf & _
        " order by " & vbCrLf & _
        "       tmf_seq " & vbCrLf
    Set rsTemp = t_dbMainDatabase.OpenRecordset(sSql, dbOpenSnapshot, dbSQLPassThrough)
    
    If Not rsTemp Is Nothing Then
        If rsTemp.RecordCount > 0 Then
            ' OPEN FILE *************************
            hFile = FreeFile()
            If bAppend Then
                Open sFile For Append As #hFile
            Else
                Open sFile For Output As #hFile
            End If
            ' WRITE LINES ***********************
            While Not rsTemp.EOF
                nLen = tfnRound(rsTemp!tmf_length)
                bWithCRLF = UCase(Trim(rsTemp!tmf_newline)) <> "N"
                sLine = rsTemp!tmf_contents & ""
                If nLen > 0 Then
                    If Len(sLine) < nLen Then sLine = sLine & Space(nLen)
                    sLine = Left(sLine, nLen)
                Else
                    sLine = ""
                End If
                If bWithCRLF Then
                    Print #hFile, sLine
                ElseIf nLen > 0 Then
                    Print #hFile, sLine;
                End If
                
                rsTemp.MoveNext
            Wend
            ' CLOSE FILE ************************
            Close #hFile
            hFile = 0
        End If
    End If
    
    Err.Clear
FINISHED:
    If Err.Number <> 0 Then
        subHandleError "WriteFile", Err.Number, Erl(), Err.Description, sSql
    End If
    If hFile <> 0 Then
        On Error Resume Next
        Close #hFile
        hFile = 0
        Err.Clear
    End If
    CleanUp rsTemp
End Function

Public Function WriteAllFiles(Optional bAppend = True)
    On Error GoTo FINISHED
    Dim rsTemp As Recordset
    Dim sSql As String
    Dim sFile As String
    Dim colFiles As New Collection
    Dim hFile As Integer
    Dim nLen As Long
    Dim bWithCRLF As Boolean
    Dim sLine As String
    Dim nLines As Long
    Dim idx As Long
    
    sSql = _
        "select distinct tmf_file " & vbCrLf & _
        "  from tmp_multi_file " & vbCrLf & _
        " order by " & vbCrLf & _
        "       1 " & vbCrLf
    Set rsTemp = t_dbMainDatabase.OpenRecordset(sSql, dbOpenSnapshot, dbSQLPassThrough)
    If Not rsTemp Is Nothing Then
        If rsTemp.RecordCount > 0 Then
            While Not rsTemp.EOF
                nLen = tfnRound(rsTemp!tmf_file)
                sFile = Trim(rsTemp!tmf_file & "")
                colFiles.Add sFile
                rsTemp.MoveNext
            Wend
        End If
    End If
    
    For idx = 1 To colFiles.count
        sFile = colFiles(idx) & ""
        WriteFile sFile, bAppend
    Next
    
    Err.Clear
FINISHED:
    If Err.Number <> 0 Then
        subHandleError "WriteAllFiles", Err.Number, Erl(), Err.Description, sSql
    End If
    If hFile <> 0 Then
        On Error Resume Next
        Close #hFile
        hFile = 0
        Err.Clear
    End If
    CleanUp rsTemp
End Function

Public Function FileHasLines(sFile As String) As Boolean
    On Error GoTo FINISHED
    Dim sSql As String
    Dim rsTemp As Recordset
    Dim cnt As Long
    
    sSql = "select count(*) cnt from tmp_multi_file where tmf_file = " & tfnSQLString(UCase(sFile)) & " "
    Set rsTemp = t_dbMainDatabase.OpenRecordset(sSql, dbOpenSnapshot, dbSQLPassThrough)
    If Not rsTemp Is Nothing Then
        cnt = tfnRound(rsTemp!cnt & "")
        FileHasLines = cnt > 0
    End If
    
    Err.Clear
FINISHED:
    If Err.Number <> 0 Then
        subHandleError "FileHasLines", Err.Number, Erl(), Err.Description, sSql
    End If
    CleanUp rsTemp
End Function

Public Function GetFirstLine(sFile As String, Optional bSkipBlanks As Boolean = False) As String
    On Error GoTo FINISHED
    Dim sSql As String
    Dim rsTemp As Recordset
    Dim nLen As Long
    Dim bWithCRLF As Boolean
    Dim sLine As String
    
    sSql = _
        "select * " & vbCrLf & _
        "  from tmp_multi_file " & vbCrLf & _
        " where tmf_file = " & tfnSQLString(UCase(sFile)) & " " & vbCrLf & _
        " order by " & vbCrLf & _
        "       tmf_seq " & vbCrLf
    Set rsTemp = t_dbMainDatabase.OpenRecordset(sSql, dbOpenSnapshot, dbSQLPassThrough)
    If Not rsTemp Is Nothing Then
        While Not rsTemp.EOF
            nLen = tfnRound(rsTemp!tmf_length)
            bWithCRLF = UCase(Trim(rsTemp!tmf_newline)) <> "N"
            sLine = rsTemp!tmf_contents & ""
            If nLen > 0 Then
                If Len(sLine) < nLen Then sLine = sLine & Space(nLen)
                sLine = Left(sLine, nLen)
            Else
                sLine = ""
            End If
            
            If sLine <> "" Then
                GetFirstLine = sLine
                CleanUp rsTemp
                Exit Function
            ElseIf Not bSkipBlanks Then
                GetFirstLine = sLine
                CleanUp rsTemp
                Exit Function
            End If
            
            rsTemp.MoveNext
        Wend
    End If
    
    Err.Clear
FINISHED:
    If Err.Number <> 0 Then
        subHandleError "FileHasLines", Err.Number, Erl(), Err.Description, sSql
    End If
    CleanUp rsTemp
End Function

Public Sub DropTempTable()
    On Error Resume Next
    Dim sSql As String
    Dim nCnt As Long
    
    sSql = "drop table tmp_multi_file"
    nCnt = t_dbMainDatabase.ExecuteSQL(sSql)
    Err.Clear

End Sub

Private Sub subHandleError(sMethName As String, nErrNum As Long, nLine As Long, sErrMsg As String, Optional sSql As String = "")
    On Error GoTo FINISHED
    Dim sMsg As String
    Dim hFile As Integer
    
    sMsg = "Error in " & sMethName & " [" & nErrNum & "] Line " & nLine & ": " & sErrMsg
    If sSql <> "" Then sMsg = sMsg & vbCrLf & vbCrLf & sSql
    
    If Not m_bQuiet Then
        MsgBox sMsg
    Else
        hFile = FreeFile()
        Open io.StandardLogFile For Append As #hFile
        Print #hFile, sMsg
        Close #hFile
        hFile = 0
    End If
    
FINISHED:
    If hFile <> 0 Then
        On Error Resume Next
        Close #hFile
    End If
    Err.Clear
End Sub

' **************************************************************
' *** In case the program you add this to needs one of these ***
' **************************************************************
'
'Public Function fnTypInvoiceNoToString(rec As typInvoiceNo) As String
'    On Error GoTo FINISHED
'    Dim buf As String
'
'    buf = buf & Left(rec.sDatabase & Space(16), 16)
'    buf = buf & Left(rec.sCustCode & Space(16), 16)
'    buf = buf & Left(rec.sEmail & Space(65), 65)
'    buf = buf & Left(rec.sInvoiceNo & Space(300), 300)
'
'    fnTypInvoiceNoToString = buf
'
'    Err.Clear
'FINISHED:
'    If Err.Number <> 0 Then
'        MsgBox "Error in fnTypInvoiceNoToString: " & Err.Description
'        Err.Clear
'    End If
'End Function
'
'Public Function tfnSQLString(ByVal vTemp As Variant, _
'                             Optional bQuotes As Boolean = True, _
'                             Optional bConvertCRLFToCaret As Boolean = True) As String
'
'    Dim nIdx As Integer
'    Dim nPos As Integer
'    Dim szParameter As String
'
'    If IsNull(vTemp) Then
'        szParameter = ""
'    Else
'        szParameter = vTemp
'    End If
'
'    ' for each ' character found, insert a double ''.  Leave "%* alone
'    nIdx = 1
'    nPos = InStr(nIdx, szParameter, "'")
'
'    While nPos <> 0
'        szParameter = Left(szParameter, nPos) & "'" & Right(szParameter, Len(szParameter) - nPos)
'        nIdx = nPos + 2
'        nPos = InStr(nIdx, szParameter, "'")
'    Wend
'
'    If bConvertCRLFToCaret Then
'        'Robert Atwood 02-08-23 373550 modified function to replace all carets with CR's
'        'using tfnCRToCaret
'        szParameter = tfnCRToCaret(szParameter)
'    End If
'
'    ' quote the whole string - optional
'    If bQuotes Then
'        tfnSQLString = "'" & szParameter & "'"
'    Else
'        tfnSQLString = szParameter
'    End If
'End Function
'
'Public Function tfnCRToCaret(strInput As String) As String
'    Dim sTemp As String
'
'    sTemp = Replace(strInput, vbCrLf, "^")
'    sTemp = Replace(sTemp, vbCr, "")
'    tfnCRToCaret = Replace(sTemp, vbLf, "")
'
'End Function
'
'Public Function tfnGetNamedString(sSource As String, sName As String) As String
'    tfnGetNamedString = ""
'    If sSource = "" Or sName = "" Then
'        Exit Function
'    End If
'
'    Dim nPos1 As Integer
'    Dim nPos2 As Integer
'    Dim sUcaseSource As String
'    Dim sUcaseName As String
'
'    sUcaseSource = UCase(sSource)
'    sUcaseName = UCase(sName)
'    nPos1 = InStr(sUcaseSource, sUcaseName)
'
'    If nPos1 > 0 Then
'        nPos1 = InStr(nPos1, sUcaseSource, "=")
'        If nPos1 > 0 Then
'            nPos2 = InStr(nPos1, sUcaseSource, ";")
'            If nPos2 = 0 Then
'                nPos2 = Len(sUcaseSource) + 1
'            End If
'            nPos1 = nPos1 + 1
'            If nPos2 > nPos1 Then
'                tfnGetNamedString = Trim(Mid(sSource, nPos1, nPos2 - nPos1))
'            End If
'        End If
'    End If
'End Function
'
'Public Function tfnRound(vTemp As Variant, _
'                         Optional vPrec As Variant) As Variant
'    Dim fTempD As Double
'    Dim sFmt As String
'    Dim nPrec As Integer
'    Dim fOffset As Double
'    Dim sTemp As String
'
'    If IsMissing(vPrec) Then
'        nPrec = 0
'    Else
'        nPrec = vPrec
'    End If
'    If IsNull(vTemp) Then
'        tfnRound = 0
'    Else
'        If Trim(vTemp) = "" Then
'            tfnRound = 0
'        Else
'            If IsNumeric(vTemp) Then
'                If nPrec >= 0 Then
'                    sFmt = "#0." & String(nPrec, "#")
'                        sTemp = CStr(vTemp)
'                        tfnRound = Val(Format(sTemp, sFmt))
'                Else
'                    sTemp = CStr(vTemp)
'                    tfnRound = Val(Format(sTemp, "#"))
'                End If
'            Else
'                tfnRound = 0
'            End If
'        End If
'    End If
'End Function
'
'Public Sub CleanUp(ByRef inRS As Recordset)
'On Error GoTo ErrorCatch
'    inRS.Close
'    Set inRS = Nothing
'Exit Sub
'ErrorCatch:
'Err.Clear
'Resume Next
'End Sub

