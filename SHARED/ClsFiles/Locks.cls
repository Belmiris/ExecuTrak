VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRecordLocks"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'***********************************************************'
'
' Copyright (c) 1996 FACTOR, A Division of W.R.Hess Company
'
' Module name   : LOCKS.CLS
' Date          : October 15, 1996
' Programmer    : Qinggang Ma
'
' This module implements a class for keeping database consistancy
' whene editing an existing record.
'
' Functions:
'
'   1. CreateLock(sTable As String, sKeyFields As String)
'       Create an environment to get and update a record
'       Parameters: Table name and key fields
'
'   2. OpenRecordset(strSQL As String) as Recordset
'        Return a recordset for the given SQL. The record will also be put
'           in a temp table for comparing when update the table
'       Parameter: A SQL statement
'
'   3. UpdateRecords(strSQL As String) As Integer
'       Retrun the number of records affected by the SQL statement
'       Parameter: A SQL statement or the set value part of the SQL.

Option Explicit
    Private Const MAX_NAME_LEN = 11
    Private Const NAME_SUFFIX = "_hold"
    Private Const FIELD_TYPE_SHORT = 3
    Private Const FIELD_TYPE_INTEGER = 4
    Private Const FIELD_TYPE_SINGLE = 6
    Private Const FIELD_TYPE_DECIMAL10 = 7
    Private Const FIELD_TYPE_DATE = 8
    Private Const FIELD_TYPE_CHAR = 10
    
    Private Type tpFields
        m_sName As String
        m_sValue As String
    End Type

    Private CRLF As String * 2
    Private m_dbRemote As Object
    Private m_sTableName As String
    Private m_sKeyFields() As String
    Private m_sWhereClause As String
    
Public Sub CreateLock(sTable As String, sKeyFields As String)

    Dim i As Integer
    m_sTableName = Trim(sTable)
    subParseString m_sKeyFields, sKeyFields
    
    For i = 0 To UBound(m_sKeyFields)
        m_sKeyFields(i) = LCase(Trim(m_sKeyFields(i)))
    Next
End Sub

Private Function fnTempTableName() As String
    
    Dim nLen As Integer
    
    nLen = Len(m_sTableName)
    If nLen > MAX_NAME_LEN Then
        fnTempTableName = Left(m_sTableName, MAX_NAME_LEN) & NAME_SUFFIX
    Else
        fnTempTableName = Left(m_sTableName, nLen) & NAME_SUFFIX
    End If

End Function

Public Function OpenRecordset(sSQL As String) As Recordset
    ' Get records from the given SQL statement
    Const sMsg = "Can not create table for locking"
    
    Dim sTable As String
    Dim nLen As Integer
    Dim rsTemp As Recordset
    Dim strSQL As String
    Dim nPos As Integer

    Set OpenRecordset = Nothing
    
    'Get the where clause
    strSQL = UCase(sSQL)
    m_sWhereClause = "WHERE "
    nPos = InStr(strSQL, " WHERE ")
    If nPos = 0 Then
        Exit Function
    End If
    If nPos <= 1 Then
        #If DEVELOP Then
            MsgBox "Syntax error in SQL:" & CRLF & sSQL, vbCritical, App.Title
        #End If
        Exit Function
    End If
    m_sWhereClause = " " & Right(sSQL, Len(sSQL) - nPos + 1)

    sTable = fnTempTableName

    fnExecuteSQL "DROP TABLE " & sTable
    If Not fnExecuteSQL(sSQL & " INTO TEMP " & sTable, sMsg) Then
        Exit Function
    End If
    
    If fnGetRecord(rsTemp, strSQL) Then
        #If DEVELOP Then
            If rsTemp.RecordCount > 1 Then
                MsgBox "Multiple records selected. This class may not handle correctly", vbExclamation, App.Title
            End If
        #End If
    End If
    
    Set OpenRecordset = rsTemp
    
End Function

Private Function fnExecuteSQL(strSQL As String, Optional sErrorMessage As Variant) As Boolean

    Dim sMsg As String
    Dim bMsg As Boolean

    On Error GoTo errExecute
    m_dbRemote.Execute strSQL, dbSQLPassThrough
    fnExecuteSQL = True

extExecute:
    On Error GoTo 0
    Exit Function

errExecute:
    If IsMissing(sErrorMessage) Then
        sMsg = "Error occurred while executing a SQL statement"
        bMsg = False
    Else
        sMsg = sErrorMessage
        bMsg = True
    End If
    sMsg = sMsg & CRLF & CRLF & "Error # " & Err.Number & CRLF & Err.Description
    #If DEVELOP Then
        sMsg = sMsg & CRLF & CRLF & strSQL
        Clipboard.SetText strSQL
        MsgBox sMsg, vbOKOnly + vbCritical, App.Title
    #Else
        If bMsg Then
            MsgBox sMsg, vbOKOnly + vbCritical, App.Title
        End If
    #End If
    Err.Clear
    fnExecuteSQL = False
    Resume extExecute

End Function


Private Sub subParseString(sParam() As String, _
                          sSrc As String, _
                          Optional vStart As Variant, _
                          Optional vEnd As Variant)
    Dim i1 As Integer
    Dim i2 As Integer
    Dim i3 As Integer
    Dim k As Integer
    Dim nEND As Integer
    If IsMissing(vStart) Then
        i1 = 1
    Else
        i1 = vStart
    End If
    If IsMissing(vEnd) Then
        nEND = Len(sSrc)
    Else
        nEND = vEnd
    End If
    If i1 < 1 Then i1 = 1
    i2 = 1
    k = 0
    ReDim sParam(10)
    While i1 < nEND And i2 > 0 And i2 < nEND
        i2 = InStr(i1, sSrc, ",")
        If i2 = i1 Then
            i1 = i1 + 1
        Else
            If i2 > i1 And i2 < nEND Then
                sParam(k) = Trim$(Mid$(sSrc, i1, i2 - i1))
                k = k + 1
                i1 = i2 + 1
                i3 = UBound(sParam)
                If k > i3 Then
                    ReDim Preserve sParam(i2 + 10)
                End If
            End If
        End If
    Wend
    If i1 <= nEND Then
        sParam(k) = Trim$(Mid$(sSrc, i1, nEND - i1 + 1))
        k = k + 1
    Else
        sParam(k - 1) = Trim$(Mid$(sSrc, i1, nEND - i1 + 1))
    End If
    If k > 0 Then
        ReDim Preserve sParam(k - 1)
    End If
End Sub


Property Set DataBase(objDB As Object)
    Set m_dbRemote = objDB
End Property


Private Sub subGetNameValues(sFields() As tpFields, sArray() As String)

    Dim i As Integer
    Dim nPos As Integer
    Dim nLen As Integer
    Dim sTemp As String
    Dim sChar As String * 1

    ReDim sFields(UBound(sArray))
    For i = 0 To UBound(sArray)
        nPos = InStr(sArray(i), "=")
        If nPos > 1 Then
            sFields(i).m_sName = LCase(Trim(Left(sArray(i), nPos - 1)))
            nLen = Len(sArray(i)) - nPos - 1
            If nLen > 0 Then
                sTemp = Trim(Right(sArray(i), nLen))
                sChar = Left(sTemp, 1)
                If sChar = "'" Or sChar = Chr(34) Then
                    If nLen > 1 Then
                        nLen = nLen - 1
                        sTemp = Right(sTemp, nLen)
                    Else
                        sTemp = ""
                    End If
                End If
                sChar = Right(sTemp, 1)
                If sChar = "'" Or sChar = Chr(34) Then
                    If nLen > 1 Then
                        sTemp = Left(sTemp, nLen - 1)
                    Else
                        sTemp = ""
                    End If
                End If
                sFields(i).m_sValue = Trim(sTemp)
            End If
        End If
    Next

End Sub

Public Function UpdateRecords_Old(sSQL As String) As Integer

    Dim sSet As String
    Dim sWhere As String
    Dim strSQL As String
    Dim nPos As Integer
    Dim sTempTable As String
    Dim sArray() As String
    Dim sFields() As String
    Dim nKeys As Integer
    Dim i As Integer
    Dim j As Integer
    Dim rsTemp As Recordset
    Dim bFound As Boolean
    
    UpdateRecords = False
    strSQL = UCase(sSQL)
    nPos = InStr(strSQL, " WHERE ")
    If nPos = 0 Then
        #If DEVELOP Then
            MsgBox "No where clause in the update SQL statement:" & CRLF & sSQL, vbCritical, App.Title
        #End If
        Exit Function
    End If
    If nPos <= 1 Then
        #If DEVELOP Then
            MsgBox "Syntax error in SQL:" & CRLF & sSQL, vbCritical, App.Title
        #End If
    End If
    sSet = Left(sSQL, nPos - 1)
    sWhere = Right(sSQL, Len(sSQL) - nPos + 1)
    
    subParseString sArray, sSet
    subGetNameValues sSetFields, sArray
    
    sTempTable = fnTempTableName
    
    strSQL = "SELECT * FROM " & sTempTable & " " & sWhere
    If Not fnGetRecord(rsTemp, strSQL) Then
        Exit Function
    End If
    If rsTemp.RecordCount = 0 Then
        #If DEVELOP Then
            MsgBox "The records are not selected for editing"
        #End If
        Exit Function
    End If

    ReDim sFields(rsTemp.Fields.Count - 1)
    For i = 0 To rsTemp.Fields.Count - 1
        sFields(i) = LCase(Trim(rsTemp.Fields(i).Name))
    Next
        
    nKeys = UBound(m_sKeyFields)
    For i = 0 To nKeys
        sWhere = sWhere & " AND " & m_sTableName & "." & m_sKeyFields(i) & " = " & sTempTable & "." & m_sKeyFields(i)
    Next i
    
    For i = 0 To UBound(sArray)
        sArray(i) = LCase(Trim(sArray(i)))
        bFound = False
        For j = 0 To rsTemp.Fields.Count - 1
            If sArray(i) = sFields(j) Then
                bFound = True
                Exit For
            End If
        Next j
        If bFound Then
            For j = 0 To nKeys
                If sArray(i) = m_sKeyFields(j) Then
                    bFound = False
                    Exit For
                End If
            Next j
        End If
        If bFound Then
            sWhere = sWhere & " AND " & m_sTableName & "." & sArray(i) & " = " & sTempTable & "." & sArray(i)
        End If
    Next i
    
    strSQL = "UPDATE " & m_sTableName & " SET " & sSet & sWhere

    On Error GoTo errUpdate
    nPos = m_dbRemote.ExecuteSQL(strSQL)
extUpdate:
    UpdateRecords = nPos
    Exit Function
errUpdate:
    nPos = 0
End Function

Public Function UpdateRecords(sSQL As String) As Integer
    'Update the record(s) that has(have) been selected in the OpenRecordset function
    Dim sSet As String
    Dim sWhere As String
    Dim strSQL As String
    Dim sACriteria As String
    Dim nPos As Integer
    Dim sTempTable As String
    Dim sArray() As String
    Dim sFields() As tpFields
    Dim sSetFields() As tpFields
    Dim nKeys As Integer
    Dim i As Integer
    Dim j As Integer
    Dim rsTemp As Recordset
    
    'Take out the where clause
    UpdateRecords = False
    strSQL = UCase(sSQL)
    nPos = InStr(strSQL, " WHERE ")
    If nPos > 1 Then
        sSet = Left(sSQL, nPos - 1)
    Else
        sSet = sSQL
    End If

    'Take out the UPDATE .. SET key words
    nPos = InStr(strSQL, " SET ")
    If nPos > 0 Then
        sSet = Right(sSet, Len(sSet) - nPos - 4)
    End If
    
    'Parse the set value part of the SQL statement
    subParseString sArray, sSet
    'Parse the field name and value
    subGetNameValues sSetFields, sArray
    
    sTempTable = fnTempTableName
    
    'Get the record from the temptable
    strSQL = "SELECT * FROM " & sTempTable & " " & m_sWhereClause
    If Not fnGetRecord(rsTemp, strSQL) Then
        Exit Function
    End If
    If rsTemp.RecordCount = 0 Then
        #If DEVELOP Then
            MsgBox "The records are not selected for editing"
        #End If
        Exit Function
    End If

    'Get the fields and values
    ReDim sFields(rsTemp.Fields.Count - 1)
    For i = 0 To rsTemp.Fields.Count - 1
        sFields(i).m_sName = LCase(Trim(rsTemp.Fields(i).Name))
        If IsNull(rsTemp.Fields(i).Value) Then
            sFields(i).m_sValue = ""
        Else
            sFields(i).m_sValue = Trim(rsTemp.Fields(i).Value)
        End If
    Next

    'Build the SET clause and WHERE clause
    sWhere = ""
    nKeys = UBound(m_sKeyFields)
    sSet = ""
    For i = 0 To UBound(sSetFields)
        nPos = -1
        'Check whether the field is changed
        For j = 0 To rsTemp.Fields.Count - 1
            If sSetFields(i).m_sName = sFields(j).m_sName Then
                If sSetFields(i).m_sValue <> sFields(j).m_sValue Then
                    nPos = j
                    Exit For
                End If
            End If
        Next j
        
        'Check whether the field is key field
        If nPos >= 0 Then
            For j = 0 To nKeys
                If sSetFields(i).m_sName = m_sKeyFields(j) Then
                    nPos = -1
                    Exit For
                End If
            Next j
        End If
        
        'If the field has been changed, add to the SET and WHERE clauses
        If nPos >= 0 Then
            If sSet = "" Then
                sSet = sSetFields(nPos).m_sName & " = " & tfnSQLString(sSetFields(nPos).m_sValue)
            Else
                sSet = sSet & ", " & sSetFields(nPos).m_sName & " = " & tfnSQLString(sSetFields(nPos).m_sValue)
            End If
            If IsNull(rsTemp.Fields(nPos).Value) Then
                sWhere = sWhere & " AND " & sFields(i).m_sName & " = NULL"
            Else
                If rsTemp.Fields(nPos).Type = FIELD_TYPE_CHAR Or rsTemp.Fields(nPos).Type = FIELD_TYPE_DATE Then
                    sWhere = sWhere & " AND " & sFields(nPos).m_sName & " = " & tfnSQLString(sFields(nPos).m_sValue)
                Else
                    sWhere = sWhere & " AND " & sFields(nPos).m_sName & " = " & sFields(nPos).m_sValue
                End If
            End If
        End If
    Next i
    
    'If no fields are changed, set the affected row numbers
    If sWhere = "" Then
        strSQL = "SELECT " & m_sKeyFields(0) & " FROM " & m_sTableName & m_sWhereClause
        If fnGetRecord(rsTemp, strSQL) Then
            UpdateRecords = rsTemp.RecordCount
        End If
        Exit Function
    End If

    'Make the SQL
    strSQL = "UPDATE " & m_sTableName & " SET " & sSet & m_sWhereClause & sWhere

'    On Error GoTo errUpdate
    nPos = m_dbRemote.ExecuteSQL(strSQL)
extUpdate:
    UpdateRecords = nPos
    Exit Function
errUpdate:
    nPos = 0

End Function

Public Function fnGetRecord(rsTemp As Recordset, _
                            strSQL As String) As Boolean
    ' Get records from the given SQL statement

    On Error GoTo SQLError
    Set rsTemp = m_dbRemote.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    
    fnGetRecord = True

quitsub:
    On Error GoTo 0
    Exit Function
SQLError:
    Dim sMsg As String
    sMsg = "An error occurred while doing a SQL query" & CRLF & CRLF & "Error# " & CStr(Err.Number) & CRLF & Err.Description
    #If DEVELOP Then
        sMsg = sMsg & CRLF & CRLF & strSQL
        Clipboard.SetText strSQL
    #End If
    MsgBox sMsg, vbOKOnly + vbInformation, App.Title
    fnGetRecord = False
    Err.Clear
    Resume quitsub
End Function
Private Sub Class_Initialize()
    CRLF = Chr(&HD) & Chr(&HA)
End Sub


