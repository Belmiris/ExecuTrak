VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRetailStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***********************************************************'
'
' Copyright (c) 1996 FACTOR, A Division of W.R.Hess Company
'
' Module name   : RSSTORE.CLS
' Date          : December 13, 1996
' Programmer(s) : Qinggang Ma
'
' This module implements a class for tranferring rs_shifthold for Retail Storage.
'
'Usage:
'   1. Properties
'       a.  dbRemote (Set only)
'               Set the remote database object
'               This property must be set before call Load or Store
'       b.  dbLocal (Set only)
'               Set the local database object
'               This property must be set before call Load or Store
'       c.  ErrorCode (Let and Get)
'               Sets or Returns the error code associated with last action
'               ErrorCode 0 means no error
'       d.  ErrorDescription (Get only)
'           Returns the error description associated with last action
'
'   4. Functions
'       a.  Load    : Sub Load(ByVal lShiftLink As Long)
'               Loads records from remote database to local database
'       b.  Store   : Sub Store(ByVal lShiftLink As Long)
'               Stores records from local database to remote database
'       c.  SetDebug : Sub SetDebug(bStatus As Boolean)
'               Sets/remotes debug mode

Option Explicit
    Private Const ERROR_STATUS_OK = 0
    Private Const ERROR_INVALID_LOCAL_DB = 1
    Private Const ERROR_INVALID_REMOTE_DB = 2
    Private Const ERROR_LOAD_FAILED = 3
    Private Const ERROR_READ_DATABASE = 4
    Private Const ERROR_CLEANUP_REMOTE_FAILED = 5
    Private Const ERROR_CLEANUP_LOCAL_FAILED = 6
    Private Const ERROR_CALCULATE_LOAD_FAILED = 7
    Private Const ERROR_STORAGE_FAILED = 8
    Private Const ERROR_INVALID_SHIFTLINK = 9
    Private Const ERROR_COPY_TABLE = 10
    Private Const ERROR_TABLEDEF_DIFF = 11
    
    
    Private Const TABLE_LOCAL = "RSShiftHold"
    Private Const TABLE_REMOTE = "rs_shifthold"
    
    Private Type tpDBCopyTable
        m_dbObject As DataBase
        m_sTable As String
        m_sCriteria As String
'        m_sKeyFieldNames() As String
'        m_sKeyFieldValues() As String
    End Type

    Private m_dbLocal As DataBase
    Private m_dbRemote As DataBase
    
    Public m_nError As Integer
    
    Private m_bDebug As Boolean
Property Set dbLocal(dbTemp As Object)
    If dbTemp Is Nothing Then
        m_nError = ERROR_INVALID_LOCAL_DB
    Else
        Set m_dbLocal = dbTemp
        m_nError = ERROR_STATUS_OK
    End If
End Property

Property Set dbRemote(dbTemp As Object)
    If dbTemp Is Nothing Then
        m_nError = ERROR_INVALID_REMOTE_DB
    Else
        Set m_dbRemote = dbTemp
        m_nError = ERROR_STATUS_OK
    End If
End Property

Property Get ErrorCode() As Integer
    ErrorCode = m_nError
End Property

Property Let ErrorCode(nCode As Integer)
    m_nError = nCode
End Property

Property Get ErrorDescription() As String
    Select Case m_nError
        Case ERROR_STATUS_OK
            ErrorDescription = "No error"
        Case ERROR_INVALID_LOCAL_DB
            ErrorDescription = "Invalid local database object"
        Case ERROR_INVALID_REMOTE_DB
            ErrorDescription = "Invalid remote database object"
        Case ERROR_LOAD_FAILED
            ErrorDescription = "Load records from remote database failed"
        Case ERROR_READ_DATABASE
            ErrorDescription = "Can not read from database"
        Case ERROR_CLEANUP_REMOTE_FAILED
            ErrorDescription = "Clean up records from remote database failed"
        Case ERROR_CLEANUP_LOCAL_FAILED
            ErrorDescription = "Clean up records from local database failed"
        Case ERROR_STORAGE_FAILED
            ErrorDescription = "Store records to reomte database failed"
        Case ERROR_INVALID_SHIFTLINK
            ErrorDescription = "Invalid shift link"
        Case ERROR_TABLEDEF_DIFF
            ErrorDescription = "Definations of table are different in remote and local database"
        Case Else
            ErrorDescription = "Unknown error"
    End Select
End Property
Private Function fnCopyRecords(strctDest As tpDBCopyTable, _
                               strctSource As tpDBCopyTable, _
                               bClearSource As Boolean) As Boolean
                             
    'Copy some records from source(strctSource) to destionation(strctDest)
    '   The defination of the tables the source and destination
    '   must be the same.

    Dim strSQL As String
    Dim strSQLd As String
    Dim rsTemp As Recordset
    Dim rsTempd As Recordset
    Dim i As Integer
    Dim nFields As Integer
    
    fnCopyRecords = False

    strSQL = "SELECT * FROM " & strctSource.m_sTable & " WHERE " & strctSource.m_sCriteria
    Set rsTemp = fnOpenRecord(strSQL, strctSource.m_dbObject)
    If Not rsTemp Is Nothing Then
        If rsTemp.RecordCount > 0 Then
            strSQLd = "SELECT * FROM " & strctDest.m_sTable
            Set rsTempd = fnOpenRecord(strSQLd, strctDest.m_dbObject)
            If rsTempd Is Nothing Then
                Exit Function
            End If
            If rsTemp.Fields.Count <> rsTempd.Fields.Count Then
                m_nError = ERROR_TABLEDEF_DIFF
                Exit Function
            End If
            
            'Clean up destination table
            strSQL = "DELETE FROM " & strctDest.m_sTable & " WHERE " & strctDest.m_sCriteria
            If fnExecuteSQL(strSQL, strctDest.m_dbObject) < 0 Then
                m_nError = ERROR_CLEANUP_REMOTE_FAILED
                Exit Function
            End If

            nFields = rsTempd.Fields.Count - 1
            strSQLd = "INSERT INTO " & strctDest.m_sTable & "("
            For i = 0 To nFields
                If i < nFields Then
                    strSQLd = strSQLd & rsTempd.Fields(i).Name & ", "
                Else
                    strSQLd = strSQLd & rsTempd.Fields(i).Name
                End If
            Next
            rsTemp.MoveFirst
            While Not rsTemp.EOF
                strSQL = strSQLd & ") VALUES ("
                For i = 0 To nFields
                    If IsNull(rsTemp.Fields(i).value) Then
                        strSQL = strSQL & "''"
                    Else
                        strSQL = strSQL & fnSQLString(rsTemp.Fields(i).value)
                    End If
                    If i < nFields Then
                        strSQL = strSQL & ", "
                    End If
                Next
                strSQL = strSQL & ")"
                If fnExecuteSQL(strSQL, strctDest.m_dbObject) < 0 Then
                    m_nError = ERROR_COPY_TABLE
                    Exit Function
                End If
                rsTemp.MoveNext
            Wend
        End If
    End If

    If bClearSource Then
        'Clean up the source table
        strSQL = "DELETE FROM " & strctSource.m_sTable & " WHERE " & strctSource.m_sCriteria
        If fnExecuteSQL(strSQL, strctSource.m_dbObject) < 0 Then
            m_nError = ERROR_CLEANUP_LOCAL_FAILED
            Exit Function
        End If
    End If

    fnCopyRecords = True
    
End Function

Private Function fnFixSeqence(ByVal lShiftLink As Long) As Boolean
    Const SUB_NAME = "fnFixSeqence"
    
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim nCount As Integer
    Dim sLastType As String
    
    strSQL = "SELECT szType, lSeqNum FROM rsShiftHold" _
           & " WHERE lShiftLink = " & CStr(lShiftLink) _
           & " ORDER BY szType, lSeqNum"
    On Error GoTo errFix
    Set rsTemp = m_dbLocal.OpenRecordset(strSQL, dbOpenDynaset)
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveLast
        rsTemp.MoveFirst
        nCount = 1
        While Not rsTemp.EOF
            If sLastType <> fnCStr(rsTemp!szType) Then
                nCount = 1
                sLastType = fnCStr(rsTemp!szType)
            End If
            rsTemp.Edit
            rsTemp!lSeqNum = CStr(nCount)
            rsTemp.Update
            nCount = nCount + 1
            rsTemp.MoveNext
        Wend
    End If
    fnFixSeqence = True
    Exit Function
errFix:
    fnFixSeqence = False
    tfnErrHandler strSQL, SUB_NAME
End Function

Private Function fnCStr(vTemp As Variant) As String
    If IsNull(vTemp) Then
        fnCStr = ""
    Else
        fnCStr = Trim(vTemp)
    End If
End Function


Private Function fnSQLString(ByVal szParameter As String, Optional vNoQuotes As Variant) As String
    
    Dim nIdx As Integer
    Dim nPos As Integer
    
    nIdx = 1
    nPos = InStr(nIdx, szParameter, "'")
    
    While nPos <> 0
        szParameter = Left(szParameter, nPos) & "'" & Right(szParameter, Len(szParameter) - nPos)
        nIdx = nPos + 2
        nPos = InStr(nIdx, szParameter, "'")
    Wend
    
    ' quote the whole string - optional
    If IsMissing(vNoQuotes) Then
        fnSQLString = "'" & szParameter & "'"
    Else
        fnSQLString = szParameter
    End If

End Function

Private Function fnExecuteSQL(strSQL As String, _
                              objDB As DataBase, _
                              Optional sErrorMessage As Variant) As Integer
    
    Const SUB_NAME = "fnExecuteSQL"
    On Error GoTo errExecute
    If objDB Is m_dbRemote Then
        fnExecuteSQL = objDB.ExecuteSQL(strSQL)
    Else
        objDB.Execute strSQL
        fnExecuteSQL = 0
    End If
    If fnExecuteSQL < 0 Then
        fnExecuteSQL = 0
    End If
extExecute:
    On Error GoTo 0
    Exit Function

errExecute:
    tfnErrHandler strSQL, SUB_NAME
    Err.Clear
    fnExecuteSQL = -1
    Resume extExecute

End Function

Public Sub Load(ByVal lShiftLink As Long)

    If lShiftLink <= 0 Then
        m_nError = ERROR_INVALID_SHIFTLINK
        Exit Sub
    End If
    If m_nError = ERROR_INVALID_REMOTE_DB Or m_nError = ERROR_INVALID_LOCAL_DB Then
        Exit Sub
    End If
    
    Dim sLink As String
    Dim strctRemote As tpDBCopyTable
    Dim strctLocal As tpDBCopyTable
    
    sLink = CStr(lShiftLink)
    With strctRemote
        Set .m_dbObject = m_dbRemote
        .m_sTable = TABLE_REMOTE
        .m_sCriteria = "rssh_shl = " & sLink
    End With
    With strctLocal
        Set .m_dbObject = m_dbLocal
        .m_sTable = TABLE_LOCAL
        .m_sCriteria = "lShiftLink = " & sLink
    End With
    If Not fnCopyRecords(strctLocal, strctRemote, False) Then
        If m_nError = ERROR_COPY_TABLE Then
            m_nError = ERROR_LOAD_FAILED
        End If
    Else
        m_nError = ERROR_STATUS_OK
    End If
    
End Sub

Private Function fnOpenRecord(strSQL As String, _
                              objDB As DataBase, _
                              Optional vMsg As Variant) As Recordset
    Const SUB_NAME = "fnOpenRecord"
    ' Get records from the given SQL statement
    Dim rsTemp As Recordset

    On Error GoTo SQLError
    If objDB Is m_dbRemote Then
        Set rsTemp = objDB.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    Else
        Set rsTemp = objDB.OpenRecordset(strSQL, dbOpenSnapshot)
        If rsTemp.RecordCount > 0 Then
            rsTemp.MoveLast
            rsTemp.MoveFirst
        End If
    End If
    Set fnOpenRecord = rsTemp
quitsub:
    On Error GoTo 0
    Exit Function
SQLError:
    tfnErrHandler strSQL, SUB_NAME
    Set fnOpenRecord = Nothing
    Err.Clear
    Resume quitsub
End Function

Public Sub SetDebug(bStatus As Boolean)
    m_bDebug = bStatus
End Sub
Public Sub Store(ByVal lShiftLink As Long)
    If lShiftLink <= 0 Then
        m_nError = ERROR_INVALID_SHIFTLINK
        Exit Sub
    End If
    If m_nError = ERROR_INVALID_REMOTE_DB Or m_nError = ERROR_INVALID_LOCAL_DB Then
        Exit Sub
    End If
    
    Dim sLink As String
    Dim strctRemote As tpDBCopyTable
    Dim strctLocal As tpDBCopyTable
    
    If fnFixSeqence(lShiftLink) Then
        sLink = CStr(lShiftLink)
        With strctRemote
            Set .m_dbObject = m_dbRemote
            .m_sTable = TABLE_REMOTE
            .m_sCriteria = "rssh_shl = " & sLink
        End With
        With strctLocal
            Set .m_dbObject = m_dbLocal
            .m_sTable = TABLE_LOCAL
            .m_sCriteria = "lShiftLink = " & sLink
        End With
    
        If Not fnCopyRecords(strctRemote, strctLocal, False) Then
            If m_nError = ERROR_COPY_TABLE Then
                m_nError = ERROR_STORAGE_FAILED
            End If
        Else
            m_nError = ERROR_STATUS_OK
        End If
    Else
        m_nError = ERROR_STORAGE_FAILED
    End If
End Sub


