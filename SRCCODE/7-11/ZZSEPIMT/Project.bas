Attribute VB_Name = "modProject"
Option Explicit

Public objCursor As clsCursor
Public cValidate As cValidateInput

Public Const sDATABASE_FAILED As String = "Failed to access remote database"
Public Const nDBRemote As Integer = 1
Public Const nDBLocal As Integer = 2

Public dbLocal As Database

' Get records from the given SQL statement
' nDB = 1 ---> Informax Database (remote)
'     = 2 ---> Access Database (local)
'This function will return a recordcount
Public Function fnGetRecord(rsTemp As Recordset, strSQL As String, nDB As Integer, sCalledFrom As String, Optional bShowError As Variant) As Long
    Const SUB_NAME = "fnGetRecord"
    
    On Error GoTo SQLError
    Select Case nDB
        Case 1  'remote
            Set rsTemp = t_dbMainDatabase.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
        Case 2  'local
            Set rsTemp = dbLocal.OpenRecordset(strSQL, dbOpenSnapshot)
    End Select
    
    If rsTemp.RecordCount > 0 Then  'the following code is used to get the correct
        rsTemp.MoveLast             'RecordCount of the RecordSet
        rsTemp.MoveFirst
    End If
    
    fnGetRecord = rsTemp.RecordCount
    
    Exit Function

SQLError:
    tfnErrHandler SUB_NAME + "," + sCalledFrom, strSQL, bShowError
    fnGetRecord = -9999
End Function

Public Function fnCreateTempTable(sTmpTableName As String, sSelectSql As String, sCalledFrom As String) As Boolean
    Const SUB_NAME = "fnCreateTempTable"
    Dim sSql As String
    
    fnDropTempTable sTmpTableName
    If UCase(Left(sSelectSql, 6)) = "SELECT" Then
        sSql = sSelectSql & " INTO TEMP " & sTmpTableName
    Else
        sSql = sSelectSql
    End If
    On Error GoTo SQLError
    t_dbMainDatabase.ExecuteSQL sSql
    fnCreateTempTable = True
    
    Exit Function

SQLError:
    tfnErrHandler SUB_NAME + "," + sCalledFrom, sSql
    fnCreateTempTable = False
End Function

Public Sub fnDropTempTable(sTmpTableName As String)
    Dim sSql As String
    sSql = "DROP TABLE " & sTmpTableName
    On Error Resume Next
    t_dbMainDatabase.ExecuteSQL sSql
    Exit Sub
End Sub

'default to do not show error message
Public Function fnManipulateRecords(nDBdatabase As Integer, sSql As String, sCalledFrom As String, Optional bShowError As Variant) As Boolean
    Const SUB_NAME = "fnManipulateRecords"
    
    On Error GoTo ErrorAccessRecords
    If nDBdatabase = nDBRemote Then
        t_dbMainDatabase.ExecuteSQL sSql
    ElseIf nDBdatabase = nDBLocal Then
        dbLocal.Execute sSql
    End If
    fnManipulateRecords = True
    Exit Function

ErrorAccessRecords:
    tfnErrHandler SUB_NAME + "," + sCalledFrom, sSql, bShowError
    fnManipulateRecords = False
End Function

