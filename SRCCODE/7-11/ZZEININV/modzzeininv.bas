Attribute VB_Name = "modzzeininv"
Option Explicit

Public dbLocal As Database
Public Const nDB_LOCAL As Integer = 1
Public Const nDB_REMOTE As Integer = 2

Public Const ColHeaderPrftctr As Integer = 0
Public Const ColHeaderPrftDesc As Integer = 1
Public Const ColHeaderRptDate As Integer = 2
Public Const ColHeaderVendor As Integer = 3
Public Const ColHeaderVendorName As Integer = 4
Public Const ColHeaderInvoice As Integer = 5
Public Const ColHeaderInvAmount As Integer = 6
Public ColHdnHeaderShift As Integer
Public ColHdnHeaderTerm As Integer
Public ColHdnHeaderType As Integer
Public ColHdnHeaderDraft As Integer
Public colHdnHeaderInvDate As Integer
Public Const ColDetailLine As Integer = 0
Public Const ColDetailItemCode As Integer = 1
Public Const ColDetailItemDesc As Integer = 2
Public Const ColDetailQty As Integer = 3
Public Const ColDetailCost As Integer = 4
Public Const ColDetailExtCost As Integer = 5
Public Const ColDetailPBCost As Integer = 6
Public Const ColDetailExtPBCost As Integer = 7
Public Const ColDetailRetail As Integer = 8
Public Const ColDetailExtRetail As Integer = 9
Public colHdnDetailUOM As Integer
Public Const nMaxDetailCol As Integer = 10

Public Function fnExecuteSQL(szSQL As String, Optional nDB As Variant, _
                Optional sCalledFrom As Variant, Optional bShowError As Variant) As Boolean
                
    On Error GoTo SQLError
    
    If IsMissing(nDB) Then
        nDB = nDB_REMOTE
    End If
    
    Select Case nDB
        Case nDB_LOCAL 'local
            dbLocal.Execute szSQL
        Case nDB_REMOTE 'remote
            t_dbMainDatabase.ExecuteSQL szSQL
    End Select
    
    fnExecuteSQL = True
    Exit Function
    
SQLError:

    If IsMissing(sCalledFrom) Then
        sCalledFrom = ""
    End If
    
    If IsMissing(bShowError) Then
        bShowError = True
    End If

    tfnErrHandler "fnExecuteSQL, " & sCalledFrom, szSQL, bShowError
    On Error GoTo 0
    
End Function

' Get records from the given SQL statement
' nDB = 1 ---> Informax Database (remote)
'     = 2 ---> Access Database (local)
'This function will return a recordcount
Public Function fnGetRecord(rsTemp As Recordset, strSQL As String, Optional nDB As Integer, Optional sCalledFrom As String, Optional bShowError As Variant) As Long
    Const SUB_NAME = "fnGetRecord"

    On Error GoTo SQLError
    
    If IsMissing(nDB) Then
        nDB = nDB_REMOTE
    End If
    
    Select Case nDB
        Case nDB_LOCAL
            Set rsTemp = dbLocal.OpenRecordset(strSQL, dbOpenSnapshot)
        Case nDB_REMOTE
            Set rsTemp = t_dbMainDatabase.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    End Select
    
    If rsTemp.RecordCount > 0 Then  'the following code is used to get the correct
        rsTemp.MoveLast             'RecordCount of the RecordSet
        rsTemp.MoveFirst
    End If
    
    fnGetRecord = rsTemp.RecordCount
    Exit Function
    
SQLError:
    
    If IsMissing(sCalledFrom) Then
        sCalledFrom = ""
    End If
    
    If IsMissing(bShowError) Then
        bShowError = True
    End If

    tfnErrHandler SUB_NAME + "," + sCalledFrom, strSQL, bShowError
    fnGetRecord = -9999
End Function


