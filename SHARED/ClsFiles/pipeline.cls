VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPipeLine"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit
    Private Const sTABLENAME As String = "ControlTable"
    Private Const sTABLECOUNT As String = "TableCount"
    Private Const sTABLEPARAM As String = "TableParam"
    
    Private Const REQUEST_TYPE_FIELD As Integer = 1
    Private Const REQUEST_TYPE_TABLE As Integer = 2
    
    Private Const nERR_FREE As Integer = 0
    Private Const nERR_CREATE_TABLE As Integer = 1
    Private Const nERR_INSERT_RECORD As Integer = 2
    Private Const nERR_UPDATE_RECORD As Integer = 3
    Private Const nERR_SQL_QUERY As Integer = 4
    Private Const nERR_NO_REQUEST As Integer = 5
    Private Const nERR_NO_PARAM As Integer = 6

    Private dbInformix As DataBase
    Private nErrorNumber As Integer
    Private nTableCount As Integer

Public Function ErrorStatus()
    'Return error status: = 0 ---> No error
    '                     > 0 ---> Error occurs.
    'Call Description to get the error description

    ErrorStatus = nErrorNumber
End Function

Private Function fnExecuteSQL(strSQL) As Boolean
    
    'Execute a SQL statement. Return true if successful

    On Error GoTo errExecute
    dbInformix.ExecuteSQL strSQL
    fnExecuteSQL = True

extExecute:
    On Error GoTo 0
    Exit Function

errExecute:
    Err.Clear
    fnExecuteSQL = False
    Resume extExecute
End Function


Public Sub GetRequest(sModuleID As String, _
                      sConsumer As String, _
                      nType As Integer, _
                      Optional sField As Variant, _
                      Optional sTable As Variant)
    'Return the consumer, type and field of a request

    Dim strSQL As String
    Dim rsTemp As Recordset

    strSQL = "SELECT tt_consumer, tt_param_type, tt_param_field, tt_param_table FROM " & sTABLENAME & " WHERE tt_supplier = '" & sModuleID & "'"
    Set rsTemp = fnGetRecordSet(strSQL)
    If Not rsTemp Is Nothing Then
        If rsTemp.RecordCount > 0 Then
            sConsumer = rsTemp!tt_consumer
            nType = rsTemp!tt_param_type
            If Not IsMissing(sField) Then
                If Not IsNull(rsTemp!tt_param_field) Then
                    sField = rsTemp!tt_param_field
                Else
                    sField = ""
                End If
            End If
            If Not IsMissing(sTable) Then
                If Not IsNull(rsTemp!tt_param_table) Then
                    sTable = rsTemp!tt_param_table
                Else
                    sTable = ""
                End If
            End If
        End If
    End If

End Sub


Public Sub WaitForData(sForID As String, Optional sFromID As Variant)
    'Return the name of parameter requested
    ' (Usually it is a field in a database table)

    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim bFlag As Boolean

    strSQL = "SELECT tt_flag FROM " & sTABLENAME & " WHERE tt_consumer = '" & sForID & "'"
    If Not IsMissing(sFromID) Then
        strSQL = strSQL & " AND tt_supplier = '" & sFromID & "'"
    End If

    bFlag = True
    While bFlag
        DoEvents
        Set rsTemp = fnGetRecordSet(strSQL)
        If rsTemp Is Nothing Then
            MsgBox "Recordset is nothing"
            Exit Sub
        End If
        If rsTemp.RecordCount = 0 Then
            MsgBox "No record"
            Exit Sub
        End If
        DoEvents
        If rsTemp!tt_flag = "P" Then
            bFlag = False
        End If
    Wend
End Sub
Public Function RequestType(sModuleID As String) As Integer
    'Return the type of parameter requested
    ' (Usually it is a field in a database table)

    Dim strSQL As String
    Dim rsTemp As Recordset
    
    RequestType = 0
    strSQL = "SELECT tt_param_type FROM " & sTABLENAME & " WHERE tt_supplier = '" & sModuleID & "'"
    Set rsTemp = fnGetRecordSet(strSQL)
    If Not rsTemp Is Nothing Then
        If rsTemp.RecordCount > 0 Then
            If Not IsNull(rsTemp!tt_param_type) Then
                RequestType = rsTemp!tt_param_type
            End If
        End If
    End If

End Function
Private Property Get TableCount() As Integer
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    TableCount = 1
    strSQL = "SELECT nCount FROM " & sTABLECOUNT
    Set rsTemp = fnGetRecordSet(strSQL)
    If rsTemp Is Nothing Then
        Exit Property
    End If
    If rsTemp.RecordCount > 0 Then
        TableCount = rsTemp!nCount
    End If

End Property
Private Property Let TableCount(nCount As Integer)
    Dim strSQL As String
    
    strSQL = "UPDATE " & sTABLECOUNT & " SET nCount = " & CStr(nCount)
    If fnExecuteSQL(strSQL) Then
        nErrorNumber = nERR_FREE
    Else
        nErrorNumber = nERR_UPDATE_RECORD
    End If
End Property

Private Function fnGetRecordSet(strSQL) As Recordset
    'Return a recordset from the given SQL if successful
    '       nothing if fails

    On Error GoTo errOpen
    Set fnGetRecordSet = dbInformix.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    nErrorNumber = nERR_FREE

extOpen:
    Exit Function
errOpen:
    nErrorNumber = nERR_SQL_QUERY
    Err.Clear
    Set fnGetRecordSet = Nothing
    Resume extOpen
End Function

Private Function fnTableSetup() As Boolean
    'Create a temp table to hold requests.
    ' Return True if successful and False if fails.
    Const sCreateTable As String = "CREATE TABLE "

    Dim strSQL As String
    Dim rsTemp As Recordset
    
    strSQL = sCreateTable & sTABLENAME & "(tt_consumer CHAR(8), tt_supplier CHAR(8), tt_param_type SMALLINT, tt_param_field CHAR(30), tt_param_table CHAR(30), tt_param_value CHAR(90), tt_flag CHAR(1))"
    If fnExecuteSQL(strSQL) Then
        nErrorNumber = nERR_FREE
    Else
        strSQL = "SELECT tt_consumer FROM " & sTABLENAME & " WHERE ROWID = 1"
        Set rsTemp = fnGetRecordSet(strSQL)
        If rsTemp Is Nothing Then
            nErrorNumber = nERR_CREATE_TABLE
        Else
            nErrorNumber = nERR_FREE
        End If
    End If

    strSQL = sCreateTable & sTABLECOUNT & "(nCount SMALLINT)"
    If fnExecuteSQL(strSQL) Then
        nErrorNumber = nERR_FREE
        strSQL = "INSERT INTO " & sTABLECOUNT & "(nCount) VALUES(1)"
        If fnExecuteSQL(strSQL) Then
            nErrorNumber = nERR_FREE
        Else
            nErrorNumber = nERR_CREATE_TABLE
        End If
    Else
        strSQL = "SELECT nCount FROM " & sTABLECOUNT
        Set rsTemp = fnGetRecordSet(strSQL)
        If rsTemp Is Nothing Then
            nErrorNumber = nERR_CREATE_TABLE
        Else
            nErrorNumber = nERR_FREE
        End If
    End If
    
    If nErrorNumber = nERR_FREE Then
        fnTableSetup = True
    Else
        fnTableSetup = False
    End If
End Function

Public Function GetConsumer(sModuleID As String) As String
    'Return the name of the consumer who put the request out

    Dim strSQL As String
    Dim rsTemp As Recordset

    GetConsumer = ""
    strSQL = "SELECT tt_consumer FROM " & sTABLENAME & " WHERE tt_supplier = '" & sModuleID & "'"
    Set rsTemp = fnGetRecordSet(strSQL)
    If Not rsTemp Is Nothing Then
        If rsTemp.RecordCount > 0 Then
            GetConsumer = rsTemp!tt_consumer
        End If
    End If

End Function


Public Function GetRecords(sModuleID As String, Optional sFromID As Variant) As Recordset
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim sTable As String
    
    strSQL = "SELECT tt_param_value FROM " & sTABLENAME & " WHERE tt_consumer = '" & sModuleID & "'"
    If Not IsMissing(sFromID) Then
        strSQL = strSQL & " AND tt_supplier = '" & sFromID & "'"
    End If
    Set rsTemp = fnGetRecordSet(strSQL)

    If rsTemp Is Nothing Then
        MsgBox "Recordset is nothing(controltable) " & CRLF & strSQL
        Exit Function
    End If
    If rsTemp.RecordCount = 0 Then
        nErrorNumber = nERR_NO_REQUEST
        MsgBox "No Records(controltable) " & CRLF & strSQL
        Exit Function
    End If
    
    sTable = Trim$(rsTemp!tt_param_value)
    strSQL = "SELECT * FROM " & sTable

    Set rsTemp = fnGetRecordSet(strSQL)
    If rsTemp Is Nothing Then
        MsgBox "Recordset is nothing(paramtable) " & CRLF & strSQL
        Exit Function
    End If
    
    If rsTemp.RecordCount = 0 Then
        nErrorNumber = nERR_NO_PARAM
        MsgBox "No records(paramtable) " & CRLF & strSQL
        Exit Function
    End If

    MsgBox strSQL
    Clipboard.SetText strSQL
    Set GetRecords = rsTemp
    
    strSQL = "DELETE FROM " & sTABLENAME & " WHERE tt_consumer = '" & sModuleID & "'"
    If Not IsMissing(sFromID) Then
        strSQL = strSQL & " AND tt_supplier = '" & sFromID & "'"
    End If
    If fnExecuteSQL(strSQL) Then
        nErrorNumber = nERR_FREE
    Else
        nErrorNumber = nERR_SQL_QUERY
    End If

    'Drop the table after it is done ?
'    If fnExecuteSQL("DROP TABLE " & sTable) Then
'        nErrorNumber = nERR_FREE
'    Else
'        nErrorNumber = nERR_SQL_QUERY
'    End If
    
End Function

Public Function ParamName(sModuleID As String) As String
    'Return the name of parameter requested
    ' (Usually it is a field in a database table)

    Dim strSQL As String
    Dim rsTemp As Recordset
    
    ParamName = ""
    strSQL = "SELECT tt_param_field FROM " & sTABLENAME & " WHERE tt_supplier = '" & sModuleID & "'"
    Set rsTemp = fnGetRecordSet(strSQL)
    If Not rsTemp Is Nothing Then
        If rsTemp.RecordCount > 0 Then
            If Not IsNull(rsTemp!tt_param_field) Then
                ParamName = rsTemp!tt_param_field
            End If
        End If
    End If

End Function

Public Function RequestNumber(sModuleID As String) As Integer
    'Return the number of parameters requested
    ' 0 means no request is set for this module

    Dim strSQL As String
    Dim rsTemp As Recordset

    RequestNumber = 0
    strSQL = "SELECT tt_supplier FROM " & sTABLENAME & " WHERE tt_supplier = '" & sModuleID & "'"
    Set rsTemp = fnGetRecordSet(strSQL)
    If Not rsTemp Is Nothing Then
        If rsTemp.RecordCount > 0 Then
            RequestNumber = rsTemp.RecordCount
        End If
    End If

End Function
Public Sub PutParam(sValue As String, sForID As String, Optional sField As Variant, Optional sModuleID As Variant)
    'Set the value of a parameter

    Dim strSQL As String

    strSQL = "UPDATE " & sTABLENAME & " SET tt_flag = 'P', tt_param_value = '" & sValue _
           & "' WHERE tt_consumer = '" & sForID & "'"
    If Not IsMissing(sField) Then
        strSQL = strSQL & " AND tt_param_field = '" & sField & "'"
    End If
    If Not IsMissing(sModuleID) Then
        strSQL = strSQL & " AND tt_supplier = '" & sModuleID & "'"
    End If

    If fnExecuteSQL(strSQL) Then
        nErrorNumber = nERR_FREE
    Else
        nErrorNumber = nERR_UPDATE_RECORD
    End If

End Sub

Public Function GetParam(sModuleID As String, Optional sField As Variant, Optional sFromID As Variant) As String
    ' Return the value of a parameter.

    Dim strSQL As String
    Dim rsTemp As Recordset

    strSQL = "SELECT tt_param_value FROM " & sTABLENAME _
           & " WHERE tt_consumer = '" & sModuleID & "'"
    If Not IsMissing(sField) Then
        strSQL = strSQL & " AND tt_param_field = '" & sField & "'"
    End If
    
    If Not IsMissing(sFromID) Then
        strSQL = strSQL & " AND tt_supplier = '" & sFromID & "'"
    End If

    Set rsTemp = fnGetRecordSet(strSQL)
    If Not rsTemp Is Nothing Then
        If rsTemp.RecordCount > 0 Then
            GetParam = rsTemp!tt_param_value
        Else
            GetParam = ""
            nErrorNumber = nERR_NO_PARAM
        End If
    End If

End Function
Public Sub PutRecords(sSQL As String, sForID As String, Optional sMyID As Variant)
    ' Put a record to a parameter table
    
    Dim strSQL As String
    Dim nCount As Integer
    Dim sTable As String
    
    nCount = TableCount
    If nErrorNumber > 0 Then
        Exit Sub
    End If

    sTable = sTABLEPARAM & CStr(nCount)
    strSQL = sSQL & " INTO TEMP " & sTable
    
    Clipboard.SetText strSQL
    MsgBox "Insertint record " & CRLF & strSQL
    
    If fnExecuteSQL(strSQL) Then
        nErrorNumber = nERR_FREE
        strSQL = "UPDATE " & sTABLENAME & " SET tt_flag = 'P', tt_param_value = '" & sTable _
               & "' WHERE tt_consumer = '" & sForID & "'"
        If Not IsMissing(sMyID) Then
            strSQL = strSQL & " AND tt_supplier = '" & sMyID & "'"
        End If
        If fnExecuteSQL(strSQL) Then
            nErrorNumber = nERR_FREE
            TableCount = nCount + 1
        Else
            nErrorNumber = nERR_INSERT_RECORD
        End If
    Else
        nErrorNumber = nERR_INSERT_RECORD
    End If
End Sub

Public Sub PutRequest(sModuleID As String, sFromID As String, nType As Integer, Optional sName As Variant)
'    Debug.Print fnTableSetup
'Exit Sub

    'Put a request for a module

    If Not fnTableSetup Then
        Exit Sub
    End If
    Dim strSQL As String
    Dim sField As String
    Dim sTable As String
    Dim sSource As String
    
    nErrorNumber = nERR_INSERT_RECORD
    sTable = ""
    If IsMissing(sName) Then
        sField = ""
    Else
        sField = sName
        If nType = REQUEST_TYPE_FIELD Then
            Dim i As Integer
            Dim l As Integer
            
            i = InStr(sName, ".")
            If i > 0 Then
                sTable = Left(sName, i - 1)
                l = Len(sName) - i
                If l > 0 Then
                    sField = Right(sName, l)
                End If
            End If
        End If
    End If
    If IsMissing(sFromID) Then
        sSource = ""
    Else
        sSource = sFromID
    End If
    strSQL = "INSERT INTO " & sTABLENAME & " VALUES('" _
           & sModuleID & "', '" & sSource & "', " & CStr(nType) _
           & ", '" & sField & "', '" & sTable & "', '', 'R')"
    On Error GoTo errInsert
    dbInformix.ExecuteSQL strSQL
    nErrorNumber = nERR_FREE
    
extRequest:
    On Error GoTo 0
    Exit Sub

errInsert:
    Err.Clear
    Resume extRequest
End Sub
Property Get RequestField() As Integer
    RequestField = REQUEST_TYPE_FIELD
End Property

Property Get RequestRecord() As Integer
    RequestRecord = REQUEST_TYPE_TABLE
End Property

Public Function TableName(sModuleID As String) As String

    'Return the name of the table in which the parameter requested is
    ' (Usually it is a field in a database table)

    Dim strSQL As String
    Dim rsTemp As Recordset
    
    TableName = ""
    strSQL = "SELECT tt_param_table FROM " & sTABLENAME & " WHERE tt_supplier = '" & sModuleID & "'"
    Set rsTemp = fnGetRecordSet(strSQL)
    If Not rsTemp Is Nothing Then
        If rsTemp.RecordCount > 0 Then
            If Not IsNull(rsTemp!tt_param_table) Then
                TableName = rsTemp!tt_param_table
            End If
        End If
    End If


End Function

Private Sub Class_Initialize()
    #If FACTOR_MENU = 0 Then
        Set dbInformix = t_dbMainDatabase
    #Else
        Dim objTemp As Object
        MsgBox "Step 1"
        Set objTemp = t_oleObject.DatabaseHandle
        MsgBox "Step 2"
        Set dbInformix = objTemp
    #End If
End Sub

