VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTableMaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'---------------------------------------------------------------------------------------
' Module    : clsTableMaint
' DateTime  : 5/13/2005 14:49
' Author    : RSanders
' Purpose   : Generic table maintenance class to maintain a single record
'---------------------------------------------------------------------------------------
'
'       This class has been designed as a generic method to replace the
'       mnt_vendor, mnt_inv_2loc, mnt_inv_master, and mnt_whlsale stored procedures.
'
'       fnUpdateRecord
'
'           The fnUpdateRecord function will take the original data as presented to
'           the user and compare it with the new values to be updated.  If there is no
'           difference then no action is taken.  If there is a difference then the
'           object will then compare the original value with the current value.
'           If we note a change, an error is raised, indicating that the operation
'           cannot proceed due to values having changed.  Otherwise, the table is
'           updated with only the fields that have changed.
'
'           After updating the table, the table is accessed once again and a recordset
'           returned to the user.  The first byte of the record set is a status flag.
'           A zero indicates a failure and a one indicates success.  The last field of
'           the record set is a status message.  Between the status flag and the status
'           message is the updated data that is returned from the table.
'
'       To use this class it is first necessary to create the object.
'
'               Dim objTableMaint As clsTableMaint
'               Set objTableMaint = New clsTableMaint
'               With objTableMaint
'
'       Then we identify the table characteristics to the object.  Note the use of the
'       "NULL" keyword.  When specified it indicates to the object what value is to be
'       used if an empty string is passed to the object.  In the example we will use
'       a null but the null could have just as easily been a "0".
'
'               .CREATE_TABLE t_dbMainDatabase, "p_vendor"
'
'               .CREATE_FIELD "pm_vendor", "integer"
'               .CREATE_FIELD "pm_status", "char", 1
'               .CREATE_FIELD "pm_hist_lvl", "smallint"
'               .CREATE_FIELD "pm_credit_limit", "decimal", 10, 2, "NULL"
'               etc..... repeating for each field in the table to be maintained
'
'               .CREATE_INDEX "pm_vendor"       ' Note: Repeat for each index
'               etc..... repeating for each index field
'
'       The user of this class may then proceed in one of two ways.  The first method is
'       preferable due to peformance considerations plus the ability to not require that
'       all fields be specified.  The second method may also be used in certain cases if
'       the impact on existing code is a consideration.
'
'           Method 1 - Setting properties directly into the object
'
'               ' First store the original field values as presented to the user
'               ' into  the object
'               .OLD_VALUE("pm_vendor") = frmAPFMASTR!txtVendor
'               .OLD_VALUE("pm_status") = fnSQLString(sOldStatus)
'               etc..... for each field in the table to be maintained
'
'               ' And now store the new field values
'               .NEW_VALUE("pm_vendor") = frmAPFMASTR!txtVendor
'               .NEW_VALUE("pm_status") = fnSQLString(Trim(frmAPFMASTR!TxtStatus))
'               etc..... for each field in the table to be maintained
'
'               ' and finally use the fnUpdateRecord function in the object
'               ' Note that the sql statement is not passed to the function
'               nRecordCount = .fnUpdateRecord("fnUpdateVendor", rsTemp, nDBRemote)
'
'           Method 2 - Using an EXECUTE PROCEDURE SQL string to pass the data to
'                      the object
'
'               ' Create the SQL string just as if we were going to access
'               ' a stored procedure
'               strSQL = "EXECUTE PROCEDURE mnt_vendor (" _
'                       + sOldVendor + "," + old_status + "," _
'                       etc..... for each field in the table
'
'               strSQL = strSQL _
'                       + frmAPFMASTR!txtVendor + "," + new_status + "," _
'                       etc..... for each field in the table
'
'               ' And finally use the fnUpdateRecord function in the object
'               ' Note that the sql string is passed to the function
'               nRecordCount = .fnUpdateRecord("fnUpdateVendor", rsTemp, nDBRemote, strSQL)
'
'       fnInsertRecord
'
'           In addition to the fnUpdateRecord function, an fnInsertRecord function has also
'           been added.  The intent is to take the values stored into the object using the
'           NEW_VALUE function, create and execute a sql string to insert a record into the
'           table.  The current value of the record will be returned in the same manner as the
'           fnUpdateRecord function.
'
'
'       fnSelectRecord
'
'           A select record functions is also available.  It will select the fields
'           as defined by the CREATE_FIELD function for the record matching the criteria
'           identified by the CREATE_INDEX function.  Once selected, the values are
'           saved, thus eliminating the need to call the OLD_VALUE and NEW_VALUE funcions
'           (for those fields that are not to be changed) prior to issuing a call to
'           fnUpdateRecord.
'
'---------------------------------------------------------------------------------------

Public status_flag              As Long             ' 0=Bad, 1=Good
Public status_message           As String           ' One of the following messages
    
Public InitializingMsg          As String           ' These are set when we do a CREATE_TABLE
Public BeginningMsg             As String           ' because we use the table name in
Public SelectedMsg              As String           ' several of them.  Note that you have the
Public NotFoundMsg              As String           ' option of changed them to messages that
Public ChangedMsg               As String           ' may be more appropriate for your specific
Public SuccessMsg               As String           ' application.
Public UpdatingMsg              As String
Public InsertingMsg             As String
Public ExceptionMsg             As String
Public OpenParenMsg             As String
Public CloseParenMsg            As String

'=================================================================================

Private Enum enumFieldTypes                         ' These are the field types we currently support
    TypeUndefined = 0
    TypeChar = 1
    TypeDate = 2
    TypeDecimal = 3
    TypeInteger = 4
    TypeSmallInt = 5
End Enum

Private arrFieldType(100)       As enumFieldTypes   ' TypeUndefined signifies end of array
Private arrFieldName(100)       As String
Private arrFieldLength(100)     As Integer
Private arrFieldPrecision(100)  As Integer
Private arrFieldValueOld(100)   As String
Private arrFieldValueNew(100)   As String
Private arrFieldNull(100)       As String
Private arrIndex(100)           As Boolean

'=================================================================================

Private Const intError          As Integer = 0      ' These values are for the status flag
Private Const intNoError        As Integer = 1

Private blnRecordFound          As Boolean
Private blnDataChanged          As Boolean

Private intRecordCount          As Integer

Private strTableName            As String
Private strInsertName           As String           ' This is for building SQL string
Private strInsertValue          As String
Private strUpdate               As String

Private dbDatabase              As Database
Private rsCurrent               As Recordset

'---------------------------------------------------------------------------------------
' Procedure : CREATE_TABLE
' DateTime  : 5/13/2005 14:44
' Author    : RSanders
' Purpose   : This function allows the user of this class to define a table to
'             be maintained.
'---------------------------------------------------------------------------------------
'
Public Function CREATE_TABLE(pDatabase As Database, pTableName As String) As Boolean
    
    CREATE_TABLE = False
    
    Set dbDatabase = pDatabase
    strTableName = LCase(Trim(pTableName))
    
    InitializingMsg = "Initializing variables."
    BeginningMsg = "Beginning " & LCase(strTableName) & " table maintenance procedure."
    SelectedMsg = "Record selected from " & LCase(strTableName) & " table."
    NotFoundMsg = "No matching record in " & LCase(strTableName) & " table."
    ChangedMsg = "Values in the " & LCase(strTableName) & " table have been changed since last query."
    SuccessMsg = "Procedure successful."
    UpdatingMsg = "Updating " & LCase(strTableName) & " table."
    InsertingMsg = "Inserting a record into the " & LCase(strTableName) & " table."
    ExceptionMsg = "Exception Error: @Err_Number, @Err_Description "
    OpenParenMsg = "SQL string lacks open parenthesis for parameters used in stored procedure"
    CloseParenMsg = "SQL string lacks close parenthesis for parameters used in stored procedure"

    CREATE_TABLE = True
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : CREATE_FIELD
' DateTime  : 5/13/2005 14:45
' Author    : RSanders
' Purpose   : This function allows the user of this class to define a field to
'             be maintained.
'---------------------------------------------------------------------------------------
'
Public Function CREATE_FIELD(pFieldName As String, _
                              pFieldType As String, _
                              Optional pFieldLength As Integer = -1, _
                              Optional pFieldPrecision As Integer = -0, _
                              Optional pNull As String = "") As Boolean

    Dim i As Integer
    Dim strFieldName As String
    Dim strFieldType As String
    
    CREATE_FIELD = False
    
    strFieldName = LCase(Trim(pFieldName))
    strFieldType = LCase(Trim(pFieldType))
    
    i = fnFindFieldName(strFieldName)
    
    If Not i < UBound(arrFieldName) - 1 Then
        Err.Raise 9999, "clsTableMaint", "The CREATE_FIELD function in clsTableMaint has exceeded the number of fields allowed while trying to add the " & strFieldName & " field."
    End If
    
    If arrFieldType(i) <> TypeUndefined Then
        Err.Raise 9999, "clsTableMaint", "The CREATE_FIELD function in clsTableMaint found a duplicate of the " & strFieldName & " field."
    End If
    
    Select Case strFieldType
        Case "char"
            arrFieldType(i) = TypeChar
        Case "decimal"
            arrFieldType(i) = TypeDecimal
        Case "integer"
            arrFieldType(i) = TypeInteger
        Case "smallint"
            arrFieldType(i) = TypeSmallInt
        Case "date"
            arrFieldType(i) = TypeDate
        Case Else
            Err.Raise 9999, "clsTableMaint", "The CREATE_FIELD function in clsTableMaint has encountered an invalid field type (" & strFieldType & ") while trying to add the " & strFieldName & " field."
    End Select
    
    arrFieldName(i) = strFieldName
    arrFieldLength(i) = pFieldLength
    arrFieldPrecision(i) = pFieldPrecision
    arrFieldNull(i) = pNull
    
    i = i + 1
    arrFieldName(i) = ""
    arrFieldType(i) = TypeUndefined
    arrIndex(i) = False
    
    CREATE_FIELD = True
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : CREATE_INDEX
' DateTime  : 5/13/2005 14:45
' Author    : RSanders
' Purpose   : This function allows the user of this class to define a unique index
'             for the table to be maintained.
'---------------------------------------------------------------------------------------
'
Public Function CREATE_INDEX(pFieldName As String) As Boolean

    Dim i As Integer
    Dim strFieldName As String
    
    CREATE_INDEX = False
    strFieldName = LCase(Trim(pFieldName))
    i = fnFindFieldName(strFieldName)
    If arrFieldType(i) = TypeUndefined Then
        Err.Raise 9999, "clsTableMaint", "The CREATE_INDEX function in clsTableMaint failed to find the " & strFieldName & " field."
    End If
    If arrIndex(i) = True Then
        Err.Raise 9999, "clsTableMaint", "The CREATE_INDEX function in clsTableMaint has found a duplicate of the " & strFieldName & " index field."
    End If
    arrIndex(i) = True
    CREATE_INDEX = True
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : OLD_VALUE
' DateTime  : 5/13/2005 14:47
' Author    : RSanders
' Purpose   : This property allows the user of this class to specify the value that
'             originally was read from the table.  Its use is intended for those
'             programs that perform their own query of the database instead of using
'             the fnSelectRecord function of this class.  Users of the fnSelectRecord
'             class should not be required to set the OLD_VALUE becase the function
'             will set both the old and new fields automatically when the data is
'             read.  Note that character and date fields stored by this property will
'             have the quotes automatically inserted at the beginning and end of the
'             string if they are missing.  Thus the get property will return a string
'             with the data wrapped in single quotes.
'---------------------------------------------------------------------------------------
'
Public Property Let OLD_VALUE(pFieldName As String, _
                              pFieldValue As String)
    
    Dim i As Integer
    Dim strFieldName As String
    
    strFieldName = LCase(Trim(pFieldName))
    i = fnFindFieldName(strFieldName)
    Select Case arrFieldType(i)
        Case TypeChar, TypeDate
            arrFieldValueOld(i) = fnCheckForNull(pFieldValue)
        Case TypeDecimal, TypeInteger, TypeSmallInt
            arrFieldValueOld(i) = fnCheckForNull(pFieldValue)
            arrFieldValueOld(i) = Replace(arrFieldValueOld(i), "'", "")
        Case Else
            Err.Raise 9999, "clsTableMaint", "The Let OLD_VALUE property failed to find the " & strFieldName & " field."
    End Select
End Property

Public Property Get OLD_VALUE(pFieldName As String) As String

    Dim i As Integer
    Dim strFieldName As String
    
    strFieldName = LCase(Trim(pFieldName))
    i = fnFindFieldName(strFieldName)
    If arrFieldType(i) = TypeUndefined Then
        Err.Raise 9999, "clsTableMaint", "The Get OLD_VALUE property failed to find the " & strFieldName & " field."
    End If
    OLD_VALUE = arrFieldValueOld(i)
    
End Property

'---------------------------------------------------------------------------------------
' Procedure : NEW_VALUE
' DateTime  : 5/13/2005 14:47
' Author    : RSanders
' Purpose   : This property allows the user of this class to specify the value that
'             is to be written to the table.  When performing an update, the OLD_VALUES
'             and the NEW_VALUES are compared and only those fields that have changed
'             will be updated.  Note that character and date fields stored by this
'             property will have the quotes automatically inserted at the beginning and
'             end of the string if they are missing.  Thus the get property will return
'             a string with the data wrapped in single quotes.
'---------------------------------------------------------------------------------------
'
Public Property Let NEW_VALUE(pFieldName As String, _
                              pFieldValue As String)
    
    Dim i As Integer
    Dim strFieldName As String
    
    strFieldName = LCase(Trim(pFieldName))
    i = fnFindFieldName(strFieldName)
    Select Case arrFieldType(i)
        Case TypeChar, TypeDate
            arrFieldValueNew(i) = fnCheckForNull(pFieldValue)
        Case TypeDecimal, TypeInteger, TypeSmallInt
            arrFieldValueNew(i) = fnCheckForNull(pFieldValue)
            arrFieldValueNew(i) = Replace(arrFieldValueNew(i), "'", "")
        Case Else
            Err.Raise 9999, "clsTableMaint", "The Let NEW_VALUE property failed to find the " & strFieldName & " field."
    End Select
    
End Property

Public Property Get NEW_VALUE(pFieldName As String) As String

    Dim i As Integer
    Dim strFieldName As String
    
    strFieldName = LCase(Trim(pFieldName))
    i = fnFindFieldName(strFieldName)
    If arrFieldType(i) = TypeUndefined Then
        Err.Raise 9999, "clsTableMaint", "The Get NEW_VALUE property failed to find the " & strFieldName & " field."
    End If
    NEW_VALUE = arrFieldValueNew(i)
    
End Property


'---------------------------------------------------------------------------------------
' Procedure : fnUpdateRecord
' DateTime  : 5/13/2005 14:47
' Author    : RSanders
' Purpose   : This function will update a record in the specified table.  Only data that
'             has changed will be updated as determined by a comparison of the OLD_VALUE
'             and the NEW_VALUE fields.  The record set specified as a parameter will be
'             updated to reflect the current data.  If the pSQL parameter is not a blank
'             then it is assumed that the string is an "EXECUTE PROCEDURE" sql string and
'             the string will be parsed to obtain the OLD_VALUES and the NEW_VALUES.
'             The first byte of the record set returned is a status byte where 0 indicates
'             a failure and 1 indicates success.  At the end of the record set is a
'             status message field which provides additional details related to the
'             failure.  Note that it is possible that a record set cannot be returned.
'             A check of the number of records returned should be done to determine if
'             the operations succeeded.
'---------------------------------------------------------------------------------------
'
Public Function fnUpdateRecord _
    (pCalledFrom As String, _
     pCurrent As Recordset, _
     Optional pDB As Integer = 1, _
     Optional pSQL As String = "", _
     Optional pShowError As Variant) As Long
    
    Dim i As Integer
    
    On Error GoTo Exception

    status_message = InitializingMsg
    status_flag = intError
    
    fnUpdateRecord = -9999
    Set pCurrent = Nothing

    blnRecordFound = False
    blnDataChanged = False
    
    intRecordCount = 0
    
    strUpdate = ""
    strInsertName = ""
    strInsertValue = ""
    
    If Trim(pSQL) <> "" Then
        fnParseSQL pSQL         ' Parse SQL into input and new fields
    End If
    
    status_message = BeginningMsg
    
    ' Get current values
    Set rsCurrent = Nothing
    Set rsCurrent = fnSelectData(pDB)
    
    If Not rsCurrent.EOF Then
        
        status_message = SelectedMsg
        blnRecordFound = True    ' We have a valid record
        
        ' The following code has been changed from the original stored procedure.
        ' Many of the fields were not included in the check and the original
        ' code would update every field instead of only those that have changed
        ' as we do here.
       
        For i = 0 To UBound(arrFieldName) - 1
            Select Case arrFieldType(i)
                Case TypeChar
                    fnCheckChar arrFieldName(i), arrFieldValueOld(i), arrFieldValueNew(i), arrFieldLength(i), arrFieldNull(i)
                Case TypeDecimal, TypeInteger, TypeSmallInt
                    fnCheckNumb arrFieldName(i), arrFieldValueOld(i), arrFieldValueNew(i), arrFieldLength(i), arrFieldPrecision(i), arrFieldNull(i)
                Case TypeDate
                    fnCheckDate arrFieldName(i), arrFieldValueOld(i), arrFieldValueNew(i), arrFieldNull(i)
                Case TypeUndefined
                    Exit For
            End Select
        Next
         
    End If
    
    If blnRecordFound = False Then
        status_message = NotFoundMsg
        
    ElseIf blnDataChanged = True Then
        status_message = ChangedMsg
        
    Else
        
        If fnUpdateData(pDB, Mid(strUpdate, 3)) Then
    
            ' Set flag to show that the update went as expected
            status_flag = intNoError
            status_message = SuccessMsg
        
            For i = 0 To UBound(arrFieldName) - 1
                arrFieldValueOld(i) = arrFieldValueNew(i)
            Next
            
        End If
        
    End If

    Set pCurrent = Nothing
    Set pCurrent = fnSelectData(pDB)
    fnUpdateRecord = fnSaveOldAndNew(pCurrent)

    Exit Function
    
Exception:
    On Error Resume Next
    status_flag = intError
    status_message = ExceptionMsg & vbCrLf & status_message
    status_message = Replace(status_message, "@err_number", Err.Number)
    status_message = Replace(status_message, "@err_description", Err.Description)
    Set pCurrent = Nothing
    Set pCurrent = fnSelectData(pDB)
    tfnErrHandler "fnUpdateRecord," & pCalledFrom, pSQL, pShowError
End Function

'---------------------------------------------------------------------------------------
' Procedure : fnInsertRecord
' DateTime  : 5/13/2005 14:48
' Author    : RSanders
' Purpose   : This function will insert a record in the specified table.  Data to be
'             inserted must be specified via the NEW_VALUE property. The record set
'             specified as a parameter will be updated to reflect the current data.
'             If the pSQL parameter is not a blank then it is assumed that the string
'             is an "EXECUTE PROCEDURE" sql string and the string will be parsed to
'             obtain the OLD_VALUES and the NEW_VALUES.  It's format is identical to
'             that used by the fnUpdateRecord function.  The first byte of the record
'             set returned is a status byte where 0 indicates a failure and 1 indicates
'             success.  At the end of the record set is a status message field which
'             provides additional details related to the failure.  Note that it is
'             possible that a record set cannot be returned.  A check of the number
'             of records returned should be done to determine if the operations succeeded.
'---------------------------------------------------------------------------------------
'
Public Function fnInsertRecord _
    (pCalledFrom As String, _
     pCurrent As Recordset, _
     Optional pDB As Integer = 1, _
     Optional pSQL As String = "", _
     Optional pShowError As Variant) As Long
    
    Dim i As Integer
    
    On Error GoTo Exception

    status_message = InitializingMsg
    status_flag = intError
    
    fnInsertRecord = -9999
    Set pCurrent = Nothing

    strUpdate = ""
    strInsertName = ""
    strInsertValue = ""
    
    If Trim(pSQL) <> "" Then
        fnParseSQL pSQL         ' Parse SQL into input and new fields
    End If
    
    status_message = BeginningMsg
       
    ' The following code has been changed from the original stored procedure.
    ' Many of the fields were not included in the check and the original
    ' code would update every field instead of only those that have changed
    ' as we do here.
    
    For i = 0 To UBound(arrFieldName) - 1
        arrFieldValueOld(i) = arrFieldValueNew(i)
        Select Case arrFieldType(i)
            Case TypeChar
                 fnCheckChar arrFieldName(i), arrFieldValueOld(i), arrFieldValueNew(i), arrFieldLength(i), arrFieldNull(i)
            Case TypeDecimal, TypeInteger, TypeSmallInt
                fnCheckNumb arrFieldName(i), arrFieldValueOld(i), arrFieldValueNew(i), arrFieldLength(i), arrFieldPrecision(i), arrFieldNull(i)
            Case TypeDate
                fnCheckDate arrFieldName(i), arrFieldValueOld(i), arrFieldValueNew(i), arrFieldNull(i)
            Case TypeUndefined
                Exit For
         End Select
    Next

    If strInsertName <> "" Then
        strInsertName = Mid(strInsertName, 3)
    End If
    
    If strInsertValue <> "" Then
        strInsertValue = Mid(strInsertValue, 3)
    End If
    
    If fnInsertData(pDB, strInsertName, strInsertValue) Then

        ' Set flag to show that the update went as expected
        status_flag = intNoError
        status_message = SuccessMsg
        
        For i = 0 To UBound(arrFieldName) - 1
            arrFieldValueOld(i) = arrFieldValueNew(i)
        Next
    
    End If

    Set pCurrent = Nothing
    Set pCurrent = fnSelectData(pDB)
    fnInsertRecord = fnSaveOldAndNew(pCurrent)
    
    Exit Function
    
Exception:
    On Error Resume Next
    status_flag = intError
    status_message = ExceptionMsg & vbCrLf & status_message
    status_message = Replace(status_message, "@err_number", Err.Number)
    status_message = Replace(status_message, "@err_description", Err.Description)
    Set pCurrent = Nothing
    Set pCurrent = fnSelectData(pDB)
    tfnErrHandler "fnInsertRecord," & pCalledFrom, pSQL, pShowError
End Function

'---------------------------------------------------------------------------------------
' Procedure : fnSelectRecord
' DateTime  : 5/13/2005 14:48
' Author    : RSanders
' Purpose   : This function will select a record from the specified table.  Upon completion
'             the OLD_VALUE and NEW_VALUE properties will be set for each field.  The first
'             byte of the record set returned is a status byte where 0 indicates a failure
'             and 1 indicates success.  At the end of the record set is a status message
'             field which provides additional details related to the failure.  Note that it
'             is possible that a record set cannot be returned.  A check of the number
'             of records returned should be done to determine if the operations succeeded.
'---------------------------------------------------------------------------------------
'
Public Function fnSelectRecord _
    (pCalledFrom As String, _
     pCurrent As Recordset, _
     Optional pDB As Integer = 1, _
     Optional pSQL As String = "", _
     Optional pShowError As Variant) As Long
    
    Dim i As Integer
    
    On Error GoTo Exception

    status_message = InitializingMsg
    status_flag = intError
    
    fnSelectRecord = -9999
    Set pCurrent = Nothing
    
    If Trim(pSQL) <> "" Then
        fnParseSQL pSQL         ' Parse SQL into input and new fields
    End If
    
    status_message = BeginningMsg

    Set pCurrent = Nothing
    Set pCurrent = fnSelectData(pDB)
    fnSelectRecord = fnSaveOldAndNew(pCurrent)
    
    Exit Function
    
Exception:
    On Error Resume Next
    status_flag = intError
    status_message = ExceptionMsg & vbCrLf & status_message
    status_message = Replace(status_message, "@err_number", Err.Number)
    status_message = Replace(status_message, "@err_description", Err.Description)
    Set pCurrent = Nothing
    Set pCurrent = fnSelectData(pDB)
    tfnErrHandler "fnSelectRecord," & pCalledFrom, pSQL, pShowError
End Function

'---------------------------------------------------------------------------------------
' Procedure : fnSelectData
' DateTime  : 5/13/2005 14:48
' Author    : RSanders
' Purpose   : This is the function that executes the sql select.
'---------------------------------------------------------------------------------------
'
Private Function fnSelectData(pDB As Integer) As Recordset

    Dim i As Integer
    Dim strSelect As String
    
    strSelect = ""
    For i = 0 To UBound(arrFieldName) - 1
        If arrFieldType(i) = TypeUndefined Then Exit For
        strSelect = strSelect + ", " & arrFieldName(i)
    Next
    
    Const SQL_SELECT As String _
        = "SELECT @status_flag AS status_flag" _
        & "@strSelect, " _
        & "'@status_message' AS status_message " _
        & "FROM @strTableName " _
        & "@fnWhereClause"

    Dim strSQL As String
    
    strSQL = SQL_SELECT
    strSQL = Replace(strSQL, "@strTableName", strTableName)
    strSQL = Replace(strSQL, "@status_flag", status_flag)
    strSQL = Replace(strSQL, "@strSelect", strSelect)
    strSQL = Replace(strSQL, "@status_message", status_message)
    strSQL = Replace(strSQL, "@fnWhereClause", fnWhereClause)
    
    Select Case pDB
        Case 1  'remote
            Set fnSelectData = dbDatabase.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
        Case 2  'local
            Set fnSelectData = dbDatabase.OpenRecordset(strSQL, dbOpenSnapshot)
    End Select

End Function

'---------------------------------------------------------------------------------------
' Procedure : fnUpdateData
' DateTime  : 5/13/2005 14:48
' Author    : RSanders
' Purpose   : This is the function that executes the sql update.
'---------------------------------------------------------------------------------------
'
Private Function fnUpdateData(pDB As Integer, pUpdate As String) As Boolean

    Const SQL_UPDATE As String _
        = "UPDATE @strTableName " _
        & "SET @pUpdate " _
        & "@fnWhereClause"

    Dim strSQL As String
    
    fnUpdateData = False
    
    ' There is nothing to update if pUpdate is blank
    If pUpdate <> "" Then
    
        status_message = UpdatingMsg
        
        strSQL = SQL_UPDATE
        strSQL = Replace(strSQL, "@strTableName", strTableName)
        strSQL = Replace(strSQL, "@pUpdate", pUpdate)
        strSQL = Replace(strSQL, "@fnWhereClause", fnWhereClause)
        
        dbDatabase.ExecuteSQL strSQL

    End If
        
    fnUpdateData = True
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : fnInsertData
' DateTime  : 5/13/2005 14:48
' Author    : RSanders
' Purpose   : This is the function that executes the sql insert.
'---------------------------------------------------------------------------------------
'
Private Function fnInsertData(pDB As Integer, pInsertName As String, pInsertValue As String) As Boolean
    Const SQL_INSERT As String _
        = "INSERT INTO @strTableName " _
        & "(@pInsertName) VALUES (@pInsertValue) "

    Dim strSQL As String
    
    fnInsertData = False
    
    ' There is nothing to update if pUpdate is blank
    If pInsertName <> "" Then
    
        status_message = InsertingMsg
        
        strSQL = SQL_INSERT
        strSQL = Replace(strSQL, "@strTableName", strTableName)
        strSQL = Replace(strSQL, "@pInsertName", pInsertName)
        strSQL = Replace(strSQL, "@pInsertValue", pInsertValue)
        
        dbDatabase.ExecuteSQL strSQL

    End If
        
    fnInsertData = True
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : fnCheckChar
' DateTime  : 5/13/2005 14:48
' Author    : RSanders
' Purpose   : This function will build the sql string for the update and insert
'             functions for a character field.
'---------------------------------------------------------------------------------------
'
Private Function fnCheckChar(pField As String, pOld As String, pNew As String, Optional pLength = -1, Optional pNull As String = "") As Boolean
    Dim sCurrent As String

    strInsertName = ", " & pField
    strInsertValue = ", " & pNew
    
    If Not fnCharsAreEqual(pOld, pNew) Then
        
        sCurrent = Trim(pNull)
        If Not rsCurrent Is Nothing Then
            If pNull = "" Or Trim(fnGetField(rsCurrent.Fields(pField))) <> "" Then
                sCurrent = fnGetField(rsCurrent.Fields(pField))
                sCurrent = fnSQLString(sCurrent)
            End If
        End If
    
        If Not fnCharsAreEqual(pOld, sCurrent) Then
            blnDataChanged = True
            Debug.Print pField & "   " & pOld & "   " & sCurrent & "   " & pNew
        Else
            strUpdate = strUpdate & ", " & pField & " = " & pNew
        End If
        
    End If
    
End Function

Private Function fnCharsAreEqual(pChar1 As String, pChar2 As String) As Boolean
    Dim sChar1 As String
    Dim sChar2 As String

    fnCharsAreEqual = False

    sChar1 = Trim(pChar1)
    sChar2 = Trim(pChar2)
    
    If sChar1 = sChar2 Then
        fnCharsAreEqual = True
    ElseIf (sChar1 = "" Or UCase(sChar1) = "NULL") And (sChar2 = "" Or UCase(sChar2) = "NULL") Then
        fnCharsAreEqual = True
    End If
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : fnCheckNumb
' DateTime  : 5/13/2005 14:48
' Author    : RSanders
' Purpose   : This function will build the sql string for the update and insert
'             functions for a numeric field.
'---------------------------------------------------------------------------------------
'
Private Function fnCheckNumb(pField As String, pOld As String, pNew As String, Optional pLength As Integer = 8, Optional pPrecision As Integer = 0, Optional pNull As String = "") As Boolean
    Dim sCurrent As String

    strInsertName = ", " & pField
    strInsertValue = ", " & pNew
    
    If Not fnNumbsAreEqual(pOld, pNew) Then
        
        sCurrent = Trim(pNull)
        If Not rsCurrent Is Nothing Then
            If pNull = "" Or fnGetField(rsCurrent.Fields(pField)) <> "" Then
                sCurrent = tfnRound(rsCurrent.Fields(pField), pPrecision)
            End If
        End If
    
        If Not fnNumbsAreEqual(pOld, sCurrent) Then
            blnDataChanged = True
            Debug.Print pField & "   " & pOld & "   " & sCurrent & "   " & pNew
        Else
            strUpdate = strUpdate & ", " & pField & " = " & pNew
        End If
        
    End If
    
End Function

Private Function fnNumbsAreEqual(pNumb1 As String, pNumb2 As String) As Boolean
    Dim sNumb1 As String
    Dim sNumb2 As String
    Dim dNumb1 As Double
    Dim dNumb2 As Double
    
    fnNumbsAreEqual = False

    sNumb1 = fnRemoveQuotes(pNumb1)
    sNumb2 = fnRemoveQuotes(pNumb2)
    
    If sNumb1 = sNumb2 Then
        fnNumbsAreEqual = True
    ElseIf (sNumb1 = "" Or UCase(sNumb1) = "NULL") And (sNumb2 = "" Or UCase(sNumb2) = "NULL") Then
        fnNumbsAreEqual = True
    Else
        If IsNumeric(sNumb1) And IsNumeric(sNumb2) Then
            dNumb1 = CDbl(sNumb1)
            dNumb2 = CDbl(sNumb2)
            If dNumb1 = dNumb2 Then
                fnNumbsAreEqual = True
            End If
        End If
    End If
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : fnCheckDate
' DateTime  : 5/13/2005 14:48
' Author    : RSanders
' Purpose   : This function will build the sql string for the update and insert
'             functions for a date field.
'---------------------------------------------------------------------------------------
'
Private Function fnCheckDate(pField As String, pOld As String, pNew As String, Optional pNull As String = "") As Boolean
    Dim sCurrent As String
    
    strInsertName = ", " & pField
    strInsertValue = ", " & pNew
    
    If Not fnDatesAreEqual(pOld, pNew) Then
    
        sCurrent = Trim(pNull)
        If Not rsCurrent Is Nothing Then
            If pNull = "" Or fnGetField(rsCurrent.Fields(pField)) <> "" Then
                sCurrent = tfnFormatDate(rsCurrent.Fields(pField))
                sCurrent = tfnDateString(sCurrent, True)
            End If
        End If
    
        If Not fnDatesAreEqual(pOld, sCurrent) Then
            blnDataChanged = True
            Debug.Print pField & "   " & pOld & "   " & sCurrent & "   " & pNew
        Else
            strUpdate = strUpdate & ", " & pField & " = " & pNew
        End If
        
    End If
        
End Function

Private Function fnDatesAreEqual(pDate1 As String, pDate2 As String) As Boolean
    Dim sDate1 As String
    Dim sDate2 As String
    Dim dDate1 As Date
    Dim dDate2 As Date
    
    fnDatesAreEqual = False

    sDate1 = fnRemoveQuotes(pDate1)
    sDate2 = fnRemoveQuotes(pDate2)
    
    If (sDate1 = "" Or UCase(sDate1) = "NULL") And (sDate2 = "" Or UCase(sDate2) = "NULL") Then
        fnDatesAreEqual = True
    Else
        If IsDate(sDate1) And IsDate(sDate2) Then
            If DateDiff("d", sDate1, sDate2) = 0 Then
                fnDatesAreEqual = True
            End If
        End If
    End If
    
End Function

Private Function fnRemoveQuotes(pString As String) As String
    Dim sString As String
    fnRemoveQuotes = ""
    sString = Trim(pString)
    If sString <> "'" Or sString <> "''" Then
        If Len(sString) > 2 And Left(sString, 1) = "'" And Right(sString, 1) = "'" Then
            fnRemoveQuotes = Mid(sString, 2, Len(sString) - 2)
        End If
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : fnCheckForNull
' DateTime  : 5/13/2005 14:48
' Author    : RSanders
' Purpose   : This function will check to see if the current value is a null.  Null
'             character strings will contain the word NULL without the leading and
'             trailing single quote.  Note that this function will also add the leading
'             and trailing single quotes if they are missing.  For numeric fields the
'             quotes will be stripped at a later time.
'---------------------------------------------------------------------------------------
'
Private Function fnCheckForNull(pValue As String) As String

    Dim sValue As String
    
    sValue = Trim(pValue)
            
    If UCase(sValue) = "NULL" Then
        sValue = "NULL"
    Else
        If Len(sValue) > 1 Then
            If Left(sValue, 1) <> "'" And Right(sValue, 1) <> "'" Then
                sValue = "'" & sValue & "'"
            End If
        Else
            sValue = "'" & sValue & "'"
        End If
    End If

    fnCheckForNull = sValue
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : fnGetField
' DateTime  : 5/13/2005 14:48
' Author    : RSanders
' Purpose   : This function will check to see if the specified field is a null and if
'             so it will return a "null string".
'---------------------------------------------------------------------------------------
'
Private Function fnGetField(fld As Variant) As String
    If IsNull(fld) Then
        fnGetField = ""
    Else
        fnGetField = Trim(CStr(fld))
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : fnSQLString
' DateTime  : 5/13/2005 15:20
' Author    : RSanders
' Purpose   : This routine has been copied from another source.
'---------------------------------------------------------------------------------------
'
Private Function fnSQLString(ByVal vTemp As Variant, Optional vNoQuotes As Variant) As String
'
' Properly quotes and formats an SQL string.  If vNoQuotes is present, the result WILL NOT BE QUOTED
' for each ' character found, insert a double ''.  Leave "%* alone
    
    Dim nIdx As Integer
    Dim nPos As Integer
    Dim szParameter As String
    
    If IsNull(vTemp) Then
        fnSQLString = "NULL"
        Exit Function
    Else
        If vTemp = "" Then
            fnSQLString = "NULL"
            Exit Function
        End If
        
        szParameter = vTemp
    End If

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

'---------------------------------------------------------------------------------------
' Procedure : fnParseSQL
' DateTime  : 5/13/2005 14:48
' Author    : RSanders
' Purpose   : This routine is designed to parse the sql string passed to the fnUpdateRecord
'             and fnInsertRecord functions into the OLD_VALUE and NEW_VALUE arrays.
'---------------------------------------------------------------------------------------
'
Private Function fnParseSQL(pSQL As String) As String
    Dim i As Integer
    Dim j As Integer
    Dim strSQL As String
    
    fnParseSQL = ""
    
    i = InStr(1, pSQL, "(")
    If Not i > 0 Then
        status_message = OpenParenMsg
        Exit Function
    End If
    
    j = InStrRev(pSQL, ")")
    If Not j > 0 Then
        status_message = CloseParenMsg
        Exit Function
    End If
    
    i = i + 1
    strSQL = Trim(Mid(pSQL, i, j - i))
    j = 1

    For i = 0 To UBound(arrFieldName) - 1
        If arrFieldName(i) = "" Then
            Exit For
        End If
        arrFieldValueOld(i) = fnParseField(strSQL, j)
    Next
    
    For i = 0 To UBound(arrFieldName) - 1
        If arrFieldName(i) = "" Then
            Exit For
        End If
        arrFieldValueNew(i) = fnParseField(strSQL, j)
    Next
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : fnParseField
' DateTime  : 5/13/2005 14:48
' Author    : RSanders
' Purpose   : This function is designed to parse a single field from a sql string
'---------------------------------------------------------------------------------------
'
Private Function fnParseField(pSQL As String, ByRef i As Integer) As String
    Dim x As Integer

    If i > Len(pSQL) Then
        fnParseField = ""
        Exit Function
    End If
    
    ' Skip any spaces
    For i = i To Len(pSQL)
        If Mid(pSQL, i, 1) <> " " Then
            Exit For
        End If
    Next
    
    ' Is the field delimited with a single quote?
    If Mid(pSQL, i, 1) = "'" Then
    
        ' yes, set the x index to one beyond the single quote
        x = i + 1
        
        ' Loop till we are at the end or we find a closing quote.
        Do While x < Len(pSQL)
        
            ' Look for the next single quote
            x = InStr(x, pSQL, "'")
            
            ' If its missing or we are at the end then get out
            If x = 0 Then
                x = Len(pSQL)
            End If
            
            If x < Len(pSQL) Then
            
                ' if we don't have a case of back to back single quotes then get out
                If Mid(pSQL, x, 2) <> "''" Then
                    Exit Do
                End If
                
                ' go to the character beyond the back to back single quotes quotes
                x = x + 2

            End If
            
        Loop

        ' Did we end with a single quote?  If not then add one.
        If Mid(pSQL, x, 1) <> "'" Then
            pSQL = pSQL & "'"
            x = x + 1
        End If
        
        ' This puts us one beyone the end of the field
        x = x + 1
    
    Else
        
        ' Look for the comma seperating fields
        x = InStr(i, pSQL, ",")
        
        ' If we don't have one then point to the end + 1
        If x = 0 Then
            x = Len(pSQL) + 1
        End If
        
    End If

    ' i = start of field, x = one beyond the end
    fnParseField = Trim(Mid(pSQL, i, x - i))
    
    If x < Len(pSQL) Then
        i = InStr(x, pSQL, ",") + 1
        If i = 1 Then
            i = Len(pSQL)
        End If
    End If

End Function

'---------------------------------------------------------------------------------------
' Procedure : fnFindFieldName
' DateTime  : 5/13/2005 14:48
' Author    : RSanders
' Purpose   : This function will return the index into the arrFieldName array for
'             a specific field name.  If the name cannot be found then it will point
'             to the last entry of the array where the arrFieldType array can be
'             checked for TypeUndefined to determine that this is in fact the end.
'---------------------------------------------------------------------------------------
'
Private Function fnFindFieldName(pFieldName As String) As Integer
    Dim i As Integer
    For i = 0 To UBound(arrFieldName) - 1
        If arrFieldType(i) = TypeUndefined Then
            Exit For
        End If
        If StrComp(pFieldName, arrFieldName(i), vbTextCompare) = 0 Then
            Exit For
        End If
    Next
    fnFindFieldName = i
End Function

'---------------------------------------------------------------------------------------
' Procedure : fnWhereClause
' DateTime  : 5/13/2005 14:48
' Author    : RSanders
' Purpose   : This function will build the WHERE clause for a sql statement based on
'             the data specified for the CREATE_INDEX function when the object was
'             initially set up.
'---------------------------------------------------------------------------------------
'
Private Function fnWhereClause() As String

    Dim i           As Integer
    Dim strWhere    As String
    Dim strDate     As String
    Dim strNumber   As String
    Dim strChar     As String
    
    fnWhereClause = ""
    strWhere = ""

    For i = 0 To UBound(arrFieldName) - 1
        If arrFieldType(i) = TypeUndefined Then
            Exit For
        End If
        If arrIndex(i) = True Then
            Select Case arrFieldType(i)
                Case TypeChar
                    strWhere = strWhere & " AND " & arrFieldName(i) & " = " & arrFieldValueOld(i)
                Case TypeDecimal, TypeInteger, TypeSmallInt
                    strWhere = strWhere & " AND " & arrFieldName(i) & " = " & tfnRound(arrFieldValueOld(i), arrFieldPrecision(i))
                Case TypeDate
                    strDate = tfnFormatDate(arrFieldValueOld(i))
                    strWhere = strWhere & " AND " & arrFieldName(i) & " = " & tfnDateString(strDate, True)
            End Select
        End If
    Next
    
    If strWhere <> "" Then
        fnWhereClause = " WHERE " & Mid(strWhere, 6)
    End If
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : fnSaveOldAndNew
' DateTime  : 5/13/2005 14:48
' Author    : RSanders
' Purpose   : This function will save the fields read into the old and new arrays.
'---------------------------------------------------------------------------------------
'
Private Function fnSaveOldAndNew(pCurrent As Recordset) As Long
    
    Dim i As Integer
    Dim sCurrent As String
    
    fnSaveOldAndNew = 0
    If Not pCurrent Is Nothing Then
        If pCurrent.RecordCount > 0 Then
        
            pCurrent.MoveLast
            pCurrent.MoveFirst
            
            For i = 0 To UBound(arrFieldName) - 1
                sCurrent = Trim(arrFieldNull(i))
                Select Case arrFieldType(i)
                    Case TypeChar
                        If arrFieldNull(i) = "" Or Trim(fnGetField(pCurrent.Fields(arrFieldName(i)))) <> "" Then
                            sCurrent = fnGetField(pCurrent.Fields(arrFieldName(i)))
                            sCurrent = fnSQLString(sCurrent)
                            arrFieldValueOld(i) = sCurrent
                        End If
                    Case TypeDecimal, TypeInteger, TypeSmallInt
                        If arrFieldNull(i) = "" Or fnGetField(pCurrent.Fields(arrFieldName(i))) <> "" Then
                            sCurrent = tfnRound(pCurrent.Fields(arrFieldName(i)), arrFieldPrecision(i))
                            arrFieldValueOld(i) = sCurrent
                        End If
                    Case TypeDate
                        If arrFieldNull(i) = "" Or fnGetField(pCurrent.Fields(arrFieldName(i))) <> "" Then
                            sCurrent = tfnFormatDate(pCurrent.Fields(arrFieldName(i)))
                            sCurrent = tfnDateString(sCurrent, True)
                            arrFieldValueOld(i) = sCurrent
                        End If
                    Case TypeUndefined
                        Exit For
                End Select
            Next
            
        End If
        fnSaveOldAndNew = pCurrent.RecordCount
    End If
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' DateTime  : 5/13/2005 14:48
' Author    : RSanders
' Purpose   : Just your basic class initialization
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    arrFieldType(0) = TypeUndefined
    arrIndex(0) = False
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Class_Terminate
' DateTime  : 5/13/2005 14:48
' Author    : RSanders
' Purpose   : And this kills it all.
'---------------------------------------------------------------------------------------
'
Private Sub Class_Terminate()
    Set rsCurrent = Nothing
    Set dbDatabase = Nothing
End Sub

