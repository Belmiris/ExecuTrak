VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRowLock"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'***********************************************************'
'
' Copyright (c) 1996 FACTOR, A Division of W.R.Hess Company
'
' Module name   : ROWLOCK.CLS
' Date          : October 23, 1996
' Programmer(s) : Qinggang Ma
'
' This module implement a class for lock records in a table of the database.
'
'Usage:
'   1. Properties need to set up
'       a. Database:        Set Obj.Database = t_dbMainDatabase
'       b. Table:               Obj.Table = "gl_master" (The table to be maintained)
'       c. MyID:                Obj.MyID = "WSFMASTR"   (The module ID)
'                            or Obj.MyID = efraPanel.Caption
'   2. Functions:
'       a. OpenRecordForEdit(strSQL As String) As Recordset
'           input parameter strSQL is the SQL statement to open the recordset
'           Example:
'               SELECT * WHERE glm_series = 100
'                   AND glm_account = 1000 AND glm_prft_ctr = 0
'           returns the recordset opened if successful
'               otherwise returns nothing
'       b. UpdateRecord(strSQL As String) As Integer
'           input parameter strSQL is the set value part of the SQL.
'           For example: for the SQL given above, to update glm_amt_period0 to 100,
'               Only the following string needs to pass in
'                   glm_amt_period0 = 100
'           returns the number of rows affected. If return 0. means not successful.
'       c. UnlockRow
'           No parameter and no values returned.

Option Explicit
    Private Const MAX_CRIT_LEN = 80
    
    Private Const sErrorMessage1 = "Can not read record from database"
    Private Const sErrorMessage2 = "Can not lock row for edit"

    Private m_nHandle As Integer
    Private m_sTable As String
    Private m_dbDatabase As Object
    Private m_sProgID As String

Property Set Database(objDB As Database)
    Set m_dbDatabase = objDB
End Property

Private Function fnOpenRecord(strSQL As String, _
                              Optional vMsg As Variant, _
                              Optional vDB As Variant) As Recordset
    ' Get records from the given SQL statement
    Dim objDB As Database
    Dim rsTemp As Recordset

    If IsMissing(vDB) Then
        Set objDB = t_dbMainDatabase
    Else
        Set objDB = vDB
    End If
    On Error GoTo SQLError
    If objDB Is t_dbMainDatabase Then
        Set rsTemp = objDB.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    Else
        Set rsTemp = objDB.OpenRecordset(strSQL, dbOpenSnapshot)
        If rsTemp.RecordCount > 0 Then
            rsTemp.MoveLast
            rsTemp.MoveFirst
        End If
    End If
    If rsTemp.RecordCount > 0 Then
        Set fnOpenRecord = rsTemp
    Else
        Set fnOpenRecord = Nothing
    End If
quitsub:
    On Error GoTo 0
    Exit Function
SQLError:
    Dim sMsg As String
    If IsMissing(vMsg) Then
        #If DEVELOP Then
            sMsg = "An error occurred while doing the SQL query" & CRLF & CRLF & "Error# " & CStr(Err.Number) & CRLF & Err.Description
            sMsg = sMsg & CRLF & CRLF & strSQL
            Clipboard.SetText strSQL
        #Else
            sMsg = ""
        #End If
    Else
        sMsg = vMsg
        #If DEVELOP Then
            sMsg = sMsg & CRLF & CRLF & strSQL
            Clipboard.SetText strSQL
        #End If
    End If
    MsgBox sMsg, vbOKOnly + vbInformation, App.Title
    Set fnOpenRecord = Nothing
    Err.Clear
    Resume quitsub
End Function


Private Function fnExecuteSQL(strSQL As String, _
                              Optional sErrorMessage As Variant, _
                              Optional vDB As Variant) As Integer

    Dim objDB As Database
    
    If IsMissing(vDB) Then
        Set objDB = t_dbMainDatabase
    Else
        Set objDB = vDB
    End If
    On Error GoTo errExecute
    If objDB Is t_dbMainDatabase Then
        fnExecuteSQL = objDB.ExecuteSQL(strSQL)
    Else
        objDB.Execute strSQL
        fnExecuteSQL = 0
    End If

extExecute:
    On Error GoTo 0
    Exit Function

errExecute:
    Dim sMsg As String
    Dim bMsg As Boolean
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
    fnExecuteSQL = -1
    Resume extExecute

End Function


Private Function fnGetHandle() As Integer
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim nNbr As Integer
    
    strSQL = "SELECT MAX(srl_nbr) handle FROM sys_row_lock"
    Set rsTemp = fnOpenRecord(strSQL, sErrorMessage1, m_dbDatabase)
    If Not rsTemp Is Nothing Then
        If IsNull(rsTemp!Handle) Then
            nNbr = 0
        Else
            nNbr = rsTemp!Handle
        End If
        Do
            nNbr = nNbr + 1
            strSQL = "SELECT srl_nbr handle FROM sys_row_lock WHERE srl_nbr = " & CStr(nNbr)
            Set rsTemp = fnOpenRecord(strSQL, sErrorMessage1, m_dbDatabase)
        Loop Until rsTemp Is Nothing
    End If
    fnGetHandle = nNbr
End Function

Private Function fnGetUserID() As String
    Dim nPos As Integer
    Dim sConnect As String
    Dim nLen As Integer
    
    fnGetUserID = ""
    sConnect = UCase(m_dbDatabase.Connect)
    nPos = InStr(sConnect, "UID=")
    If nPos > 0 Then
        nLen = InStr(nPos, sConnect, ",")
        If nLen > nPos Then
            fnGetUserID = Mid(sConnect, nPos + 1, nLen - nPos)
        End If
    End If

End Function

Private Function fnLockRow(strSQL As String) As Boolean
    
    Const sWHERE = " WHERE "

    Dim sCrit As String
    Dim sSQL As String
    Dim rsTemp As Recordset
    Dim nPos As Integer
    Dim nNbr As Integer
    Dim sTemp As String
    Dim rsLock As Recordset
    Dim sUserID As String
    Dim bFlag As Boolean
    
    sUserID = fnGetUserID

    m_nHandle = fnGetHandle
    fnLockRow = False
    sSQL = UCase(strSQL)
    nPos = InStr(sSQL, sWHERE)
    If nPos = 0 Then
        #If DEVELOP Then
            MsgBox "The SQL statement needs to have WHERE clause to open record for update" & CRLF & CRLF & strSQL, vbOKOnly + vbCritical
        #End If
        Exit Function
    End If

    sCrit = Trim(Right(sSQL, Len(strSQL) - nPos - Len(sWHERE) + 1))

    nPos = Len(sCrit)
    nNbr = 1
    While nPos >= 0
        If nPos > MAX_CRIT_LEN Then
            nPos = MAX_CRIT_LEN
        End If
        sSQL = "INSERT INTO sys_row_lock VALUES(" & CStr(m_nHandle) & ", " _
               & CStr(nNbr) & ", '" & m_sTable & "', '" & sUserID & "', '" _
               & m_sProgID & "', '" & Left$(sCrit, nPos) & "')"
        If fnExecuteSQL(sSQL, sErrorMessage2, m_dbDatabase) < 0 Then
            Exit Function
        End If
        nPos = nPos - MAX_CRIT_LEN
        If nPos > 0 Then
            sCrit = Right(sCrit, nPos)
            nNbr = nNbr + 1
        End If
    Wend

    sSQL = "SELECT srl_nbr, srl_seq_nbr, srl_criteria FROM sys_row_lock WHERE srl_nbr <> " & CStr(m_nHandle) & " AND srl_table = '" & m_sTable & "' ORDER BY srl_nbr, srl_seq_nbr"
    Set rsTemp = fnOpenRecord(sSQL, sErrorMessage1, m_dbDatabase)
    If Not rsTemp Is Nothing Then
        rsTemp.MoveFirst
        While Not rsTemp.EOF
            nNbr = rsTemp!srl_nbr
            If Not IsNull(rsTemp!srl_criteria) Then
                sCrit = Trim(rsTemp!srl_criteria)
            End If
            Do
                rsTemp.MoveNext
                bFlag = rsTemp.EOF
                If Not bFlag Then
                    If nNbr <> rsTemp!srl_nbr Then
                        bFlag = True
                    End If
                    If Not IsNull(rsTemp!srl_criteria) Then
                        sCrit = sCrit & Trim(rsTemp!srl_criteria)
                    End If
                End If
            Loop Until bFlag

            sSQL = strSQL & " AND " & sCrit
            Set rsLock = fnOpenRecord(sSQL, sErrorMessage1, m_dbDatabase)
            If Not rsLock Is Nothing Then
                sSQL = "SELECT srl_user_id, srl_prog_id FROM sys_row_lock WHERE srl_nbr = " & CStr(nNbr)
                Set rsLock = fnOpenRecord(sSQL, sErrorMessage1, m_dbDatabase)
                If Not rsLock Is Nothing Then
                    If Not IsNull(rsLock!srl_user_id) Then
                        sTemp = " by " & rsLock!srl_user_id
                    Else
                        sTemp = ""
                    End If
                    MsgBox "The record is locked" & sTemp, vbOKOnly, App.Title
                End If
                UnLockRow
                fnLockRow = False
                Exit Function
            End If
        Wend
    End If
    
    fnLockRow = True

End Function


Property Let MyID(sID As String)
    m_sProgID = UCase(Trim(sID))
End Property

Public Function OpenRecordForEdit(strSQL As String) As Recordset
    'Input variables:
    '   strSQL: The SQL statement to open the recordset
    '   Return the recordset

    UnLockRow

    Set OpenRecordForEdit = Nothing
    If m_sTable = "" Then
        #If DEVELOP Then
            MsgBox "The maintaining table must be supplied", vbCritical + vbOKOnly
        #End If
        Exit Function
    End If
    If m_sProgID = "" Then
        #If DEVELOP Then
            MsgBox "The program ID must be supplied", vbCritical + vbOKOnly
        #End If
        Exit Function
    End If
    If m_dbDatabase Is Nothing Then
        #If DEVELOP Then
            MsgBox "The database must be supplied", vbCritical + vbOKOnly
        #End If
        Exit Function
    End If

    If fnLockRow(strSQL) Then
        Set OpenRecordForEdit = fnOpenRecord(strSQL, , m_dbDatabase)
    End If

End Function
Public Sub UnLockRow()
    'Unlock the locked row by deleting the record from sys_row_lock

    Dim sSQL As String

    If m_nHandle < 0 Then
        Exit Sub
    End If
    sSQL = "DELETE FROM sys_row_lock WHERE srl_nbr = " & CStr(m_nHandle)
    If fnExecuteSQL(sSQL, , m_dbDatabase) > 0 Then
        m_nHandle = -1
    End If

End Sub

Property Let Table(sTemp As String)
    m_sTable = UCase(Trim(sTemp))
End Property

Public Function UpdateRecord(strSQL As String) As Integer
    'Update the locked record and unlock the record if it is successful
    ' The SQL string need only the field and value part
    ' Return the number of rows affected. 0 means not successful

    If m_dbDatabase Is Nothing Then
        #If DEVELOP Then
            MsgBox "The database must be supplied", vbCritical + vbOKOnly
        #End If
        Exit Function
    End If
    
    Dim sSQL As String
    Dim rsTemp As Recordset
    Dim nRet As Integer
    Dim sCrit As String

    sSQL = "SELECT srl_seq_nbr, srl_criteria FROM sys_row_lock WHERE srl_nbr = " & CStr(m_nHandle) & " ORDER BY srl_seq_nbr"
    Set rsTemp = fnOpenRecord(sSQL, sErrorMessage1, m_dbDatabase)
    If rsTemp Is Nothing Then
        #If DEVELOP Then
            MsgBox "The record to be updated is not locked", vbOKOnly
        #End If
        UpdateRecord = 0
    Else
        rsTemp.MoveFirst
        sCrit = ""
        While Not rsTemp.EOF
            If Not IsNull(rsTemp!srl_criteria) Then
                sCrit = sCrit & rsTemp!srl_criteria
            End If
            rsTemp.MoveNext
        Wend
        sSQL = "UPDATE " & m_sTable & " SET " & strSQL & " WHERE " & sCrit
        nRet = fnExecuteSQL(sSQL, , m_dbDatabase)
        If nRet > 0 Then
            UnLockRow
        End If
        UpdateRecord = nRet
    End If
End Function

Private Sub Class_Initialize()
    m_nHandle = -1
End Sub


Private Sub Class_Terminate()
    UnLockRow
End Sub


