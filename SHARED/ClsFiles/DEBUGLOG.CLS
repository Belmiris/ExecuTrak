VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDebugLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private m_frmFormMain As Form
Private m_sFormID As String
'

Public Sub Initialize(formMain As Form, sID As String)
    Set m_frmFormMain = formMain
    
    If sID = "" Then
        MsgBox "Main Program ID is required.", vbExclamation
        m_sFormID = "NONE"
    Else
        m_sFormID = sID
    End If
End Sub

Private Sub Class_Terminate()
    m_sFormID = ""
    Set m_frmFormMain = Nothing
End Sub

'david 05/06/2003  #408907
Public Sub StartDebugLog()
    Dim sField As String
    Dim sValue As String
    
    #If Not DEBUG_LOG Then
        Exit Sub
    #End If
    
    If Not fnTableExists("sys_ini") Then
        Exit Sub
    End If
    
    If m_sFormID = "" Then
        MsgBox "Main Program ID is required.", vbExclamation
        Exit Sub
    End If
    
    sField = Format(Now, "yyyymmdd-hhmmss") + "-start[" & m_frmFormMain.hwnd & "]"
    sValue = " [" + ComputerName() + "] <" + UserName() + "> [" + tfnGetDataSourceName() + "]" + tfnGetUserName()
    
    WriteSysIni m_sFormID, "AUTO", "DEBUG_LOG", sField, sValue
End Sub

Public Sub StopDebugLog()
    Dim sField As String
    Dim sValue As String
    
    #If Not DEBUG_LOG Then
        Exit Sub
    #End If
    
    If Not fnTableExists("sys_ini") Then
        Exit Sub
    End If
    
    If m_sFormID = "" Then
        MsgBox "Main Program ID is required.", vbExclamation
        Exit Sub
    End If
    
    sField = Format(Now, "yyyymmdd-hhmmss") + "-stop[" & m_frmFormMain.hwnd & "]"
    sValue = " [" + ComputerName() + "] <" + UserName() + "> [" + tfnGetDataSourceName() + "]" + tfnGetUserName()
    
    WriteSysIni m_sFormID, "AUTO", "DEBUG_LOG", sField, sValue
End Sub

Public Sub LogMsg(sMsg As String, Optional bQueryFirst As Boolean = False)
    Dim sField As String
    
    If Not fnTableExists("sys_ini") Then
        Exit Sub
    End If
    
    sField = Format(Now, "yyyymmdd-hhmmss") + "[" & m_frmFormMain.hwnd & "]"
    
    WriteSysIni m_sFormID, "AUTO", "PROG_LOG", sField, " [" + ComputerName() + _
        "] <" + UserName() + "> " + sMsg, bQueryFirst, False
End Sub

Public Function fnTableExists(sTableName As String) As Boolean
    Const SUB_NAME As String = "fnTableExists"
    
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    strSQL = "SELECT * FROM systables WHERE tabname = " + tfnSQLString(sTableName)
    
    On Error GoTo errTrap
    Set rsTemp = t_dbMainDatabase.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    
    If rsTemp.RecordCount > 0 Then
        fnTableExists = True
    End If
    Exit Function
    
errTrap:
    If Not objErrHandler Is Nothing Then
        tfnErrHandler SUB_NAME, strSQL, True
    End If
End Function

'return error message if any
'will set the sValue upon return
Public Function ReadSysIni(sFilename As String, _
                              sUserID As String, _
                              sSECTION As String, _
                              sField As String, _
                              ByRef sValue As String, _
                              ByRef bExist As Boolean, _
                              Optional bShowError As Boolean = True) As String
                              
    Const SUB_NAME As String = "fnReadSysIni"
    
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    sValue = ""
    bExist = False
    
    strSQL = "SELECT ini_value FROM sys_ini WHERE ini_file_name = " + tfnSQLString(sFilename)
    strSQL = strSQL + " AND ini_user_id = " + tfnSQLString(sUserID)
    strSQL = strSQL + " AND ini_section = " + tfnSQLString(sSECTION)
    strSQL = strSQL + " AND ini_field_name = " + tfnSQLString(sField)
    
    On Error GoTo errTrap
    Set rsTemp = t_dbMainDatabase.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    
    If rsTemp.RecordCount > 0 Then
        sValue = CStr(rsTemp!ini_value & "")
        bExist = True
    End If
    Exit Function
    
errTrap:
    If Not objErrHandler Is Nothing Then
        tfnErrHandler SUB_NAME, strSQL, bShowError
    End If
    
    ReadSysIni = "Failed to read data from sys_ini"
End Function

Public Function WriteSysIni(sFilename As String, _
                              sUserID As String, _
                              sSECTION As String, _
                              sField As String, _
                              sValue As String, _
                              Optional bQueryFirst As Boolean = True, _
                              Optional bShowError As Boolean = True) As String

    Const SUB_NAME As String = "fnWriteSysIni"
    
    Dim strSQL As String
    Dim sData As String
    Dim bExist As Boolean
    
    bExist = False
    
    If bQueryFirst Then
        ReadSysIni sFilename, sUserID, sSECTION, sField, sData, bExist
    End If
    
    On Error GoTo errTrap
    
    If bExist Then
        strSQL = "UPDATE sys_ini SET ini_value = " + tfnSQLString(sValue)
        strSQL = strSQL + " WHERE ini_file_name = " + tfnSQLString(sFilename)
        strSQL = strSQL + " AND ini_user_id = " + tfnSQLString(sUserID)
        strSQL = strSQL + " AND ini_section = " + tfnSQLString(sSECTION)
        strSQL = strSQL + " AND ini_field_name = " + tfnSQLString(sField)
    Else
        strSQL = "INSERT INTO sys_ini (ini_file_name,ini_user_id,ini_section,"
        strSQL = strSQL + "ini_field_name,ini_value) VALUES ("
        strSQL = strSQL + tfnSQLString(sFilename) + ","
        strSQL = strSQL + tfnSQLString(sUserID) + ","
        strSQL = strSQL + tfnSQLString(sSECTION) + ","
        strSQL = strSQL + tfnSQLString(sField) + ","
        strSQL = strSQL + tfnSQLString(sValue) + ")"
    End If
    
    t_dbMainDatabase.ExecuteSQL strSQL
    Exit Function

errTrap:
    If Not objErrHandler Is Nothing Then
        tfnErrHandler SUB_NAME, strSQL, bShowError
    End If
    
    WriteSysIni = "Failed to write data to sys_ini"
End Function

Public Function ComputerName() As String
    Dim lngReturnCode As Long
    Dim strHostname As String
    strHostname = Space$(50)
    lngReturnCode = GetComputerName(strHostname, 50)
    ComputerName = Left$(strHostname, InStr(strHostname, Chr$(0)) - 1)
End Function

Public Function UserName() As String
    Dim nSize As Long
    Dim sName As String
    
    nSize = 128
    sName = Space(nSize)
    GetUserName sName, nSize
    
    sName = UCase(fnGetField(Left(sName, nSize - 1)))
    
    UserName = Trim(sName)
End Function
