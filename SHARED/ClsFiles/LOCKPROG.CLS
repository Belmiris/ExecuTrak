VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLockProgram"
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
    Else
        m_sFormID = sID
    End If
    
    subDeleteOldLocks
End Sub

Private Sub Class_Terminate()
    m_sFormID = ""
    Set m_frmFormMain = Nothing
End Sub

Public Function LockProgram(Optional sProgramID As String = "") As Boolean
    Const SUB_NAME As String = "LockProgram"
    
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    Dim sErrMsg As String
    Dim i As Integer
    
    Dim lHWnd As Long
    Dim sUser As String
    Dim sCriteria As String
    
    On Error GoTo errTrap
    
    If m_frmFormMain Is Nothing Then
        MsgBox "Main Form not set.", vbExclamation
        Exit Function
    End If
    
    If sProgramID = "" Then
        sProgramID = m_sFormID
    End If
    
    If sProgramID = "" Then
        MsgBox "Program ID is required.", vbExclamation
        Exit Function
    End If
    
    sErrMsg = "N/A"
    
    lHWnd = m_frmFormMain.hwnd
    sUser = UserName()
    
    If Len(sUser) > 12 Then
        sUser = Left(sUser, 12)
    End If
    
    sCriteria = "LOCKPROG-" + UCase(sProgramID) + "-" + ComputerName + "[" + tfnGetDataSourceName + "]"
    
    'get/check the lock
    strSQL = "SELECT * FROM sys_row_lock"
    strSQL = strSQL + " WHERE srl_criteria = " + tfnSQLString(sCriteria)
    strSQL = strSQL + " and srl_user_id = " + tfnSQLString(sUser)
    
    If fnGetRecord(rsTemp, strSQL, SUB_NAME, False) < 0 Then
        sErrMsg = "Failed to select from the lock table."
        GoTo errTrap
    End If
    
    sErrMsg = "N/A"
    
    If rsTemp.RecordCount > 0 Then
        strSQL = "UPDATE sys_row_lock SET srl_table = " + tfnSQLString(lHWnd)
        strSQL = strSQL + " WHERE srl_criteria = " + tfnSQLString(sCriteria)
        strSQL = strSQL + " and srl_user_id = " + tfnSQLString(sUser)

        If Not fnExecuteSQL(strSQL, SUB_NAME, False) Then
            sErrMsg = "Failed to update to the lock table."
            GoTo errTrap
        End If
        
        LockProgram = True
    Else
        'insert the lock entry
        strSQL = "INSERT INTO sys_row_lock (srl_table,srl_prog_id,srl_user_id,"
        strSQL = strSQL + "srl_criteria,srl_nbr) VALUES ("
        strSQL = strSQL + tfnSQLString(lHWnd) + ","
        strSQL = strSQL + tfnSQLString(sProgramID) + ","
        strSQL = strSQL + tfnSQLString(sUser) + ","
        strSQL = strSQL + tfnSQLString(sCriteria) + ","
        strSQL = strSQL + "0" + ")"

        If Not fnExecuteSQL(strSQL, SUB_NAME, False) Then
            sErrMsg = "Failed to insert into the lock table."
            GoTo errTrap
        End If
    
        'get/check the lock again
        strSQL = "SELECT * FROM sys_row_lock"
        strSQL = strSQL + " WHERE srl_criteria = " + tfnSQLString(sCriteria)
        strSQL = strSQL + " and srl_user_id = " + tfnSQLString(sUser)
        'strSQL = strSQL + " ORDER BY srl_nbr"
        
        'david 01/26/2005  #472512
        'timing issue
        'let's try 5 times to retrieve the lock entry,
        'wait 1 second between tries.
        For i = 1 To 5
            sErrMsg = "N/A"
            
            If fnGetRecord(rsTemp, strSQL, SUB_NAME) < 0 Then
                sErrMsg = "Failed to select again from the lock table."
                GoTo errTrap
            End If
            
            If rsTemp.RecordCount > 0 Then
                'inserted lock entries found
                LockProgram = True
                Exit For
            Else
                'SOMETHING WRONG???!!!
                'try again, loops
                tfnWaitSeconds 1
            End If
        Next i
        
        If i > 5 Then
            sErrMsg = "lock entry not found."
            GoTo errTrap
        End If
    End If
    
    Exit Function
    
errTrap:
    'SOMETHING WRONG???!!!
    MsgBox "Failed to lock the Program." + vbCrLf _
        + "ErrDesc: " + sErrMsg + vbCrLf + vbCrLf _
        + "Please report this message to Support.", vbCritical
End Function

Public Sub UnLockProgram(Optional sProgramID As String = "")
    Const SUB_NAME As String = "UnLockProgram"
    
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    Dim sUser As String
    Dim sCriteria As String
    
    On Error GoTo errTrap
    
    If sProgramID = "" Then
        sProgramID = m_sFormID
    End If
    
    If sProgramID = "" Then
        MsgBox "Program ID is required.", vbExclamation
        Exit Sub
    End If
    
    sUser = UserName()
    
    If Len(sUser) > 12 Then
        sUser = Left(sUser, 12)
    End If
    
    sCriteria = "LOCKPROG-" + UCase(sProgramID) + "-" + ComputerName + "[" + tfnGetDataSourceName + "]"
    
    'delete the program lock
    strSQL = "DELETE FROM sys_row_lock"
    strSQL = strSQL + " WHERE srl_criteria = " + tfnSQLString(sCriteria)
    strSQL = strSQL + " and srl_user_id = " + tfnSQLString(sUser)
        
    If Not fnExecuteSQL(strSQL, SUB_NAME, False) Then
        GoTo errTrap
    End If

    Exit Sub
    
errTrap:
    'SOMETHING WRONG???!!!
    MsgBox "Failed to unlock the Program." + vbCrLf + vbCrLf + "Please report this message to Support.", vbCritical
End Sub

'return True if the program is locked/running
Public Function IsProgramLocked(sProgramID As String, _
                                 Optional bShowMsgBox As Boolean = True) As Boolean
    Const SUB_NAME As String = "IsProgramLocked"
    
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    Dim sUser As String
    Dim sCriteria As String
    
    On Error GoTo errTrap
    
    If sProgramID = "" Then
        sProgramID = m_sFormID
    End If
    
    If sProgramID = "" Then
        MsgBox "Program ID is required.", vbExclamation
        Exit Function
    End If
    
    sUser = UserName()
    
    If Len(sUser) > 12 Then
        sUser = Left(sUser, 12)
    End If
    
    sCriteria = "LOCKPROG-" + UCase(sProgramID) + "-" + ComputerName + "[" + tfnGetDataSourceName + "]"

    'get/check the program lock
    strSQL = "SELECT * FROM sys_row_lock"
    strSQL = strSQL + " WHERE srl_criteria = " + tfnSQLString(sCriteria)
    strSQL = strSQL + " and srl_user_id = " + tfnSQLString(sUser)
    
    If fnGetRecord(rsTemp, strSQL, SUB_NAME, False) < 0 Then
        GoTo errTrap
    End If
    
    If rsTemp.RecordCount > 0 Then
        IsProgramLocked = True
    End If
    
    Exit Function
    
errTrap:
    'SOMETHING WRONG???!!!
    If bShowMsgBox Then
        MsgBox "Failed to check the lock." + vbCrLf + vbCrLf + "Please report this message to Support.", vbCritical
    End If
End Function

Private Sub subDeleteOldLocks()
    Const SUB_NAME As String = "LockProgram"
    
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    
    Dim sUser As String
    Dim sCriteria As String
    
    If m_sFormID = "" Then
        MsgBox "Program ID is required.", vbExclamation
        Exit Sub
    End If
    
    sUser = UserName()
    
    If Len(sUser) > 12 Then
        sUser = Left(sUser, 12)
    End If
    
    sCriteria = "LOCKPROG-" + UCase(m_sFormID) + "-" + ComputerName + "[" + tfnGetDataSourceName + "]"
    
    'delete the old lock
    strSQL = "DELETE FROM sys_row_lock"
    strSQL = strSQL + " WHERE srl_criteria = " + tfnSQLString(sCriteria)
    strSQL = strSQL + " and srl_user_id = " + tfnSQLString(sUser)
    
    If Not fnExecuteSQL(strSQL, SUB_NAME, False) Then
        MsgBox "Failed to delete old lock.", vbExclamation
    End If
End Sub

Private Function fnGetRecord(rsTemp As Recordset, strSQL As String, _
                            sCalledFrom As String, Optional bShowError As Boolean = True, _
                            Optional bMoveLast As Boolean = True) As Long
    ' Get records from the given SQL statement
    ' nDB = 1 ---> Informax Database (remote)
    '     = 2 ---> Access Database (local)

    On Error GoTo SQLError
    
    Set rsTemp = t_dbMainDatabase.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    
    If bMoveLast Then
        If rsTemp.RecordCount > 0 Then
            rsTemp.MoveLast
            rsTemp.MoveFirst
        End If
    End If
    
    fnGetRecord = rsTemp.RecordCount
    Exit Function

SQLError:
    If Not objErrHandler Is Nothing Then
        tfnErrHandler "fnGetRecord," + sCalledFrom, strSQL, bShowError
    End If
    
    fnGetRecord = -9999
End Function

Private Function fnExecuteSQL(strSQL As String, sCalledFrom As String, Optional bShowError As Boolean = True) As Boolean
    On Error GoTo ErrorAccessRecords
    
    t_dbMainDatabase.ExecuteSQL strSQL
    
    fnExecuteSQL = True
    Exit Function

ErrorAccessRecords:
    If Not objErrHandler Is Nothing Then
        tfnErrHandler "fnExecuteSQL," + sCalledFrom, strSQL, bShowError
    End If
    
    fnExecuteSQL = False
End Function

Private Function ComputerName() As String
    Dim lngReturnCode As Long
    Dim strHostname As String
    strHostname = Space$(50)
    lngReturnCode = GetComputerName(strHostname, 50)
    ComputerName = Left$(strHostname, InStr(strHostname, Chr$(0)) - 1)
End Function

Private Function UserName() As String
    Dim nSize As Long
    Dim sName As String
    
    nSize = 128
    sName = Space(nSize)
    GetUserName sName, nSize
    
    sName = UCase(fnCStr(Left(sName, nSize - 1)))
    
    UserName = Trim(sName)
End Function

Private Function fnCStr(sIn) As String
    fnCStr = Trim(sIn & "")
End Function
