VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLockLocalDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    
Const szPERCENT As String = "%"

Private m_frmFormMain As Form
Private m_sFormID As String

Private m_lLockedSerial As Long
'

Public Sub Initialize(formMain As Form, sID As String)
    Set m_frmFormMain = formMain
    
    If sID = "" Then
        MsgBox "Main Program ID is required.", vbExclamation
    Else
        m_sFormID = sID
    End If
    
    subDeleteOldLocks
    
    m_lLockedSerial = 0
End Sub

Private Sub Class_Terminate()
    m_sFormID = ""
    Set m_frmFormMain = Nothing
End Sub

Public Function LockLocalTable(ByRef bStealTheLock As Boolean) As Boolean
    Const SUB_NAME As String = "LockLocalTable"
    
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    Dim lHWnd As Long
    Dim sUser As String
    Dim sCriteria As String
    Dim sDateTime As String
    Dim sUserList As String
    
    On Error GoTo errTrap
    
    m_lLockedSerial = 0
    bStealTheLock = False
    
    If m_frmFormMain Is Nothing Then
        MsgBox "Main Form not set.", vbExclamation
        Exit Function
    End If
    
    If m_sFormID = "" Then
        MsgBox "Main Program ID is required.", vbExclamation
        Exit Function
    End If
    
    lHWnd = m_frmFormMain.hwnd
    sUser = UserName()
    sCriteria = m_sFormID + "-" + ComputerName + "[" + tfnGetDataSourceName + "]"
    sDateTime = Format(Now, "yyyymmdd-hhmmss")
    
    'get/check the lock
    strSQL = "SELECT * FROM sys_row_lock"
    strSQL = strSQL + " WHERE srl_criteria LIKE " + tfnSQLString(sCriteria + szPERCENT)
    
    If fnGetRecord(rsTemp, strSQL, SUB_NAME, False) < 0 Then
        GoTo errTrap
    End If
    
    If rsTemp.RecordCount > 1 Then
        'ERROR!!! local table has already been locked by more than one user
        Beep
        Beep
        Beep
        
        While Not rsTemp.EOF
            sUserList = sUserList + fnCStr(rsTemp!srl_user_id) + "[" + fnCStr(rsTemp!srl_table) + "]" + vbCrLf
            rsTemp.MoveNext
        Wend
        
        MsgBox "Local Database is currently in use by more than one user:" _
            + vbCrLf + vbCrLf + sUserList + vbCrLf + _
            "Please report this message to Support.", vbCritical
        Exit Function
    ElseIf rsTemp.RecordCount > 0 Then
        'local table has already been locked
        Beep
        Beep
        Beep
        
        m_lLockedSerial = tfnRound(rsTemp!srl_nbr)
        
        If MsgBox("Local Database is currently in use by " + fnCStr(rsTemp!srl_user_id) + "[" + _
            fnCStr(rsTemp!srl_table) + "]." + vbCrLf + vbCrLf + "This may be caused by a Data" + _
            " Source Name has already been use to login to the Factor Menu and run this program." + _
            " Choose 'No', then exit from the Factor Menu and login again using different Data Source" + _
            " Name." + vbCrLf + vbCrLf + "Do you want to continue?", _
            vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
                Beep
                Beep
                Beep
                
                If MsgBox("WARNING:" + vbCrLf + vbCrLf + "IF PROCEED, ANOTHER INSTANCE OF THE PROGRAM WILL BE TERMINATED! " + _
                   "ANY CHANGES MADE IN THAT PROGRAM WILL BE LOST." + vbCrLf + vbCrLf + "Are you sure you want to continue?", _
                   vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
                        m_lLockedSerial = 0
                        Exit Function
                End If
        Else
            m_lLockedSerial = 0
            Exit Function
        End If
        
        'steal the lock
        bStealTheLock = True
        
        If Not objErrHandler Is Nothing Then
            tfnErrHandler SUB_NAME, -1, "Local Database is currently in use by user " + _
                "[" + fnCStr(rsTemp!srl_table) + "] - PROCEEDED.", False
        End If
        
        strSQL = "UPDATE sys_row_lock SET srl_table = " + tfnSQLString(lHWnd)
        strSQL = strSQL + ", srl_prog_id = " + tfnSQLString(m_sFormID)
        strSQL = strSQL + ", srl_user_id = " + tfnSQLString(sUser)
        strSQL = strSQL + ", srl_criteria = " + tfnSQLString(sCriteria + "-" + sDateTime)
        strSQL = strSQL + " WHERE srl_nbr = " & m_lLockedSerial

        If Not fnExecuteSQL(strSQL, SUB_NAME, False) Then
            GoTo errTrap
        End If
    
        '**********************
        ' STEAL THE LOCK!!!
        '**********************
        LockLocalTable = True
    Else
        'insert the lock entry
        strSQL = "INSERT INTO sys_row_lock (srl_table,srl_prog_id,srl_user_id,"
        strSQL = strSQL + "srl_criteria,srl_nbr) VALUES ("
        strSQL = strSQL + tfnSQLString(lHWnd) + ","
        strSQL = strSQL + tfnSQLString(m_sFormID) + ","
        strSQL = strSQL + tfnSQLString(sUser) + ","
        strSQL = strSQL + tfnSQLString(sCriteria + "-" + sDateTime) + ","
        strSQL = strSQL + "0" + ")"

        If Not fnExecuteSQL(strSQL, SUB_NAME, False) Then
            GoTo errTrap
        End If
    
        'get/check the lock again
        strSQL = "SELECT * FROM sys_row_lock"
        strSQL = strSQL + " WHERE srl_criteria LIKE " + tfnSQLString(sCriteria + szPERCENT)
        strSQL = strSQL + " ORDER BY srl_nbr"
        
        If fnGetRecord(rsTemp, strSQL, SUB_NAME, False) < 0 Then
            GoTo errTrap
        End If
        
        If rsTemp.RecordCount = 0 Then
            'SOMETHING WRONG???!!!
            MsgBox "Failed to lock Local Database." + vbCrLf + vbCrLf + "Please report this message to Support.", vbCritical
            Exit Function
        Else
            m_lLockedSerial = tfnRound(rsTemp!srl_nbr)
            
            If tfnRound(rsTemp!srl_table) <> lHWnd Then
                'locked failed, another user locked it first
                MsgBox "Local Database is currently in use by " + fnCStr(rsTemp!srl_user_id) + _
                    "[" + fnCStr(rsTemp!srl_table) + "].", vbCritical
                
                'delete the bad lock(s)
                strSQL = "DELETE FROM sys_row_lock"
                strSQL = strSQL + " WHERE srl_criteria LIKE " + tfnSQLString(sCriteria + szPERCENT)
                strSQL = strSQL + " AND srl_nbr <> " & m_lLockedSerial
        
                If Not fnExecuteSQL(strSQL, SUB_NAME, False) Then
                    GoTo errTrap
                End If
                
                m_lLockedSerial = 0
                Exit Function
            Else
                '**********************
                ' GOT THE LOCK!!!
                '**********************
                LockLocalTable = True
            End If
        End If
    End If
    
    Exit Function
    
errTrap:
    'SOMETHING WRONG???!!!
    MsgBox "Failed to lock Local Database." + vbCrLf + vbCrLf + "Please report this message to Support.", vbCritical
    m_lLockedSerial = 0
End Function

Public Sub UnlockLocalTable()
    Const SUB_NAME As String = "LockLocalTable"
    
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    Dim lHWnd As Long
    Dim sUser As String
    Dim sCriteria As String
    Dim sUserList As String
    
    On Error GoTo errTrap
    
    If m_lLockedSerial <> 0 Then
        'delete the lock
        strSQL = "DELETE FROM sys_row_lock"
        strSQL = strSQL + " WHERE srl_nbr = " & m_lLockedSerial

        If Not fnExecuteSQL(strSQL, SUB_NAME, False) Then
            GoTo errTrap
        End If
        
        m_lLockedSerial = 0
    End If

    Exit Sub
    
errTrap:
    'SOMETHING WRONG???!!!
    MsgBox "Failed to unlock Local Database." + vbCrLf + vbCrLf + "Please report this message to Support.", vbCritical
End Sub

Public Function CheckLockLocalTable() As Boolean
    Const SUB_NAME As String = "CheckLockLocalTable"
    
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    On Error GoTo errTrap
    
    If m_frmFormMain Is Nothing Then
        MsgBox "Main Form not set.", vbExclamation
        Exit Function
    End If
    
    If m_sFormID = "" Then
        MsgBox "Main Program ID is required.", vbExclamation
        Exit Function
    End If
    
    If m_lLockedSerial = 0 Then
        'no lock
        Exit Function
    End If

    'get/check the lock
    strSQL = "SELECT * FROM sys_row_lock"
    strSQL = strSQL + " WHERE srl_nbr = " & m_lLockedSerial
    
    If fnGetRecord(rsTemp, strSQL, SUB_NAME, False) < 0 Then
        GoTo errTrap
    End If
    
    If rsTemp.RecordCount = 0 Then
        Exit Function
    End If
    
    If tfnRound(rsTemp!srl_table) <> m_frmFormMain.hwnd Then
        '***************************************
        ' LOCK IS GONE!!! STEALED BY SOMEBODY!!!
        '***************************************
        MsgBox "Local Database Lock is lost! " + vbCrLf + vbCrLf + "Local Database is now in use by " + _
            fnCStr(rsTemp!srl_user_id) + "[" + fnCStr(rsTemp!srl_table) + "].", vbCritical
        
        m_lLockedSerial = 0
        Exit Function
    Else
        '************************
        ' STILL HAVE THE LOCK!!!
        '************************
        CheckLockLocalTable = True
    End If
    
    Exit Function
    
errTrap:
    'SOMETHING WRONG???!!!
    MsgBox "Failed to check the lock." + vbCrLf + vbCrLf + "Please report this message to Support.", vbCritical
End Function

Private Sub subDeleteOldLocks()
    Const SUB_NAME As String = "LockLocalTable"
    
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    Dim sCriteria As String
    Dim sDate As String
    Dim sTemp As String
    Dim nPosi As Integer
    
    Dim sSerialList As String
    
    On Error GoTo errTrap
    
    sCriteria = m_sFormID + "-" + ComputerName + "[" + tfnGetDataSourceName + "]"
    sDate = Format(Date, "yyyymmdd")
    
    'get/check the lock
    strSQL = "SELECT srl_nbr, srl_criteria FROM sys_row_lock"
    strSQL = strSQL + " WHERE srl_criteria LIKE " + tfnSQLString(sCriteria + szPERCENT)
    
    If fnGetRecord(rsTemp, strSQL, SUB_NAME, False) < 0 Then
        GoTo errTrap
    End If
    
    sSerialList = ""
    
    While Not rsTemp.EOF
        'get the date
        sTemp = fnCStr(rsTemp!srl_criteria)
        nPosi = InStrRev(sTemp, "]-")
        
        If nPosi > 0 Then
            sTemp = Left(Mid(sTemp, nPosi + 2), 8)
            
            If Len(sTemp) = 8 Then
                If sTemp < sDate Then
                    If sSerialList <> "" Then
                        sSerialList = sSerialList + ","
                    End If
                
                    sSerialList = sSerialList & tfnRound(rsTemp!srl_nbr)
                End If
            End If
        End If
        
        rsTemp.MoveNext
    Wend
    
    'delete the old lock
    If sSerialList <> "" Then
        strSQL = "DELETE FROM sys_row_lock"
        strSQL = strSQL + " WHERE srl_nbr IN (" + sSerialList + ")"

        If Not fnExecuteSQL(strSQL, SUB_NAME, False) Then
            GoTo errTrap
        End If
    End If
    
    Exit Sub
    
errTrap:
    'SOMETHING WRONG???!!!
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
    
    sName = UCase(fnGetField(Left(sName, nSize - 1)))
    
    UserName = Trim(sName)
End Function

Private Function fnCStr(sIn) As String
    fnCStr = Trim(sIn & "")
End Function
