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
Private Declare Sub PokeLong Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, Optional ByVal length As Long = 4)

Private colPrograms As New Collection

Public Function ProgramAdd(sProg As String, sDesc As String)
    colPrograms.Add sDesc, sProg
End Function

Public Function ProgramCount() As Long
   ProgramCount = colPrograms.Count
End Function

Public Sub ProgramDelete(sProg As String)
   colPrograms.Remove sProg
End Sub

Public Function ProgramDesc(sProg As String) As String
   ProgramDesc = colPrograms.Item(sProg)
End Function


Public Sub Initialize(formMain As Form, sID As String, _
                      Optional bUseFactMenuUserID As Boolean = False)
    Set m_frmFormMain = formMain
    
    If sID = "" Then
        MsgBox "Main Program ID is required.", vbExclamation
    Else
        m_sFormID = sID
    End If
    
    subDeleteOldLocks bUseFactMenuUserID
End Sub

Private Sub Class_Terminate()
    Dim num As Long
    
    m_sFormID = ""
    Set m_frmFormMain = Nothing
    
   If Not colPrograms Is Nothing Then
        For num = 1 To colPrograms.Count
            colPrograms.Remove 1
        Next
        
        Set colPrograms = Nothing
    End If
End Sub

Public Function LockProgram(Optional sProgramID As String = "", _
                            Optional bIgnoreComputerName As Boolean = False, _
                            Optional bUseFactMenuUserID As Boolean = False) As Boolean
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
    
    lHWnd = m_frmFormMain.hWnd
    
    If bUseFactMenuUserID Then
        sUser = tfnGetUserName()
    Else
        sUser = UserName()
    End If
    
    If Len(sUser) > 12 Then
        sUser = Left(sUser, 12)
    End If
    
    If bIgnoreComputerName Then
        sCriteria = "LOCKPROG-" + UCase(sProgramID)
    Else
        sCriteria = "LOCKPROG-" + UCase(sProgramID) + "-" + ComputerName + "[" + tfnGetDataSourceName + "]"
    End If
    
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

Public Sub UnLockProgram(Optional sProgramID As String = "", _
                         Optional bIgnoreComputerName As Boolean = False, _
                         Optional bUseFactMenuUserID As Boolean = False)
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
    
    
    If bUseFactMenuUserID Then
        sUser = tfnGetUserName()
    Else
        sUser = UserName()
    End If
    
    If Len(sUser) > 12 Then
        sUser = Left(sUser, 12)
    End If
    
    If bIgnoreComputerName Then
        sCriteria = "LOCKPROG-" + UCase(sProgramID)
    Else
        sCriteria = "LOCKPROG-" + UCase(sProgramID) + "-" + ComputerName + "[" + tfnGetDataSourceName + "]"
    End If
    
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
                                 Optional bShowMsgBox As Boolean = True, _
                                 Optional bIgnoreComputerName As Boolean = False, _
                                 Optional bIgnoreUserName As Boolean = False, _
                                 Optional ByRef sOtherUser As String, _
                                 Optional bUseFactMenuUserID As Boolean = False) As Boolean
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
    
    If bUseFactMenuUserID Then
        sUser = tfnGetUserName()
    Else
        sUser = UserName()
    End If
    
    If Len(sUser) > 12 Then
        sUser = Left(sUser, 12)
    End If
    
    If bIgnoreComputerName Then
        sCriteria = "LOCKPROG-" + UCase(sProgramID)
    Else
        sCriteria = "LOCKPROG-" + UCase(sProgramID) + "-" + ComputerName + "[" + tfnGetDataSourceName + "]"
    End If
    'get/check the program lock
    strSQL = "SELECT * FROM sys_row_lock"
    strSQL = strSQL + " WHERE srl_criteria = " + tfnSQLString(sCriteria)
    If Not bIgnoreUserName Then
        strSQL = strSQL + " and srl_user_id = " + tfnSQLString(sUser)
    End If
    
    If fnGetRecord(rsTemp, strSQL, SUB_NAME, False) < 0 Then
        GoTo errTrap
    End If
    
    If rsTemp.RecordCount > 0 Then
        sOtherUser = GetField(rsTemp!srl_user_id)
        IsProgramLocked = True
    End If
    
    Exit Function
    
errTrap:
    'SOMETHING WRONG???!!!
    If bShowMsgBox Then
        MsgBox "Failed to check the lock." + vbCrLf + vbCrLf + "Please report this message to Support.", vbCritical
    End If
End Function

Private Sub subDeleteOldLocks(Optional bUseFactMenuUserID As Boolean = False)
    Const SUB_NAME As String = "LockProgram"
    
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    
    Dim sUser As String
    Dim sCriteria As String
    
    If m_sFormID = "" Then
        MsgBox "Program ID is required.", vbExclamation
        Exit Sub
    End If
    
    If bUseFactMenuUserID Then
        sUser = tfnGetUserName()
    Else
        sUser = UserName()
    End If
    
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

Public Function ExclusiveProgram(Optional vProgramIsRunning) As Boolean
    
    Const SUB_NAME = "Program"
    
    Dim nRet As Integer
    Dim i As Integer
    Dim sProg As String
    Dim sUser As String
        
    nRet = vbNo
    

    For i = 1 To ProgramCount()
        sProg = ItemKey(i, colPrograms)
        If IsProgramLocked(sProg, , True, True, sUser, True) Then
            MsgBox "The " & colPrograms(i) & " (" & sProg & ") program is being run by " & sUser & "." _
                + " Cannot run at the same time.", vbInformation + vbOKOnly
            nRet = vbAbort
        End If
    Next
    
    If nRet = vbAbort Then
        If IsMissing(vProgramIsRunning) Then
            Exit Function
        Else
            vProgramIsRunning = True
            ExclusiveProgram = True
            Exit Function
        End If
    End If
    'end check program running
    
    If Not LockProgram(, True, True) Then
        Exit Function
    End If
    
    ExclusiveProgram = True

End Function


Private Function ItemKey(ByVal Index As Long, Coll As Collection) As String
  'optimized get collection sKey by index
  'Private Declare Sub PokeLong Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, Optional ByVal Length As Long = 4)
  Dim i     As Long
  Dim Ptr   As Long
  Dim sKey  As String

   If Coll Is Nothing Then                             'oops!
       Err.Raise 91                                    'No object
     Else 'NOT COLL...
       Select Case Index
         Case Is < 1, Is > Coll.Count                  'oops!
           Err.Raise 9                                 'Index out of range
         Case Is <= Coll.Count / 2                     'walk items upwards from first
           PokeLong Ptr, ByVal ObjPtr(Coll) + 24       'first Ptr
           For i = 2 To Index
               PokeLong Ptr, ByVal Ptr + 24            'next Ptr
           Next i
         Case Else                                     'walk items downwards from last
           PokeLong Ptr, ByVal ObjPtr(Coll) + 28       'last Ptr
           For i = Coll.Count - 1 To Index Step -1
               PokeLong Ptr, ByVal Ptr + 20            'prev Ptr
           Next i
        End Select
       i = StrPtr(sKey)                                'save StrPtr
       PokeLong ByVal VarPtr(sKey), ByVal Ptr + 16     'replace StrPtr by that from collection sKey (which is null if there ain't no sKey)
       ItemKey = sKey                                  'now copy it to function value
       PokeLong ByVal VarPtr(sKey), i                  'and finally restore original StrPtr
   End If

End Function


Private Function ItemIndex(ByVal Key As String, Coll As Collection, Optional ByVal Compare As VbCompareMethod = vbTextCompare) As Long
 'get collection index by key
 'Private Declare Sub PokeLong Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, Optional ByVal Length As Long = 4)
 Dim Ptr   As Long
 Dim sKey  As String
 Dim aKey  As Long
   
   If Coll Is Nothing Then                             'oops!
       Err.Raise 91                                    'No object
     Else 'NOT COLL...
       If Coll.Count Then
           aKey = StrPtr(sKey)                         'save StrPtr
           PokeLong Ptr, ByVal ObjPtr(Coll) + 24       'first Ptr
           ItemIndex = 1                               'walk items upwards from first
           Do
               PokeLong ByVal VarPtr(sKey), ByVal Ptr + 16
               If StrComp(Key, sKey, Compare) = 0 Then 'equal
                   Exit Do                             'found
               End If
               ItemIndex = ItemIndex + 1               'next Index
               PokeLong Ptr, ByVal Ptr + 24            'next Ptr
           Loop Until Ptr = 0                          'end of chain
           PokeLong ByVal VarPtr(sKey), aKey           'restore original StrPtr
       End If
       If Ptr = 0 Then
           ItemIndex = -1                              'key not found
       End If
   End If

End Function


