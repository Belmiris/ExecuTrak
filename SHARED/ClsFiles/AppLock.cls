VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AppLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const PROCESS_ALL_ACCESS = &H1F0FFF
Private Const STILL_ACTIVE = &H103
Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const PROCESS_VM_READ = &H10

Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function GetProcessId Lib "kernel32" (ByVal hProcess As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

Private dBase As Database
Private applicationName As String
Private processId As Long
Private recordExists As Boolean
Private userLocked As String
Private bCloseDb As Boolean

Public Function Load(appName As String, Optional db As Database = Nothing, Optional connectionString = "") As Boolean
    Dim en As New DBEngine
    Dim ws As Workspace
    
    If Not db Is Nothing Then
        Set dBase = db
    Else
        If Len(connectionString) < 1 Then Err.Raise -1, , "AppLock.Load must pass either Database or Connection String"
        Set ws = en.Workspaces(0)
        Set dBase = ws.OpenDatabase("", False, False, connectionString)
        bCloseDb = True
    End If
    
    applicationName = appName
    LoadInfo
    Load = recordExists

End Function

Public Sub Unload()
    If bCloseDb And Not dBase Is Nothing Then
        dBase.Close
        Set dBase = Nothing
    End If
End Sub

Public Sub CreateLock(userName As String)
    Dim sql$
    Dim pid As Long
    
    DeleteLock
    
    pid = GetCurrentProcessId()
    sql = "insert into sys_data_lock (sdl_data, sdl_prog, sdl_user) values ('" & pid & "', '" & applicationName & "', '" & userName & "')"
    dBase.ExecuteSQL sql
    LoadInfo
    
End Sub

Public Function DeleteLock() As Boolean
    Dim sql$
    
    sql = "delete from sys_data_lock where sdl_prog = '" + applicationName + "'"
    dBase.ExecuteSQL sql
    
End Function

Public Property Get ProcessActive() As Boolean
    
    LoadInfo
    ProcessActive = ProcessIsActive(processId)
    
End Property

Public Property Get Exists() As Boolean
    
    LoadInfo
    Exists = recordExists
    
End Property

Private Sub LoadInfo()
    Dim sql As String
    Dim pid As String
        
    If dBase Is Nothing Then Err.Raise -1, , "AppLock: Load must be called with a valid database!"
    If Len(applicationName) < 1 Then Err.Raise -1, , "AppLock: Load must be called with an application name!"
    
    recordExists = False
    userLocked = ""
    processId = 0
    
    sql = "select sdl_user, sdl_data from sys_data_lock where sdl_prog = '" + applicationName + "'"
    With dBase.OpenRecordset(sql, dbOpenSnapshot, dbSQLPassThrough)
        If Not .EOF Then
            recordExists = True
            userLocked = Trim$(.Fields(0).value & "")
            pid = Trim$(.Fields(1).value & "")
        End If
        .Close
    End With
    
    If Len(pid) > 0 Then
        processId = CLng(pid)
    End If
    
End Sub

Public Function ProcessIsActive(procId As Long) As Boolean
    Dim lExit As Long
    Dim hdl As Long
    
    If procId <> 0 Then
        hdl = OpenProcess(PROCESS_QUERY_INFORMATION, 0, procId)
        GetExitCodeProcess hdl, lExit
        If hdl <> 0 Then
            If procId = GetProcessId(hdl) And lExit = STILL_ACTIVE Then
                ProcessIsActive = True
            End If
        End If
    End If

End Function
