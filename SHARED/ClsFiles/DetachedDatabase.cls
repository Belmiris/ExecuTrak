VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DetachedDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum DetachedDatabaseType
    ddtNone = 0
    ddtAccess = 1
    ddtInformix = 2
End Enum

Private m_database As Database
Private m_databaseType As DetachedDatabaseType
Private m_engine As DBEngine
Private m_connection As String
Private m_workspace As Workspace
'

Private Sub Class_Initialize()
    m_databaseType = ddtNone
    
End Sub

Private Sub Class_Terminate()
    CloseDatabase
    
End Sub

Public Property Get Connection() As String
    Connection = m_connection

End Property

Public Property Get DatabaseType() As DetachedDatabaseType
    DatabaseType = m_databaseType

End Property

Public Function AssignConnection(ByVal strConnection As String, ByVal dbType As DetachedDatabaseType)
    
    If dbType = ddtNone Then Err.Raise -5381, "DetachedDatabase.AssignConnection", "Invalid Connection Type passed"
    If Len(strConnection) < 1 Then Err.Raise -5380, "DetachedDatabase.AssignConnection", "Connection string cannot be empty"
        
    m_databaseType = dbType
    m_connection = strConnection
    
End Function

Public Sub CloseDatabase()
    On Error Resume Next

    Set m_engine = Nothing
    Set m_workspace = Nothing

    If Not m_database Is Nothing Then
        m_database.Close
        Set m_database = Nothing
    End If

    Err.Clear
End Sub

Public Function CopyRecordset(rsTemp As Recordset, Optional bMoveFirst As Boolean = True) As DetachedRecordset
    Dim rsDetached As DetachedRecordset
    
    If rsTemp Is Nothing Then _
        Err.Raise -2000, "DetachedDatabase.CopyRecordset", "DetachedDatabase.CopyRecordset Empty rsTemp passed"
    
    Set rsDetached = New DetachedRecordset
    
    If rsTemp.RecordCount > 0 Then
        If bMoveFirst Then rsTemp.MoveFirst
        
        While Not rsTemp.EOF
            rsDetached.CopyCurrentRecord rsTemp
            rsTemp.MoveNext
        Wend
    End If
    
    Set CopyRecordset = rsDetached
    
End Function

Public Function ExecuteSQL(strSQL As String) As Long
    
    ExecuteSQL = Execute(strSQL)

End Function

Public Function Execute(strSQL As String) As Long
    
    OpenDatabase
    
    Select Case m_databaseType
        Case ddtAccess
            m_database.Execute strSQL
            Execute = m_database.RecordsAffected
        
        Case ddtInformix
            Execute = m_database.ExecuteSQL(strSQL)
        
        Case Else
            Err.Raise -5368, "DetachedDatabase.Execute", "Database Type unhandled or not set (" & m_databaseType & "')"
    End Select
    
    CloseDatabase
    
End Function

Public Function AssignFactorMdb() As Boolean
    Dim dsnPath As String
    
    CloseDatabase
    
    dsnPath = LOCAL_FACTOR_PATH & UCase(Trim(tfnGetDataSourceName)) + "\"
    
    If Not io Is Nothing Then
        If io.MultiDatabase Then
            dsnPath = io.ApplicationPath
        End If
    End If
    
    m_connection = IIf(Right(dsnPath, 1) = "\", dsnPath & "FACTOR.MDB", dsnPath & "\FACTOR.MDB")
    
    m_databaseType = ddtAccess
    
    If io.FileExists(m_connection) Then
        AssignFactorMdb = True
    End If
    
End Function

Public Function OpenRecordset(strSQL As String, rsType As RecordsetTypeEnum) As DetachedRecordset
    Dim rsDetached As DetachedRecordset
    Dim rsTemp As Recordset
    
    Set rsDetached = New DetachedRecordset
    
    OpenDatabase
    
    Select Case m_databaseType
        Case ddtAccess
            Set rsTemp = m_database.OpenRecordset(strSQL, dbOpenSnapshot)
        
        Case ddtInformix
            Set rsTemp = m_database.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
        
        Case Else
            Err.Raise -5368, "DetachedDatabase.OpenRecordset", "Database Type unhandled or not set (" & m_databaseType & "')"
    End Select
    
    If Not rsTemp Is Nothing Then
        If rsTemp.RecordCount > 0 Then
            While Not rsTemp.EOF
                rsDetached.CopyCurrentRecord rsTemp
                rsTemp.MoveNext
            Wend
        End If
    End If
    
    rsDetached.MoveFirst
    
    CloseDatabase
    
    Set OpenRecordset = rsDetached
    
End Function

Private Sub OpenDatabase()

    CloseDatabase
    
    Set m_engine = New DBEngine
    Set m_workspace = m_engine.Workspaces(0)
    
    Select Case m_databaseType
        Case ddtAccess
            If Len(m_connection) = 0 Then _
                Err.Raise -5364, "DetachedDatabase.OpenDatabase", "Access Database file path empty"
    
            If Dir(m_connection) = "" Then _
                Err.Raise -5365, "DetachedDatabase.OpenDatabase", "Access Database file not found '" & m_connection & "'"
            
            Set m_database = m_workspace.OpenDatabase(m_connection)
            
        Case ddtInformix
            If m_connection = "" Then _
                Err.Raise -5366, "DetachedDatabase.OpenDatabase", "Informix connection string empty"
            
            Set m_database = m_workspace.OpenDatabase("", False, False, m_connection)
        
        Case Else
            Err.Raise -5367, "DetachedDatabase.OpenDatabase", "Database Type unhandled or not set (" & m_databaseType & "')"
    End Select

End Sub
