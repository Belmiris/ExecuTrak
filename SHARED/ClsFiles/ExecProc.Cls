VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
END
Attribute VB_Name = "clsSQLExecProcedure"
Attribute VB_Creatable = True
Attribute VB_Exposed = True
    Private t_dbMainDatabase As Database
    Private t_wsWorkSpace As Workspace
    Private t_engFactor As DBEngine

    Private m_sConnect As String         'This holds the ODBC connect string passed from oleObject
    Private m_sSQL As String
    Private m_bDatabaseOpened As Boolean
Property Let Connect(sConnect As String)
    m_sConnect = sConnect
End Property


Public Function ExecuteSQL() As Integer
    On Error GoTo errExecute
    
    fnOpenDatabase
    
    t_dbMainDatabase.Execute m_sSQL, dbSQLPassThrough

    ExecuteSQL = 0
    On Error GoTo 0
    Exit Function

errExecute:
    ExecuteSQL = Err.Number
End Function

Private Sub fnOpenDatabase()

    If m_bDatabaseOpened Then
        Exit Sub
    End If
    
    Set t_engFactor = New DBEngine 'create a new dDBEngine
    
    t_engFactor.IniPath = "C:\WINDOWS\SYSTEM"
    
    Set t_wsWorkSpace = t_engFactor.Workspaces(0) 'set the default workspace handle
    
    Set t_dbMainDatabase = t_wsWorkSpace.OpenDatabase("", False, False, m_sConnect)

    m_bDatabaseOpened = True

End Sub

Property Let SQL(sSQL As String)
    m_sSQL = sSQL
End Property

Private Sub Class_Initialize()
    m_bDatabaseOpened = False
End Sub


Private Sub Class_Terminate()
    If Not t_dbMainDatabase Is Nothing Then
    t_dbMainDatabase.Close
    Set t_dbMainDatabase = Nothing
    Set t_wsWorkSpace = Nothing
    Set t_engFactor = Nothing
    End If
    m_bDatabaseOpened = False
End Sub

