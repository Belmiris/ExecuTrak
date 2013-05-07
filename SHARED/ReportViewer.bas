Attribute VB_Name = "ReportViewer"
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
    (ByVal lpApplicationName As String, _
     ByVal lpKeyName As Any, _
     ByVal lpDefault As String, _
     ByVal lpReturnedString As String, _
     ByVal nSize As Long, _
     ByVal lpFileName As String) As Long
     
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
    (ByVal lpApplicationName As String, _
     ByVal lpKeyName As Any, _
     ByVal lpString As Any, _
     ByVal lpFileName As String) As Long

Option Explicit

Public Function ShellVehicleAnalysis(custNum As String) As Boolean
    
    ShellVehicleAnalysis = ShellReportViewer("&CATEGORY=MNR&REPORT=Vehicle Analysis and Inv|mnrvehcl|mnmenu|", "PARMf3=" & custNum)
    
End Function

Private Function ShellReportViewer(xtraPostData As String, settings As String) As Boolean
    On Error GoTo FINISHED
    
    Dim engine As DAO.DBEngine
    Dim workspace As DAO.workspace
    Dim database As DAO.database
    
    Dim strTempHostName$
    Dim strTempDatabaseName$
    Dim strProgPath$
    Dim strPostData$
    Dim strReportServerURL$
    Dim strUser$
    Dim strPass$
    Dim strDBName$
    Dim strCmd$
    Dim i As Integer
    Dim lTempInstance As Long
    Dim lTempHWND As Long
    
    Set engine = New DBEngine
    Set workspace = engine.Workspaces(0)
    Set database = workspace.OpenDatabase("", False, False, t_oleObject.MainConnectString)
    
    ' HOST NAME
    strTempHostName = tfnGetNamedString(t_oleObject.MainConnectString, "HOST")
    If Trim(strTempHostName) = "" Then
        strTempHostName = tfnGetNamedString(database.Connect, "SRVR")
    End If
    
    ' DATABASE NAME
    strTempDatabaseName = ""
    strTempDatabaseName = tfnGetNamedString(database.Connect, "DATABASE")
    
    If strTempDatabaseName = "ssfactor" Then
        strTempDatabaseName = database.name
    End If
    
    If strTempDatabaseName = "" Then
        strTempDatabaseName = tfnGetNamedString(database.Connect, ";DB")
    End If
    
    ' 4GL PATH
    strProgPath = Trim(fnDefaultParm("SETUP OF 4GL PROGRAMS", "PROG PATH", "/usr/factor"))
        
    ' USER
    strUser = tfnGetNamedString(t_oleObject.SecurityConnectString, "UID")
        
    ' PASSWORD
    strPass = tfnGetNamedString(database.Connect, "PWD")
    If Trim(strPass) = "" Then
        strPass = tfnGetNamedString(t_oleObject.SecurityConnectString, "PWD")
    End If
            
    ' DATABSE NAME
    strDBName = database.name
    
    ' CLOSE DATABASE
    database.Close
    Set database = Nothing
    
    ' Determine actual host for program
    strTempHostName = DetermineReportHost(strTempHostName, strProgPath)
    
    ' BUILD COMMAND LINE
    strReportServerURL = "http://" & strTempHostName & "/cgi-bin/reportsrv.cgi"
    strPostData = "USERNAME=" & strUser & "&PASSWORD=" & strPass
    strPostData = strPostData & "&DATABASE=" & strTempDatabaseName
    strPostData = strPostData & "&PROGPATH=" & strProgPath
    If Len(xtraPostData) Then strPostData = strPostData & xtraPostData
    
    strCmd = App.path & "\" & "ReportViewer.exe"
    If Dir$(strCmd) = "" Then
        MsgBox "Unable to find file '" + strCmd + "'"
        ShellReportViewer = False
        GoTo FINISHED
    End If
        
    If Len(settings) > 0 Then
        strCmd = strCmd + " " + _
                 Chr$(34) + strPostData + Chr$(34) + " " + _
                 Chr$(34) + strReportServerURL + Chr$(34) + " " + _
                 Chr$(34) + strDBName + Chr$(34) + " " + _
                 Chr$(34) + settings + Chr$(34) + _
                 Chr$(0)
    Else
        strCmd = strCmd + " " + _
                 Chr$(34) + strPostData + Chr$(34) + " " + _
                 Chr$(34) + strReportServerURL + Chr$(34) + " " + _
                 Chr$(34) + strDBName + Chr$(34) + _
                 Chr$(0)
    End If
    
    Shell strCmd, vbNormalFocus
    
    ShellReportViewer = True
    
    Err.Clear
FINISHED:
    If Err.Number <> 0 Then
        MsgBox "Error in ShellReportViewer: " & Err.Description
        ShellReportViewer = False
    End If
End Function

Private Function fnDefaultParm(sSECTION As String, _
                               sKey As String, _
                               sDefault As String) As String
    Dim sIniFileName As String
    Dim nLength As Long
    Dim sBuffer As String
    Dim bStatus As Boolean

    sIniFileName = "C:\FACTOR\FACTOR.INI"
    
    sBuffer = Space(MAX_STRING_LENGTH)
    
    nLength = GetPrivateProfileString(sSECTION, sKey, "", sBuffer, MAX_STRING_LENGTH, sIniFileName)
    
    If nLength <> 0 Then
        fnDefaultParm = Left(sBuffer, nLength)
    Else
        WritePrivateProfileString sSECTION, sKey, sDefault, sIniFileName
        fnDefaultParm = sDefault
    End If

End Function

Private Function DetermineReportHost(host As String, strProgPath As String)
    
    DetermineReportHost = host
    
    'ServiceTrak    4958
    Dim nIdx As Integer
    Dim strSvcHostName As String                        'Alternate host name for serviceTrak
        
    nIdx = InStr(strProgPath, "/")                      'Get the first slash in /user/factor
    If nIdx > 0 Then
        nIdx = InStr(nIdx + 1, strProgPath, "/")        'Get the second slash
        If nIdx > 0 Then
            strSvcHostName = Mid(strProgPath, nIdx + 1) 'Extract the host name after the second slash, like 'factor'
        End If
    End If
    
    If StrComp(strSvcHostName, "factor", vbTextCompare) <> 0 Then
        DetermineReportHost = strSvcHostName
    End If
    
End Function

