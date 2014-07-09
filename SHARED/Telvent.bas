Attribute VB_Name = "Telvent"
Option Explicit
'**********************************************************
' Requires template.bas file be in project
'**********************************************************

Public Function IsTelventSetup(Msg As String) As Boolean
'    Dim SiteId As String
    Dim siteAddress As String
    Dim logonString As String
    Dim logonPwd As String
    Dim outputFile As String
    Dim outputFolder As String
    Dim ftpUser As String
    Dim ftpPwd As String
    Dim rs As Recordset
    Dim sql As String
    Dim val As String
    Dim Key As String
    
    Msg = ""
    
    sql = "select * from sys_ini where ini_file_name = 'TELVENT' and ini_section = 'DTN'"
    Set rs = t_dbMainDatabase.OpenRecordset(sql, dbOpenSnapshot, dbSQLPassThrough)
    If rs.RecordCount > 0 Then
        While Not rs.EOF
            Key = Trim(rs("ini_field_name").value & "")
            Select Case UCase(Key)
'                Case "SITEID"
'                    SiteId = Trim(rs("ini_value").value & "")
                Case "LOGON-STRING"
                    logonString = Trim(rs("ini_value").value & "")
                Case "LOGON-PWD"
                    logonPwd = Trim(rs("ini_value").value & "")
                Case "OUTPUT-FILE"
                    outputFile = Trim(rs("ini_value").value & "")
                Case "OUTPUT-FOLDER"
                    outputFolder = Trim(rs("ini_value").value & "")
                Case "SITEADDRESS"
                    siteAddress = Trim(rs("ini_value").value & "")
                Case "FTPUSER"
                    ftpUser = Trim(rs("ini_value").value & "")
                Case "FTPPWD"
                    ftpPwd = Trim(rs("ini_value").value & "")
            End Select
            rs.MoveNext
        Wend
    End If
    
'    If Len(SiteId) < 1 Then Msg = Msg & ", " & "Site ID is empty"
    If Len(logonString) < 1 Then Msg = Msg & ", " & "Logon String is empty"
    If Len(logonPwd) < 1 Then Msg = Msg & ", " & "Logon Password is empty"
    If Len(siteAddress) < 1 Then Msg = Msg & ", " & "Site Address is empty"
    If Len(outputFile) < 1 Then Msg = Msg & ", " & "Output File is empty"
    If Len(outputFolder) < 1 Then
        Msg = Msg & ", " & "Output Folder is empty"
    ElseIf Not io.DirectoryExists(outputFolder) Then
        Msg = Msg & ", " & "Output Folder '" & outputFolder & "' does not exist"
    End If
    If Len(ftpUser) < 1 Then Msg = Msg & ", " & "Ftp User is empty"
    If Len(ftpPwd) < 1 Then Msg = Msg & ", " & "Ftp Password is empty"
    If Len(Msg) > 3 Then Msg = Mid(Msg, 3)
        
    IsTelventSetup = Len(Msg) = 0
    
End Function

Public Function RunTelventSetup() As Double
    Dim arftvdtn As String
    Dim s As String
    Dim Cmd As String
    
    'tfnRun "ARFTVDTN.EXE", 1, False, t_szConnect
    arftvdtn = IIf(Right(App.Path, 1) = "\", App.Path & "ARFTVDTN.EXE", App.Path & "\ARFTVDTN.EXE")
    If io.FileExists(arftvdtn) Then
        Cmd = IIf(Right(App.Path, 1) = "\", App.Path & "ARFTVDTN.EXE", App.Path & "\ARFTVDTN.EXE")
        Cmd = Cmd & " " & Chr(34) & t_szConnect & Chr(34)
        RunTelventSetup = Shell(Cmd, vbNormalFocus)
    Else
        MsgBox "The program '" & arftvdtn & "' was not found"
        RunTelventSetup = -1#
    End If
    
End Function
