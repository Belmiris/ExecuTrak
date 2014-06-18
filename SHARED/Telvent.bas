Attribute VB_Name = "Telvent"
Option Explicit

Public Function IsTelventSetup(msg As String) As Boolean
    Dim siteId As String
    Dim siteAddress As String
    Dim ftpUser As String
    Dim ftpPwd As String
    Dim rs As Recordset
    Dim sql As String
    Dim Val As String
    Dim key As String
    
    msg = ""
    
    sql = "select * from sys_ini where ini_file_name = 'TELVENT' and ini_section = 'DTN'"
    Set rs = t_dbMainDatabase.OpenRecordset(sql, dbOpenSnapshot, dbSQLPassThrough)
    If rs.RecordCount > 0 Then
        While Not rs.EOF
            key = Trim(rs("ini_field_name").value & "")
            Select Case UCase(key)
                Case "SITEID"
                    siteId = Trim(rs("ini_value").value & "")
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
    
    If Len(siteId) < 1 Then msg = msg & ", " & "Site ID is empty"
    If Len(siteAddress) < 1 Then msg = msg & ", " & "Site Address is empty"
    If Len(ftpUser) < 1 Then msg = msg & ", " & "Ftp User is empty"
    If Len(ftpPwd) < 1 Then msg = msg & ", " & "Ftp Password is empty"
    If Len(msg) > 3 Then msg = Mid(msg, 3)
    
    IsTelventSetup = Len(msg) = 0
    
End Function

Public Function RunTelventSetup()
    
    tfnRun "ARFTVDTN.EXE", 1, False, t_szConnect
    
End Function
