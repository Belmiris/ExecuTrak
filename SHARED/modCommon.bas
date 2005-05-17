Attribute VB_Name = "modCommon"
Option Explicit

'Enums
Public Enum DatabaseLocation
    RemoteDB = 1
    LocalDB = 2
End Enum

Public Enum ButtonStatus
    Disable = 0
    Enable = 1
End Enum

Public dbLocal As DAO.DataBase 'Local MS Access Database

Public Function GetSysParm(ByVal ParmNum As Long, Optional ByVal Default As String = vbNullString, Optional ByVal Reload As Boolean = False) As String
    Static SysParms As Collection
    Dim SQL         As String
    Dim rs          As DAO.Recordset
    
    On Error GoTo ErrHandler
    
    If (SysParms Is Nothing) Or Reload Then
        Set SysParms = New Collection
        SQL = "SELECT Parm_Nbr,Parm_Field FROM Sys_Parm"
        If fnRecordset(rs, SQL) > 0 Then
            Do While Not rs.EOF
                SysParms.Add Trim$(rs(1).value & vbNullString), "sp" & rs(0).value
                rs.MoveNext
            Loop
        End If
        rs.Close
        Set rs = Nothing
    End If
    
    GetSysParm = SysParms("sp" & ParmNum)
    Exit Function
    
ErrHandler:
    GetSysParm = Default
    Err.Clear
End Function
Public Sub SelectAllText()
    On Error GoTo ErrHandler
    With Screen.ActiveControl
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Exit Sub
    
ErrHandler:
    Err.Clear
End Sub
Function SQLParm(ByVal SQL As String, ParamArray Parms()) As String
    Dim MaxIndex As Integer
    Dim Index    As Integer
    Dim sTemp    As String
    
    'Get Index of the Last Parameter Name that has a Replacement Value following it
    MaxIndex = UBound(Parms) - 1
    If MaxIndex Mod 2 Then
        'The last parm specified was not given a replacement value.
        'Should have been an even number of parameters but since we
        'have an odd number of parameters, we'll simply ignore the last one.
        MaxIndex = MaxIndex - 1
    End If
    
    'Replace all the specified Parameters
    If MaxIndex >= 0 Then
        Index = 0
        Do While Index <= MaxIndex
            sTemp = Parms(Index + 1)
            If InStrB(1, sTemp, "'") > 0 Then
                sTemp = Replace(sTemp, "'", "''")
            End If
            SQL = Replace(SQL, Parms(Index), sTemp, , , vbTextCompare)
            Index = Index + 2
        Loop
    End If
    
    SQLParm = SQL
    
End Function

Public Function fnRecordset(rsTemp As Recordset, SQL As String, _
                   Optional dbLocation As DatabaseLocation = RemoteDB, Optional sCalledFrom As String, _
                   Optional bShowErrow As Boolean) As Long
Attribute fnRecordset.VB_Description = "Modifies the passed recordset with data from the passed SQL and returns the recordcount of the recordset."
    On Error GoTo SQLError
    
    Select Case dbLocation
        Case DatabaseLocation.LocalDB
            Set rsTemp = dbLocal.OpenRecordset(SQL, dbOpenSnapshot)
        Case DatabaseLocation.RemoteDB
            Set rsTemp = t_dbMainDatabase.OpenRecordset(SQL, dbOpenSnapshot, dbSQLPassThrough)
    End Select
    
    If rsTemp.RecordCount > 0 Then
       rsTemp.MoveLast
       rsTemp.MoveFirst
    End If
    
    fnRecordset = rsTemp.RecordCount
    
    Exit Function
    
SQLError:
    If IsMissing(sCalledFrom) Then
        sCalledFrom = vbNullString
    End If
    
    If IsMissing(bShowErrow) Then
        bShowErrow = True
    End If
    
    fnRecordset = -1
    tfnErrHandler "fnRecordset," & sCalledFrom, SQL, bShowErrow
    
    On Error GoTo 0
    
End Function

Public Function fnExecSQL(SQL As String, Optional dbLocation As DatabaseLocation = RemoteDB, _
                Optional sCalledFrom As Variant = vbNullString, Optional bShowError As Variant = True) As Boolean

On Error GoTo SQLError

    Select Case dbLocation
        Case DatabaseLocation.LocalDB
            dbLocal.Execute SQL, dbSQLPassThrough
        Case DatabaseLocation.RemoteDB
            t_dbMainDatabase.ExecuteSQL SQL
            
    End Select
    
    fnExecSQL = True
    
    Exit Function
    
SQLError:
    tfnErrHandler "fnExecSQL, " & sCalledFrom, SQL, bShowError
      
    On Error GoTo 0
End Function

Public Sub subSetButtonStatus(ByRef objButton As FactorFrame, Status As ButtonStatus, _
                                                Optional ByVal ContextForm As Form = Nothing)
    If ContextForm Is Nothing Then
        Set ContextForm = frmContext
    End If
    
    objButton.Enabled = Status
    If Status = Enable Then
        objButton.Picture = ContextForm.LoadPicture(SEARCH_UP)
    Else
        objButton.Picture = ContextForm.LoadPicture(SEARCH_DOWN)
    End If
End Sub

Public Function IsUserSure(Prompt, Optional ByVal YesNoType As Boolean = True, Optional DefaultToNo As Boolean = True) As Boolean
    '-------------------------------------------------------------------
    '   Author.: DenBorg
    '   Written: 04/27/2005
    '
    '   This function simplifies asking the user a Yes/No or OK/Cancel
    '   question and getting True/False as a result.
    '
    '   To ask a Yes/No question, set YesNoType to TRUE (the default)
    '   To ask a OK/Cancel question, set YesNoType to FALSE
    '
    '   By default, the default button on the MsgBox is the second
    '   button (No or Cancel). The default button can be changed to the
    '   first button (Yes or OK) by setting the DefaultToNo parameter
    '   to FALSE.
    '-------------------------------------------------------------------
    Dim Buttons As VbMsgBoxStyle
    Dim YesOK   As VbMsgBoxResult
    
    If YesNoType Then
        Buttons = vbYesNo
        YesOK = vbYes
    Else
        Buttons = vbOKCancel
        YesOK = vbOK
    End If
    If DefaultToNo Then
        Buttons = Buttons + vbDefaultButton2
    End If
    Buttons = Buttons + vbQuestion
    
    IsUserSure = (MsgBox(Prompt, Buttons) = YesOK)
End Function
'-------------------------------------------------------------------
'   Author.: DenBorg
'   Written: 04/27/2005
'
'   This function overrides VBA's MsgBox() function so that the
'   Application Title (App.Title) can automatically be added to
'   the MsgBox's Title Bar if not already present, as per Factor
'   standards.
'
'   All MsgBox statements in the program passes through this
'   function, which in turn calls the VBA.MsgBox function.
'-------------------------------------------------------------------
Public Function MsgBox(Prompt, Optional Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal Title, Optional HelpFile, Optional Context) As VbMsgBoxResult

    If IsMissing(Title) Then
        Title = App.Title
    Else
        If StrComp(Left$(Title, Len(App.Title)), App.Title, vbTextCompare) Then
            Title = App.Title & " - " & Title
        End If
    End If
    MsgBox = VBA.MsgBox(Prompt, Buttons, Title, HelpFile, Context)
End Function

Public Function AppFile(ByVal Filename As String) As String
    AppFile = AppPath() & Filename
End Function
 
Public Function AppPath() As String
    AppPath = FixPath(App.Path)
End Function
 
Public Function FileExists(ByVal Filename As String) As Boolean
    Dim bExists As Boolean
    
    On Error Resume Next
    FileLen Filename
    bExists = (Err.Number = 0)
    If bExists Then
        bExists = (GetAttr(Filename) And vbDirectory = 0)
    End If
    
    FileExists = bExists
End Function
 
Public Function DirExists(ByVal DirName As String) As Boolean
    Dim bExists As Boolean
    
    On Error Resume Next
    If Right$(DirName, 1) = "\" Then
        DirName = Left$(DirName, Len(DirName) - 1)
    End If
    FileLen DirName
    bExists = (Err.Number = 0)
    If bExists Then
        bExists = (GetAttr(DirName) And vbDirectory = vbDirectory)
    End If
    
    DirExists = bExists
End Function
 
Public Function FixPath(ByVal Path As String) As String
    If Right$(Path, 1) <> "\" Then
        Path = Path & "\"
    End If
    
    FixPath = Path
End Function
