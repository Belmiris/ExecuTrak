Attribute VB_Name = "modCommon"
Option Explicit

Private Const GWL_STYLE = (-16)

'Enums
Public Enum TextBoxStyles
    UpperCase = &H8&
    LowerCase = &H10&
    Numeric = &H2000&
End Enum

Public Enum DatabaseLocation
    RemoteDB = 1
    LocalDB = 2
End Enum

Public Enum ButtonStatus
    Disable = 0
    Enable = 1
End Enum

Public Const INTO_TEMP As String = " into temp "
Public Const SQL_DROP_TABLE As String = "drop table @table"
Public Const SQL_TABLE_EXISTS As String = _
    "select tabname from systables where tabname = '@table'"

Public dbLocal As DAO.Database 'Local MS Access Database

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
    ByVal hWnd As Long, _
    ByVal nIndex As Long _
) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
    ByVal hWnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long _
) As Long

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
                SysParms.Add Trim$(rs(1).Value & vbNullString), "sp" & rs(0).Value
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
Public Function ReadEntireFile(ByVal Filename As String) As String
    Dim hFile As Integer
    
    If FileExists(Filename) Then
        hFile = FreeFile()
        Open Filename For Binary As #hFile
        ReadEntireFile = Input(LOF(hFile), hFile)
        Close #hFile
    End If
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
Public Sub SetTextBoxStyle(TextBox As TextBox, ByVal Style As TextBoxStyles, Optional ByVal EnableStyle As Boolean = True)
    With TextBox
        If EnableStyle Then
            Style = GetWindowLong(.hWnd, GWL_STYLE) Or Style
        Else
            Style = GetWindowLong(.hWnd, GWL_STYLE) And (Not Style)
        End If
        
        SetWindowLong .hWnd, GWL_STYLE, Style
    End With
End Sub
Public Function StringAppend(ByRef StrValue, ByVal Delimeter As String, ParamArray AppendValues() As Variant)
    If LenB(Delimeter) = 0 Then
        Delimeter = ","
    End If
    
    If LenB(StrValue) Then
        StrValue = StrValue & Delimeter
    End If
    
    StringAppend = StrValue & Join(AppendValues, Delimeter)
End Function
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
            sTemp = Trim$(Parms(Index + 1) & vbNullString)
            
            If sTemp = "0" Or sTemp = vbNullString Then
                If InStrB(1, SQL, "'" & Parms(Index)) > 0 Then
                    sTemp = vbNullString
                Else
                    sTemp = "0"
                End If
            Else
                If InStrB(1, sTemp, "'") > 0 Then
                    sTemp = Replace(sTemp, "'", "''")
                End If
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

'---------------------------------------------------------------------------------------
' Procedure : fnQueryForField
' DateTime  : 5/26/2005 15:43
' Author    : Chris Albrecht
' Purpose   : Get a single field from the result of a query
'---------------------------------------------------------------------------------------
Public Function fnQueryForField(SQL As String, Optional FieldName As String, _
                   Optional dbLocation As DatabaseLocation = RemoteDB, Optional sCalledFrom As String, _
                   Optional bShowErrow As Boolean = True) As Variant
    
    Dim rsTemp As Recordset
    
    On Error GoTo SQLError
    
    Select Case dbLocation
        Case DatabaseLocation.LocalDB
            Set rsTemp = dbLocal.OpenRecordset(SQL, dbOpenSnapshot)
        Case DatabaseLocation.RemoteDB
            Set rsTemp = t_dbMainDatabase.OpenRecordset(SQL, dbOpenSnapshot, dbSQLPassThrough)
    End Select
    
    If rsTemp.RecordCount > 0 Then
        If Not IsMissing(FieldName) And FieldName <> vbNullString Then
            fnQueryForField = fnGetField(rsTemp(FieldName))
        Else
            fnQueryForField = fnGetField(rsTemp.Fields(0))
        End If
    End If
        
    Exit Function
    
SQLError:
    If IsMissing(sCalledFrom) Then
        sCalledFrom = vbNullString
    End If
        
    fnQueryForField = vbNullString
    
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

Public Function fnCreateTempTable(SQL As String, TableName As String) As Boolean

    fnCreateTempTable = fnExecSQL(SQL & INTO_TEMP & TableName)
    
End Function

Public Sub subDropTable(TableName As String)
    Dim sSQL As String
    
    'In case the table doesn't exist, just continue
    On Error Resume Next
    sSQL = SQLParm(SQL_DROP_TABLE, "@table", TableName)
    
    fnExecSQL sSQL, , , False
    
End Sub

Public Function fnTableExists(TableName As String) As Boolean
    Dim sSQL As String
    Dim rsTemp As Recordset
    
    sSQL = SQLParm(SQL_TABLE_EXISTS, _
                          "@table", TableName)
                          
    If fnRecordset(rsTemp, sSQL) > 0 Then
        fnTableExists = True
    End If
    
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
'
Public Function IsUserSure(Prompt, Optional ByVal YesNoType As Boolean = True, Optional DefaultToNo As Boolean = True) As Boolean
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
        bExists = ((GetAttr(Filename) And vbDirectory) = 0)
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
Public Sub UnloadAllForms()
    Dim Form As Form
    
    For Each Form In Forms
        Unload Form
    Next 'Form
    Set Form = Nothing
End Sub

Public Function AsciiUCase(ByVal KeyAscii As Integer) As Integer
    If (KeyAscii >= 97) And (KeyAscii <= 122) Then
        KeyAscii = KeyAscii - 32
    End If
    
    AsciiUCase = KeyAscii
End Function
