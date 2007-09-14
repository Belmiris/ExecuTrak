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

Public Enum HourglassStatus
    ShowHourglass = vbHourglass
    HideHourglass = vbNormal
End Enum

Public Enum Justify
    LeftJustify = 0
    CenterJustify = 1
    RightJustify = 2
End Enum

Public Const INTO_TEMP As String = " into temp "
Public Const SQL_DROP_TABLE As String = _
    "DROP table @table"
Public Const SQL_TABLE_EXISTS As String = _
    "SELECT tabName FROM systables WHERE tabName = '@table'"
Public Const SQL_TEMP_TABLE_EXISTS As String = _
    "SELECT * FROM @table WHERE 1 = 2"
    
Public Const SQL_COLUMN_EXISTS As String = _
    " SELECT tabName, colName                   " & _
    "   FROM systables, syscolumns              " & _
    "  WHERE systables.tabid = syscolumns.tabid " & _
    "    AND tabName         = '@table'         " & _
    "    AND colName         = '@column'        "
    
Public Const SQL_SELECT_LOCK As String = _
        " select sdl_user from sys_data_lock where sdl_prog = '@program' "

Public Const SQL_INSERT_DATA_LOCK As String = _
    " INSERT into sys_data_lock                 " & _
    "        (sdl_user, sdl_prog, sdl_data)     " & _
    " VALUES ('@user', '@program', today)       "

Public Const SQL_DELETE_LOCK As String = _
    " DELETE from sys_data_lock                 " & _
    "  WHERE sdl_user = '@user'                 " & _
    "    AND sdl_prog = '@program'              "

#If Not dbLocalDef Then
Public dbLocal As DAO.Database 'Local MS Access Database
#End If

Public Const VK_LBUTTON = &H1
Public Const VK_RBUTTON = &H2
Public Const VK_MBUTTON = &H4

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
    ByVal hWnd As Long, _
    ByVal nIndex As Long _
) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
    ByVal hWnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long _
) As Long

Private Declare Sub SleepAPI Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)

Public debugCount As Integer
Private suppressMsgBox As Boolean
Private printSQL As Boolean

Public Property Let PrintSQLStatements(bPrint As Boolean)
    printSQL = bPrint
End Property

Public Function Q_Str(ByVal Str As String, Optional ByVal Quote As String = """") As String
    Q_Str = Quote & Str & Quote
End Function

'---------------------------------------------------------------------------------------
' Procedure : BackupFileName
' DateTime  : 11/23/2005 14:54
' Author    : DenBorg
' Magic     : 468962
' Purpose   : This routine, given a file Name and a backup directory, will return the
'             Name for a new backup file. Backup files are in the form of AAAA.###
'             where AAA is the Name of the file, and ### is a 3-digit counter. The
'             3-digit counter replaces the file's original file extension.
'
'             For example, for fileName CUSTOMER.DAT, this routine might return
'             something like CUSTOMER.017 (if there were already 16 backup files in the
'             specified backup path)
'---------------------------------------------------------------------------------------
'
Public Function BackupFileName(ByVal Filename As String, ByVal BackupPath As String, Optional ByVal UseFileExt As Boolean = False) As String
    Dim FileExt As String
    Dim FileNum As Integer
    Dim CurFile As String
    
    If Not UseFileExt Then
        '------------------------------------------------------------------------------------
        'Strip off Path and Extension from FileName
        '------------------------------------------------------------------------------------
        FileNameParts Filename, , Filename
    End If
    
    '------------------------------------------------------------------------------------
    'See if the file already has existing backup copies in BackupPath.
    'If so, take note of the highest backup counter value.
    '------------------------------------------------------------------------------------
    BackupPath = FixPath(BackupPath)
    CurFile = Dir(BackupPath & Filename & ".???")
    Do While LenB(CurFile)
        FileNameParts CurFile, , , FileExt
        If FileExt Like "###" Then
            If val(FileExt) > FileNum Then
                FileNum = val(FileExt)
            End If
        End If
        
        CurFile = Dir() 'Get next fileName
    Loop
    
    '------------------------------------------------------------------------------------
    'Increment counter
    '------------------------------------------------------------------------------------
    If FileNum < 999 Then
        FileNum = FileNum + 1
    Else
        FileNum = 999
    End If
    
    '------------------------------------------------------------------------------------
    'Return the Name for the new Backup File
    '------------------------------------------------------------------------------------
    BackupFileName = BackupPath & Filename & "." & Format$(FileNum, "000")
End Function

Public Sub AlignWithControl(Ctl As Control, AlignWith As Control)
    Dim OldMode As Integer
    Dim ChgMode As Boolean
    
    With Ctl
        If (TypeOf .Container Is Form) Or (TypeOf .Container Is PictureBox) Then
            ChgMode = True
            With .Container
                OldMode = .ScaleMode
                .ScaleMode = vbTwips
            End With
        End If
        
        .Left = ContainerToForm(AlignWith, 0) - ContainerToForm(Ctl.Container, 0) - PicBoxBorderSize(Ctl.Container)
        
        If ChgMode Then
            .Container.ScaleMode = OldMode
        End If
    End With
End Sub

Public Function ContainerToForm(Ctl As Object, ByVal CoordType As Integer) As Single
    Dim PrevMode          As Integer
    Dim value             As Single
    Dim BorderSize        As Single
    Dim IsContainerForm   As Boolean
    Dim IsContainerPicBox As Boolean
    
    If Not (TypeOf Ctl Is Form) Then
        With Ctl
            If TypeOf .Container Is PictureBox Then
                IsContainerPicBox = True
            ElseIf TypeOf .Container Is Form Then
                IsContainerForm = True
            End If
            
            If IsContainerForm Or IsContainerPicBox Then
                With .Container
                    PrevMode = .ScaleMode
                    .ScaleMode = vbTwips
                    If IsContainerPicBox Then
                        BorderSize = PicBoxBorderSize(Ctl.Container)
                    End If
                End With
            End If
            
            If CoordType = 0 Then
                value = .Left + BorderSize
            Else
                value = .Top + BorderSize
            End If
            
            If IsContainerForm Or IsContainerPicBox Then
                .Container.ScaleMode = PrevMode
            End If
            
            If Not IsContainerForm Then
                value = value + ContainerToForm(.Container, CoordType)
            End If
        End With
    End If
    
    ContainerToForm = value
End Function

Public Function PicBoxBorderSize(PicBox As Object) As Single
    Dim OldMode As Integer
    Dim ChgMode As Boolean
    Dim Size    As Single
    
    If TypeOf PicBox Is PictureBox Then
        With PicBox
            If (TypeOf .Container Is PictureBox) Or (TypeOf .Container Is Form) Then
                ChgMode = True
                With .Container
                    OldMode = .ScaleMode
                    .ScaleMode = vbTwips
                End With
            End If
            
            Size = (.Width - .ScaleWidth) / 2
            
            If ChgMode Then
                .Container.ScaleMode = OldMode
            End If
        End With
    End If
    
    PicBoxBorderSize = Size
End Function
'------------------------------------------------------------------------------------------
' Procedure : ArrayValueIndex
' DateTime  : 7/29/2005 11:36
' Author    : DenBorg
'
' Purpose   : This function looks up SearchValue in the column identified by SearchColumn
'             of DataArray and returns the Row Index where that value was found. If the
'             value is NOT found, then -1 is returned.
'
' **NOTE**  : This function works with zero-based arrays whose dimensions are (Column, Row),
'             instead of the traditional (Row, Column).
'------------------------------------------------------------------------------------------
'
Public Function ArrayValueIndex(DataArray As Variant, ByVal SearchColumn As Long, ByVal SearchValue As Variant) As Long
    Dim SearchRow As Long 'Row Index where SearchValue is found; -1 if not found.
    Dim row       As Long
    
    SearchRow = -1 'Assume SearchValue is not found
    Do While row <= UBound(DataArray, 2)
        If DataArray(SearchColumn, row) = SearchValue Then
            SearchRow = row
            Exit Do
        End If
        
        row = row + 1
    Loop
    
    ArrayValueIndex = SearchRow
End Function

Function IsLeapYear(ByVal YearDate As Variant) As Boolean
    Dim LeapYear As Boolean
    
    If IsDate(YearDate) Then
        'Convert to just year
        YearDate = Year(YearDate)
    End If
    
    If YearDate Mod 4 = 0 Then
        If YearDate Mod 100 = 0 Then
            If YearDate Mod 400 = 0 Then
                LeapYear = True
            End If
        Else
            LeapYear = True
        End If
    End If
    
    IsLeapYear = LeapYear
End Function

Public Sub EnableControls(ByVal Enabled As Boolean, ParamArray Controls() As Variant)
    Dim Index As Long
    
    On Error Resume Next 'Just in case parameter does not have a property Named 'Enabled'
    For Index = LBound(Controls) To UBound(Controls)
        Controls(Index).Enabled = Enabled
    Next 'Index
    On Error GoTo 0
End Sub

'---------------------------------------------------------------------------------------
' Procedure : EnableCtrlArray
' DateTime  : 6/15/2005 10:24
' Author    : DenBorg
' Purpose   : Enables/Disables all elements in a control array if no control array
'             indices are specified in Indices(). When indices are specified, then
'             the specified subset of elements in the control array are enabled/disabled.
'---------------------------------------------------------------------------------------
'
Public Sub EnableCtrlArray(CtrlArray As Object, ByVal Enabled As Boolean, ParamArray Indices() As Variant)
    Dim ParmIndex As Long
    Dim ctrl      As Object
    
    If UBound(Indices) = -1 Then
        'No specific elements were targetted, so enable/disable all elements in control array
        For Each ctrl In CtrlArray
            ctrl.Enabled = Enabled
        Next 'Ctrl
        Set ctrl = Nothing
    Else
        'Certain elements were specified, so only enable/disable those that are listed.
        For ParmIndex = 0 To UBound(Indices)
            CtrlArray(Indices(ParmIndex)).Enabled = Enabled
        Next 'ParmIndex
    End If
End Sub

Public Function GetSysParm(ByVal ParmNum As Long, Optional ByVal DEFAULT As String = vbNullString, Optional ByVal Reload As Boolean = False) As String
    Static SysParms As Collection
    Dim SQL         As String
    Dim rs          As DAO.Recordset
    
    On Error GoTo errHandler
    
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
    
    If LenB(Trim$(GetSysParm)) = 0 Then
        GetSysParm = DEFAULT
    End If
    
    Exit Function
    
errHandler:
    GetSysParm = DEFAULT
    Err.Clear
End Function

Public Function IsFormLoaded(ByVal FormName As String) As Boolean
    Dim Form As Form
    
    FormName = UCase$(FormName)
    For Each Form In Forms
        If UCase$(Form.Name) = FormName Then
            IsFormLoaded = True
            Exit For
        End If
    Next 'Form
    Set Form = Nothing
End Function

'---------------------------------------------------------------------------------------
' Procedure : IsKeyPressed
' DateTime  : 6/20/2005 12:03
' Author    : DenBorg
' Purpose   : Returns TRUE if the specified key is currently pressed; returns FALSE if not.
'---------------------------------------------------------------------------------------
'
Public Function IsKeyPressed(VirtualKey As Long) As Boolean
    IsKeyPressed = CBool(GetKeyState(VirtualKey) And &H80)
End Function

Public Function Nz(ByVal value As Variant, Optional ByVal ValueIfNull As Variant = vbNullString) As Variant
    If Not IsNull(value) Then
        Nz = value
    Else
        Nz = ValueIfNull
    End If
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

'---------------------------------------------------------------------------------------
' Procedure : RecordArray
' DateTime  : 7/15/2005 14:50
' Author    : DenBorg
' Purpose   : Returns records generated from SQL statement in a 2-dimensional array.
'             First Index represents the Field; Second Index represents the Row.
'             If there were no records, UBound() returns -1
'             If there was an error accessing the database, EMPTY is returned instead
'             of an array.
'---------------------------------------------------------------------------------------
'
Public Function RecordArray(SQL As String) As Variant
    Dim Data     As Variant
    Dim rs       As DAO.Recordset
    Dim RecCount As Long
    
    RecCount = fnRecordset(rs, SQL)
    If RecCount >= 0 Then
        With rs
            If RecCount > 0 Then
                Data = .GetRows(RecCount)
            Else
                Data = Array() 'No records ... return empty array
            End If
            .Close
        End With
    End If
    Set rs = Nothing
    
    RecordArray = Data
End Function
Public Sub SelectAllText()
    On Error GoTo errHandler
    With Screen.ActiveControl
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Exit Sub
    
errHandler:
    Err.Clear
End Sub
Public Sub SetTextBoxStyle(Textbox As Textbox, ByVal Style As TextBoxStyles, Optional ByVal EnableStyle As Boolean = True)
    With Textbox
        If EnableStyle Then
            Style = GetWindowLong(.hWnd, GWL_STYLE) Or Style
        Else
            Style = GetWindowLong(.hWnd, GWL_STYLE) And (Not Style)
        End If
        
        SetWindowLong .hWnd, GWL_STYLE, Style
    End With
End Sub
'---------------------------------------------------------------------------------------
' Procedure : SQL_FieldValue
' DateTime  : 7/26/2005 11:24
' Author    : DenBorg
'
' Purpose   : This function is used when building Strings containing SQL Statements,
'             such as in an UPDATE statement in which you set the value of various fields.
'
'             This function returns the Value with Quotes (if applicable) if Value is
'             not Null. If the Value *is* Null, then it returns the string "NULL". It
'             will optionally append a comma at the end.
'---------------------------------------------------------------------------------------
'
Public Function SQL_FieldValue(ByVal value As Variant, ByVal DataType As DAO.DataTypeEnum, Optional ByVal AppendComma As Boolean = False, Optional ByVal EmptyStringAsNull As Boolean = False) As String
    Dim FV    As String
    Dim Quote As String
    
    If EmptyStringAsNull Then
        If (Not IsNull(value)) And (LenB(value & "") = 0) Then
            value = Null
        End If
    End If
    Select Case DataType
        Case dbChar, dbGUID, dbText
            Quote = "'"
            If Not IsNull(value) Then
                value = Replace(value, "'", "''")
            End If
        Case dbDate, dbTime, dbTimeStamp
            Quote = "'"
    End Select

    If Not IsNull(value) Then
        FV = Quote & value & Quote
    Else
        FV = "NULL"
    End If
    
    If AppendComma Then
        FV = FV & ","
    End If
    
    SQL_FieldValue = FV
End Function

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
            
            If sTemp = vbNullString Then
                'Check to if the default value should be string or
                'numeric.  String values will be enclosed in single
                'quotes, so check the SQL string for a preceding
                'quote on the parm Name e.g.
                If InStrB(1, SQL, "'" & Parms(Index)) = 0 Then
                    sTemp = "0"
                End If
            Else
                'If the second character in the Parameter Name is a #,
                'then we want to skip replacing single quotes with
                'double quotes.  This would usually occur when the
                'replacing value is an 'IN' clause of string values
                'e.g.  ( and nbrstring in (@numstring) - with a parm value of  '123','456','789' )
                If Mid$(Parms(Index), 2, 1) <> "#" Then
                    If InStrB(1, sTemp, "'") > 0 Then
                        sTemp = Replace(sTemp, "'", "''")
                    End If
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
        
    If printSQL Then
        Debug.Print SQL & ";"
    End If
    
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
    
    If printSQL Then
        Debug.Print SQL & ";"
    End If
    
    Select Case dbLocation
        Case DatabaseLocation.LocalDB
            Set rsTemp = dbLocal.OpenRecordset(SQL, dbOpenSnapshot)
        Case DatabaseLocation.RemoteDB
            Set rsTemp = t_dbMainDatabase.OpenRecordset(SQL, dbOpenSnapshot, dbSQLPassThrough)
    End Select
    
    If rsTemp.RecordCount > 0 Then
        If Not IsMissing(FieldName) And FieldName <> vbNullString Then
            fnQueryForField = GetField(rsTemp(FieldName))
        Else
            fnQueryForField = GetField(rsTemp.Fields(0))
        End If
    End If
                
    Set rsTemp = Nothing
                
    Exit Function
    
SQLError:
    If IsMissing(sCalledFrom) Then
        sCalledFrom = vbNullString
    End If
        
    fnQueryForField = vbNullString
    
    tfnErrHandler "fnRecordset," & sCalledFrom, SQL, bShowErrow
    
    Set rsTemp = Nothing
    
    On Error GoTo 0
    
End Function

Public Function fnExecSQL(SQL As String, Optional dbLocation As DatabaseLocation = RemoteDB, _
                Optional sCalledFrom As Variant = vbNullString, Optional bShowError As Variant = True) As Boolean

On Error GoTo SQLError
    
    If printSQL Then
        Debug.Print SQL & ";"
    End If

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

Public Function fnDataExists(SQL As String) As Boolean
    Dim rsTemp As Recordset
    
    fnDataExists = fnRecordset(rsTemp, SQL) > 0
    
    Set rsTemp = Nothing
End Function

Public Function CreateTempTable(SQL As String, TableName As String) As Boolean

    CreateTempTable = fnExecSQL(SQL & INTO_TEMP & TableName)
    
End Function

Public Sub subDropTable(TableName As String)
    Dim sSql As String
            
    'In case the table doesn't exist, just continue
    On Error Resume Next
    sSql = SQLParm(SQL_DROP_TABLE, "@table", TableName)
            
    fnExecSQL sSql, , , False
    
End Sub

Public Function fnTableExists(TableName As String, Optional bTemp As Boolean = False) As Boolean
    Dim SQL As String
    Dim rsTemp As Recordset
    
    If bTemp Then
        SQL = SQL_TEMP_TABLE_EXISTS
    Else
        SQL = SQL_TABLE_EXISTS
    End If
    
    SQL = SQLParm(SQL, "@table", TableName)
            
    If bTemp Then
        fnRecordset rsTemp, SQL
        If Not rsTemp Is Nothing Then
            fnTableExists = True
        End If
    Else
        fnTableExists = fnDataExists(SQL)
    End If
                              
End Function

Public Function fnColumnExists(TableName As String, ColumnName As String) As Boolean
    Dim sSql As String
    
    sSql = SQLParm(SQL_COLUMN_EXISTS, _
                        "@table", TableName, _
                        "@column", ColumnName)

    fnColumnExists = fnDataExists(sSql)

End Function

Public Sub EnableButton(Status As ButtonStatus, ParamArray buttons() As Variant)
    Dim Index As Integer
    On Error Resume Next 'Just in case parameter does not have a property Named 'Picture'
    
    For Index = LBound(buttons) To UBound(buttons)
    
        If TypeOf buttons(Index) Is FactorFrame Then
            If Status = Enable Then
                buttons(Index).Picture = frmContext.LoadPicture(SEARCH_UP)
            Else
                buttons(Index).Picture = frmContext.LoadPicture(SEARCH_DOWN)
            End If
            
            buttons(Index).Enabled = CBool(Status)
        End If
        
    Next
    
    On Error GoTo 0

End Sub


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
    Dim buttons As VbMsgBoxStyle
    Dim YesOK   As VbMsgBoxResult
    
    If YesNoType Then
        buttons = vbYesNo
        YesOK = vbYes
    Else
        buttons = vbOKCancel
        YesOK = vbOK
    End If
    If DefaultToNo Then
        buttons = buttons + vbDefaultButton2
    End If
    buttons = buttons + vbQuestion
    
    IsUserSure = (MsgBox(Prompt, buttons) = YesOK)
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
Public Function MsgBox(Prompt, Optional buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal Title, Optional HelpFile, Optional Context) As VbMsgBoxResult

    If SuppressMessageBox Then
        Exit Function
    End If
    
    If IsMissing(Title) Then
        Title = App.Title
    Else
        If StrComp(Left$(Title, Len(App.Title)), App.Title, vbTextCompare) Then
            Title = App.Title & " - " & Title
        End If
    End If
    MsgBox = VBA.MsgBox(Prompt, buttons, Title, HelpFile, Context)
End Function

Public Property Let SuppressMessageBox(suppress As Boolean)
    suppressMsgBox = suppress
End Property

Public Property Get SuppressMessageBox() As Boolean
    SuppressMessageBox = suppressMsgBox
End Property

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
    On Error GoTo 0 'Clear Err & disable error handler
    
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
    On Error GoTo 0 'Clear Err & disable error handler
    
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

Public Function GetField(field As Variant) As String

    If IsNull(field) Then
        GetField = vbNullString
    Else
        GetField = Trim(CStr(field))
    End If
    
End Function

Public Sub SendTab()
    SendKeys "{TAB}"
End Sub

Public Sub ClearText(ParamArray Parms())
    Dim i As Integer
    
    For i = 0 To UBound(Parms)
        If TypeOf Parms(i) Is Textbox Then
            Parms(i).Text = vbNullString
        ElseIf TypeOf Parms(i) Is Label Then
            Parms(i).Caption = vbNullString
        Else
            'Try to reset the .text property...but ignore any error that may occur
            On Error Resume Next
            Parms(i).Text = vbNullString
        End If
    Next

End Sub
Public Sub FileNameParts(ByVal FullFileName As String, Optional ByRef filePath As Variant = vbNullString, Optional ByRef Filename As Variant = vbNullString, Optional ByRef FileExt As Variant = vbNullString)
    Dim pos As Long
    
    '------------------------------------------------------------------------------------
    'Init - Needed for Optional Params that had pre-existing values
    '------------------------------------------------------------------------------------
    filePath = vbNullString
    Filename = vbNullString
    FileExt = vbNullString
    
    '------------------------------------------------------------------------------------
    'Extract the Path
    '------------------------------------------------------------------------------------
    pos = InStrRev(FullFileName, "\")
    If pos = 0 Then
        pos = InStr(FullFileName, ":")
    End If
    If pos Then
        filePath = Left$(FullFileName, pos)
        FullFileName = Mid$(FullFileName, pos + 1)
    End If
    
    '------------------------------------------------------------------------------------
    'Extract the File Extension
    '------------------------------------------------------------------------------------
    pos = InStrRev(FullFileName, ".")
    If pos Then
        FileExt = Mid$(FullFileName, pos + 1)
        FullFileName = Left$(FullFileName, pos - 1)
    End If
    
    '------------------------------------------------------------------------------------
    'Extract the File Name
    '------------------------------------------------------------------------------------
    Filename = FullFileName 'Only thing left is the File NAME itself
End Sub

Public Function CaseInSensitiveString(ByVal S As String, Optional addAsterisk As Boolean = False) As String
    Dim i As Integer
    Dim sRet As String
    Dim sChar As String
    Dim bStartInserted As Boolean
    
    S = Trim(S)
    
    If Trim(S) <> "" Then
        For i = 1 To Len(S)
            sChar = Mid(S, i, 1)
            Select Case sChar
            Case " "
                sRet = sRet + " "
                bStartInserted = False
            Case "\", "?", "*"
                sRet = sRet + "\" + sChar
                bStartInserted = False
            Case Else
                If IsAlphabet(sChar) Then
                    sRet = sRet + "[" + UCase(sChar) + LCase(sChar) + "]"
                ElseIf sChar = "_" Then
                    sRet = sRet + "?"
                ElseIf sChar = "%" Then
                    sRet = sRet + "*"
                    bStartInserted = True
                Else
                    sRet = sRet + sChar
                End If
                
                bStartInserted = False
            End Select
        Next i
    End If
    
    If addAsterisk Then
        sRet = sRet + "*"
    End If
    
    CaseInSensitiveString = sRet
End Function

Private Function IsAlphabet(ByVal sChar As String) As Boolean
    sChar = UCase(sChar)
    
    If sChar >= "A" And sChar <= "Z" Then
        IsAlphabet = True
    End If
End Function

Public Function DPrint(Text As String)
    
    debugCount = debugCount + 1
    Debug.Print " " & Format(debugCount, "##0") & " " & Text

End Function

Public Sub Sleep(milliseconds As Long)
    SleepAPI milliseconds
End Sub

Public Function Align(value As String, just As Justify, Optional length As Integer = 0) As String

Dim wholeLen As Integer
Dim valueLen As Integer
Dim spaceCount As Integer
Dim leftSpace As Integer
Dim rightSpace As Integer

If length > 0 Then
    wholeLen = length
Else
    wholeLen = Len(value)
End If

valueLen = Len(Trim(value))
spaceCount = wholeLen - valueLen

If spaceCount > 0 Then
    If just = LeftJustify Then
        value = Trim(value) & Space(spaceCount)
        
    ElseIf just = RightJustify Then
        value = Space(spaceCount) & Trim(value)
        
    ElseIf just = CenterJustify Then
        
        leftSpace = spaceCount / 2
        
        If spaceCount Mod 2 > 0 Then
            rightSpace = spaceCount - leftSpace
        Else
            rightSpace = leftSpace
        End If
        
        value = Space(leftSpace) & Trim(value) & Space(rightSpace)
    End If
End If

Align = value

End Function

Public Function LockProcess(PROGRAM As String) As Boolean
    Dim SQL As String
    Dim USER As String
    Dim answer As VbMsgBoxResult
    
    SQL = SQLParm(SQL_SELECT_LOCK, "@program", PROGRAM)
        
    USER = fnQueryForField(SQL, "sdl_user")
    
    'If same user, then the process is already locked...give user option to OK bypassing this check.
    If USER = tfnGetUserName() Then
        answer = MsgBox("Process is already locked for this user id (" & USER & ").  Continue?", _
                        vbYesNo + vbExclamation, "Process Locked")
        
        'Keep existing lock
        If answer = vbYes Then
            LockProcess = True
            Exit Function
        End If
    End If
    
    If USER <> vbNullString Then
        MsgBox "Process is being run by " & USER & "." & vbCrLf & _
               "Unable to process at this time"
        
        Exit Function
    Else
        SQL = SQLParm(SQL_INSERT_DATA_LOCK, "@user", tfnGetUserName, "@program", PROGRAM)
        
        
        If fnExecSQL(SQL) Then
            LockProcess = True
        Else
            MsgBox "Unable to lock process, please try again."
        End If
    End If

End Function

Public Function UnlockProcess(PROGRAM) As Boolean
    Dim SQL As String
    
    SQL = SQLParm(SQL_DELETE_LOCK, "@user", tfnGetUserName, "@program", PROGRAM)

    UnlockProcess = fnExecSQL(SQL)

End Function
