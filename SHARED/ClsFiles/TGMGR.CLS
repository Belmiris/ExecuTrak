VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTGManager"
Attribute VB_Creatable = True
Attribute VB_Exposed = False
' clsTGManager
'
' This class manages a collection of True DBGrids.  This is the publicly exported
' class for this OLE server
'
' Michael A. Donahue

Option Explicit

Private Const KEY_CTRL_C = &H3
Private Const KEY_CTRL_V = &H16
Private Const KEY_CTRL_X = &H18

Private Const szROW_STYLE As String = "RowStyle"
Private Const szCURRENT_CELL_STYLE As String = "CurrCellStyle"
Private Const szHEADING_STYLE As String = "Heading"
Private Const szNORMAL_STYLE As String = "Normal"

Private m_styRow As Object            ' row style
Private m_styCurrentCell As Object    ' current cell style
Private m_colTables As Collection     ' collection of grid controls

Private Sub Class_Initialize()

    subClearClass
    
End Sub

Private Sub Class_Terminate()

    subClearClass
    
End Sub

Private Sub subClearClass()

    Set m_colTables = Nothing

End Sub

Property Get ROW_STYLE_ID() As String

    ROW_STYLE_ID = szROW_STYLE

End Property

Property Get CURRENT_CELL_STYLE_ID() As String

    CURRENT_CELL_STYLE_ID = szCURRENT_CELL_STYLE
    
End Property

Property Get HEADING_STYLE_ID() As String

    HEADING_STYLE_ID = szHEADING_STYLE

End Property

Property Get NORMAL_STYLE_ID() As String

    NORMAL_STYLE_ID = szNORMAL_STYLE
    
End Property

Public Function AddTable(frmParent As Object, tblGrid As Object) As Boolean
'
' Add a True DBGrid to the collection and set its default properties
'
    Dim clsData As clsTrueGridData
    Dim szBuf As String
    Dim n As Integer
    
    If frmParent Is Nothing Then
        AddTable = False
        Exit Function
    End If
    
    On Error GoTo AddToCollectionError
    
    If m_colTables Is Nothing Then
        Set m_colTables = New Collection
    End If
    
    ' fill in the data structure
    Set clsData = New clsTrueGridData
        With clsData
            .m_nOriginalRowHeight = tblGrid.RowHeight
            .m_nOriginalRowFontSize = tblGrid.Styles(szNORMAL_STYLE).Font.Size
            .m_nOriginalHeadFontSize = tblGrid.Styles(szHEADING_STYLE).Font.Size
        End With
    
    Set clsData.m_colOriginalColumnWidths = New Collection
    Set clsData.m_colPatterns = New Collection
    
    ' save original column widths
    For n = 0 To tblGrid.Columns.Count - 1
        clsData.m_colOriginalColumnWidths.Add tblGrid.Columns(n).Width
    Next
    
    ' add table data to collection
    Set clsData.m_tblGrid = tblGrid
    Set clsData.m_frmParent = frmParent
    m_colTables.Add clsData, tblGrid.Name

    ' set some properties for the table
    Set m_styRow = tblGrid.Styles.Add(szROW_STYLE)
    Set m_styCurrentCell = tblGrid.Styles.Add(szCURRENT_CELL_STYLE)
    
    ' create highlight row style
    With m_styRow
        .BackColor = &HFF0000
        .ForeColor = &HFFFFFF
    End With
    
    ' create current cell style
    With m_styCurrentCell
        .BackColor = &HFFFFFF
        .ForeColor = &H0
    End With
    
    ' set default properties for the table
    With tblGrid
        .AllowAddNew = False
        .AllowDelete = False
        .AllowUpdate = True
        .Enabled = False
        .MarqueeUnique = True
        .EditDropDown = False
        .RowDividerStyle = dbgBlackLine
        .BackColor = &HFFFFFF
        .ForeColor = &H0
        .HeadLines = 1
        .Appearance = dbg3D
        .BorderStyle = dbgFixedSingle
        .DataMode = dbgBound
        .AllowArrows = True
        .TabAcrossSplits = True
        .WrapCellPointer = False
        .ExposeCellMode = dbgScrollOnEdit
        .TabAction = dbgControlNavigation
        .Splits(0).Locked = False
        .Splits(0).AllowSizing = False
        .Splits(0).AllowFocus = True
        .AllowRowSizing = False
        .ExtendRightColumn = True
        .ScrollBars = dbgAutomatic
        .MarqueeStyle = dbgHighlightRowRaiseCell
        .RecordSelectors = False
    End With
    
    ' set column default properties
    For n = 0 To tblGrid.Columns.Count - 1
        With tblGrid.Columns(n)
            .DividerStyle = dbgBlackLine
            .AllowSizing = False
            .Visible = True
            .WrapText = False
        End With
    Next
    
    ' add the new styles to the table
    tblGrid.AddCellStyle dbgMarqueeRow, m_styRow
    For n = 0 To tblGrid.Columns.Count - 1
        With tblGrid.Columns(n)
            .AddCellStyle dbgMarqueeRow + dbgCurrentCell, m_styCurrentCell
            .AddCellStyle dbgMarqueeRow + dbgUpdatedCell + dbgCurrentCell, m_styCurrentCell
            .AddCellStyle dbgMarqueeRow + dbgUpdatedCell, m_styRow
            .AddCellStyle dbgUpdatedCell, m_styCurrentCell
        End With
    Next

    On Error GoTo 0
    AddTable = True
    Exit Function
    
AddToCollectionError:
    
    On Error GoTo 0
    AddTable = False

End Function

Public Sub BeforeColUpdate(tblGrid As Object, ByVal ColIndex As Integer, OldValue As Variant, CANCEL As Integer)
'
' Fired when a cell has changed, before the internal copy buffer is changed
'
    Dim clsGridData As clsTrueGridData
    Dim bCode As Boolean
    
    ' fetch the grid data item
    Set clsGridData = m_colTables(tblGrid.Name)
    If clsGridData Is Nothing Then
        Exit Sub
    End If

    ' is this cell valid?
    If Not CellIsValid(tblGrid, ColIndex) Then
        
        ' if we are not on the grid anymore, restore the cell and let the user leave
        If Not clsGridData.m_frmParent.ActiveControl Is tblGrid Then
            CANCEL = False
            tblGrid.Columns(ColIndex).Value = OldValue
        Else
            CANCEL = True
        
            ' display the error message if there is one
            If Not clsGridData.m_colErrorMsgs Is Nothing Then
                subShowStatusBarErrorMessage clsGridData.m_lblStatusBar, _
                                             clsGridData.m_colErrorMsgs.Item(ColIndex + 1)
            End If
        End If
    End If
    
End Sub

Public Sub BeforeUpdate(tblGrid As Object, CANCEL As Integer)

    If Not RowIsValid(tblGrid) Then
        CANCEL = True
    End If

End Sub

Public Function CellIsValid(tblGrid As Object, nColumn As Integer) As Boolean
'
' Validates the contents of a cell
'
    Dim clsGridData As clsTrueGridData
    Dim szData As String
    Dim szPattern As String
    Dim nCode As Integer
    
    ' fetch the grid data item
    Set clsGridData = m_colTables(tblGrid.Name)
    If clsGridData Is Nothing Then
        CellIsValid = True
        Exit Function
    End If

    ' get the cell's value
    szData = tblGrid.Columns(nColumn).Text
    
    ' check for blank
    If szData = "" Then
        If fnIsRequired(clsGridData, nColumn) Then
            CellIsValid = False
        Else
            CellIsValid = True
        End If
        Exit Function
    End If
    
    ' check the value against the pattern
    szPattern = fnGetPattern(clsGridData, nColumn)
    
    ' if there is no regular expression, the text is good
    If szPattern = "" Then
        CellIsValid = True
        Exit Function
    End If
        
    ' check the pattern
    nCode = PRegExpMatch(szPattern, Len(szPattern), szData)
    
    ' is it good or not?
    CellIsValid = (nCode = 0)
    
End Function

Public Function Change(tblGrid As Object, Optional CancelBeep As Variant) As Boolean
'
' Check the changed text against the regular expression for that column
'

    Dim szPattern As String
    Dim szText As String
    Dim nCode As Integer
    Dim clsGridData As clsTrueGridData
        
    Change = False
    
    If tblGrid Is Nothing Then
        Exit Function
    End If
    
    ' get collection data for this grid control
    Set clsGridData = m_colTables(tblGrid.Name)
    If clsGridData Is Nothing Then
        Exit Function
    End If
        
    ' get the pattern
    szPattern = fnGetPattern(clsGridData, tblGrid.Col)
    If szPattern = "" Then
        Change = True
        Exit Function
    End If
    
    ' is the column text blank?
    szText = tblGrid.Columns(tblGrid.Col).Text
    If szText = "" Then
        Change = Not fnIsRequired(clsGridData, tblGrid.Col)
        Exit Function
    End If
    
    ' check the pattern
    nCode = PRegExpMatch(szPattern, Len(szPattern), szText)
        
    ' highlight any bad text
    If nCode > 0 Then
        tblGrid.SelStart = nCode - 1
        tblGrid.SelLength = Len(szText) - tblGrid.SelStart
        If tblGrid.SelLength > 0 And Not IsMissing(CancelBeep) Then
            Beep
        End If
    End If
    
    Change = (nCode = 0)    ' true if text is okay
    
End Function

Public Sub GotFocus(tblGrid As Object)
'
'   Table got focus - set the status bar message
'
    Dim clsGridData As clsTrueGridData
    Dim szMessage As String
    
    Set clsGridData = m_colTables(tblGrid.Name)
    
    If clsGridData Is Nothing Then
        Exit Sub
    End If
    
    ' if there are no records, don't do anything
    If Not tblGrid.Enabled Then
        Exit Sub
    End If
    
    ' display the appropriate message for this cell
    If CellIsValid(tblGrid, tblGrid.Col) Then
        szMessage = fnGetStatusMessage(clsGridData, tblGrid.Col)
        subShowStatusBarMessage clsGridData.m_lblStatusBar, szMessage
    Else
        szMessage = fnGetStatusErrorMessage(clsGridData, tblGrid.Col)
        subShowStatusBarErrorMessage clsGridData.m_lblStatusBar, szMessage
    End If

End Sub

Public Function KeyPress(tblGrid As Object, KeyAscii As Integer) As Boolean
'
' Check the key pressed against the regular expression
'
    Dim szPattern As String
    Dim szText As String
    Dim nCode As Integer
    Dim clsGridData As clsTrueGridData
    
    KeyPress = False
    
    If tblGrid Is Nothing Then
        Exit Function
    End If
    
    ' get data for this grid control
    Set clsGridData = m_colTables(tblGrid.Name)
    If clsGridData Is Nothing Then
        Exit Function
    End If
    
    ' get the pattern
    szPattern = fnGetPattern(clsGridData, tblGrid.Col)
    If szPattern = "" Then
        KeyPress = True
        Exit Function
    End If
      
    ' check enter, backspace, escape
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyEscape Then
        KeyPress = True
        Exit Function
    End If
    
    ' check for cut/copy/paste keys
    If KeyAscii = KEY_CTRL_C Or KeyAscii = KEY_CTRL_V Or KeyAscii = KEY_CTRL_X Then
        KeyPress = True
        Exit Function
    End If
        
    ' get the data up to the cursor position and tack on the key pressed
    szText = Left(tblGrid.Columns(tblGrid.Col).Text, tblGrid.SelStart) & Chr(KeyAscii)
    
    nCode = PRegExpMatch(szPattern, Len(szPattern), szText)
    
    If nCode < 0 Or (nCode <> 0 And nCode <> tblGrid.SelStart + 2) Then
        KeyPress = False
        Beep
    Else
        KeyPress = True
    End If

End Function

Public Sub RowColChange(tblGrid As Object, LastRow As Variant, ByVal LastCol As Integer)
'
' sets the status bar message
'
    Dim clsGridData As clsTrueGridData
    Dim szMessage As String
    
    Set clsGridData = m_colTables(tblGrid.Name)
    If clsGridData Is Nothing Then
        Exit Sub
    End If
    
    If Not tblGrid.Enabled Then
        Exit Sub
    End If
    
    ' display the appropriate message
    If CellIsValid(tblGrid, tblGrid.Col) Then
        szMessage = fnGetStatusMessage(clsGridData, tblGrid.Col)
        subShowStatusBarMessage clsGridData.m_lblStatusBar, szMessage
    Else
        szMessage = fnGetStatusErrorMessage(clsGridData, tblGrid.Col)
        subShowStatusBarErrorMessage clsGridData.m_lblStatusBar, szMessage
    End If

End Sub

Public Function RowIsValid(tblGrid As Object) As Boolean
'
' returns True is the current row is entirely valid
'
    Dim clsGridData As clsTrueGridData
    Dim nIdx As Integer
    
    Set clsGridData = m_colTables(tblGrid.Name)
    If clsGridData Is Nothing Then
        RowIsValid = True
        Exit Function
    End If
    
    For nIdx = 0 To tblGrid.Columns.Count - 1
        If Not CellIsValid(tblGrid, nIdx) Then
            RowIsValid = False
            Exit Function
        End If
    Next
    
    RowIsValid = True
    
End Function

Public Sub SelChange(tblGrid As Object, CANCEL As Integer)
'
' cancel column selection
'
    
    CANCEL = True
    
End Sub

Public Function SetColumnPatterns(tblGrid As Object, szarryPatterns() As String) As Boolean
'
'   store regular expressions for the columns
'
    Dim szBuffer As String
    Dim clsGridData As clsTrueGridData
    Dim nIdx As Integer
    Dim nSize As Integer
    
    On Error GoTo ERROR_NEW_COLLECTION
    
    SetColumnPatterns = False
    szBuffer = Space(4096)
    
    ' fetch the grid data item
    Set clsGridData = m_colTables(tblGrid.Name)
    If clsGridData Is Nothing Then
        On Error GoTo 0
        Exit Function
    End If
    
    ' remove the old pattern collection
    If Not (clsGridData.m_colPatterns Is Nothing) Then
        While clsGridData.m_colPatterns.Count > 0
            clsGridData.m_colPatterns.Remove 1
        Wend
    End If
    
    Set clsGridData.m_colPatterns = New Collection
        
    ' add the new patterns
    For nIdx = 1 To tblGrid.Columns.Count
                
        If szarryPatterns(nIdx - 1) <> "" Then
            If InStrB(szarryPatterns(nIdx - 1), Chr(0)) = 0 Then
                nSize = GetRegExpPCode(szarryPatterns(nIdx - 1), szBuffer)
            Else
                szBuffer = szarryPatterns(nIdx - 1)
                nSize = Len(szBuffer)
            End If
            
            If nSize <= 0 Then
                Exit Function
            End If
        
            clsGridData.m_colPatterns.Add Left(szBuffer, nSize)
        
        Else
            clsGridData.m_colPatterns.Add ""
        End If
        
    Next
    
    On Error GoTo 0
    SetColumnPatterns = True
    Exit Function
    
ERROR_NEW_COLLECTION:
    On Error GoTo 0
    SetColumnPatterns = False
    Exit Function
    
End Function

Public Function SetColumnsRequired(tblGrid As Object, szarryReq() As Boolean) As Boolean
'
' store the flags for required columns
'

    Dim clsGridData As clsTrueGridData
    Dim nIdx As Integer
    Dim nSize As Integer
    
    On Error GoTo ERROR_NEW_COLLECTION
    
    SetColumnsRequired = False
    
    ' fetch the grid data item
    Set clsGridData = m_colTables(tblGrid.Name)
    If clsGridData Is Nothing Then
        On Error GoTo 0
        Exit Function
    End If
    
    ' remove the old collection
    If Not (clsGridData.m_colColumnsRequired Is Nothing) Then
        While clsGridData.m_colColumnsRequired.Count > 0
            clsGridData.m_colColumnsRequired.Remove 1
        Wend
    End If
    
    Set clsGridData.m_colColumnsRequired = New Collection
        
    ' add the new patterns
    For nIdx = 1 To tblGrid.Columns.Count
        clsGridData.m_colColumnsRequired.Add szarryReq(nIdx - 1)
    Next
    
    On Error GoTo 0
    SetColumnsRequired = True
    Exit Function
    
ERROR_NEW_COLLECTION:
    On Error GoTo 0
    SetColumnsRequired = False
    Exit Function

End Function

Public Function SetStatusBar(tblGrid As Object, lblStatusBar As Object) As Boolean
'
' set the pointer to the status bar
'
    Dim clsGridData As clsTrueGridData
    
    On Error GoTo ERROR_SET
    
    SetStatusBar = False
    
    ' fetch the grid data item
    Set clsGridData = m_colTables(tblGrid.Name)
    If clsGridData Is Nothing Then
        On Error GoTo 0
        Exit Function
    End If

    Set clsGridData.m_lblStatusBar = lblStatusBar
    SetStatusBar = True
    On Error GoTo 0
    Exit Function
    
ERROR_SET:
    On Error GoTo 0
    SetStatusBar = False
    
End Function

Public Function SetStatusBarErrorMessages(tblGrid As Object, szarryMsgs() As String) As Boolean
'
'   store the status bar error messages
'

    Dim clsGridData As clsTrueGridData
    Dim nIdx As Integer
    
    On Error GoTo ERROR_NEW_COLLECTION

    SetStatusBarErrorMessages = False
    
    ' fetch the grid data item
    Set clsGridData = m_colTables(tblGrid.Name)
    If clsGridData Is Nothing Then
        On Error GoTo 0
        Exit Function
    End If
    
    ' remove the old messages
    If Not (clsGridData.m_colErrorMsgs Is Nothing) Then
        While clsGridData.m_colErrorMsgs.Count > 0
            clsGridData.m_colErrorMsgs.Remove 1
        Wend
    End If
    
    Set clsGridData.m_colErrorMsgs = New Collection
        
    ' add the new messages
    For nIdx = 1 To tblGrid.Columns.Count
       clsGridData.m_colErrorMsgs.Add szarryMsgs(nIdx - 1)
    Next
    
    On Error GoTo 0
    SetStatusBarErrorMessages = True
    Exit Function
    
ERROR_NEW_COLLECTION:
    On Error GoTo 0
    SetStatusBarErrorMessages = False
    Exit Function

End Function

Public Function SetStatusBarMessages(tblGrid As Object, szarryMsgs() As String) As Boolean
'
'   store the status bar messages
'

    Dim clsGridData As clsTrueGridData
    Dim nIdx As Integer
    
    On Error GoTo ERROR_NEW_COLLECTION

    SetStatusBarMessages = False
    
    ' fetch the grid data item
    Set clsGridData = m_colTables(tblGrid.Name)
    If clsGridData Is Nothing Then
        On Error GoTo 0
        Exit Function
    End If
    
    ' remove the old messages
    If Not (clsGridData.m_colStatusMsgs Is Nothing) Then
        While clsGridData.m_colStatusMsgs.Count > 0
            clsGridData.m_colStatusMsgs.Remove 1
        Wend
    End If
    
    Set clsGridData.m_colStatusMsgs = New Collection
        
    ' add the new messages
    For nIdx = 1 To tblGrid.Columns.Count
        clsGridData.m_colStatusMsgs.Add szarryMsgs(nIdx - 1)
    Next
    
    On Error GoTo 0
    SetStatusBarMessages = True
    Exit Function
    
ERROR_NEW_COLLECTION:
    On Error GoTo 0
    SetStatusBarMessages = False
    Exit Function

End Function

Private Function fnGetPattern(clsGridData As Object, nColumn As Integer) As String
'
' retrieves the regular expression for a coulmn
'
    If clsGridData.m_colPatterns Is Nothing Then
        fnGetPattern = ""
    Else
        fnGetPattern = clsGridData.m_colPatterns.Item(nColumn + 1)
    End If
    
End Function

Private Function fnGetStatusMessage(clsGridData As Object, nColumn As Integer) As String
'
' retrieves the status bar message for a column
'
    If clsGridData.m_colStatusMsgs Is Nothing Then
        fnGetStatusMessage = ""
    Else
        fnGetStatusMessage = clsGridData.m_colStatusMsgs.Item(nColumn + 1)
    End If
    
End Function

Private Function fnGetStatusErrorMessage(clsGridData As Object, nColumn As Integer) As String
'
' retrieves the error message for a column
'
    If clsGridData.m_colErrorMsgs Is Nothing Then
        fnGetStatusErrorMessage = ""
    Else
        fnGetStatusErrorMessage = clsGridData.m_colErrorMsgs.Item(nColumn + 1)
    End If
    
End Function

Private Function fnIsRequired(clsGridData As Object, nColumn As Integer) As Boolean
'
' returns true or false if a value is required for this column
'
    If clsGridData.m_colColumnsRequired Is Nothing Then
        fnIsRequired = False
    Else
        fnIsRequired = clsGridData.m_colColumnsRequired.Item(nColumn + 1)
    End If
    
End Function

Private Sub subShowStatusBarMessage(lblStatusBar As Object, szText As String)
'
' display an error message in the status bar
'
    If lblStatusBar Is Nothing Or szText = "" Then
        Exit Sub
    End If
    
    lblStatusBar.ForeColor = STANDARD_TEXT_COLOR
    lblStatusBar.Font.Bold = False
    lblStatusBar.Caption = szText
    lblStatusBar.Refresh

End Sub

Private Sub subShowStatusBarErrorMessage(lblStatusBar As Object, szText As String)
'
' display a message in the status bar
'

    If lblStatusBar Is Nothing Or szText = "" Then
        Exit Sub
    End If
    
    lblStatusBar.ForeColor = ERROR_TEXT_COLOR
    lblStatusBar.Font.Bold = True
    lblStatusBar.Caption = szText
    Beep
    lblStatusBar.Refresh

End Sub

