VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTGSpreadSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***********************************************************'
'
' Copyright (c) 1996 FACTOR, A Division of W.R.Hess Company
'
' Module name   : TGEDITOR.CLS
' Date          : May 16, 1996
' Programmer(s) : Qinggang Ma
'
' This module defines a class implementing unbound editing using true grid.


' An true grid editing is implemented
' Usage:
'      i. Include class files TGEDITOR.CLS in the project
'     ii. Define and create a class: clsTGSpreadSheet
'    iii. Initialization (See sub fnInitializeSpreadSheet in the form)
'     iv. Event calls:
'           a. call .AfterColEdit ColIndex
'           b. call .BeforeColEdit ColIndex, KeyAscii, Cancel
'           c. call .Change
'           d. call .GotFocus
'           e. call .KeyDown KeyCode, Shift
'           f. call .Keypress KeyAscii
'           g. call .LostFocus
'           h. call .RowColChange LastRow, LastCol
'           i. Cancel = True    in the SelChange event
'           j. call .ReadData RowBuf, StartLocation, ReadPriorRows
'       v. Setup public function: fnValidCellValue for validation

Option Explicit
    
#Const REGULAR_EXP_INCLUDED = True

Private Const KEY_P_FINISHED = 0
Private Const KEY_P_OTHER = 1
Private Const KEY_P_CHANGE_ROW = 2

Private Const nCRITICAL = vbCritical + vbOKOnly
Private Const ROW_INCREMENT = 5
Private Const ASC_CTRL_C = 3
Private Const ASC_CTRL_V = &H16
Private Const ASC_CTRL_X = &H18
Private Const vbKeyCtrlY = &H19

'david 04/29/2002
Private Const ASC_CTRL_A = 1
Private Const ASC_CTRL_I = 9
'''''''''''''''''

Private Const STATUS_NOCHANGE = 0
Private Const STATUS_CHANGED = 1
Private Const STATUS_WARNING = 2
Private Const STATUS_VALIDATED = 4

Private Const ROW_INITIAL As Integer = 0
Private Const ROW_ADDED As Integer = 1
Private Const ROW_EDITED As Integer = 2

Private Const BACK_ONCE = 0
Private Const BACK_ALWAYS = 1

Private Const NOCURRENT = 1
Private Const NEWRECORD = 2
Const MAX_CONSTRAINTS = 20
Const KEYS_NOCURRENT = "nocurrent"
Const KEYS_NEW = "new"
Const OPR_EQUAL = "="
Const OPR_GREATER = ">"
Const OPR_GREATER_EQUAL = ">="
Const OPR_LESS = "<"
Const OPR_LESS_EQUAL = "<="

Private Const DB_INFORMIX As Integer = 1
Private Const DB_ACCESS As Integer = 2

Private Const RETURN_KEY_NO_ACTION = 0
Private Const RETURN_KEY_RIGHT = 1
Private Const RETURN_KEY_DOWN = 2

Const szROW_STYLE As String = "RowStyle"
Const szCURRENT_CELL_STYLE As String = "CurrCellStyle"
Const STYLE_INITIAL As Integer = 0
Const STYLE_PLAIN As Integer = 1
Const STYLE_HIGHLIGHT As Integer = 2

Private Type tpConstraints
    nColumn As Integer
    sOperator As String
    vValue As Variant
End Type

Private Type tpESCapable
    nColumn As Integer
    nColumnFrom As Integer
End Type

Private nMaxCol As Integer           'Number of columns
Private nRowCount As Long              'Number of rows
Private nMaxRow As Long

Private GridArray() As Variant      'Array to store the data - our data source
Private nCellChanged() As Integer
Private bValidCell() As Boolean

Private nEditColumns() As Integer
Private sKeyFields() As Variant
Private nKeyColumns() As Integer
Private bmChangedRows() As Variant

Private sPatternCode() As String
Private sPrompt() As String
Private sErrMessage() As String
Private nColumnType() As Integer
Private bFieldRequired() As Boolean
Private bCheckRegExp() As Boolean

Private sDefaultValue() As String
Private sDisplayFormat() As String

Private nESCColumns() As tpESCapable

Private sDownArrowMessage As String
Private nColumnCount As Integer
Private nFlagColumn  As Integer
Private nBookMarkColumn As Integer
Private nChangedCount As Long
Private nMaxChanged As Long
Private lEditingRow As Long
Private lBackFirstRow As Integer
Private nLastColumn As Integer
Private nLastTableRow As Integer
Private nColumnForNewRow As Integer
Private lBookMarkCount As Long
Private nKeyColumnCount As Integer
Private lKeyRows As Long
Private lMaxKeys As Long
Private lMaxRows As Long

Private nOriginalTableWidth As Integer
Private nOriginalTableHeight As Integer
Private nFontSize As Integer
Private nESCCount As Integer
Private nInputKey As Integer
Private nReturnKeyAction As Integer
Private m_nOldValue As Variant

Private nKeyProcessing As Integer
Private nAddRowCount As Integer
Private bRowChanging As Boolean
Private bValidating As Boolean
Private bCellEditing As Boolean
Private bForcedEditing As Boolean
Private bAddingNewRow As Boolean
Private bClearData As Boolean
Private bErrorFlag As Boolean
Private bDataChanged As Boolean
Private bFirstTimeGotFocus As Boolean
Private m_bKeepData As Boolean

Private bHighlightInvalidCell As Boolean
Private bCtrlYClearRow As Boolean
Private bCtrlYClearLocked As Boolean
Private bLocked As Boolean
Private bAllowAddNew As Boolean
Private m_bAllowRevalidate As Boolean

Private tgTable As TDBGrid
Private lblStatusbar As Control
Private frmParent As Form
Private engDBEngine As DBEngine

Private m_nRowHighLighted As Integer

Public AddingValidate As Boolean

'david 01/15/2001
'variables for Column Sort enhancement
Private Const m_NUMERIC_TYPE As Byte = 100
Private Const m_STRING_TYPE As Byte = 101
Private Const m_DATE_TYPE = 106
Private Const m_ASCENDING As Byte = 102
Private Const m_DESCENDING As Byte = 103
Private Const m_CASE_SENSITIVE As Byte = 104
Private Const m_CASE_INSENSITIVE As Byte = 105

Private Type typSortKey
    nColIndex As Byte
    nColType As Byte
    nSortOrder As Byte
    nCaseSensitive As Byte
    nInitSortOrder As Byte
End Type

Private Const MAX_SORT_KEYS As Integer = 5
Private m_arySortCol() As typSortKey
Private m_SortColCount As Integer

Private Const nCOL_IN_GRID As Integer = 0
Private Const nCOL_KEYS_COUNT As Integer = 1
Private Const nCOL_SORT_ENABLED As Integer = 2
Private Const MAX_COL_arySortKeysCount As Integer = 2

Private m_arySortKeysCount() As Integer
Private m_bSortOnCol As Boolean
Private m_bSortWhenGridNotvalid As Boolean

'david 03/21/2001
Private bIn_SetRowHighLighted As Boolean

'david 04/04/2001
Private Const ERROR_TEXT_COLOR As Long = &HFF&
Private Const CORRECT_TEXT_COLOR As Long = &H8000&
Private Const STANDARD_TEXT_COLOR As Long = &H0&
Private Const szEMPTY As String = ""
'
'david 04/26/2002
Private m_bColIsSorted As Boolean
Private aryHighlightedID() As Byte
Private lCountHighlightedID As Long
'''''''''''''''''

'david 06/05/2002
'refresh the text of a cell automatically
'when the content is changed in CellValue() let property.
Private m_bCellValueAutoRebind As Boolean
'''''''''''''''''

'david 08/09/2002
'reversed sort column on subsequence HeadClick()
Private m_bReverseSortOnCol As Boolean
'''''''''''''''''

'david 02/09/2004  #401172-1
Private m_tgsSelector As Object
''''''''''''''''''''''''''''

'david 03/23/2006  #513569
Private m_bIgnoreCellValueError As Boolean
''''''''''''''''''''''''''
'

Public Sub AddEditColumn(ByVal nIndex As Integer, _
                         Optional sMsg As Variant, _
                         Optional vPtn As Variant, _
                         Optional vFormat As Variant, _
                         Optional vRequired As Variant, _
                         Optional nColType As Variant)
    If nIndex >= nMaxCol Or nColumnCount >= nMaxCol Then
        Exit Sub
    End If
    
    Dim i As Integer
    Dim j As Integer
    
    bClearData = True
    
    If nColumnCount = 0 Then
        If nIndex < tgTable.Columns.Count Then
            tgTable.Col = nIndex
            nLastColumn = nIndex
        End If
        If nColumnForNewRow < 0 Then
            nColumnForNewRow = nIndex
        End If
    Else
        If nIndex < tgTable.Col Then
            tgTable.Col = nIndex
            nLastColumn = nIndex
        End If
    End If
    
    bClearData = False
    
    For i = 0 To nColumnCount - 1
        If nIndex = nEditColumns(i) Then
            If nIndex < tgTable.Columns.Count Then
                tgTable.Columns(nIndex).Locked = False
            End If
            Exit Sub
        End If
    Next
    
    i = 0
    
    While i < nColumnCount And nIndex > nEditColumns(i)
        i = i + 1
    Wend
    
    For j = nColumnCount To i + 1 Step -1
        nEditColumns(j) = nEditColumns(j - 1)
        sPrompt(j) = sPrompt(j - 1)
        sPatternCode(j) = sPatternCode(j - 1)
        sDefaultValue(j) = sDefaultValue(j - 1)
        nColumnType(j) = nColumnType(j - 1)
        bFieldRequired(j) = bFieldRequired(j - 1)
        bCheckRegExp(j) = bCheckRegExp(j - 1)
    Next

    nEditColumns(i) = nIndex
    
    If IsMissing(sMsg) Then
        sPrompt(i) = ""
    Else
        sPrompt(i) = sMsg
    End If
    
    If IsMissing(vPtn) Then
        sPatternCode(i) = ""
    Else
        Dim sPtn As String
        sPtn = vPtn
        sPatternCode(i) = fnGetPatternCode(sPtn)
    End If
    
    If IsMissing(vRequired) Then
        bFieldRequired(i) = True
    Else
        bFieldRequired(i) = vRequired
    End If
    
    If Not IsMissing(vFormat) Then
        sDisplayFormat(nIndex) = vFormat
    End If
    
    If IsMissing(nColType) Then
        nColumnType(i) = BACK_ONCE
    Else
        nColumnType(i) = nColType
    End If
    
    bCheckRegExp(i) = True
    nColumnCount = nColumnCount + 1
    
    If Not m_bKeepData Then
        If nColumnCount = 1 Then
            SetupTable True
        End If
    End If
    
    With tgTable.Columns(nIndex)
        .Locked = False
        .FetchStyle = True
    End With
    
    If Not m_bKeepData Then
        fnInitialParams
    End If
End Sub

Property Get AddedRow() As Integer
    AddedRow = ROW_ADDED
End Property

Public Sub AddESCColumn(ByVal nCol As Integer, nFrom As Integer)
    If nCol < 0 Or nCol >= nMaxCol Then
        Exit Sub
    End If
    
    If nESCCount = 0 Then
        ReDim nESCColumns(0)
    Else
        ReDim Preserve nESCColumns(nESCCount)
    End If
    Dim i As Integer
    For i = 0 To nESCCount - 1
        If nCol = nESCColumns(i).nColumn Then
            If nFrom = nESCColumns(i).nColumnFrom Then
                Exit Sub
            End If
        End If
    Next
    
    nESCColumns(nESCCount).nColumn = nCol
    nESCColumns(nESCCount).nColumnFrom = nFrom
    nESCCount = nESCCount + 1
End Sub

Public Function AddHiddenField(sKey As String) As Integer
    Dim i As Integer
    Dim sField As String
    sField = LCase(Trim(sKey))
    For i = tgTable.Columns.Count - 1 To 0 Step -1
        If LCase(Trim(tgTable.Columns(i).DataField)) = sField Then
            AddHiddenField = i
            Exit Function
        End If
    Next
    
    tgTable.Columns.Add tgTable.Columns.Count
    nMaxCol = tgTable.Columns.Count
    tgTable.Columns(nMaxCol - 1).DataField = LCase(Trim(sField))
    tgTable.Columns(nMaxCol - 1).Visible = False
    tgTable.Columns(nMaxCol - 1).Width = -100

    nFlagColumn = nMaxCol
    nBookMarkColumn = nMaxCol + 1
    AddHiddenField = nMaxCol - 1
    ReDim GridArray(nMaxCol + 1, 0)
    ReDim Preserve sDisplayFormat(nMaxCol - 1)
    sDisplayFormat(nMaxCol - 1) = ""
End Function

Public Sub AddKeyFields(sFields As String)
    Dim sTemp(20) As String
    Dim nCol As Integer

    fnParseString sTemp, sFields
    nKeyColumnCount = 0
    While sTemp(nKeyColumnCount) <> ""
        nCol = fnIndexFromField(sTemp(nKeyColumnCount))
        If nCol < 0 Then
            MsgBox "Key fields not in the table"
        Else
            ReDim Preserve nKeyColumns(nKeyColumnCount)
            nKeyColumns(nKeyColumnCount) = nCol
        End If
        nKeyColumnCount = nKeyColumnCount + 1
    Wend
    fnInitialParams
End Sub

Property Let AddNewErrorMessage(sMsg As String)
    sDownArrowMessage = sMsg
End Property

Public Sub AfterColEdit(ByVal ColIndex As Integer)
    Dim nCol As Integer
    Dim lRow As Long

    If IsNull(tgTable.Row) Then
        Exit Sub
    Else
        If Val(tgTable.Row) < 0 Then
            Exit Sub
        End If
    End If
    
    nCol = fnColumnNumber(ColIndex)
    lRow = GetCurrentRowNumber
    If lEditingRow < 0 Then
        lEditingRow = lRow
    End If
    
    'david 03/21/2001 WORK AROUND!!!
    'fix the bug of the last row cellvalue being reset to zero
    If bIn_SetRowHighLighted Then
        tgTable.Columns(ColIndex).Text = fnGetUserData(MakeBookmark(lRow), ColIndex)
        bIn_SetRowHighLighted = False
        Exit Sub
    End If
    
    'david 12/06/00 commented out
    'm_nOldValue = GridArray(ColIndex, lRow)
    If fnGetUserData(MakeBookmark(lRow), ColIndex) = fnCellValue(ColIndex, lRow) Then
        lEditingRow = -1
        Exit Sub
    End If
    
    DataEdited = True
    fnSaveBookMarks Bookmark(lRow)
    GridArray(ColIndex, lRow) = fnCellValue(ColIndex, lRow)
    If sDisplayFormat(ColIndex) <> "" Then
        tgTable.Columns(ColIndex).Text = fnGetUserData(MakeBookmark(lRow), ColIndex)
    End If
    
    If nCol >= 0 Then
        nCellChanged(nCol, lRow) = nCellChanged(nCol, lRow) And (Not STATUS_VALIDATED) Or STATUS_CHANGED
        GridArray(nFlagColumn, lRow) = ROW_EDITED
        bErrorFlag = fnValidate(ColIndex, lRow, fnCellValue(ColIndex, lRow))
        If Not bAllowAddNew Then
            If ColIndex = LastEditableCol Then
                If lRow = nRowCount - 1 Then
                    If bErrorFlag Then
                        nCellChanged(nCol, lRow) = nCellChanged(nCol, lRow) Or STATUS_WARNING
                    End If
                    fnShowStatus lRow, ColIndex
                End If
            End If
        End If
    End If
    
    'david 02/09/2004  #401172-1
    If Not m_tgsSelector Is Nothing Then
        m_tgsSelector.ClearSelectedCells
    End If
    ''''''''''''''''''''''''''''
    
    bForcedEditing = False
    lEditingRow = -1
    SetRowHighLighted True, True
End Sub

Property Let AllowAddNew(bStatus As Boolean)
    If bAllowAddNew Then
        nRowCount = RowCount
    End If
    bAllowAddNew = bStatus
End Property

Property Get AllowAddNew() As Boolean
    AllowAddNew = bAllowAddNew
End Property

Property Let AllowRevalidate(bTemp As Boolean)
    m_bAllowRevalidate = bTemp
End Property

'david 02/09/2004  #401172-1
Property Set SelectorClass(tgsSelector As Object)
    Set m_tgsSelector = tgsSelector
End Property
''''''''''''''''''''''''''''

Public Sub Append(vData() As Variant)
    InsertRow vData, nRowCount
    If IsNull(tgTable.FirstRow) Then
        Exit Sub
    End If
    
    If tgTable.FirstRow + tgTable.Row >= nRowCount - 1 Then
        tgTable.Rebind
    End If
End Sub

Property Get BackAlways() As Integer
    BackAlways = BACK_ALWAYS
End Property

Property Get BackOnce() As Integer
    BackOnce = BACK_ONCE
End Property

Public Sub BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    
    DataEdited = False
    If bLocked Then
        Cancel = True
        Exit Sub
    End If
    If IsNull(tgTable.Row) Then
        Cancel = True
        Exit Sub
    Else
        If Val(tgTable.Row) < 0 Then
            Cancel = True
            Exit Sub
        End If
    End If

    Dim i As Integer
    Dim nCol As Integer
    Dim sText As String

    bCellEditing = True
    nCol = fnColumnNumber(ColIndex)
    If nCol < 0 Then
        Exit Sub
    End If

    subAddNewRow ColIndex
    
    lEditingRow = GetCurrentRowNumber()
    If lEditingRow < 0 Then
        Exit Sub
    End If
    
    fnSaveChangedKeys
'    fnSaveBookMarks Bookmark(GetCurrentRowNumber)
    'david 01/24/2001
    If KeyAscii <> -1 Then
        m_nOldValue = GridArray(ColIndex, lEditingRow)
    Else
        KeyAscii = 0
    End If
    
    nCellChanged(nCol, lEditingRow) = STATUS_CHANGED
    GridArray(nFlagColumn, lEditingRow) = ROW_EDITED
    
    With tgTable.Columns(ColIndex)
        If .ValueItems.Count > 0 Then
            If nInputKey < 0 Then
                i = fnNextValueItemIndex(ColIndex)
            Else
                i = nInputKey
                nInputKey = -1
            End If
            If i >= 0 Then
                GridArray(ColIndex, lEditingRow) = .ValueItems.Item(i).value
                nCellChanged(nCol, lEditingRow) = nCellChanged(nCol, lEditingRow) And (Not STATUS_VALIDATED) Or STATUS_CHANGED
                GridArray(nFlagColumn, lEditingRow) = ROW_EDITED
                bErrorFlag = fnValidate(ColIndex, lEditingRow, GridArray(ColIndex, lEditingRow))
                fnSaveBookMarks Bookmark(lEditingRow)
            End If
        End If
    End With
    If nKeyColumnCount > 0 Then
        If IsNull(sKeyFields(nKeyColumnCount, lKeyRows)) Or sKeyFields(nKeyColumnCount, lKeyRows) = "" Then
            sKeyFields(nKeyColumnCount, lKeyRows) = "  "
            For i = 0 To nKeyColumnCount - 1
                sKeyFields(i, lKeyRows) = GridArray(nKeyColumns(i), lEditingRow)
            Next
        End If
    End If
End Sub

Property Get BlockRowColChange() As Boolean
    BlockRowColChange = bClearData
End Property

Property Let BlockRowColChange(bStatus As Boolean)
    bClearData = bStatus
End Property

Property Let CtrlYClearLocked(bTemp As Boolean)
    bCtrlYClearLocked = bTemp
End Property

Property Get DataChanged()
    DataChanged = STATUS_CHANGED
End Property

Property Get CellValue(ByVal nCol As Integer, ByVal lRow As Long) As Variant
    If nRowCount < 0 Then
        Exit Property
    End If
    If nCol >= nMaxCol Or nCol < 0 Or lRow > nRowCount Or lRow < 0 Then
        Exit Property
    End If
    If Not IsNull(GridArray(nCol, lRow)) Then
        CellValue = GridArray(nCol, lRow)
    Else
        CellValue = ""
    End If
End Property

Public Sub ClearData()
    Dim bFlag As Boolean
    Dim i As Integer
    Dim j As Integer
    
    If tgTable Is Nothing Then
        Exit Sub
    End If
    
    'david 02/09/2004  #401172-1
    If Not m_tgsSelector Is Nothing Then
        m_tgsSelector.ClearSelectedCells
    End If
    ''''''''''''''''''''''''''''
    
    ReDim GridArray(nMaxCol + 1, 0)
    lBookMarkCount = 0
    GridArray(nBookMarkColumn, 0) = fnGetNewBookmark   'MakeBookmark(0)
    
    If m_bKeepData Then
        For j = 0 To nMaxRow - 1
            For i = 0 To nColumnCount - 1
                bValidCell(i, j) = False
                nCellChanged(i, j) = 0
            Next i
        Next j
    Else
        If nColumnCount > 0 Then
            ReDim bValidCell(nColumnCount - 1, 0)
            ReDim nCellChanged(nColumnCount - 1, 0)
        End If
    End If
    ReDim bmChangedRows(0)

    nChangedCount = 0
    nMaxChanged = 0
    nRowCount = 0
    nMaxRow = 0
    nLastColumn = 0
    nLastTableRow = 0
    lEditingRow = -1
    bAddingNewRow = False
    AddingValidate = False
    
    If FirstEditableCol < tgTable.Columns.Count Then
        tgTable.Col = FirstEditableCol
    End If
    
    bFlag = bClearData
    bClearData = True
    tgTable.Refresh
    DoEvents
    bClearData = bFlag

    'david 08/10/2002
    'restore the original sort order
    subResetSortOrder
    '''''''''''''''''
End Sub

Property Let ColumnForNewRow(nCol As Integer)
    nColumnForNewRow = nCol
End Property

Property Get ColumnForNewRow() As Integer
    ColumnForNewRow = nColumnForNewRow
End Property

Property Get ColumnNumber(sKey As String)
    ColumnNumber = fnIndexFromField(sKey)
End Property

Property Let CtrlYClearRow(bTemp As Boolean)
    bCtrlYClearRow = bTemp
End Property

Property Get DataEdited() As Boolean
    DataEdited = bDataChanged
End Property

Property Let DataEdited(bTemp As Boolean)
    bDataChanged = bTemp
End Property

Property Let DisplayFormat(ByVal nCol As Integer, sFormat As String)
    If nCol >= 0 And nCol < nMaxCol Then
        sDisplayFormat(nCol) = sFormat
    End If
End Property

Property Get DisplayFormat(ByVal nCol As Integer) As String
    If nCol > 0 And nCol < nMaxCol Then
        DisplayFormat = sDisplayFormat(nCol)
    End If
End Property

Property Set engFactor(objTemp As DBEngine)
    Set engDBEngine = objTemp
End Property

Property Let ErrorMessage(nCol As Integer, sMsg As String)
    Dim nFCol As Integer
    nFCol = fnColumnNumber(nCol)

    If nFCol >= 0 Then
        sErrMessage(nFCol) = sMsg
    End If
End Property

Property Let CheckRegExp(nCol As Integer, bFlag As Boolean)
    Dim nFCol As Integer
    nFCol = fnColumnNumber(nCol)

    If nFCol >= 0 Then
        bCheckRegExp(nFCol) = bFlag
    End If
End Property

Public Sub FetchCellStyle(ByVal Condition As Integer, _
                          ByVal Split As Integer, _
                          Bookmark As Variant, _
                          ByVal Col As Integer, _
                          ByVal CellStyle As Object)
    Dim nECol As Integer
    Dim lRow As Long
    
    nECol = fnColumnNumber(Col)
    If nECol >= 0 Then
        lRow = Val(Bookmark)
        If Not bValidCell(nECol, lRow) Then
            If (nCellChanged(nECol, lRow) And STATUS_WARNING) = STATUS_WARNING Then
                CellStyle.ForeColor = vbWhite
                CellStyle.BackColor = vbRed
            End If
        End If
    End If
End Sub

Public Property Get FirstRow() As Long
    If IsNull(tgTable.FirstRow) Then
        FirstRow = 0
    Else
        FirstRow = Val(tgTable.FirstRow)
    End If
End Property

Public Sub FirstRowChange()
    If lBackFirstRow > FirstRow Then
        nLastTableRow = nLastTableRow + 1
    ElseIf lBackFirstRow < FirstRow Then
        nLastTableRow = nLastTableRow - 1
        If nLastTableRow < 0 Then
            nLastTableRow = 0
        End If
    End If
    lBackFirstRow = FirstRow
End Sub

Property Get EditColumnCount() As Integer
    EditColumnCount = nColumnCount
End Property

Private Function fnCellValue(ByVal nCol As Integer, _
                             ByVal lRow As Long) As Variant
    Dim sVal As Variant
    Dim lCurrRow As Long
    
    On Error Resume Next
    If IsNull(tgTable.FirstRow) Then
        lCurrRow = lRow
    Else
        lCurrRow = lRow - Val(tgTable.FirstRow)
    End If
    With tgTable.Columns(nCol)
        If lCurrRow = tgTable.Row Then
            sVal = .value
        Else
            sVal = GridArray(nCol, lRow)
        End If
        If .ValueItems.Count > 0 Then
            Dim i As Integer
            For i = 0 To .ValueItems.Count - 1
                If .ValueItems(i).value = .value Then
                    sVal = .ValueItems(i).value
                    Exit For
                End If
            Next
        End If
    End With

    If IsNull(sVal) Then
        fnCellValue = ""
    Else
        fnCellValue = sVal
    End If
End Function

Private Sub fnChangeToEditingCol(ByVal nLastCol As Integer)
    If nColumnCount <= 0 Then
        Exit Sub
    End If

    Dim nColL As Integer
    Dim nColC As Integer
    Dim i As Integer
    
    nColL = fnColumnNumber(nLastCol)
    nColC = fnColumnNumber(tgTable.Col)
    If nLastCol < tgTable.Col Then
        If nColL < nColumnCount - 1 Then
            i = tgTable.Col
            While nColC < 0 And i < nMaxCol
                i = i + 1
                nColC = fnColumnNumber(i)
            Wend
            If i < nMaxCol And nColC >= 0 Then
                If i < tgTable.Columns.Count Then
                    tgTable.Col = i
                End If
            Else
                If nColC < 0 Then
                    If nLastCol < tgTable.Columns.Count Then
                        tgTable.Col = nLastCol
                    End If
                End If
            End If
        Else
            If fnGotoNextRow Then
                If FirstEditableCol < tgTable.Columns.Count Then
                    tgTable.Col = FirstEditableCol
                End If
            Else
                If nLastColumn < tgTable.Columns.Count Then
                    tgTable.Col = nLastColumn
                End If
            End If
        End If
    ElseIf nLastCol > tgTable.Col Then
        If nColL > 0 Then
            i = tgTable.Col
            While nColC < 0 And i > 0
                i = i - 1
                nColC = fnColumnNumber(i)
            Wend
            If i < nMaxCol And nColC >= 0 Then
                If i < tgTable.Columns.Count Then
                    tgTable.Col = i
                End If
            Else
                If nColC < 0 Then
                    If nLastCol < tgTable.Columns.Count Then
                        tgTable.Col = nLastCol
                    End If
                End If
            End If
        Else
            'david 12/11/2002
            'WHY GOTO THE PREVIOUS ROW ???
            'THE CAUSE PROBLEM WHEN IMPLEMENTING THE POPUP ON DOUBLE CLICKING A ROW!!!
'            If fnGotoPreviousRow Then
'                If LastEditableCol < tgTable.Columns.Count Then
'                    tgTable.Col = LastEditableCol
'                End If
'            Else
'                If nLastCol < tgTable.Columns.Count Then
'                    tgTable.Col = nLastCol
'                End If
'            End If
            tgTable.Col = FirstEditableCol
        End If
    End If

End Sub

Private Sub fnCheckForError(ByVal lRow As Long)
    Const LARGE_NUMBER = 10000
    
    Dim i As Integer
    Dim nCol As Integer
    Dim nECol As Integer
    Dim vTemp As Variant
    Dim sTemp As String
    Dim nRetCol As Integer

    fnValidate nLastColumn, lRow, fnGetUserData(MakeBookmark(lRow), nLastColumn)

    nCol = fnColumnNumber(nLastColumn)
    bErrorFlag = False
    If nCol >= 0 Then
        If (nCellChanged(nCol, lRow) And STATUS_WARNING) <> STATUS_WARNING Or nColumnType(nCol) = BACK_ALWAYS Then
            bErrorFlag = Not bValidCell(nCol, lRow)
            nCellChanged(nCol, lRow) = nCellChanged(nCol, lRow) Or STATUS_WARNING
        End If
    End If
    If bErrorFlag Then
        Exit Sub
    End If
    
    i = 0
    nRetCol = LARGE_NUMBER
    While Not bErrorFlag And i < nESCCount
        If nLastColumn = nESCColumns(i).nColumn Then
            nECol = nESCColumns(i).nColumnFrom
            fnValidate nECol, lRow, fnGetUserData(MakeBookmark(lRow), nECol)
            nCol = fnColumnNumber(nECol)
            
            If nCol >= 0 Then
                If ((nCellChanged(nCol, lRow) And STATUS_WARNING) <> STATUS_WARNING) Or nColumnType(nECol) = BACK_ALWAYS Then
                    bErrorFlag = Not bValidCell(nCol, lRow)
                    nCellChanged(nCol, lRow) = nCellChanged(nCol, lRow) Or STATUS_WARNING
                End If
            Else
                bErrorFlag = False
            End If
            If bErrorFlag Then
                If nRetCol > nESCColumns(i).nColumnFrom Then
                    nRetCol = nESCColumns(i).nColumnFrom
                End If
            End If
        End If
        i = i + 1
    Wend
    
    If nRetCol < LARGE_NUMBER Then
        nLastColumn = nRetCol     'nEditColumns(nCol)
    End If
End Sub

Private Sub fnCheckForValueItem(nKey As Integer)
    Dim i As Integer
    Dim sKey As String * 1
    
    sKey = UCase(Chr(nKey))
    nInputKey = -1
    With tgTable.Columns(tgTable.Col).ValueItems
        If .Count > 0 Then
            For i = 0 To .Count - 1
                If sKey = UCase(Left(.Item(i).DisplayValue, 1)) Then
                    nInputKey = i
                    Exit For
                End If
            Next
        End If
    End With
End Sub

Private Function fnCheckRowValid(ByVal nCol As Integer, _
                                 ByVal lRow As Long) As Boolean
    Dim szText As String
    
    fnCheckRowValid = False
    fnValidate nEditColumns(nCol), lRow, fnCellValue(nEditColumns(nCol), lRow)
    
    If bValidCell(nCol, lRow) Then
        fnCheckRowValid = True
    Else
        If Not bLocked Then
            fnShowStatus lRow, nEditColumns(nCol), True
            If bAllowAddNew Then
                'david 12/06/00
                'if red error message then Beep
                If lblStatusbar.ForeColor = &HFF Then
                    Beep
                End If
            End If
            nAddRowCount = nAddRowCount + 1
            If bHighlightInvalidCell Then
                If nCol = nColumnCount - 1 Then
                    bForcedEditing = True
                    BeforeColEdit nEditColumns(nCol), 0, False
                    tgTable.EditActive = True
                    tgTable.SelStart = 0
                    tgTable.SelLength = Len(szText)
                End If
            End If
        End If
        fnCheckRowValid = False
    End If
End Function

Private Function fnComapreVariant(vTemp1 As Variant, vTemp2 As Variant) As Integer
    'Returns -1, if vTemp1 < vTemp2
    'Returns 0, if vTemp1 = vTemp2
    'Returns 1, if vTemp1 > vTemp2
    If IsDate(vTemp1) And IsDate(vTemp2) Then
        Dim dTemp1 As Date
        Dim dTemp2 As Date
        dTemp1 = CDate(vTemp1)
        dTemp2 = CDate(vTemp2)
        If dTemp1 < dTemp2 Then
            fnComapreVariant = -1
        ElseIf dTemp1 > dTemp2 Then
            fnComapreVariant = 1
        Else
            fnComapreVariant = 0
        End If
    ElseIf IsNumeric(vTemp1) And IsNumeric(vTemp2) Then
        Dim fTemp1 As Double
        Dim fTemp2 As Double
        fTemp1 = CDbl(vTemp1)
        fTemp2 = CDbl(vTemp2)
        If Abs(fTemp1 - fTemp2) < 0.000001 Then
            fnComapreVariant = 0
        ElseIf fTemp1 < fTemp2 Then
            fnComapreVariant = -1
        Else
            fnComapreVariant = 1
        End If
    Else
        Dim sTemp1 As String
        Dim sTemp2 As String
        sTemp1 = Trim(vTemp1)
        sTemp2 = Trim(vTemp2)
        If sTemp1 < sTemp2 Then
            fnComapreVariant = -1
        ElseIf sTemp1 > sTemp2 Then
            fnComapreVariant = 1
        Else
            fnComapreVariant = 0
        End If
    End If
End Function

Private Function fnEscapable(ByVal nColFrom As Integer, nColTo As Integer) As Boolean
    Dim i As Integer
    
    fnEscapable = False
    For i = 0 To nESCCount - 1
        If nColTo = nESCColumns(i).nColumn Then
            If nESCColumns(i).nColumnFrom = nColFrom Then
                fnEscapable = True
                Exit For
            End If
        End If
    Next
End Function

Private Function fnGotoNextColumn() As Boolean
    Dim nCol As Integer
    
    fnGotoNextColumn = False
    nCol = fnColumnNumber(tgTable.Col)
    If nCol >= 0 Then
        nCol = nCol + 1
        If nCol < nColumnCount Then
            tgTable.Col = nEditColumns(nCol)
            fnGotoNextColumn = True
        End If
    End If
    
End Function

Private Function fnGotoPrevColumn() As Boolean
    Dim nCol As Integer
    
    fnGotoPrevColumn = False
    nCol = fnColumnNumber(tgTable.Col)
    If nCol >= 0 Then
        nCol = nCol - 1
        If nCol >= 0 Then
            tgTable.Col = nEditColumns(nCol)
            fnGotoPrevColumn = True
        End If
    End If
    
End Function

Private Function fnIndexFromField(sField As String) As Integer
    Dim i As Integer
    Dim sKey As String
    
    fnIndexFromField = -1
    sKey = LCase(sField)
    With tgTable.Columns
        For i = 0 To .Count - 1
            If LCase(Trim(.Item(i).DataField)) = sKey Then
                fnIndexFromField = i
                Exit Function
            End If
        Next
    End With
    #If DEVELOP Then
        MsgBox "The field(" & sField & ") you passed in is not in the grid"
    #End If

End Function

Private Function fnRowNumberByIndex(Bookmk As Variant) As Long
    Dim Index As Long
    Index = -9999
    For Index = 0 To nRowCount - 1
        If Bookmk = Bookmark(Index) Then
            Exit For
        End If
    Next
    
    fnRowNumberByIndex = Index

End Function

Private Sub fnSetEditProperties()

    Dim szBuf As String
    Dim n As Integer

    bAllowAddNew = True

    ' set column default properties
    For n = 0 To tgTable.Columns.Count - 1
        With tgTable.Columns(n)
            .DividerStyle = dbgBlackLine
            .AllowSizing = False
            .WrapText = False
            .Locked = True
        End With
    Next
    
    ' set default properties for the table
    With tgTable
        .AllowAddNew = False
        .DataMode = dbgUnbound
        .MarqueeUnique = True
        .EditDropDown = False
        .ExposeCellMode = dbgScrollOnEdit
        .TabAction = dbgControlNavigation
        .AllowRowSizing = False
        .ExtendRightColumn = True
        .ScrollBars = dbgAutomatic
        If nColumnCount > 0 Then
            .MarqueeStyle = dbgHighlightRowRaiseCell
        Else
            .MarqueeStyle = dbgHighlightRow
        End If
        .RecordSelectors = False
    End With
    
    SetRowHighLighted True
    
End Sub

Private Function fnNextValueItemIndex(nCol As Integer)
    Dim i As Integer
    Dim vTemp As Variant
    
    fnNextValueItemIndex = -1
    vTemp = tgTable.Columns(nCol).value
    With tgTable.Columns(nCol).ValueItems
        For i = 0 To .Count - 1
            If .Item(i).value = vTemp Then
                Exit For
            End If
        Next i
        If i < .Count - 1 Then
            i = i + 1
        Else
            i = 0
        End If
    End With
    fnNextValueItemIndex = i

End Function

Property Let FormatedCell(ByVal nCol As Integer, _
                          ByVal lRow As Long, _
                          vValue As Variant)
    
    If RowCount < 0 Then
        Exit Property
    End If
    If nCol >= nMaxCol Or lRow > nRowCount Then
        Exit Property
    End If
    GridArray(nCol, lRow) = vValue

End Property

Public Sub GetColumn(vData() As Variant, _
                     lCount As Long, _
                     ByVal nCol As Integer)

    lCount = 0
    If nCol < 0 Or nCol > nMaxCol Then
        Exit Sub
    End If
    If nRowCount <= 0 Then
        Exit Sub
    End If
    
    Dim i As Long
    
    ReDim vData(nRowCount - 1)
    For i = 0 To nRowCount - 1
        vData(i) = GridArray(nCol, i)
    Next
    lCount = nRowCount
End Sub

Property Let HighlightInvalidCell(bStatus As Boolean)
    bHighlightInvalidCell = bStatus
End Property

Property Get HighlightInvalidCell() As Boolean
    HighlightInvalidCell = bHighlightInvalidCell
End Property

Public Function IsEmptyRow(Optional vRow As Variant) As Boolean
    If IsMissing(vRow) Then
        IsEmptyRow = GridEmpty(GetCurrentRowNumber)
    Else
        IsEmptyRow = GridEmpty(vRow)
    End If
End Function

Property Get KeepData() As Boolean
    KeepData = m_bKeepData
End Property

Property Let KeepData(bTemp As Boolean)
    m_bKeepData = bTemp
End Property

Property Let MaxRows(lTemp As Long)
    lMaxRows = lTemp
End Property

'david 04/27/2002
'THIS CALL IS REQUIRED FOR THE IMPLEMENTATION OF:
'WHEN CLICKED ON THE COLUMN HEADER TO SORT THE GRID
'TO SET/RESTORE HE HIGHLIGHTED ROW PROPERLY

'SHOULD BE CALLED ALSO IF THE SELECTOR (ROW HIGHLIGHTING) FEATURE IS USED
'''''''''''''''''
Public Sub MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    nKeyProcessing = KEY_P_FINISHED
    subSaveHighlightedRowID Button, Shift, x, y
End Sub

Public Property Get OldValue() As Variant
    OldValue = m_nOldValue
End Property

'david 10/20/2002
'when program change the cellvalue, the BeforeColEdit did not fire.
'and thus the 'm_nOldValue' does not set and is not correct!
'this method SHOULD BE called if 'OldValue' get method is called
'after the program changed the cellvalue.
Public Property Let OldValue(sCellValueBeforeChanged As Variant)
    m_nOldValue = sCellValueBeforeChanged
End Property

Public Property Get CellValueBeforeColEdit() As Variant
    CellValueBeforeColEdit = m_nOldValue
End Property


Public Function IsEndOfTable() As Boolean
    'calculate the different because the visible rows in table may be more than one over the max. rows
    On Error Resume Next
    If tgTable.Bookmark >= nRowCount - 1 Then
        IsEndOfTable = True
    Else
        IsEndOfTable = False
    End If
End Function

Public Function ChangedRowValid(ByVal lRow As Long, _
                                Optional ByVal bAllBlankIsValid As Variant, _
                                Optional ByVal bAnyRow As Variant, _
                                Optional vVisibleCol) As Boolean
'check changed row to see if the whole row's data is valid.  *** Take into consideration: If the all columns
'in the last row are empty, the changed row (ie. last row) is valid.
    Dim bkMark As Variant, vData() As Variant, lRowNumber As Long, i As Long
    Dim nVisibleCol As Integer, bRowIsBlank As Boolean
    
    bkMark = bmChangedRows(lRow)
    lRowNumber = fnIndexFromBookmark(bkMark, False)
    
    If lRowNumber < 0 Then
        ChangedRowValid = True
        Exit Function
    End If
    
    bRowIsBlank = True
    
    If IsMissing(vVisibleCol) Then
        nVisibleCol = nMaxCol - 1
    Else
        nVisibleCol = vVisibleCol - 1
    End If
    
    'check row is blank
    For i = 0 To nVisibleCol
        If Not IsNull(GridArray(i, lRowNumber)) Then
            If Trim(GridArray(i, lRowNumber)) <> "" Then
                bRowIsBlank = False
                Exit For
            End If
        End If
    Next i
    
    If bRowIsBlank Then
        If Not IsMissing(bAllBlankIsValid) Then
            If bAllBlankIsValid Then
                If IsMissing(bAnyRow) Then
                    If lRowNumber = nMaxRow - 1 Then ChangedRowValid = True
                Else
                    If bAnyRow Then ChangedRowValid = True
                End If
            End If
        End If
    Else
        ChangedRowValid = ValidData(lRowNumber)  ', True  david 08-18-98
    End If
End Function

Public Sub UnboundAddData(ByVal RowBuf As RowBuffer, NewRowBookmark As Variant)
    'no code, for compatibility only
End Sub

'UnboundDeleteRow event - David
Public Sub UnboundDeleteRow(Bookmark As Variant)
    If Not fnDeleteRow(Bookmark) Then Bookmark = Null
    tgTable.Rebind
End Sub

Public Function Prompt(Optional vIndex As Variant) As String
    Dim nIndex As Integer
    
    If IsMissing(vIndex) Then
        nIndex = tgTable.Col
    Else
        nIndex = vIndex
    End If
    If nIndex >= nColumnCount Then
        Exit Function
    End If
    Dim nCol As Integer
    nCol = fnColumnNumber(nIndex)
    If nCol < 0 Then
        Exit Function
    End If
    Prompt = sPrompt(nCol)
End Function

Public Sub ResetFlags(Optional vRow As Variant, _
                      Optional vCol As Variant, _
                      Optional vStatus As Variant)
    Dim j As Long
    Dim i As Integer
    Dim nStatus As Integer
    
    If IsMissing(vStatus) Then
        nStatus = STATUS_NOCHANGE
    Else
        nStatus = vStatus
    End If
    If IsMissing(vRow) Then
        j = GetCurrentRowNumber
    Else
        j = Val(vRow)
    End If
    If IsMissing(vCol) Then
        For i = 0 To nColumnCount - 1
            nCellChanged(i, j) = nStatus
        Next
    Else
        i = fnColumnNumber(vCol)
        If i >= 0 Then
            nCellChanged(i, j) = nStatus
        End If
    End If
End Sub

Property Get RowHighLighted() As Boolean
    RowHighLighted = (m_nRowHighLighted = STYLE_HIGHLIGHT)
End Property

Public Sub SetGoUpDown()
    nReturnKeyAction = RETURN_KEY_DOWN
End Sub

Public Sub SetGoLeftRight()
    nReturnKeyAction = RETURN_KEY_RIGHT
End Sub

Public Sub InsertRow(vData() As Variant, ByVal lRow As Long)

    If fnUbound(vData) < nMaxCol - 1 Then
        Exit Sub
    End If
    
    Dim i As Long
    Dim j As Integer
    Dim bkMark As Variant

    nRowCount = nRowCount + 1
    bkMark = fnGetNewBookmark
    For i = nRowCount - 1 To lRow Step -1
        For j = 0 To nMaxCol + 1
            GridArray(j, i + 1) = GridArray(j, i)
        Next j
    Next i
    
    GridArray(nFlagColumn, lRow) = ROW_INITIAL
    GridArray(nBookMarkColumn, lRow) = bkMark
    
    For j = 0 To nMaxCol - 1
        GridArray(j, lRow) = vData(j)
    Next

End Sub

'fnNewInsertRow functin - David
Public Function InsertNewRow(bookm As Variant, ByVal bInsertAfter As Boolean) As Boolean

    ReDim vData(nMaxCol + 1) As Variant
    Dim lRow As Long
    Dim bmRow As Variant
    
    lRow = RowByBookMark(bookm)
    If bInsertAfter Then
        lRow = lRow + 1
    End If
    If lRow >= RowCount Then
        InsertRow vData, RowCount
    Else
        InsertRow vData, lRow
    End If
    
    tgTable.Rebind
    InsertNewRow = True
Exit Function

    Dim Index As Long, lMaxRow As Long, lSkipRow, i As Long, j As Integer
    
    If bookm = 0 And nMaxRow = 0 Then 'insert the very first row
        GridArray(nBookMarkColumn, 0) = fnGetNewBookmark
        fnSaveBookMarks GridArray(nBookMarkColumn, 0)
        
        GridArray(nFlagColumn, 0) = ROW_ADDED
        bClearData = True
        
        For j = 0 To nMaxCol - 1
            GridArray(j, 0) = ""
        Next j
        
        tgTable.Rebind
        bClearData = False
        InsertNewRow = True
        Exit Function
    End If
    
    Index = fnIndexFromBookmark(bookm, False)
    If Index < 0 Or Index >= nMaxRow Then
        InsertNewRow = False
        Exit Function
    End If
    
    nMaxRow = nMaxRow + 1
    nRowCount = nRowCount + 1
    
    Dim nTemp As Integer
    
    ReDim Preserve GridArray(nMaxCol + 1, nMaxRow)
    
    nTemp = UBound(nCellChanged)
    ReDim Preserve nCellChanged(nTemp, nMaxRow + 1)
    nTemp = UBound(bValidCell)
    ReDim Preserve bValidCell(nTemp, nMaxRow + 1)

    If bInsertAfter Then
        lSkipRow = 1
    Else
        lSkipRow = 0
    End If
    
    'If New row is not the last row then shift all rows downward from the new row
    If Index + lSkipRow < nMaxRow - 1 Then
        'Shift downword the data in the array
        For i = nMaxRow - 1 To Index + lSkipRow Step -1
            For j = 0 To nMaxCol - 1
                GridArray(j, i) = GridArray(j, i - 1)
            Next j
            For j = 0 To nColumnCount - 1
                nCellChanged(j, i) = nCellChanged(j, i - 1)
                bValidCell(j, i) = bValidCell(j, i - 1)
            Next j
        Next i
        i = Index + lSkipRow
    ElseIf Index + lSkipRow = nMaxRow - 1 Then  'insert at the end of the table
        i = nMaxRow - 1
    End If
        
    'blank out all columns for the new row
    For j = 0 To nMaxCol - 1
        GridArray(j, i) = ""
    Next j
    
    GridArray(nFlagColumn, i) = ROW_INITIAL
    GridArray(nBookMarkColumn, i) = MakeBookmark(i)
    
    're-new the bookmark in bmChangedRows array is needed!
    'because all bookmarks in the GridArray will be re-sequence after rebind
    
    For j = 0 To nChangedCount - 1
        For i = 0 To nMaxRow - 1
            If bmChangedRows(j) = GridArray(nBookMarkColumn, i) Then
                bmChangedRows(j) = MakeBookmark(i)
                Exit For
            End If
        Next i
    Next j
    
    'Now re-new Bookmark Column
    For i = 0 To nMaxRow - 1
        GridArray(nBookMarkColumn, i) = MakeBookmark(i)
    Next i
    
    InsertNewRow = True
    
    On Error Resume Next
    tgTable.Rebind
'    tgTable.SetFocus

End Function

Property Let ColumnCount(nCol As Integer)
    If nCol > 0 Then
        nMaxCol = nCol
    
        nFlagColumn = nMaxCol
        nBookMarkColumn = nMaxCol + 1
        
        ReDim sPatternCode(nMaxCol - 1)
        ReDim sPrompt(nMaxCol - 1)
        ReDim nEditColumns(nMaxCol - 1)
        ReDim nColumnType(nMaxCol - 1)
        ReDim sDefaultValue(nMaxCol - 1)
        ReDim bFieldRequired(nMaxCol - 1)
        ReDim bCheckRegExp(nMaxCol - 1)
        ReDim sDisplayFormat(nMaxCol - 1)
        ReDim sErrMessage(nMaxCol - 1)
        
        ReDim GridArray(nMaxCol + 1, 0)
            
        Dim i As Integer
        
        For i = 0 To nMaxCol - 1
            sDisplayFormat(i) = ""
        Next
    End If
End Property

Property Get GetEditColumnCount() As Integer
    GetEditColumnCount = nColumnCount
End Property

Property Get LastRow() As Long
    If IsNull(tgTable.FirstRow) Then
        LastRow = nLastTableRow
    Else
        LastRow = Val(tgTable.FirstRow) + nLastTableRow
    End If
End Property

Public Function GridEmpty(Optional ByVal vRow As Variant) As Boolean
    Dim i As Long
    Dim j As Integer
    Dim lStart As Long
    Dim lEnd As Long
    
    If IsMissing(vRow) Then
        lStart = 0
        lEnd = nRowCount - 1
    Else
        If vRow < 0 Then
            Exit Function
        End If
        lStart = vRow
        If lStart > fnUbound(GridArray, 2) Then
            Exit Function
        End If
        lEnd = vRow
    End If
    
    For i = lStart To lEnd
        For j = 0 To nMaxCol - 1
            If Not IsNull(GridArray(j, i)) Then
                If Trim(GridArray(j, i)) <> "" Then
                    GridEmpty = False
                    Exit Function
                End If
            End If
        Next j
    Next i
    
    GridEmpty = True
End Function

Property Let Locked(bFlag As Boolean)
    bLocked = bFlag
End Property

Property Get Locked() As Boolean
    Locked = bLocked
End Property

Public Sub NewRow()

End Sub

Public Sub Refresh()
    tgTable.Refresh
End Sub

Property Get FirstTimeGotFocus()
    FirstTimeGotFocus = bFirstTimeGotFocus
End Property

Property Let FirstTimeGotFocus(bOnOff)
    bFirstTimeGotFocus = bOnOff
End Property

Public Function Revalidate(ByVal nCol As Integer, ByVal lRow As Long) As Boolean
    Dim szText As String
    Dim vTemp As Variant
    Dim nECol As Integer
    
    nECol = fnColumnNumber(nCol)
    If nECol >= 0 And nECol < nColumnCount Then
        nCellChanged(nECol, lRow) = nCellChanged(nECol, lRow) And (Not STATUS_VALIDATED)
        Revalidate = fnValidate(nCol, lRow, fnGetUserData(MakeBookmark(lRow), nCol))
    End If
End Function

Public Function RowByBookMark(bkMark As Variant) As Long
    'Returns the row number for the given bookmark.

    Dim nMid As Long
    Dim nEnd As Long
    
    RowByBookMark = -1
    nEnd = nRowCount
    If nEnd <= fnUbound(GridArray, 2) Then
        If IsNull(GridArray(nBookMarkColumn, nEnd)) Then
            nEnd = nRowCount - 1
        End If
    End If
    For nMid = 0 To nEnd
        If bkMark = GridArray(nBookMarkColumn, nMid) Then
            RowByBookMark = nMid
            Exit For
        End If
    Next
End Function

Property Get RowIsChanged(ByVal lRow As Long) As Boolean
    Dim i As Integer
    
    RowIsChanged = False
    Dim bkMark As Variant
    bkMark = Bookmark(lRow)
    For i = 0 To nChangedCount - 1
        If bkMark = bmChangedRows(i) Then
            RowIsChanged = True
            Exit For
        End If
    Next
End Property

Property Let RowIsChanged(ByVal lRow As Long, bStatus As Boolean)
    If bStatus Then
        fnSaveBookMarks Bookmark(lRow)
    Else
        If nChangedCount = 0 Then
            Exit Property
        End If
        
        Dim i As Integer
        Dim bkMark As Variant
        
        bkMark = Bookmark(lRow)
        
        For i = 0 To nChangedCount - 1
            If bkMark = bmChangedRows(i) Then
                Exit For
            End If
        Next
        
        If i < nChangedCount Then
            While i < nChangedCount
                bmChangedRows(i) = bmChangedRows(i + 1)
                i = i + 1
            Wend
            nChangedCount = nChangedCount - 1
        End If
    End If
End Property

Public Sub RefreshFlags()
    Dim i As Long
    Dim j As Integer
    
    For i = 0 To nRowCount
        For j = 0 To nColumnCount - 1
            nCellChanged(j, i) = STATUS_NOCHANGE
            bValidCell(j, i) = True
        Next j
    Next i
    
    For i = 0 To nChangedCount - 1
        bmChangedRows(i) = Null
    Next
    
    nChangedCount = 0
End Sub

Public Sub RemoveEditColumn(ByVal nIndex As Integer)
    If nIndex >= nMaxCol Or nColumnCount > nMaxCol Or nColumnCount = 0 Then
        Exit Sub
    End If
    Dim i As Integer
    i = 0
    While nEditColumns(i) <> nIndex And i < nColumnCount
        i = i + 1
    Wend
    If i < nColumnCount Then
        If nIndex < tgTable.Columns.Count Then
            tgTable.Columns(nIndex).Locked = True
            nColumnCount = nColumnCount - 1
            While i < nColumnCount
                nEditColumns(i) = nEditColumns(i + 1)
                sPrompt(i) = sPrompt(i + 1)
                sPatternCode(i) = sPatternCode(i + 1)
                nColumnType(i) = nColumnType(i + 1)
                sDefaultValue(i) = sDefaultValue(i + 1)
                bFieldRequired(i) = bFieldRequired(i + 1)
                bCheckRegExp(i) = bCheckRegExp(i + 1)
                i = i + 1
            Wend
        End If
    End If
    
    If Not m_bKeepData Then
        fnInitialParams
    End If
End Sub

Property Get EditedRow() As Integer
    EditedRow = ROW_EDITED
End Property

Private Sub fnFetchDefault(nCol As Integer, ByVal lCurrRow As Long)
    If nCellChanged(nCol, lCurrRow) = STATUS_NOCHANGE And fnCellValue(nLastColumn, lCurrRow) = "" Then
        If nCol = 0 Then
            If IsNull(Bookmark(lCurrRow)) Then
                GridArray(nBookMarkColumn, lCurrRow) = fnGetNewBookmark()
                nRowCount = nRowCount + 1
            End If
        End If
        If Not IsNull(Bookmark(lCurrRow)) Then
            If sDefaultValue(nCol) <> "" Then
                tgTable.Columns(nLastColumn).Text = sDefaultValue(nCol)
                If nKeyProcessing = KEY_P_FINISHED Then
                    DoEvents
                End If
                nCellChanged(nCol, lCurrRow) = nCellChanged(nCol, lCurrRow) Or STATUS_CHANGED
                GridArray(nLastColumn, lCurrRow) = sDefaultValue(nCol)
            Else
                If sPatternCode(nCol) = "" Then
                    bValidCell(nCol, lCurrRow) = True
                End If
            End If
        End If
    End If
End Sub

Private Sub fnInitialParams()
    Dim nTemp As Integer
    
    nRowCount = 0
    nChangedCount = 0
    nLastColumn = tgTable.Col
    nLastTableRow = tgTable.Row
    bCellEditing = False
    lBookMarkCount = 0
    lEditingRow = -1
    If nKeyColumnCount > 0 Then
        lMaxKeys = 0
        lKeyRows = 0
        ReDim sKeyFields(nKeyColumnCount, lMaxKeys)
    End If
    ReDim GridArray(nMaxCol + 1, nMaxRow + 1)
    If m_bKeepData Then
        nTemp = UBound(bValidCell)
        ReDim bValidCell(nTemp, nMaxRow + 1)
        nTemp = UBound(nCellChanged)
        ReDim nCellChanged(nTemp, nMaxRow + 1)
    Else
        If nColumnCount > 0 Then
            ReDim bValidCell(nColumnCount - 1, nMaxRow + 1)
            ReDim nCellChanged(nColumnCount - 1, nMaxRow + 1)
        End If
    End If
End Sub

Private Sub fnSaveChangedKeys()
    Dim nCol As Integer
    Dim bmTemp As Variant

    If lEditingRow < 0 Then
        Exit Sub
    End If
    bmTemp = Bookmark(lEditingRow)
    If IsNull(bmTemp) Then
        Exit Sub
    End If
    
    If nKeyColumnCount <= 0 Then
        Exit Sub
    End If
    If sKeyFields(nKeyColumnCount, lKeyRows) = "  " Then
        sKeyFields(nKeyColumnCount, lKeyRows) = bmTemp
        If Val(GridArray(nFlagColumn, lEditingRow)) = ROW_ADDED Then
            For nCol = 0 To nKeyColumnCount - 1
                sKeyFields(nCol, lKeyRows) = GridArray(nKeyColumns(nCol), lEditingRow)
            Next
        End If
        lKeyRows = lKeyRows + 1
        If lKeyRows >= lMaxKeys Then
            lMaxKeys = lMaxKeys + ROW_INCREMENT
            ReDim Preserve sKeyFields(nKeyColumnCount, lMaxKeys)
        End If
    End If
End Sub

Private Function fnValidate(nCurrCol As Integer, _
                            lCurrRow As Long, _
                            vText As Variant) As Boolean

    Dim szText As String

    If bValidating And Not m_bAllowRevalidate Then
        Exit Function
    End If
    bValidating = True
                                                                
    Dim nCol As Integer
    Dim nCode As Integer
    
    If IsNull(vText) Then
        szText = ""
    Else
        szText = vText
    End If
    
    nCol = fnColumnNumber(nCurrCol)
    
    If nCol < 0 Or lCurrRow > nRowCount Then
        bValidating = False
        Exit Function
    End If
    If ((nCellChanged(nCol, lCurrRow) And STATUS_VALIDATED) = STATUS_VALIDATED) Then
        bValidating = False
        Exit Function
    End If
    
    If bFieldRequired(nCol) Then
        If szText = "" Then
            #If FACTOR_MENU = 1 Then
                bValidCell(nCol, lCurrRow) = True
            #Else
                bValidCell(nCol, lCurrRow) = frmParent.fnValidCellValue(tgTable, nCurrCol, lCurrRow, szText)
            #End If
            nCellChanged(nCol, lCurrRow) = nCellChanged(nCol, lCurrRow) Or STATUS_VALIDATED
            fnValidate = Not bValidCell(nCol, lCurrRow)
            bValidating = False
            Exit Function
        End If
    Else
        bValidCell(nCol, lCurrRow) = True
        nCellChanged(nCol, lCurrRow) = nCellChanged(nCol, lCurrRow) Or STATUS_VALIDATED
        bValidating = False
        Exit Function
    End If
    
    nCode = 0
    If sPatternCode(nCol) <> "" Then
        #If REGULAR_EXP_INCLUDED Then
            If bCheckRegExp(nCol) Then
                nCode = PRegExpMatch(sPatternCode(nCol), Len(sPatternCode(nCol)), szText)
            End If
        #End If
    End If
    If nCode = 0 Or nCode = Len(szText) + 1 Then
        #If FACTOR_MENU = 1 Then
            bValidCell(nCol, lCurrRow) = True
        #Else
            bValidCell(nCol, lCurrRow) = frmParent.fnValidCellValue(tgTable, nCurrCol, lCurrRow, szText)
        #End If
    Else
        bValidCell(nCol, lCurrRow) = False
    End If
    fnValidate = Not bValidCell(nCol, lCurrRow)
    nCellChanged(nCol, lCurrRow) = nCellChanged(nCol, lCurrRow) Or STATUS_VALIDATED
    bValidating = False
End Function

Public Sub GetChangedKey(vData() As Variant, nCount As Integer, ByVal lIndex As Long)
    Dim i As Integer
    Dim lRow As Long
    If nKeyColumnCount <= 0 Then
        Exit Sub
    End If
    ReDim vData(nKeyColumnCount - 1)
    
    If lIndex < 0 Or lIndex >= nChangedCount Then
        nCount = 0
        Exit Sub
    End If
    For lRow = 0 To lKeyRows - 1
        If sKeyFields(nKeyColumnCount, lRow) = bmChangedRows(lIndex) Then
            Exit For
        End If
    Next
    If lRow >= lKeyRows Then
        MsgBox "Key is not recorded"
        Exit Sub
    End If
    If lRow >= 0 And lRow < nRowCount Then
        For i = 0 To nKeyColumnCount - 1
            vData(i) = sKeyFields(i, lRow)
        Next i
    End If
    nCount = nKeyColumnCount
End Sub

Public Sub GetRecords(vData() As Variant, _
                      lRows As Long, _
                      sStatement As String)
    If nRowCount <= 0 Then
        Exit Sub
    End If

    Dim i As Long
    Dim j As Integer
    
    If sStatement = "" Then
        lRows = nRowCount
        If lRows >= 1 Then
            ReDim vData(nMaxCol - 1, lRows - 1)
            For i = 0 To lRows - 1
                For j = 0 To nMaxCol - 1
                    vData(j, i) = GridArray(j, i)
                Next j
            Next i
        End If
        Exit Sub
    End If

    Dim sParams(20) As String
    Dim nParams(20) As Integer
    Dim sParamValue(20) As Variant
    Dim nParamCount As Integer
    Dim nPos As Integer
    Dim nLen As Integer
    Dim lRowNos() As Long
    
    nPos = InStr(LCase(sStatement), "where")

    nLen = Len(sStatement)
    If nPos > 1 Then
        fnParseString sParams, sStatement, , 0, nPos - 1
    Else
        fnParseString sParams, sStatement, , 0, nLen
    End If
    fnParseAssigments nParams(), sParamValue(), nParamCount, sParams()
    If nParamCount = 0 Then
        Exit Sub
    End If
    
    If nPos > 1 Then
        subGetAffectedRows lRowNos, lRows, Right(sStatement, nLen - (nPos + 5))
    Else
        subGetAffectedRows lRowNos, lRows, ""
    End If
    If lRows > 0 Then
        ReDim vData(nParamCount, lRows - 1)
        
        For i = 0 To lRows - 1
            For j = 0 To nParamCount - 1
                vData(j, i) = GridArray(nParams(j), lRowNos(i))
            Next j
            vData(nParamCount, i) = lRowNos(i)
        Next i
    End If
End Sub

Public Sub GetRow(vData() As Variant, nCount As Integer, ByVal lIndex As Long)
    Dim i As Integer
    ReDim vData(nMaxCol)
    
    nCount = 0
    If lIndex >= 0 And lIndex < RowCount Then
        For i = 0 To nMaxCol
            vData(i) = GridArray(i, lIndex)
        Next i
        nCount = nMaxCol
    End If
End Sub

Public Sub GotFocus()
    If Not bLocked Then
        If fnColumnNumber(tgTable.Col) < 0 Then
            If nColumnCount > 0 Then
                If FirstEditableCol < tgTable.Columns.Count Then
                    tgTable.Col = FirstEditableCol
                End If
            End If
        End If
        If nKeyProcessing = KEY_P_FINISHED Then
            DoEvents
        End If
        fnShowStatus GetCurrentRowNumber, tgTable.Col
        If nRowCount = 0 And IsNull(Bookmark(0)) Then
            If nColumnCount > 0 Then
                fnFetchDefault 0, 0
            End If
        End If
    End If
End Sub

Public Sub GotoNextColumn(ByVal nLastCol As Integer)
    If bClearData Then
        Exit Sub
    End If

    Dim nCol As Integer
    Dim bFlag As Boolean
    
    bFlag = bClearData
    bClearData = True
    nCol = fnColumnNumber(tgTable.Col)
    If nCol < 0 Or nCol + 1 >= nColumnCount Then
        If nCol < 0 Then
            If fnGotoPreviousRow Then
                nCol = nColumnCount - 1
            End If
        Else
            If fnGotoNextRow Then
                nCol = 0
            End If
        End If
        If nCol >= 0 Then
            If nEditColumns(nCol) < tgTable.Columns.Count Then
                tgTable.Col = nEditColumns(nCol)
            End If
        End If
        bClearData = bFlag
        Exit Sub
    End If
    If tgTable.Col < nLastCol And nLastCol = nEditColumns(nCol + 1) Then
        If nCol = 0 Then
            If fnGotoPreviousRow Then
                nCol = nColumnCount - 1
            End If
        Else
            nCol = nCol - 1
        End If
    Else
        If nCol = nColumnCount - 1 Then
            If fnGotoNextRow Then
                nCol = 0
            End If
        Else
            nCol = nCol + 1
        End If
    End If
    If nEditColumns(nCol) < tgTable.Columns.Count Then
        tgTable.Col = nEditColumns(nCol)
    End If
    bClearData = bFlag
    If nKeyProcessing = KEY_P_FINISHED Then
        DoEvents
    End If
End Sub

Property Get InitialRow() As Integer
    InitialRow = ROW_INITIAL
End Property

Public Function IsNewRow() As Boolean
    If GetCurrentRowNumber >= 0 Then
        If IsNull(Bookmark(GetCurrentRowNumber)) Then
            IsNewRow = True
        Else
            IsNewRow = False
        End If
    End If
End Function

Public Sub Rebind()
    If tgTable.Row > 0 Then
        On Error Resume Next
        bClearData = True
        tgTable.Row = tgTable.Row - 1
        tgTable.Rebind
        If bAllowAddNew Then
            If GetCurrentRowNumber < RowCount - 1 Then
                tgTable.Row = tgTable.Row + 1
            End If
        Else
            tgTable.Row = tgTable.Row + 1
        End If
        bClearData = False
        On Error GoTo 0
    Else
        tgTable.Rebind
    End If
End Sub

Public Sub ResetSetup()
    nMaxCol = 0
    nRowCount = 0
    nMaxRow = 0

    ReDim GridArray(0, 0)
    ReDim nCellChanged(0)
    ReDim bValidCell(0)
    
    ReDim sKeyFields(0)
    ReDim nKeyColumns(0)
    ReDim bmChangedRows(0)
    ReDim sPatternCode(0)
    ReDim sPrompt(0)
    ReDim nEditColumns(0)
    ReDim nColumnType(0)
    ReDim sDefaultValue(0)
    
    nColumnCount = 0
    nFlagColumn = 0
    nBookMarkColumn = 0
    nChangedCount = 0
    nMaxChanged = 0
    lEditingRow = -1
    nLastColumn = 0
    nLastTableRow = 0
'    nCurrCol = 0
    lBookMarkCount = 0
    nKeyColumnCount = 0
    lKeyRows = 0
    lMaxKeys = 0

    bErrorFlag = False
    bClearData = False
    bAllowAddNew = False
    bRowChanging = False
    bValidating = False
    nKeyProcessing = KEY_P_FINISHED
    bCellEditing = False
    
    Set tgTable = Nothing
End Sub

Public Sub Resize()
    Dim fFactor As Single
    Dim nTW As Integer
    Dim i As Integer
    fFactor = frmParent.Height / frmParent.OriginalHeight
    tgTable.Height = nOriginalTableHeight * fFactor
    tgTable.Font.Size = nFontSize * fFactor
    tgTable.HeadFont.Size = nFontSize * fFactor
    
    fFactor = frmParent.Width / frmParent.OriginalWidth
    nTW = tgTable.Width
    tgTable.Width = nOriginalTableWidth * fFactor
    fFactor = tgTable.Width / nTW
    For i = 0 To tgTable.Columns.Count - 1
        tgTable.Columns(i).Width = fFactor * tgTable.Columns(i).Width
    Next i
End Sub

'change return type from integer to long type by junsong Oct. 9, 2000
Property Get RowCount() As Long
    Dim i As Integer
    RowCount = nRowCount
    If bAllowAddNew Then
        On Error GoTo errRowCount
        If nRowCount <= fnUbound(GridArray, 2) Then
            For i = 0 To nMaxCol - 1
                    If Not IsNull(GridArray(i, nRowCount)) Then
                        If GridArray(i, nRowCount) <> "" Then
                            RowCount = nRowCount + 1
                        End If
                    End If
            Next
        End If
    End If
    Exit Property
    
errRowCount:
    RowCount = 0
End Property

Property Let CellValue(ByVal nCol As Integer, ByVal lRow As Long, vValue As Variant)
    Dim nCol1 As Integer

    If nRowCount < 0 Then
        Exit Property
    End If
    
    If nCol >= nMaxCol Or lRow < 0 Or lRow > nRowCount Then
        Exit Property
    End If
    
    'david 03/23/2006  #513569
    On Error GoTo errTrap
    ''''''''''''''''''''''''''
    
    GridArray(nCol, lRow) = vValue
    nCol1 = fnColumnNumber(nCol)
    
    If nCol1 >= 0 Then
        nCellChanged(nCol1, lRow) = nCellChanged(nCol1, lRow) And (Not STATUS_VALIDATED)
    End If

    If tgTable.Columns(nCol).Visible Then
        'david 06/05/2002
        'refresh the text of a cell automatically
        'when the content is changed in CellValue() let property.
        If m_bCellValueAutoRebind Then
            Dim lCurrRow As Long
            Dim bTempLocked As Boolean
            
            lCurrRow = tgTable.Row
            
            If lCurrRow < 0 Then
                Exit Property
            End If
            
            If tgTable.FirstRow = "" Then
                Exit Property
            End If
            
            bTempLocked = bLocked
            bLocked = True
            
            If lRow >= tgTable.FirstRow And lRow < tgTable.FirstRow + tgTable.VisibleRows Then
                If lRow = tgTable.FirstRow + tgTable.VisibleRows - 1 And fnIsLastEditableCol() Then
                    tgTable.Row = lRow - tgTable.FirstRow
                ElseIf lRow < tgTable.FirstRow + tgTable.VisibleRows - 1 Then
                    tgTable.Row = lRow - tgTable.FirstRow
                ElseIf lRow <= tgTable.VisibleRows Then
                    tgTable.Row = lRow
                End If
                
                If sDisplayFormat(nCol) = "" Then
                    If tgTable.Columns(nCol).ValueItems.Count = 0 Then
                        tgTable.Columns(nCol).Text = vValue
                    End If
                Else
                    tgTable.Columns(nCol).Text = Format$(vValue, sDisplayFormat(nCol))
                End If
                
                tgTable.Row = lCurrRow
            End If
            
            bLocked = bTempLocked
        '''''''''''''''''
        Else
            If lRow = GetCurrentRowNumber() Then
                Dim bkMark As Variant
                bkMark = MakeBookmark(lRow)
                tgTable.Columns(nCol).Text = fnGetUserData(bkMark, nCol)
            End If
        End If
    End If
    
    Exit Property
    
errTrap:
    'david 03/23/2006  #513569
    'Occationally the assignment statement "tgTable.Row = "
    'will cause RTE 6148, Invalid row number. I don't know the exact causes.
    'One reason could be the sum of the height of the rows in the grid
    'is equal to the height of the grid
    'Before, when this RTE occurs, the program will simply quit
    'I will show the error and RESUME NEXT
    'set IgnoreCellValueError = True to suppress the error message
    If Not m_bIgnoreCellValueError Then
        Dim sErrMsg As String
        sErrMsg = "An error has occurred in CellValue()." & vbCrLf & vbCrLf
        sErrMsg = sErrMsg & "Err Code: " & Err.Number & vbCrLf & "Err Desc: " & Err.Description
        MsgBox sErrMsg, vbExclamation
    End If
    
    Resume Next
    ''''''''''''''''''''''''''
End Property

'david 03/23/2006  #513569
Property Let IgnoreCellValueError(bFlag As Boolean)
    m_bIgnoreCellValueError = bFlag
End Property

Property Get IgnoreCellValueError() As Boolean
    IgnoreCellValueError = m_bIgnoreCellValueError
End Property
''''''''''''''''''''''''''

Public Sub SetCurrentRowData(vData() As Variant)
    Dim i As Integer
    Dim nCol As Integer
    Dim lRow As Long
    Dim bkMark As Variant
    lRow = GetCurrentRowNumber
    bkMark = MakeBookmark(lRow)
    
    If lRow < 0 Then
        Exit Sub
    End If
    For i = 0 To fnUbound(vData)
        If Not IsNull(vData(i)) Then
            GridArray(i, lRow) = vData(i)
            tgTable.Columns(i).Text = fnGetUserData(bkMark, i)
            nCol = fnColumnNumber(i)
            If nCol >= 0 Then
                nCellChanged(nCol, lRow) = (nCellChanged(nCol, lRow)) And (Not STATUS_VALIDATED)
            End If
        End If
    Next
    GridArray(nFlagColumn, lRow) = ROW_ADDED
End Sub

Public Sub SetRowHighLighted(bHighLighted As Boolean, _
                             Optional bForceHighLighted As Boolean = False, _
                             Optional bSelectedRow As Boolean = False)
    
    If bHighLighted Then
        If m_nRowHighLighted = STYLE_HIGHLIGHT And (Not bForceHighLighted) Then
            Exit Sub
        End If
    Else
        If m_nRowHighLighted = STYLE_PLAIN Then
            Exit Sub
        End If
    End If
    
    Dim m_styRow As DBTRUEGRID.Style            ' row style
    Dim m_styCurrentCell As DBTRUEGRID.Style    ' current cell style
    Dim m_styHighLightedRow As DBTRUEGRID.Style

    tgTable.ClearCellStyle dbgAllCells
    
    ' set some properties for the table
    Set m_styRow = tgTable.Styles.Item(szROW_STYLE)
    Set m_styCurrentCell = tgTable.Styles.Item(szCURRENT_CELL_STYLE)
    Set m_styHighLightedRow = tgTable.Styles.Item(szROW_STYLE)
    
    If bHighLighted Then
        ' create highlight row style
        With m_styRow
            .BackColor = &HFF0000
            .ForeColor = &HFFFFFF
        End With

        ' create current cell style
        If nColumnCount > 0 Then
            With m_styCurrentCell
                .BackColor = &HFFFFFF
                .ForeColor = &H0
            End With
        Else
            Set m_styCurrentCell = m_styRow
        End If
    
        m_nRowHighLighted = STYLE_HIGHLIGHT
        
        'david 03/21/2001 WORK AROUND!!!
        'fix the bug of the last row cellvalue being reset to zero
        If bForceHighLighted Then
            bIn_SetRowHighLighted = True
        End If
            
        If nColumnCount > 0 Then
            If bSelectedRow Then
                If tgTable.SelBookmarks.Count > 0 Then
                    With m_styHighLightedRow
                        .BackColor = tgTable.SelectedBackColor
                        .ForeColor = &HFFFFFF
                    End With
                    
                    tgTable.AddCellStyle dbgSelectedRow, m_styHighLightedRow
                    tgTable.MarqueeStyle = dbgNoMarquee
                Else
                    tgTable.MarqueeStyle = dbgHighlightRowRaiseCell
                End If
            Else
                tgTable.MarqueeStyle = dbgHighlightRowRaiseCell
            End If
        Else
            tgTable.MarqueeStyle = dbgHighlightRow
        End If
    Else
        ' create plain row style
        With m_styRow
            .BackColor = &HFFFFFF
            .ForeColor = &H0
        End With
        
        m_nRowHighLighted = STYLE_PLAIN
        tgTable.MarqueeStyle = dbgNoMarquee
    End If

    ' add the new styles to the table
    tgTable.AddCellStyle dbgMarqueeRow, m_styRow
    tgTable.AddCellStyle dbgMarqueeRow + dbgCurrentCell, m_styCurrentCell
    tgTable.AddCellStyle dbgMarqueeRow + dbgUpdatedCell + dbgCurrentCell, m_styCurrentCell
    tgTable.AddCellStyle dbgMarqueeRow + dbgUpdatedCell, m_styRow
    tgTable.AddCellStyle dbgUpdatedCell, m_styCurrentCell

    bIn_SetRowHighLighted = False
End Sub

Public Sub SetupTable(ByVal bHighLighted As Boolean)

    Dim szBuf As String
    Dim n As Integer

    bAllowAddNew = True

'    Dim m_styRow As Object            ' row style
'    Dim m_styCurrentCell As Object    ' current cell style

'    On Error GoTo AddToCollectionError
    ' set column default properties
    For n = 0 To tgTable.Columns.Count - 1
        With tgTable.Columns(n)
            .DividerStyle = dbgBlackLine
            .AllowSizing = False
            .WrapText = False
            .Locked = True
            If .ValueItems.Count > 0 Then
                .ValueItems.CycleOnClick = True
            End If
        End With
    Next
    
    ' set default properties for the table
    With tgTable
        .AllowAddNew = False
        .DataMode = dbgUnbound
        .MarqueeUnique = True
        .EditDropDown = False
        .ExposeCellMode = dbgScrollOnEdit
        .TabAction = dbgControlNavigation
        .AllowRowSizing = False
        .ExtendRightColumn = True
        .ScrollBars = dbgAutomatic
        If nColumnCount > 0 Then
            .MarqueeStyle = dbgHighlightRowRaiseCell
        Else
            .MarqueeStyle = dbgHighlightRow
        End If
        .RecordSelectors = False
    End With
    
    SetRowHighLighted bHighLighted
    lBackFirstRow = 0
    
    On Error GoTo 0

End Sub

Public Sub subAddNewRow(nCol As Integer)
    Dim vRow As Variant
    Dim i As Integer
    Dim lRow As Long
    Dim szText As String
    If nCol < nColumnForNewRow Then
        Exit Sub
    End If
    If lMaxRows > 0 Then
        If UBound(GridArray, 2) >= lMaxRows - 1 Then
            Exit Sub
        End If
    End If
    If bAllowAddNew Then
        lRow = GetCurrentRowNumber
        If lRow >= nRowCount - 1 And lRow > 0 Then
            For i = nColumnCount - 1 To 0 Step -1
                fnValidate nEditColumns(i), lRow - 1, fnCellValue(nEditColumns(i), lRow - 1)
                If Not bValidCell(i, lRow - 1) Then
                    If Not bLocked Then
                        fnShowStatus lRow - 1, nEditColumns(i), True
                    End If
                    Exit Sub
                End If
            Next
        End If
        If lRow = nRowCount Then
            bAddingNewRow = True
            nRowCount = nRowCount + 1
            GridArray(nBookMarkColumn, nRowCount) = fnGetNewBookmark
            tgTable.Rebind
            If Not AddingValidate Then
                DoEvents
            End If
        End If
        
        If fnUbound(GridArray, 2) <= nRowCount Then
            If IsNull(GridArray(nBookMarkColumn, nRowCount)) Then
                GridArray(nBookMarkColumn, nRowCount) = fnGetNewBookmark
            Else
                If Trim(GridArray(nBookMarkColumn, nRowCount)) = "" Then
                    GridArray(nBookMarkColumn, nRowCount) = fnGetNewBookmark
                End If
            End If
        End If
    End If
End Sub

Private Sub subClearRow(ByVal lRow As Long)
    Dim i As Integer
    
    If lRow >= 0 Then
        For i = 0 To nColumnCount - 1
            Revalidate i, lRow
            nCellChanged(i, lRow) = STATUS_NOCHANGE
        Next
        For i = 0 To nMaxCol - 1
            If (Not tgTable.Columns(i).Locked) Or bCtrlYClearLocked Then
                GridArray(i, lRow) = ""
                If lRow = GetCurrentRowNumber Then
                    tgTable.Columns(i).Text = ""
                End If
            End If
        Next
    End If
End Sub

Private Sub subGetAffectedRows(lRows() As Long, lCount As Long, sConditions As String)
    
    Dim nLen As Integer
    Dim sConds(MAX_CONSTRAINTS - 1) As String
    Dim nFlag As Integer
    Dim i As Long
    Dim j As Integer
    Dim bFound As Boolean
    Dim arryConstraints(20) As tpConstraints
    Dim nCol As Integer
    Dim sOper As String
    Dim sValue As String
    Dim nCCount As Integer
    
    nLen = Len(sConditions)
    ReDim lRows(nRowCount - 1)
    If nLen = 0 Then
        'Get all
        lCount = nRowCount
        For i = 0 To lCount - 1
            lRows(i) = i
        Next
    Else
        fnParseString sConds, sConditions, , 0, nLen
        nFlag = 0
        nCCount = 0
        'Check for key words
        While nCCount < MAX_CONSTRAINTS And sConds(nCCount) <> ""
            If LCase(sConds(nCCount)) = KEYS_NOCURRENT Or LCase(sConds(nCCount)) = KEYS_NEW Then
                If LCase(sConds(nCCount)) = KEYS_NOCURRENT Then
                    nFlag = nFlag Or NOCURRENT
                ElseIf LCase(sConds(nCCount)) = KEYS_NEW Then
                    nFlag = nFlag Or NEWRECORD
                End If
                j = nCCount
                While sConds(j + 1) <> ""
                    sConds(j) = sConds(j + 1)
                    j = j + 1
                Wend
                sConds(j) = ""
                nCCount = nCCount - 1
            End If
            nCCount = nCCount + 1
        Wend
        
        'Parsing the constraints
        'ReDim arryConstraints(nCCount)
        For j = 0 To nCCount - 1
            subParseExpression nCol, sOper, sValue, sConds(j)
            With arryConstraints(j)
                .nColumn = nCol
                .sOperator = sOper
                .vValue = sValue
            End With
        Next
        
        lCount = 0
        For i = 0 To nRowCount - 1
            bFound = True
            If (nFlag And NOCURRENT) = NOCURRENT And i = GetCurrentRowNumber Then
                bFound = False
            End If
            If (nFlag And NEWRECORD) = NEWRECORD And GridArray(nFlagColumn, i) <> ROW_ADDED Then
                bFound = False
            End If
            If bFound Then
                For j = 0 To nCCount - 1
                    With arryConstraints(j)
                        If .nColumn < 0 Then
                            bFound = False
                        Else
                            If IsNull(GridArray(.nColumn, i)) Then
                                If .vValue <> "NULL" Then
                                    bFound = False
                                    Exit For
                                End If
                            Else
                                Select Case .sOperator
                                    Case OPR_EQUAL
                                        If fnComapreVariant(GridArray(.nColumn, i), .vValue) <> 0 Then
                                            bFound = False
                                            Exit For
                                        End If
                                    Case OPR_GREATER
                                        If fnComapreVariant(GridArray(.nColumn, i), .vValue) <= 0 Then
                                            bFound = False
                                            Exit For
                                        End If
                                    Case OPR_GREATER_EQUAL
                                        If fnComapreVariant(GridArray(.nColumn, i), .vValue) < 0 Then
                                            bFound = False
                                            Exit For
                                        End If
                                    Case OPR_LESS
                                        If fnComapreVariant(GridArray(.nColumn, i), .vValue) >= 0 Then
                                            bFound = False
                                            Exit For
                                        End If
                                    Case OPR_LESS_EQUAL
                                        If fnComapreVariant(GridArray(.nColumn, i), .vValue) > 0 Then
                                            bFound = False
                                            Exit For
                                        End If
                                End Select
                            End If
                        End If
                    End With
                Next j
            End If
            If bFound Then
                lRows(lCount) = i
                lCount = lCount + 1
            End If
        Next i
    End If

End Sub

Private Sub subParseExpression(nCol As Integer, _
                               sOper As String, _
                               sValue As String, _
                               sExpr As String)
    Const MAX_CONSTRAINTS = 20
    Const KEYS_NOCURRENT = "NOCURRENT"
    Const KEYS_NEW = "NEW"
    Const OPR_EQUAL = "="
    Const OPR_GREATER = ">"
    Const OPR_GREATER_EQUAL = ">="
    Const OPR_LESS = "<"
    Const OPR_LESS_EQUAL = "<="
    
    Dim sParms(2) As String
    fnParseString sParms, sExpr, OPR_LESS_EQUAL
    If sParms(1) <> "" Then
        nCol = fnIndexFromField(sParms(0))
        sOper = OPR_LESS_EQUAL
        sValue = sParms(1)
        Exit Sub
    End If
    fnParseString sParms, sExpr, OPR_GREATER_EQUAL
    If sParms(1) <> "" Then
        nCol = fnIndexFromField(sParms(0))
        sOper = OPR_GREATER_EQUAL
        sValue = sParms(1)
        Exit Sub
    End If
    fnParseString sParms, sExpr, OPR_EQUAL
    If sParms(1) <> "" Then
        nCol = fnIndexFromField(sParms(0))
        sOper = OPR_EQUAL
        sValue = sParms(1)
        Exit Sub
    End If
    fnParseString sParms, sExpr, OPR_GREATER
    If sParms(1) <> "" Then
        nCol = fnIndexFromField(sParms(0))
        sOper = OPR_GREATER
        sValue = sParms(1)
        Exit Sub
    End If
    fnParseString sParms, sExpr, OPR_LESS
    If sParms(1) <> "" Then
        nCol = fnIndexFromField(sParms(0))
        sOper = OPR_LESS
        sValue = sParms(1)
        Exit Sub
    End If
    
    #If DEVELOP Then
        MsgBox "Syntax error in where clause"
    #End If
End Sub

Property Get Table() As Object
    Set Table = tgTable
End Property

Public Sub UnboundWriteData(ByVal RowBuf As RowBuffer, WriteLocation As Variant)
    'no code, for compatibility only
End Sub

Public Sub ValidateCell(ByVal nCol As Integer, ByVal lRow As Long)
    If lRow < 0 Or bClearData Then
        Exit Sub
    End If
    Dim nFCol As Integer
    Dim sTemp As String
    Dim vTemp As Variant
    
    nFCol = fnColumnNumber(nCol)
    If nFCol >= 0 Then
        nCellChanged(nFCol, lRow) = nCellChanged(nFCol, lRow) And (Not STATUS_VALIDATED)
        fnValidate nCol, lRow, fnGetUserData(MakeBookmark(lRow), nCol)
    End If
End Sub

Property Let ValidatedFlag(ByVal nCol As Integer, ByVal lRow As Long, ByVal bStatus As Boolean)
    Dim nFCol As Integer
    nFCol = fnColumnNumber(nCol)
    If nFCol >= 0 Then
        If bStatus Then
            nCellChanged(nFCol, lRow) = nCellChanged(nFCol, lRow) Or STATUS_VALIDATED
        Else
            nCellChanged(nFCol, lRow) = nCellChanged(nFCol, lRow) And (Not STATUS_VALIDATED)
        End If
    End If

End Property

Property Get ValidCell(ByVal nCol As Integer, ByVal lRow As Long) As Boolean
    If lRow < 0 Or nRowCount < 0 Or lRow >= nRowCount Then
        If IsNull(GridArray(nCol, lRow)) Then
            Exit Property
        End If
        If Trim(GridArray(nCol, lRow)) = "" Then
            Exit Property
        End If
    End If
    
    Dim nFCol As Integer
    Dim sTemp As String
    Dim vTemp As Variant

    nFCol = fnColumnNumber(nCol)
    If nFCol >= 0 Then
        If Not bClearData Then
            If (nCellChanged(nFCol, lRow) And STATUS_VALIDATED) <> STATUS_VALIDATED Then
                fnValidate nCol, lRow, fnGetUserData(MakeBookmark(lRow), nCol)
            End If
        End If
        
        ValidCell = bValidCell(nFCol, lRow)
    Else
        ValidCell = True
    End If
End Property

Property Let ValidCell(ByVal nCol As Integer, ByVal lRow As Long, bStatus As Boolean)
    If lRow < 0 Or nRowCount < 0 Or lRow >= nRowCount Then
        Exit Property
    End If
    Dim nFCol As Integer
    nFCol = fnColumnNumber(nCol)
    If nFCol >= 0 Then
        bValidCell(nFCol, lRow) = bStatus
        nCellChanged(nFCol, lRow) = nCellChanged(nFCol, lRow) Or STATUS_VALIDATED
    End If
End Property

Public Sub SetDefault(nIndex As Integer, Optional vDefault As Variant)
    Dim nCol As Integer
    nCol = fnColumnNumber(nIndex)
    If nCol >= 0 And nCol < nMaxCol Then
        If IsMissing(vDefault) Then
            sDefaultValue(nCol) = ""
        Else
            sDefaultValue(nCol) = vDefault
        End If
    End If
End Sub

Public Sub SetValueByField(sStatement As String, Optional nStatus As Variant)
    Dim sParams(20) As String
    Dim nParams(20) As Integer
    Dim sParamValue(20) As Variant
    Dim nParamCount As Integer
    Dim bkMark As Variant
    Dim nPos As Integer
    Dim nLen As Integer
    Dim i As Long
    Dim j As Integer
    Dim lRows() As Long
    Dim lCount As Long
    
    If nRowCount <= 0 Or Trim(sStatement) = "" Then
        Exit Sub
    End If
    
    nPos = InStr(LCase(sStatement), "where")
    If nPos > 1 Then
        fnParseString sParams, sStatement, , 0, nPos - 1
    Else
        fnParseString sParams, sStatement
    End If
    
    nLen = Len(sStatement)
    fnParseAssigments nParams(), sParamValue(), nParamCount, sParams()
    If nLen - (nPos + 5) <= 0 Then
        Exit Sub
    End If
    If nPos > 1 Then
        subGetAffectedRows lRows, lCount, Right(sStatement, nLen - (nPos + 5))
    Else
        subGetAffectedRows lRows, lCount, ""
    End If
    For i = 0 To lCount - 1
        For j = 0 To nParamCount - 1
            If UCase(sParamValue(j)) = "NULL" Then
                GridArray(nParams(j), lRows(i)) = ""
            Else
                GridArray(nParams(j), lRows(i)) = sParamValue(j)
            End If
            If i = GetCurrentRowNumber Then
                bkMark = MakeBookmark(lRows(i))
                tgTable.Columns(nParams(j)).Text = fnGetUserData(bkMark, nParams(j))
            End If
        Next j
        If IsMissing(nStatus) Then
            GridArray(nFlagColumn, lRows(i)) = ROW_INITIAL
        Else
            GridArray(nFlagColumn, lRows(i)) = nStatus
            If nStatus <> ROW_INITIAL Then
                fnSaveBookMarks Bookmark(lRows(i))
            End If
        End If
        nPos = fnColumnNumber(j)
        If nPos >= 0 Then
            nCellChanged(nPos, lRows(i)) = nCellChanged(nPos, lRows(i)) And (Not STATUS_VALIDATED)
        End If
    Next i
    
End Sub

Private Sub fnFieldNameValue(sName As String, sValue As String, sSrc As String)
    sValue = ""
    If sSrc = "" Then
        sName = ""
        Exit Sub
    End If

    Dim nPos As Integer
    nPos = InStr(sSrc, "=")
    If nPos = 0 Then
        sName = sSrc
    Else
        sName = Trim$(Left$(sSrc, nPos - 1))
        nPos = Len(sSrc) - nPos
        If nPos > 0 Then
            sValue = Trim$(Right(sSrc, nPos))
        End If
    End If
End Sub

Private Sub fnParseString(sParam() As String, _
                          sSrc As String, _
                          Optional vDelm As Variant, _
                          Optional vStart As Variant, _
                          Optional vEnd As Variant)
    Dim i1 As Integer
    Dim i2 As Integer
    Dim k As Integer
    Dim nEnd As Integer
    Dim sDelm As String
    
    If IsMissing(vDelm) Then
        sDelm = ","
    Else
        sDelm = vDelm
    End If
    If IsMissing(vStart) Then
        i1 = 1
    Else
        i1 = vStart
    End If
    If IsMissing(vEnd) Then
        nEnd = Len(sSrc)
    Else
        nEnd = vEnd
    End If
    If i1 < 1 Then i1 = 1
    i2 = 1
    k = 0
    While i1 < nEnd And i2 > 0 And i2 < nEnd And k < UBound(sParam)
        i2 = InStr(i1, sSrc, sDelm)
        If i2 > i1 And i2 < nEnd Then
            sParam(k) = Trim$(Mid$(sSrc, i1, i2 - i1))
            k = k + 1
            i1 = i2 + Len(sDelm)
        End If
    Wend
    If i1 < nEnd Then
        sParam(k) = Trim$(Mid$(sSrc, i1, nEnd - i1 + 1))
        k = k + 1
    Else
        sParam(k - 1) = Trim$(Mid$(sSrc, i1, nEnd - i1 + 1))
    End If
    sParam(k) = ""
End Sub

Private Function fnParseAssigments(nCols() As Integer, sValue() As Variant, nCount As Integer, sSrc() As String) As Boolean
    Dim sName As String
    Dim sTValue As String
    Dim i As Long
    Dim nPos As Integer

    i = 0
    nCount = 0
    While i < 20 And sSrc(i) <> ""
        fnFieldNameValue sName, sTValue, sSrc(i)
        nPos = fnIndexFromField(LCase(sName))
        If nPos >= 0 Then
            nCols(nCount) = nPos
            sValue(nCount) = sTValue
            nCount = nCount + 1
        Else
            MsgBox "Fields specified in the criteria not in the table"
            Exit Function
        End If
        i = i + 1
    Wend

End Function

Public Sub SetValueByRow(sStatement As String, ByVal lRow As Long, Optional nStatus As Variant)
    Dim sParams(20) As String
    Dim nParams(20) As Integer
    Dim sParamValue(20) As Variant
    Dim nParamCount As Integer

    Dim nPos As Integer
    Dim nLen As Integer
    Dim j As Integer
    
    nLen = Len(sStatement)
    fnParseString sParams, sStatement
    fnParseAssigments nParams(), sParamValue(), nParamCount, sParams()
    
    For j = 0 To nParamCount - 1
        GridArray(nParams(j), lRow) = sParamValue(j)
    Next j
    If IsMissing(nStatus) Then
        GridArray(nFlagColumn, lRow) = ROW_INITIAL
    Else
        GridArray(nFlagColumn, lRow) = nStatus
    End If
End Sub

Public Function ValidData(Optional ByVal lRow As Variant) As Boolean
    If tgTable Is Nothing Or nRowCount < 0 Then
        ValidData = False
        Exit Function
    End If
    
    Dim i As Long
    Dim j As Integer
    Dim lBegin As Long
    Dim lEnd As Long
    Dim sTemp As String
    Dim vTemp As Variant
    Dim bFlag As Boolean
    
    ValidData = True
    If IsMissing(lRow) Then
        lBegin = 0
        lEnd = nRowCount - 1
        If lEnd < GetCurrentRowNumber Then
            lEnd = GetCurrentRowNumber
        End If
        bFlag = True
    Else
        i = lRow
        If i < 0 Or i > nRowCount Then
            ValidData = False
            Exit Function
        End If
        For j = 0 To nColumnCount - 1
            fnValidate nEditColumns(j), i, fnGetUserData(MakeBookmark(i), nEditColumns(j))
            
            If Not bValidCell(j, i) Then
                ValidData = False
                Exit Function
            End If
        Next j
        Exit Function
    End If
    
    If bAllowAddNew Then
        Do
            For j = 0 To nColumnCount - 1
                If GridArray(nEditColumns(j), lEnd) <> "" Then
                    bFlag = False
                    Exit Do
                End If
            Next j
            lEnd = lEnd - 1
        Loop Until Not bFlag Or lEnd < lBegin
    End If

    For i = lBegin To lEnd
        If i < lEnd Or Not bAllowAddNew Or GridArray(nFlagColumn, i) <> "" Then   'Problem ?
            For j = 0 To nColumnCount - 1
                fnValidate nEditColumns(j), i, fnGetUserData(MakeBookmark(i), nEditColumns(j))
                
                If Not bValidCell(j, i) Then
                    ValidData = False
                    Exit Function
                End If
            Next j
        End If
    Next i
End Function

Public Function FillWithSQL(dbSource As DataBase, _
                            strSQL As String, _
                            Optional nDB As Variant, _
                            Optional bRefresh As Variant, _
                            Optional vValidate As Variant, _
                            Optional vClear As Variant, _
                            Optional bResetSortOrder As Boolean = True) As Boolean
    Dim rsTemp As Recordset
    Dim nDBType As Integer
    If IsMissing(nDB) Then
        nDBType = DB_INFORMIX
    Else
        nDBType = nDB
    End If
    On Error GoTo SQLError
    Select Case nDBType
        Case DB_INFORMIX
            Set rsTemp = dbSource.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
        Case DB_ACCESS
            Set rsTemp = dbSource.OpenRecordset(strSQL, dbOpenSnapshot)
    End Select
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveLast
        rsTemp.MoveFirst
    End If
    
    FillWithRecordset rsTemp, bRefresh, vValidate, vClear, bResetSortOrder
    rsTemp.Close
    Set rsTemp = Nothing
    FillWithSQL = True

quitsub:
    On Error GoTo 0
    Exit Function
SQLError:
    subShowODBCError "", strSQL
    FillWithSQL = False
    Resume quitsub
    
End Function

Private Sub subShowODBCError(Optional vMsg As Variant, _
                             Optional vSQL As Variant)
    Dim i As Integer
    Dim sMsgs As String
    Dim sNumbers As String
    Dim sODBCErrors As String
    
    #If DEVELOP Then
        Dim strSQL As String
        If IsMissing(vSQL) Then
            strSQL = ""
        Else
            strSQL = vSQL
        End If
    #End If
    
    sODBCErrors = Err.Description
    If Err.Number = 3146 Then
        If Not engDBEngine Is Nothing Then
            With engDBEngine.Errors
                If .Count > 0 Then
                    For i = 0 To .Count - 2
                        sMsgs = sMsgs & "Number: " & .Item(i).Number & Space(5) & .Item(i).Description & vbCrLf
                    Next
                End If
                If .Count <= 2 Then
                    sNumbers = ""
                Else
                    sNumbers = "s"
                End If
            End With
            sODBCErrors = "The following error" & sNumbers & " occurred while doing an ODBC query:" & vbCrLf & vbCrLf _
                           & vbCrLf & sMsgs
        End If
    End If
    
    Dim sMsg As String
    If IsMissing(vMsg) Then
        #If DEVELOP Then
            sMsg = "An error occurred while doing a SQL query" & vbCrLf & vbCrLf & "Error# " & CStr(Err.Number) & vbCrLf & Err.Description
            sMsg = sMsg & vbCrLf & vbCrLf & strSQL & vbCrLf & vbCrLf & sODBCErrors
            Clipboard.SetText strSQL
        #Else
            sMsg = ""
        #End If
    Else
        sMsg = vMsg
        #If DEVELOP Then
            If Trim(sMsg) = "" Then
                sMsg = "SQL: " & strSQL & vbCrLf & vbCrLf & sODBCErrors
            Else
                sMsg = sMsg & vbCrLf & vbCrLf & "SQL: " & strSQL & vbCrLf & vbCrLf & sODBCErrors
            End If
            Clipboard.SetText strSQL
        #Else
            If Trim(sMsg) = "" Then
                sMsg = sODBCErrors
            Else
                sMsg = sMsg & vbCrLf & vbCrLf & sODBCErrors
            End If
        #End If
    End If
    If sMsg <> "" Then
        MsgBox sMsg, vbOKOnly + vbCritical, App.Title
    End If
    Err.Clear

End Sub

Property Get dbAccess() As Integer
    dbAccess = DB_ACCESS
End Property

Property Get dbInformix() As Integer
    dbInformix = DB_INFORMIX
End Property

Public Function GetCurrentRowNumber() As Long

    Dim lRow As Long
    If nRowCount = 0 Then
        GetCurrentRowNumber = 0
        Exit Function
    End If
    
    If IsNull(tgTable.FirstRow) Then
        lRow = 0
    Else
        lRow = Val(tgTable.FirstRow)
    End If
    If IsNull(tgTable.Row) Then
        GetCurrentRowNumber = lRow
    Else
        GetCurrentRowNumber = lRow + Val(tgTable.Row)
    End If

End Function

Public Function GetChangedRowNumber(ByVal lIndex As Long) As Long
    Dim lRow As Long
    
    If lIndex < 0 Or lIndex >= nChangedCount Then
        GetChangedRowNumber = -1
        Exit Function
    End If
    GetChangedRowNumber = fnRowNumberByIndex(bmChangedRows(lIndex))

End Function

Public Function Bookmark(Optional ByVal vRow As Variant) As Variant
    Dim nRow As Long
    
    If IsMissing(vRow) Then
        nRow = GetCurrentRowNumber
    Else
        nRow = vRow
    End If
    If nRow >= 0 And nRow <= nRowCount Then
        Bookmark = GridArray(nBookMarkColumn, nRow)
    Else
        Bookmark = Null
    End If
End Function

Private Sub fnSaveBookMarks(Bookmk)
    ' Save the book marks of the changed rows
    Dim lRow As Long

    If nChangedCount > nMaxChanged Then
        nMaxChanged = nMaxChanged + ROW_INCREMENT
        ReDim Preserve bmChangedRows(nMaxChanged)
    End If
    For lRow = 0 To nChangedCount - 1
        If bmChangedRows(lRow) = Bookmk Then
            Exit Sub
        End If
    Next
    bmChangedRows(lRow) = Bookmk
    If lRow >= nChangedCount Then
        nChangedCount = nChangedCount + 1
    End If
End Sub

Private Sub fnShowStatus(lRow As Long, nCol As Integer, Optional vAddRow As Variant)
    If lblStatusbar Is Nothing Then
        Exit Sub
    End If

    Dim nECol As Integer
    Dim bFlag As Boolean
    
    nECol = fnColumnNumber(nCol)
    If nECol < 0 Or lRow < 0 Then
        Exit Sub
    End If

    If bValidCell(nECol, lRow) Then
        If (nCellChanged(nECol, lRow) And STATUS_CHANGED) = STATUS_CHANGED Then
            lblStatusbar.ForeColor = &H8000&  'Green text
            lblStatusbar.Caption = "Data is valid"
            lblStatusbar.Font.Bold = False
        Else
            lblStatusbar.Caption = sPrompt(nECol)
            lblStatusbar.ForeColor = &H80000012
            lblStatusbar.Font.Bold = False
        End If
    Else
        Dim bAddRow As Boolean
        bAddRow = False
        If Not IsMissing(vAddRow) Then
            bAddRow = vAddRow
        End If
        If bForcedEditing Or bAddRow Or ((nCellChanged(nECol, lRow) And STATUS_WARNING) = STATUS_WARNING) Then
            If nAddRowCount >= 1 Then
                lblStatusbar.Caption = sDownArrowMessage
            Else
                If sErrMessage(nECol) = "" Then
                    lblStatusbar.Caption = "Data is invalid; " & sPrompt(nECol)
                Else
                    lblStatusbar.Caption = sErrMessage(nECol)
                End If
            End If
            lblStatusbar.ForeColor = &HFF&       'Red Text on
            lblStatusbar.Font.Bold = True
        Else
            lblStatusbar.Caption = sPrompt(nECol)
            lblStatusbar.ForeColor = &H80000012
            lblStatusbar.Font.Bold = False
        End If
    End If
    lblStatusbar.Refresh
End Sub

Public Function Change() As Boolean
'
' Check the changed text against the regular expression for that column
'

    Dim nCol As Integer
    Dim szText As String
    Dim nCode As Integer
        
    Change = False
    
    If tgTable Is Nothing Then
        Exit Function
    End If

    nCol = fnColumnNumber(tgTable.Col)
    If nCol < 0 Then
        Exit Function
    End If
    
    szText = fnCellValue(tgTable.Col, GetCurrentRowNumber)

    If sPatternCode(nCol) = "" Then
        Change = True
        Exit Function
    Else
        If szText = "" Then
            Change = False
            Exit Function
        End If
    End If

    ' check the pattern
    #If REGULAR_EXP_INCLUDED Then
        nCode = PRegExpMatch(sPatternCode(nCol), Len(sPatternCode(nCol)), szText)
    #End If
    ' highlight any bad text
    If nCode > 0 Then
        tgTable.SelStart = nCode - 1
        tgTable.SelLength = Len(szText) - tgTable.SelStart
        If tgTable.SelLength > 0 Then
            Beep
        End If
    End If
    
    Change = (nCode = 0)    ' true if text is okay

End Function

Public Sub DeleteRow(Optional vBookmark As Variant, Optional ByVal vRebind As Variant)
    'The bookmark of the true grid may be different from
    ' the bookmark of the class because once a line is deleted
    ' the bookmark of the grid will shift.
    'The bookmark required for this function is that of the class
    'which can be obtained by calling Bookmark suppling the row number
    
    If nRowCount = 0 Then
        ClearData
        Exit Sub
    End If
    
    Dim vBkMark As Variant
    
    If IsMissing(vBookmark) Then
        'Current row book mark
        vBkMark = Bookmark(GetCurrentRowNumber)
    Else
        vBkMark = vBookmark
    End If
    If Not fnDeleteRow(vBkMark) Then 'Bookmark = Null
        MsgBox "Delete row error"
    Else
        If nRowCount = 0 Then
            If nColumnCount > 0 Then
                nLastColumn = FirstEditableCol
                tgTable.Col = nLastColumn
                nCellChanged(0, 0) = STATUS_NOCHANGE
            End If
        End If
            
        If Not m_tgsSelector Is Nothing Then
            m_tgsSelector.ClearSelectedCells
        End If
        
        If IsMissing(vRebind) Then
            Rebind
        Else
            If vRebind Then
                Rebind
            End If
        End If
        If nLastTableRow >= nRowCount And nLastTableRow > 0 Then
            nLastTableRow = nLastTableRow - 1
        End If
    End If
End Sub

'david 05/26/2002
'NOTE:
'(FillWithArray vData, True, False) IS NOT WORKING
'if a cell validation in the main form uses the value from other cells,
'the validation will fail because those values are not available.

'THIS MEANS THE CELL(S) VALIDATION IN THIS FUNCTION IS NOT 100% GURANTEE WORKING
'THE REVALIDATION OF THE GRID SHOULD BE PERFORMED IN THE CALLING FUNCTION
'AFTER THE GRID HAS BEEN FILLED WITH THE DATA
Public Sub FillWithArray(arryData() As Variant, _
                         Optional vClear As Variant, _
                         Optional vValid As Variant, _
                         Optional bResetSortOrder As Boolean = True, _
                         Optional bDoEvents As Boolean = True)
    Dim i As Long
    Dim lRow As Long
    Dim nCol As Integer
    Dim nMCol As Integer
    Dim bClear As Boolean
    Dim bValid As Boolean
    Dim lStart As Long
    Dim nECol As Integer
    
    'david 02/09/2004  #401172-1
    If Not m_tgsSelector Is Nothing Then
        m_tgsSelector.ClearSelectedCells
    End If
    ''''''''''''''''''''''''''''
    
    If IsMissing(vClear) Then
        bClear = True
    Else
        bClear = vClear
    End If
    If IsMissing(vValid) Then
        bValid = True
    Else
        bValid = vValid
    End If
    
    'david 05/06/2003  #408907
    'to trap the out of memory RTE
    On Error GoTo errHandler
    ''''''''''''''''''''''''''

    If bClear Then
        lStart = 0
        nMaxRow = fnUbound(arryData, 2)
        nMCol = fnUbound(arryData, 1) + 1
        If nMCol > nMaxCol Then
            nMCol = nMaxCol
        End If
        fnInitialParams
    Else
        lStart = nRowCount
        nMaxRow = nRowCount + fnUbound(arryData, 2)
        ReDim Preserve GridArray(nMaxCol + 1, nMaxRow + 1)
        If nColumnCount > 0 Then
            ReDim Preserve bValidCell(nColumnCount - 1, nMaxRow + 1)
            ReDim Preserve nCellChanged(nColumnCount - 1, nMaxRow + 1)
        End If
        nMCol = fnUbound(arryData, 1) + 1
        If nMCol > nMaxCol Then
            nMCol = nMaxCol
        End If
    End If

    For i = lStart To nMaxRow
        lRow = i - lStart
        For nCol = 0 To nMCol - 1
            GridArray(nCol, i) = arryData(nCol, lRow)
            nECol = fnColumnNumber(nCol)
            If nECol >= 0 Then
                If bValid Then
                    bValidCell(nECol, i) = True
                    nCellChanged(nECol, i) = STATUS_VALIDATED
                Else
                    nCellChanged(nECol, i) = STATUS_CHANGED
                    fnValidate nCol, i, GridArray(nCol, i)
                End If
            End If
        Next nCol
        GridArray(nFlagColumn, i) = ROW_INITIAL
        GridArray(nBookMarkColumn, i) = fnGetNewBookmark
        nRowCount = nRowCount + 1
    Next i

    bClearData = True
    bErrorFlag = False
    bValidating = True
    tgTable.Row = 0
    If bClear Then
        If FirstEditableCol < tgTable.Columns.Count Then
            tgTable.Col = FirstEditableCol
        End If
        tgTable.Rebind
        
        'david 03/16/2004
        If bDoEvents Then
            DoEvents
        End If
    Else
        Rebind
        'david 03/16/2004
        If bDoEvents Then
            DoEvents
        End If
    End If
    
    bClearData = False
    bValidating = False
    
    'david 08/10/2002
    'restore the original sort order
    If bResetSortOrder Then
        subResetSortOrder
    End If
    '''''''''''''''''
    
extFill:
    On Error GoTo 0
    Exit Sub

errHandler:
    'david 05/06/2003  #408907
    'to trap the out of memory RTE
    MsgBox "An error has occurred in function FillWithArray()." + vbCrLf + vbCrLf + "Error Code: " & Err.Number _
        & vbCrLf & "Error Desc:" & Err.Description & vbCrLf + vbCrLf + "Ppopulating grid data failed.", vbCritical
    ClearData
    Resume extFill
End Sub

Private Function fnUbound(arryTemp() As Variant, Optional vIndex As Variant)
    Dim nIndex As Integer
    
    If IsMissing(vIndex) Then
        nIndex = 1
    Else
        nIndex = vIndex
    End If
    
    fnUbound = -1
    On Error Resume Next
    fnUbound = UBound(arryTemp, nIndex)
End Function

Public Sub FillWithRecordset(rsTemp As Recordset, _
                             Optional bRefresh As Variant, _
                             Optional vValidate As Variant, _
                             Optional vClear As Variant, _
                             Optional bResetSortOrder As Boolean = True)
    Dim lRow As Long
    Dim nCol As Integer
    Dim i As Integer
    Dim sField As String
    Dim bValidate As Boolean
    Dim bClear As Boolean
    Dim lStart As Long
    ReDim nColPtr(nMaxCol - 1) As Integer
    
    'david 02/09/2004  #401172-1
    If Not m_tgsSelector Is Nothing Then
        m_tgsSelector.ClearSelectedCells
    End If
    ''''''''''''''''''''''''''''

    'david 08/21/00
    If IsMissing(vClear) Then
        bClear = True
    Else
        bClear = vClear
    End If

    If rsTemp.RecordCount = 0 Then
        If bClear Then
            ClearData
            nRowCount = 0
        End If
        Exit Sub
    Else
        rsTemp.MoveLast
    End If
    
    If IsMissing(vValidate) Then
        bValidate = False
    Else
        bValidate = vValidate
    End If

    'david 05/06/2003  #408907
    'to trap the out of memory RTE
    On Error GoTo errHandler
    ''''''''''''''''''''''''''

    If bClear Then
        nMaxRow = rsTemp.RecordCount - 1
        lStart = 0
        fnInitialParams
    Else
        lStart = nRowCount
        nMaxRow = nRowCount + rsTemp.RecordCount - 1
        
        ReDim Preserve GridArray(nMaxCol + 1, nMaxRow + 1)
        
        If nColumnCount > 0 Then
            ReDim Preserve bValidCell(nColumnCount - 1, nMaxRow + 1)
            ReDim Preserve nCellChanged(nColumnCount - 1, nMaxRow + 1)
        End If
    End If
    
    rsTemp.MoveFirst
    
    On Error GoTo errFill
    
    For nCol = 0 To nMaxCol - 1
        sField = LCase(Trim(tgTable.Columns(nCol).DataField))
        nColPtr(nCol) = -1
        For i = 0 To rsTemp.Fields.Count - 1
            If LCase(Trim(rsTemp.Fields(i).Name)) = sField Then
                nColPtr(nCol) = i
                Exit For
            End If
        Next i
    Next nCol

    For lRow = lStart To nMaxRow
        For nCol = 0 To nMaxCol - 1
            If nColPtr(nCol) >= 0 Then
                If IsNull(rsTemp.Fields(nColPtr(nCol))) Then
                    GridArray(nCol, lRow) = Null
                Else
                    GridArray(nCol, lRow) = Trim(rsTemp.Fields(nColPtr(nCol)))
                End If
            End If
        Next nCol
        
        GridArray(nFlagColumn, lRow) = ROW_INITIAL
        GridArray(nBookMarkColumn, lRow) = fnGetNewBookmark
        nRowCount = nRowCount + 1
        
        For nCol = 0 To nColumnCount - 1
            bValidCell(nCol, lRow) = True
            If bValidate Then
                nCellChanged(nCol, lRow) = STATUS_NOCHANGE 'STATUS_CHANGED
            Else
                nCellChanged(nCol, lRow) = STATUS_VALIDATED
            End If
        Next nCol
        rsTemp.MoveNext
    Next lRow
    
    bClearData = True
    bValidating = True
    tgTable.Row = 0
    
    If IsMissing(bRefresh) Then
        tgTable.Rebind
    Else
        If bRefresh Then
            tgTable.Refresh
        End If
    End If

    DoEvents
    bClearData = False
    bValidating = False

    'david 08/10/2002
    'restore the original sort order
    If bResetSortOrder Then
        subResetSortOrder
    End If
    '''''''''''''''''

extFill:
    On Error GoTo 0
    Exit Sub

errFill:
    MsgBox Err.Description, nCRITICAL   ', App.Caption
    Err.Clear
    Resume extFill

errHandler:
    'david 05/06/2003  #408907
    'to trap the out of memory RTE
    MsgBox "An error has occurred in function FillWithArray()." + vbCrLf + vbCrLf + "Error Code: " & Err.Number _
        & vbCrLf & "Error Desc:" & Err.Description & vbCrLf + vbCrLf + "Ppopulating grid data failed.", vbCritical
    ClearData
    Resume extFill
End Sub

Private Function fnColumnNumber(ByVal nIndex As Integer) As Integer
    Dim i As Integer
    
    fnColumnNumber = -1
    For i = 0 To nColumnCount - 1
        If nEditColumns(i) = nIndex Then
            fnColumnNumber = i
            Exit Function
        End If
    Next
End Function

Private Function fnDeleteRow(bkMark As Variant) As Boolean
    fnDeleteRow = False

    Dim Index As Long
    
    Index = RowByBookMark(bkMark)

    If Index < 0 Or Index > nRowCount Then
        ' invalid row, delete is invalid
        Exit Function
    End If
    
    Dim i%, j%
    If Index = nRowCount And bAllowAddNew Then
        For j% = 0 To nMaxCol + 1
            GridArray(j%, Index) = ""
        Next j%
    Else
        nRowCount = nRowCount - 1 ' decrease number of rows
        
        On Error Resume Next
        ' Shift the data in the array, moving the rows down one
        For i% = Index To nRowCount
            If i = nRowCount Then
                For j% = 0 To nMaxCol + 1
                    GridArray(j%, i) = ""
                Next j%
            Else
                For j% = 0 To nMaxCol + 1
                    GridArray(j%, i%) = GridArray(j%, i% + 1)
                Next j%
            End If
            For j = 0 To nColumnCount - 1
                nCellChanged(j, i) = nCellChanged(j, i + 1)
                bValidCell(j, i) = bValidCell(j, i + 1)
            Next j
        Next i%
    End If
    
    fnDeleteRow = True
    
    For i = 0 To nChangedCount - 1
        If bmChangedRows(i) = bkMark Then
            Exit For
        End If
    Next i
    If i < nChangedCount Then
        For j = i To nChangedCount - 2
            bmChangedRows(j) = bmChangedRows(j + 1)
        Next
        nChangedCount = nChangedCount - 1
    End If
End Function

Private Function fnGetPatternCode(sRegExp As String)
'
'   store regular expressions for the columns
'
    Dim szBuffer As String
    Dim nSize As Integer

    szBuffer = Space(4096)
    
    #If REGULAR_EXP_INCLUDED Then
        nSize = GetRegExpPCode(sRegExp, szBuffer)
        If nSize <= 0 Then
            fnGetPatternCode = ""
        Else
            fnGetPatternCode = Left$(szBuffer, nSize)
        End If
    #Else
        fnGetPatternCode = ""
    #End If
End Function

Private Function fnGotoNextRow(Optional bAction As Variant) As Boolean
    Dim lRow As Long
    Dim i As Integer
    Dim nECol As Integer
    
    On Error GoTo retFalse
    fnGotoNextRow = False
    lRow = GetCurrentRowNumber()
    nECol = fnColumnNumber(tgTable.Col)
    
    If lRow >= nRowCount - 1 Then
        If Not bAllowAddNew Then
            Exit Function
        End If
    End If
    
    If nColumnCount = 1 Then
        If lRow > nRowCount Or (lRow >= nRowCount And nRowCount > 0) Then
            Exit Function
        End If
    ElseIf nColumnCount > 1 Then
        If lRow > nRowCount Then
            Exit Function
        End If
    End If
    
    If lRow >= nRowCount - 1 Then
        If nECol >= 0 Then
            If Not fnCheckRowValid(nECol, lRow) Then
                Exit Function
            End If
        End If
        
        For i = nColumnCount - 1 To 0 Step -1
            If Not fnCheckRowValid(i, lRow) Then
                Exit Function
            End If
        Next
    Else
        If bHighlightInvalidCell And bForcedEditing Then
            If nColumnCount > 0 Then
                If nECol = nColumnCount - 1 And lRow >= 0 Then
                    fnCheckRowValid nECol, lRow
                End If
            End If
        End If
    End If
    
    If lRow >= nRowCount - 1 Then
        subAddNewRow nColumnForNewRow
    End If

    If IsMissing(bAction) Then
        tgTable.Row = tgTable.Row + 1
    Else
        If bAction Then
            tgTable.Row = tgTable.Row + 1
        End If
    End If
    
    fnGotoNextRow = True
    Exit Function
    
retTrue:
    On Error GoTo 0
    Exit Function
retFalse:
    If Err.Number = 13 Then
        lRow = tgTable.Row
        Resume Next
    ElseIf Err.Number = 6148 Then
        SendKeys "{DOWN}", True
    Else
        fnGotoNextRow = False
        Err.Clear
        Resume retTrue
    End If
End Function

Private Function fnGotoPreviousRow() As Boolean
    On Error GoTo retFalse1
    fnGotoPreviousRow = False
    If nMaxRow = 0 Then
        Exit Function
    End If
    If tgTable.Row > 0 Then
        tgTable.Row = tgTable.Row - 1
    Else
        If tgTable.FirstRow > 0 Then
            tgTable.FirstRow = tgTable.FirstRow - 1
        Else
            Exit Function
        End If
        'tgTable.Row = tgTable.Row - 1
    End If
    fnGotoPreviousRow = True
retTrue1:
    On Error GoTo 0
    Exit Function
retFalse1:
    Err.Clear
    Resume retTrue1
End Function

Property Set Form(frmMain As Object)
    Set frmParent = frmMain
End Property

Property Get GetChangedRowCount()
    GetChangedRowCount = nChangedCount
End Property

Public Sub GetChangedRow(vData() As Variant, nCount As Integer, ByVal lIndex As Long)
    Dim i As Integer
    Dim lRow As Long
    ReDim vData(nMaxCol - 1)
    
    If lIndex < 0 Or lIndex >= nChangedCount Then
        nCount = 0
        Exit Sub
    End If
    lRow = fnRowNumberByIndex(bmChangedRows(lIndex))
    If lRow >= 0 And lRow < RowCount Then
        For i = 0 To nMaxCol - 1
            vData(i) = GridArray(i, lRow)
        Next i
    End If
    nCount = nMaxCol
End Sub

'return True if the text is changed
Public Function KeyDown(KeyCode As Integer, Shift As Integer, Optional bRegExpCheck As Boolean = False) As Boolean
    Const sValidKeys As String = " `1234567890-=qwertyuiop[]\asdfghjkl;'zxcvbnm,./~!@#$%^&*()_+QWERTYUIOP{}|ASDFGHJKL:""""ZXCVBNM<>?"
    
    Dim lRow As Long
    Dim nCol As Integer
    
    'david 12/11/2002
    Dim bShiftDown As Boolean
    Dim bAltDown As Boolean
    Dim bCtrlDown As Boolean
    
    bShiftDown = (Shift And vbShiftMask) > 0
    bAltDown = (Shift And vbAltMask) > 0
    bCtrlDown = (Shift And vbCtrlMask) > 0
    '''''''''''''''''
    
    On Error GoTo errKeyDown
    
    'david 05/07/2002
    If Shift = 0 And bRegExpCheck Then
        Dim nCode As Integer
        Dim nLen As Integer
        Dim szText As String
        Dim sTemp As String
        
        'navigation keys
        Select Case KeyCode
        Case vbKeyReturn, vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyHome, vbKeyEnd, _
           vbKeyPageUp, vbKeyPageDown, vbKeyInsert, vbKeyF1, vbKeyF2, vbKeyF3, vbKeyF4, vbKeyF5, _
           vbKeyF6, vbKeyF7, vbKeyF8, vbKeyF9, vbKeyF10, vbKeyF11, vbKeyF12
            KeyDown = False
        Case Else
            If tgTable.EditActive Then
                sTemp = fnCellValue(tgTable.Col, lEditingRow)
                
                If KeyCode = vbKeyBack Then
                    If sTemp <> "" Then
                        If tgTable.SelStart <= 0 And tgTable.SelLength = 0 Then
                            szText = sTemp
                            KeyDown = False
                        Else
                            On Error Resume Next
                            If tgTable.SelLength = 0 Then
                                If tgTable.SelStart >= Len(sTemp) Then
                                    szText = Left(sTemp, tgTable.SelStart - 1)
                                Else
                                    szText = Mid(sTemp, 1, tgTable.SelStart - 1) + Mid(sTemp, tgTable.SelStart + 1)
                                End If
                            Else
                                szText = Mid(sTemp, 1, tgTable.SelStart) + Mid(sTemp, tgTable.SelStart + tgTable.SelLength + 1)
                            End If
                            On Error GoTo 0
                        End If
                        
                        If szText <> sTemp Then
                            KeyDown = True
                        End If
                    Else
                        szText = ""
                    End If
                ElseIf KeyCode = vbKeyDelete Then
                    If sTemp <> "" Then
                        If tgTable.SelStart >= Len(sTemp) And tgTable.SelLength = 0 Then
                            szText = sTemp
                            KeyDown = False
                        Else
                            If tgTable.SelStart <= 0 Then
                                On Error Resume Next
                                If tgTable.SelLength = 0 Then
                                    szText = Mid(sTemp, tgTable.SelStart + 2)
                                Else
                                    szText = Mid(sTemp, tgTable.SelStart + tgTable.SelLength + 1)
                                End If
                                On Error GoTo 0
                            Else
                                On Error Resume Next
                                If tgTable.SelLength = 0 Then
                                    szText = Mid(sTemp, 1, tgTable.SelStart) + Mid(sTemp, tgTable.SelStart + 2)
                                Else
                                    szText = Mid(sTemp, 1, tgTable.SelStart) + Mid(sTemp, tgTable.SelStart + tgTable.SelLength + 1)
                                End If
                                On Error GoTo 0
                            End If
                            
                            If szText <> sTemp Then
                                KeyDown = True
                            End If
                        End If
                    Else
                        szText = ""
                    End If
                ElseIf Shift = 0 And InStr(sValidKeys, Chr(KeyCode)) > 0 Then
                    If tgTable.SelLength = 0 Then
                        If tgTable.SelStart <= 0 Then
                            szText = Chr(KeyCode) & sTemp
                        ElseIf tgTable.SelStart >= Len(sTemp) Then
                            szText = sTemp & Chr(KeyCode)
                        Else
                            szText = Mid(sTemp, 1, tgTable.SelStart) + Chr(KeyCode) + Mid(sTemp, tgTable.SelStart + 1)
                        End If
                    Else
                        szText = Mid(sTemp, 1, tgTable.SelStart) + Chr(KeyCode) + Mid(sTemp, tgTable.SelStart + tgTable.SelLength + 1)
                    End If
                End If

                If Not KeyDown And sTemp <> szText Then
                    nCode = PRegExpMatch(sPatternCode(nCol), Len(sPatternCode(nCol)), szText)
                    
                    If nCode < 0 Or (nCode <> 0 And nCode <> tgTable.SelStart + 2) Then
                        KeyDown = False
                    Else
                        KeyDown = True
                    End If
                End If
            Else
                KeyDown = True
            End If
        End Select
    End If
    
    If nKeyProcessing <> KEY_P_FINISHED Then
        If nKeyProcessing = KEY_P_CHANGE_ROW Then
            nKeyProcessing = KEY_P_FINISHED
        Else
            KeyCode = 0
            Exit Function
        End If
    End If
    
    nKeyProcessing = KEY_P_OTHER
    
    If nReturnKeyAction = RETURN_KEY_RIGHT Then
        If nColumnCount = 0 Then
            If KeyCode = vbKeyReturn Then
                fnGotoNextRow
            ElseIf bCtrlDown And (Not bShiftDown) Then
                'david 12/11/2002
                'new feature, ctrl-home, ctrl-end
                If KeyCode = vbKeyHome Or KeyCode = vbKeyEnd Then
                    If Not tgTable.EditActive Then
                        If KeyCode = vbKeyHome Then
                            lRow = 0
                            tgTable.Bookmark = MakeBookmark(lRow)
                            KeyCode = 0
                            bCellEditing = False
                            tgTable.Col = FirstEditableCol
                        Else
                            lRow = RowCount - 1
                            tgTable.Bookmark = MakeBookmark(lRow)
                            tgTable.Col = LastEditableCol
                        End If
        
                        nKeyProcessing = KEY_P_CHANGE_ROW
                    End If
                End If
                '''''''''''''''''
            End If
            
            nKeyProcessing = KEY_P_FINISHED
            Exit Function
        End If
    End If

    If (KeyCode = vbKeyReturn And nReturnKeyAction = RETURN_KEY_RIGHT) Then
        If bCellEditing Then
            bCellEditing = False
            If tgTable.Col = LastEditableCol Then
                If fnGotoNextRow(False) Then
                    nKeyProcessing = KEY_P_CHANGE_ROW
                    If FirstEditableCol < tgTable.Columns.Count Then
                        If FirstEditableCol >= 0 Then
                            tgTable.Col = FirstEditableCol
                        End If
                        'david 02/14/2001
                        'commented out the DoEvents. NOT SURE WHAT'S THE USE.
                        'sometime the it causes problem in last editable column
                        'when the data entered is not valid. The error row is set
                        'correctly but the focus will go to the first editable column.
                        'DoEvents
                    End If
                    If tgTable.Col = FirstEditableCol Then
                        On Error Resume Next
                        tgTable.Row = tgTable.Row + 1
                    End If
                    KeyCode = 0
                    nKeyProcessing = KEY_P_FINISHED
                    Exit Function
                End If
            End If
        End If

        If tgTable.EditActive Then
            If tgTable.Col < LastEditableCol Then
                tgTable.EditActive = False
            Else
                KeyCode = 0
            End If
        End If
    End If
    
    If KeyCode = vbKeyRight Or (KeyCode = vbKeyReturn And nReturnKeyAction = RETURN_KEY_RIGHT) Then
        If Not tgTable.EditActive Or tgTable.SelStart = Len(fnCellValue(tgTable.Col, GetCurrentRowNumber)) Then
            'DoEvents don't know the effects of commenting out this line
            If tgTable.Col >= LastEditableCol Then
                'add a row
                If fnGotoNextRow(False) Then
                    nKeyProcessing = KEY_P_CHANGE_ROW
                    If FirstEditableCol < tgTable.Columns.Count Then
                        tgTable.Col = FirstEditableCol
                        'comments the doevents to block the event rowcolChange, let it fire later
                        'otherwise if you have method like GotoNextcolumn or change the column of
                        'the grid, if will not go to the new row even you press enter key or arrow key
                        'in the last column of the grid, but do so may cause the message showing incorrectly
                        'or validation problem, I'm not sure, if meet the problem, be careful!!
                        'junsong 05/25/01
                        'DoEvents
                    End If
                    If tgTable.Col = FirstEditableCol Then
                        On Error Resume Next
                        tgTable.Row = tgTable.Row + 1
                    End If
                End If
                KeyCode = 0
            End If
            bCellEditing = False
        End If
    ElseIf KeyCode = vbKeyLeft Then
        If Not tgTable.EditActive Then
            If tgTable.Col <= FirstEditableCol Then
                If fnGotoPreviousRow Then
                    If LastEditableCol < tgTable.Columns.Count Then
                        tgTable.Col = LastEditableCol
                    End If
                End If
                KeyCode = 0
            End If
            bCellEditing = False
        End If
    ElseIf KeyCode = vbKeyDown Or (KeyCode = vbKeyReturn And nReturnKeyAction = RETURN_KEY_DOWN) Then
        If fnGotoNextRow Then
            If GetCurrentRowNumber = nRowCount Then
                If tgTable.Columns(tgTable.Col).Text = "" Then
                    nCol = fnColumnNumber(tgTable.Col)
                    If nCol >= 0 Then
                        Do
                            If Not IsNull(tgTable.Columns(nEditColumns(nCol)).Text) Then
                                If tgTable.Columns(nEditColumns(nCol)).Text <> "" Then
                                    Exit Do
                                End If
                            End If
                            nCol = nCol - 1
                        Loop Until nCol < 0
                        If nCol < 0 Then
                            tgTable.Col = FirstEditableCol
                        Else
                            If nCol + 1 < nColumnCount Then
                                tgTable.Col = nEditColumns(nCol + 1)
                            End If
                        End If
                    End If
                End If
            End If
        Else
            If nReturnKeyAction = RETURN_KEY_DOWN Then
                If GetCurrentRowNumber = nRowCount - 1 Then
                    If fnGotoNextColumn Then
                        tgTable.FirstRow = 0
                    End If
                End If
            End If
        End If
        KeyCode = 0
        bCellEditing = False
    ElseIf KeyCode = vbKeyUp Then
        If nReturnKeyAction = RETURN_KEY_DOWN Then
            If GetCurrentRowNumber = 0 Then
                If fnGotoPrevColumn Then
                    lRow = RowCount - tgTable.VisibleRows
                    If lRow < 0 Then
                        lRow = 0
                    End If
                    tgTable.FirstRow = lRow
                    lRow = tgTable.VisibleRows
                    If lRow > RowCount Then
                        lRow = RowCount
                    End If
                    tgTable.Row = lRow - 1
                    If lRow <= (tgTable.Height \ tgTable.RowHeight) - tgTable.HeadLines Then
                        KeyCode = 0
                    End If
                End If
            End If
            
            'Special treatment is need for the following case. Don't know why
            If tgTable.Row = 0 And tgTable.FirstRow = 1 Then
                tgTable.FirstRow = 0
                tgTable.Row = 0
            End If
        End If
    
    ElseIf (KeyCode = vbKeyHome Or KeyCode = vbKeyEnd) And (Not bCtrlDown) Then
        'david 12/11/2002
        'new feature, home goto first editable column, end goto last editable column
        If Not tgTable.EditActive Then
            If KeyCode = vbKeyHome Then
                KeyCode = 0
                bCellEditing = False
                tgTable.Col = FirstEditableCol
            Else
                tgTable.Col = LastEditableCol
            End If
        End If
        '''''''''''''''''
    End If
    
    'new feature, ctrl-home, ctrl-end
    If bShiftDown And bCtrlDown And bAltDown Then
        'do nothing
    ElseIf bShiftDown And bAltDown Then
        'do nothing
    ElseIf bShiftDown And bCtrlDown Then
        'do nothing
    ElseIf bCtrlDown And bAltDown Then
        'do nothing
    ElseIf bShiftDown Then
        'do nothing
    ElseIf bCtrlDown And (Not bShiftDown) Then
        If KeyCode = vbKeyHome Or KeyCode = vbKeyEnd Then
            If Not tgTable.EditActive Then
                If KeyCode = vbKeyHome Then
                    lRow = 0
                    tgTable.Bookmark = MakeBookmark(lRow)
                    KeyCode = 0
                    bCellEditing = False
                    tgTable.Col = FirstEditableCol
                Else
                    lRow = RowCount - 1
                    tgTable.Bookmark = MakeBookmark(lRow)
                    tgTable.Col = LastEditableCol
                End If

                nKeyProcessing = KEY_P_CHANGE_ROW
            End If
        End If
    ElseIf bAltDown Then
        'do nothing
    ElseIf bShiftDown = 0 Then  'no special key is pressed
        'do nothing
    End If
    
    If nKeyProcessing <> KEY_P_CHANGE_ROW Then
        nKeyProcessing = KEY_P_FINISHED
    End If
    
    Exit Function
    
errKeyDown:
    If Err.Number <> 28 Then
        On Error GoTo 0
        Resume
    End If
    Err.Clear
    nKeyProcessing = KEY_P_FINISHED

End Function

Public Function Keypress(KeyAscii As Integer) As Boolean

'
' Check the key pressed against the regular expression
'
    Dim nCol As Integer
    Dim szText As String
    Dim nCode As Integer
    Dim nLen As Integer
    Dim sTemp As String
    Dim lRow As Long
    
    Keypress = False
    
    If tgTable Is Nothing Then
        Exit Function
    End If
    
    nCol = fnColumnNumber(tgTable.Col)
    If nCol < 0 Then
        Exit Function
    End If

    ' check enter, backspace, escape
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyEscape Then
        Keypress = True
        If bHighlightInvalidCell Then
            If bAllowAddNew Then
                If GetCurrentRowNumber = RowCount - 1 Then
                    KeyAscii = 0
                Else
                    SendKeys "{Right}", True
                End If
                Beep
            Else
                If KeyAscii = vbKeyReturn Then
                    SendKeys "{Right}", True
                End If
            End If
        End If
        Exit Function
    End If

    ' check for cut/copy/paste keys
    If KeyAscii = ASC_CTRL_C Or KeyAscii = ASC_CTRL_V Or KeyAscii = ASC_CTRL_X _
       Or KeyAscii = ASC_CTRL_A Or KeyAscii = ASC_CTRL_I Then
        Keypress = True
        Exit Function
    ElseIf KeyAscii = vbKeyCtrlY Then
        If bCtrlYClearRow Then
            lRow = GetCurrentRowNumber
            If nColumnCount > 0 Then
                If tgTable.EditActive Then
                    tgTable.EditActive = False
                    DoEvents
                End If
                bClearData = True
                tgTable.Col = FirstEditableCol
                DoEvents
                bClearData = False
            End If
            subClearRow lRow
        End If
        Keypress = False
        Exit Function
    End If
        
    fnCheckForValueItem KeyAscii
    If sPatternCode(nCol) = "" Then
        Keypress = True
        If bHighlightInvalidCell Then
            If lEditingRow >= 0 And lEditingRow < nMaxRow Then
                nCellChanged(nCol, lEditingRow) = STATUS_CHANGED
            End If
        End If
        Exit Function
    End If

    ' get the data up to the cursor position and tack on the key pressed
    sTemp = fnCellValue(tgTable.Col, lEditingRow)
    If tgTable.EditActive Then
        szText = Left(sTemp, tgTable.SelStart) & Chr(KeyAscii)
        nLen = Len(sTemp)
        If tgTable.SelStart < nLen Then
            nLen = nLen - (tgTable.SelStart + tgTable.SelLength)
            If nLen > 0 Then
                szText = szText & Right(sTemp, nLen)
            End If
        End If
    Else
        szText = Chr(KeyAscii)
    End If
    #If REGULAR_EXP_INCLUDED Then
        nCode = PRegExpMatch(sPatternCode(nCol), Len(sPatternCode(nCol)), szText)
    #End If
    
    If nCode < 0 Or (nCode <> 0 And nCode <> tgTable.SelStart + 2) Then
        Keypress = False
        Beep
    Else
        Keypress = True
        If bForcedEditing Then
            BeforeColEdit tgTable.Col, 0, False
            bForcedEditing = False
        End If
        If lEditingRow >= 0 And lEditingRow < nMaxRow Then
            nCellChanged(nCol, lEditingRow) = STATUS_CHANGED
            GridArray(nFlagColumn, lEditingRow) = ROW_EDITED
        End If
    End If
    
End Function

Public Sub LostFocus()
    If lEditingRow >= 0 Then
        fnSaveChangedKeys
        fnSaveBookMarks Bookmark(GetCurrentRowNumber)
        lEditingRow = -1
    End If
    bFirstTimeGotFocus = False
End Sub

Public Sub ReadData(ByVal RowBuf As RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
    Dim Bookmk As Variant
    Dim nIndex As Integer
    Dim RelPos As Integer
    Dim lRows As Long
    
    Bookmk = StartLocation

    If bAllowAddNew Then
        lRows = nRowCount + 1
    Else
        lRows = nRowCount
    End If
    
    If ReadPriorRows Then
        ' the grid is requesting data in rows prior to StartLocation
        RelPos = -1
    Else
        ' the grid is requesting data in rows after StartLocation
        RelPos = 1
    End If
    
'    If nRowCount = 0 Then
'        nIndex = nIndex + RelPos
'        If nIndex >= 0 Then
'            RowBuf.bookmark(nRowCount) = MakeBookmark(nRowCount) 'GridArray(nBookMarkColumn, nIndex)
'            RowBuf.RowCount = 1
'        End If
'        Exit Sub
'    End If
    If IsNull(StartLocation) Then
        If ReadPriorRows Then
            nIndex = lRows
        Else
            nIndex = -1
        End If
    Else
        nIndex = Val(StartLocation)
        If nIndex < 0 Or nIndex >= lRows Then
            nIndex = -9999
        End If
    End If
        

    Dim RowsFetched As Integer
    RowsFetched = 0
    Dim i%, j%
    
    ' loop for each row in the RowBuffer object (there will be 10 rows; start at Row 0)
    For i% = 0 To RowBuf.RowCount - 1
    
        nIndex = nIndex + RelPos
        
        If nIndex < 0 Or (nIndex >= lRows) Then
            Exit For
        End If

        For j% = 0 To RowBuf.ColumnCount - 1
        ' Get the value of each cell in the row and give it to the grid
        '    in the RowBuf's columns
        ' Notice that the GetUserData function expects column argument in the
        '    second position, while our array has the column index in the first
        '    position (this is necessary so we can change the number of rows with
        '    the ReDim statement
            If j% >= 0 And j% < nMaxCol Then
                If sDisplayFormat(j%) <> "" Then
                    RowBuf.value(i%, j%) = Format$(GridArray(j%, nIndex), sDisplayFormat(j%))
                Else
                    RowBuf.value(i%, j%) = GridArray(j%, nIndex)
                End If
            Else
                RowBuf.value(i%, j%) = GridArray(j%, nIndex)
            End If
        
        Next j%
        
        ' Tell the grid what the row's bookmark is.  The grid will store the bookmarks for the
        '     rows in addition to the data for the rows.  The bookmarks are used for navigating
        '     through the grid as is done with a bound grid and data control.
        RowBuf.Bookmark(i%) = MakeBookmark(nIndex) 'GridArray(nBookMarkColumn, nIndex)
        ' Increment the count of fetched rows
        ' This is important because you must tell the grid how many rows you have put into the
        '    RowBuf object.
        RowsFetched = RowsFetched + 1
        
    Next i%
    
    ' tell the grid how many rows were fetched
    RowBuf.RowCount = RowsFetched

End Sub

'david 04/06/2004
'add optional argument to do the last row validation check
'instead of the second last row validation check being done currently
Public Function RowColChange(vLastRow As Variant, ByVal LastCol As Integer, _
                             Optional bCheckLastRow As Boolean = False) As Boolean
    
    On Error GoTo errRowChanging
    
    RowColChange = False
    nAddRowCount = 0
    If tgTable Is Nothing Or bLocked Or frmParent.ActiveControl.TabIndex <> tgTable.TabIndex Then
        Exit Function
    End If
    If Not AddingValidate Then
        If bAddingNewRow Then
            bAddingNewRow = False
            'Comment out the following row for problems happened when adding new row
            nLastColumn = tgTable.Col
            'david 08/22/2006  #529111 NDR
            'add - 1 so that the correct value of LastRow() will be returned
            If bCheckLastRow And GetCurrentRowNumber() >= RowCount() - 1 Then
                nLastTableRow = tgTable.Row - 1
            Else
                nLastTableRow = tgTable.Row
            End If
            ''''''''''''''''''''''''''''''
            fnShowStatus GetCurrentRowNumber, tgTable.Col
            Exit Function
        End If
    End If
    
    If tgTable.Row < 0 Then
        tgTable.Row = nLastTableRow
        Exit Function
    End If
    
    Dim nCurrRow As Long
    Dim nCol As Integer
    Dim bFlag As Boolean
    
    nCurrRow = GetCurrentRowNumber()
    nCol = fnColumnNumber(tgTable.Col)
    
    bFlag = True
    
    If bClearData Or (nKeyProcessing = KEY_P_CHANGE_ROW And bAllowAddNew) Or (nCurrRow < LastRow And bFlag) Then
        fnChangeToEditingCol nLastColumn
        fnShowStatus nCurrRow, tgTable.Col
        nLastColumn = tgTable.Col
        nLastTableRow = tgTable.Row
        RowColChange = True
        Exit Function
    End If
    
    Dim nCurrCol As Integer
    
    If nCurrRow < 0 Then
        RowColChange = True
        Exit Function
    End If
    
    fnValidate tgTable.Col, nCurrRow, fnGetUserData(MakeBookmark(nCurrRow), tgTable.Col)
    fnShowStatus nCurrRow, tgTable.Col
    
    If bRowChanging Or nLastTableRow < 0 Or LastRow > nRowCount Then
        nLastColumn = tgTable.Col
        nLastTableRow = tgTable.Row
        RowColChange = True
        Exit Function
    End If
    If LastRow < tgTable.Row Then
        'Row changed
        'david 04/06/2004
        'add optional argument to do the last row validation check
        'instead of the second last row validation check being done currently
        If GetCurrentRowNumber >= (RowCount - IIf(bCheckLastRow, 0, 1)) And bAllowAddNew Then
        '''''''''''''''''
            'The last row
            Dim i As Integer
            For i = nColumnCount - 1 To 0 Step -1
                If Not fnCheckRowValid(i, LastRow) Then
                    tgTable.Row = nLastTableRow
                    Exit Function
                End If
            Next
        End If
    End If
    
    bRowChanging = True
    
    If GetCurrentRowNumber <> lEditingRow Then
        If lEditingRow >= 0 Then
            fnSaveChangedKeys
            fnSaveBookMarks Bookmark(lEditingRow)
            lEditingRow = -1
        End If
    End If
    
    If fnEscapable(nLastColumn, tgTable.Col) Or IsNull(GridArray(nLastColumn, LastRow)) And LastRow = nRowCount - 1 Then
        bRowChanging = False
        nLastColumn = tgTable.Col     'tgTable.col
        nLastTableRow = tgTable.Row        'tgTable.Row
        RowColChange = True
        Exit Function
    End If
    If nReturnKeyAction = RETURN_KEY_DOWN Then
        If ((nLastTableRow <> tgTable.Row) Or nColumnCount = 1) Then
            fnCheckForError LastRow
        End If
    Else
        If ((nLastColumn <> tgTable.Col) Or nColumnCount = 1) Then
            fnCheckForError LastRow
        End If
    End If
    If bErrorFlag Then
        If nCol >= 0 Then
            If (nColumnType(nCol) <> BACK_ALWAYS) Then
                bErrorFlag = False
            End If
        End If
        On Error Resume Next
        tgTable.Col = nLastColumn
        tgTable.Row = nLastTableRow

        If bHighlightInvalidCell Then
            bForcedEditing = True
            tgTable.EditActive = True
            tgTable.SelStart = 0
            If nColumnCount > 0 Then
                nCellChanged(nCol, lEditingRow) = nCellChanged(nCol, lEditingRow) And (Not STATUS_VALIDATED) Or STATUS_CHANGED
            End If
            If Not IsNull(tgTable.Columns(nLastColumn).Text) Then
                tgTable.SelLength = Len(tgTable.Columns(nLastColumn).Text)
            End If
        End If

        On Error GoTo 0
        If nKeyProcessing = KEY_P_FINISHED Then
            DoEvents
        End If
        fnShowStatus LastRow, tgTable.Col
        'david 12/06/00
        'if red error message then Beep
        If lblStatusbar.ForeColor = &HFF Then
            Beep
        End If
        bRowChanging = False
        RowColChange = False
        Exit Function
    End If

    nCurrCol = tgTable.Col
    If nCol < 0 Then
        fnChangeToEditingCol nLastColumn
    Else
        nLastColumn = nCurrCol     'tgTable.col
        'david 08/22/2006  #529111 NDR
        'add - 1 so that the correct value of LastRow() will be returned
        If bCheckLastRow And GetCurrentRowNumber() >= RowCount() - 1 Then
            nLastTableRow = tgTable.Row - 1
        Else
            nLastTableRow = tgTable.Row
        End If
        ''''''''''''''''''''''''''''''
    End If
    
    nCol = fnColumnNumber(nLastColumn)
    If nCol >= 0 And nCol < nColumnCount Then
        If tgTable.FirstRow = "" Then
            nCurrRow = tgTable.Row
            If nCurrRow <= nRowCount Then
                fnFetchDefault nCol, nCurrRow
            End If
        Else
            nCurrRow = tgTable.FirstRow + tgTable.Row
            fnFetchDefault nCol, nCurrRow
        End If
    End If

    bRowChanging = False
    RowColChange = True
    Exit Function

errRowChanging:
    Err.Clear
    bRowChanging = False
End Function

Private Function fnRound(vLastRow As Variant) As Long
    If IsNull(vLastRow) Then
        fnRound = 0
    Else
        fnRound = Val(vLastRow)
    End If
End Function

Public Sub SetPattern(ByVal nIndex As Integer, Optional vPtn As Variant, Optional bRequired As Variant)
    Dim nCol As Integer
    
    nCol = fnColumnNumber(nIndex)
    If nCol >= nColumnCount Then
        Exit Sub
    End If
    If nCol < 0 Then
        Exit Sub
    End If
    If IsMissing(vPtn) Then
        sPatternCode(nCol) = ""
    Else
        Dim sPtn As String
        sPtn = vPtn
        sPatternCode(nCol) = fnGetPatternCode(sPtn)
    End If

End Sub

Public Sub SetPrompt(ByVal nIndex As Integer, Optional sMsg As Variant)
    If nIndex >= tgTable.Columns.Count Then
        Exit Sub
    End If
    Dim nCol As Integer
    nCol = fnColumnNumber(nIndex)
    If nCol < 0 Then
        Exit Sub
    End If
    If IsMissing(sMsg) Then
        sPrompt(nCol) = ""
    Else
        sPrompt(nCol) = sMsg
    End If
End Sub

Property Set StatusBar(lblBar As Object)
    Set lblStatusbar = lblBar
End Property

Public Sub WriteData(ByVal RowBuf As RowBuffer, WriteLocation As Variant)

End Sub

Private Function fnGetNewBookmark() As Variant
' used to add new rows to the data source
' Creates a bookmark for a newly added row, called from the UnboundAddData event
' ReDim the array and create the new bookmark using the MakeBookmark function

'    fnGetNewBookmark = MakeBookmark(nRowCount)
    Dim nCol As Integer
    fnGetNewBookmark = MakeBookmark(lBookMarkCount)
    lBookMarkCount = lBookMarkCount + 1
    If nRowCount > nMaxRow Then
        nMaxRow = nMaxRow + ROW_INCREMENT
        ReDim Preserve GridArray(nMaxCol + 1, nMaxRow + 1)
        If nColumnCount > 0 Then
            nCol = UBound(nCellChanged)
            ReDim Preserve nCellChanged(nCol, nMaxRow + 1)
            ReDim Preserve bValidCell(nCol, nMaxRow + 1)
        End If
    End If
End Function

Property Set Table(tblGrid As Object)
'
' Add a True DBGrid to the collection and set its default properties
'
    Dim m_styRow As Object            ' row style
    Dim m_styCurrentCell As Object    ' current cell style
    
    Set tgTable = tblGrid
    nOriginalTableWidth = tgTable.Width
    nOriginalTableHeight = tgTable.Height
    nFontSize = tgTable.Font.Size

    ColumnCount = tblGrid.Columns.Count
    
    On Error Resume Next
    tgTable.Styles.Remove szROW_STYLE
    tgTable.Styles.Remove szCURRENT_CELL_STYLE
    On Error GoTo 0

    ' set some properties for the table
    Set m_styRow = tgTable.Styles.Add(szROW_STYLE)
    Set m_styCurrentCell = tgTable.Styles.Add(szCURRENT_CELL_STYLE)
End Property

Private Function fnGetUserData(Bookmk As Variant, Colm As Integer) As Variant
' Returns a data value from the data source (our array)
' This function will be called from the UnboundReadData event for each cell
'    of data (or each column in a row)
' Bookmk - bookmark of the row from which we are requesting data
' Colm - column number for the field from which we are requesting data

' Notice that the GetUserData function expects column argument in the
'    second position, while our array has the column index in the first
'    position (this is necessary so we can change the number of rows with
'    the ReDim statement

    Dim Index As Long
    ' need the row index to get the value out of the array, column index is passed in as Colm
    Index = fnIndexFromBookmark(Bookmk, False)  ' the second argument is not needed here
    
    If Index < 0 Or Index > nRowCount Or Colm < 0 Or Colm >= nMaxCol Then 'make sure the index is valid
        fnGetUserData = Null   'return nothing for a bogus request for data (e.g., EOF, BOF)
    Else
        If Colm >= 0 And Colm < nMaxCol Then
            If sDisplayFormat(Colm) <> "" Then
                fnGetUserData = Format$(GridArray(Colm, Index), sDisplayFormat(Colm))
            Else
                fnGetUserData = GridArray(Colm, Index) 'return the actual data
            End If
        Else
            fnGetUserData = GridArray(Colm, Index) 'return the actual data
        End If
    End If
End Function

Private Function fnIndexFromBookmark(Bookmk As Variant, ReadPriorRows As Boolean) As Long

    If IsNull(Bookmk) Then
    ' A null bookmark signifies EOF or BOF (which one depends on whether we are trying to
    '    Read Rows Prior to Bookmk or rows after Bookmk)
    
        If ReadPriorRows = True Then  ' the grid is asking for rows prior to Bookmk
        ' we want EOF since you can not read rows prior to BOF
        ' (nRowCount - 1) is the last element of the array, so nRowCount is the index for our EOF marker
            fnIndexFromBookmark = nRowCount
        Else  ' the grid is asking for rows after Bookmk
        ' we want BOF since you can not read rows after EOF
        ' the first element of the array is 0, so -1 is the index for our BOF marker
            fnIndexFromBookmark = -1
        End If
    Else
    ' the row is a valid row, so find the index number using the bookmark
        Dim Index As Long
        ' take the Val() to get the array index
        ' this is the opposite of the function in MakeBookmark()
        Index = -9999
        fnIndexFromBookmark = Val(Bookmk)
    End If
End Function

Public Function MakeBookmark(ByVal Index As Long) As Variant
    'Generate a bookmark
    MakeBookmark = Str$(Index)
End Function

Public Sub AddData(ByVal RowBuf As RowBuffer, NewRowBookmark As Variant)

End Sub

Private Sub Class_Initialize()
    
    nMaxChanged = -1
    nMaxRow = 0
    nRowCount = 0
    nESCCount = 0
    nInputKey = -1
    m_nRowHighLighted = STYLE_INITIAL
    nColumnForNewRow = -1
    HighlightInvalidCell = False
    bCtrlYClearRow = False
    nReturnKeyAction = RETURN_KEY_RIGHT
    bFirstTimeGotFocus = True
    m_bKeepData = False
    m_bAllowRevalidate = False
    lMaxRows = -1
    sDownArrowMessage = "No new rows can be added unless previous input is valid"

    m_bSortOnCol = False
    m_bSortWhenGridNotvalid = False
    m_bReverseSortOnCol = False

    'david 03/23/2006  #513569
    'david 06/28/2006  #523455
    'default to not showing error message
    m_bIgnoreCellValueError = True
    ''''''''''''''''''''''''''

End Sub

Private Sub Class_Terminate()
    Set tgTable = Nothing
    Set lblStatusbar = Nothing
    Set frmParent = Nothing
    Set engDBEngine = Nothing
End Sub

'==============================================================================================
'david 01/15/2001
'properties and functions for Column Sort enhancement
Property Let SortByColumn(bStatus As Boolean)
    m_bSortOnCol = bStatus
End Property

Property Get SortByColumn() As Boolean
    SortByColumn = m_bSortOnCol
End Property

Property Let SortColEnabled(nColIndex As Integer, bStatus As Boolean)
    Dim j As Long
    
    If Not m_bSortOnCol Then
        Exit Property
    End If
    
    If nColIndex < 0 Then
        Exit Property
    End If
    
    'check the exists of sort column
    For j = 0 To m_SortColCount - 1
        If m_arySortKeysCount(nCOL_IN_GRID, j) = nColIndex Then
            m_arySortKeysCount(nCOL_SORT_ENABLED, j) = bStatus
            Exit Property
        End If
    Next j
    
    MsgBox "Column Index " & nColIndex & " not found.", vbExclamation
End Property

Property Get SortColEnabled(nColIndex As Integer) As Boolean
    Dim j As Long
    
    If Not m_bSortOnCol Then
        Exit Property
    End If
    
    If nColIndex < 0 Then
        Exit Property
    End If
    
    'check the exists of sort column
    For j = 0 To m_SortColCount - 1
        If m_arySortKeysCount(nCOL_IN_GRID, j) = nColIndex Then
            SortColEnabled = m_arySortKeysCount(nCOL_SORT_ENABLED, j)
            Exit Property
        End If
    Next j
    
    MsgBox "Column Index " & nColIndex & " not found.", vbExclamation
End Property

Property Let SortIgnoreGridValidation(bStatus As Boolean)
    m_bSortWhenGridNotvalid = bStatus
End Property

Property Get SortIgnoreGridValidation() As Boolean
    SortIgnoreGridValidation = m_bSortWhenGridNotvalid
End Property

Property Get NUMERIC_TYPE() As Integer
    NUMERIC_TYPE = m_NUMERIC_TYPE
End Property

Property Get STRING_TYPE() As Integer
    STRING_TYPE = m_STRING_TYPE
End Property

Property Get DATE_TYPE() As Integer
    DATE_TYPE = m_DATE_TYPE
End Property

Property Get ASCENDING() As Integer
    ASCENDING = m_ASCENDING
End Property

Property Get DESCENDING() As Integer
    DESCENDING = m_DESCENDING
End Property

Property Get CASE_SENSITIVE() As Integer
    CASE_SENSITIVE = m_CASE_SENSITIVE
End Property

Property Get CASE_INSENSITIVE() As Integer
    CASE_INSENSITIVE = m_CASE_INSENSITIVE
End Property

'david 08/09/2002
'reverse sort column on subsequence HeadClick()
Property Let ReverseSortByColumn(bStatus As Boolean)
    m_bReverseSortOnCol = bStatus
End Property
'''''''''''''''''

'calling syntax
'AddSortColumn
'nGridColIndex, nCol, <NUMERIC_TYPE|STRING_TYPE>, <ASCENDING|DESCEDING>, <CASE_SENSITIVE|CASE_INSENSITIVE>
'[, nCol, <NUMERIC_TYPE|STRING_TYPE>, <ASCENDING|DESCEDING>, <CASE_SENSITIVE|CASE_INSENSITIVE>]
'[...]
'Note: if Sort Column already exists, it will be replaced.
'
Public Sub AddSortColumn(nColIndex As Integer, ParamArray argSortOnKeyList())
    Dim nMaxArg As Integer
    Dim i As Integer
    Dim j As Integer
    Dim nSortKeyCount As Integer
    
    On Error GoTo errTrap
    nMaxArg = UBound(argSortOnKeyList) + 1
    
    If nMaxArg Mod 4 <> 0 Then
        MsgBox "Invalid Number of Arguments.", vbExclamation
        Exit Sub
    End If
    
    If nMaxArg / 4 > MAX_SORT_KEYS Then
        MsgBox "Max. keys allowed is " & MAX_SORT_KEYS & ".", vbExclamation
        Exit Sub
    End If
    
    'check the exists of sort column
    For j = 0 To m_SortColCount - 1
        If m_arySortKeysCount(nCOL_IN_GRID, j) = nColIndex Then
            'MsgBox "Sort Column already exists, and will be replaced.", vbExclamation
            Exit For
        End If
    Next j
    
    'new sort column
    If j > m_SortColCount - 1 Then
        m_SortColCount = m_SortColCount + 1
        j = m_SortColCount - 1
        ReDim Preserve m_arySortCol(MAX_SORT_KEYS - 1, j)
        ReDim Preserve m_arySortKeysCount(MAX_COL_arySortKeysCount, j)
        m_arySortKeysCount(nCOL_IN_GRID, j) = nColIndex
        m_arySortKeysCount(nCOL_SORT_ENABLED, j) = True
    End If
    
    nSortKeyCount = -1
    
    For i = 0 To nMaxArg - 1
        Select Case i Mod 4
        Case 0  'nCOL
            
            nSortKeyCount = nSortKeyCount + 1
            m_arySortCol(nSortKeyCount, j).nColIndex = argSortOnKeyList(i)
        
        Case 1  '[NUMERIC_TYPE|STRING_TYPE]
            
            If argSortOnKeyList(i) <> m_NUMERIC_TYPE And argSortOnKeyList(i) <> m_STRING_TYPE _
               And argSortOnKeyList(i) <> m_DATE_TYPE Then
                MsgBox "Invalid Column Type [NUMERIC_TYPE|STRING_TYPE|DATE_TYPE].", vbExclamation
                Exit Sub
            End If
            
            m_arySortCol(nSortKeyCount, j).nColType = argSortOnKeyList(i)
        
        Case 2  '[ASCENDING|DESCEDING]
            
            If argSortOnKeyList(i) <> m_ASCENDING And argSortOnKeyList(i) <> m_DESCENDING Then
                MsgBox "Invalid Sort Order [ASCENDING|DESCEDING].", vbExclamation
                Exit Sub
            End If
        
            m_arySortCol(nSortKeyCount, j).nSortOrder = argSortOnKeyList(i)
            'david 08/11/2002
            'save the original sort order
            m_arySortCol(nSortKeyCount, j).nInitSortOrder = m_arySortCol(nSortKeyCount, j).nSortOrder
            '''''''''''''''''
        
        Case 3  '[CASE_SENSITIVE|CASE_INSENSITIVE]
            
            If argSortOnKeyList(i) <> m_CASE_SENSITIVE And argSortOnKeyList(i) <> m_CASE_INSENSITIVE Then
                MsgBox "Invalid Sort Order [CASE_SENSITIVE|CASE_INSENSITIVE].", vbExclamation
                Exit Sub
            End If
        
            m_arySortCol(nSortKeyCount, j).nCaseSensitive = argSortOnKeyList(i)
        
        End Select
    Next i
    
    m_arySortKeysCount(nCOL_KEYS_COUNT, j) = nSortKeyCount
    
    Exit Sub
    
errTrap:
    MsgBox "Error has occurred while setting up Sort Column." + vbCrLf + vbCrLf + "Error Code: " & _
        Err.Number & vbCrLf & "Error Desc: " & Err.Description, vbExclamation
End Sub

'david 05/26/2002
'NOTE:
'(FillWithArray vData, True, False) IS NOT WORKING in SortCol function
'if a cell validation in the main form uses the value from other cells,
'the validation will fail because those values are not available.
'THE REVALIDATION OF THE GRID SHOULD BE PERFORMED IN THE CALLING FUNCTION
'AFTER THE COLUMNS HAS BEEN SORT

'remove optional argument... SORRY FOR THE INCONVENIENCE (-:
'return True if the sorting operation is actually done.

'Public Sub HeadClick(ColIndex As Integer, Optional bReValidate As Boolean = False)
Public Function HeadClick(ColIndex As Integer) As Boolean
    If Not m_bSortOnCol Then
        Exit Function
    End If
    
    'david 05/26/2002
    'NOTE:
    '(FillWithArray vData, True, False) IS NOT WORKING in SortCol function
    'if a cell validation in the main form uses the value from other cells,
    'the validation will fail because those values are not available.
    'THE REVALIDATION OF THE GRID NEEDS TO BE DONE AFTER THE COLUMNS HAS BEEN SORT
    
    'SortCol ColIndex, bReValidate
    HeadClick = SortCol(ColIndex)
End Function

'david 05/26/2002
'NOTE:
'(FillWithArray vData, True, False) IS NOT WORKING
'if a cell validation in the main form uses the value from other cells,
'the validation will fail because those values are not available.
'THE REVALIDATION OF THE GRID SHOULD BE PERFORMED IN THE CALLING FUNCTION
'AFTER THE COLUMNS HAS BEEN SORT

'remove optional argument
'return True if the sorting operation is actually done.

'Public Sub SortCol(nSortCol As Integer, Optional bReValidate As Boolean = False)
Public Function SortCol(nSortCol As Integer) As Boolean
    Dim vData()
    Dim lDummyRow As Long
    Dim nIndex As Integer
    
    Dim i As Integer
    Dim j As Long
    Dim sSortCols As String
    Dim nCol As Integer
    
    On Error GoTo errTrap
    
    m_bColIsSorted = False
    
    If Not IsColForSorting(nSortCol, nIndex) Then
        Exit Function
    End If
    
    For j = 0 To m_arySortKeysCount(nCOL_KEYS_COUNT, nIndex)
        sSortCols = sSortCols + "Col " & m_arySortCol(j, nIndex).nColIndex & " "
        If m_arySortCol(j, nIndex).nSortOrder = ASCENDING Then
            sSortCols = sSortCols + " ASCENDING, "
        Else  'descending
            sSortCols = sSortCols + " DESCENDING, "
        End If
    Next j
    
    sSortCols = Mid(sSortCols, 1, Len(sSortCols) - 2)
    lblStatusbar.ForeColor = STANDARD_TEXT_COLOR
    lblStatusbar.Font.Bold = False
    lblStatusbar = "Sorting " + sSortCols + " ..."
    lblStatusbar.Refresh
'Debug.Print sSortCols
    Screen.MousePointer = vbHourglass
    
    'read data
    GetRecords vData, lDummyRow, szEMPTY

    'begin shell sort
    Dim NumEls As Long
    Dim Index As Long
    Dim Index2 As Long
    Dim FirstItem As Long
    Dim Distance As Long
    Dim value()

    NumEls = UBound(vData, 2) + 1
    FirstItem = LBound(vData, 2)
    nCol = UBound(vData, 1)

    'Find the best value for distance.
    Do
        Distance = Distance * 3 + 1
    Loop Until Distance > NumEls

    ReDim value(nCol)

    'Sort the array.
    Do
        Distance = Distance \ 3

        For Index = Distance + FirstItem To NumEls + FirstItem - 1
            'Value = vData(Index)
            'array assignment
            For i = 0 To nCol
                value(i) = vData(i, Index)
            Next
            
            Index2 = Index

            'Do While (vData(Index2 - Distance) > Value)
            Do While (fnCompareSortKeysNeedToSwap2(nIndex, vData, Index2 - Distance, value))
                'vData(Index2) = vData(Index2 - Distance)
                'array assignment
                For i = 0 To nCol
                    vData(i, Index2) = vData(i, Index2 - Distance)
                Next
                
                m_bColIsSorted = True

                Index2 = Index2 - Distance

                If Index2 - Distance < FirstItem Then
                    Exit Do
                End If
            Loop

            'vData(Index2) = Value
            'array assignment
            For i = 0 To nCol
                vData(i, Index2) = value(i)
            Next
        Next Index
    Loop Until Distance = 1
    'end shell sort
    
    If m_bColIsSorted Then
        'david 05/10/2001
        'NOTE:
        '(FillWithArray vData, True, False) IS NOT WORKING
        'if a cell validation in the main form uses the value from other cells,
        'the validation will fail because those values are not available.
        'THE REVALIDATION OF THE GRID NEEDS TO BE DONE AFTER THE COLUMNS HAS BEEN SORT
        FillWithArray vData, True, , False  ', Not bReValidate
        SortCol = True
    End If
    
    'reversed sort column on subsequence HeadClick()
    If m_bReverseSortOnCol Then
        'reverse the "order by" variable of the first sorting column
        'reverse the order
        If m_arySortCol(0, nIndex).nSortOrder = m_ASCENDING Then
            m_arySortCol(0, nIndex).nSortOrder = m_DESCENDING
        Else
            m_arySortCol(0, nIndex).nSortOrder = m_ASCENDING
        End If
    End If
    
errTrap:
    
    Screen.MousePointer = vbDefault
End Function

Private Function fnCompareSortKeysNeedToSwap(nIndex, vData(), j As Long, k As Long) As Boolean
    Dim n As Integer
    Dim dfData1 As Double
    Dim dfData2 As Double
    Dim sData1 As String
    Dim sData2 As String
    Dim sKey1 As String
    Dim sKey2 As String
    Dim nSortCount As Integer
    
    On Error GoTo errTrap
    
    nSortCount = -1
    
    For n = 0 To m_arySortKeysCount(nCOL_KEYS_COUNT, nIndex)
        Select Case m_arySortCol(n, nIndex).nColType
        Case m_NUMERIC_TYPE
            If IsNull(vData(m_arySortCol(n, nIndex).nColIndex, j)) Then
                dfData1 = 0
            Else
                dfData1 = Val(vData(m_arySortCol(n, nIndex).nColIndex, j))
            End If
        
            If IsNull(vData(m_arySortCol(n, nIndex).nColIndex, k)) Then
                dfData2 = 0
            Else
                dfData2 = Val(vData(m_arySortCol(n, nIndex).nColIndex, k))
            End If
            
            If m_arySortCol(n, nIndex).nSortOrder = m_ASCENDING Then
                If dfData1 > dfData2 Then
                    fnCompareSortKeysNeedToSwap = True
                    Exit Function
                ElseIf dfData1 < dfData2 Then
                    'no swap
                    Exit Function
                End If
            Else  'descending
                If dfData1 < dfData2 Then
                    fnCompareSortKeysNeedToSwap = True
                    Exit Function
                ElseIf dfData1 > dfData2 Then
                    'no swap
                    Exit Function
                End If
            End If
        Case m_STRING_TYPE
            If IsNull(vData(m_arySortCol(n, nIndex).nColIndex, j)) Then
                sData1 = ""
            Else
                If m_arySortCol(n, nIndex).nCaseSensitive = m_CASE_SENSITIVE Then
                    sData1 = vData(m_arySortCol(n, nIndex).nColIndex, j)
                Else
                    sData1 = UCase(vData(m_arySortCol(n, nIndex).nColIndex, j))
                End If
            End If
        
            If IsNull(vData(m_arySortCol(n, nIndex).nColIndex, k)) Then
                sData2 = ""
            Else
                If m_arySortCol(n, nIndex).nCaseSensitive = m_CASE_SENSITIVE Then
                    sData2 = vData(m_arySortCol(n, nIndex).nColIndex, k)
                Else
                    sData2 = UCase(vData(m_arySortCol(n, nIndex).nColIndex, k))
                End If
            End If
        
            If m_arySortCol(n, nIndex).nSortOrder = m_ASCENDING Then
                If sData1 > sData2 Then
                    fnCompareSortKeysNeedToSwap = True
                    Exit Function
                ElseIf sData1 < sData2 Then
                    'no swap
                    Exit Function
                End If
            Else  'descending
                If sData1 < sData2 Then
                    fnCompareSortKeysNeedToSwap = True
                    Exit Function
                ElseIf sData1 > sData2 Then
                    'no swap
                    Exit Function
                End If
            End If
        Case m_DATE_TYPE
            If IsNull(vData(m_arySortCol(n, nIndex).nColIndex, j)) Then
                sData1 = ""
            Else
                sData1 = tfnDateString(vData(m_arySortCol(n, nIndex).nColIndex, j))
                
                If Not IsDate(sData1) Then
                    sData1 = ""
                End If
            End If
        
            If IsNull(vData(m_arySortCol(n, nIndex).nColIndex, k)) Then
                sData2 = ""
            Else
                sData2 = tfnDateString(vData(m_arySortCol(n, nIndex).nColIndex, k))
                
                If Not IsDate(sData2) Then
                    sData2 = ""
                End If
            End If
        
            If sData1 <> "" And sData2 <> "" Then
                If m_arySortCol(n, nIndex).nSortOrder = m_ASCENDING Then
                    If CDate(sData1) > CDate(sData2) Then
                        fnCompareSortKeysNeedToSwap = True
                        Exit Function
                    ElseIf sData1 < sData2 Then
                        'no swap
                        Exit Function
                    End If
                Else  'descending
                    If CDate(sData1) < CDate(sData2) Then
                        fnCompareSortKeysNeedToSwap = True
                        Exit Function
                    ElseIf sData1 > sData2 Then
                        'no swap
                        Exit Function
                    End If
                End If
            Else  'sData1 = "" Or sData2 = ""
                If m_arySortCol(n, nIndex).nSortOrder = m_ASCENDING Then
                    If sData1 <> "" Then
                        fnCompareSortKeysNeedToSwap = True
                        Exit Function
                    Else
                        'no swap
                        Exit Function
                    End If
                Else  'descending
                    If sData2 <> "" Then
                        fnCompareSortKeysNeedToSwap = True
                        Exit Function
                    Else
                        'no swap
                        Exit Function
                    End If
                End If
            End If
        End Select
    Next n

    Exit Function
    
errTrap:
    'error, no sort
End Function

Private Function fnCompareSortKeysNeedToSwap2(nIndex, vData(), j As Long, vValue()) As Boolean
    Dim n As Integer
    Dim dfData1 As Double
    Dim dfData2 As Double
    Dim sData1 As String
    Dim sData2 As String
    Dim sKey1 As String
    Dim sKey2 As String
    Dim nSortCount As Integer
    
    On Error GoTo errTrap
    
    nSortCount = -1
    
    For n = 0 To m_arySortKeysCount(nCOL_KEYS_COUNT, nIndex)
        Select Case m_arySortCol(n, nIndex).nColType
        Case m_NUMERIC_TYPE
            If Trim(vData(m_arySortCol(n, nIndex).nColIndex, j) & "") = "" Then
                dfData1 = 0
            Else
                dfData1 = CDbl(vData(m_arySortCol(n, nIndex).nColIndex, j))
            End If
        
            If Trim(vValue(m_arySortCol(n, nIndex).nColIndex) & "") = "" Then
                dfData2 = 0
            Else
                dfData2 = CDbl(vValue(m_arySortCol(n, nIndex).nColIndex))
            End If
            
            If m_arySortCol(n, nIndex).nSortOrder = m_ASCENDING Then
                If dfData1 > dfData2 Then
                    fnCompareSortKeysNeedToSwap2 = True
                    Exit Function
                ElseIf dfData1 < dfData2 Then
                    'no swap
                    Exit Function
                End If
            Else  'descending
                If dfData1 < dfData2 Then
                    fnCompareSortKeysNeedToSwap2 = True
                    Exit Function
                ElseIf dfData1 > dfData2 Then
                    'no swap
                    Exit Function
                End If
            End If
        Case m_STRING_TYPE
            If IsNull(vData(m_arySortCol(n, nIndex).nColIndex, j)) Then
                sData1 = ""
            Else
                If m_arySortCol(n, nIndex).nCaseSensitive = m_CASE_SENSITIVE Then
                    sData1 = vData(m_arySortCol(n, nIndex).nColIndex, j)
                Else
                    sData1 = UCase(vData(m_arySortCol(n, nIndex).nColIndex, j))
                End If
            End If
        
            If IsNull(vValue(m_arySortCol(n, nIndex).nColIndex)) Then
                sData2 = ""
            Else
                If m_arySortCol(n, nIndex).nCaseSensitive = m_CASE_SENSITIVE Then
                    sData2 = vValue(m_arySortCol(n, nIndex).nColIndex)
                Else
                    sData2 = UCase(vValue(m_arySortCol(n, nIndex).nColIndex))
                End If
            End If
        
            If m_arySortCol(n, nIndex).nSortOrder = m_ASCENDING Then
                If sData1 > sData2 Then
                    fnCompareSortKeysNeedToSwap2 = True
                    Exit Function
                ElseIf sData1 < sData2 Then
                    'no swap
                    Exit Function
                End If
            Else  'descending
                If sData1 < sData2 Then
                    fnCompareSortKeysNeedToSwap2 = True
                    Exit Function
                ElseIf sData1 > sData2 Then
                    'no swap
                    Exit Function
                End If
            End If
        Case m_DATE_TYPE
            If IsNull(vData(m_arySortCol(n, nIndex).nColIndex, j)) Then
                sData1 = ""
            Else
                sData1 = tfnDateString(vData(m_arySortCol(n, nIndex).nColIndex, j))
                
                If Not IsDate(sData1) Then
                    sData1 = ""
                End If
            End If
        
            If IsNull(vValue(m_arySortCol(n, nIndex).nColIndex)) Then
                sData2 = ""
            Else
                sData2 = tfnDateString(vValue(m_arySortCol(n, nIndex).nColIndex))
                
                If Not IsDate(sData2) Then
                    sData2 = ""
                End If
            End If
        
            If sData1 <> "" And sData2 <> "" Then
                If m_arySortCol(n, nIndex).nSortOrder = m_ASCENDING Then
                    If CDate(sData1) > CDate(sData2) Then
                        fnCompareSortKeysNeedToSwap2 = True
                        Exit Function
                    ElseIf sData1 < sData2 Then
                        'no swap
                        Exit Function
                    End If
                Else  'descending
                    If CDate(sData1) < CDate(sData2) Then
                        fnCompareSortKeysNeedToSwap2 = True
                        Exit Function
                    ElseIf sData1 > sData2 Then
                        'no swap
                        Exit Function
                    End If
                End If
            Else  'sData1 = "" Or sData2 = ""
                If m_arySortCol(n, nIndex).nSortOrder = m_ASCENDING Then
                    If sData1 <> "" Then
                        fnCompareSortKeysNeedToSwap2 = True
                        Exit Function
                    Else
                        'no swap
                        Exit Function
                    End If
                Else  'descending
                    If sData2 <> "" Then
                        fnCompareSortKeysNeedToSwap2 = True
                        Exit Function
                    Else
                        'no swap
                        Exit Function
                    End If
                End If
            End If
        End Select
    Next n

    Exit Function
    
errTrap:
    'error, no sort
End Function

Public Function GetErrorMessage(lRow As Long, nCol As Integer) As String
    Dim nECol As Integer
    
    nECol = fnColumnNumber(nCol)
    
    If nECol < 0 Or lRow < 0 Then
        Exit Function
    End If

    If Not bValidCell(nECol, lRow) Then
        If (nCellChanged(nECol, lRow) And STATUS_WARNING) = STATUS_WARNING Then
            If sErrMessage(nECol) = "" Then
                GetErrorMessage = "Data is invalid; " & sPrompt(nECol)
            Else
                GetErrorMessage = sErrMessage(nECol)
            End If
        End If
    End If
End Function

Private Function IsColForSorting(nSortCol As Integer, Optional Index As Integer) As Boolean
    Dim j As Long
    
    Index = -1
    
    If Not m_bSortOnCol Then
        Exit Function
    End If
    
    If nSortCol < 0 Then
        Exit Function
    End If
    
    If RowCount <= 1 Then
        Exit Function
    End If
    
    If Not m_bSortWhenGridNotvalid Then
        If Not ValidData Then
            Exit Function
        End If
    End If
    
    'check the exists of sort column
    For j = 0 To m_SortColCount - 1
        If m_arySortKeysCount(nCOL_IN_GRID, j) = nSortCol Then
            If m_arySortKeysCount(nCOL_SORT_ENABLED, j) Then
                Index = j
                IsColForSorting = True
            End If
                
            Exit Function
        End If
    Next j
End Function

'david 04/27/2002
'THIS CALL IS REQUIRED FOR THE IMPLEMENTATION OF:
'WHEN CLICKED ON THE COLUMN HEADER TO SORT THE GRID
'TO SET/RESTORE THE HIGHLIGHTED ROW PROPERLY

'SHOULD BE CALLED ALSO IF THE SELECTOR (ROW HIGHLIGHTING) FEATURE IS USED
Public Sub SelChange(Cancel As Integer)
    If lCountHighlightedID = 0 Then
        Cancel = True
    Else
        Cancel = False
        tgTable.SelStartCol = -1
        tgTable.SelEndCol = -1
        subReHighlightRows
    End If
End Sub

Private Sub subGetSelected(lTemp() As Long, nCount As Long)
    'Return the array of selected row numbers
    'and the count of the selected rows
    Dim i As Long
    
    With tgTable.SelBookmarks
        nCount = .Count
        
        If nCount > 0 Then
            ReDim lTemp(nCount - 1)
            For i = 0 To nCount - 1
                lTemp(i) = Val(.Item(i))
            Next i
        End If
    End With
End Sub

Private Sub subClearSelBookmarks()
    'selected column
    tgTable.SelStartCol = -1
    tgTable.SelEndCol = -1
    
    'selected row
    With tgTable.SelBookmarks
        While .Count > 0
            .Remove 0
        Wend
    End With
End Sub

Private Sub subSelectRow(ByVal lRow As Long)
    With tgTable.SelBookmarks
        .Add MakeBookmark(lRow)
        SetRowHighLighted True, True, True
    End With
End Sub

Private Sub subSaveHighlightedRowID(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim j As Long
    Dim k As Long
    Dim lSelected() As Long
    
    lCountHighlightedID = 0
    m_bColIsSorted = False
    
    If Not m_bSortOnCol Then
        Exit Sub
    End If
    
    'is not header click
    If y >= tgTable.RowTop(0) Then
        Exit Sub
    End If
    
    On Error GoTo errTrap
    
    'store the highlighted row id
    subGetSelected lSelected, lCountHighlightedID
    
    ReDim aryHighlightedID(RowCount - 1)
    
    If lCountHighlightedID > 0 Then
        For j = 0 To RowCount - 1
            For k = 0 To lCountHighlightedID - 1
                If j = lSelected(k) Then
                    aryHighlightedID(j) = 1
                    Exit For
'Debug.Print "j="; j
                End If
            Next k
        Next j
    End If
    
    subClearSelBookmarks
    SetRowHighLighted True, True
errTrap:
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub subReHighlightRows()
    Dim j As Long
    
    If lCountHighlightedID = 0 Then
        With tgTable.SelBookmarks
            While .Count > 0
                .Remove 0
            Wend
        End With
    Else
        subClearSelBookmarks
        
        For j = 0 To RowCount - 1
            If aryHighlightedID(j) = 1 Then
                subSelectRow j
            End If
        Next j
    End If
End Sub

'david 06/05/2002
'refresh the text of a cell automatically
'when the content is changed in CellValue() let property.
Property Let AutoRefreshCellValue(bStatus As Boolean)
    m_bCellValueAutoRebind = bStatus
End Property

Private Function fnIsLastEditableCol(Optional nCol As Integer = -1) As Boolean
    Dim i As Integer
    
    If nColumnCount = 0 Then
        Exit Function
    End If
    
    If nCol = -1 Then
        nCol = tgTable.Col
    End If
    
    fnIsLastEditableCol = nCol = nEditColumns(nColumnCount - 1)
End Function

'david 08/10/2002
Private Sub subResetSortOrder()
    Dim j As Integer
    
    If m_bReverseSortOnCol Then
        For j = 0 To m_SortColCount - 1
            m_arySortCol(0, j).nSortOrder = m_arySortCol(0, j).nInitSortOrder
        Next j
    End If
End Sub

Public Function LastEditableCol(Optional bReturnColZeroIfNone As Boolean = True) As Integer
    If nColumnCount = 0 Then
        If bReturnColZeroIfNone Then
            LastEditableCol = 0
        Else
            LastEditableCol = 9999
        End If
    Else
        LastEditableCol = nEditColumns(nColumnCount - 1)
    End If
End Function

Public Function FirstEditableCol(Optional bReturnColZeroIfNone As Boolean = True) As Integer
    If nColumnCount = 0 Then
        If bReturnColZeroIfNone Then
            FirstEditableCol = 0
        Else
            FirstEditableCol = -1
        End If
    Else
        FirstEditableCol = nEditColumns(0)
    End If
End Function

'david 02/26/2004  #401172-1
Property Get IsEditableCol(nCol As Integer) As Boolean
    Dim i As Integer
    
    For i = 0 To nColumnCount - 1
        If nCol = nEditColumns(i) Then
            IsEditableCol = True
            Exit For
        End If
    Next i
End Property
''''''''''''''''''''''''''''

'david 03/26/2004
'to be call in clsTGSelector class
'for the fix to allow the highlighting to go to the last new row
'when the last row is valid

'when AllowAddNew
'this function returns True if the current row is able to change to next row
'returns False if the current row is NOT able to change to next row.
'one reason is: when the current row is reaching end of the grid and the last row is not valid
'
'if bAction = True then the selected row will be changed to next row
'if bAction = False then the selected row/current row will remain unchanged
Public Function GotoNextRow(bAction As Boolean) As Boolean
    GotoNextRow = fnGotoNextRow(bAction)
End Function

'david 12/12/2002
'Implement new function to show the progress bar
'while populating the grid with recordset.

'NOTE: A PUBLIC Call back function called "ShowProgress(fPercent As Single,...)"
'is required in the Form (frmParent) where the progress bar resides.

'IN ORDER TO USE THIS NEW FEATURE,
'THE CONDITIONAL COMPILATION FLAG 'TGEDITOR_EX' MUST BE INCLUDED AND SET TO -1
'IN THE PROJECT

#If TGEDITOR_EX Then
    Public Sub FillWithRecordset_EX(rsTemp As Recordset, _
                                 Optional bRefresh As Variant, _
                                 Optional vValidate As Variant, _
                                 Optional vClear As Variant, _
                                 Optional bResetSortOrder As Boolean = True, _
                                 Optional bShowProgress As Boolean = False, _
                                 Optional ByVal sngTotalPercent As Single, _
                                 Optional ByVal sngPercentToGo As Single)
        
        Dim lRow As Long
        Dim nCol As Integer
        Dim i As Integer
        Dim sField As String
        Dim bValidate As Boolean
        Dim bClear As Boolean
        Dim lStart As Long
        ReDim nColPtr(nMaxCol - 1) As Integer
    
        Dim sngPercent As Single
    
        'david 02/09/2004  #401172-1
        If Not m_tgsSelector Is Nothing Then
            m_tgsSelector.ClearSelectedCells
        End If
        ''''''''''''''''''''''''''''
        
        'david 08/21/00
        If IsMissing(vClear) Then
            bClear = True
        Else
            bClear = vClear
        End If
    
        If rsTemp.RecordCount = 0 Then
            If bClear Then
                ClearData
                nRowCount = 0
            End If
            Exit Sub
        Else
            rsTemp.MoveLast
        End If
        
        If IsMissing(vValidate) Then
            bValidate = False
        Else
            bValidate = vValidate
        End If
        
        'david 05/06/2003  #408907
        'to trap the out of memory RTE
        On Error GoTo errHandler
        ''''''''''''''''''''''''''
    
        If bClear Then
            nMaxRow = rsTemp.RecordCount - 1
            lStart = 0
            fnInitialParams
        Else
            lStart = nRowCount
            nMaxRow = nRowCount + rsTemp.RecordCount - 1
            ReDim Preserve GridArray(nMaxCol + 1, nMaxRow + 1)
            
            If nColumnCount > 0 Then
                ReDim Preserve bValidCell(nColumnCount - 1, nMaxRow + 1)
                ReDim Preserve nCellChanged(nColumnCount - 1, nMaxRow + 1)
            End If
        End If
        
        Screen.MousePointer = vbHourglass
        
        rsTemp.MoveFirst
        
        On Error GoTo errFill
        For nCol = 0 To nMaxCol - 1
            sField = LCase(Trim(tgTable.Columns(nCol).DataField))
            nColPtr(nCol) = -1
            
            For i = 0 To rsTemp.Fields.Count - 1
                If LCase(Trim(rsTemp.Fields(i).Name)) = sField Then
                    nColPtr(nCol) = i
                    Exit For
                End If
            Next i
        Next nCol
    
        For lRow = lStart To nMaxRow
            Screen.MousePointer = vbHourglass
            
            For nCol = 0 To nMaxCol - 1
                If nColPtr(nCol) >= 0 Then
                    If IsNull(rsTemp.Fields(nColPtr(nCol))) Then
                        GridArray(nCol, lRow) = Null
                    Else
                        GridArray(nCol, lRow) = Trim(rsTemp.Fields(nColPtr(nCol)))
                    End If
                End If
            Next nCol
            
            GridArray(nFlagColumn, lRow) = ROW_INITIAL
            GridArray(nBookMarkColumn, lRow) = fnGetNewBookmark
            nRowCount = nRowCount + 1
            
            For nCol = 0 To nColumnCount - 1
                bValidCell(nCol, lRow) = True
                
                If bValidate Then
                    nCellChanged(nCol, lRow) = STATUS_NOCHANGE 'STATUS_CHANGED
                Else
                    nCellChanged(nCol, lRow) = STATUS_VALIDATED
                End If
            Next nCol
            
            If bShowProgress And nMaxRow >= 0 Then
                '#Protect division by zero problem
                '#Weiging, 07/18/03, 416179
                'sngPercent = sngTotalPercent + ((lRow - lStart + 1) / nMaxRow) * sngPercentToGo
                sngPercent = sngTotalPercent + ((lRow - lStart + 1) / (nMaxRow + 1)) * sngPercentToGo
                frmParent.ShowProgress sngPercent
            End If
            
            rsTemp.MoveNext
        Next lRow
        
        bClearData = True
        bValidating = True
        tgTable.Row = 0
        
        If IsMissing(bRefresh) Then
            tgTable.Rebind
        Else
            If bRefresh Then
                tgTable.Refresh
            End If
        End If
    
        DoEvents
        bClearData = False
        bValidating = False
    
        'david 08/10/2002
        'restore the original sort order
        If bResetSortOrder Then
            subResetSortOrder
        End If
        '''''''''''''''''
    
        Screen.MousePointer = vbDefault
    
extFill:
        On Error GoTo 0
        Exit Sub

errFill:
        MsgBox Err.Description, nCRITICAL   ', App.Caption
        Err.Clear
        Resume extFill
    
errHandler:
        'david 05/06/2003  #408907
        'to trap the out of memory RTE
        MsgBox "An error has occurred in function FillWithArray()." + vbCrLf + vbCrLf + "Error Code: " & Err.Number _
            & vbCrLf & "Error Desc:" & Err.Description & vbCrLf + vbCrLf + "Ppopulating grid data failed.", vbCritical
        ClearData
        Resume extFill
    End Sub
#End If

