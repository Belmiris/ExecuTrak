VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTGUnbound"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'***********************************************************'
'
' Copyright (c) 1996 FACTOR, A Division of W.R.Hess Company
'
' Module name   : TGUNBMGR.CLS - TGUnbound
' Date          : Nov 20, 1996
' Programmer(s) : David Chai, modified TGEDITOR.CLS - tgSpreadSheet of Qinggang Ma
'
' This module defines a class implementing unbound editing using true grid.


' An true grid editing is implemented
' Usage:
'      i. Include class files TGUNBMGR.CLS in the project
'     ii. Define and create a class: clsTGUnbound
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
'       v. Setup public function: fnValidateCellValue for validation
    
    Option Explicit
    
    Private Const nNULL_ROW_BOOKMARK As Integer = -2000
    Private Const nCRITICAL = vbCritical + vbOKOnly
    Private Const ROW_INCREMENT = 1
    Private Const ASC_RETURN = 13
    Private Const ASC_LEFT = 37
    Private Const ASC_RIGHT = 39
    Private Const ASC_DOWN = 40
    Private Const ASC_ESCAPE = 27
    Private Const ASC_BACK = 8
    Private Const ASC_CTRL_C = 3
    Private Const ASC_CTRL_V = &H16
    Private Const ASC_CTRL_X = &H18
    
    Private Const STATUS_NOCHANGE = 0
    Private Const STATUS_CHANGED = 1
    Private Const STATUS_WARNING = 2
    Private Const STATUS_VALIDATED = 4

    Private Const ROW_INITIAL As Integer = 0
    Private Const ROW_ADDED As Integer = 1
    Private Const ROW_EDITED As Integer = 2
    Private Const ROW_ADDED_EDITED As Integer = 3
    
    Private Const BACK_ONCE = 0
    Private Const BACK_ALWAYS = 1
    
    Private Const NOCURRENT = 1
    Private Const NEWRECORD = 2

    Const szROW_STYLE As String = "RowStyle"
    Const szCURRENT_CELL_STYLE As String = "CurrCellStyle"
    Const STYLE_INITIAL As Integer = 0
    Const STYLE_PLAIN As Integer = 1
    Const STYLE_HIGHLIGHT As Integer = 2
    
    Private CRLF As String * 2
    Private nMaxCol As Integer           'Number of columns
    Private nRowCount As Long              'Number of rows
    Private nMaxRow As Long
    Private nEditColumnCount As Integer   'Total number of editable columns

    Private GridArray() As Variant      'Array to store the data - our data source
    Private nCellChanged() As Integer
    Private bValidCell() As Boolean
    
    Private bmChangedRows() As Variant
    Private sPatternCode() As String
    Private sPrompt() As String
    Private sErrMessage() As String
    Private nEditColumns() As Integer
    Private nColumnType() As Integer
    Private sDefaultValue() As String
    Private sDisplayFormat() As String
    Private bFieldRequired() As Boolean
    
    Private Type tpESCapable
        nColumn As Integer
        nColumnFrom As Integer
    End Type
    Private nESCColumns() As tpESCapable
    
    Private sLastCellValue As String
    Private nColumnCount As Integer
    Private nFlagColumn  As Integer
    Private nBookMarkColumn As Integer
    Private nChangedCount As Long
    Private nMaxChanged As Long
    Private lEditingRow As Long
    Private nLastColumn As Integer
    Private lLastRow As Long
    Private bmLastRow As Variant
    Private bRowChanging As Boolean
    Private bValidating As Boolean
    Private bKeyProcessing As Boolean
    Private nCurrCol As Integer
    Private bCellEdited As Boolean
    Private bBlockAddNew As Boolean
    Private lBookMarkCount As Long
    
    Private bClearData As Boolean
    Private bFirstEdit As Boolean
    Private bAllowAddNew As Boolean
    Private nOriginalTableWidth As Integer
    Private nOriginalTableHeight As Integer
    Private nFontSize As Integer
    Private bBlockRowColChange As Boolean
    Private bFirstTimeGotFocus As Boolean
    Private nESCCount As Integer
    
    Private vCellValueBeforeColEdit As Variant

    Private bErrorFlag As Boolean

    Private tgTable As TDBGrid
    Private lblStatusBar As Label
    Private frmParent As Form
    
    Private m_nRowHighLighted As Integer
'


'========================================
'Public interface functions or procedures
'========================================

Public Sub AddEditColumn(nIndex As Integer, _
                         Optional sMsg As Variant, _
                         Optional vPtn As Variant, _
                         Optional vFormat As Variant, _
                         Optional bLocked As Variant, _
                         Optional vRequired As Variant, _
                         Optional vDefault As Variant, _
                         Optional nColType As Variant)
    
    If nIndex >= nMaxCol Or nColumnCount >= nMaxCol Then
        Exit Sub
    End If
    
    Dim i As Integer, j As Integer, aryValidCell() As Boolean, aryChangedCell() As Boolean
    Dim nPointOfInsert As Integer, sPtn As String
    
    bRowChanging = True
    If nColumnCount = 0 Then
        tgTable.col = nIndex
        nLastColumn = nIndex
    Else
        If nIndex < tgTable.col Then
            tgTable.col = nIndex
            nLastColumn = nIndex
        End If
    End If
    bRowChanging = False
    
    For i = 0 To nColumnCount - 1
        If nIndex = nEditColumns(i) Then
            'replace message, pattern, default value and column type
            If IsMissing(sMsg) Then
                sPrompt(i) = ""
            Else
                sPrompt(i) = sMsg
            End If
            If IsMissing(vPtn) Then
                sPatternCode(i) = ""
            Else
                sPtn = vPtn
                sPatternCode(i) = fnGetPatternCode(sPtn)
            End If
            If Not IsMissing(vDefault) Then
                sDefaultValue(i) = vDefault
            End If
            If IsMissing(nColType) Then
                nColumnType(i) = BACK_ONCE
            Else
                nColumnType(i) = nColType
            End If
            
            Exit Sub
        End If
    Next
    i = 0
    While i < nColumnCount And nIndex > nEditColumns(i)
        i = i + 1
    Wend
    
    nPointOfInsert = i
    
    For j = nColumnCount To i + 1 Step -1
        nEditColumns(j) = nEditColumns(j - 1)
        sPrompt(j) = sPrompt(j - 1)
        sPatternCode(j) = sPatternCode(j - 1)
        sDefaultValue(j) = sDefaultValue(j - 1)
        nColumnType(j) = nColumnType(j - 1)
        bFieldRequired(j) = bFieldRequired(j - 1)
    Next

    nEditColumns(i) = nIndex
    If IsMissing(bLocked) Then
        tgTable.Columns(nIndex).Locked = False
    ElseIf bLocked Then
        tgTable.Columns(nIndex).Locked = True
    Else
        tgTable.Columns(nIndex).Locked = False
    End If
    If IsMissing(vRequired) Then
        bFieldRequired(i) = True
    Else
        bFieldRequired(i) = vRequired
    End If
    If IsMissing(sMsg) Then
        sPrompt(i) = ""
    Else
        sPrompt(i) = sMsg
    End If
    If IsMissing(vPtn) Then
        sPatternCode(i) = ""
    Else
        sPtn = vPtn
        sPatternCode(i) = fnGetPatternCode(sPtn)
    End If
    If Not IsMissing(vDefault) Then
        sDefaultValue(i) = vDefault
    End If
    If IsMissing(nColType) Then
        nColumnType(i) = BACK_ONCE
    Else
        nColumnType(i) = nColType
    End If
    If Not IsMissing(vFormat) Then
        sDisplayFormat(nIndex) = vFormat
    End If
    subReDimAddEditCell nPointOfInsert
    nEditColumnCount = nEditColumnCount + 1

    If nEditColumnCount = 1 Then
        SetRowHighLighted False
        SetRowHighLighted True
    End If
End Sub

Public Sub SetRowHighLighted(bHighLighted As Boolean)
    If bHighLighted Then
        If m_nRowHighLighted = STYLE_HIGHLIGHT Then
            Exit Sub
        End If
    Else
        If m_nRowHighLighted = STYLE_PLAIN Then
            Exit Sub
        End If
    End If
    
    Dim m_styRow As Object            ' row style
    Dim m_styCurrentCell As Object    ' current cell style

    tgTable.ClearCellStyle dbgAllCells
    
    ' set some properties for the table
    Set m_styRow = tgTable.Styles.Item(szROW_STYLE)
    Set m_styCurrentCell = tgTable.Styles.Item(szCURRENT_CELL_STYLE)
    
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
    Else
        ' create plain row style
        With m_styRow
            .BackColor = &HFFFFFF
            .ForeColor = &H0
        End With
        m_nRowHighLighted = STYLE_PLAIN

    End If

    ' add the new styles to the table
    tgTable.AddCellStyle dbgMarqueeRow, m_styRow
    tgTable.AddCellStyle dbgMarqueeRow + dbgCurrentCell, m_styCurrentCell
    tgTable.AddCellStyle dbgMarqueeRow + dbgUpdatedCell + dbgCurrentCell, m_styCurrentCell
    tgTable.AddCellStyle dbgMarqueeRow + dbgUpdatedCell, m_styRow
    tgTable.AddCellStyle dbgUpdatedCell, m_styCurrentCell
End Sub

Public Sub AddESCColumn(nCol As Integer, nFrom As Integer)
    If nCol < 0 Or nCol >= nMaxCol Then
        Exit Sub
    End If
    
    If nESCCount = 0 Then
        ReDim nESCColumns(0)
    Else
        ReDim Preserve nESCColumns(nESCCount)
    End If
    nESCColumns(nESCCount).nColumn = nCol
    nESCColumns(nESCCount).nColumnFrom = nFrom
    nESCCount = nESCCount + 1
End Sub

Public Sub AddHiddenField(sField As String)
    
    tgTable.Columns.Add tgTable.Columns.Count
    nMaxCol = tgTable.Columns.Count
    tgTable.Columns(nMaxCol - 1).DataField = sField
    tgTable.Columns(nMaxCol - 1).Visible = False
    tgTable.Columns(nMaxCol - 1).Width = 0
    
    nFlagColumn = nMaxCol
    nBookMarkColumn = nMaxCol + 1
    ReDim GridArray(nMaxCol + 1, 0)
    ReDim Preserve sDisplayFormat(nMaxCol - 1)
    sDisplayFormat(nMaxCol - 1) = ""

End Sub

Public Sub AddKeyFields(sFields As String)
    MsgBox "Sub AddKeyFields() is no long supported"
End Sub

Public Sub AfterColEdit(ByVal ColIndex As Integer)
    Dim nCol As Integer, lRow As Long
    
    If MaxRows < 1 Then
        Exit Sub
    End If
    nCol = fnColumnNumber(ColIndex)
    lRow = GetCurrentRowNumber
    If lEditingRow < 0 Then
        lEditingRow = lRow
    End If
    If tgTable.Columns(ColIndex).ValueItems.Count = 0 Then
        CellValue(ColIndex, lRow) = tgTable.Columns(ColIndex).Value
    End If
    If nCol >= 0 Then
        nCellChanged(nCol, lRow) = nCellChanged(nCol, lRow) And (Not STATUS_VALIDATED)
        bErrorFlag = fnValidate(ColIndex, lRow, CellValue(ColIndex, lRow))
        bValidating = False
    End If
    
    On Error GoTo ErrorInvalidType
    If CellValue(ColIndex, lRow) = vCellValueBeforeColEdit Then
        lEditingRow = -1
        Exit Sub
    End If
    If Trim$(Str$(CellValue(ColIndex, lRow))) = Trim$(Str$(vCellValueBeforeColEdit)) Then
        lEditingRow = -1
        Exit Sub
    End If
    fnSaveBookMarks fnMakeBookmark(GetCurrentRowNumber)
    Exit Sub
ErrorInvalidType:
    On Error GoTo 0
    If tgTable.Columns(ColIndex).ValueItems.Count > 0 Then
        fnSaveBookMarks fnMakeBookmark(GetCurrentRowNumber)
    End If
End Sub

Public Sub BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    Dim i As Integer
    Dim nCol As Integer

    If MaxRows < 1 Then
        Exit Sub
    End If
    If tgTable.Columns(ColIndex).Locked Then
        On Error Resume Next
        tgTable.col = nEditColumns(0)
        Exit Sub
    End If
    
    bCellEdited = True
    nCol = fnColumnNumber(ColIndex)
    If nCol >= 0 Then
        lEditingRow = GetCurrentRowNumber()
    End If
    
    nCellChanged(nCol, lEditingRow) = STATUS_CHANGED
    
    vCellValueBeforeColEdit = CellValue(ColIndex, lEditingRow)
End Sub

Public Sub ClearData(Optional bSetFocus As Variant)
    ReDim GridArray(nMaxCol + 1, 0)
    ReDim bValidCell(nColumnCount, 0)
    ReDim nCellChanged(nColumnCount, 0)
    ReDim bmChangedRows(0)

    Dim i As Integer
    
    Class_Initialize
    
    nChangedCount = 0
    nLastColumn = 0
    lLastRow = 0
    lBookMarkCount = 0
    lEditingRow = -1
    tgTable.col = nEditColumns(0)
    bClearData = True
    GridArray(nBookMarkColumn, 0) = fnMakeBookmark(0)
    tgTable.Refresh
    bClearData = False
    If Not IsMissing(bSetFocus) Then
        tgTable.SetFocus
    End If
End Sub

Public Function IsRowChanged(lRow As Long) As Boolean
    Dim i As Integer
    
    IsRowChanged = False
    For i = 0 To nColumnCount - 1
        If (nCellChanged(i, lRow) And STATUS_CHANGED) = STATUS_CHANGED Then
            IsRowChanged = True
            Exit Function
        End If
    Next
End Function

Public Function IsEmptyRow(Optional ByVal lTempRow As Variant) As Boolean
    Dim i As Integer, lRow As Long
    
    If IsMissing(lTempRow) Then
        lRow = GetCurrentRowNumber
    Else
        lRow = lTempRow
        If lRow < 0 Or lRow > nMaxRow Then
'            MsgBox "Invalid row number"
            IsEmptyRow = False
            Exit Function
        End If
    End If
    For i = 0 To nMaxCol - 1
        If Not IsNull(GridArray(i, lRow)) Then
            If GridArray(i, lRow) <> "" Then
                IsEmptyRow = False
                Exit Function
            End If
        End If
    Next i
    IsEmptyRow = True
End Function

Public Function IsEndOfTable() As Boolean
    Dim nDiff As Integer
    
    'calculate the different because the visible rows in table may be more than one over the max. rows
    If tgTable.VisibleRows > nMaxRow Then
        nDiff = tgTable.VisibleRows - nMaxRow
    End If
    
    IsEndOfTable = False
    If tgTable.Row >= tgTable.VisibleRows - 1 - nDiff Then
        IsEndOfTable = True
    End If

End Function

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

Public Sub RemoveEditColumn(nIndex As Integer)
    Dim nPointOfRemove As Integer, i As Integer
    
    If nIndex >= nMaxCol Or nColumnCount = 0 Then
        Exit Sub
    End If
    
    i = 0
    While nEditColumns(i) <> nIndex And i < nColumnCount
        i = i + 1
    Wend
    
    If i >= nColumnCount Then
        Exit Sub
    End If
    
    nPointOfRemove = i
    
    If i < nColumnCount And nEditColumns(i) = nIndex Then
        tgTable.Columns(nIndex).Locked = True
        While i < nColumnCount - 1
            nEditColumns(i) = nEditColumns(i + 1)
            sPrompt(i) = sPrompt(i + 1)
            sPatternCode(i) = sPatternCode(i + 1)
            nColumnType(i) = nColumnType(i + 1)
            sDefaultValue(i) = sDefaultValue(i + 1)
            bFieldRequired(i) = bFieldRequired(i + 1)
            i = i + 1
        Wend
        subReDimRemoveEditCell nPointOfRemove
    End If
    nEditColumnCount = nEditColumnCount - 1
    If nEditColumnCount < 0 Then
        nEditColumnCount = 0
    End If
    If nEditColumnCount = 0 Then
        SetRowHighLighted False
        SetRowHighLighted True
    End If
End Sub

Public Sub GetRecords(vData() As Variant, lRows As Long, sStatement As String)
    Const ROWINC = 1
    Dim sParams(20) As String
    Dim nParams(20) As Integer
    Dim sParamValue(20) As Variant
    Dim nParamCount As Integer
    Dim sConds(20) As String
    Dim sCondValue(20) As Variant
    Dim nCond(20) As Integer
    Dim nCondCount As Integer
    Dim bFound As Boolean
    Dim sName As String
    Dim sValue As String
    Dim lMRows As Long
    Dim nPos As Integer
    Dim nLen As Integer
    Dim nFlag As Integer
    Dim i As Long
    Dim j As Integer
    
    nPos = InStr(sStatement, "WHERE")
    If nPos = 0 Then
        MsgBox "Syntax error when get records", nCRITICAL, frmParent.Caption
        Exit Sub
    End If

    nLen = Len(sStatement)
    fnParseString sParams, sStatement, 0, nPos - 1
    fnParseString sConds, sStatement, nPos + 5, nLen
    
    nFlag = 0
    i = 0
    While i < 20 And sConds(i) <> ""
        If sConds(i) = "NOCURRENT" Or sConds(i) = "NEW" Then
            If sConds(i) = "NOCURRENT" Then
                nFlag = nFlag Or NOCURRENT
            ElseIf sConds(i) = "NEW" Then
                nFlag = nFlag Or NEWRECORD
            End If
            j = i
            While sConds(j + 1) <> ""
                sConds(j) = sConds(j + 1)
                j = j + 1
            Wend
            sConds(j) = ""
            i = i - 1
        End If
        i = i + 1
    Wend
    
    fnParseAssigments nCond(), sCondValue(), nCondCount, sConds()
    fnParseAssigments nParams(), sParamValue(), nParamCount, sParams()

    If nParamCount = 0 Then
        Exit Sub
    End If
    lMRows = ROWINC
    ReDim vData(nParamCount - 1, lMRows)
    lRows = 0
    For i = 0 To nRowCount - 1
        bFound = True
        If (nFlag And NOCURRENT) = NOCURRENT And i = GetCurrentRowNumber Then
            bFound = False
        End If
        If (nFlag And NEWRECORD) = NEWRECORD And GridArray(nFlagColumn, i) <> ROW_ADDED Then
            bFound = False
        End If
        j = 0
        While bFound And j < nCondCount
            If IsNull(GridArray(nCond(j), i)) Then
                bFound = False
            Else
                If Trim$(GridArray(nCond(j), i)) <> sCondValue(j) Then
                    bFound = False
                End If
            
            End If
            j = j + 1
        Wend
        If bFound Then
            For j = 0 To nParamCount - 1
                vData(j, lRows) = GridArray(nParams(j), i)
            Next j
            lRows = lRows + 1
            If lRows > lMRows Then
                lMRows = lMRows + ROWINC
                ReDim Preserve vData(nParamCount - 1, lMRows)
            End If
        End If
    Next i
End Sub

Public Sub GetRow(vData() As Variant, ByRef nCount As Integer, ByVal lIndex As Long)
    Dim i As Integer
    Dim lRow As Long
    ReDim vData(nMaxCol)
    
    nCount = 0
    If lIndex >= 0 And lIndex <= nRowCount Then
        For i = 0 To nMaxCol
            vData(i) = GridArray(i, lIndex)
        Next i
        nCount = nMaxCol
    End If
End Sub

Public Function GetBlankRowCount() As Long
    
    Dim i As Long, j As Long, bkMark As Variant, lOrgRowNo As Long, bBlank As Boolean, lCount As Long
    
    lCount = 0
    
    For i = 0 To nChangedCount - 1
        bkMark = bmChangedRows(i)
        lOrgRowNo = fnIndexFromBookmark(bkMark, False)
        If lOrgRowNo = nNULL_ROW_BOOKMARK Then
            GetBlankRowCount = 0
            Exit Function
        End If
        bBlank = True
        For j = 0 To nMaxCol - 1
            If Not IsNull(GridArray(j, lOrgRowNo)) Then
                If GridArray(j, lOrgRowNo) <> "" Then
                    bBlank = False
                    Exit For
                End If
            End If
        Next j
        If bBlank Then
            lCount = lCount + 1
        End If
    Next i

    GetBlankRowCount = lCount
End Function

Public Sub FillRowWithArray(vData() As Variant, ByVal lIndex As Long, lTableRow As Long)
    Dim i As Integer, lCurrRow
    
    lCurrRow = tgTable.Row
    
    If lIndex >= 0 And lIndex <= nRowCount Then
        For i = LBound(vData) To UBound(vData)
            GridArray(i, lIndex) = vData(i)
        Next i
        'the ltablerow is 3 means 4th row but total visible rows only 3, thus add (- 1)
        If lTableRow >= tgTable.FirstRow And lTableRow <= tgTable.FirstRow + tgTable.VisibleRows Then
            tgTable.Row = lTableRow
            For i = 0 To nMaxCol - 1
                tgTable.Columns(i).Text = vData(i)
            Next i
            tgTable.Row = lCurrRow
        End If
    End If
End Sub

Public Sub GotFocus()
    fnShowStatus GetCurrentRowNumber, tgTable.col
    If nRowCount = 0 And IsNull(fnMakeBookmark(0)) Then
        fnFetchDefault 0, 0
    End If
End Sub

Public Sub GotoCol(nCol As Integer)
    If nCol >= 0 And nCol < nMaxCol - 2 Then
        tgTable.col = nCol
        fnShowStatus GetCurrentRowNumber, nCol
    End If
End Sub

Public Function IsNewRow() As Boolean
    If IsNull(fnMakeBookmark(GetCurrentRowNumber)) Then
        IsNewRow = True
    Else
        IsNewRow = False
    End If
End Function

Public Sub Rebind()
    tgTable.Rebind
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

Public Sub SetCurrentRowData(vData() As Variant)
    Dim i As Integer
    Dim nCol As Integer
    Dim lRow As Long
    lRow = GetCurrentRowNumber
    If lRow < 0 Then
        Exit Sub
    End If
    For i = 0 To UBound(vData)
        If vData(i) <> "" Then
            tgTable.Columns(i).Text = vData(i)
            GridArray(i, lRow) = vData(i)
            nCol = fnColumnNumber(i)
            If nCol >= 0 Then
                nCellChanged(i, lRow) = (nCellChanged(i, lRow) Or STATUS_CHANGED) And (Not STATUS_VALIDATED)
            End If
        End If
    Next
    GridArray(nFlagColumn, lRow) = ROW_ADDED
End Sub

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
    Dim sCrts(20) As String
    Dim sCrtValue(20) As Variant
    Dim nCrts(20) As Integer
    Dim nCrtCount As Integer
    Dim bFound As Boolean
    Dim sName As String
    Dim sValue As String
    Dim nPos As Integer
    Dim nLen As Integer
    Dim i As Long
    Dim j As Integer
    
    nPos = InStr(sStatement, "WHERE")
    If nPos = 0 Then
        MsgBox "Syntax error when set a cell value", nCRITICAL, frmParent.Caption
        Exit Sub
    End If

    nLen = Len(sStatement)
    fnParseString sParams, sStatement, 0, nPos - 1
    fnParseString sCrts, sStatement, nPos + 5, nLen
    fnParseAssigments nCrts(), sCrtValue(), nCrtCount, sCrts()
    fnParseAssigments nParams(), sParamValue(), nParamCount, sParams()
    
    For i = 0 To nRowCount - 1
        For j = 0 To nCrtCount - 1
            bFound = True
            If IsNull(GridArray(nCrts(j), i)) Then
                bFound = False
                Exit For
            Else
                If Trim$(GridArray(nCrts(j), i)) <> sCrtValue(j) Then
                    bFound = False
                    Exit For
                End If
            End If
        Next j
        If bFound Then
            For j = 0 To nParamCount - 1
                GridArray(nParams(j), i) = sParamValue(j)
                If i = GetCurrentRowNumber Then
                    tgTable.Columns(nParams(j)).Text = sParamValue(j)
                End If
            Next j
            If IsMissing(nStatus) Then
                GridArray(nFlagColumn, i) = ROW_INITIAL
            Else
                GridArray(nFlagColumn, i) = nStatus
                If nStatus <> ROW_INITIAL Then
                    fnSaveBookMarks fnMakeBookmark(i)
                End If
            End If
        End If
    Next i
End Sub

Public Sub SetValueByRow(sStatement As String, lRow As Long, Optional nStatus As Variant)
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

Public Function ValidData(Optional lRow As Variant) As Boolean
    Dim i As Long
    Dim j As Integer
    Dim lBegin As Long
    Dim lEnd As Long
    Dim sTemp As String
    
    If lRow = nNULL_ROW_BOOKMARK Then
        ValidData = False
        Exit Function
    End If
    
    If IsMissing(lRow) Then
        lBegin = 0
        lEnd = nRowCount - 1
        If lEnd < GetCurrentRowNumber Then
            lEnd = GetCurrentRowNumber
        End If
    Else
        lBegin = lRow
        lEnd = lRow
    End If

    ValidData = True
    For i = lBegin To lEnd
        If i <= lEnd Or GridArray(nFlagColumn, i) <> "" Then  'Problem ?
            For j = 0 To nColumnCount - 1
                sTemp = GridArray(j, i)
                fnValidate j, i, sTemp
                If Not bValidCell(j, i) Then
                    ValidData = False
                    Exit Function
                End If
            Next j
        Else
            For j = 0 To nColumnCount - 1
                If GridArray(j, i) <> "" Then
                    ValidData = False
                    Exit Function
                End If
            Next j
        End If
    Next i
End Function

'add optional argument-  ndbType, remote or local database indentifier
'if ndbType is missing then get the data from the remote database
Public Sub FillWithSQL(dbSource As DataBase, strSQL As String, Optional bRefresh As Variant, Optional nDBType As Variant)
    Dim rsTemp As Recordset
    On Error GoTo SQLError
    If IsMissing(nDBType) Then
        Set rsTemp = dbSource.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    Else
        Set rsTemp = dbSource.OpenRecordset(strSQL, dbOpenSnapshot)
    End If
    FillWithRecordset rsTemp, bRefresh
quitsub:
    On Error GoTo 0
    Exit Sub
SQLError:
    Err.Clear
    MsgBox "An error occurred while doing the SQL query" & CRLF & CRLF & Err.Description, vbOKOnly + vbInformation, frmParent.Caption
    Resume quitsub
    
End Sub

'FillWithBlank functin - David
Public Sub FillWithBlank(Optional lTempRow As Variant)
    Dim i As Integer, lRow As Long
    
    If IsMissing(lTempRow) Then
        lRow = GetCurrentRowNumber
    Else
        lRow = lTempRow
        If lRow < 0 Or lRow > nMaxRow Then
            Exit Sub
        End If
    End If
    bClearData = True
    For i = 0 To nMaxCol - 1
        CellValue(i, lRow) = ""
    Next i
    bClearData = False
End Sub

Public Function GetCurrentRowNumber() As Long

    Dim lRow As Long
    
    If nRowCount = 0 Then
        GetCurrentRowNumber = 0
        Exit Function
    End If
    
    If tgTable.FirstRow = "" Then
        Exit Function
    End If
    
    lRow = tgTable.FirstRow + tgTable.Row
    If lRow < 0 Then
        GetCurrentRowNumber = 0
    Else
        GetCurrentRowNumber = lRow
    End If

End Function

Public Function GetChangedRowNumber(ndxChangedRow As Long) As Long
    Dim lOrgRowNo As Long, bkMark As Variant
    
    bkMark = bmChangedRows(ndxChangedRow)
    lOrgRowNo = fnIndexFromBookmark(bkMark, False)
    GetChangedRowNumber = lOrgRowNo
End Function

Public Function Change() As Boolean
'
' Check the changed text against the regular expression for that column
'

    Dim nCol As Integer, i As Integer
    Dim szText As String
    Dim nCode As Integer
        
    Change = False
    
    If tgTable Is Nothing Then
        Exit Function
    End If

    nCol = fnColumnNumber(tgTable.col)
    If nCol < 0 Then
        Exit Function
    End If
    
    If sPatternCode(nCol) = "" Then
        Change = True
        Exit Function
    End If

    ' is the column text blank?
    szText = tgTable.Columns(tgTable.col).Text
    
    ' check the pattern
    nCode = PRegExpMatch(sPatternCode(nCol), Len(sPatternCode(nCol)), szText)
        
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

Public Sub DeleteRow(Optional ByVal lCurrRow As Variant)  '(Bookmark As Variant)
' If the DeleteRow function doesn't fail set the Bookmark to Null for this row being deleted.
' The actual data manipulation for the delete is done in the DeleteRow function.
    Dim bookmark As Variant
    Dim lRow As Long
    Dim lFRow As Long
    
    If nRowCount = 0 Or tgTable.Row = nRowCount Then
        Exit Sub
    End If

    If Not IsMissing(lCurrRow) Then
        lRow = lCurrRow
        bookmark = fnMakeBookmark(lRow)
    Else
        lRow = tgTable.Row
        bookmark = tgTable.bookmark
    End If
    
'Comment out and change March, 26 1997
'    If lRow = 0 Then
'        If nRowCount > 1 Then
'            tgTable.Row = tgTable.Row + 1
'        End If
'    Else
'        tgTable.Row = tgTable.Row - 1
'    End If
    If lRow > 0 Then
'        tgTable.Row = tgTable.Row - 1  changed 4-21-97
        tgTable.Row = lRow - 1
    End If
    
    If Not fnDeleteRow(bookmark) Then 'Bookmark = Null
        MsgBox "Delete row error"
    Else
        On Error Resume Next
        tgTable.Rebind
'???????? is these lines needed ??????????????????
'        If lRow >= nRowCount Then
'            lRow = nRowCount - 1
'        End If
'        If lRow >= 0 Then
'            tgTable.Row = lRow
'        End If
    End If
End Sub

Public Sub FillWithArray(arryData() As Variant, Optional vClear As Variant)
    Dim lRow As Long, lArrayRow As Long
    Dim nCol As Integer
    Dim nMCol As Integer
    
    Dim i As Long, bClear As Boolean, lStart As Long
    
    If IsMissing(vClear) Then
        bClear = True
    Else
        bClear = vClear
    End If

''''''''''''''vvvvvvvvv    Ma's code below  vvvvvvvvvvvvvvvvvvvvv''''''''''''''''''''
    
    'this codes from Ma's TGSpreadSheet class 4-2-97
    'when the old data needed to be kept in the table
    If bClear Then
        lStart = 0
        On Error GoTo ErrorOutofRange
        lArrayRow = UBound(arryData, 2)
        On Error GoTo 0
        nMCol = UBound(arryData, 1) + 1
        If nMCol > nMaxCol Then
            nMCol = nMaxCol
        End If
        fnInitialParams
    Else
        lStart = MaxRows
        On Error GoTo ErrorOutofRange
        lArrayRow = MaxRows + UBound(arryData, 2)
        On Error GoTo 0
        ReDim Preserve GridArray(nMaxCol + 1, lArrayRow)
        If nColumnCount > 0 Then
            ReDim Preserve bValidCell(nColumnCount, lArrayRow + 1)
            ReDim Preserve nCellChanged(nColumnCount, lArrayRow + 1)
        End If
        nMCol = UBound(arryData, 1) + 1
        If nMCol > nMaxCol Then
            nMCol = nMaxCol
        End If
    End If
    
    For i = lStart To lArrayRow
        lRow = i - lStart
        For nCol = 0 To nMCol - 1
            GridArray(nCol, i) = arryData(nCol, lRow)
        Next nCol
        GridArray(nFlagColumn, i) = ROW_INITIAL
        GridArray(nBookMarkColumn, i) = fnGetNewBookmark
        nRowCount = nRowCount + 1
        For nCol = 0 To nColumnCount - 1
            bValidCell(nCol, i) = True
            nCellChanged(nCol, i) = STATUS_VALIDATED
        Next nCol
    Next i

'    bmLastRow = BookMark(0)
    bClearData = True
    bValidating = True
    tgTable.Row = 0
    If bClear Then
        If nEditColumns(0) < tgTable.Columns.Count Then
            tgTable.col = nEditColumns(0)
        End If
        tgTable.Rebind
        DoEvents
    Else
        Rebind
        DoEvents
    End If
    bClearData = False
    bValidating = False
    Exit Sub
''''''''''''''''  ^^^^^^^^^^^^^^^^^^^^^^  Ma's codes above  ^^^^^^^^^^^^^^^^ ''''''''''''''''
    
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'my old FillWithArray(arryData() As Variant, Optional vRow As Variant) codes
    If bClear Then
        On Error GoTo ErrorOutofRange
        lArrayRow = UBound(arryData, 2)
        nMCol = UBound(arryData, 1) + 1
        If nMCol > nMaxCol Then
            nMCol = nMaxCol
        End If
    
        On Error GoTo 0
        fnInitialParams
    
        For lRow = 0 To lArrayRow
            For nCol = 0 To nMCol - 1
                GridArray(nCol, lRow) = arryData(nCol, lRow)
            Next nCol
            GridArray(nFlagColumn, lRow) = ROW_INITIAL
            GridArray(nBookMarkColumn, lRow) = fnGetNewBookmark
            For nCol = 0 To nColumnCount - 1
                bValidCell(nCol, lRow) = True
                nCellChanged(nCol, lRow) = STATUS_VALIDATED
            Next nCol
        Next lRow
        
        bmLastRow = fnMakeBookmark(0)
        bRowChanging = True
        bValidating = True
        
        sLastCellValue = GridArray(nEditColumns(0), 0)
        tgTable.col = nEditColumns(0)
        tgTable.Row = 0
    
        tgTable.Refresh
    
        bRowChanging = False
        bValidating = False
        
        bmLastRow = fnMakeBookmark(0)
        BlockRowColChange = False
    'added 8-20-96
        nLastColumn = tgTable.col
        lLastRow = tgTable.Row
        Exit Sub
    End If
    
    Exit Sub
    
ErrorOutofRange:
    #If DEBUG_DAVID Then
        MsgBox "FillWithArray: Array was not initialized"
    #End If
    On Error GoTo 0
End Sub

Public Sub FillWithRecordset(rsTemp As Recordset, Optional bRefresh As Variant)
    Dim lRow As Long
    Dim nCol As Integer
    Dim i As Integer
    Dim sField As String
    ReDim nColPtr(nMaxCol - 1) As Integer

    If rsTemp.RecordCount = 0 Then
        Exit Sub
    End If

    rsTemp.MoveLast
    rsTemp.MoveFirst
    
    fnInitialParams
    
    On Error GoTo errFill
    
    For nCol = 0 To nMaxCol - 1
        sField = tgTable.Columns(nCol).DataField
        nColPtr(nCol) = -1
        For i = 0 To rsTemp.Fields.Count - 1
            If rsTemp.Fields(i).Name = sField Then
                nColPtr(nCol) = i
                Exit For
            End If
        Next i
    Next nCol

    For lRow = 0 To rsTemp.RecordCount - 1
        For nCol = 0 To nMaxCol - 1
            If nColPtr(nCol) >= 0 Then
                If rsTemp.Fields(nColPtr(nCol)) <> "" Then
                    GridArray(nCol, lRow) = Trim$(rsTemp.Fields(nColPtr(nCol)))
                End If
            End If
        Next nCol
        GridArray(nFlagColumn, lRow) = ROW_INITIAL
        GridArray(nBookMarkColumn, lRow) = fnGetNewBookmark
        For nCol = 0 To nColumnCount - 1
            If GridArray(nCol, lRow) <> "" Then
                bValidCell(nCol, lRow) = True
                nCellChanged(nCol, lRow) = STATUS_VALIDATED
            End If
        Next nCol
        rsTemp.MoveNext
    Next lRow
    
    bmLastRow = fnMakeBookmark(0)
    bRowChanging = True
    bValidating = True
    
    sLastCellValue = GridArray(nEditColumns(0), 0)
    tgTable.Row = 0
    
    tgTable.Refresh
    
    bRowChanging = False
    bValidating = False

extFill:
    On Error GoTo 0
    Exit Sub
errFill:
    MsgBox Err.Description, nCRITICAL   ', App.Caption
    Err.Clear
    Resume extFill
End Sub

Public Sub GetOriginalRow(vData() As Variant, nCount As Integer, lIndex As Long)
    subNoSupport "Sub GetOriginalRow()"
End Sub

Public Sub GetChangedRow(vData() As Variant, nCount As Integer, lIndex As Long)
    Dim i As Integer
    Dim lRow As Long
    ReDim vData(nMaxCol - 1)
    
    If lIndex < 0 Or lIndex >= nChangedCount Or nMaxRow <= 0 Then
        nCount = 0
        Exit Sub
    End If
    lRow = fnIndexFromBookmark(bmChangedRows(lIndex), False)
    If lRow >= 0 And lRow < nRowCount Then
        For i = 0 To nMaxCol - 1
            vData(i) = GridArray(i, lRow)
        Next i
    End If
    nCount = nMaxCol
End Sub

Public Sub KeyDown(KeyCode As Integer, Shift As Integer)
    If tgTable.Row < 0 Then
        tgTable.Row = 0
    End If
'    If KeyCode = ASC_RETURN Then
    If KeyCode = ASC_RETURN Or KeyCode = ASC_DOWN Then 'changed 10-29-96
        If bCellEdited Then
            bCellEdited = False
            If tgTable.col = nEditColumns(nColumnCount - 1) Then
                If fnGotoNextRow Then
                    tgTable.col = nEditColumns(0)
                    KeyCode = 0
                    Exit Sub
                End If
            End If
        End If
        If tgTable.EditActive Then
            If tgTable.col < nEditColumns(nColumnCount - 1) Then
                tgTable.EditActive = False
            Else
                KeyCode = 0
            End If
        End If
    End If
        
    If bKeyProcessing Then
        bKeyProcessing = False
        Exit Sub
    End If
    
    bKeyProcessing = True
    If nColumnCount = 0 Then
        bKeyProcessing = False
        Exit Sub
    End If
    
    If KeyCode = ASC_RIGHT Or KeyCode = ASC_RETURN Then
        If tgTable.col >= nEditColumns(nColumnCount - 1) Then
            KeyCode = 0
            If IsEmptyRow() Then  'no row is added if the current row is a empty row
                bKeyProcessing = False
                Exit Sub
            End If
            'add a row
            If fnGotoNextRow(False) Then
                tgTable.col = nEditColumns(0)
            End If
            KeyCode = 0
        End If
        bCellEdited = False
    ElseIf KeyCode = ASC_LEFT Then
        If tgTable.col <= nEditColumns(0) Then
            If fnGotoPreviousRow Then
                tgTable.col = nEditColumns(nColumnCount - 1)
            End If
            KeyCode = 0
        End If
        bCellEdited = False
    ElseIf KeyCode = ASC_DOWN Then
        KeyCode = 0
        If IsEmptyRow() Then  'no row is added if the current row is a empty row
            bKeyProcessing = False
            lblStatusBar.Font.Bold = True
            lblStatusBar.Caption = "No new rows can be added before the current row has been filled"
            lblStatusBar.ForeColor = &HFF&
            lblStatusBar.Refresh
            Exit Sub
        End If
        If fnGotoNextRow Then
            If GetCurrentRowNumber = nRowCount - 1 Then
                If tgTable.Columns(tgTable.col).Text = "" Then
                    tgTable.col = nEditColumns(0)
                End If
            End If
        End If
        KeyCode = 0
        bCellEdited = False
    End If
    bKeyProcessing = False
End Sub

Public Function KeyPress(KeyAscii) As Boolean

'
' Check the key pressed against the regular expression
'
    Dim nCol As Integer
    Dim szText As String
    Dim nCode As Integer
    
    KeyPress = False
    
    If tgTable Is Nothing Or nMaxRow <= 0 Then
        Exit Function
    End If
    
    ' check enter, backspace, escape
    If KeyAscii = ASC_RETURN Or KeyAscii = ASC_BACK Or KeyAscii = ASC_ESCAPE Then
        KeyPress = True
        Exit Function
    End If

    ' check for cut/copy/paste keys
    If KeyAscii = ASC_CTRL_C Or KeyAscii = ASC_CTRL_V Or KeyAscii = ASC_CTRL_X Then
        KeyPress = True
        Exit Function
    End If
        
    nCol = fnColumnNumber(tgTable.col)
    If nCol < 0 Then
        Exit Function
    End If

    If sPatternCode(nCol) = "" Then
        KeyPress = True
        Exit Function
    End If

    ' get the data up to the cursor position and tack on the key pressed
    szText = Left(tgTable.Columns(tgTable.col).Text, tgTable.SelStart) & Chr(KeyAscii)

    nCode = PRegExpMatch(sPatternCode(nCol), Len(sPatternCode(nCol)), szText)
    
    If nCode < 0 Or (nCode <> 0 And nCode <> tgTable.SelStart + 2) Then
        KeyPress = False
        If KeyAscii <> 32 Then
            KeyAscii = 0
        End If
        Beep
    Else
        KeyPress = True
        If lEditingRow >= 0 And lEditingRow < nMaxRow Then
            If GridArray(nFlagColumn, lEditingRow) = ROW_ADDED Then
                GridArray(nFlagColumn, lEditingRow) = ROW_ADDED_EDITED
            Else
                GridArray(nFlagColumn, lEditingRow) = ROW_EDITED
            End If
        End If
    End If
End Function

Public Sub LostFocus()
    lblStatusBar.Font.Bold = False
    bFirstTimeGotFocus = True
    If nMaxRow <= 0 Then
        Exit Sub
    End If
    
    If lEditingRow >= 0 Then
        fnSaveBookMarks tgTable.bookmark
        lEditingRow = -1
    End If
End Sub

Public Sub RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    DoEvents
    If bClearData Or tgTable Is Nothing Or nColumnCount = 0 Or bBlockRowColChange Then
        If bBlockRowColChange Then
            bBlockRowColChange = False
        End If
        Exit Sub
    End If
    
    Dim nCurrRow As Long
    Dim nCurrCol As Integer
    
    nCurrRow = GetCurrentRowNumber()
    If nCurrRow < 0 Then
        Exit Sub
    End If
    fnShowStatus nCurrRow, tgTable.col
    
    If bRowChanging Or lLastRow < 0 Or lLastRow > nRowCount Then
        nLastColumn = tgTable.col
        lLastRow = tgTable.Row
        Exit Sub
    End If
    bRowChanging = True
    
    Dim nCol As Integer

    If GetCurrentRowNumber <> lEditingRow Then
        If lEditingRow >= 0 Then
            fnSaveBookMarks fnMakeBookmark(lEditingRow)
            lEditingRow = -1
        End If
    End If
    
    If fnEscapable(tgTable.col) = LastCol Then
        bRowChanging = False
'        nLastColumn = tgTable.col     'tgTable.col
'        lLastRow = tgTable.Row        'tgTable.Row
        Exit Sub
    End If
    
    Dim sLastCellValue As String
    
    sLastCellValue = GridArray(nLastColumn, lLastRow)
    nCol = fnColumnNumber(nLastColumn)
    
    fnValidate nLastColumn, lLastRow, sLastCellValue
    If nCol >= 0 Then
        If (nCellChanged(nCol, lLastRow) And STATUS_WARNING) <> STATUS_WARNING Then
            bErrorFlag = Not bValidCell(nCol, lLastRow)
            nCellChanged(nCol, lLastRow) = nCellChanged(nCol, lLastRow) Or STATUS_WARNING
        End If
    End If
    If bErrorFlag Then
        tgTable.col = nLastColumn
        On Error Resume Next
        tgTable.Row = lLastRow
        On Error GoTo 0
        DoEvents
        bRowChanging = False
        nCol = fnColumnNumber(tgTable.col)
        If nCol >= 0 Then
            If nColumnType(nCol) <> BACK_ALWAYS Then
                bErrorFlag = False
            End If
        End If
        fnShowStatus lLastRow, tgTable.col
        Exit Sub
    End If

    nCol = fnColumnNumber(tgTable.col)
    nCurrCol = tgTable.col
    If nCol < 0 Then
        fnChangeToEditingCol nLastColumn
    Else
        nLastColumn = nCurrCol     'tgTable.col
        lLastRow = nCurrRow        'tgTable.Row
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
End Sub

Public Sub SetPattern(nIndex As Integer, Optional vPtn As Variant)
    If nIndex >= nColumnCount Then
        Exit Sub
    End If
    Dim nCol As Integer
    nCol = fnColumnNumber(nIndex)
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

Public Sub SetPrompt(nIndex As Integer, Optional sMsg As Variant)
'    If nIndex >= nColumnCount Then
'        Exit Sub
'    End If
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

Public Function ChangedRowValid(ByVal lRow As Long, Optional ByVal bAllBlankIsValid As Variant, Optional ByVal bAnyRow As Variant) As Boolean
'check changed row to see if the whole row's data is valid.  *** Take into consideration: If the all columns
'in the last row are empty, the changed row (ie. last row) is valid.
    Dim bkMark As Variant, vData() As Variant, lRowNumber As Long, i As Long
    
    bkMark = bmChangedRows(lRow)
    lRowNumber = fnIndexFromBookmark(bkMark, False)
    ChangedRowValid = ValidData(lRowNumber)
    If (Not ChangedRowValid) Then
        If Not IsMissing(bAllBlankIsValid) Then
            If bAllBlankIsValid Then
                For i = 0 To nMaxCol - 1
                    If Not IsNull(GridArray(i, lRowNumber)) Then
                        If Trim$(GridArray(i, lRowNumber)) <> "" Then
                            Exit Function
                        End If
                    End If
                Next i
            Else
                Exit Function
            End If
            If IsMissing(bAnyRow) Then
                If lRowNumber = nMaxRow - 1 Then ChangedRowValid = True
            Else
                If bAnyRow Then ChangedRowValid = True
            End If
        End If
    End If
End Function


'===================
'True Grid Functions
'===================

'UnboundReadData event - David
Public Sub UnboundReadData(ByVal RowBuf As RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
    
    Dim bookm As Variant, RelPos As Integer, RowsFetched As Integer
    Dim i As Long, j As Integer
    
    bookm = StartLocation
        
    RowBuf.RowCount = nMaxRow
    
    If ReadPriorRows Then
        RelPos = -1 ' Requesting data in rows prior to StartLocation
    Else
        RelPos = 1  ' Requesting data in rows after StartLocation
    End If
    
    RowsFetched = 0
    
    For i = 0 To RowBuf.RowCount - 1
        ' Get the bookmark of the next available row
        bookm = fnGetRelativeBookmark(bookm, RelPos)
    
        ' If the next is BOF or EOF, then done
        If IsNull(bookm) Then Exit For
    
        For j = 0 To RowBuf.ColumnCount - 1
            RowBuf.Value(i, j) = fnGetUserData(bookm, j)
        Next j
    
        ' Set the bookmark for the row
        RowBuf.bookmark(i) = bookm
    
        ' Increment the count of fetched rows
        RowsFetched = RowsFetched + 1
    Next i
    
    ' Tell the grid how many rows were fetched
    RowBuf.RowCount = RowsFetched

End Sub

'UnboundWriteData event - David
Public Sub UnboundWriteData(ByVal RowBuf As RowBuffer, WriteLocation As Variant)
' Assume that a VB function StoreUserData(bookm, col, value)
' takes a row bookmark, a column index, and a variant with the
' appropriate data to be stored in an array or database.  The
' function returns True if the data is acceptable and can be
' stored, False otherwise.

' Loop over all the columns of the row, storing non-Null values
    
    Dim i As Integer
    
    For i = 0 To RowBuf.ColumnCount - 1
        If Not IsNull(RowBuf.Value(0, i)) Then
            If Not fnStoreUserData(WriteLocation, i, RowBuf.Value(0, i)) Then
            
                RowBuf.RowCount = 0 ' Tell the grid the update failed.
                Exit Sub            ' It failed, so exit the event.
            End If
        End If
    Next i

End Sub

'UnboundAddData event - David
Public Sub UnboundAddData(ByVal RowBuf As RowBuffer, NewRowBookmark As Variant)
' Assume that a VB function StoreUserData(bookm, col, value)
' takes a row bookmark, a column index, and a variant with the
' appropriate data to be stored in an array or database.  The
' function returns True if the data is acceptable and can be
' stored, False otherwise.

' First, get a bookmark for the new row.  Do this with a VB
' function, GetNewBookmark.  This function allocates a new
' row of data in the storage medium (array or database), and

' returns a variant containing a bookmark for the added row.
    
    Dim newval As Variant, i As Integer
    
    NewRowBookmark = fnGetNewBookmark()
    
    ' Loop over all the columns of the row, storing non-Null values
    For i = 0 To RowBuf.ColumnCount - 1
        newval = RowBuf.Value(0, i)
        If IsNull(newval) Then
            ' The RowBuf does not contain a value for this column.
            ' A default value should be set.  A convenient value
            ' is the default value for the column.
            newval = tgTable.Columns(i).DefaultValue
    
        End If
    
        ' Now store the new values.
        If Not fnStoreUserData(NewRowBookmark, i, newval) Then
            ' Storage of the data has failed.  Delete the added
            ' row using a VB function, fnDeleteRow, which takes a
            ' bookmark as an argument.  Also, fail the update by
            ' clearing the RowCount.
            fnDeleteRow NewRowBookmark
            RowBuf.RowCount = 0 ' Tell the grid the update failed.
            Exit Sub            ' It failed, so exit the event.
    
        End If
    Next i

End Sub

'UnboundDeleteRow event - David
Public Sub UnboundDeleteRow(bookmark As Variant)

    If Not fnDeleteRow(bookmark) Then bookmark = Null
    tgTable.Rebind
End Sub

'fnNewInsertRow functin - David
Public Function InsertNewRow(bookm As Variant, ByVal bInsertAfter As Boolean) As Boolean

    Dim Index As Long, lMaxRow As Long, lSkipRow, i As Long, j As Integer
    
    If bookm = 0 And nMaxRow = 0 Then 'insert the very first row
        GridArray(nBookMarkColumn, 0) = fnGetNewBookmark
        fnSaveBookMarks GridArray(nBookMarkColumn, 0)
        
        GridArray(nFlagColumn, 0) = ROW_ADDED
        
        bmLastRow = fnMakeBookmark(0)
        bClearData = True
        sLastCellValue = GridArray(nEditColumns(0), 0)
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
    
    ReDim Preserve GridArray(nMaxCol + 1, nMaxRow)
    ReDim Preserve nCellChanged(nColumnCount, nMaxRow + 1)
    ReDim Preserve bValidCell(nColumnCount, nMaxRow + 1)

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
    GridArray(nBookMarkColumn, i) = fnMakeBookmark(i)
    
    're-new the bookmark in bmChangedRows array is needed!
    'because all bookmarks in the GridArray will be re-sequence after rebind
    
    For j = 0 To nChangedCount - 1
        For i = 0 To nMaxRow - 1
            If bmChangedRows(j) = GridArray(nBookMarkColumn, i) Then
                bmChangedRows(j) = fnMakeBookmark(i)
                Exit For
            End If
        Next i
    Next j
    
    'Now re-new Bookmark Column
    For i = 0 To nMaxRow - 1
        GridArray(nBookMarkColumn, i) = fnMakeBookmark(i)
    Next i
    
    InsertNewRow = True
    
    On Error Resume Next
    tgTable.Rebind
    tgTable.SetFocus

End Function

Public Sub fnShowStatus(lRow As Long, nCol As Integer, Optional vLastCol As Variant)
    Dim nECol As Integer
    Dim bFlag As Boolean
    Dim nLastCol As Integer
    
    
    nECol = fnColumnNumber(nCol)
    If nECol < 0 Then
        Exit Sub
    End If

    If tgTable.Columns(nCol).Locked Then
        Exit Sub
    End If
    
    If Not IsMissing(vLastCol) Then
        nLastCol = vLastCol
        If nLastCol < 0 Then
            nLastCol = 0
        End If
        If fnEscapable(nCol) = nLastCol And GridArray(nLastCol, lRow) = "" Then
            lblStatusBar.Font.Bold = False
            lblStatusBar.Caption = sPrompt(nECol)
            lblStatusBar.ForeColor = &H80000012
            Exit Sub
        End If
    End If
    
    If bValidCell(nECol, lRow) Then
'        If (nCellChanged(nECol, lRow) And STATUS_CHANGED) = STATUS_CHANGED Then
        If (nCellChanged(nECol, lRow) And STATUS_CHANGED) = STATUS_CHANGED And bFieldRequired(nCol) Then
            lblStatusBar.Font.Bold = True
            lblStatusBar.ForeColor = &H8000&  'Green text
            lblStatusBar.Caption = "Data is valid"
        Else
            lblStatusBar.Font.Bold = False
            lblStatusBar.Caption = sPrompt(nECol)
            lblStatusBar.ForeColor = &H80000012
        End If
    Else
        If (nCellChanged(nECol, lRow) And STATUS_WARNING) = STATUS_WARNING Then
            lblStatusBar.Font.Bold = True
            lblStatusBar.ForeColor = &HFF&       'Red Text on
            If sErrMessage(nECol) = "" Then
                lblStatusBar.Caption = "Data is invalid; " & sPrompt(nECol)
            Else
                lblStatusBar.Caption = sErrMessage(nECol)
            End If
        Else
            lblStatusBar.Font.Bold = False
            lblStatusBar.Caption = sPrompt(nECol)
            lblStatusBar.ForeColor = &H80000012
        End If
    End If
    lblStatusBar.Refresh
End Sub

Public Sub EnableRecordSelector(bOnOff As Boolean)
    tgTable.RecordSelectors = bOnOff
End Sub

'============================
'Get, Set, and Let properties
'============================
Property Let ErrorMessage(nCol As Integer, sMsg As String)
    Dim nFCol As Integer
    nFCol = fnColumnNumber(nCol)

    If nFCol >= 0 Then
        sErrMessage(nFCol) = sMsg
    End If
End Property

Property Get BlockRowColChange()
    BlockRowColChange = bBlockRowColChange
End Property

Property Let BlockRowColChange(bOnOff)
    bBlockRowColChange = bOnOff
End Property

Property Get FirstTimeGotFocus()
    FirstTimeGotFocus = bFirstTimeGotFocus
End Property

Property Let FirstTimeGotFocus(bOnOff)
    bFirstTimeGotFocus = bOnOff
End Property

Property Get MaxRows() As Long
    MaxRows = nMaxRow
End Property

Property Get GetEditColumnCount() As Integer
    GetEditColumnCount = nEditColumnCount
End Property

Property Get GetRowStatus(ByVal lRow As Long) As Integer
    GetRowStatus = GridArray(nFlagColumn, lRow)
End Property

Property Get AddedRow() As Integer
    AddedRow = ROW_ADDED
End Property

Property Let AllowAddNew(bStatus As Boolean)
    bAllowAddNew = bStatus
End Property

Property Get BackAlways() As Integer
    BackAlways = BACK_ALWAYS
End Property

Property Get BackOnce() As Integer
    BackOnce = BACK_ONCE
End Property

Property Let OrigCellValue(nCol As Integer, lRow As Long, vValue As Variant)
    subNoSupport "Property Let OrigCellValue()"
End Property

Property Get OrigCellValue(nCol As Integer, lRow As Long) As Variant
    subNoSupport "Property Get OrigCellValue()"
End Property

Property Get CellValue(nCol As Integer, ByVal lRow As Long) As Variant
    If nCol >= nMaxCol Or lRow > nMaxRow Then
        Exit Property
    End If
    CellValue = GridArray(nCol, lRow)
End Property

Property Get ColumnNumber(sField As String)
    Dim i As Integer
    ColumnNumber = -1
    For i = 0 To tgTable.Columns.Count - 1
        If Trim$(tgTable.Columns(i).DataField) = Trim$(sField) Then
            ColumnNumber = i
            Exit Property
        End If
    Next
End Property

Property Get EditedRow() As Integer
    EditedRow = ROW_EDITED
End Property

Property Get InitialRow() As Integer
    InitialRow = ROW_INITIAL
End Property

Property Get RowCount() As Integer
    RowCount = nRowCount
End Property

Property Let CellValue(nCol As Integer, ByVal lRow As Long, vValue As Variant)
    
    Dim lCurrRow As Long
    
    If nCol >= nMaxCol Or lRow > nMaxRow Then
        Exit Property
    End If
    
    If tgTable.Columns(nCol).Visible Then
        lCurrRow = tgTable.Row
        If lCurrRow < 0 Then
            Exit Property
        End If
        If tgTable.FirstRow = "" Then
            Exit Property
        End If
        If lRow >= tgTable.FirstRow And lRow < tgTable.FirstRow + tgTable.VisibleRows Then
            tgTable.Row = lRow - tgTable.FirstRow
            If sDisplayFormat(nCol) = "" Then
                If tgTable.Columns(nCol).ValueItems.Count = 0 Then
                    tgTable.Columns(nCol).Text = vValue
                End If
            Else
                tgTable.Columns(nCol).Text = Format$(vValue, sDisplayFormat(nCol))
            End If
            tgTable.Row = lCurrRow
        End If
    End If
    
    GridArray(nCol, lRow) = vValue
    
End Property

'Property Let ValidatedFlag(nCol As Integer, lRow As Long, bStatus As Boolean)
Property Let ValidatedFlag(nCol As Integer, lRow As Long, nStatus As Integer)
    Dim nFCol As Integer
    nFCol = fnColumnNumber(nCol)
    If nFCol >= 0 Then
        nCellChanged(nFCol, lRow) = nStatus
'        If bStatus Then
'            nCellChanged(nFCol, lRow) = nCellChanged(nFCol, lRow) Or STATUS_VALIDATED
'        Else
'            nCellChanged(nFCol, lRow) = nCellChanged(nFCol, lRow) And (Not STATUS_VALIDATED)
'        End If
    End If

End Property

Property Get ValidCell(nCol As Integer, lRow As Long) As Boolean
    If lRow < 0 Then
        Exit Property
    End If
    Dim nFCol As Integer
    Dim sTemp As String
    nFCol = fnColumnNumber(nCol)
    If nFCol >= 0 Then
        If (nCellChanged(nFCol, lRow) And STATUS_VALIDATED) <> STATUS_VALIDATED Then
            sTemp = GridArray(nCol, lRow)
            fnValidate nCol, lRow, sTemp
        End If
        ValidCell = bValidCell(nFCol, lRow)
    Else
        ValidCell = True
    End If
End Property

Property Let ValidCell(nCol As Integer, lRow As Long, bStatus As Boolean)

    On Error GoTo ExpectedError:
    
    Dim nFCol As Integer
    nFCol = fnColumnNumber(nCol)
    If nFCol >= 0 Then
        bValidCell(nFCol, lRow) = bStatus
        If bStatus Then
            nCellChanged(nFCol, lRow) = STATUS_VALIDATED
        Else
            nCellChanged(nFCol, lRow) = STATUS_CHANGED
        End If
    End If
    
    Exit Property
    
ExpectedError:
' this error occur due to row col change,  and will be ignore
' err.number=9, err.description=subscript out of range
    Resume Next
End Property

Property Get CellValueBeforeColEdit() As Variant
    CellValueBeforeColEdit = vCellValueBeforeColEdit
End Property

Property Set Form(frmMain As Object)
    Set frmParent = frmMain
End Property

Property Get GetChangedRowCount()
    GetChangedRowCount = nChangedCount
End Property

Property Set StatusBar(lblBar As Object)
    Set lblStatusBar = lblBar
End Property

Property Get Table() As Object
    Set Table = tgTable
End Property

Property Set Table(tblGrid As Object)
'
' Add a True DBGrid to the collection and set its default properties
'
    Dim n As Integer
    Dim m_styRow As Object            ' row style
    Dim m_styCurrentCell As Object    ' current cell style
    
    bAllowAddNew = True
    
    On Error GoTo AddToCollectionError
    
    Set tgTable = tblGrid
    nOriginalTableWidth = tgTable.Width
    nOriginalTableHeight = tgTable.Height
    nFontSize = tgTable.Font.Size

    nMaxCol = tblGrid.Columns.Count
    nFlagColumn = nMaxCol
    nBookMarkColumn = nMaxCol + 1
    nEditColumnCount = 0
    ReDim sPatternCode(nMaxCol - 1)
    ReDim sPrompt(nMaxCol - 1)
    ReDim sErrMessage(nMaxCol - 1)
    ReDim nEditColumns(nMaxCol - 1)
    ReDim nColumnType(nMaxCol - 1)
    ReDim sDefaultValue(nMaxCol - 1)
    ReDim sDisplayFormat(nMaxCol - 1)
    ReDim bFieldRequired(nMaxCol - 1)

    nCurrCol = 0

    On Error GoTo 0
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
    
    On Error Resume Next
    tgTable.Styles.Remove szROW_STYLE
    tgTable.Styles.Remove szCURRENT_CELL_STYLE
    On Error GoTo 0

    ' set some properties for the table
    Set m_styRow = tgTable.Styles.Add(szROW_STYLE)
    Set m_styCurrentCell = tgTable.Styles.Add(szCURRENT_CELL_STYLE)
    
    'initial row style setting
'    SetRowHighLighted False
    SetRowHighLighted True
    Exit Property
    
AddToCollectionError:
    
    On Error GoTo 0

End Property

Property Let DisplayFormat(nCol As Integer, sFormat As String)
    If nCol > 0 And nCol < nMaxCol Then
        sDisplayFormat(nCol) = sFormat
    End If
End Property

Property Get DisplayFormat(nCol As Integer) As String
    If nCol > 0 And nCol < nMaxCol Then
        DisplayFormat = sDisplayFormat(nCol)
    End If
End Property

Property Get RowIsChanged(ByVal lRow As Long) As Boolean
    Dim i As Integer
    
    RowIsChanged = False
    Dim bkMark As Variant
    bkMark = bookmark(lRow)
    For i = 0 To nChangedCount - 1
        If bkMark = bmChangedRows(i) Then
            RowIsChanged = True
            Exit For
        End If
    Next
End Property

Property Let RowIsChanged(ByVal lRow As Long, bStatus As Boolean)
    If bStatus Then
        If lRow < 0 Then lRow = 0
        fnSaveBookMarks bookmark(lRow)
    Else
        If nChangedCount = 0 Then
            Exit Property
        End If
        Dim i As Integer
        Dim bkMark As Variant
        bkMark = bookmark(lRow)
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


'================================
'private functions and procedures
'================================

Private Sub subReDimAddEditCell(nIndex As Integer)

    Dim i As Long, j As Integer, TempCellArray() As Boolean, TempChangedArray() As Integer
    
    ReDim TempCellArray(nColumnCount, nMaxRow)
    ReDim TempChangedArray(nColumnCount, nMaxRow)
    
    'save valid cell and changed cell array
    For i = 0 To nMaxRow - 1
        For j = 0 To nColumnCount - 1
            TempCellArray(j, i) = bValidCell(j, i)
            TempChangedArray(j, i) = nCellChanged(j, i)
        Next j
    Next i
    
    nColumnCount = nColumnCount + 1
    
    If nColumnCount > 0 Then
        ReDim bValidCell(nColumnCount, nMaxRow)
        ReDim nCellChanged(nColumnCount, nMaxRow)
        'restore valid cell and changed cell array
        If nIndex < nColumnCount - 1 Then
            For i = 0 To nMaxRow - 1
                For j = 0 To nIndex - 1
                    bValidCell(j, i) = TempCellArray(j, i)
                    nCellChanged(j, i) = TempChangedArray(j, i)
                Next j
                For j = nIndex + 1 To nColumnCount - 1
                    bValidCell(j, i) = TempCellArray(j - 1, i)
                    nCellChanged(j, i) = TempChangedArray(j - 1, i)
                Next j
            Next i
        End If
    End If
    
End Sub

Private Sub subReDimRemoveEditCell(nIndex As Integer)

    Dim i As Long, j As Integer, TempCellArray() As Boolean, TempChangedArray() As Integer
    
    ReDim TempCellArray(nColumnCount, nMaxRow)
    ReDim TempChangedArray(nColumnCount, nMaxRow)
    
    'save valid cell and changed cell array
    For i = 0 To nMaxRow - 1
        For j = 0 To nColumnCount - 1
            TempCellArray(j, i) = bValidCell(j, i)
            TempChangedArray(j, i) = nCellChanged(j, i)
        Next j
    Next i
    
    nColumnCount = nColumnCount - 1
    
    If nColumnCount > 0 Then
        ReDim bValidCell(nColumnCount, nMaxRow)
        ReDim nCellChanged(nColumnCount, nMaxRow)
        'restore valid cell and changed cell array
        If nIndex < nColumnCount - 1 Then
            For i = 0 To nMaxRow - 1
                For j = 0 To nIndex - 1
                    bValidCell(j, i) = TempCellArray(j, i)
                    nCellChanged(j, i) = TempChangedArray(j, i)
                Next j
                For j = nIndex + 1 To nColumnCount
                    bValidCell(j - 1, i) = TempCellArray(j, i)
                    nCellChanged(j - 1, i) = TempChangedArray(j, i)
                Next j
            Next i
        End If
    End If

End Sub

Private Function fnEscapable(nCol As Integer) As Integer
    Dim i As Integer
    
    fnEscapable = -1
    For i = 0 To nESCCount - 1
        If nCol = nESCColumns(i).nColumn Then
            fnEscapable = nESCColumns(i).nColumnFrom
            Exit For
        ElseIf nCol = nESCColumns(i).nColumnFrom Then
            fnEscapable = nESCColumns(i).nColumn
            Exit For
        End If
    Next
End Function

Private Sub fnChangeToEditingCol(nLastCol As Integer)
    Dim nColL As Integer
    Dim nColC As Integer
    Dim i As Integer
    
    nColL = fnColumnNumber(nLastCol)
    nColC = fnColumnNumber(tgTable.col)
    If nLastCol < tgTable.col Then
        If nColL < nColumnCount - 1 Then
            i = tgTable.col
            While nColC < 0 And i < nMaxCol
                i = i + 1
                nColC = fnColumnNumber(i)
            Wend
            If i < nMaxCol Then
                tgTable.col = i
            End If
        Else
            If fnGotoNextRow Then
                tgTable.col = nEditColumns(0)
            Else
                tgTable.col = nLastColumn
            End If
        End If
    Else
        If nColL > 0 Then
            i = tgTable.col
            While nColC < 0 And i > 0
                i = i - 1
                nColC = fnColumnNumber(i)
            Wend
            If i < nMaxCol Then
                tgTable.col = i
            End If
        Else
            If fnGotoPreviousRow Then
                tgTable.col = nEditColumns(nColumnCount - 1)
            Else
                tgTable.col = nLastCol
            End If
        End If
    End If

End Sub

Private Sub fnFetchDefault(nCol As Integer, lCurrRow As Long)
    If lCurrRow < 0 Then Exit Sub
    
    If nCellChanged(nCol, lCurrRow) = STATUS_NOCHANGE And tgTable.Columns(nLastColumn).Text = "" Then
        If nCol = 0 Then
            If IsNull(fnMakeBookmark(lCurrRow)) Then
                GridArray(nBookMarkColumn, lCurrRow) = fnMakeBookmark(lCurrRow)
            End If
        End If
        If Not IsNull(fnMakeBookmark(lCurrRow)) Then
            If sDefaultValue(nCol) <> "" Then
                If sDisplayFormat(nCol) = "" Then
                    tgTable.Columns(nLastColumn).Text = sDefaultValue(nCol)
                Else
                    tgTable.Columns(nLastColumn).Text = Format$(sDefaultValue(nCol), sDisplayFormat(nCol))
                End If
                DoEvents
                nCellChanged(nCol, lCurrRow) = nCellChanged(nCol, lCurrRow) Or STATUS_CHANGED
'                bCellEdited = True
                GridArray(nLastColumn, lCurrRow) = sDefaultValue(nCol)
            Else
                If sPatternCode(nCol) = "" Then
                    bValidCell(nCol, lCurrRow) = True
                End If
            End If
        End If
    End If
End Sub

Private Function fnGetFieldNumber(sSrc As String)
    Dim i As Integer
    fnGetFieldNumber = -1
    For i = 0 To tgTable.Columns.Count - 1
        If tgTable.Columns(i).DataField = sSrc Then
            fnGetFieldNumber = i
            Exit For
        End If
    Next
End Function

Private Sub fnInitialParams()
    nRowCount = 0
    nMaxRow = 0
    nChangedCount = 0
    nLastColumn = tgTable.col
    lLastRow = tgTable.Row
    bCellEdited = False
    lBookMarkCount = 0
    lEditingRow = -1
    ReDim GridArray(nMaxCol + 1, nMaxRow)
    If nColumnCount > 0 Then
        ReDim bValidCell(nColumnCount, nMaxRow + 1)
        ReDim nCellChanged(nColumnCount, nMaxRow + 1)
    End If
End Sub

Private Function fnValidate(nCurrCol As Integer, lCurrRow As Long, szText As String) As Boolean
    Dim nCol As Integer
    Dim nCode As Integer
    
    If bValidating Then
        Exit Function
    End If
    bValidating = True
    
    nCol = fnColumnNumber(nCurrCol)
    If nCol < 0 Then
        bValidating = False
        Exit Function
    End If
    If ((nCellChanged(nCol, lCurrRow) And STATUS_VALIDATED) = STATUS_VALIDATED) Then
        bValidating = False
        Exit Function
    End If
    
    If bFieldRequired(nCol) Then
        If szText = "" Then
            bValidCell(nCol, lCurrRow) = frmParent.fnValidateCellValue(nCurrCol, GetCurrentRowNumber, szText, tgTable.TabIndex)
            nCellChanged(nCol, lCurrRow) = nCellChanged(nCol, lCurrRow) Or STATUS_VALIDATED
            bValidating = False
            Exit Function
        End If
    Else
        bValidCell(nCol, lCurrRow) = True
        nCellChanged(nCol, lCurrRow) = nCellChanged(nCol, lCurrRow) Or STATUS_VALIDATED
        bValidating = False
        Exit Function
    End If
    
    If sPatternCode(nCol) <> "" Then
        nCode = PRegExpMatch(sPatternCode(nCol), Len(sPatternCode(nCol)), szText)
    Else
        nCode = 0
    End If
    If nCode = 0 Then
        bValidCell(nCol, lCurrRow) = frmParent.fnValidateCellValue(nCurrCol, GetCurrentRowNumber, szText, tgTable.TabIndex)
    Else
        bValidCell(nCol, lCurrRow) = False
    End If
    fnValidate = Not bValidCell(nCol, lCurrRow)
    nCellChanged(nCol, lCurrRow) = nCellChanged(nCol, lCurrRow) Or STATUS_VALIDATED
    bValidating = False
End Function

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
                          Optional vStart As Variant, _
                          Optional vEnd As Variant)
    Dim i1 As Integer
    Dim i2 As Integer
    Dim k As Integer
    Dim nEnd As Integer
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
        i2 = InStr(i1, sSrc, ",")
        If i2 > i1 And i2 < nEnd Then
            sParam(k) = Trim$(Mid$(sSrc, i1, i2 - i1))
            k = k + 1
            i1 = i2 + 1
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
        nPos = fnGetFieldNumber(sName)
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

Public Function bookmark(Optional ByVal vRow As Variant) As Variant
    Dim nRow As Long
    
    If IsMissing(vRow) Then
        nRow = GetCurrentRowNumber
    Else
        nRow = vRow
    End If
    If nRow <= nRowCount Then
        bookmark = GridArray(nBookMarkColumn, nRow)
    Else
        bookmark = Null
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
            Exit For
        End If
    Next
    bmChangedRows(lRow) = Bookmk
    If lRow >= nChangedCount Then
        nChangedCount = nChangedCount + 1
    End If
End Sub

Private Function fnColumnNumber(nIndex) As Integer
    Dim i As Integer
    
    fnColumnNumber = -1
    For i = 0 To nColumnCount - 1
        If nEditColumns(i) = nIndex Then
            fnColumnNumber = i
            Exit Function
        End If
    Next
End Function

Private Function fnGetPatternCode(sRegExp As String)
'
'   store regular expressions for the columns
'
    Dim szBuffer As String
    Dim nSize As Integer

    szBuffer = Space(4096)
    
    nSize = GetRegExpPCode(sRegExp, szBuffer)

    If nSize <= 0 Then
        szBuffer = ""
    End If
            
    fnGetPatternCode = szBuffer
End Function

Private Function fnGotoNextRow(Optional bAction As Variant) As Boolean
    Dim lRow As Long
    Dim i As Integer
    Dim szText As String

    On Error GoTo retFalse
    fnGotoNextRow = False
    lRow = tgTable.FirstRow + tgTable.Row
    If tgTable.AllowAddNew Then
        lRow = lRow - 1
    Else
        If lRow = nRowCount - 1 Then
            If Not bAllowAddNew Then
                Exit Function
            End If
        End If
    End If
    If lRow >= nRowCount Then
        Exit Function
    End If
    lRow = GetCurrentRowNumber()
    If lRow >= nRowCount - 1 Then
        For i = 0 To nColumnCount - 1
            szText = tgTable.Columns(nEditColumns(i)).Text
            fnValidate nEditColumns(i), lRow, szText
            If Not bValidCell(i, lRow) Then
                lblStatusBar.Font.Bold = True
                lblStatusBar.Caption = "No new rows can be added before previous inputs are valid"
                lblStatusBar.ForeColor = &HFF&
                lblStatusBar.Refresh
                Exit Function
            End If
        Next
    End If
    If lRow >= nRowCount - 1 Then
        'append a new row
        GridArray(nBookMarkColumn, nRowCount - 1) = fnGetNewBookmark
        fnSaveBookMarks GridArray(nBookMarkColumn, nRowCount - 1)
        
        GridArray(nFlagColumn, nRowCount - 1) = ROW_ADDED
        For i = 0 To nMaxCol - 1
            GridArray(i, nRowCount - 1) = ""
        Next i
        tgTable.Rebind
        DoEvents
    End If
    
    tgTable.Row = tgTable.Row + 1
    fnGotoNextRow = True
    Exit Function
    
    If IsMissing(bAction) Then
        SendKeys "{DOWN}", True
    Else
        If bAction Then
            SendKeys "{DOWN}", True
        End If
    End If
    DoEvents
    fnGotoNextRow = True
retTrue:
    On Error GoTo 0
    Exit Function
retFalse:
    If Err.Number = 13 Then
        lRow = tgTable.Row
        Resume Next
    ElseIf Err.Number = 6148 Then
        SendKeys "{DOWN}", True
        fnGotoNextRow = True
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
    End If
    fnGotoPreviousRow = True
retTrue1:
    On Error GoTo 0
    Exit Function
retFalse1:
    tgTable.bookmark = 0
    tgTable.Row = 0
    Err.Clear
    Resume retTrue1
End Function

Private Sub Class_Initialize()
    nMaxChanged = -1
    nMaxRow = 0
    nRowCount = 0
    bFirstEdit = True
    bFirstTimeGotFocus = True
    CRLF = Chr(&HD) & Chr(&HA)
    m_nRowHighLighted = STYLE_INITIAL
End Sub


'=========================
' TDBGrid global functions
'=========================

'MakeBookmark function - David
Private Function fnMakeBookmark(Index As Long) As Variant

' This support function is used only by the remaining support
' functions.  It is not used directly by the unbound events.

    fnMakeBookmark = Str$(Index)
    
End Function

Public Function MakeBookmark(ByVal Index As Long) As Variant

' This support function is used only by the remaining support
' functions.  It is not used directly by the unbound events.

    MakeBookmark = Str$(Index)
    
End Function

'IndexFromBookmark function - David
Private Function fnIndexFromBookmark(bookm As Variant, ReadPriorRows As Boolean) As Long
' This support function is used only by the remaining support
' functions.  It is not used directly by the unbound events.

    If IsNull(bookm) Then
        If ReadPriorRows Then
            fnIndexFromBookmark = nMaxRow
        Else
            fnIndexFromBookmark = -1
        End If
    Else
        Dim Index As Long
        Index = val(bookm)
        If Index < 0 Or Index >= nMaxRow Then Index = nNULL_ROW_BOOKMARK
        fnIndexFromBookmark = Index
    End If

End Function

'GetNewBookmark function - David
Private Function fnGetNewBookmark() As Variant
' used to add new rows to the data source
' Creates a bookmark for a newly added row, called from the UnboundAddData event
' ReDim the array and create the new bookmark using the fnMakeBookmark function

    nMaxRow = nMaxRow + 1

    nRowCount = nRowCount + 1
    
    On Error Resume Next
    ReDim Preserve GridArray(nMaxCol + 1, nMaxRow)
    ReDim Preserve nCellChanged(nColumnCount, nMaxRow + 1)
    ReDim Preserve bValidCell(nColumnCount, nMaxRow + 1)

    fnGetNewBookmark = fnMakeBookmark(nMaxRow - 1)
    
End Function

'GetRelativeBookmark function - David
Private Function fnGetRelativeBookmark(bookm As Variant, RelPos As Integer) As Variant

    Dim Index As Long

    Index = fnIndexFromBookmark(bookm, RelPos < 0) + RelPos
    If Index < 0 Or Index >= nMaxRow Then
        fnGetRelativeBookmark = Null
    Else
        fnGetRelativeBookmark = fnMakeBookmark(Index)
    End If
    
End Function

'GetUserData function - David

Private Function fnGetUserData(bookm As Variant, col As Integer) As Variant

    Dim Index As Long

    Index = fnIndexFromBookmark(bookm, False)
    If Index < 0 Or Index >= nMaxRow Or col < 0 Or col >= nMaxCol Then
        fnGetUserData = Null
    Else
        If col >= 0 And col < nMaxCol Then
            If sDisplayFormat(col) <> "" Then
                fnGetUserData = Format$(GridArray(col, Index), sDisplayFormat(col))
            Else
                fnGetUserData = GridArray(col, Index) 'return the actual data
            End If
        Else
            fnGetUserData = GridArray(col, Index) 'return the actual data
        End If
    End If
    
End Function

'StoreUserData function - David
Private Function fnStoreUserData(bookm As Variant, Colm As Integer, userval As Variant) As Boolean

    Dim Index As Long
    Index = fnIndexFromBookmark(bookm, False)
    If Index < 0 Or Index > nMaxRow Or Colm < 0 Or Colm >= nMaxCol Then
        ' unable to write data so send a fail message
        fnStoreUserData = False
    Else
        ' storing data was successful
        fnStoreUserData = True
        GridArray(Colm, Index) = userval
    End If

End Function

'fnDeleteRow function - David
Private Function fnDeleteRow(bookm As Variant) As Boolean

    Dim Index As Long, vBookmark As Variant
    
    ' get the array index number to relate the bookmark to our data source
    Index = fnIndexFromBookmark(bookm, False) ' the second argument will not be necessary
    If Index < 0 Or Index >= nRowCount Then
        ' invalid row, delete is invalid
        fnDeleteRow = False
        Exit Function
    End If
        
    Dim i%, j%
    
    nRowCount = nRowCount - 1 ' decrease number of rows
    nMaxRow = nMaxRow - 1
    
    For i% = Index To nMaxRow - 1
        For j% = 0 To nMaxCol + 1
            GridArray(j%, i%) = GridArray(j%, i% + 1)
        Next j%
        For j = 0 To nColumnCount - 1
            nCellChanged(j, i) = nCellChanged(j, i + 1)
            bValidCell(j, i) = bValidCell(j, i + 1)
        Next j
    Next i%
    
    're-new the bookmark in bmChangedRows array is needed!
    'because all bookmarks in the GridArray will be re-sequence after rebind
    
    For i = 0 To nChangedCount - 1
        For Index = 0 To nMaxRow - 1
            If bmChangedRows(i) = GridArray(nBookMarkColumn, Index) Then
                bmChangedRows(i) = fnMakeBookmark(Index)
                ' Shift the data in the array, moving the rows up one
                nChangedCount = nChangedCount - 1
                For j = i To nChangedCount - 1
                    bmChangedRows(j) = bmChangedRows(j + 1)
                Next j
                Exit For
            End If
        Next Index
    Next i
    
    'Now re-new Bookmark Column
    For Index = 0 To nMaxRow - 1
        GridArray(nBookMarkColumn, Index) = fnMakeBookmark(Index)
    Next Index
    
    fnDeleteRow = True

End Function

Private Sub subNoSupport(sMsg As String)
    MsgBox "clsTGUnbound - " & sMsg & " is no longer supported"
End Sub
