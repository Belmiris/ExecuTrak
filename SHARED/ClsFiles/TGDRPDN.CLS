VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDropDown"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'***********************************************************'
'
' Copyright (c) 1996 FACTOR, A Division of W.R.Hess Company
'
' Module name   : TGDRPDN.CLS
' Date          : May 25, 1996
' Programmer(s) : Qinggang Ma
'
' This module defines a class implementing a drop down box using true grid.

' Usage:
'      i. Include class files COLUMN.CLS and FLTDRPDN.CLS in the project
'     ii. Define and create a class: clsFloatingDropDown
'    iii. Initialization (See sub fnInitializeDropDown in the form)
'     iv. Event calls:
'           2. In the dropdown true grid:
'                a. call .TableClick           in Click event
'                b. call .KeyPress  KeyAscii   in Keypress event
'                c. call .LostFocus            in LostFocus event
'                d. call .RowColChange         in RowColChange event
'                e. call .MouseUp Y            in MouseUp event
'                f. call .ReadData             in UnBoundReadData
'           3. In the command button
'                a. call .ButtonClick          in Click event

Option Explicit
    Private Const ASC_RETURN = 13

    Private CRLF As String * 2

    Private GridArray() As Variant
    Private tgDropDown As TDBGrid
    Private frmParent As Form
    Private txtBox As TextBox
    Private cmdSearch As Control
    Private nMaxRows As Integer
    Private nMaxCols As Integer
    Private bTableClicked As Boolean
    Private nClickPointY As Single
    Private bValidSelection As Boolean
    
    Private nOriginalTableWidth As Integer
    Private nOriginalTableHeight As Integer
    Private nFontSize As Integer

Public Sub AddRow(vData() As Variant)
    Dim nCol As Integer
    ReDim Preserve GridArray(nMaxCols - 1, nMaxRows)
    
    For nCol = 0 To nMaxCols - 1
        GridArray(nCol, nMaxRows) = vData(nCol)
    Next
    
    nMaxRows = nMaxRows + 1
End Sub

Private Function fnGetContainer(n_hWnd As Integer) As Object
    Dim i As Integer

    If frmParent.hWnd = n_hWnd Then
        Set fnGetContainer = Me
        Exit Function
    End If
    For i = 0 To frmParent.Controls.count - 1
        If TypeOf frmParent.Controls(i) Is VideoSoftElastic Or TypeOf frmParent.Controls(i) Is VideoSoftIndexTab Then
            If frmParent.Controls(i).hWnd = n_hWnd Then
                Set fnGetContainer = frmParent.Controls(i)
                Exit Function
            End If
        End If
    Next
End Function


Private Sub fnPositionTable()
    Const nDropDownRowNumbers = 8
    Const nScrollBarWidth = 252

    Dim nTbl_hWnd As Integer
    Dim n_hWnd As Integer
    Dim nLeft As Integer
    Dim nTop As Integer
    Dim objTemp As Object

    tgDropDown.Width = txtBox.Width + nScrollBarWidth
    
    nTbl_hWnd = GetParent(tgDropDown.hWnd)
    n_hWnd = GetParent(txtBox.hWnd)
    nLeft = txtBox.Left
    nTop = txtBox.Top + txtBox.Height
    While n_hWnd <> nTbl_hWnd And n_hWnd <> 0
        Set objTemp = fnGetContainer(n_hWnd)
        If TypeOf objTemp Is VideoSoftElastic Or TypeOf objTemp Is VideoSoftIndexTab Then
            If objTemp.Left > 0 And objTemp.Left < 2000 Then
                nLeft = nLeft + objTemp.Left
            End If
            If objTemp.Top > 0 And objTemp.Top < 2000 Then
                nTop = nTop + objTemp.Top
            End If
        End If
        n_hWnd = GetParent(n_hWnd)
    Wend
    tgDropDown.Top = nTop
    tgDropDown.Left = nLeft
    nOriginalTableHeight = tgDropDown.Height
    nOriginalTableWidth = tgDropDown.Width
End Sub


Property Set EditBox(tBox As Object)
    Set txtBox = tBox
    If Not tgDropDown Is Nothing Then
        fnPositionTable
    End If
End Property

Private Sub fnLoadTable()
    Const MaxVisibleRows = 5
    Const nSCROLLBAR As Integer = 252
    Dim nRows As Integer

    'Make our mouse pointer an hourglass
    Screen.MousePointer = vbHourglass
    bValidSelection = False
    
    nRows = nMaxRows
    If nRows > MaxVisibleRows Then
        nRows = MaxVisibleRows
    End If
    tgDropDown.Height = nRows * tgDropDown.RowHeight + 2 * Screen.TwipsPerPixelY
    tgDropDown.ZOrder 0
    tgDropDown.Visible = True
    
'    tgDropDown.Refresh
    tgDropDown.SetFocus

    'Reset our mouse pointer to the default
    Screen.MousePointer = vbDefault
End Sub

Public Sub Keypress(cntl As Object, KeyAscii As Integer)
    Const ASC_ESC = 27
    If cntl.TabIndex = tgDropDown.TabIndex Then
        If KeyAscii = ASC_RETURN Then
            fnGetText
        ElseIf KeyAscii = ASC_ESC Then
            txtBox.SetFocus
            tgDropDown.Visible = False
        End If
    ElseIf cntl.TabIndex = txtBox.TabIndex Then
        If KeyAscii = ASC_RETURN Then
            fnLoadTable
        End If
    End If
End Sub

Public Sub ReadData(ByVal RowBuf As RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
    Dim Bookmk As Variant
    Bookmk = StartLocation

    Dim RelPos As Integer
    If ReadPriorRows Then
        RelPos = -1
    Else
        RelPos = 1
    End If
    
    Dim RowsFetched As Integer
    RowsFetched = 0
    Dim i%, j%
    Dim lRow As Integer
    
    For i% = 0 To RowBuf.RowCount - 1
    
        Bookmk = fnGetRelativeBookmark(Bookmk, RelPos)
        If IsNull(Bookmk) Then
            Exit For
        Else
            If Bookmk = "" Then Exit For
        End If
        For j% = 0 To RowBuf.ColumnCount - 1
            lRow = fnIndexFromBookmark(Bookmk, False)
            If lRow >= 0 And lRow < nMaxRows Then
                RowBuf.Value(i%, j%) = GridArray(j%, lRow)
            End If
        Next j%
        RowBuf.bookmark(i%) = Bookmk
        RowsFetched = RowsFetched + 1
    Next i%
    RowBuf.RowCount = RowsFetched
End Sub

Private Function fnGetRelativeBookmark(Bookmk As Variant, RelPos As Integer) As Variant

    Dim Index As Long
    Index = fnIndexFromBookmark(Bookmk, RelPos < 0) + RelPos
    If Index < 0 Or Index >= nMaxRows Then
        fnGetRelativeBookmark = Null ' invalid row (e.g., EOF or BOF)
    Else
        fnGetRelativeBookmark = fnMakeBookmark(Index)
    End If
End Function

Private Function fnIndexFromBookmark(Bookmk As Variant, ReadPriorRows As Boolean) As Long
    If IsNull(Bookmk) Then
        If ReadPriorRows = True Then  ' the grid is asking for rows prior to Bookmk
            fnIndexFromBookmark = nMaxRows
        Else  ' the grid is asking for rows after Bookmk
            fnIndexFromBookmark = -1
        End If
    Else
        fnIndexFromBookmark = val(Bookmk)
    End If
End Function


Private Function fnMakeBookmark(Index As Long) As Variant
    fnMakeBookmark = Str$(Index)
End Function

Public Sub Resize()
    Dim fFactor As Single
    Dim nTW As Integer
    Dim i As Integer
    fFactor = frmParent.Height / frmParent.OriginalHeight
    tgDropDown.Height = nOriginalTableHeight * fFactor
    tgDropDown.Font.Size = nFontSize * fFactor
    tgDropDown.HeadFont.Size = nFontSize * fFactor
    
    fFactor = frmParent.Width / frmParent.OriginalWidth
    nTW = tgDropDown.Width
    tgDropDown.Width = nOriginalTableWidth * fFactor
    fFactor = tgDropDown.Width / nTW
    For i = 0 To tgDropDown.Columns.count - 1
        tgDropDown.Columns(i).Width = fFactor * tgDropDown.Columns(i).Width
    Next i
End Sub



Property Set DropDownTable(tblGrid As Object)
    
    Const szROW_STYLE As String = "RowStyle"
    Const szCURRENT_CELL_STYLE As String = "CurrCellStyle"
    
    Dim szBuf As String
    Dim n As Integer
    Dim m_styRow As Object            ' row style
    Dim m_styCurrentCell As Object    ' current cell style
    
    On Error GoTo AddToCollectionError
    
    Set tgDropDown = tblGrid
    nMaxCols = tblGrid.Columns.count
    nFontSize = tblGrid.Font.Size
    
    ' set some properties for the table
    Set m_styRow = tblGrid.Styles.Add(szROW_STYLE)
    Set m_styCurrentCell = tblGrid.Styles.Add(szCURRENT_CELL_STYLE)
    
    ' create current cell style
    With m_styCurrentCell
        .BackColor = &HFF0000
        .ForeColor = &HFFFFFF
    End With
    
    ' create highlight row style
    With m_styRow
        .BackColor = &HFF0000
        .ForeColor = &HFFFFFF
    End With
    
    For n = 0 To tblGrid.Columns.count - 1
        With tblGrid.Columns(n)
            .DividerStyle = dbgBlackLine
            .Locked = True
            .AllowSizing = False
            .Visible = True
            .WrapText = False
            .Button = False
            .FetchStyle = False
        End With
    Next
  
    ' add the new styles to the table
    tblGrid.AddCellStyle dbgMarqueeRow, m_styRow
    tblGrid.AddCellStyle dbgMarqueeRow + dbgCurrentCell, m_styCurrentCell
    
    If Not txtBox Is Nothing Then
        fnPositionTable
    End If

  ' set default properties
    With tblGrid
        .AllowAddNew = False
        .AllowDelete = False
        .AllowUpdate = False
        .Enabled = True
        .DataMode = dbgUnbound
        .MarqueeUnique = True
        .EditDropDown = False
        .RowDividerStyle = dbgBlackLine
        .BackColor = &HFFFFFF
        .ForeColor = &H0
        .HeadLines = 1
        .Appearance = dbg3D
        .BorderStyle = dbgFixedSingle
        .AllowArrows = True
        .TabAcrossSplits = False
        .WrapCellPointer = False
        .ExposeCellMode = dbgScrollOnSelect
        .TabAction = dbgControlNavigation
        .Splits(0).Locked = True
        .Splits(0).AllowFocus = True
        .Splits(0).AllowSizing = False
        .AllowRowSizing = False
        .ExtendRightColumn = True
        .ScrollBars = dbgAutomatic
        .MarqueeStyle = dbgHighlightRow
        .RecordSelectors = False
        .ColumnHeaders = False
        .Visible = False
    End With
    
    On Error GoTo 0
    Exit Property
    
AddToCollectionError:
    
    On Error GoTo 0

End Property

Public Sub ButtonClick(ParamArray vMoreFields() As Variant)
    fnLoadTable
End Sub


Private Sub fnGetText()
    Dim nPos As Integer
    Dim sTemp As String
    sTemp = Trim$(tgDropDown.Columns(0).Text)
    nPos = InStr(sTemp, "(")
    If nPos > 1 Then
        sTemp = Trim$(Left$(sTemp, nPos - 1))
    End If
    txtBox.Text = sTemp
    txtBox.SetFocus
    tgDropDown.Visible = False
    bValidSelection = True
    tgDropDown.ZOrder 1
End Sub

Public Sub LostFocus()
    txtBox.SetFocus
    tgDropDown.Visible = False
End Sub

Public Sub MouseUp(Y As Single)
    nClickPointY = Y
End Sub

Property Set Form(frmMain As Object)
    Set frmParent = frmMain
End Property

Property Set Button(btn As Object)
    Set cmdSearch = btn
    btn.Picture = LoadResPicture(SEARCH_UP, vbResBitmap)
End Property

Public Sub RowColChange()
'    Dim nTCol As Integer
'    Dim nMCol As Integer
'    nTCol = tgMain.col
'    nMCol = fnColumnNumber(nTCol)
    Dim lDropDown As Long
    Dim lCol As Long
    
    If bTableClicked Then
        If lDropDown >= 0 Then
            fnGetText
        End If
        bTableClicked = False
    End If
End Sub
Public Sub SetData(vData() As Variant)
    Dim nCol As Integer
    Dim nRow As Integer
    nMaxRows = UBound(vData, 2) + 1
    
    ReDim GridArray(nMaxCols - 1, nMaxRows - 1)
    
    For nRow = 0 To nMaxRows - 1
        For nCol = 0 To nMaxCols - 1
            GridArray(nCol, nRow) = vData(nCol, nRow)
        Next nCol
    Next nRow
End Sub

Public Sub TableClick()
    Dim nRowTop As Integer
    Dim nRowBtm As Integer
    If tgDropDown.Row >= 0 Then
        nRowTop = tgDropDown.RowTop(tgDropDown.Row)
        nRowBtm = nRowTop + tgDropDown.RowHeight
        If nClickPointY >= nRowTop And nClickPointY <= nRowBtm Then
            fnGetText
        Else
            bTableClicked = True
        End If
    Else
        bTableClicked = True
    End If
End Sub

Property Get ValidSelection() As Boolean
    ValidSelection = bValidSelection
End Property
Property Let ValidSelection(bStatus As Boolean)
    bValidSelection = bStatus
End Property

Private Sub Class_Initialize()
    CRLF = Chr(13) + Chr(10)
End Sub

Private Sub Class_Terminate()
    Set tgDropDown = Nothing
    Set frmParent = Nothing
    Set txtBox = Nothing
    Set cmdSearch = Nothing
End Sub

