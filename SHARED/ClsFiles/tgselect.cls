VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTGSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***********************************************************'
'
' Copyright (c) 1996 FACTOR, A Division of W.R.Hess Company
'
' Module name   : TGSELECT.CLS
' Date          : November 16, 1996
' Programmer(s) : Qinggang Ma
'
' This module defines a class implementing select/unselect items
'    in a true grid table

'Usage:
'   Setup properties:
'       1. Spreadsheet class : EditorClass as clsTGSpreadSheet (Required)
'       2. Selector class    : Partner clsTGSelector (Optional)
'   Example:
'       Set tgcSelector = New clsTGSelector
'       Set tgcSelector.EditorClass = tgcSpreadSheet
'       Set tgcSelector.Partner = tgcSelector1
'
'   Event calls in the true grid:
'       1. Click        :   Click
'       2. GotFocus     :   GotFocus
'       3. KeyDown      :   KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
'       4. LostFocus    :   LostFocus
'       5. MousUp       :   MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal Y As Integer)
'       6. RowColChange :   RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'       7. SelChange    :   SelChange(Cancel As Integer)
'
'   Auxilary functions and properties:
'       1. Clear()(Sub):
'           Clear all the selected rows
'       2. Copy()(Sub )
'           Copy selected items to another grid controlled by tgcDest
'           Require to setup the Partner property
'       3. Count() (Property Get) :
'           Returns the number of selected rows
'       4. Delete() (Sub)
'           Delete selected rows
'       5. GetSelected(lTemp() As Long, nCount As Long) (Sub)
'           Returns the row numbers selected in the grid
'       7. RowHighLighted() (Property Let and Get)
'           High light (or not) current row
'       6. SelectAll() (Sub)
'           Selects all the rows in the grid

Option Explicit
    
Private Const MODE_KEY_CONTROL = 1
Private Const MODE_NO_KEY_CONTROL = 2
Private Const MyShiftMask = -&H1000

Private Const szROW_STYLE As String = "RowStyle"
Private Const szCURRENT_CELL_STYLE As String = "CurrCellStyle"

Private m_tgTable As Object
Private m_tgcEditor As clsTGSpreadSheet
Private m_tgcSelector As clsTGSelector

Private m_bTableGotFocus As Boolean
Private m_bTableClicked As Boolean
Private m_nClickPointY As Integer
Private m_vLastRow As Variant

Private m_nIdleColumn As Integer
Private m_nSelectMode As Integer
Private m_nShiftControl As Integer
Private m_nLastClickedRow As Integer
Private m_bRowHighLighted As Boolean
Private m_bAvoidBeep As Boolean
Private m_bSelectCurrRow As Boolean
Private m_vLastSelected As Variant
'

Property Let AllowMultipleSelect(bTemp As Boolean)

    If bTemp Then
        m_nSelectMode = MODE_KEY_CONTROL
    Else
        m_nSelectMode = MODE_NO_KEY_CONTROL
    End If
    
End Property

Property Let AvoidBeep(ByVal bTemp As Boolean)
    m_bAvoidBeep = bTemp
End Property

Public Sub Clear()
    With m_tgTable.SelBookmarks
        While .Count > 0
            .Remove 0
        Wend
    End With
End Sub

Public Sub Copy(Optional vSelIndex As Variant)
    If m_tgcSelector Is Nothing Then
        Exit Sub
    End If
    
    Dim nRows As Long
    Dim bmTemp As SelBookmarks
    Dim tgcDest As Object
    
    Set tgcDest = m_tgcSelector.EditorClass
    Set bmTemp = m_tgcEditor.Table.SelBookmarks
    nRows = bmTemp.Count
    If nRows = 0 Then
        Exit Sub
    End If

    Dim i As Long
    Dim j As Integer
    Dim vBuffer() As Variant
    Dim nCol As Integer
    Dim nColCount As Integer
    Dim vData() As Variant
    
    nColCount = m_tgcEditor.Table.Columns.Count
    ReDim vData(nColCount - 1, nRows - 1)
    
    If IsMissing(vSelIndex) Then
        'store the selected bookmarks
        For i = 0 To nRows - 1
            j = Val(bmTemp.Item(i))
            m_tgcEditor.GetRow vBuffer, nCol, j
            For j = 0 To nColCount - 1
                If j <> m_nIdleColumn Then
                    vData(j, i) = vBuffer(j)
                End If
            Next j
        Next i
        tgcDest.FillWithArray vData, False
    Else
        If IsNumeric(vSelIndex) Then
            If vSelIndex >= 0 And vSelIndex < nRows Then
                i = Val(bmTemp.Item(vSelIndex))
                ReDim vData(nColCount - 1, 0)
                m_tgcEditor.GetRow vBuffer, nCol, i
                For j = 0 To nColCount - 1
                    If j <> m_nIdleColumn Then
                        vData(j, 0) = vBuffer(j)
                    End If
                Next j
                tgcDest.FillWithArray vData, False
            End If
        End If
    End If

End Sub

Public Sub Delete(Optional vRowIndex As Variant)
    If Count > 0 Then
        Dim i As Long
        If IsMissing(vRowIndex) Then
            Dim nRow As Long
            ReDim vBkMark(Count - 1) As Variant
            For i = 0 To Count - 1
                nRow = Val(m_tgTable.SelBookmarks.Item(i))
                vBkMark(i) = m_tgcEditor.Bookmark(nRow)
            Next
            For i = Count - 1 To 0 Step -1
                m_tgcEditor.DeleteRow vBkMark(i), False
            Next
            m_tgcEditor.Rebind
        Else
            If IsNumeric(vRowIndex) Then
                i = Val(vRowIndex)
                If i >= 0 And i < Count Then
                    nRow = Val(m_tgTable.SelBookmarks.Item(i))
                    m_tgcEditor.DeleteRow m_tgcEditor.Bookmark(nRow), False
                End If
            End If
        End If
    End If
End Sub

Property Set EditorClass(clsTemp As Object)
    Set m_tgcEditor = clsTemp
    Set Table = m_tgcEditor.Table
End Property

Property Get EditorClass() As Object
    Set EditorClass = m_tgcEditor
End Property


Public Sub Click()
    If m_bTableGotFocus Then
        If m_tgTable.Row = m_tgTable.RowContaining(m_nClickPointY) Then
            subMarkSelected m_vLastRow
            subSetCurrentRow
            If Not m_tgTable.EditActive Then
                SetRowHighLighted m_bRowHighLighted
            End If
        Else
            m_bTableClicked = True
        End If
        If m_bAvoidBeep Then
            If m_nIdleColumn >= 0 Then
                If m_nIdleColumn >= m_tgTable.Columns.Count Then
                    m_nIdleColumn = m_tgTable.Columns.Count - 1
                End If
                If m_tgTable.Columns(m_nIdleColumn).Visible Then
                    If m_tgTable.Columns.Count > 0 Then
                        m_nIdleColumn = m_tgTable.Columns.Count - 1
                        If m_tgTable.Columns(m_nIdleColumn).Visible Then
                            m_nIdleColumn = m_nIdleColumn + 1
                            m_tgTable.Columns.Add m_nIdleColumn
                            m_tgTable.Columns(m_nIdleColumn).Width = 50
                        End If
                    End If
                End If
                m_tgTable.col = m_nIdleColumn
            End If
        End If
    End If
End Sub

Private Function fnLong(vTemp As Variant) As Long
    If IsNull(vTemp) Then
        fnLong = -1
    Else
        fnLong = Val(vTemp)
    End If
End Function

Private Function fnMarkCurrent(ByVal KeyCode As Integer, _
                               ByVal Shift As Integer) As Integer
    Dim nRow As Integer
    Dim vCurrent As Variant
    Dim nCurrent As Integer
    Dim nLast As Integer
    Dim nStart As Integer
    Dim nEnd As Integer
    Dim i As Integer

    nRow = m_tgcEditor.GetCurrentRowNumber
    If KeyCode = vbKeyDown Then
        If nRow >= m_tgcEditor.RowCount Then
            Exit Function
        End If
        nRow = nRow + 1
    ElseIf KeyCode = vbKeyUp Then
        If nRow <= 0 Then
            Exit Function
        End If
        nRow = nRow - 1
    ElseIf KeyCode = vbKeyPageUp Then
        If nRow <= 0 Then
            Exit Function
        End If
        nRow = nRow - m_tgTable.VisibleRows + 1
        If nRow < 0 Then
            nRow = 0
        End If
    ElseIf KeyCode = vbKeyPageDown Then
        If nRow >= m_tgcEditor.RowCount Then
            Exit Function
        End If
        nRow = nRow + m_tgTable.VisibleRows - 1
        If nRow >= m_tgcEditor.RowCount Then
            nRow = m_tgcEditor.RowCount - 1
        End If
    Else
        fnMarkCurrent = -1
        Exit Function
    End If
    If nRow < m_tgcEditor.RowCount Then
        vCurrent = m_tgcEditor.MakeBookmark(nRow)
    '        If Shift = MyShiftMask Then
        If Shift = vbShiftMask Then
            nCurrent = fnRowSelected(vCurrent)
            nLast = fnRowSelected(m_vLastRow)
            If nLast >= 0 Then
                If nCurrent >= 0 Then
                    If nCurrent < nLast Then
                        nStart = nCurrent + 1
                        nEnd = nLast
                    Else
                        nStart = nLast
                        nEnd = nCurrent - 1
                    End If
                    For i = nEnd To nStart Step -1
                        m_tgTable.SelBookmarks.Remove i
                    Next
                Else
                    nCurrent = Val(vCurrent)
                    nLast = Val(m_vLastRow)
                    If nCurrent < nLast Then
                        nStart = nCurrent
                        nEnd = nLast - 1
                    Else
                        nEnd = nCurrent
                        nStart = nLast + 1
                    End If
                    If nStart < 0 Then
                        nStart = 0
                    End If
                    For i = nStart To nEnd
                        m_tgTable.SelBookmarks.Add m_tgcEditor.MakeBookmark(i)
                    Next
                End If
            End If
        Else
            subMarkSelected vCurrent
            fnMarkCurrent = nRow
        End If
    End If
End Function


Private Function fnRowSelected(vBkMark As Variant) As Integer
    Dim bFound As Boolean
    Dim i As Integer

    fnRowSelected = -1
    With m_tgTable
        For i = .SelBookmarks.Count - 1 To 0 Step -1
            If vBkMark = .SelBookmarks.Item(i) Then
                fnRowSelected = i
                Exit For
            End If
        Next
    End With

End Function

Public Sub GetSelected(lTemp() As Long, nCount As Long)
    'Return the row numbers of selected rows

    Dim i As Long
    
    With m_tgTable.SelBookmarks
        nCount = .Count
        If nCount > 0 Then
            ReDim lTemp(nCount - 1)
            For i = 0 To nCount - 1
                lTemp(i) = Val(.Item(i))
            Next i
        End If
    End With

End Sub

Public Sub GotFocus()
    m_bTableGotFocus = True
    If Not m_tgcEditor Is Nothing Then
        If m_tgcEditor.RowCount > 0 Then
            SetRowHighLighted m_bRowHighLighted
        End If
    End If
End Sub

Property Let IdleColumn(nTemp As Integer)
    m_nIdleColumn = nTemp
End Property

Public Sub KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    
    If m_nSelectMode = MODE_NO_KEY_CONTROL Then
        If KeyCode = vbKeyReturn Then
            m_nShiftControl = Shift
            subMarkSelected m_vLastRow
        End If
    ElseIf m_nSelectMode = MODE_KEY_CONTROL Then
        m_nShiftControl = Shift  'MyShiftMask
        If Shift = vbShiftMask Then
            If m_bSelectCurrRow Then
                If Count > 0 Then
                    fnMarkCurrent KeyCode, m_nShiftControl
                End If
            Else
                If Count = 0 Then
                    If m_tgcEditor.GetCurrentRowNumber >= 0 Then
                        m_tgTable.SelBookmarks.Add m_tgcEditor.MakeBookmark(m_tgcEditor.GetCurrentRowNumber)
                    End If
                Else
                    fnMarkCurrent KeyCode, m_nShiftControl
                End If
            End If
            'subSetCurrentRow
        Else
            'david 02/01/2001
            'david 03/11/2002 add 'And KeyCode <> vbKeyTab'
            If Shift = 0 And KeyCode <> 0 And KeyCode <> vbKeyTab Then
                Dim i As Integer
                
                'david 10/19/2001
                If m_tgTable.SelBookmarks.Count > 1 Then
                    For i = m_tgTable.SelBookmarks.Count - 1 To 0 Step -1
                        m_tgTable.SelBookmarks.Remove i
                    Next i
                'david 01/22/2002
                ElseIf m_tgTable.SelBookmarks.Count = 1 And KeyCode = vbKeyReturn Then
                    m_tgTable.SelBookmarks.Remove 0
                End If
                
                m_nLastClickedRow = fnMarkCurrent(KeyCode, Shift)
                '''''''''''''''''
            ElseIf (Shift And vbCtrlMask) > 0 And (KeyCode = 65 Or KeyCode = 97) And (Not m_tgTable.EditActive) Then
                SelectAll
                Shift = 0
                KeyCode = 0
                
                Exit Sub
            Else
                m_nLastClickedRow = fnMarkCurrent(KeyCode, Shift)
            End If
        End If
    End If
    If KeyCode = vbKeyDown Then
        If Not m_tgcEditor Is Nothing Then
            If m_tgcEditor.GetCurrentRowNumber >= m_tgcEditor.RowCount - 1 Then
                KeyCode = 0
                Exit Sub
            End If
        End If
    End If
End Sub


Public Sub LostFocus()
    If m_nSelectMode = MODE_NO_KEY_CONTROL Then
        m_bTableGotFocus = False
    End If
    m_bTableClicked = False
    If Not m_tgcEditor Is Nothing Then
        SetRowHighLighted False
    End If
End Sub

Public Sub MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal y As Integer)
    If Button = vbLeftButton Then
        If Shift <> vbShiftMask Then
            m_nLastClickedRow = -1
        End If
        m_nShiftControl = Shift
        m_nClickPointY = y
        DoEvents
    End If
End Sub


Property Set Partner(tgcTemp As Object)
    Set m_tgcSelector = tgcTemp
End Property

Property Get RowIsSelected(ByVal lRow As Long) As Boolean

    Dim i As Long
    
    For i = 0 To m_tgTable.SelBookmarks.Count - 1
        If lRow = Val(m_tgTable.SelBookmarks.Item(i)) Then
            RowIsSelected = True
            Exit Property
        End If
    Next i
    RowIsSelected = False
    
End Property

Property Let RowIsSelected(ByVal lRow As Long, bSelect As Boolean)

    Dim i As Long
    
    If bSelect Then
        If Not RowIsSelected(lRow) Then
            If lRow >= 0 Then
                m_tgTable.SelBookmarks.Add m_tgcEditor.MakeBookmark(lRow)
            End If
        End If
    Else
        If RowIsSelected(lRow) Then
            For i = 0 To m_tgTable.SelBookmarks.Count - 1
                If lRow = Val(m_tgTable.SelBookmarks.Item(i)) Then
                    m_tgTable.SelBookmarks.Remove i
                    Exit Property
                End If
            Next i
        End If
    End If
    
End Property


Public Sub SelChange(Cancel As Integer)
    Cancel = True
'    If Not m_tgcSelector Is Nothing Then
'        m_tgcSelector.Clear
'    End If
End Sub


Public Sub SelectAll()

    'changed for two variables from integer type to long by junsong Oct. 9, 2000
    Dim i As Long
    Dim nRows As Long
    
    If Not m_tgcEditor Is Nothing Then
        nRows = m_tgcEditor.RowCount
    End If
    With m_tgTable.SelBookmarks
        For i = 0 To nRows - 1
            .Add m_tgcEditor.MakeBookmark(i)
            SetRowHighLighted m_bRowHighLighted
        Next
    End With

End Sub

Property Get Count() As Long
    Count = m_tgTable.SelBookmarks.Count
End Property


Property Let RowHighLighted(bStatus As Boolean)
    m_bRowHighLighted = bStatus
    SetRowHighLighted bStatus
End Property

Property Get RowHighLighted() As Boolean
    RowHighLighted = m_bRowHighLighted
End Property

Property Let SelectCurrRow(ByVal bTemp As Boolean)
    m_bSelectCurrRow = bTemp
End Property

Public Sub SelectRow(ByVal lRow As Long)
    subMarkSelected m_tgcEditor.MakeBookmark(lRow)
End Sub

Private Sub subMarkSelected(vBkMark As Variant)
    If m_tgcEditor.RowCount <= 0 Then
        Exit Sub
    End If
    Dim i As Integer
    Dim nEnd As Integer
    Dim nStart As Integer
    Dim bFound As Boolean

    If m_nSelectMode = MODE_KEY_CONTROL Then
        nStart = m_nLastClickedRow
        If m_tgcEditor Is Nothing Then
            nEnd = m_tgTable.Row
        Else
            nEnd = m_tgcEditor.GetCurrentRowNumber
        End If
        If m_nShiftControl = vbShiftMask Then
            Clear
            If nStart > nEnd Then
                i = nStart
                nStart = nEnd
                nEnd = i
            End If
            If nStart < 0 Then
                nStart = 0
            End If
            For i = nStart To nEnd
                m_tgTable.SelBookmarks.Add m_tgcEditor.MakeBookmark(i)
            Next
            'm_nLastClickedRow = -1
        ElseIf m_nShiftControl = vbCtrlMask Then
            subToggleSelected vBkMark
'        ElseIf m_nShiftControl = MyShiftMask Then
'            If nStart = nEnd Then
'                Clear
'                m_tgTable.SelBookmarks.Add vBkMark
'            End If
'            subToggleSelected vBkMark
'            Stop
        Else
            Clear
            If m_bSelectCurrRow Then
                If fnLong(vBkMark) >= 0 Then
                    m_tgTable.SelBookmarks.Add vBkMark
                End If
            End If
            
'            subToggleSelected vBkMark
'            If nStart = nEnd Then
'                If Count > 0 Then
'                    Clear
'                Else
'                    m_tgTable.SelBookmarks.Add vBkMark
'                End If
'            Else
'                Clear
'                m_tgTable.SelBookmarks.Add vBkMark
'            End If
        End If
    Else
        Clear
        subToggleSelected vBkMark
    End If
End Sub

Public Sub RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If m_tgcEditor.RowCount > 0 Then
        m_vLastRow = m_tgTable.Bookmark
        If m_nSelectMode = MODE_KEY_CONTROL Then
            If m_bTableClicked Then
                subMarkSelected m_vLastRow
                m_bTableClicked = False
                If m_nShiftControl <> vbShiftMask Then
                    subSetCurrentRow
                End If
            End If
        Else
            subMarkSelected m_vLastRow
        End If
        m_vLastSelected = m_tgcEditor.Bookmark
    End If

    SetRowHighLighted m_bRowHighLighted
End Sub


Private Sub subSetCurrentRow()
'    If m_nLastClickedRow < 0 Then
        If m_tgcEditor Is Nothing Then
            m_nLastClickedRow = m_tgTable.Row
        Else
            m_nLastClickedRow = m_tgcEditor.GetCurrentRowNumber
        End If
'    End If
End Sub

Private Sub subToggleSelected(vBkMark As Variant)
    Dim i As Integer

    i = fnRowSelected(vBkMark)
    With m_tgTable
        If i >= 0 Then
            .SelBookmarks.Remove i
        Else
            If fnLong(vBkMark) >= 0 Then
                .SelBookmarks.Add vBkMark
            End If
        End If
    End With

End Sub


Property Set Table(objTemp As Object)
    Set m_tgTable = objTemp
    If m_bAvoidBeep Then
        With m_tgTable.Columns
            If .Item(.Count - 1).Visible Then
                m_nIdleColumn = .Count
                .Add .Count
                .Item(m_nIdleColumn).Width = 50
            End If
        End With
    End If
End Property

Property Get Temp() As Integer
    Temp = m_nShiftControl
End Property

Private Sub SetRowHighLighted(bHighLighted As Boolean)
    Dim m_styRow As Style  ' row style
    Dim m_styCurrentCell As Style   ' current cell style
    Dim m_styHighLightedRow As Style
    
    m_tgTable.ClearCellStyle dbgAllCells
    
    ' set some properties for the table
    Set m_styRow = m_tgTable.Styles.Item(szROW_STYLE)
    Set m_styHighLightedRow = m_tgTable.Styles.Item(szROW_STYLE)
    Set m_styCurrentCell = m_tgTable.Styles.Item(szCURRENT_CELL_STYLE)
    
    If bHighLighted Then
        ' create highlight row style
        With m_styRow
            .BackColor = &HFF0000
            .ForeColor = &HFFFFFF
        End With

        ' create current cell style
        If m_tgcEditor.EditColumnCount > 0 Then
            With m_styCurrentCell
                .BackColor = &HFFFFFF
                .ForeColor = &H0
            End With
        Else
            Set m_styCurrentCell = m_styRow
        End If
    
        If m_tgcEditor.EditColumnCount > 0 Then
            If m_tgTable.SelBookmarks.Count > 0 Then
                With m_styHighLightedRow
                    .BackColor = m_tgTable.SelectedBackColor
                    .ForeColor = &HFFFFFF
                End With
                
                m_tgTable.AddCellStyle dbgSelectedRow, m_styHighLightedRow
                m_tgTable.MarqueeStyle = dbgNoMarquee
            Else
                m_tgTable.MarqueeStyle = dbgHighlightRowRaiseCell
            End If
        Else
            m_tgTable.MarqueeStyle = dbgHighlightRow
        End If
    Else
        ' create plain row style
        With m_styRow
            .BackColor = &HFFFFFF
            .ForeColor = &H0
        End With
        m_tgTable.MarqueeStyle = dbgNoMarquee
    End If

    'add the new styles to the table
    m_tgTable.AddCellStyle dbgMarqueeRow, m_styRow
    m_tgTable.AddCellStyle dbgMarqueeRow + dbgCurrentCell, m_styCurrentCell
    m_tgTable.AddCellStyle dbgMarqueeRow + dbgUpdatedCell + dbgCurrentCell, m_styCurrentCell
    m_tgTable.AddCellStyle dbgMarqueeRow + dbgUpdatedCell, m_styRow
    m_tgTable.AddCellStyle dbgUpdatedCell, m_styCurrentCell
End Sub

Private Sub Class_Initialize()
    m_nSelectMode = MODE_KEY_CONTROL
    m_bTableGotFocus = True
    m_nLastClickedRow = -1
    m_nIdleColumn = -1
    m_bAvoidBeep = True
    m_bSelectCurrRow = True
End Sub

Private Sub Class_Terminate()
    Set m_tgTable = Nothing
    Set m_tgcEditor = Nothing
    Set m_tgcSelector = Nothing
End Sub


