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

Private Const szROW_STYLE As String = "RowStyle"
Private Const szCURRENT_CELL_STYLE As String = "CurrCellStyle"

Private m_tgTable As Object
Private m_tgcEditor As clsTGSpreadSheet
Private m_tgcSelector As clsTGSelector

Private m_bTableGotFocus As Boolean
Private m_bTableClicked As Boolean
Private m_nClickPointY As Single
Private m_vLastRow As Variant

Private m_nIdleColumn As Integer
Private m_nSelectMode As Integer
Private m_nShiftControl As Integer
Private m_nLastClickedRow As Integer
Private m_bRowHighLighted As Boolean
Private m_bAvoidBeep As Boolean
Private m_bSelectCurrRow As Boolean
Private m_vLastSelected As Variant

'david 02/09/2004  #401172-1
Private m_bMultiSelectRowFeature As Boolean

Private m_objCellColor As Object

Private m_colMultiSelCellColumns As Collection
Private m_colMultiSelectedCells As Collection

Private m_nClickPointX As Single
Private m_nClickCol As Integer
Private m_sLastClickedCell As String
''''''''''''''''''''''''''''
'

Property Set EditorClass(clsTemp As Object)
    Set m_tgcEditor = clsTemp
    Set Table = m_tgcEditor.Table
    
    Set m_tgcEditor.SelectorClass = Me
    
    'david 02/09/2004  #401172-1
    'turn off the FetchCell
    Dim i As Integer
    
    If m_objCellColor Is Nothing Then
        For i = 0 To m_tgTable.Columns.Count - 1
            If Not m_tgcEditor.IsEditableCol(i) Then
                m_tgTable.Columns(i).FetchStyle = False
            End If
        Next i
    Else
        For i = 0 To m_tgTable.Columns.Count - 1
            If m_objCellColor.IsColumnForCellColor(i) Then
                m_tgTable.Columns(i).FetchStyle = False
            End If
        Next i
    End If
    ''''''''''''''''''''''''''''
End Property

Property Get EditorClass() As Object
    Set EditorClass = m_tgcEditor
End Property

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
            
            'david 09/17/2003  #412827-4
            subSortArray vBkMark
            ''''''''''''''''''''''''''''
            
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

Public Sub Click()
    'david 02/09/2004  #401172-1
    m_nClickCol = m_tgTable.ColContaining(m_nClickPointX)
    
    If IsMultiSelectCellColumn(m_nClickCol) Then
        If m_bTableGotFocus Then
            Clear
            
            If IsIDE() Then
                If m_nClickPointY < m_tgTable.RowTop(0) Then
                    'header click, do nothing
                Else
                    If m_tgTable.Row = m_tgTable.RowContaining(m_nClickPointY) Then
                        If m_tgcEditor.RowCount > 0 Then
                            Dim sColRow As String
                            
                            sColRow = fnMakeKey(m_nClickCol, m_tgcEditor.GetCurrentRowNumber)
                            
                            If Not m_tgTable.EditActive Then
                                subMarkCellSelected sColRow
                            End If
                        
                            subSetLastClickedCell sColRow
                        End If
                    Else
                        m_bTableClicked = True
                    End If
                End If
            Else
                If m_nClickPointY < m_tgTable.RowTop(0) Then
                    'header click, do nothing
                Else
                    m_bTableClicked = True
                End If
            End If
        End If
    ''''''''''''''''''''''''''''
    Else
        If m_bTableGotFocus Then
            '02/09/2004
            'when the selected cell count is greater than zero (one or more cell is selected)
            'we need to do the rebind to clear the selected cell(s)
            'note:when doing rebind, one click cannot get into cell edit mode
            If SelectedCellCount > 0 Then
                ClearSelectedCells True
            Else
                ClearSelectedCells
            End If
            '''''''''''
            
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
                    m_tgTable.Col = m_nIdleColumn
                End If
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
    Dim lRow As Long
    Dim vCurrent As Variant
    Dim nCurrent As Integer
    Dim nLast As Integer
    Dim nStart As Integer
    Dim nEnd As Integer
    Dim i As Integer

    Dim bShiftDown As Boolean
    Dim bAltDown As Boolean
    Dim bCtrlDown As Boolean
    
    bShiftDown = (Shift And vbShiftMask) > 0
    bAltDown = (Shift And vbAltMask) > 0
    bCtrlDown = (Shift And vbCtrlMask) > 0
     
    lRow = m_tgcEditor.GetCurrentRowNumber

    'david 12/10/2002
    'new codes
    If bShiftDown And bCtrlDown And bAltDown Then
        'do nothing
    ElseIf bShiftDown And bAltDown Then
        'do nothing
    ElseIf bShiftDown And bCtrlDown Then
        If KeyCode = vbKeyHome Or KeyCode = vbKeyPageUp Then
            lRow = 0
            m_tgTable.Bookmark = m_tgcEditor.MakeBookmark(lRow)
        ElseIf KeyCode = vbKeyEnd Or KeyCode = vbKeyPageDown Then
            lRow = m_tgcEditor.RowCount - 1
            m_tgTable.Bookmark = m_tgcEditor.MakeBookmark(lRow)
        End If
    ElseIf bCtrlDown And bAltDown Then
        'do nothing
    ElseIf bShiftDown Or (bShiftDown = 0 And (Not bCtrlDown)) Then
        If KeyCode = vbKeyDown Then
            'david 03/26/2004
            'when allow add new
            'fix to allow the highlighting to go to the last new row
            'when the last row is valid
            'If lRow < m_tgcEditor.RowCount - 1 Then
            If lRow <= m_tgcEditor.RowCount - 1 Then
                If m_tgcEditor.GotoNextRow(False) Then
                    lRow = lRow + 1
                End If
            'Else
            '    lRow = m_tgcEditor.RowCount - 1
            End If
            '''''''''''''''''
        ElseIf KeyCode = vbKeyUp Then
            If lRow > 0 Then
                lRow = lRow - 1
            Else
                lRow = 0
            End If
        ElseIf KeyCode = vbKeyPageUp Then
            If lRow <= 0 Then
                Exit Function
            End If
            
            'david 12/11/2002
            'fixed pageup highlighting row
            'lRow = lRow - m_tgTable.VisibleRows + 1
            lRow = lRow - m_tgTable.VisibleRows + 2
            '''''''''''''''''
            
            If lRow < 0 Then
                lRow = 0
            End If
        ElseIf KeyCode = vbKeyPageDown Then
            If lRow >= m_tgcEditor.RowCount Then
                Exit Function
            End If
            
            lRow = lRow + m_tgTable.VisibleRows - 1
            
            If lRow >= m_tgcEditor.RowCount Then
                lRow = m_tgcEditor.RowCount - 1
            End If
        
        'david 12/10/2002
        'fixed the highlighting problem
        'when focus is in last editable cell and right arrow key is pressed
        'or when focus is in first editable cell and left arrow key is pressed
        ElseIf KeyCode = vbKeyLeft And m_tgTable.Col = m_tgcEditor.FirstEditableCol Then
            If lRow <= 0 Then
                Exit Function
            End If
            
            lRow = lRow - 1
        ElseIf KeyCode = vbKeyRight And m_tgTable.Col = m_tgcEditor.LastEditableCol Then
            If lRow >= m_tgcEditor.RowCount Then
                Exit Function
            End If
            
            'david 03/26/2004
            'when allow add new
            'fix to allow the highlighting to go to the last new row
            'when the last row is valid
            If m_tgcEditor.GotoNextRow(False) Then
                lRow = lRow + 1
            End If
            '''''''''''''''''
        '''''''''''''''''
        Else
            fnMarkCurrent = -1
            Exit Function
        End If
    ElseIf bCtrlDown Then
        If KeyCode = vbKeyHome Then
            lRow = 0
            m_tgTable.Bookmark = m_tgcEditor.MakeBookmark(lRow)
            
            If m_tgcEditor.FirstEditableCol >= 0 Then
                m_tgTable.Col = m_tgcEditor.FirstEditableCol
            End If
        ElseIf KeyCode = vbKeyEnd Then
            lRow = m_tgcEditor.RowCount - 1
            m_tgTable.Bookmark = m_tgcEditor.MakeBookmark(lRow)
            
            If m_tgcEditor.LastEditableCol >= 0 Then
                m_tgTable.Col = m_tgcEditor.LastEditableCol
            End If
        End If
    ElseIf bAltDown Then
        'do nothing
    End If
    
    'david 03/26/2004
    'when allow add new
    'fix to allow the highlighting to go to the last new row
    '''''''''''''''''
    'If lRow < m_tgcEditor.RowCount Then
    If lRow <= m_tgcEditor.RowCount Then
        vCurrent = m_tgcEditor.MakeBookmark(lRow)
        
        If bShiftDown Or (bShiftDown And bCtrlDown) Then
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
            fnMarkCurrent = lRow
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
    'Return the array of selected row numbers
    'and the count of the selected rows
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
    Dim i As Integer
    
    Dim bShiftDown As Boolean
    Dim bAltDown As Boolean
    Dim bCtrlDown As Boolean
    
    bShiftDown = (Shift And vbShiftMask) > 0
    bAltDown = (Shift And vbAltMask) > 0
    bCtrlDown = (Shift And vbCtrlMask) > 0
     
    
    'david 02/09/2004  #401172-1
    If IsMultiSelectCellColumn(m_nClickCol) Then
        m_nShiftControl = Shift
        
        If KeyCode = vbKeyHome Or KeyCode = vbKeyPageUp Or _
           KeyCode = vbKeyEnd Or KeyCode = vbKeyPageDown Or _
           KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or _
           KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or _
           KeyCode = vbKeyBack Or KeyCode = vbKeyReturn _
        Then
            ClearSelectedCells True
        End If
    End If
    ''''''''''''''''''''''''''''
        
    If m_nSelectMode = MODE_NO_KEY_CONTROL Then
        If KeyCode = vbKeyReturn Then
            m_nShiftControl = Shift
            subMarkSelected m_vLastRow
        End If
    ElseIf m_nSelectMode = MODE_KEY_CONTROL Then
        m_nShiftControl = Shift
        
        If bShiftDown And bCtrlDown And bAltDown Then
            'do nothing
        ElseIf bShiftDown And bAltDown Then
            'do nothing
        ElseIf bShiftDown And bCtrlDown Then
            'david 12/10/2002
            'add shift-ctrl-home, shift-ctrl-pageup, shift-ctrl-end, shift-ctrl-pagedown feature
            If Not m_tgTable.EditActive Then
                If KeyCode = vbKeyHome Or KeyCode = vbKeyPageUp Or _
                   KeyCode = vbKeyEnd Or KeyCode = vbKeyPageDown Then
                    m_nLastClickedRow = fnMarkCurrent(KeyCode, Shift)
                    '''''''''''''''''
                    Shift = 0
                    KeyCode = 0
                    Exit Sub
                End If
            End If
        ElseIf bCtrlDown And bAltDown Then
            'do nothing
        ElseIf bShiftDown Then
            'david 02/09/2004  #401172-1
            If (Not m_tgTable.EditActive) And m_bMultiSelectRowFeature Then
            ''''''''''''''''''''''''''''
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
            End If
        ElseIf bCtrlDown Then
            If Not m_tgTable.EditActive Then
                If KeyCode = vbKeyA Then
                    SelectAll
                    Shift = 0
                    KeyCode = 0
                    Exit Sub
                ElseIf KeyCode = vbKeyHome Or KeyCode = vbKeyEnd Then
                    'david 10/19/2001
                    If m_tgTable.SelBookmarks.Count > 1 Then
                        For i = m_tgTable.SelBookmarks.Count - 1 To 0 Step -1
                            m_tgTable.SelBookmarks.Remove i
                        Next i
                    'david 01/22/2002
                    ElseIf m_tgTable.SelBookmarks.Count = 1 Then
                        m_tgTable.SelBookmarks.Remove 0
                    End If
                    
                    m_nLastClickedRow = fnMarkCurrent(KeyCode, Shift)
                    '''''''''''''''''
                    Shift = 0
                    KeyCode = 0
                    Exit Sub
                ElseIf KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
                    'david 12/12/2002
                    'fixed ctrl-up, ctrl-down
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
                End If
            End If
        ElseIf bAltDown Then
            'do nothing
        ElseIf bShiftDown = 0 Then  'no special key is pressed
            'david 02/01/2001
            'david 03/11/2002 add 'And KeyCode <> vbKeyTab'
            If KeyCode <> 0 And KeyCode <> vbKeyTab Then
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
            Else
                m_nLastClickedRow = fnMarkCurrent(KeyCode, Shift)
            End If
        End If
    End If
    
    If KeyCode = vbKeyDown Then
        If Not m_tgcEditor Is Nothing Then
            If m_tgcEditor.GetCurrentRowNumber > m_tgcEditor.RowCount - 1 Then
                KeyCode = 0
                Exit Sub
            End If
        End If
    End If
End Sub

'david 04/24/2002
'when the grid lost focus, the highlighted rows always be reset
'sometime we want to keep the rows highlighted... Pass in False to keep it.
Public Sub LostFocus(Optional bClearHighlight As Boolean = True)
    If m_nSelectMode = MODE_NO_KEY_CONTROL Then
        m_bTableGotFocus = False
    End If
    
    m_bTableClicked = False
    
    If Not m_tgcEditor Is Nothing Then
        If bClearHighlight Then
            SetRowHighLighted False
        End If
    End If
End Sub

'david 04/12/2004  #401172-1
Public Sub MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ShiftTest As Integer
   
    If Button = vbLeftButton Then
        ShiftTest = Shift And 7
        
        'Select Case ShiftTest
        'Case 1 ' or vbShiftMask
        'Case 2 ' or vbCtrlMask
        'Case 4 ' or vbAltMask
        'Case 3  'SHIFT and CTRL
        'Case 5  'SHIFT and ALT
        'Case 6  'CTRL and ALT
        'Case 7  'SHIFT, CTRL, and ALT
        'End Select
        
        'david 02/09/2004  #401172-1
        If IsMultiSelectCellColumn(m_tgTable.ColContaining(X)) Then
            If ShiftTest <> vbShiftMask And Shift <> vbCtrlMask And Shift <> 3 Then
                If ShiftTest <> vbShiftMask Then
                    m_nLastClickedRow = -1
                End If
            
                ClearSelectedCells
            End If
        End If
        ''''''''''''''''''''''''''''
        
        m_nShiftControl = Shift
        m_nClickPointX = X
        m_nClickPointY = Y
    End If
End Sub
''''''''''''''''''''''''''''

'david 02/09/2004  #401172-1
'NEED TO ADD THE X parameter into this function
'The calls made to this function need to be changed
'Sorry for the inconvinence
Public Sub MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Integer)
    If Button = vbLeftButton Then
        m_nShiftControl = Shift
        m_nClickPointX = X
        m_nClickPointY = Y
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

Public Sub SelChange(CANCEL As Integer)
    CANCEL = True
End Sub

Public Sub SelectAll()
    'changed for two variables from integer type to long by junsong Oct. 9, 2000
    Dim i As Long
    Dim nRows As Long
    
    'david 02/09/2004  #401172-1
    If Not m_bMultiSelectRowFeature Then
        Exit Sub
    End If
    ''''''''''''''''''''''''''''
    
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
    'david 02/09/2004  #401172-1
    If Not m_bMultiSelectRowFeature Then
        Exit Sub
    End If
    ''''''''''''''''''''''''''''
    
    'david 04/24/2002
    'subMarkSelected m_tgcEditor.MakeBookmark(lRow)
    
    If Not m_tgcEditor Is Nothing Then
        With m_tgTable.SelBookmarks
            .Add m_tgcEditor.MakeBookmark(lRow)
            SetRowHighLighted m_bRowHighLighted
        End With
    End If
End Sub

'david 04/24/2002
'Invert Select
Public Sub InvertSelect()
    Dim i As Long
    Dim j As Long
    Dim nRows As Long
    Dim lSelected() As Long
    Dim lCountSelected As Long
    Dim lMidPoint As Long
    Dim bAdd As Boolean
    
    'david 02/09/2004  #401172-1
    If Not m_bMultiSelectRowFeature Then
        Exit Sub
    End If
    ''''''''''''''''''''''''''''
    
    If m_tgcEditor Is Nothing Then
        Exit Sub
    End If
    
    nRows = m_tgcEditor.RowCount
    
    If nRows <= 0 Then
        Exit Sub
    End If
    
    GetSelected lSelected, lCountSelected
    
    If lCountSelected > 0 Then
        Clear
        lMidPoint = (lCountSelected \ 2) - 1
    Else
        lMidPoint = m_tgcEditor.GetCurrentRowNumber
    End If
    
    With m_tgTable.SelBookmarks
        For i = 0 To nRows - 1
            bAdd = True
            
            If lCountSelected > 0 Then
                If lSelected(lMidPoint) = i Then
                    bAdd = False
                ElseIf lSelected(lMidPoint) > i Then
                    For j = lMidPoint - 1 To 0 Step -1
                        If lSelected(j) = i Then
                            bAdd = False
                        End If
                    Next j
                Else
                    For j = lMidPoint + 1 To lCountSelected - 1
                        If lSelected(j) = i Then
                            bAdd = False
                        End If
                    Next j
                End If
            Else  'consider the current highlighted row as selected row
                If lMidPoint = i Then
                    bAdd = False
                End If
            End If
            
            If bAdd Then
                .Add m_tgcEditor.MakeBookmark(i)
                SetRowHighLighted m_bRowHighLighted
            End If
        Next
    End With
End Sub

Private Sub subMarkSelected(vBkMark As Variant)
    Dim i As Integer
    Dim nEnd As Integer
    Dim nStart As Integer
    Dim bFound As Boolean
    
    'david 02/09/2004  #401172-1
    If Not m_bMultiSelectRowFeature Then
        Exit Sub
    End If
    ''''''''''''''''''''''''''''
    
    If m_tgcEditor.RowCount <= 0 Then
        Exit Sub
    End If

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
        ElseIf m_nShiftControl = vbCtrlMask Then
            subToggleSelected vBkMark
        Else
            Clear
            
            If m_bSelectCurrRow Then
                If fnLong(vBkMark) >= 0 Then
                    m_tgTable.SelBookmarks.Add vBkMark
                End If
            End If
        End If
    Else
        Clear
        subToggleSelected vBkMark
    End If
End Sub

Public Sub RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    'david 02/09/2004  #401172-1
    If IsMultiSelectCellColumn(m_nClickCol) Then
        Clear
    Else
        ClearSelectedCells
    End If

    m_nClickCol = m_tgTable.Col

    If IsMultiSelectCellColumn(m_nClickCol) Then
        If m_tgcEditor.RowCount > 0 Then
            If m_bTableClicked Then
                Dim sColRow As String
                sColRow = fnMakeKey(m_nClickCol, m_tgcEditor.GetCurrentRowNumber)
                
                subMarkCellSelected sColRow, True  'IIf(CStr(LastRow) <> CStr(m_tgTable.Row), True, False)
                m_bTableClicked = False
                subSetLastClickedCell
            End If
        End If
    End If
    ''''''''''''''''''''''''''''
        
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
    If m_tgcEditor Is Nothing Then
        m_nLastClickedRow = m_tgTable.Row
    Else
        m_nLastClickedRow = m_tgcEditor.GetCurrentRowNumber
    End If
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

'david 09/17/2003  #412827-4
Private Sub subSortArray(v())
    Dim lRow As Long
    Dim j As Long
    Dim k As Long
    
    Dim vTemp As Variant
    Dim bSwapped As Boolean
    
    On Error GoTo errTrap
    
    lRow = UBound(v)
    
    For j = 0 To lRow
        For k = j + 1 To lRow - 1
            If tfnRound(v(j)) > tfnRound(v(k)) Then
                vTemp = v(j)
                v(j) = v(k)
                v(k) = vTemp
            End If
        Next k
    Next j

errTrap:
    'do nothing
End Sub

Private Sub Class_Initialize()
    m_nSelectMode = MODE_KEY_CONTROL
    m_bTableGotFocus = True
    m_nLastClickedRow = -1
    m_nIdleColumn = -1
    m_bAvoidBeep = True
    m_bSelectCurrRow = True
    m_bMultiSelectRowFeature = True
    m_nClickCol = -1
End Sub

Private Sub Class_Terminate()
    While m_colMultiSelCellColumns.Count > 0
        m_colMultiSelCellColumns.Remove 1
    Wend
    
    While m_colMultiSelectedCells.Count > 0
        m_colMultiSelectedCells.Remove 1
    Wend
    
    Set m_colMultiSelCellColumns = Nothing
    Set m_colMultiSelectedCells = Nothing
    
    Set m_tgTable = Nothing
    Set m_tgcEditor = Nothing
    Set m_tgcSelector = Nothing
End Sub

'david 02/09/2004  #401172-1
Property Set CellColor(objCellColor As Object)
    If objCellColor Is Nothing Then
        MsgBox "CellColor object is not valid. Setup the clsCellColor before setting up clsSelector.", vbInformation
        Exit Property
    End If
    
    Set m_objCellColor = objCellColor
    
    Dim i As Integer
    
    'turn the FetchStyle on if the column has the cellcolor feature turned on
    For i = 0 To m_objCellColor.Table.Columns.Count - 1
        If m_objCellColor.IsColumnForCellColor(i) Then
            m_objCellColor.Table.Columns(i).FetchStyle = True
        End If
    Next i
End Property

Public Sub AddMultiSelectCellColumn(nCol As Integer)
    On Error Resume Next
    
    If m_colMultiSelCellColumns Is Nothing Then
        Set m_colMultiSelCellColumns = New Collection
    End If

    If m_colMultiSelectedCells Is Nothing Then
        Set m_colMultiSelectedCells = New Collection
    End If
    
    m_colMultiSelCellColumns.Add "+", fnMakeKey(nCol)
    m_tgTable.Columns(nCol).FetchStyle = True
End Sub

Public Sub RemoveMultiSelectCellColumn(nCol As Integer)
    On Error Resume Next
    m_colMultiSelCellColumns.Remove fnMakeKey(nCol)
    m_tgTable.Columns(nCol).FetchStyle = False
End Sub

Property Let MultiSelectRowFeature(bOnOff As Boolean)
    m_bMultiSelectRowFeature = bOnOff
End Property

Property Get MultiSelectRowFeature() As Boolean
    MultiSelectRowFeature = m_bMultiSelectRowFeature
End Property

Property Let FetchStyle(nCol As Integer, bOnOff As Boolean)
    m_tgTable.Columns(nCol).FetchStyle = bOnOff
End Property

Public Sub FetchCellStyle(ByVal Condition As Integer, _
                          ByVal Split As Integer, _
                          Bookmark As Variant, _
                          ByVal Col As Integer, _
                          ByVal CellStyle As Object)
    
    'if column and row found in the collection
    If IsCellSelected(Col, Bookmark) Then
        CellStyle.ForeColor = vbBlue
        CellStyle.BackColor = &HE0E0E0
    ElseIf Condition = dbgSelectedRow + dbgMarqueeRow Or Condition = dbgSelectedRow + dbgMarqueeRow + dbgCurrentCell Or Condition = dbgSelectedRow Or Condition = dbgSelectedRow + dbgUpdatedCell Then 'selected row(s)
        CellStyle.ForeColor = &H8000000E
        CellStyle.BackColor = &H8000000D
    End If

End Sub

Public Sub HeadClick(ColIndex As Integer)
    If Not IsMultiSelectCellColumn(ColIndex) Then
        Exit Sub
    End If
    
    On Error Resume Next
    
    Dim i As Long
    Dim sColRow As String
    Dim bShiftDown As Boolean
    Dim bAltDown As Boolean
    Dim bCtrlDown As Boolean
    
    bShiftDown = (m_nShiftControl And vbShiftMask) > 0
    bCtrlDown = (m_nShiftControl And vbCtrlMask) > 0
     
    'If bShiftDown And bCtrlDown Then
    If (Not bShiftDown) And (Not bCtrlDown) Then
        ClearSelectedCells True
    End If
    
    m_nClickCol = ColIndex
    
    For i = 0 To m_tgcEditor.RowCount - 1
        subAddSelectedCell ColIndex, i
    Next i
End Sub

Property Get IsMultiSelectCellColumn(ByVal nCol As Integer) As Boolean
    On Error GoTo notfound
    
    Dim sTemp As String
    sTemp = m_colMultiSelCellColumns.Item(fnMakeKey(nCol))
    IsMultiSelectCellColumn = True
    
    Exit Property
    
notfound:
    IsMultiSelectCellColumn = False
End Property

Property Get IsCellSelected(ByVal Col As Integer, ByVal Bookmark As Variant) As Boolean
    On Error GoTo notfound
    
    Dim sTemp As String
    sTemp = m_colMultiSelectedCells.Item(fnMakeKey(Col, Bookmark))
    IsCellSelected = True
    
    Exit Property
    
notfound:
    IsCellSelected = False
End Property

Property Get SelectedCellCount() As Long
    SelectedCellCount = m_colMultiSelectedCells.Count
End Property

'return the selectedcellcount and put the selected cell "col,row" into the string array
Public Function GetSelectedCells(ary() As String) As Long
    Dim i As Long
    
    If m_colMultiSelectedCells.Count > 0 Then
        ReDim ary(m_colMultiSelectedCells.Count - 1)
        
        For i = 1 To m_colMultiSelectedCells.Count
            ary(i) = m_colMultiSelectedCells.Item(i)
        Next i
    Else
        ReDim ary(0)
    End If
    
    GetSelectedCells = m_colMultiSelectedCells.Count
End Function

'return the selected cell "col,row" by using index
'return empty string if not found
Public Function GetSelectedCell(ndx As Long) As String
    On Error Resume Next
    GetSelectedCell = ""
    GetSelectedCell = m_colMultiSelectedCells.Item(ndx)
End Function

'set nCol and lRow if found
'set nCol=-1 and lRow=-1 if not found
Public Sub GetSelectedColRow(ndx As Long, nCol As Integer, lRow As Long)
    Dim sTemp As String
    
    On Error Resume Next
    sTemp = m_colMultiSelectedCells.Item(ndx)
    
    If sTemp <> "" Then
        subGetColRow sTemp, nCol, lRow
    Else
        nCol = -1
        lRow = -1
    End If
End Sub

Public Sub ClearSelectedCells(Optional bRebind As Boolean = False)
    If m_colMultiSelectedCells Is Nothing Then
        Exit Sub
    End If
    
    While m_colMultiSelectedCells.Count > 0
        m_colMultiSelectedCells.Remove 1
    Wend
    
    m_sLastClickedCell = ""
    m_nClickCol = -1
    
    If bRebind Then
        m_tgTable.Rebind
    End If
End Sub

'need to have 3 parm in a group
'namely Column, Operator (=, +, -, *, /), Value
'= is direct assignment, + is increase, - is decrease, * is multply, / is divide
'e.g. 2, "=", "02/03/02"
'e.g. 6, "+", "," + sMoreMessage
'e.g. 5, "=", "1.059"
'e.g. 4, "*", "0.95"
'return True if any cellvalue is changed in the function
Public Function ChangeValues(ParamArray aryColOpValue() As Variant) As Boolean
    Dim i As Integer
    Dim j As Long
    Dim nCol As Integer
    Dim sOP As String
    Dim sValue As String
    Dim nCellCol As Integer
    Dim lCellRow As Long
    Dim bChanged As Boolean
    
    On Error GoTo errTrap
    
    i = -1
    
    On Error Resume Next
    i = UBound(aryColOpValue)
    On Error GoTo 0
    
    If i < 0 Then
        Exit Function
    End If
    
    'need to have 3 parm in a group
    'namely Column, Operator (=, +, -, *, /), Value
    '= is direct assignment, + is increase, - is decrease, * is multply, / is divide
    'e.g. 2, "=", "02/03/02"
    'e.g. 6, "+", "," + sMoreMessage
    'e.g. 5, "=", "1.059"
    'e.g. 4, "*", "0.95"
    
    If (i + 1) Mod 3 <> 0 Then
        Exit Function
    End If
    
    i = (i + 1) / 3
    
    For i = 0 To i
        If Not IsNull(aryColOpValue(i)) Then
            nCol = Val(aryColOpValue(i))
        Else
            nCol = -1
        End If
        
        sOP = CStr(aryColOpValue(i + 1) & "")
        sValue = CStr(aryColOpValue(i + 2) & "")
        
        If nCol >= 0 And IsMultiSelectCellColumn(nCol) Then
            'update the cellvalue
            For j = 1 To m_colMultiSelectedCells.Count
                GetSelectedColRow j, nCellCol, lCellRow
                
                If nCellCol = nCol Then
                    Select Case sOP
                    Case "="  'direct assignment
                        If CStr(m_tgcEditor.CellValue(nCol, lCellRow) & "") <> sValue Then
                            m_tgcEditor.CellValue(nCol, lCellRow) = sValue
                            'revalidate column
                            m_tgcEditor.Revalidate nCol, i
                            bChanged = True
                        End If
                    Case "+"
                        m_tgcEditor.CellValue(nCol, lCellRow) = m_tgcEditor.CellValue(nCol, lCellRow) + sValue
                        bChanged = True
                    Case "-"
                        m_tgcEditor.CellValue(nCol, lCellRow) = m_tgcEditor.CellValue(nCol, lCellRow) - sValue
                        bChanged = True
                    Case "*"
                        m_tgcEditor.CellValue(nCol, lCellRow) = m_tgcEditor.CellValue(nCol, lCellRow) * Val(sValue)
                        bChanged = True
                    Case "/"
                        If Val(sValue) <> 0 Then
                            m_tgcEditor.CellValue(nCol, lCellRow) = m_tgcEditor.CellValue(nCol, lCellRow) / Val(sValue)
                            bChanged = True
                        End If
                    End Select
                End If
            Next j
        End If
    Next i
    
    Exit Function
    
errTrap:
    'do nothing
End Function

Private Function fnMakeKey(ByVal nCol As Integer, Optional ByVal lBookmark As Long = -1) As String
    'On Error Resume Next
    fnMakeKey = CStr(nCol) + "," + CStr(lBookmark)
End Function

Private Sub subSetLastClickedCell(Optional sColRow As String = "")
    On Error Resume Next
    
    If sColRow = "" Then
        If m_tgcEditor Is Nothing Then
            If m_tgTable.Row >= 0 Then
                sColRow = fnMakeKey(m_nClickCol, m_tgTable.Bookmark)
            End If
        Else
            sColRow = fnMakeKey(m_nClickCol, m_tgcEditor.GetCurrentRowNumber)
        End If
    End If
    
    m_sLastClickedCell = sColRow
End Sub

Private Sub subMarkCellSelected(ByVal sColRow As String, Optional ByVal bRebind As Boolean = False)
    If m_tgcEditor.RowCount <= 0 Then
        Exit Sub
    End If

    Clear

    Dim bShiftDown As Boolean
    Dim bAltDown As Boolean
    Dim bCtrlDown As Boolean
     
    Dim nCol As Integer
    Dim lRow As Long
    Dim nStartCol As Integer
    Dim nCurCol As Integer
    Dim lStartRow As Long
    Dim lCurRow As Long
    
    bShiftDown = (m_nShiftControl And vbShiftMask) > 0
    bAltDown = (m_nShiftControl And vbAltMask) > 0
    bCtrlDown = (m_nShiftControl And vbCtrlMask) > 0
     
    subGetColRow sColRow, nCurCol, lCurRow
     
    'If bShiftDown And bCtrlDown Then
    If bShiftDown Then
        nStartCol = -1
        lStartRow = -1
        
        If m_sLastClickedCell <> "" Then
            subGetColRow m_sLastClickedCell, nStartCol, lStartRow
        End If
            
        If nStartCol >= 0 And lStartRow >= 0 And nCurCol >= 0 And lCurRow >= 0 Then
            'swap start and end, make sure start is less than end
            If nStartCol > nCurCol Then
                nCol = nStartCol
                nStartCol = nCurCol
                nCurCol = nCol
            End If
            If lStartRow > lCurRow Then
                lRow = lStartRow
                lStartRow = lCurRow
                lCurRow = lRow
            End If
            
            For nCol = nStartCol To nCurCol
                If IsMultiSelectCellColumn(nCol) Then
                    For lRow = lStartRow To lCurRow
                        subAddSelectedCell nCol, lRow
                    Next lRow
                End If
            Next nCol
                    
            If Not bRebind Then
                If lStartRow <> lCurRow Or Abs(nStartCol - nCurCol) > 1 Then
                    bRebind = True
                End If
            End If
        End If
    ElseIf bCtrlDown Then
        subToggleCellSelected nCurCol, lCurRow
    Else
        'all selected cells have already been cleared in MouseUp event
        'no need to call ClearSelectedCells() function here
        subAddSelectedCell nCurCol, lCurRow
    End If

    If bRebind Then
        m_tgTable.Rebind
    End If
End Sub

Private Sub subAddSelectedCell(ByVal nCol As Integer, ByVal lRow As Long)
    Dim sColRow As String
    
    sColRow = fnMakeKey(nCol, lRow)
    
    On Error Resume Next
    m_colMultiSelectedCells.Add sColRow, sColRow
End Sub

Private Sub subGetColRow(sColRow As String, nCol As Integer, lRow As Long)
    Dim nPosi As Integer
    
    nPosi = InStr(sColRow, ",")
    
    If nPosi > 0 Then
        nCol = Val(Left(sColRow, nPosi - 1))
        lRow = Mid(sColRow, nPosi + 1)
    Else
        nCol = Val(sColRow)
        lRow = -1
    End If
End Sub

Private Sub subToggleCellSelected(ByVal nCol As Integer, ByVal lRow As Long)
    If IsCellSelected(nCol, lRow) Then
        m_colMultiSelectedCells.Remove fnMakeKey(nCol, lRow)
    Else
        subAddSelectedCell nCol, lRow
    End If
End Sub
'''''''''''''''''''''''''''''''''''''

' ***************************************************************************
' * Instructions: Paste the lines below into a module, or into your form, and
' *               call using the example below.
' * Example:
' *          MsgBox "Is Running In VB IDE: " & IsIDE()
' ***************************************************************************
Private Function IsIDE() As Boolean
    On Error GoTo ExitHandler
    
    ' in a compiled VB program, the Debug.Print line would not
    ' be compiled.  So on this function being called, there would be
    ' no Error (Err=0) and the return value would be False
    ' If however, it was design time, an error would be generated due to
    ' "Divide by Zero" and non-false (True) would be returned
    Debug.Print 1 / 0

ExitHandler:
    IsIDE = Err
End Function

