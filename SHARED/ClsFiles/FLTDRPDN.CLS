VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFloatingDropDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***********************************************************'
'
' Copyright (c) 1996 FACTOR, A Division of W.R.Hess Company
'
' Module name   : FLTDRPDN.CLS
' Date          : May 16, 1996
' Programmer(s) : Qinggang Ma
'
' This module defines a class for floating dropdown in a true grid.


' A floating drop down class using true grid
' Usage:
'      i. Include class files COLUMN.CLS and FLTDRPDN.CLS in the project
'     ii. Define and create a class: clsFloatingDropDown
'    iii. Initialization (See sub fnInitializeDropDown in the form)
'     iv. Event calls:
'           1. In the main true grid:
'                a. call .TableClick True-Grid-Name           in Click event
'                b. call .GotFocus                            in GotFocus event
'                c. call .Keypress True-Grid-Name, KeyAscii   in Keypress event
'                d. call .LostFocus True-Grid-Name            in LostFocus event
'                e. call .RowColChange True-Grid-Name         in RowColChange event
'                f. call .Scroll                              in Scroll event
'                g. Cancel = True                             in SelChange event
'           2. In the dropdown true grid:
'                a. call .TableClick True-Grid-Name           in Click event
'                b. call .KeyPress True-Grid-Name, KeyAscii   in Keypress event
'                c. call .LostFocus True-Grid-Name            in LostFocus event
'                d. call .MouseUp Y                           in MouseUp event
'                e. call .RowColChange True-Grid-Name         in RowColChange event
'           3. In the command button
'                a. call .ButtonClick                         in Click event
'                b. call .LostFocus Button-Name               in LostFocus event

Option Explicit
    
    Private Const TYPE_STRING = 1
    Private Const TYPE_INTEGER = 2
    Private Const TYPE_HIDDEN = 3
    
    Private Const TYPE_DRPDN_FACTORSEARCH = 1
    Private Const TYPE_DRPDN_COMMONDROPDOWN = 2
    Private Const TYPE_DRPDN_NO_CRITERIA = 3
    Private Const TYPE_DRPDN_NO_CRITERIA2 = 3
    
    Private Const EXTRA_COL_RIGHT = 1
    Private Const EXTRA_COL_LEFT = 2
    
    Private Const DB_INFORMIX As Integer = 1
    Private Const DB_ACCESS As Integer = 2

    Private Const EXTRACOLUMNOFFSET = 1000
    Private Const sINITIAL As Integer = -1
    
    Private Const DEFAULT_VISIBLE_ROWS As Integer = 5
    Private Const SCROLLBAR_WIDTH = 252

    Private nCurrentDropDown As Integer
    
    Private colDropDowns As Collection

    Private nColumnCount As Integer

    Private tgMain As TDBGrid
    Private tgDropDown As TDBGrid
    Private lblStatusbar As Control
    Private frmParent As Form
    Private cmdSearch As Control
    Private datLink As Data
    Private dbMain As DataBase
    
    Private bTableClicked As Boolean
    Private nClickPointY As Single
    Private bValidSelection As Boolean
    Private nLastColumn As Integer
    Private nDatabase As Integer
    Private nMainTableTop As Integer
    Private nMainTableLeft As Integer
    Private nTopBoundary As Integer
    Private nBottomBoundary As Integer
    Private nMaxVisibleRows As Integer
    Private nVisibleColumns As Integer
    Private bSearchOnReturn As Boolean
    Private bKeyPressed As Boolean
    Private bNavigating As Boolean
    Private bShowList As Boolean
    Private bAutoHide As Boolean

    Private nOriginalTableWidth As Integer
    Private nOriginalTableHeight As Integer
    Private nFontSize As Integer

    Private Const STATUS_INIT = 0
    Private Const STATUS_FILLED = 1

    Private m_clsEdit As Object
    Private m_LastRow As Integer
    Private m_nSeries As Integer
    Private m_nCurrentDropDown As Integer
    Private m_bDefaultCursorControl As Boolean
    Private m_bLocked As Boolean

    #If Win32 Then
        Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
    #Else
        Private Declare Function ScreenToClient Lib "USER" (ByVal hwnd As Integer, lpPoint As POINTAPI) As Boolean
    #End If
    

Property Let AutoHide(ByVal bTemp As Boolean)
    bAutoHide = bTemp
End Property

Property Get CurrentDropDown() As Integer
    CurrentDropDown = m_nCurrentDropDown
End Property

Public Sub FirstRowChange()
'    If Not bNavigating Then
    If tgMain.Enabled Then
        fnMoveButton
        LostFocus tgDropDown
    Else
        cmdSearch.Visible = False
    End If
'    End If
End Sub


Public Function fnSQLString(ByVal vTemp As Variant) As String
'
' Properly quotes and formats an SQL string.  If vNoQuotes is present, the result WILL NOT BE QUOTED
' for each ' character found, insert a double ''.  Leave "%* alone
    
    Dim nIdx As Integer
    Dim nPos As Integer
    Dim szParameter As String
    
    If IsNull(vTemp) Then
        szParameter = ""
    Else
        szParameter = Trim(vTemp)
    End If

    nIdx = 1
    nPos = InStr(nIdx, szParameter, "'")
    
    While nPos <> 0
        szParameter = Left(szParameter, nPos) & "'" & Right(szParameter, Len(szParameter) - nPos)
        nIdx = nPos + 2
        nPos = InStr(nIdx, szParameter, "'")
    Wend
    
    ' quote the whole string - optional
    fnSQLString = "'" & szParameter & "'"

End Function


Property Let DefaultCursorControl(bTemp As Boolean)
    m_bDefaultCursorControl = bTemp
End Property

Property Set EditClass(objTemp As Object)
    Set m_clsEdit = objTemp
End Property


Public Function AddDropDown(Optional nKey As Variant, _
                            Optional sSql As Variant, _
                            Optional nType As Variant, _
                            Optional nVRows) As Integer
    Dim colColumn As New Collection
    Dim stCol As New clsColumns

'    If IsMissing(nKey) Then
        nCurrentDropDown = colDropDowns.Count + 1
'    Else
'        nCurrentDropDown = nKey
'    End If
    If Not IsMissing(sSql) Then
        stCol.SQL = sSql
    End If
    If IsMissing(nType) Then
        stCol.DropDownType = TYPE_DRPDN_FACTORSEARCH
    Else
        stCol.DropDownType = nType
    End If
    If nMaxVisibleRows <= 0 Then
        If IsMissing(nVRows) Then
            nMaxVisibleRows = DEFAULT_VISIBLE_ROWS
        Else
            nMaxVisibleRows = nVRows
        End If
    End If
    With stCol
        .Unique = False
        .Visible = True
        .PopulateBox = True
        .NoDataMessage = True
        .ShowCaption = True
        .CursorControl = m_bDefaultCursorControl
    End With
    
    On Error GoTo errAddDropDown
    colColumn.Add stCol, Str$(sINITIAL)
    colDropDowns.Add colColumn, Str$(nCurrentDropDown)
    nCurrentDropDown = colDropDowns.Count
    AddDropDown = nCurrentDropDown
extAddDropDown:
    On Error GoTo 0
    Exit Function
errAddDropDown:
    MsgBox Err.Description
    AddDropDown = -1
    Resume extAddDropDown
End Function

Property Let ColumnWidth(nIndex As Integer, nWidth As Integer)
    Dim lDropDown As Long
    Dim lColumn As Long
    Dim colColumn As Collection
    
    fnGetKeys lDropDown, lColumn, nIndex
    If lDropDown > 0 Then
        Set colColumn = colDropDowns.Item(lDropDown)
        colColumn.Item(lColumn).ColWidth = nWidth
    End If

End Property

Property Get ColumnWidth(nIndex As Integer) As Integer
    Dim lDropDown As Long
    Dim lColumn As Long
    Dim colColumn As Collection
    
    fnGetKeys lDropDown, lColumn, nIndex
    If lDropDown > 0 Then
        Set colColumn = colDropDowns.Item(lDropDown)
        ColumnWidth = colColumn.Item(lColumn).ColWidth
    End If

End Property
Property Get CommonDropdownButton() As Integer
    CommonDropdownButton = TYPE_DRPDN_COMMONDROPDOWN
End Property
Property Get ExtendLeft() As Integer
    ExtendLeft = EXTRA_COL_LEFT
End Property
Property Get ExtendRight() As Integer
    ExtendRight = EXTRA_COL_RIGHT
End Property


Property Get FactorSerachButton() As Integer
    FactorSerachButton = TYPE_DRPDN_FACTORSEARCH
End Property


Private Function fnVisibleColumns(nCol As Integer) As Integer
    Dim i As Integer
    Dim nCount As Integer
    
    nCount = 0
    For i = 0 To nCol
        If tgMain.Columns(i).Visible Then
            nCount = nCount + 1
        End If
    Next
    
    fnVisibleColumns = nCount

End Function

Private Function fnLocateDropDownTable(colTemp As Collection)
    Dim lCol As Long
    Dim i As Integer
    Dim nExtra As Integer
    Dim nTWidth As Integer
    Dim objCol As Column
    Dim nRight As Integer
    Dim nTemp As Integer
    Dim bFlag As Boolean
    
    nTWidth = SCROLLBAR_WIDTH
    If nVisibleColumns < 0 Then
        nVisibleColumns = fnVisibleColumns(tgMain.Columns.Count - 1)
    End If
    nExtra = 0
    For i = 1 To colTemp.Count
        If colTemp.Item(i).ColumnIndex >= EXTRACOLUMNOFFSET Then
            nExtra = nExtra + 1
        End If
    Next
    If colTemp.Item(1).ShowCaption Then
        tgDropDown.HeadLines = 1
    Else
        tgDropDown.HeadLines = 0
    End If
    If colTemp.Item(1).ExtDirection = EXTRA_COL_LEFT And nExtra > 0 Then
        nRight = nMainTableLeft
        bFlag = True
        For lCol = 1 To colTemp.Count
            With colTemp.Item(lCol)
                If .ColumnIndex < EXTRACOLUMNOFFSET Then
                    i = nExtra + lCol - 1
                    Set objCol = tgMain.Columns(.ColumnIndex)
                    nTemp = objCol.Left + objCol.Width
                    If nRight < nTemp Then
                        nRight = nTemp
                    End If
                    If .ColWidth > 0 Then
                        tgDropDown.Columns(i).Width = .ColWidth
                    Else
                        If fnVisibleColumns(.ColumnIndex) >= nVisibleColumns Then
                            bFlag = False
                            If tgMain.ExtendRightColumn Then
                                nTemp = tgMain.Width
                                tgDropDown.Columns(i).Width = nTemp - objCol.Left - SCROLLBAR_WIDTH
                                If nRight < nTemp Then
                                    nRight = nTemp
                                End If
                            Else
                                tgDropDown.Columns(i).Width = objCol.Width - SCROLLBAR_WIDTH
                            End If
                        Else
                                                                                                                                                                                                                     tgDropDown.Columns(i).Width = objCol.Width
                        End If
                    End If
                    If .Caption = "" Then
                        tgDropDown.Columns(i).Caption = objCol.Caption
                    Else
                        tgDropDown.Columns(i).Caption = .Caption
                    End If
                    If .DataField = "" Then
                        tgDropDown.Columns(i).DataField = objCol.DataField
                    Else
                        tgDropDown.Columns(i).DataField = .DataField
                    End If
                Else
                    i = colTemp.Count - lCol
                    tgDropDown.Columns(i).Width = .ColWidth
                    tgDropDown.Columns(i).Caption = .Caption
                    tgDropDown.Columns(i).DataField = .DataField
                End If
            End With
            tgDropDown.Columns(lCol - 1).Visible = True
            nTWidth = nTWidth + tgDropDown.Columns(i).Width
            nTemp = nRight - nTWidth
            If bFlag Then
                nTemp = nTemp + SCROLLBAR_WIDTH
            End If
            If nTemp < 0 Then
                nTemp = 0
            End If
            tgDropDown.Left = nMainTableLeft + nTemp
        Next
    Else
        For lCol = 1 To colTemp.Count
            With colTemp.Item(lCol)
                If .ColumnIndex < EXTRACOLUMNOFFSET Then
                    Set objCol = tgMain.Columns(.ColumnIndex)
                    If lCol = 1 Then
                        tgDropDown.Left = nMainTableLeft + objCol.Left
                    End If
                    If .ColWidth > 0 Then
                        tgDropDown.Columns(lCol - 1).Width = .ColWidth
                    Else
                        If fnVisibleColumns(.ColumnIndex) >= nVisibleColumns Then
                            If tgMain.ExtendRightColumn Then
                                tgDropDown.Columns(lCol - 1).Width = tgMain.Width - objCol.Left - SCROLLBAR_WIDTH
                            Else
                                tgDropDown.Columns(lCol - 1).Width = objCol.Width - SCROLLBAR_WIDTH
                            End If
                        Else
                            tgDropDown.Columns(lCol - 1).Width = objCol.Width
                        End If
                    End If
                    If .Caption = "" Then
                        tgDropDown.Columns(lCol - 1).Caption = objCol.Caption
                    Else
                        tgDropDown.Columns(lCol - 1).Caption = .Caption
                    End If
                    If .DataField = "" Then
                        tgDropDown.Columns(lCol - 1).DataField = objCol.DataField
                    Else
                        tgDropDown.Columns(lCol - 1).DataField = .DataField
                    End If
                Else
                    tgDropDown.Columns(lCol - 1).Width = .ColWidth
                    tgDropDown.Columns(lCol - 1).Caption = .Caption
                    tgDropDown.Columns(lCol - 1).DataField = .DataField
                End If
            End With
            tgDropDown.Columns(lCol - 1).Visible = True
            nTWidth = nTWidth + tgDropDown.Columns(lCol - 1).Width
            If tgDropDown.Left + nTWidth > nMainTableLeft + tgMain.Width Then
                nTWidth = nMainTableLeft + tgMain.Width - tgDropDown.Left
            End If
        Next
    End If
    tgDropDown.Width = nTWidth
    tgDropDown.Top = nMainTableTop + tgMain.RowTop(tgMain.Row) + tgMain.RowHeight + Screen.TwipsPerPixelY
    If tgDropDown.Top + tgDropDown.Height > frmParent.ScaleHeight Then
        tgDropDown.Top = nMainTableTop + tgMain.RowTop(tgMain.Row) - tgDropDown.Height
    End If
    tgDropDown.col = 0
End Function

Private Sub fnLocateMainTable()
    
    If tgMain Is Nothing Or tgDropDown Is Nothing Or frmParent Is Nothing Or nMainTableTop > 0 Then
        Exit Sub
    End If
    
    #If Win32 Then
        Dim nMain_hWnd As Long
        Dim nDrpDn_hWnd As Long
        Dim n_ParenthWnd As Long
    #Else
        Dim nMain_hWnd As Integer
        Dim nDrpDn_hWnd As Integer
        Dim n_ParenthWnd As Integer
    #End If
    Dim nLeft As Integer
    Dim nTop As Integer
    Dim nClientLeft As Single
    Dim nClientTop As Single

    nMain_hWnd = tgMain.hwnd
    n_ParenthWnd = GetParent(nMain_hWnd)
    nDrpDn_hWnd = GetParent(tgDropDown.hwnd)
    nTopBoundary = frmParent.Top          'objTemp.Top
    nBottomBoundary = nTopBoundary + frmParent.ScaleHeight
    nLeft = 0   'tgMain.Left
    nTop = 0    'tgMain.Top
    While nMain_hWnd <> frmParent.hwnd And n_ParenthWnd > 0
        fnGetLeftTop nClientLeft, nClientTop, n_ParenthWnd, nMain_hWnd
        nLeft = nLeft + nClientLeft
        nTop = nTop + nClientTop
        nBottomBoundary = nBottomBoundary - nTop
        nMain_hWnd = n_ParenthWnd
        n_ParenthWnd = GetParent(nMain_hWnd)
    Wend
    nMainTableLeft = nLeft - Screen.TwipsPerPixelX
    nMainTableTop = nTop - Screen.TwipsPerPixelY
End Sub

#If Win32 Then
    Private Function fnGetLeftTop(nLeft As Single, _
                                  nTop As Single, _
                                  ByVal nParenthWnd As Long, _
                                  ByVal n_hWnd As Long)
#Else
    Private Function fnGetLeftTop(nLeft As Single, _
                                  nTop As Single, _
                                  ByVal nParenthWnd As Integer, _
                                  ByVal n_hWnd As Integer)
#End If

    Dim myRect As RECT
    Dim myPoint As POINTAPI
    
    GetWindowRect n_hWnd, myRect

    myPoint.x = myRect.Left
    myPoint.y = myRect.Top
    
    If ScreenToClient(nParenthWnd, myPoint) Then
'        nLeft = (myPoint.X - 1) * Screen.TwipsPerPixelX    'For True Gridtable
        nLeft = myPoint.x * Screen.TwipsPerPixelX     'For textbox, VedioSoftElastic
        If nLeft < -6000 Or nLeft > 6000 Then
            nLeft = Screen.TwipsPerPixelX
        End If
'        nTop = (myPoint.Y - 1) * Screen.TwipsPerPixelY    'For True Gridtable
        nTop = myPoint.y * Screen.TwipsPerPixelY       'For textbox, VedioSoftElastic
        If nTop < 0 Or nTop > 10000 Then
            nTop = Screen.TwipsPerPixelX
        End If
    End If

End Function


Public Sub AddExtraColumn(sCap As String, _
                          sField As String, _
                          nWidth As Integer, _
                          Optional nDropDown As Variant)
    Dim colTemp As Collection
    Dim clsCol As New clsColumns

    If Not IsMissing(nDropDown) Then
        nCurrentDropDown = nDropDown
    End If
    If fnKeyExist(colDropDowns, nCurrentDropDown) Then
        Set colTemp = colDropDowns.Item(Str$(nCurrentDropDown))
    Else
        Exit Sub
    End If
    If colTemp Is Nothing Then
        MsgBox "Error adding new column"
    End If
    clsCol.ColumnIndex = EXTRACOLUMNOFFSET + colTemp.Count
    clsCol.Caption = sCap
    clsCol.DataField = sField
    clsCol.ColWidth = nWidth
    
    colTemp.Add clsCol
End Sub


Property Let ColumnCaption(nIndex As Integer, sCap As String)
    Dim lDropDown As Long
    Dim lColumn As Long
    Dim colColumn As Collection
    
    fnGetKeys lDropDown, lColumn, nIndex
    If lDropDown > 0 Then
        Set colColumn = colDropDowns.Item(lDropDown)
        colColumn.Item(lColumn).Caption = sCap
    End If
End Property
Property Get ColumnCaption(nIndex As Integer) As String
    Dim lDropDown As Long
    Dim lColumn As Long
    Dim colColumn As Collection
    
    fnGetKeys lDropDown, lColumn, nIndex
    If lDropDown > 0 Then
        Set colColumn = colDropDowns.Item(lDropDown)
        ColumnCaption = colColumn.Item(lColumn).Caption
    End If
End Property

Property Let ColumnDataField(nIndex As Integer, sField As String)

    Dim lDropDown As Long
    Dim lColumn As Long
    Dim colColumn As Collection
    
    fnGetKeys lDropDown, lColumn, nIndex
    If lDropDown > 0 Then
        Set colColumn = colDropDowns.Item(lDropDown)
        colColumn.Item(lColumn).DataField = sField
    End If

End Property

Property Get ColumnDataField(nIndex As Integer) As String

    Dim lDropDown As Long
    Dim lColumn As Long
    Dim colColumn As Collection
    
    fnGetKeys lDropDown, lColumn, nIndex
    If lDropDown > 0 Then
        Set colColumn = colDropDowns.Item(lDropDown)
        ColumnDataField = colColumn.Item(lColumn).DataField
    End If

End Property

Property Get DataBaseType() As Integer
    DataBaseType = nDatabase
End Property

Property Let DataBaseType(nType As Integer)
    nDatabase = nType
End Property
Property Get dbAccess() As Integer
    dbAccess = DB_ACCESS
End Property
Property Get dbInformix() As Integer
    dbInformix = DB_INFORMIX
End Property

Property Get DropDownCount() As Integer
    DropDownCount = colDropDowns.Count
End Property

Private Sub fnGetKeys(lDropDown As Long, lColumn As Long, ByVal nCol As Integer)
    Dim i As Long
    Dim j As Long
    Dim colTemp As Collection
    
    For i = 1 To colDropDowns.Count
        Set colTemp = colDropDowns.Item(i)
        For j = 1 To colTemp.Count
            If colTemp.Item(j).ColumnIndex = nCol Then
                lDropDown = i
                lColumn = j
                Exit Sub
            End If
        Next j
    Next i
'    MsgBox "Column not defined for drop down"
    lDropDown = -1
End Sub

Private Function fnKeyExist(col As Collection, nKey As Integer) As Boolean
    Dim vTemp As Variant

    fnKeyExist = True
    On Error GoTo errfnKeyExist
    Set vTemp = col.Item(Str$(nKey))
extfnKeyExist:
    On Error GoTo 0
    Exit Function
errfnKeyExist:
    If Err.Number = 357 Or Err.Number = 5 Then
        fnKeyExist = False
        Resume extfnKeyExist
    Else
        On Error GoTo 0
        Resume
    End If
End Function

Private Function fnLoadTable(nColumn As Integer) As Boolean

    'Declare Sub's Variables
    Dim szCode As String
    Dim szQuery As String
    Dim sTemp As String
    Dim rsTemp As Recordset
    Dim szLink As String
    Dim i As Integer
    Dim nPos As Integer
    Dim lDropDown As Long
    Dim lCol As Long
    Dim lColIndex As Long
    Dim colTemp As Collection
    Dim vData() As Variant
    Dim lCount As Long
    Dim sGroupByClause As String
    
'    Dim nTWidth As Integer
    If dbMain Is Nothing Then
        Exit Function
    End If
    If tgMain.Row < 0 Then
        On Error Resume Next
        tgMain.SetFocus
        fnMoveButton
        Exit Function
    End If

    DoEvents

    'Make our mouse pointer an hourglass
    fnLoadTable = False
    Screen.MousePointer = vbHourglass
    bValidSelection = False
    
    tgDropDown.Visible = False
    fnGetKeys lDropDown, lColIndex, nColumn
    subSetFlagSelection lDropDown, False
    
    If lDropDown < 0 Then
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    fnLoadTable = True
    If Not colDropDowns.Item(lDropDown).Item(1).Visible Then
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    Set colTemp = colDropDowns.Item(lDropDown)
    
    szQuery = colTemp.Item(1).SQL
    nPos = InStr(UCase(szQuery), " GROUP BY ")
    If nPos > 1 Then
        sGroupByClause = Right(szQuery, Len(szQuery) - nPos + 1)
        szQuery = Left(szQuery, nPos - 1)
    Else
        sGroupByClause = ""
    End If
    
    If colTemp.Item(1).DropDownType <> TYPE_DRPDN_NO_CRITERIA And colTemp.Item(1).DropDownType <> TYPE_DRPDN_NO_CRITERIA2 Then
        nPos = InStr(UCase(szQuery), "UNION")
        If nPos = 0 Then
            nPos = InStr(UCase$(szQuery), "WHERE")
            If nPos = 0 Then
                szLink = " WHERE ("
            Else
                szLink = " AND ("
            End If
            
            For lCol = 1 To colTemp.Count
                If colTemp.Item(lCol).ColumnIndex < EXTRACOLUMNOFFSET Then
                    szCode = Trim$(tgMain.Columns(colTemp.Item(lCol).ColumnIndex).Text)
                    ' if a code mask has been entered, add to the SQL
                    If szCode <> "" Then
                        If colTemp.Item(lCol).ColType = TYPE_STRING Then
                            szQuery = szQuery & szLink & colTemp.Item(lCol).DataField & " LIKE " & fnSQLString(szCode & "%") & ")"
                            szLink = " AND ("
                        ElseIf colTemp.Item(lCol).ColType = TYPE_INTEGER Then
                            szQuery = szQuery & szLink & fnGenetateSearchString(colTemp.Item(lCol).DataField, szCode) & ")"
                            szLink = " AND ("
                        End If
                    End If
                End If
            Next
            If colTemp.Item(1).Unique Then
                If Not m_clsEdit Is Nothing Then
                    m_clsEdit.GetColumn vData, lCount, nColumn
                    lCol = m_clsEdit.GetCurrentRowNumber
                    sTemp = ""
                    For i = 0 To lCount - 1
                        If i <> lCol Then
                            If Not IsNull(vData(i)) Then
                                If Trim(vData(i)) <> "" Then
                                    If sTemp = "" Then
                                        sTemp = fnSQLString(vData(i))
                                    Else
                                        sTemp = sTemp & "," & fnSQLString(vData(i))
                                    End If
                                End If
                            End If
                        End If
                    Next
                    If sTemp <> "" Then
                        If InStr(UCase(szQuery), "WHERE ") > 0 Then
                            szQuery = szQuery & " AND " & colTemp.Item(lColIndex).DataField & " NOT IN (" & sTemp & ")"
                        Else
                            szQuery = szQuery & " WHERE " & colTemp.Item(lColIndex).DataField & " NOT IN (" & sTemp & ")"
                        End If
                    End If
                End If
            End If
            sTemp = colTemp.Item(lColIndex).DataField
        Else
            lCol = fnCoulmnOrder(lDropDown, nColumn)
            If lCol > 0 Then
                sTemp = Str$(lCol)
            End If
        End If
    Else
        sTemp = colTemp.Item(lColIndex).DataField
    End If
    
    szQuery = szQuery & sGroupByClause & " ORDER BY " & sTemp
    'Get the records
    If Not fnGetRecord(rsTemp, szQuery, nDatabase) Then
        Screen.MousePointer = vbDefault
        ' Set the return value false
        tgDropDown.Visible = False
        On Error Resume Next
        tgMain.SetFocus
        Exit Function
    End If

    'Check there is something for us to display before we show the results
    If rsTemp.RecordCount > 0 Then
        ' Set load the data table
        Set datLink.Recordset = rsTemp

        ' Set the return values
        If rsTemp.RecordCount = 1 And colTemp.Item(1).PopulateBox Then
            fnGetText nColumn
        Else
            While tgDropDown.Columns.Count > 0
                tgDropDown.Columns.Remove 0
            Wend
            For lCol = 1 To colTemp.Count
                tgDropDown.Columns.Add lCol - 1
            Next

            If rsTemp.RecordCount < nMaxVisibleRows Then
                tgDropDown.Height = (tgDropDown.HeadLines + rsTemp.RecordCount + 1) * tgDropDown.RowHeight + 2 * Screen.TwipsPerPixelY
            Else
                tgDropDown.Height = (tgDropDown.HeadLines + nMaxVisibleRows) * tgDropDown.RowHeight + 2 * Screen.TwipsPerPixelY
            End If
            fnLocateDropDownTable colTemp

            tgDropDown.Visible = True
            On Error Resume Next
            bShowList = True
            tgDropDown.ZOrder 0
            tgDropDown.SetFocus
        End If
    Else
        If colDropDowns.Item(lDropDown).Item(1).NoDataMessage Then
            ' Set the return value false
            Dim sMsg As String
            sMsg = "No record found for the given criteria"
            MsgBox sMsg, vbOKOnly + vbInformation, frmParent.Caption
            On Error Resume Next
            tgMain.SetFocus
        End If
    End If

    'Reset our mouse pointer to the default
    Screen.MousePointer = vbDefault
End Function

Public Sub InitialState()
    cmdSearch.Visible = False
    cmdSearch.ZOrder 1
End Sub

Public Sub KeyUp()
    bNavigating = False
'    fnMoveButton
End Sub

Public Sub LeftColChange()
    If m_bLocked Then
        Exit Sub
    End If
    fnMoveButton
End Sub

Property Let Locked(bTemp As Boolean)
    m_bLocked = bTemp
End Property

Property Get NoCriteria() As Integer
    NoCriteria = TYPE_DRPDN_NO_CRITERIA
End Property


Property Get NoCriteria2() As Integer
    NoCriteria2 = TYPE_DRPDN_NO_CRITERIA2
End Property

Property Let SearchOnReturn(bFlag As Boolean)
    bSearchOnReturn = bFlag
End Property

Public Sub SetPopulateBox(nCol As Integer, bStatus As Boolean)
    Dim lIndex As Long
    Dim lColIndex As Long
    
    fnGetKeys lIndex, lColIndex, nCol
    If lIndex > 0 Then
        colDropDowns.Item(lIndex).Item(1).PopulateBox = bStatus
    End If

End Sub


Public Sub SetSQL(ByVal nCol As Integer, sSql As String)
    Dim lIndex As Long
    Dim lColIndex As Long
    
    fnGetKeys lIndex, lColIndex, nCol
    If lIndex > 0 Then
        DropDownSQL(lIndex) = sSql
    End If

End Sub

Public Function GetSQL(ByVal nCol As Integer) As String
    Dim lIndex As Long
    Dim lColIndex As Long
    
    fnGetKeys lIndex, lColIndex, nCol
    If lIndex > 0 Then
        GetSQL = DropDownSQL(lIndex)
    End If

End Function

Public Sub SetUnique(nCol As Integer, bStatus As Boolean)
    Dim lIndex As Long
    Dim lColIndex As Long
    
    fnGetKeys lIndex, lColIndex, nCol
    If lIndex > 0 Then
        colDropDowns.Item(lIndex).Item(1).Unique = bStatus
    End If

End Sub

Property Let ShowCaption(nCol As Integer, bFlag As Boolean)
    Dim lIndex As Long
    Dim lColIndex As Long
    
    fnGetKeys lIndex, lColIndex, nCol
    If lIndex > 0 Then
        colDropDowns.Item(lIndex).Item(1).ShowCaption = bFlag
    End If
End Property

Public Sub ShowNoDataMessage(nCol As Integer, bStatus As Boolean)
    Dim lIndex As Long
    Dim lColIndex As Long
    
    fnGetKeys lIndex, lColIndex, nCol
    If lIndex > 0 Then
        colDropDowns.Item(lIndex).Item(1).NoDataMessage = bStatus
    End If

End Sub

Private Function fnEditorLocked() As Boolean
    If Not m_clsEdit Is Nothing Then
        If m_clsEdit.Locked Then
            fnEditorLocked = True
        Else
            fnEditorLocked = tgMain.Columns(tgMain.col).Locked
        End If
    End If
End Function

Private Sub subSetFlagSelection(ByVal lDropDown As Long, _
                                bStatus As Boolean)
    Dim i As Integer
    m_nSeries = lDropDown
    If m_nSeries > 0 Then
'        colDropDowns.Item(m_nSeries).Item(1).Status = bValidSelection
        colDropDowns.Item(m_nSeries).Item(1).Status = bStatus
    End If
End Sub

Private Function fnCoulmnOrder(nDrpDn As Long, nCol As Integer) As Integer
    Dim colTemp As Collection
    Dim stCol As clsColumns
    Dim i As Integer

    fnCoulmnOrder = 0
    Set colTemp = colDropDowns.Item(nDrpDn)
    For i = 1 To colTemp.Count
        If colTemp.Item(i).ColumnIndex = nCol Then
            fnCoulmnOrder = i
            Exit For
        End If
    Next
End Function


Property Get COLUMN_TYPE_INTEGER() As Integer
    COLUMN_TYPE_INTEGER = TYPE_INTEGER
End Property


Private Sub fnResizeTable(tgTable As TDBGrid)
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

Public Sub GetColumns(nTemp() As Integer, nCount As Integer, ByVal nColumn As Integer)
    Dim lDropDown As Long
    Dim lColIndex As Long
    Dim colTemp As Collection
    Dim i As Integer

    fnGetKeys lDropDown, lColIndex, nColumn
    If lDropDown > 0 Then
        Set colTemp = colDropDowns.Item(lDropDown)
        If colTemp.Count > 0 Then
            ReDim nTemp(colTemp.Count - 1)
            nCount = 0
            For i = 1 To colTemp.Count
                If colTemp.Item(i).ColumnIndex < EXTRACOLUMNOFFSET Then
                    nTemp(nCount) = colTemp.Item(i).ColumnIndex
                    nCount = nCount + 1
                End If
            Next
        End If
    End If
End Sub

Public Sub KeyDown(tgTable As Object, KeyAscii As Integer)
'    If KeyAscii = vbKeyDown Or KeyAscii = vbKeyUp Or KeyAscii = vbKeyPageDown Or KeyAscii = vbKeyPageUp Then
'        bNavigating = True
'        cmdSearch.Visible = False
'        Exit Sub
'    End If
    Dim sKey As String
    
    If bKeyPressed Then
        KeyAscii = 0
        Exit Sub
    End If
    bKeyPressed = True
    If tgTable.TabIndex = tgDropDown.TabIndex Then
        If frmParent.ActiveControl Is tgMain Then
            On Error Resume Next
            tgMain.SetFocus
            Select Case KeyAscii
                Case vbKeyReturn
                    sKey = "~"
                Case vbKeyLeft
                    sKey = "{LEFT}"
                Case vbKeyRight
                    sKey = "{RIGHT}"
                Case vbKeyUp
                    sKey = "{UP}"
                Case vbKeyDown
                    sKey = "{DOWN}"
            End Select
            SendKeys sKey
        Else
            If KeyAscii = vbKeyReturn Then
                bShowList = False
                fnGetText tgMain.col
                subInputSelected
            ElseIf KeyAscii = vbKeyEscape Then
                bShowList = False
                tgDropDown.Visible = False
                On Error Resume Next
                tgMain.SetFocus
            End If
        End If
    ElseIf tgTable.TabIndex = tgMain.TabIndex Then
        If bSearchOnReturn Then
            If KeyAscii = vbKeyReturn Then
                If Not IsNull(tgMain.Row) Then
                    If Val(tgMain.Row) >= 0 Then
                        If fnLoadTable(nLastColumn) Then
                            KeyAscii = 0
                        End If
                    End If
                End If
                subInputSelected
            End If
        End If
    End If
    bKeyPressed = False
End Sub

Public Sub Keypress(tgTable As Object, KeyAscii As Integer)
    If m_bLocked Then
        Exit Sub
    End If
    If Not t_dbMainDatabase Is Nothing Then
        If frmParent.ActiveControl Is tgMain And tgTable Is tgDropDown Then
            'Work arround for true grid problems
            On Error Resume Next
            If IsNumeric(Chr(KeyAscii)) Then
                tgMain.SetFocus
                SendKeys KeyAscii - &H30, True
            ElseIf (KeyAscii >= 32 And KeyAscii <= 47) Or (KeyAscii >= 58 And KeyAscii <= 127) Then
                tgMain.SetFocus
                SendKeys "{" & UCase(Chr(KeyAscii)) & "}", True
            End If
        Else
            KeyDown tgTable, KeyAscii
        End If
    End If
End Sub

Public Sub RemoveDropDown(Optional nKey As Variant)
    Dim colColumn As New Collection
    Dim nCurrDropDown As Integer

    If IsMissing(nKey) Then
        While colDropDowns.Count > 0
            colDropDowns.Remove 1
        Wend
    Else
        On Error GoTo errRemoveDropDown
        colDropDowns.Remove nKey
    End If
extRemoveDropDown:
    On Error GoTo 0
    Exit Sub
errRemoveDropDown:
    MsgBox Err.Description
    Resume extRemoveDropDown

End Sub

Public Sub Resize()
    Dim lDropDown As Long
    Dim lColIndex As Long
    Dim nColumn As Integer
    Dim colTemp As Collection

'    fnResizeTable tgMain

    nMainTableTop = -1
    DoEvents
    fnLocateMainTable
    If tgDropDown.Visible Then
        fnGetKeys lDropDown, lColIndex, tgMain.col
        If lDropDown < 0 Then
            Exit Sub
        End If
    
        Set colTemp = colDropDowns.Item(lDropDown)
        fnLocateDropDownTable colTemp
    End If
    If cmdSearch.Visible Then
        fnMoveButton
    End If
End Sub


Public Sub AddColumn(nIndex As Integer, _
                     Optional sField As Variant, _
                     Optional nColType As Variant, _
                     Optional nColWidth As Variant, _
                     Optional nDropDown As Variant)

    Dim i As Integer
    Dim j As Integer
    Dim stCol As New clsColumns
    Dim colTemp As Collection
    
    If Not IsMissing(nDropDown) Then
        nCurrentDropDown = nDropDown
    End If
    If fnKeyExist(colDropDowns, nCurrentDropDown) Then
        Set colTemp = colDropDowns.Item(Str$(nCurrentDropDown))
    Else
        Set colTemp = New Collection
        colDropDowns.Add colTemp, Str$(nCurrentDropDown)
    End If
    If fnKeyExist(colTemp, nIndex) Then
        Set stCol = colTemp.Item(Str$(nIndex))
    Else
        If fnKeyExist(colTemp, sINITIAL) Then
            Set stCol = colTemp.Item(Str$(sINITIAL))
            colTemp.Remove Str$(sINITIAL)
        End If
        colTemp.Add stCol, Str$(nIndex)
    End If
    
    stCol.ColumnIndex = nIndex
    If Not IsMissing(sField) Then
        stCol.DataField = sField
    Else
        stCol.DataField = ""
    End If

    If IsMissing(nColType) Then
        stCol.ColType = TYPE_STRING
    Else
        stCol.ColType = nColType
    End If
    
    If IsMissing(nColWidth) Then
        stCol.ColWidth = 0
    Else
        stCol.ColWidth = nColWidth
    End If
End Sub

Property Get DropDownSQL(ByVal nIndex As Integer) As String
    Dim colColumn As Collection

    If nIndex > 0 And nIndex <= colDropDowns.Count Then
        Set colColumn = colDropDowns.Item(nIndex)
        DropDownSQL = colColumn.Item(1).SQL
    End If
End Property
Property Get DropDownType(nIndex As Integer) As Integer
    Dim colColumn As Collection

    If nIndex > 0 And nIndex <= colDropDowns.Count Then
        Set colColumn = colDropDowns.Item(nIndex)
        DropDownType = colColumn.Item(1).DropDownType
    End If
End Property

Property Get ColumnType(nIndex As Integer) As Integer
    Dim lDropDown As Long
    Dim lColumn As Long
    Dim colColumn As Collection
    
    fnGetKeys lDropDown, lColumn, nIndex
    If lDropDown > 0 Then
        Set colColumn = colDropDowns.Item(lDropDown)
        ColumnType = colColumn.Item(lColumn).ColType
    End If
End Property

Property Let ColumnType(nIndex As Integer, nType As Integer)
    Dim lDropDown As Long
    Dim lColumn As Long
    Dim colColumn As Collection
    
    fnGetKeys lDropDown, lColumn, nIndex
    If lDropDown > 0 Then
        Set colColumn = colDropDowns.Item(lDropDown)
        colColumn.Item(lColumn).ColType = nType
    End If
End Property
Property Set DataBase(dbSource As Object)
    Set dbMain = dbSource
    If nDatabase < 0 Then
        nDatabase = 1
    End If
End Property

Property Set DataLink(dLink As Object)
    Set datLink = dLink
End Property

Property Set DropDownTable(tblGrid As Object)
    
    Const szROW_STYLE As String = "RowStyle"
    Const szCURRENT_CELL_STYLE As String = "CurrCellStyle"
    
    Dim szBuf As String
    Dim n As Integer
    Dim m_styRow As Object            ' row style
    Dim m_styCurrentCell As Object    ' current cell style
    
    On Error GoTo AddToCollectionError
    
    tblGrid.Top = 1200
    If Not frmParent Is Nothing Then
        frmParent.Refresh
    End If
    
    Set tgDropDown = tblGrid

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
        
    For n = 0 To tblGrid.Columns.Count - 1
        With tblGrid.Columns(n)
            .DividerStyle = dbgBlackLine
            .Locked = False
            .AllowSizing = True
            .Visible = True
            .WrapText = False
            .Button = False
            .FetchStyle = False
        End With
    Next
  
  ' set default properties
    With tblGrid
        .AllowAddNew = False
        .AllowDelete = False
        .AllowUpdate = False
        .Enabled = True
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
        .Visible = False
    End With
    
    ' add the new styles to the table
    tblGrid.AddCellStyle dbgMarqueeRow, m_styRow
    tblGrid.AddCellStyle dbgMarqueeRow + dbgCurrentCell, m_styCurrentCell
    
    On Error GoTo 0
    
    fnLocateMainTable
    
    Set m_styRow = Nothing
    Set m_styCurrentCell = Nothing

    Exit Property
    
AddToCollectionError:
    
    On Error GoTo 0

    Set m_styRow = Nothing
    Set m_styCurrentCell = Nothing
End Property

Private Function fnGenetateSearchString(sField As String, sPrefix As String) As String
    Dim sTemp As String
    Dim sTrimedPrefix As String
    Dim nLenNumber As Integer
    Dim nloop As Integer
    Dim nxloop As Integer
    
    sTrimedPrefix = CStr(Val(sPrefix))
    sTemp = "(" & sField & " = " & sTrimedPrefix & ")"
    If Val(sPrefix) > 0 Then
        nLenNumber = Len(sTrimedPrefix)
        For nloop = 1 To 9 - nLenNumber
            sTemp = sTemp & " OR (" & sField & " >= " & sTrimedPrefix
            For nxloop = 1 To nloop
                 sTemp = sTemp + "0"
            Next
            
            sTemp = sTemp & " AND " & sField & " <= " & sTrimedPrefix
        
            For nxloop = 1 To nloop
                 sTemp = sTemp & "9"
            Next
            
            sTemp = sTemp & ")"
        Next
    End If
    fnGenetateSearchString = sTemp
End Function

Public Sub ButtonClick(ParamArray vMoreFields() As Variant)

    fnLoadTable tgMain.col
    subInputSelected
End Sub


Public Function fnGetRecord(rsTemp As Recordset, strSQL As String, nDB As Integer)
    ' Get records from the given SQL statement
    ' nDB = 1 ---> Informax Database (remote)
    '     = 2 ---> Access Database (local)

    On Error GoTo SQLError
    Select Case nDB
        Case DB_INFORMIX
            Set rsTemp = dbMain.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
        Case DB_ACCESS
            Set rsTemp = dbMain.OpenRecordset(strSQL, dbOpenSnapshot)
        Case Else
            fnGetRecord = False
            MsgBox "Drop down database not setup", vbOKOnly + vbCritical, frmParent.Caption
            Exit Function
    End Select
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveLast
        rsTemp.MoveFirst
    End If
    
    fnGetRecord = True
quitsub:
    On Error GoTo 0
    Exit Function
SQLError:
    subShowODBCError "", strSQL
    Screen.MousePointer = vbDefault
    fnGetRecord = False
    Resume quitsub
End Function

Private Sub fnGetText(nCol As Integer)
    Dim i As Integer
    Dim lDropDown As Long
    Dim lColumn As Long
    Dim colColumn As Collection
    
    If datLink.Recordset Is Nothing Then
        Exit Sub
    End If
    fnGetKeys lDropDown, lColumn, nCol

    If lDropDown > 0 Then
        If Not colDropDowns.Item(lDropDown).Item(1).PopulateBox Then
            Exit Sub
        End If
        bValidSelection = True
        subSetFlagSelection lDropDown, True
        Set colColumn = colDropDowns.Item(lDropDown)
        If Not m_clsEdit Is Nothing Then
            m_clsEdit.AddingValidate = True
        End If
        For i = 1 To colColumn.Count
            With colColumn.Item(i)
                If .ColumnIndex < EXTRACOLUMNOFFSET Then
                    If Not m_clsEdit Is Nothing Then
                        m_clsEdit.BeforeColEdit .ColumnIndex, 0, False
                    End If
                    If IsNull(datLink.Recordset(.DataField)) Then
                        tgMain.Columns(.ColumnIndex).Text = ""
                    Else
                        On Error GoTo errGetText
                        tgMain.Columns(.ColumnIndex).Text = Trim$(datLink.Recordset(.DataField))
                        On Error GoTo 0
                    End If
                    If Not m_clsEdit Is Nothing Then
                        m_clsEdit.AfterColEdit .ColumnIndex
                    End If
                End If
            End With
        Next i
        DoEvents
        If Not m_clsEdit Is Nothing Then
            m_clsEdit.AddingValidate = False
        End If
        tgDropDown.Visible = False
        m_nCurrentDropDown = lDropDown
'        subSetFlagSelection lDropDown, True
        tgDropDown.ZOrder 1
        On Error Resume Next
        tgMain.SetFocus
    End If
    Exit Sub

errGetText:
    tgMain.Columns(colColumn.Item(i).ColumnIndex).Text = ""
    Err.Clear
    Resume Next
End Sub

Private Sub fnMoveButton()
    Dim lDropDown As Long
    Dim lCol As Long
    Dim lRow As Long
    Dim Factor As Single
    Dim nLeft As Integer
    Dim nVisibleRows As Integer
    
    If fnEditorLocked Then
        cmdSearch.Visible = False
        Exit Sub
    End If

    fnGetKeys lDropDown, lCol, tgMain.col
    
    If lDropDown <= 0 Then
        Exit Sub
    End If
    
    If Not colDropDowns.Item(lDropDown).Item(1).Visible Then
        If Not (cmdSearch Is Nothing) Then
            cmdSearch.Visible = False
        End If
        Exit Sub
    End If
    nVisibleRows = -1
    If tgMain.ScrollBars = 2 Or tgMain.ScrollBars = 4 Then
        If Not m_clsEdit Is Nothing Then
            nVisibleRows = tgMain.Height / tgMain.RowHeight - tgMain.HeadLines
            If tgMain.Caption <> "" Then
                nVisibleRows = nVisibleRows - 1
            End If
            If IsNull(tgMain.Row) Then
                lRow = 0
            Else
                lRow = Val(tgMain.Row)
            End If
            If lRow >= nVisibleRows Or lRow < 0 Then
                cmdSearch.Visible = False
                Exit Sub
            End If
        End If
    End If
    
    If nVisibleColumns < 0 Then
        nVisibleColumns = fnVisibleColumns(tgMain.Columns.Count - 1)
    End If
    
    If tgMain.Row >= 0 Then
        Dim objCol As Column
        Dim colTemp As Collection
        Set colTemp = colDropDowns.Item(lDropDown)
        Set objCol = tgMain.Columns(colTemp.Item(lCol).ColumnIndex)
        If Not (cmdSearch Is Nothing) Then
            cmdSearch.Height = tgMain.RowHeight - 2 * Screen.TwipsPerPixelY
            On Error Resume Next
            Select Case colTemp.Item(1).DropDownType
                Case TYPE_DRPDN_FACTORSEARCH
                    cmdSearch.Width = tgMain.RowHeight
                    cmdSearch.Picture = frmContext.LoadPicture(SEARCH_UP)
                Case TYPE_DRPDN_COMMONDROPDOWN
                    cmdSearch.Width = 0.6 * tgMain.RowHeight
                    cmdSearch.Picture = frmContext.LoadPicture(DROPDOWN_UP)
                Case TYPE_DRPDN_NO_CRITERIA
                    cmdSearch.Width = 0.6 * tgMain.RowHeight
                    cmdSearch.Picture = frmContext.LoadPicture(DROPDOWN_UP)
                Case TYPE_DRPDN_NO_CRITERIA2
                    cmdSearch.Width = 0.6 * tgMain.RowHeight
                    cmdSearch.Picture = frmContext.LoadPicture(SEARCH_UP)
            End Select
            On Error GoTo 0
            If tgMain.col < tgMain.LeftCol Or (tgMain.col > tgMain.LeftCol And objCol.Left = 0) Or objCol.Left + objCol.Width >= tgMain.Width Then
                cmdSearch.Visible = False
                cmdSearch.ZOrder 1
            Else
                cmdSearch.Top = tgMain.Top + tgMain.RowTop(tgMain.Row) + Screen.TwipsPerPixelY
                If fnVisibleColumns(colTemp.Item(lCol).ColumnIndex) >= nVisibleColumns Then
                    If tgMain.ExtendRightColumn Then
                        nLeft = tgMain.Left + tgMain.Width - cmdSearch.Width
                    Else
                        nLeft = tgMain.Left + objCol.Left + objCol.Width - cmdSearch.Width
                    End If
                    If nVisibleRows > 0 Then
                        If m_clsEdit.RowCount >= nVisibleRows Then
                            nLeft = nLeft - SCROLLBAR_WIDTH
                        End If
                    End If
                Else
                    nLeft = tgMain.Left + objCol.Left + objCol.Width - cmdSearch.Width
                End If
                cmdSearch.Left = nLeft
                If tgMain.Enabled Then
                    cmdSearch.Visible = True
                    If Not frmParent.ActiveControl Is tgMain Then
                        tgMain.SetFocus
                    End If
                End If
                cmdSearch.ZOrder 0
            End If
        End If
    End If
End Sub

Public Sub GotFocus(Optional vCtrl As Variant)
    If m_bLocked Then
        Exit Sub
    End If
    If IsMissing(vCtrl) Then
        fnMoveButton
        nLastColumn = tgMain.col
        If Not m_clsEdit Is Nothing Then
            m_LastRow = m_clsEdit.GetCurrentRowNumber
        End If
    Else
        Dim objCtrl As Object
        Set objCtrl = vCtrl
        If objCtrl Is tgMain Then
            fnMoveButton
            nLastColumn = tgMain.col
            If Not m_clsEdit Is Nothing Then
                m_LastRow = m_clsEdit.GetCurrentRowNumber
            End If
        ElseIf objCtrl Is tgDropDown Then
            If tgDropDown.Visible = False Then
                DoEvents
                On Error Resume Next
                tgMain.SetFocus
            End If
        ElseIf objCtrl Is cmdSearch Then
            DoEvents
            On Error Resume Next
            tgMain.SetFocus
        End If
    End If
End Sub


Public Sub LostFocus(objControl As Object)
    Dim nTab As Integer
    On Error Resume Next
    
    If m_bLocked Then
        Exit Sub
    End If
    
    nTab = frmParent.ActiveControl.TabIndex

    If objControl.TabIndex = tgMain.TabIndex Then
        If nTab <> cmdSearch.TabIndex And nTab <> tgDropDown.TabIndex And nTab <> tgMain.TabIndex Then
            cmdSearch.Visible = False
            cmdSearch.ZOrder 1
        Else
            If nTab = tgDropDown.TabIndex Then
                If Not tgDropDown.Visible Then
                    DoEvents
                    tgMain.SetFocus
                End If
            End If
        End If
    ElseIf objControl.TabIndex = tgDropDown.TabIndex Then
        If bShowList And Not bAutoHide Then
            tgDropDown.SetFocus
        Else
            tgDropDown.Visible = False
            tgMain.SetFocus
        End If
    ElseIf objControl.TabIndex = cmdSearch.TabIndex Then
        If nTab <> tgMain.TabIndex And nTab <> tgDropDown.TabIndex Then
            cmdSearch.Visible = False
            tgDropDown.Visible = False
            cmdSearch.ZOrder 1
        End If
    End If
End Sub

Public Sub MouseUp(y As Single)
    If m_bLocked Then
        Exit Sub
    End If
    nClickPointY = y
End Sub
Property Set Form(frmMain As Object)
    Set frmParent = frmMain
    fnLocateMainTable
    If Not tgDropDown Is Nothing Then
        frmParent.Refresh
    End If
End Property
Public Sub Scroll()
    If m_bLocked Then
        Exit Sub
    End If
End Sub

Property Set SearchButton(btn As Object)
    Set cmdSearch = btn
    cmdSearch.TabStop = False
    btn.Visible = False
    btn.ZOrder 1
End Property

Property Set MainTable(tblGrid As Object)
'
' Add a True DBGrid to the collection
'
    Set tgMain = tblGrid
    nOriginalTableWidth = tgMain.Width
    nOriginalTableHeight = tgMain.Height
    nFontSize = tgMain.Font.Size

    fnLocateMainTable
End Property

Private Function fnColumnNumber(nIndex) As Integer
    Dim i As Integer
    
    fnColumnNumber = -1
    For i = 0 To nColumnCount - 1
'        If nColumns(i) = nIndex Then
'            fnColumnNumber = i
'            Exit Function
'        End If
    Next
End Function


Public Function RowColChange(tgTable As Object) As Boolean
    RowColChange = False
'    If bNavigating Then
'        Exit Function
'    End If
    
    Dim lDropDown As Long
    Dim lCol As Long
    Dim lOldRow As Integer
    Dim lLastRow As Long
    
    
    If m_bLocked Then
        Exit Function
    End If
    If frmParent.ActiveControl Is Nothing Then
        Exit Function
    End If
    
    fnGetKeys lDropDown, lCol, tgMain.col
    
    RowColChange = True
    If tgTable.TabIndex = tgDropDown.TabIndex Then
        If bTableClicked Then
            If lDropDown >= 0 Then
                bShowList = False
                lOldRow = tgMain.Row
                fnGetText tgMain.col
                DoEvents
                If lOldRow = tgMain.Row Then
                    subInputSelected
                End If
            End If
            bTableClicked = False
        End If
    ElseIf tgTable.TabIndex = tgMain.TabIndex Then
        If frmParent.ActiveControl.TabIndex <> tgMain.TabIndex Then
            If frmParent.ActiveControl.TabIndex <> cmdSearch.TabIndex Then
                Exit Function
            End If
        End If
        If lDropDown < 0 Then
            cmdSearch.Visible = False
            tgDropDown.Visible = False
            If m_nSeries > 0 Then
                colDropDowns.Item(m_nSeries).Item(1).Status = False
            End If
        Else
            fnMoveButton
    
            If colDropDowns.Item(lDropDown).Item(1).CursorControl Then
                If m_nSeries = lDropDown And nLastColumn <= tgMain.col Then
                    If colDropDowns.Item(lDropDown).Item(1).Status Then
                        If Not m_clsEdit Is Nothing Then
                            If m_LastRow = m_clsEdit.GetCurrentRowNumber Then
                                lLastRow = m_clsEdit.GetCurrentRowNumber
                                m_clsEdit.GotoNextColumn nLastColumn
                                DoEvents
                                If lLastRow <> m_clsEdit.GetCurrentRowNumber Then
                                    colDropDowns.Item(lDropDown).Item(1).Status = False
                                End If
                                On Error Resume Next
                                tgMain.SetFocus
                            End If
                            RowColChange = False
                        End If
                    End If
                Else
                    If m_nSeries > 0 Then
                        colDropDowns.Item(m_nSeries).Item(1).Status = False
                        'RowColChange = False
                    End If
                End If
            End If
        End If
        nLastColumn = tgMain.col
        If Not m_clsEdit Is Nothing Then
            m_LastRow = m_clsEdit.GetCurrentRowNumber
        End If
        m_nSeries = lDropDown
    End If
End Function
Property Let DropDownSQL(ByVal nIndex As Integer, SQL As String)
    Dim colColumn As Collection

    If nIndex > 0 And nIndex <= colDropDowns.Count Then
        Set colColumn = colDropDowns.Item(nIndex)
        colColumn.Item(1).SQL = SQL
    End If
End Property


Property Let DropDownType(nIndex As Integer, nType As Integer)
    Dim colColumn As Collection

    If nIndex > 0 And nIndex <= colDropDowns.Count Then
        Set colColumn = colDropDowns.Item(nIndex)
        colColumn.Item(1).DropDownType = nType
    End If
End Property

Property Get COLUMN_TYPE_STRING() As Integer
    COLUMN_TYPE_STRING = TYPE_STRING
End Property

Public Sub SetExtendRight(Optional vIndex As Variant)
    Dim nIndex As Integer
    Dim colColumn As Collection
    
    If IsMissing(vIndex) Then
        nIndex = nCurrentDropDown
    Else
        nIndex = Val(vIndex)
    End If

    If nIndex > 0 And nIndex <= colDropDowns.Count Then
        Set colColumn = colDropDowns.Item(nIndex)
        colColumn.Item(1).ExtDirection = EXTRA_COL_RIGHT
    End If
    
End Sub

Public Sub SetExtendLeft(Optional vIndex As Variant)
    Dim nIndex As Integer
    Dim colColumn As Collection
    
    If IsMissing(vIndex) Then
        nIndex = nCurrentDropDown
    Else
        nIndex = Val(vIndex)
    End If

    If nIndex > 0 And nIndex <= colDropDowns.Count Then
        Set colColumn = colDropDowns.Item(nIndex)
        colColumn.Item(1).ExtDirection = EXTRA_COL_LEFT
    End If
    
End Sub
Public Sub TableClick(tgTable As Object)
    If m_bLocked Then
        Exit Sub
    End If
    If tgTable.TabIndex = tgDropDown.TabIndex Then
        Dim nRowTop As Integer
        Dim nRowBtm As Integer
        bShowList = False
        If tgDropDown.Row >= 0 Then
            nRowTop = tgDropDown.RowTop(tgDropDown.Row)
            nRowBtm = nRowTop + tgDropDown.RowHeight
            If nClickPointY >= nRowTop And nClickPointY <= nRowBtm Then
                fnGetText tgMain.col
                On Error Resume Next
                tgMain.SetFocus
                subInputSelected
            Else
                bTableClicked = True
            End If
        Else
            bTableClicked = True
        End If
    ElseIf tgTable.TabIndex = tgMain.TabIndex Then
        fnMoveButton
    End If
End Sub

Property Get ValidSelection() As Boolean
    ValidSelection = bValidSelection
    bValidSelection = False
'    subSetFlagSelection m_nSeries
End Property
Property Let ValidSelection(bStatus As Boolean)
    bValidSelection = bStatus
End Property

Property Let Visible(nIndex As Integer, bStatus As Boolean)
    
    Dim colColumn As Collection

    If nIndex > 0 And nIndex <= colDropDowns.Count Then
        Set colColumn = colDropDowns.Item(nIndex)
        colColumn.Item(1).Visible = bStatus
        If Not bStatus Then
            cmdSearch.Visible = False
        End If
    End If

End Property
Public Sub SetVisible(nCol As Integer, bStatus As Boolean)
    Dim lIndex As Long
    Dim lColIndex As Long
    
    fnGetKeys lIndex, lColIndex, nCol
    If lIndex > 0 Then
        colDropDowns.Item(lIndex).Item(1).Visible = bStatus
        If tgMain.col = nCol Then
            If bStatus Then
                fnMoveButton
                cmdSearch.Visible = True
            Else
                cmdSearch.Visible = False
            End If
        End If
    End If

End Sub

Property Let CursorControl(nIndex As Integer, bStatus As Boolean)
    
    If nIndex > 0 And nIndex <= colDropDowns.Count Then
        colDropDowns.Item(nIndex).Item(1).CursorControl = bStatus
    End If

End Property

Private Sub subInputSelected()
    Dim i As Integer
    Dim colTemp As Collection
    Dim lLastRow As Long
    
    DoEvents
    If m_nSeries > 0 Then
        Set colTemp = colDropDowns.Item(m_nSeries)
        If colTemp.Item(1).Status Then
            If Not m_clsEdit Is Nothing Then
                lLastRow = m_clsEdit.GetCurrentRowNumber
                If colTemp.Item(1).CursorControl Then
                    m_clsEdit.GotoNextColumn nLastColumn
                    DoEvents
                    If lLastRow = m_clsEdit.GetCurrentRowNumber Then
                        For i = 1 To colTemp.Count
                            If colTemp.Item(i).ColumnIndex = m_clsEdit.ColumnForNewRow Then
                                m_clsEdit.subAddNewRow m_clsEdit.ColumnForNewRow
                                If m_nSeries > 0 Then
                                    colDropDowns.Item(m_nSeries).Item(1).Status = False
                                End If
                                Exit For
                            End If
                        Next
                    End If
                    On Error Resume Next
                    tgMain.SetFocus
                End If
            End If
        End If
    End If
End Sub


Property Let VisibleRows(nRows As Integer)
    nMaxVisibleRows = nRows
End Property

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
    
    If Err.Number = 3146 Then
        With t_engFactor.Errors
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
    Else
        sODBCErrors = Err.Description
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


Private Sub Class_Initialize()
    Set colDropDowns = New Collection
    nMainTableTop = -1
    nMaxVisibleRows = -1
    nVisibleColumns = -1
    nDatabase = -1
    m_bDefaultCursorControl = False
    bKeyPressed = False
    bSearchOnReturn = False
    bAutoHide = True
    m_bLocked = False
End Sub
Private Sub Class_Terminate()
    Dim colTemp As clsColumns
    While colDropDowns.Count > 0
        colDropDowns.Remove 1
    Wend
    Set colDropDowns = Nothing
    Set tgMain = Nothing
    Set tgDropDown = Nothing
    Set lblStatusbar = Nothing
    Set frmParent = Nothing
    Set cmdSearch = Nothing
    Set datLink = Nothing
    Set dbMain = Nothing
    Set m_clsEdit = Nothing
End Sub


