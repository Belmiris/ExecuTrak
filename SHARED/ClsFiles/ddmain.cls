VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDropDown"
Attribute VB_Creatable = True
Attribute VB_Exposed = True
Option Explicit
    
    Private Const TYPE_STRING As Integer = 1
    Private Const TYPE_LONG As Integer = 2
    Private Const TYPE_DATE As Integer = 3

    Private Const FACTORSEARCH = 1
    Private Const COMMONDROPDOWN = 2
    
    Private Const DB_INFORMIX As Integer = 1
    Private Const DB_ACCESS As Integer = 2
    
    Private Const nDROP_DOWN As Integer = 1
    Private Const nDROP_UP As Integer = 2

    Private Const CONTROL_TYPE_TEXTBOX As Integer = 1
    Private Const CONTROL_TYPE_BUTTON As Integer = 2
    
    Private Const nVISIBLEROWS_DEFAULT As Integer = 7
    Private Const SCROLLBAR_WIDTH As Integer = 264

    Private Const EXTRACOLUMNOFFSET = 1000
    Private Const KEY_CTRL_Y = &H19
    Private Const sINITIAL As Integer = -1
    
    Private Const nHEIGHTOFFSET = 3

    Private Const sSTRING_PATTERN As String = "P"
    Private Const sLONG_PATTERN As String = "^(#{0,10}|#{1,10}[:\-]#{0,10})$"

    Private Type POINTAPI 'Point structure
        X As Integer
        Y As Integer
    End Type
    
    Private Type RECT 'used for tooltip and context menus
        Left As Integer
        Top As Integer
        Right As Integer
        Bottom As Integer
    End Type
        
    Private nCurrentDD As Integer
    Private colComboControls As Collection

    Private tgDropDown As Object
    Private frmParent As Object

    Private bTableClicked As Boolean
    Private nClickPointY As Single
    Private bValidSelection As Boolean
    Private ListBuffer() As Variant
    Private nMaxRows As Integer
    Private m_nSelectedRow As Integer

    Private Declare Function SetParent Lib "USER" (ByVal hWndChild As Integer, ByVal hWndNewParent As Integer) As Integer
    Private Declare Function GetParent Lib "USER" (ByVal hWnd As Integer) As Integer
    Private Declare Sub GetWindowRect Lib "USER" (ByVal hWnd As Integer, lpRect As RECT)
    Private Declare Function ScreenToClient Lib "USER" (ByVal hWnd As Integer, lpPoint As POINTAPI) As Boolean

Public Sub AddColumn(txtBox As Object, Optional vWidth As Variant)
    Dim nDropDown As Integer
    Dim nType As Integer
    Dim clsDD As clsDDSingle
    Dim nWidth As Integer
    
    If IsMissing(vWidth) Then
        nWidth = txtBox.Width
    Else
        nWidth = vWidth
    End If

    fnGetKeys nDropDown, nType, txtBox
    
    If nDropDown > 0 Then
        Set clsDD = colComboControls.Item(nDropDown)
        clsDD.AddColumn nWidth
        clsDD.TableWidth = clsDD.TableWidth + nWidth
    End If

End Sub

Private Sub fnGetParsingColumns(nA() As Integer, sSrc As String)
    Dim nPos As Integer
    Dim sTemp As String
    Dim nCount As Integer
    
    sTemp = Trim$(sSrc)
    If sTemp = "" Then
        Exit Sub
    End If

    nCount = 0
    ReDim nA(nCount)
    Do
        nPos = InStr(sTemp, ",")
        If nPos > 0 Then
            ReDim Preserve nA(nCount)
            nA(nCount) = val(Left(sTemp, nPos - 1))
            nCount = nCount + 1
            sTemp = Trim$(Right(sTemp, Len(sTemp) - nPos))
        Else
            If sTemp <> "" Then
                ReDim Preserve nA(nCount)
                nA(nCount) = val(sTemp)
            End If
        End If
    Loop Until nPos = 0
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


Public Sub ReadData(ByVal RowBuf As RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
    Dim Bookmk As Variant
    
    If colComboControls.count = 0 Then
        Exit Sub
    End If
    If nMaxRows = 0 Then
        Exit Sub
    End If

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
                RowBuf.Value(i%, j%) = ListBuffer(j%, lRow)
            End If
        Next j%
        RowBuf.bookmark(i%) = Bookmk
        RowsFetched = RowsFetched + 1
    Next i%
    RowBuf.RowCount = RowsFetched
End Sub


Property Let Position(nType As Integer)
    colComboControls.Item(nCurrentDD).Position = nType
End Property



Property Get DropUp() As Integer
    DropUp = nDROP_UP
End Property

Property Let RegExpPattern(objCOntrol As Object, sPattern As String)
    Dim lDropDown As Integer
    Dim nctrlType As Integer
    Dim clsTemp As clsDDSingle

    fnGetKeys lDropDown, nctrlType, objCOntrol
    If lDropDown < 0 Then
        Exit Property
    End If
    
    Set clsTemp = colComboControls.Item(lDropDown)
    clsTemp.RegExpPattern = fnGetPCode(sPattern)
End Property

Property Get SelectedRow() As Integer
    SelectedRow = m_nSelectedRow
End Property

Public Sub SetColumnCaptions(ParamArray sCaps())

    Dim nCol As Integer
    Dim colTemp As Collection

    nCol = UBound(sCaps)
    If nCol < 0 Then
        Exit Sub
    End If

    Dim clsTemp As clsDDSingle
    Dim i As Integer

    Set clsTemp = colComboControls.Item(nCurrentDD)
    i = 0
    For i = 0 To nCol
        If i < colTemp.count Then
            clsTemp.Caption = sCaps(i)
        End If
    Next
    clsTemp.HeadLines = 1
    
End Sub

Public Function AddDropDown(Optional nRows As Variant, _
                            Optional nPos As Variant) As Integer

    Dim colCombo As New Collection
    Dim clsDD As New clsDDSingle

    If IsMissing(nRows) Then
        clsDD.VisibleRows = nVISIBLEROWS_DEFAULT
    Else
        clsDD.VisibleRows = nRows
    End If

    If IsMissing(nPos) Then
        clsDD.Position = nDROP_DOWN
    Else
        clsDD.Position = nPos
    End If
    
    On Error GoTo errAddDropDown
    clsDD.HeadLines = 0
    clsDD.ColumnCount = 0
    colComboControls.Add clsDD
    nCurrentDD = colComboControls.count
    AddDropDown = nCurrentDD

extAddDropDown:
    On Error GoTo 0
    Exit Function
errAddDropDown:
    MsgBox Err.Description
    Resume extAddDropDown
End Function

Public Sub Change(objCOntrol As Object)
    If objCOntrol Is Nothing Then
        Exit Sub
    End If
    
    If objCOntrol.TabIndex = tgDropDown.TabIndex Then
        Exit Sub
    End If

    Dim lDropDown As Integer
    Dim nctrlType As Integer

    fnGetKeys lDropDown, nctrlType, objCOntrol
    
    If lDropDown < 0 Then
        Exit Sub
    End If
    
    On Error Resume Next
    
    If nctrlType = CONTROL_TYPE_TEXTBOX Then
        
        Dim clsTemp As clsDDSingle
        
        Set clsTemp = colComboControls.Item(lDropDown)
        
        tfnRegExpControlChange objCOntrol, clsTemp.RegExpPattern
        
    End If

    On Error GoTo 0

End Sub

Private Sub fnGetOffsets(nLeft As Integer, _
                         nTop As Integer, _
                         txtBox As Object)

    'Find out the offsets of the text box in the table

    Dim nTbl_hWnd As Integer
    Dim n_hWnd As Integer
    Dim n_ParenthWnd As Integer
    Dim nClientLeft As Integer
    Dim nClientTop As Integer
    
    nTbl_hWnd = frmParent.hWnd
    n_hWnd = txtBox.hWnd
    n_ParenthWnd = GetParent(n_hWnd)
    nLeft = 0
    nTop = txtBox.Height
    While n_hWnd <> nTbl_hWnd And n_ParenthWnd > 0
        fnGetLeftTop nClientLeft, nClientTop, n_ParenthWnd, n_hWnd
        nLeft = nLeft + nClientLeft
        nTop = nTop + nClientTop
        n_hWnd = n_ParenthWnd
        n_ParenthWnd = GetParent(n_hWnd)
    Wend
End Sub
Private Function fnGetLeftTop(nLeft As Integer, nTop As Integer, nParenthWnd As Integer, n_hWnd)
    
    Dim myRect As RECT
    Dim myPoint As POINTAPI
    
    GetWindowRect n_hWnd, myRect

    myPoint.X = myRect.Left
    myPoint.Y = myRect.Top
    
    If ScreenToClient(nParenthWnd, myPoint) Then
'        nLeft = (myPoint.X - 1) * Screen.TwipsPerPixelX    'For True Gridtable
        nLeft = myPoint.X * Screen.TwipsPerPixelX     'For textbox
        If nLeft < -10000 Or nLeft > 10000 Then
            nLeft = Screen.TwipsPerPixelX
        End If
'        nTop = (myPoint.Y - 1) * Screen.TwipsPerPixelY    'For True Gridtable
        nTop = myPoint.Y * Screen.TwipsPerPixelY       'For textbox
        If nTop < 0 Or nTop > 10000 Then
            nTop = Screen.TwipsPerPixelX
        End If
    End If

End Function

Property Let CurrentDropDown(c As Integer)
    nCurrentDD = c
End Property
Property Get CurrentDropDown() As Integer
    CurrentDropDown = nCurrentDD
End Property


Property Get DDCount() As Integer
    DDCount = colComboControls.count
End Property

Private Sub fnGetKeys(nDropDown As Integer, nType As Integer, objKey As Object)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim colTemp As Collection
    
    For i = 1 To colComboControls.count
        With colComboControls.Item(i)
            If Not (.TextBox Is Nothing) Then
                If .TextBox.TabIndex = objKey.TabIndex Then
                    nDropDown = i
                    nType = CONTROL_TYPE_TEXTBOX
                    Exit Sub
                End If
                If .Button.TabIndex = objKey.TabIndex Then
                    nDropDown = i
                    nType = CONTROL_TYPE_BUTTON
                    Exit Sub
                End If
            End If
        End With
    Next i
'    MsgBox "Column not defined for drop down"
    nDropDown = -1
End Sub

Private Sub fnSetFocus(cntlTemp As Object)
    Const nTrialNumber As Integer = 1
    Dim nCount As Integer

    nCount = 0

    On Error GoTo errSetFocus
    cntlTemp.SetFocus
    
extSetFocus:
    On Error GoTo 0
    Exit Sub
    
errSetFocus:
    If nCount < nTrialNumber Then
        nCount = nCount + 1
        DoEvents
        Resume
    Else
        Resume extSetFocus
    End If
End Sub

Private Function fnLoadTable(ByVal lDropDown As Integer) As Boolean
    
    'Declare Sub's Variables
    Dim szCode As String
    Dim szQuery As String
    Dim sTemp As String
    Dim rsTemp As Object
    Dim szLink As String
    Dim i As Integer
    Dim nPos As Integer
    Dim lCol As Integer
    Dim clsDD As clsDDSingle

    'Make our mouse pointer an hourglass
    Screen.MousePointer = vbHourglass
    bValidSelection = False
    
    tgDropDown.Visible = False
    frmParent.Cls

    If lDropDown < 0 Then
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set clsDD = colComboControls.Item(lDropDown)
    nCurrentDD = lDropDown

    While tgDropDown.Columns.count > 0
        tgDropDown.Columns.Remove 0
    Wend
    For lCol = 1 To clsDD.ColumnCount
        tgDropDown.Columns.Add lCol - 1
        With tgDropDown.Columns(lCol - 1)
            .Width = clsDD.ColWidth(lCol - 1)
            .Caption = clsDD.Caption
            .DividerStyle = dbgNoDividers
            .Visible = True
        End With
    Next
    tgDropDown.HeadLines = clsDD.HeadLines

    tgDropDown.Left = clsDD.TableLeft
    tgDropDown.Width = clsDD.TableWidth + SCROLLBAR_WIDTH

    Dim nH As Integer
    If clsDD.RowCount < clsDD.VisibleRows Then
        nH = clsDD.RowCount
    Else
        nH = clsDD.VisibleRows
    End If
    tgDropDown.Height = (tgDropDown.HeadLines + nH) * tgDropDown.RowHeight + nHEIGHTOFFSET * Screen.TwipsPerPixelY
    
    If clsDD.Position = nDROP_UP Then
        tgDropDown.Top = clsDD.TableTop - clsDD.TextBox.Height - tgDropDown.Height
    Else
        tgDropDown.Top = clsDD.TableTop
    End If

    clsDD.GetArray ListBuffer
    nMaxRows = clsDD.RowCount

    tgDropDown.Refresh

    tgDropDown.ZOrder 0
    tgDropDown.Visible = True
    fnSetFocus tgDropDown
    fnLoadTable = True

    'Reset our mouse pointer to the default
    Screen.MousePointer = vbDefault
End Function
Public Sub SelChange(Cancel As Integer)
    Cancel = True
End Sub

Public Function KeyPress(objCOntrol As Object, KeyAscii As Integer) As Boolean
    If objCOntrol.TabIndex = tgDropDown.TabIndex Then
        If KeyAscii = vbKeyReturn Then
            fnSetFocus colComboControls.Item(nCurrentDD).TextBox
            fnGetText
        ElseIf KeyAscii = vbKeyEscape Then
            tgDropDown.Visible = False
            frmParent.Cls
            fnSetFocus colComboControls.Item(nCurrentDD).TextBox
        End If
    Else
        Dim lDropDown As Integer
        Dim nctrlType As Integer

        fnGetKeys lDropDown, nctrlType, objCOntrol
        If lDropDown < 0 Then
            Exit Function
        End If

        If nctrlType = CONTROL_TYPE_BUTTON Then
            fnLoadTable lDropDown
        Else
            If KeyAscii = vbKeyReturn Then
                fnLoadTable lDropDown
                KeyPress = False
            ElseIf KeyAscii = KEY_CTRL_Y Then
                ' wipe all text boxes' text
                If nctrlType = CONTROL_TYPE_TEXTBOX Then
                    objCOntrol.TextBox.Text = ""
                End If
                KeyAscii = 0
                KeyPress = True
            ElseIf KeyAscii = vbKeyEscape Then
                ' wipe the current text box's text
                objCOntrol.Text = ""
                KeyAscii = 0
                KeyPress = True
            Else
                ' validate text + key against the regular expression
                Dim clsTemp As clsDDSingle
                Set clsTemp = colComboControls.Item(lDropDown)
                KeyPress = tfnRegExpControlKeyPress(objCOntrol, KeyAscii, clsTemp.RegExpPattern)
            End If
        End If
    End If
End Function


Public Sub RemoveDropDown(Optional nKey As Variant)
    Dim colColumn As New Collection
    Dim nCurrDropDown As Integer

    If IsMissing(nKey) Then
        While colComboControls.count > 0
            colComboControls.Remove 1
        Wend
    Else
        On Error GoTo errRemoveDropDown
        colComboControls.Remove nKey
    End If
extRemoveDropDown:
    On Error GoTo 0
    Exit Sub
errRemoveDropDown:
    MsgBox Err.Description
    Resume extRemoveDropDown

End Sub


Public Sub AddDDCombo(txtBox As Object, _
                      cmdButton As Object, _
                      Optional nWidth As Variant, _
                      Optional sCaption As Variant, _
                      Optional nType As Variant, _
                      Optional nCombo As Variant)

    Dim nCurr As Integer
    Dim clsDD As clsDDSingle
    Dim colTemp As Collection
    
    If IsMissing(nCombo) Then
        nCurr = nCurrentDD
    Else
        nCurr = nCombo
    End If
    
    If nCurr > colComboControls.count Then
        Set clsDD = New clsDDSingle
        If colComboControls Is Nothing Then
            Set colComboControls = New Collection
        End If
        colComboControls.Add clsDD
        nCurr = colComboControls.count
    Else
        Set clsDD = colComboControls.Item(nCurr)
    End If
    
    Set clsDD.TextBox = txtBox
    Set clsDD.Button = cmdButton
    clsDD.ColumnCount = 0
    If Not IsMissing(sCaption) Then
        clsDD.Caption = sCaption
        clsDD.HeadLines = 1
    End If
    
    If IsMissing(nType) Then
        clsDD.ComboType = TYPE_STRING
        clsDD.RegExpPattern = fnGetPCode(TYPE_STRING)
    Else
        clsDD.ComboType = nType
        clsDD.RegExpPattern = fnGetPCode(nType)
    End If
    
    Dim nW As Integer
    If IsMissing(nWidth) Then
        nW = txtBox.Width
    Else
        nW = nWidth
    End If

    clsDD.AddColumn nW

    Dim nLeft As Integer
    Dim nTop As Integer
    fnGetOffsets nLeft, nTop, txtBox
    clsDD.TableLeft = nLeft  ' + txtBox.Left
    clsDD.TableTop = nTop    '+ txtBox.Height     '+ txtBox.Top
    clsDD.TableWidth = nW    'colTemp.Item(nBoxCount).TextBox.Width
    clsDD.VisibleRows = nVISIBLEROWS_DEFAULT
End Sub

Property Set Table(tblGrid As Object)
    
    
    Const szROW_STYLE As String = "RowStyle"
    Const szCURRENT_CELL_STYLE As String = "CurrCellStyle"
    
    Dim szBuf As String
    Dim n As Integer
    Dim m_styRow As Object            ' row style
    Dim m_styCurrentCell As Object    ' current cell style
    
    On Error GoTo AddToCollectionError
    
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
    
    ' add the new styles to the table
    tblGrid.AddCellStyle dbgMarqueeRow, m_styRow
    tblGrid.AddCellStyle dbgMarqueeRow + dbgCurrentCell, m_styCurrentCell
    
  ' set default properties
    With tblGrid
        .AllowAddNew = False
        .AllowDelete = False
        .AllowUpdate = False
        .Enabled = True
        .MarqueeUnique = True
        .EditDropDown = False
        .RowDividerStyle = dbgNoDividers
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

Public Sub Click(objCOntrol As Object, Optional v)
    
    If objCOntrol.TabIndex = tgDropDown.TabIndex Then
        fnTableClick objCOntrol
    Else
        Dim lDropDown As Integer
        Dim nctrlType As Integer
        fnGetKeys lDropDown, nctrlType, objCOntrol
        
        If nctrlType = CONTROL_TYPE_TEXTBOX Then
            Exit Sub
        End If
    
        fnLoadTable lDropDown
    End If

End Sub


Private Sub fnGetText()
    Dim i As Integer
    Dim lDropDown As Long
    Dim lColumn As Long
    Dim colColumn As Collection
    Dim sTemp As String

    With colComboControls.Item(nCurrentDD)
        If Not .TextBox Is Nothing Then
            If .TextSource = "" Then
                sTemp = Trim$(tgDropDown.Columns(0).Value)
            Else
                Dim nA() As Integer
                fnGetParsingColumns nA, .TextSource
                sTemp = ""
                For i = 0 To UBound(nA)
                    sTemp = sTemp & Trim$(tgDropDown.Columns(nA(i)).Value) & "  "
                Next
            End If
            On Error GoTo errGetText
            .TextBox.Text = Trim$(sTemp)
            On Error GoTo 0
        End If
    End With
    tgDropDown.Visible = False
    frmParent.Cls
    bValidSelection = True
    tgDropDown.ZOrder 1
    
    If IsNull(tgDropDown.FirstRow) Then
        m_nSelectedRow = tgDropDown.Row
    Else
        m_nSelectedRow = tgDropDown.FirstRow + tgDropDown.Row
    End If
    Exit Sub

errGetText:
    colComboControls.Item(nCurrentDD).TextBox.Text = ""
    Err.Clear
    Resume Next
End Sub
Public Sub SetData(A() As Variant, Optional vDD As Variant)
    Dim nDD As Integer

    If IsMissing(vDD) Then
        nDD = nCurrentDD
    Else
        nDD = vDD
    End If
    
    colComboControls.Item(nDD).SetArray A
    
End Sub

Public Sub GotFocus(objCOntrol As Object)
    
    If objCOntrol Is Nothing Then
        Exit Sub
    End If
    
    If objCOntrol.TabIndex = tgDropDown.TabIndex Then
        Exit Sub
    End If

    Dim lDropDown As Integer
    Dim nctrlType As Integer
    Dim clsTemp  As clsDDSingle

    fnGetKeys lDropDown, nctrlType, objCOntrol
    
    If lDropDown < 0 Then
        Exit Sub
    End If
    
    On Error Resume Next
    
    If nctrlType = CONTROL_TYPE_TEXTBOX Then
        
        Set clsTemp = colComboControls.Item(lDropDown)
        
        tfnRegExpControlChange objCOntrol, clsTemp.RegExpPattern
        
        clsTemp.TextBox.SelStart = 0

        clsTemp.TextBox.SelLength = Len(clsTemp.TextBox.Text)

    End If

    On Error GoTo 0

End Sub


Private Function fnGetPCode(ByVal nType As Integer)
    
    'get the PCode for the regular expression
    Dim nCode As Integer
    
    Dim sBuffer As String * 4096
    sBuffer = Space(4096)

    Select Case nType
        Case TYPE_STRING
            nCode = GetRegExpPCode(sSTRING_PATTERN, sBuffer)
        Case TYPE_LONG
            nCode = GetRegExpPCode(sLONG_PATTERN, sBuffer)
        Case TYPE_DATE
            nCode = GetRegExpPCode(szDatePattern, sBuffer)
        Case Else
            MsgBox "Invalid data type"
    End Select

    fnGetPCode = Left(sBuffer, nCode) & Chr(0)

End Function


Public Sub LostFocus(objCOntrol As Object)
    Dim nTab As Integer
    nTab = frmParent.ActiveControl.TabIndex

    If objCOntrol.TabIndex = tgDropDown.TabIndex Then
        tgDropDown.Visible = False
        frmParent.Cls
        If nCurrentDD > 0 Then
            fnSetFocus colComboControls.Item(nCurrentDD).TextBox
        End If
    End If
End Sub
Public Sub MouseUp(Y As Single)
    nClickPointY = Y
End Sub
Public Sub TableMouseUp(Y As Single)
    nClickPointY = Y
End Sub

Property Set Form(frmMain As Object)
    Set frmParent = frmMain
End Property
Public Sub RowColChange()
    Dim lDropDown As Long
    Dim lCol As Long
    
    If bTableClicked Then
        If lDropDown >= 0 Then
            fnSetFocus colComboControls.Item(nCurrentDD).TextBox
            fnGetText
        End If
        bTableClicked = False
    End If
End Sub
Private Sub fnTableClick(tgTable As Object)
    If tgTable.TabIndex = tgDropDown.TabIndex Then
        Dim nRowTop As Integer
        Dim nRowBtm As Integer
        If tgDropDown.Row >= 0 Then
            nRowTop = tgDropDown.RowTop(tgDropDown.Row)
            nRowBtm = nRowTop + tgDropDown.RowHeight
            If nClickPointY >= nRowTop And nClickPointY <= nRowBtm Then
                fnSetFocus colComboControls.Item(nCurrentDD).TextBox
                fnGetText
            Else
                bTableClicked = True
            End If
        Else
            bTableClicked = True
        End If
    End If
End Sub

Public Sub TextColumn(txtBox As Object, sSrc As String)
    Dim nDD As Integer
    Dim nType As Integer
    
    fnGetKeys nDD, nType, txtBox
    
    If nDD > 0 Then
        colComboControls.Item(nDD).TextSource = sSrc
    End If
End Sub

Property Get ValidSelection() As Boolean
    ValidSelection = bValidSelection
End Property
Property Get SingleRecordSelected() As Boolean
    SingleRecordSelected = bValidSelection
End Property
Property Let SingleRecordSelected(bFlag As Boolean)
    bValidSelection = bFlag
End Property

Property Let ValidSelection(bStatus As Boolean)
    bValidSelection = bStatus
End Property

Property Let VisibleRows(nRows As Integer)
    
    colComboControls.Item(nCurrentDD).VisibleRows = nRows

End Property

Property Get VisibleRows() As Integer
    
    VisibleRows = colComboControls.Item(nCurrentDD).VisibleRows

End Property

Property Get DropDown() As Integer
    DropDown = nDROP_DOWN
End Property

Private Sub Class_Initialize()
    Set colComboControls = New Collection
    nCurrentDD = 1
    nMaxRows = 0
End Sub

Private Sub Class_Terminate()
    Dim colTemp As clsDDSingle
    While colComboControls.count > 0
        Set colComboControls.Item(1).TextBox = Nothing
        Set colComboControls.Item(1).Button = Nothing
        colComboControls.Remove 1
    Wend
    Set colComboControls = Nothing
End Sub


