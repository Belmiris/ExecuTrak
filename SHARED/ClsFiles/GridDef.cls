VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GridDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' See program FMFFRGT as an example that exploits all grid features,
' including multi-CELL selection, conditional cell coloring, etc.
'
' Other classes that might be incorporated into this class:
'   ColorCond.cls
'   CellColor.cls
'

Private Const SCROLLBAR_WIDTH = 252
Private Const TIMER_CTL_NAME = "tmrKeyboard"

'==================================
'Enums
'==================================
Public Enum GridDefSortOptions
    NoSort = 0           'TGM.SortByColumn = False: TGM.ReverseSortByColumn = False
    AllowSort            'TGM.SortByColumn = True : TGM.ReverseSortByColumn = False
    AllowSortWithReverse 'TGM.SortByColumn = True : TGM.ReverseSortByColumn = True
End Enum

'==================================
'Events Raised
'==================================
Event FirstRowChange()
Event SelectChange(Grid As GridDef)
Event TableClick(CANCEL As Boolean)
Event TableGotFocus(Table As TDBGrid)
Event CellChange(Col As Integer, row As Integer, oldValue As String, newValue As String)

'==================================
'Member Variables for Properties
'==================================
'--- PUBLIC  ---
Public Name                   As String

'--- PRIVATE ---
'Those With Events
Private WithEvents m_Table    As TDBGrid
Attribute m_Table.VB_VarHelpID = -1
Private WithEvents m_Timer    As Timer
Attribute m_Timer.VB_VarHelpID = -1
Private WithEvents m_DDTable  As TDBGrid
Attribute m_DDTable.VB_VarHelpID = -1
Private WithEvents m_DDButton As FactorFrame
Attribute m_DDButton.VB_VarHelpID = -1

'Those WithOut Events:
Private m_GridColumns         As Collection
Private m_Form                As Form
Private m_StatusBar           As Object
Private m_TGM                 As clsTGSpreadSheet
Private m_FTX                 As clsColumnExtension
Private m_TGS                 As clsTGSelector
Private m_FDD                 As clsFloatingDropDown
Private m_FTX_PicBox          As Object
Private m_Locked              As Boolean
Private m_AllowMultiSel       As Boolean            'Keep track ourselves, since TGS.AllowMultipleSelect is ReadOnly
Private m_HideSelection       As Boolean            'Hide Selection when Lose Focus
Private m_SortOption          As GridDefSortOptions

Public Function AddColumn(ByVal Key As String, _
                          ByVal Header As String, _
                          ByVal WidthPct As Single, _
                          Optional ByVal FieldName As String = vbNullString, _
                          Optional ByVal Editable As Boolean = True, _
                          Optional ByVal RegExp As String = vbNullString) As GridColDef
    Dim GCDef As GridColDef
    
    Set GCDef = New GridColDef
    With GCDef
        .Key = Key
        .Header = Header
        .WidthPct = WidthPct
        .FieldName = FieldName
        .Editable = Editable
        .RegExp = RegExp
        .Index = m_GridColumns.Count
    End With
    
    m_GridColumns.Add GCDef, Key
    
    Set AddColumn = GCDef
End Function
Public Property Let AllowMultipleSelect(ByVal Allow As Boolean)
    m_AllowMultiSel = Allow
    If Not (m_TGS Is Nothing) Then
        m_TGS.AllowMultipleSelect = Allow
    End If
End Property
Public Property Get AllowMultipleSelect() As Boolean
    AllowMultipleSelect = m_AllowMultiSel
End Property
Public Property Get Column(vKeyIndex As Variant) As GridColDef
Attribute Column.VB_UserMemId = 0
    Set Column = m_GridColumns(vKeyIndex)
End Property
Public Property Get ColumnCount() As Long
    ColumnCount = m_GridColumns.Count
End Property
Public Function ColumnTotal(ByVal Col As Integer) As Double
    Dim Total As Double
    Dim row   As Long
    
    With m_TGM
        For row = 0 To .RowCount - 1
            Total = Total + .CellValue(Col, row)
        Next 'Row
    End With
    
    ColumnTotal = Total
End Function
Public Property Set Form(ByVal pForm As Form)
    Set m_Form = pForm
End Property
Public Property Let HideSelection(ByVal HideSel As Boolean)
    m_HideSelection = HideSel
End Property
Public Property Get HideSelection() As Boolean
    HideSelection = m_HideSelection
End Property
Public Function InitFloatingDropDown(SearchButton As FactorFrame, DropDownTable As TDBGrid, DBase As DAO.DataBase, DataCtl As Data) As clsFloatingDropDown
    Set m_FDD = New clsFloatingDropDown
    With m_FDD
        Set .SearchButton = SearchButton
        Set m_DDButton = SearchButton
        Set .MainTable = m_Table
        Set .DropDownTable = DropDownTable
        Set m_DDTable = DropDownTable
        Set .DataBase = DBase
        Set .DataLink = DataCtl
        Set .Form = m_Form
        Set .EditClass = m_TGM
        .DefaultCursorControl = True
        .SearchOnReturn = False
    End With
    
    Set InitFloatingDropDown = m_FDD
End Function
Public Sub Initialize(pForm As Form, pTable As TDBGrid, pStatusBar As Object)
    Set m_Form = pForm
    Set m_Table = pTable
    Set m_StatusBar = pStatusBar
    
    Set m_FTX_PicBox = pForm.Controls.Add("VB.PictureBox", "pb_" & pTable.Name)
    With m_FTX_PicBox
        .Appearance = 0 'Flat (not 3-D)
        Set .Container = pTable.Container
    End With
End Sub
Public Function InitSelector() As clsTGSelector
    Set m_TGS = New clsTGSelector
    Set InitSelector = m_TGS
End Function
Public Property Let Locked(ByVal pLocked As Boolean)
    m_TGM.Locked = pLocked
End Property
Public Property Get Locked() As Boolean
    Locked = m_TGM.Locked
End Property
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = m_GridColumns.[_NewEnum]
End Property
Public Property Get ColumnExtension() As clsColumnExtension
    Set ColumnExtension = m_FTX
End Property
Public Sub RemoveColumn(ByVal vKeyIndex As Variant)
    m_GridColumns.Remove vKeyIndex
End Sub
Public Property Get Selector() As clsTGSelector
    Set Selector = m_TGS
End Property
Public Sub SetupTable()
    Const METHOD = ".SetupTable"
    Dim TableWidth As Single
    Dim ColDef     As GridColDef
    Dim ColIndex   As Long
    Dim HasEditCol As Boolean    'True if there is one or more editable columns
    
    On Error GoTo ErrHandler
    
    With m_Table
        TableWidth = .width - SCROLLBAR_WIDTH
        With .Columns
            Do While .Count
                .Remove 0
            Loop
            
            For ColIndex = 0 To ColumnCount - 1
                Set ColDef = m_GridColumns(ColIndex + 1)
                If ColDef.Editable Then
                    HasEditCol = True
                End If
                If Not ColDef.HiddenField Then
                    With .Add(ColIndex)
                        .width = ColDef.WidthPct * TableWidth
                        .DataField = ColDef.FieldName
                        .Caption = ColDef.Header
                        .Visible = ColDef.Visible
                        .Alignment = ColDef.Alignment
                    End With
                End If
            Next 'ColIndex
            Set ColDef = Nothing
        End With
    End With
    
    With m_TGM
        Set .Table = m_Table
        Set .StatusBar = m_StatusBar
        Set .Form = m_Form
        Set .engFactor = t_engFactor
        If Not HasEditCol Then
            ' Only call this method if grid does
            ' *NOT* have any editable columns
            .SetupTable True
        End If
        '.CtrlYClearRow = True
    End With

    With m_FTX
        Set .Table = m_Table
        Set .StatusBar = m_StatusBar
        Set .Form = m_Form
        If Not (m_FTX_PicBox Is Nothing) Then
            Set .PictureBox = m_FTX_PicBox
        End If
        .PositionDefault = .PositionLeft
        .StyleDefault = .ShowCellText
        .UseEditorRow = True
    End With
    Set m_Timer = m_Form(TIMER_CTL_NAME)
    
    For ColIndex = 0 To ColumnCount - 1
        Set ColDef = m_GridColumns(ColIndex + 1)
        If Not ColDef.HiddenField Then
            m_FTX.AddColumn ColIndex, , m_FTX.PositionTop
        End If
        If ColDef.Editable Then
            m_TGM.AddEditColumn ColIndex, ColDef.Prompt, ColDef.RegExp
            m_Table.Columns(ColIndex).Locked = ColDef.Locked
        ElseIf ColDef.HiddenField Then
            m_TGM.AddHiddenField ColDef.FieldName
        End If
        '<NUMERIC_TYPE|STRING_TYPE>, <ASCENDING|DESCEDING>, <CASE_SENSITIVE|CASE_INSENSITIVE>
'        m_TGM.AddSortColumn CInt(ColIndex), CInt(ColIndex), m_TGM.STRING_TYPE, m_TGM.DESCENDING, m_TGM.CASE_INSENSITIVE
        m_TGM.DisplayFormat(ColIndex) = ColDef.DisplayFormat
    Next 'ColIndex
    Set ColDef = Nothing
    
    If Not (m_TGS Is Nothing) Then
        With m_TGS
            Set .EditorClass = m_TGM
            .AvoidBeep = False
            .AllowMultipleSelect = m_AllowMultiSel
            .SelectCurrRow = True
            .RowHighLighted = True
        End With
    End If

    m_TGM.AllowAddNew = m_Table.AllowAddNew
    Exit Sub
    
ErrHandler:
    With Err
        Select Case .number
            Case 730 'Control Not Found
                With m_Form.Controls.Add("VB.Timer", TIMER_CTL_NAME)
                    .Tag = "Dynamically Created"
                    .Interval = 250
                    .Enabled = True
                End With
                Resume
            Case Else
                ErrorTrapMessage Me.Name & METHOD, Err
        End Select
    End With
    Exit Sub
End Sub
Public Sub ErrorTrapMessage(p_function As String, _
                            p_error As ErrObject, _
                   Optional p_Style As VbMsgBoxStyle = vbCritical, _
                   Optional p_SQL As String = "")
                    
    Dim strReason As String
    
    strReason = App.EXEName & " has detected an error in function " & p_function & vbCrLf & vbCrLf _
              & "Error Code: " & p_error.number & vbCrLf _
              & "Error Desc: " & p_error.Description & vbCrLf & vbCrLf
    
    If LenB(p_SQL) = 0 Then
        MsgBox strReason, p_Style
    Else
        tfnErrHandler p_function, p_SQL, True
    End If
    
End Sub
Public Sub ClearData()
    If Not (m_TGS Is Nothing) Then
        m_TGS.Clear
    End If
    m_TGM.ClearData
End Sub
Public Property Let SortOption(ByVal Sort As GridDefSortOptions)
    m_SortOption = Sort
    With m_TGM
        .SortByColumn = (Sort > NoSort)
        .ReverseSortByColumn = (Sort = AllowSortWithReverse)
    End With
End Property
Public Property Get SortOption() As GridDefSortOptions
    SortOption = m_SortOption
End Property
Public Property Get SpreadSheet() As clsTGSpreadSheet
    Set SpreadSheet = m_TGM
End Property
Public Property Set StatusBar(ByVal StatBar As Object)
    Set m_StatusBar = StatBar
End Property
Public Property Set Table(pTable As TDBGrid)
    Set m_Table = pTable
End Property
Public Property Get Table() As TDBGrid
    Set Table = m_Table
End Property
Private Sub Class_Initialize()
    Name = "GridDef"
    Set m_GridColumns = New Collection
    Set m_TGM = New clsTGSpreadSheet
    Set m_FTX = New clsColumnExtension
'    Set m_TGS = New clsTGSelector
End Sub
Private Sub Class_Terminate()
    Set m_GridColumns = Nothing
    Set m_Table = Nothing
    Set m_TGM = Nothing
    Set m_FTX = Nothing
    Set m_FDD = Nothing
    Set m_DDTable = Nothing
    Set m_DDButton = Nothing
End Sub
Private Sub m_DDButton_Click()
    m_FDD.ButtonClick
    
    If m_FDD.ValidSelection Then
        m_Table_AfterColEdit m_Table.Col
        m_FDD.ValidSelection = True
    End If
End Sub
Private Sub m_DDButton_GotFocus()
    m_FDD.GotFocus m_DDButton
End Sub

Private Sub m_DDButton_LostFocus()
    m_FDD.LostFocus m_DDButton
End Sub

Private Sub m_DDTable_Click()
    m_FDD.TableClick m_DDTable
End Sub

Private Sub m_DDTable_KeyDown(KeyCode As Integer, Shift As Integer)
    m_FDD.KeyDown m_DDTable, KeyCode
End Sub
Private Sub m_DDTable_KeyPress(KeyAscii As Integer)
    m_FDD.Keypress m_DDTable, KeyAscii
End Sub
Private Sub m_DDTable_LostFocus()
    m_FDD.LostFocus m_DDTable
End Sub
Private Sub m_DDTable_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    m_FDD.MouseUp y
End Sub
Private Sub m_DDTable_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    m_FDD.RowColChange m_DDTable
End Sub
Private Sub m_Table_AfterColEdit(ByVal ColIndex As Integer)
    Dim row As Long
    Dim oldValue As String, newValue As String
    
    row = m_TGM.GetCurrentRowNumber
    
    oldValue = m_TGM.CellValue(ColIndex, row)
    
    m_TGM.AfterColEdit ColIndex
    
    newValue = m_TGM.CellValue(ColIndex, row)
        
    RaiseEvent CellChange(ColIndex, m_TGM.GetCurrentRowNumber, oldValue, newValue)
    
End Sub
Private Sub m_Table_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, CANCEL As Integer)
    m_TGM.BeforeColEdit ColIndex, KeyAscii, CANCEL
End Sub
Private Sub m_Table_Change()
    m_TGM.Change
End Sub
Private Sub m_Table_Click()
    Dim CANCEL As Boolean
    
    RaiseEvent TableClick(CANCEL)
    If Not CANCEL Then
        If Not (m_TGS Is Nothing) Then
            m_TGS.Click
        End If
        If Not (m_FDD Is Nothing) Then
            m_FDD.TableClick m_Table
        End If
        RaiseEvent SelectChange(Me)
    End If
End Sub

Private Sub m_Table_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As Object)
    If Not (m_TGS Is Nothing) Then
        m_TGS.FetchCellStyle Condition, Split, Bookmark, Col, CellStyle
    End If
End Sub
Private Sub m_Table_FirstRowChange()
    m_TGM.FirstRowChange
    If Not (m_FDD Is Nothing) Then
        m_FDD.FirstRowChange
    End If
    RaiseEvent FirstRowChange
End Sub
Private Sub m_Table_GotFocus()
    If Not (m_TGS Is Nothing) Then
        With m_TGS
            .GotFocus
            .SelectRow m_TGM.GetCurrentRowNumber
        End With
    End If
    m_TGM.GotFocus
    If Not (m_FDD Is Nothing) Then
        m_FDD.GotFocus
'        If m_FDD.ValidSelection Then
'            '
'        End If
    End If
    DoEvents
    RaiseEvent TableGotFocus(m_Table)
End Sub
Private Sub m_Table_HeadClick(ByVal ColIndex As Integer)
    m_TGM.HeadClick ColIndex
    If Not (m_TGS Is Nothing) Then
        m_TGS.HeadClick ColIndex
    End If
End Sub
Private Sub m_Table_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim SelChange As Boolean
    
    SelChange = ((Shift = vbCtrlMask) And (KeyCode = vbKeyA)) _
             Or (KeyCode = vbKeyDown) Or (KeyCode = vbKeyUp)
    If Not (m_TGS Is Nothing) Then
        m_TGS.KeyDown KeyCode, Shift
    End If
    m_TGM.KeyDown KeyCode, Shift
    If Not (m_FDD Is Nothing) Then
        If KeyCode = vbKeyReturn Then
            m_FDD.Keypress m_Table, KeyCode
'            If m_FDD.ValidSelection Then
'                '
'            End If
        End If
    End If
    If SelChange Then
        RaiseEvent SelectChange(Me)
    End If
End Sub
Private Sub m_Table_KeyPress(KeyAscii As Integer)
    If Not m_TGM.Keypress(KeyAscii) Then
        KeyAscii = 0
    End If
End Sub
Private Sub m_Table_LostFocus()
    If Not (m_TGS Is Nothing) Then
        m_TGS.LostFocus m_HideSelection
    End If
    m_TGM.LostFocus
    If Not (m_FDD Is Nothing) Then
        m_FDD.LostFocus m_Table
    End If
End Sub
Private Sub m_Table_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not (m_TGS Is Nothing) Then
        m_TGS.MouseDown Button, Shift, x, y
    End If
    m_TGM.MouseDown Button, Shift, x, y
    'm_FTX.MouseDown Button, X, Y
End Sub
Private Sub m_Table_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'    m_FTX.MouseMove X, Y
End Sub
Private Sub m_Table_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not (m_TGS Is Nothing) Then
        m_TGS.MouseUp Button, Shift, x, y
    End If
'    m_FTX.MouseUp
End Sub

Private Sub m_Table_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim RetVal As Boolean

    If Not (m_FDD Is Nothing) Then
        RetVal = m_FDD.RowColChange(m_Table)
    Else
        RetVal = True
    End If
    
    If RetVal Then
        If Not (m_TGS Is Nothing) Then
            m_TGS.RowColChange LastRow, LastCol
        End If
        m_TGM.RowColChange LastRow, LastCol
    End If
'    m_FTX.RowColChange
    RaiseEvent SelectChange(Me)
End Sub

Private Sub m_Table_Scroll(CANCEL As Integer)
    If Not (m_FDD Is Nothing) Then
        m_FDD.Scroll
    End If
End Sub
Private Sub m_Table_SelChange(CANCEL As Integer)
    m_TGM.SelChange CANCEL
    If Not (m_TGS Is Nothing) Then
        m_TGS.SelChange CANCEL
    End If
'    Cancel = True
End Sub
Private Sub m_Table_UnboundReadData(ByVal RowBuf As DBTrueGrid.RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
    m_TGM.ReadData RowBuf, StartLocation, ReadPriorRows
End Sub

Private Sub m_Timer_Timer()
    Dim ShowExt As Boolean

    ShowExt = m_Table.Visible And (Not (m_FTX Is Nothing)) And (Not (m_FTX_PicBox Is Nothing))
    If ShowExt Then
        m_FTX.ShowColumnExt
    Else
        m_FTX_PicBox.Visible = False
    End If
End Sub

Public Property Get FloatingDropDown() As clsFloatingDropDown
    Set FloatingDropDown = m_FDD
End Property

Public Function CurrentRowIsEmpty() As Boolean
    Dim i As Integer
    Dim row As Integer
    Dim isEmpty As Boolean
    
    isEmpty = True
    
    With Me
        row = .SpreadSheet.GetCurrentRowNumber
        
        For i = 0 To .ColumnCount
            If .SpreadSheet.CellValue(i, row) <> vbNullString Then
                isEmpty = False
                Exit For
            End If
        Next
    End With
    
    CurrentRowIsEmpty = isEmpty

End Function

Public Function ValueExists(Value As String, Col As Integer, Optional ignoreRow As Integer = -1) As Boolean
    Dim i As Integer
    
    With Me
        
        If ignoreRow = -1 Then
            ignoreRow = .SpreadSheet.GetCurrentRowNumber
        End If
        
        For i = 0 To .SpreadSheet.RowCount - 1
            If i <> ignoreRow Then
                If .SpreadSheet.CellValue(Col, i) = Value Then
                    ValueExists = True
                    Exit Function
                End If
            End If
        Next
    
    End With

End Function
