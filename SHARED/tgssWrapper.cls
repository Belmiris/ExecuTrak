VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tgssWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : tgssWrapper
' DateTime  : 5/17/2005 09:04
' Author    : DenBorg
' Purpose   : This class is a wrapper for the class clsTGSpreadSheet (tgEditor.cls)
'             When an instance of this tgssWrapper class is created, it automatically
'             creates an instance of clsTGSpreadSheet for you, and it automatically
'             forwards the Table events to the Spread Sheet class for you. This way,
'             you don't have to worry about forgetting to code the event procedures
'             of the table grid (True DB Grid). You still have access to the actual
'             spread sheet class instance via the tgSpreadSheet property.
'
'             Usage Example:
'
'                Dim tgwTable1 As tgssWrapper
'                Dim SQL       As String
'
'                SQL = SQLParm(SQL_SELECT_CUSTOMER, "@CustID", 1234)
'                Set tgwTable1 = New tgssWrapper
'                With tgwTable1
'                   Set .Form = Me
'                   Set .StatusBar = efraStatusBar
'                   Set .Table = Table1
'                   .tgSpreadSheet.FillWithSQL t_dbMainDatabase, SQL
'                End With
'
'
'                The following Table1 Events are automatically forwarded:
'
'                   - AfterColEdit(ColIndex)
'                   - BeforeColEdit(ColIndex, KeyAscii, Cancel)
'                   - Change
'                   - GotFocus
'                   - KeyDown(KeyCode, Shift)
'                   - KeyPress(KeyAscii)
'                   - LostFocus
'                   - RowColChange(LastRow, LastCol)
'                   - SelChange
'                   - UnboundReadData(RowBuf, StartLocation, ReadPriorRows)
'
'---------------------------------------------------------------------------------------

Option Explicit

Private m_tgSpreadSheet   As clsTGSpreadSheet
Private m_tgSelector As clsTGSelector
Private m_Form            As Form
Private m_StatusBar       As FactorFrame
Private m_HasSelector As Boolean
Private WithEvents m_Grid As TDBGrid
Attribute m_Grid.VB_VarHelpID = -1

Public Property Get Form() As Form
    Set Form = m_Form
End Property

Public Property Set Form(pForm As Form)
    Set m_Form = pForm
    Set m_tgSpreadSheet.Form = pForm
End Property

Public Property Get Table() As TDBGrid
    Set Table = m_Grid
End Property

Public Property Set Table(pTable As TDBGrid)
    Set m_Grid = pTable
    Set m_tgSpreadSheet.Table = pTable
End Property

Public Function CreateSelector()
    Set m_tgSelector = New clsTGSelector
    Set m_tgSelector.EditorClass = m_tgSpreadSheet
    m_HasSelector = True
End Function

Public Property Get tgSelector() As clsTGSelector
    Set tgSelector = m_tgSelector
End Property

Public Property Get HasSelector() As Boolean
    HasSelector = m_HasSelector
End Property

Private Property Let HasSelector(bHasSelector As Boolean)
    m_HasSelector = bHasSelector
End Property

Private Sub Class_Initialize()
    Set m_tgSpreadSheet = New clsTGSpreadSheet
End Sub

Private Sub Class_Terminate()
    Set m_tgSpreadSheet = Nothing
    Set m_Form = Nothing
    Set m_Grid = Nothing
    Set m_StatusBar = Nothing
End Sub

Private Sub m_Grid_AfterColEdit(ByVal ColIndex As Integer)
    m_tgSpreadSheet.AfterColEdit ColIndex
End Sub

Private Sub m_Grid_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, CANCEL As Integer)
    m_tgSpreadSheet.BeforeColEdit ColIndex, KeyAscii, CANCEL
End Sub

Private Sub m_Grid_Change()
    m_tgSpreadSheet.Change
End Sub

Private Sub m_grid_Click()
    If HasSelector Then
        m_tgSelector.Click
    End If
End Sub

Private Sub m_grid_GotFocus()
    m_tgSpreadSheet.GotFocus
    
    If HasSelector Then
        m_tgSelector.GotFocus
    End If
End Sub

Private Sub m_Grid_HeadClick(ByVal ColIndex As Integer)
    m_tgSpreadSheet.HeadClick ColIndex
End Sub

Private Sub m_Grid_KeyDown(KeyCode As Integer, Shift As Integer)
    m_tgSpreadSheet.KeyDown KeyCode, Shift
    
    If HasSelector Then
        m_tgSelector.KeyDown KeyCode, Shift
    End If
End Sub

Private Sub m_grid_KeyPress(KeyAscii As Integer)
    If Not m_tgSpreadSheet.Keypress(KeyAscii) Then
        KeyAscii = 0
    End If
End Sub

Private Sub m_grid_LostFocus()
    m_tgSpreadSheet.LostFocus
    
    If HasSelector Then
        m_tgSelector.LostFocus
    End If
End Sub

Private Sub m_Grid_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If HasSelector Then
        m_tgSelector.MouseDown Button, Shift, x, y
    End If
End Sub

Private Sub m_grid_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If HasSelector Then
        m_tgSelector.MouseUp Button, Shift, x, y
    End If
End Sub

Private Sub m_grid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    m_tgSpreadSheet.RowColChange LastRow, LastCol
    
    If HasSelector Then
        m_tgSelector.RowColChange LastRow, LastCol
    End If
End Sub

Private Sub m_grid_SelChange(CANCEL As Integer)
    CANCEL = True
    
    If HasSelector Then
        m_tgSelector.SelChange CANCEL
    End If
End Sub

Private Sub m_Grid_UnboundReadData(ByVal RowBuf As DBTRUEGRID.RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
    m_tgSpreadSheet.ReadData RowBuf, StartLocation, ReadPriorRows
End Sub

Public Property Get StatusBar() As FactorFrame
    Set StatusBar = m_StatusBar
End Property

Public Property Set StatusBar(pStatusBar As FactorFrame)
    Set m_StatusBar = pStatusBar
    Set m_tgSpreadSheet.StatusBar = pStatusBar
End Property

Public Property Get tgSpreadSheet() As clsTGSpreadSheet
    Set tgSpreadSheet = m_tgSpreadSheet
End Property
