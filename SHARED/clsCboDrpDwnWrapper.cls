VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCboDrpDwnWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private WithEvents m_grid As TDBGrid
Attribute m_grid.VB_VarHelpID = -1
Private m_tgcComboBox As Object

Private Sub Class_Initialize()
    Set m_tgcComboBox = CreateObject(t_szOLECOMBO)
End Sub

Private Sub Class_Terminate()
    Set m_tgcComboBox = Nothing
    Set m_grid = Nothing
End Sub

Public Property Get ComboDropDown() As Object
    Set ComboDropDown = m_tgcComboBox
End Property

Public Property Set Table(pTable As TDBGrid)
    Set m_grid = pTable
    Set m_tgcComboBox.Table = pTable
End Property

Public Property Set Form(pForm As Form)
    Set m_tgcComboBox.Form = pForm
End Property

Public Property Set DBEngine(pDBEngine As DBEngine)
    Set m_tgcComboBox.DBEngine = pDBEngine
End Property

Public Property Set DataBase(pDataBase As DataBase)
    Set m_tgcComboBox.DataBase = pDataBase
End Property

Public Property Set DataLink(pDataLink As Data)
    Set m_tgcComboBox.DataLink = pDataLink
End Property


Private Sub m_grid_Click()
    m_tgcComboBox.Click m_grid
End Sub

Private Sub m_grid_GotFocus()
    m_tgcComboBox.GotFocus m_grid
End Sub

Private Sub m_grid_KeyPress(KeyAscii As Integer)
    m_tgcComboBox.Keypress m_grid, KeyAscii
End Sub

Private Sub m_grid_LostFocus()
    m_tgcComboBox.LostFocus m_grid
End Sub

Private Sub m_grid_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    m_tgcComboBox.TableMouseUp y
End Sub

Private Sub m_grid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    m_tgcComboBox.RowColChange
End Sub

Private Sub m_grid_SelChange(CANCEL As Integer)
    m_tgcComboBox.SelChange CANCEL
End Sub
