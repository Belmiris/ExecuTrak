VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPRFHOURS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Form As Form
Private tbToolbar As Toolbar
Private ffraStatusbar As FactorFrame
Private txtEmployeeNumber As Textbox
Private cmdEmployeeNumber As FactorFrame
Private txtEmployeeName As Textbox
Private cmdEmployeeName As FactorFrame
Private txtSSN As Textbox
Private cmdSSN As FactorFrame
Private datComboDropDown As Data
Private datFloating As Data
Private tblComboDropdown As TDBGrid
Private tblTimeCard As TDBGrid
Private tblProfitCenter As TDBGrid
Private tblFloating As TDBGrid
Private cmdFloatingBtn As PictureBox
Private txtTotalDollars As Textbox
Private txtTotal As Textbox
Private cmdAddBtn  As FactorFrame
Private cmdEditBtn  As FactorFrame
Private cmdDeleteBtn  As FactorFrame
Private cmdUpdateInsertBtn As FactorFrame
Private cmdRefreshSelectBtn As FactorFrame
Private cmdExitCancelBtn As FactorFrame
Private cmdExitBtn As FactorFrame
Private efraBridge As FactorFrame

Private mnuCancel As Menu

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private t_bStartupFlag As Boolean 'optional startup flag
Private t_bDataChanged As Boolean 'data changed flag
Private t_bUpdateTable As Boolean 'update data flag

Private t_nFormMode As Integer         'global used to track the current form operating mode

Private Const IDLE_MODE As Integer = 0 'idle mode activates the NoDrop Cursor
Private Const ADD_MODE As Integer = 1  'flag set when in the add mode
Private Const EDIT_MODE As Integer = 2 'flag set when in the edit mode
Private Const DELETE_MODE As Integer = 3 'flag set when in the delete mode

'========================
'Standard Button Captions
'========================
Private Const t_szCAPTION_INSERT As String = "&Insert"
Private Const t_szCAPTION_UPDATE As String = "&Update"
Private Const t_szCAPTION_REFRESH As String = "&Refresh"
Private Const t_szCAPTION_SELECT As String = "&Select"
Private Const t_szCAPTION_CANCEL As String = "&Cancel"
Private Const t_szCAPTION_EXIT As String = "E&xit"

'==========================
'Status Bar Default Strings
'==========================
Private Const t_szCLEAR As String = ""
Private Const t_szADD As String = "Select Add, Edit or Exit."
Private Const t_szEDIT As String = "Select Add, Edit or Exit."
Private Const t_szDELETE As String = "Delete."
Private Const t_szINSERT As String = "Insert."
Private Const t_szUPDATE As String = "Update."
Private Const t_szREFRESH As String = "Refresh."
Private Const t_szSELECT As String = "Select Add, Edit or Exit."
Private Const t_szEXIT As String = "Exit"
Private Const t_szCANCEL As String = "Cancel"
Private Const t_szADDEDIT As String = "Select Add, Edit or Exit."
Private Const t_szPRINT As String = "Print"
Private Const t_szCOPYFROM As String = "Copy From"
Private Const t_szRUN1 As String = "Add Profit Center."
Private Const t_szRUN2 As String = "Add Employee."
Private Const t_szRUN3 As String = "Run #3"
Private Const t_szHELP As String = "Help"
Private s_EmployeeNumber As String
Private s_EmployeeName As String
Private s_SNumber As String
Private cValidate As cValidateInput
Private tgcEditor As clsTGSpreadSheet
Private tgcEditorPrft As clsTGSpreadSheet
Private tgcFEditor As clsFloatingDropDown
Private b_TableFilled As Boolean
Private nsArray() As Variant
Private nsArrayCenter() As Variant
Private bNeverDelet As Boolean
Private sDecimalString As String

Private sTempTable As String
Private sTempTable1 As String
Private nDataStatus As Integer
Private Const DATA_INI = 0
Private Const DATA_LOADED = 1
Private Const DATA_CHANGED = 2
Private Const COL_DATE = 0
Private Const COL_PRFT = 1
Private Const COL_CODE = 2
Private Const COL_TYPE = 3
Private Const COL_HOUR = 4
Private COL_SOURCE As Integer
Private bWarned As Boolean 'for warning of Date with glp_status ='W'

Private bTableExists As Boolean
'

'===========================
'public properties/functions
'===========================
Property Set MainForm(frmMain As Form)
    Set m_Form = frmMain
End Property
Property Set FormToolBar(tb As Toolbar)
    Set tbToolbar = tb
End Property
Property Set StatusBar(efra As FactorFrame)
    Set ffraStatusbar = efra
End Property
Property Set EmpNumTexBox(txtEmpNum As Textbox)
    Set txtEmployeeNumber = txtEmpNum
End Property
Property Set EmpNumButton(cmdEmpNum As FactorFrame)
    Set cmdEmployeeNumber = cmdEmpNum
End Property
Property Set EmpNameTexBox(txtEmpName As Textbox)
    Set txtEmployeeName = txtEmpName
End Property
Property Set EmpNameButton(cmdEmpName As FactorFrame)
    Set cmdEmployeeName = cmdEmpName
End Property
Property Set SSNTexBox(SSNTextbox As Textbox)
    Set txtSSN = SSNTextbox
End Property
Property Set SSNButton(SSNButton As FactorFrame)
    Set cmdSSN = SSNButton
End Property
Property Set ComboDropDownData(datGrid As Data)
    Set datComboDropDown = datGrid
End Property
Property Set FloatingData(datGrid As Data)
    Set datFloating = datGrid
End Property
Property Set ComboDropdownGrid(tblGrid As TDBGrid)
    Set tblComboDropdown = tblGrid
End Property
Property Set TimeCardGrid(tblGrid As TDBGrid)
    Set tblTimeCard = tblGrid
End Property
Property Set ProfitCenterGrid(tblGrid As TDBGrid)
    Set tblProfitCenter = tblGrid
End Property
Property Set FloatingGrid(tblGrid As TDBGrid)
    Set tblFloating = tblGrid
End Property
Property Set FloatingButton(Pic As PictureBox)
    Set cmdFloatingBtn = Pic
End Property
Property Set TotalDollarsTexBox(TotalDollarsTextbox As Textbox)
    Set txtTotalDollars = TotalDollarsTextbox
End Property
Property Set TotalTexBox(TotalTextbox As Textbox)
    Set txtTotal = TotalTextbox
End Property
Property Set AddButton(btn As FactorFrame)
    Set cmdAddBtn = btn
End Property
Property Set EditButton(btn As FactorFrame)
    Set cmdEditBtn = btn
End Property
Property Set DeleteButton(btn As FactorFrame)
    Set cmdDeleteBtn = btn
End Property
Property Set UpdateInsertButton(btn As FactorFrame)
    Set cmdUpdateInsertBtn = btn
End Property
Property Set RefreshSelectButton(btn As FactorFrame)
    Set cmdRefreshSelectBtn = btn
End Property
Property Set CancelButton(btn As FactorFrame)
    Set cmdExitCancelBtn = btn
End Property
Property Set CancelMenuButton(btn As Menu)
    Set mnuCancel = btn
End Property
Property Set ExitButton(btn As FactorFrame)
    Set cmdExitBtn = btn
End Property
Property Set Bridge(efra As FactorFrame)
    Set efraBridge = efra
End Property
'


Public Sub Form_Load()
    Dim bCode As Boolean
    
    '***************************************************
    ' INSERT YOUR FORM LOAD CODE HERE
    ' | | | | | | |
    ' v v v v v v v
     
    'subSetupToolBar
     subCreateTempTable
     subSetupCombos
     subSetValidtion
     subSetGrid
     subSetFloatingDropDown
           
    ' ^ ^ ^ ^ ^ ^ ^
    ' | | | | | | |
    '***************************************************
    
    tfnResetScreen 'set the default screen

End Sub

Public Sub cmdFloatingBtn_Click()
    subResetFloatingSQL2
    tgcFEditor.ButtonClick cmdFloatingBtn
End Sub

Public Sub cmdFloatingBtn_GotFocus()
    tgcFEditor.GotFocus cmdFloatingBtn
End Sub

Public Sub cmdFloatingBtn_LostFocus()
    tgcFEditor.LostFocus cmdFloatingBtn
End Sub

Public Sub efraBridge_GotFocus()
    cValidate.GotFocus efraBridge
End Sub

Public Sub cmdAddBtn_GotFocus()
    tfnSetStatusBarMessage t_szADD
    On Error Resume Next
    DoEvents
    tgcEditor.ClearData
    tgcEditorPrft.ClearData
End Sub

Public Sub cmdEditBtn_GotFocus()
    tfnSetStatusBarMessage t_szEDIT
End Sub

Public Sub cmdDeleteBtn_GotFocus()
    tfnSetStatusBarMessage t_szDELETE
End Sub

Public Sub cmdUpdateInsertBtn_GotFocus()
    If cmdUpdateInsertBtn.Caption = t_szCAPTION_UPDATE Then
        tfnSetStatusBarMessage t_szUPDATE
    Else
        tfnSetStatusBarMessage t_szINSERT
    End If
End Sub

Public Sub cmdRefreshSelectBtn_GotFocus()
    If cmdRefreshSelectBtn.Caption = t_szCAPTION_REFRESH Then
        tfnSetStatusBarMessage t_szREFRESH
    Else
        tfnSetStatusBarMessage t_szSELECT
    End If
End Sub

Public Sub cmdExitCancelBtn_GotFocus()
    If cmdExitCancelBtn.Caption = t_szCAPTION_EXIT Then
        tfnSetStatusBarMessage t_szEXIT
    Else
        tfnSetStatusBarMessage t_szCANCEL
    End If
End Sub

Public Sub cmdAddBtn_Click()
    t_bStartupFlag = False
    cmdAddBtn.Enabled = False
    cmdEditBtn.Enabled = False
    
    cmdUpdateInsertBtn.Caption = t_szCAPTION_INSERT
    cmdUpdateInsertBtn.Enabled = False
    cmdRefreshSelectBtn.Caption = t_szCAPTION_REFRESH
    cmdRefreshSelectBtn.Enabled = False
    
    frmContext.ButtonEnabled(CANCEL_UP) = True
    cmdExitCancelBtn.Enabled = True
    mnuCancel.Enabled = True
    
    cValidate.ResetFlags
    t_nFormMode = ADD_MODE
    subReSetComboSql
    fnEnableControls True
    tgcEditor.AllowAddNew = True
    tgcEditorPrft.AllowAddNew = True
   
    'tgcEditor.AddEditColumn COL_TYPE, "Enter Pay Type", "^P{1,2}$"
    tblTimeCard.Enabled = False
    tblProfitCenter.Enabled = False
    txtEmployeeNumber.SetFocus
    
End Sub

Public Sub cmdEditBtn_Click()
    t_bStartupFlag = False
    cmdAddBtn.Enabled = False
    cmdEditBtn.Enabled = False
    
    cmdUpdateInsertBtn.Caption = t_szCAPTION_UPDATE
    cmdUpdateInsertBtn.Enabled = False
    
    cmdRefreshSelectBtn.Caption = t_szCAPTION_REFRESH
    cmdRefreshSelectBtn.Enabled = False

    frmContext.ButtonEnabled(CANCEL_UP) = True
    cmdExitCancelBtn.Enabled = True
    mnuCancel.Enabled = True
    
    cValidate.ResetFlags
    t_nFormMode = EDIT_MODE
    
    subReSetComboSql
    fnEnableControls True
    tgcEditor.AllowAddNew = False
    tgcEditorPrft.AllowAddNew = False
    tgcEditor.RemoveEditColumn COL_TYPE '0
    tblTimeCard.Enabled = False
    tblProfitCenter.Enabled = False
    txtEmployeeNumber.SetFocus
    
End Sub

Public Sub cmdDeleteBtn_Click()
    Dim bResponse As Integer
    Dim sStyle As String, sMsg As String
    sMsg = "Are you sure you want to delete this record?"
    sStyle = vbYesNo + vbQuestion + vbDefaultButton2
    bResponse = MsgBox(sMsg, sStyle)
    If bResponse = vbYes Then
        If fnDeleteData = True Then
            fnGetTotalDollars
            fnFillProfitTable
            bNeverDelet = False
            tfnSetStatusBarCorrect ("The Record was Deleted Successfully.")
        End If
    End If
End Sub

Public Sub cmdUpdateInsertBtn_Click()
    tfnSetStatusBarMessage t_szEDIT
    If cValidate.FirstInvalidInput > 0 Then
        Exit Sub
    End If
    t_bUpdateTable = fnInsertUpdate
    If t_bUpdateTable Then
        cmdUpdateInsertBtn.Enabled = False
        t_bDataChanged = False
        tfnResetScreen
    End If
End Sub

Public Sub cmdExitCancelBtn_Click()
    subCancel
End Sub

Public Sub cmdRefreshSelectBtn_Click()
    
    Dim bDone As Boolean
    Dim nResponse As Integer
    
    bDone = False
    nResponse = MsgBox(t_szREFRESH_MESSAGE, _
        vbYesNo + vbQuestion + vbDefaultButton2)
                      
    If nResponse = vbYes Then
        bDone = fnRefresh
    Else
        Exit Sub
    End If
    
    If bDone Then
        nDataStatus = DATA_LOADED
        ' code here about re-validaton
        cmdDeleteBtn.Enabled = True
    Else
        MsgBox "Refresh failed"
        subSetFocus cmdExitCancelBtn
    End If

End Sub

Public Sub mnuAdd_Click()
    cmdAddBtn_Click
End Sub

Public Sub mnuEdit_Click()
    cmdEditBtn_Click
End Sub

Public Sub mnudelete_Click()
    cmdDeleteBtn_Click
End Sub

Public Sub mnuCancel_Click()
    subCancel
End Sub


Public Sub tblFloating_Click()
    tgcFEditor.TableClick tblFloating
End Sub

Public Sub tblFloating_KeyPress(KeyAscii As Integer)
     tgcFEditor.Keypress tblFloating, KeyAscii
End Sub

Public Sub tblFloating_LostFocus()
    tgcFEditor.LostFocus tblFloating
End Sub

Public Sub tblFloating_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tgcFEditor.MouseUp Y
End Sub

Public Sub tblFloating_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
     tgcFEditor.RowColChange tblFloating
End Sub

Public Sub tblProfitCenter_AfterColEdit(ByVal ColIndex As Integer)
    tgcEditorPrft.AfterColEdit ColIndex
End Sub

Public Sub tblProfitCenter_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, CANCEL As Integer)
     tgcEditorPrft.BeforeColEdit ColIndex, KeyAscii, CANCEL
End Sub

Public Sub tblProfitCenter_Change()
       If t_bStartupFlag Then
      Exit Sub
   End If
       tgcEditorPrft.Change
End Sub

Public Sub tblProfitCenter_GotFocus()
        
    Screen.MousePointer = vbDefault
    
    If t_bStartupFlag Then
      Exit Sub
   End If
         tgcEditorPrft.GotFocus
End Sub

Public Sub tblProfitCenter_KeyDown(KeyCode As Integer, Shift As Integer)
    tgcEditorPrft.KeyDown KeyCode, Shift
End Sub

Public Sub tblProfitCenter_KeyPress(KeyAscii As Integer)
    tgcEditorPrft.Keypress KeyAscii
    KeyAscii = 0
End Sub

Public Sub tblProfitCenter_LostFocus()
        tgcEditorPrft.LostFocus
End Sub

Public Sub tblProfitCenter_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If t_bStartupFlag Then
      Exit Sub
   End If
   tgcEditorPrft.RowColChange LastRow, LastCol
End Sub

Public Sub tblProfitCenter_SelChange(CANCEL As Integer)
    CANCEL = True
End Sub

Public Sub tblProfitCenter_UnboundReadData(ByVal RowBuf As RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
    If t_bStartupFlag Then
        Exit Sub
    End If
    tgcEditorPrft.ReadData RowBuf, StartLocation, ReadPriorRows
End Sub

'====
'Grid
'====
Public Sub tblTimeCard_AfterColEdit(ByVal ColIndex As Integer)
    
    Dim iCurrentRow As Integer
    
    tgcEditor.AfterColEdit ColIndex
     
    iCurrentRow = tgcEditor.GetCurrentRowNumber
    If ColIndex = 0 Then
        On Error Resume Next
        tgcEditor.CellValue(0, iCurrentRow) = _
            tfnFormatDate(tgcEditor.CellValue(0, iCurrentRow))
    End If
    
    If ColIndex = 2 Then
        On Error Resume Next
        tgcEditor.CellValue(2, iCurrentRow) = _
            UCase(tgcEditor.CellValue(2, iCurrentRow))
    End If
'   If ColIndex = 4 Then
'       On Error Resume Next
'       tgcEditor.CellValue(4, iCurrentRow) = _
'            tfnFormatDecimal(tgcEditor.CellValue(4, iCurrentRow), 2)
'   End If
    fnGetTotalDollars
    'If colIndex = 4 Or colIndex = 1 Then
    '    fnFillProfitTable
    'End If
    subSetFormStatus
 
End Sub

Public Sub tblTimeCard_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, CANCEL As Integer)
    tgcEditor.BeforeColEdit ColIndex, KeyAscii, CANCEL
    cmdUpdateInsertBtn.Enabled = False
End Sub

Public Sub tblTimeCard_Change()
    tgcEditor.Change
    'If tblTimeCard.col = COL_CODE Then
    '    tblTimeCard.Columns(COL_TYPE).Value = ""
    '    tgcEditor.CellValue(COL_TYPE, tgcEditor.GetCurrentRowNumber) = ""
    'End If
    If tblTimeCard.col = COL_DATE Then
        tblTimeCard.Columns(COL_PRFT).Value = ""
        tgcEditor.CellValue(COL_PRFT, tgcEditor.GetCurrentRowNumber) = ""
        tblTimeCard.Columns(COL_CODE).Value = ""
        tgcEditor.CellValue(COL_CODE, tgcEditor.GetCurrentRowNumber) = ""
        If t_nFormMode = ADD_MODE Then
            tblTimeCard.Columns(COL_TYPE).Value = ""
            tgcEditor.CellValue(COL_TYPE, tgcEditor.GetCurrentRowNumber) = ""
        End If
    End If
    
    If nDataStatus = DATA_LOADED Or t_nFormMode = ADD_MODE Then
        nDataStatus = DATA_CHANGED
    End If
End Sub

Public Sub tblTimeCard_Click()
    tgcFEditor.TableClick tblTimeCard
End Sub

Public Sub tblTimeCard_FirstRowChange()
    On Error Resume Next
    tgcFEditor.FirstRowChange
    tgcEditor.FirstRowChange
End Sub

Public Sub tblTimeCard_GotFocus()
    Screen.MousePointer = vbDefault
    
    
    tgcEditor.GotFocus
    If t_nFormMode = ADD_MODE Then
        tgcFEditor.GotFocus
        If tgcFEditor.ValidSelection Then
            tblTimeCard_AfterColEdit 1
            tblTimeCard_AfterColEdit 2
        End If
    End If
    If t_nFormMode = EDIT_MODE Then
        tgcFEditor.GotFocus
        If tgcFEditor.ValidSelection Then
            tblTimeCard_AfterColEdit 2
        End If
    End If
End Sub

Public Sub tblTimeCard_KeyDown(KeyCode As Integer, Shift As Integer)
    
    tgcFEditor.Keypress tblTimeCard, KeyCode
    
  '  If KeyCode = 38 And tblTimeCard.Row = 0 Then
  '       tblTimeCard.Scroll 0, -1
  '  End If
    '39 for rightarrow, 40 for downarrow
   ' If (KeyCode = 40 Or KeyCode = 39 Or KeyCode = vbKeyReturn) _
   '     And tblTimeCard.Row = 7 Then
   '      tblTimeCard.Scroll 0, 1
   ' End If
    
    If tgcFEditor.ValidSelection Then
        nDataStatus = DATA_CHANGED
    End If
    
    tgcEditor.KeyDown KeyCode, Shift

End Sub

Public Sub tblTimeCard_KeyPress(KeyAscii As Integer)
    
    If tblTimeCard.col = COL_CODE Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If

    If Not tgcEditor.Keypress(KeyAscii) Then
        KeyAscii = 0
    End If
    tgcFEditor.Keypress tblTimeCard, KeyAscii
          
End Sub

Public Sub tblTimeCard_LostFocus()
    tgcEditor.LostFocus
    tgcFEditor.LostFocus tblTimeCard
    fnGetTotalDollars
   'fnFillProfitTable
   'fnCheckupdateBtn
    subSetFormStatus
End Sub

Public Sub tblTimeCard_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tblTimeCard.col = 1 And Button = RIGHT_BUTTON Then
        frmContext.MouseDown Button, PROFITCENTER_UP
    End If
End Sub

Public Sub tblTimeCard_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    If t_bStartupFlag Then
        Exit Sub
    End If
    
    If tgcFEditor.RowColChange(tblTimeCard) Then
        tgcEditor.RowColChange LastRow, LastCol
    End If
    If LastCol = 1 Then
        If Not subCheckTotalPrft Then
            tgcEditor.ErrorMessage(COL_PRFT) = "Can't allow more than 5 different profit centers! "
        End If
    End If
    If LastCol = 4 Or LastCol = 1 Then
        fnFillProfitTable
    End If
    fnGetTotalDollars
    subSetFormStatus
End Sub

Public Sub tblTimeCard_Scroll(CANCEL As Integer)
    tgcFEditor.Scroll
End Sub

Public Sub tblTimeCard_SelChange(CANCEL As Integer)
    CANCEL = True
End Sub

Public Sub tblTimeCard_UnboundReadData(ByVal RowBuf As RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
    tgcEditor.ReadData RowBuf, StartLocation, ReadPriorRows
End Sub

Public Sub txtEmployeeName_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmContext.MouseDown Button, EMP_MST_UP
End Sub

Public Sub txtEmployeeNumber_Change()
    
    txtEmployeeName.Text = ""
    txtSSN.Text = ""
    tgcDropdown.Change txtEmployeeNumber
    cValidate.Change txtEmployeeNumber
    nDataStatus = DATA_INI
    subSetFormStatus
            
End Sub

Public Sub txtEmployeeNumber_Click()
    tgcDropdown.Click txtEmployeeNumber
End Sub

Public Sub txtEmployeeNumber_GotFocus()

    subTextSelected
 
    tgcDropdown.GotFocus txtEmployeeNumber
    cValidate.GotFocus txtEmployeeNumber
    Screen.MousePointer = vbDefault
    
    If tgcDropdown.SingleRecordSelected Then
        subCheckEmpNum
    End If
    
End Sub

Public Sub txtEmployeeNumber_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Screen.MousePointer = vbHourglass
    End If

    If tgcDropdown.Keypress(txtEmployeeNumber, KeyAscii) Then
        cValidate.Keypress txtEmployeeNumber, KeyAscii
        tfnRegExpControlKeyPress txtEmployeeNumber, KeyAscii, szLongPattern
    Else
        If KeyAscii = vbKeyReturn Then
            If tgcDropdown.SingleRecordSelected Then
                subCheckEmpNum
            End If
        End If
        KeyAscii = 0
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Public Sub txtEmployeeNumber_LostFocus()

   ' tgcDropdown.LostFocus txtEmployeeNumber
    
    If cValidate.LostFocus(txtEmployeeNumber, cmdEmployeeNumber, _
        txtEmployeeName, cmdEmployeeName, _
        txtSSN, cmdSSN, tblComboDropdown) Then
        If cValidate.ValidInput(txtEmployeeNumber) Then
            subEnableEmployee False
            tblTimeCard.Enabled = True
            subSetFocus tblTimeCard
            subSetFormStatus
        End If
    End If

End Sub

Public Sub cmdEmployeeNumber_Click()
    Screen.MousePointer = vbHourglass
    tgcDropdown.Click cmdEmployeeNumber
    Screen.MousePointer = vbDefault
End Sub

Public Sub txtEmployeeName_Change()
    txtSSN.Text = ""
    tgcDropdown.Change txtEmployeeName
    cValidate.Change txtEmployeeName
End Sub

Public Sub txtEmployeeName_Click()
    tgcDropdown.Click txtEmployeeName
End Sub

Public Sub txtEmployeeName_GotFocus()

    subTextSelected
    tgcDropdown.GotFocus txtEmployeeName
    cValidate.GotFocus txtEmployeeName
    Screen.MousePointer = vbDefault
    
    If tgcDropdown.SingleRecordSelected Then
        subCheckEmpNum
    End If
    
End Sub

Public Sub txtEmployeeName_KeyPress(KeyAscii As Integer)

    'KeyAscii = Asc(UCase(Chr(KeyAscii))) ' Weiping 12/17/99
    
    If KeyAscii = vbKeyReturn Then
        Screen.MousePointer = vbHourglass
    End If

    If tgcDropdown.Keypress(txtEmployeeName, KeyAscii) Then
        cValidate.Keypress txtEmployeeName, KeyAscii
        tfnRegExpControlKeyPress txtEmployeeName, KeyAscii, "^P{1,42}$"
    Else
        If KeyAscii = vbKeyReturn Then
            If tgcDropdown.SingleRecordSelected Then
                subCheckEmpNum
            End If
        End If
        KeyAscii = 0
    End If
    Screen.MousePointer = vbDefault
End Sub

Public Sub txtEmployeeName_LostFocus()
    
    If cValidate.LostFocus(txtEmployeeNumber, cmdEmployeeNumber, _
        txtEmployeeName, cmdEmployeeName, _
        txtSSN, cmdSSN, tblComboDropdown) Then
        If cValidate.ValidInput(txtEmployeeName) Then
            subEnableEmployee False
            subSetFocus tblTimeCard
            subSetFormStatus
        End If
    End If

End Sub

Public Sub cmdEmployeeName_Click()
    Screen.MousePointer = vbHourglass
    tgcDropdown.Click cmdEmployeeName
    Screen.MousePointer = vbDefault
End Sub

Public Sub txtEmployeeNumber_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmContext.MouseDown Button, EMP_MST_UP
End Sub

Public Sub txtSSN_Change()
    
    'txtEmployeeNumber.Text = ""
    'txtEmployeeName.Text = ""
    If txtSSN.Text = "" Then
        txtEmployeeName.Text = ""
    End If
    
    tgcDropdown.Change txtSSN
    cValidate.Change txtSSN
    tfnRegExpControlChange txtSSN, "^###(\-)?##(\-)?####$"
End Sub

Public Sub txtSSN_Click()
    tgcDropdown.Click txtSSN
End Sub

Public Sub txtSSN_GotFocus()

    subTextSelected
    tgcDropdown.GotFocus txtSSN
    cValidate.GotFocus txtSSN
    Screen.MousePointer = vbDefault
    
    If tgcDropdown.SingleRecordSelected Then
        subCheckEmpNum
    End If
    
End Sub

Public Sub txtSSN_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Screen.MousePointer = vbHourglass
    End If

    If tgcDropdown.Keypress(txtSSN, KeyAscii) Then
        cValidate.Keypress txtSSN, KeyAscii
        tfnRegExpControlKeyPress txtSSN, KeyAscii, "^###(\-)?##(\-)?####$"
    Else
        If KeyAscii = vbKeyReturn Then
            If tgcDropdown.SingleRecordSelected Then
                subCheckEmpNum
            End If
        End If
        KeyAscii = 0
    End If
    
    Screen.MousePointer = vbDefault
    
End Sub

Public Sub txtSSN_LostFocus()

    If cValidate.LostFocus(txtEmployeeNumber, cmdEmployeeNumber, _
        txtEmployeeName, cmdEmployeeName, _
        txtSSN, cmdSSN, tblComboDropdown) Then
        If cValidate.ValidInput(txtSSN) Then
            subEnableEmployee False
            tblTimeCard.Enabled = True
            subSetFocus tblTimeCard
            subSetFormStatus
        End If
    End If

End Sub

Public Sub cmdSSN_Click()
    Screen.MousePointer = vbHourglass
    tgcDropdown.Click cmdSSN
    Screen.MousePointer = vbDefault
End Sub


Public Sub txtSSN_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmContext.MouseDown Button, EMP_MST_UP
End Sub

Public Function fnValidCellValue(tgTable As TDBGrid, ByVal nCol As Integer, ByVal nRow As Long, sText As String) As Boolean
    
    Select Case tgTable.TabIndex
        Case tblTimeCard.TabIndex
            Select Case nCol
                Case 0
                    fnValidCellValue = fnValidDate(sText, nRow)
                Case 1
                     fnValidCellValue = fnValidProfitCtr(sText, nRow)
                Case 2
                     fnValidCellValue = fnValidPayCode(sText)
                Case 3
                     fnValidCellValue = True
                Case 4
                    fnValidCellValue = fnValidHours(sText)
            End Select
        Case tblProfitCenter.TabIndex
            Select Case nCol
                Case 0
                     fnValidCellValue = True 'fnValidProfitCenter(sText)
                Case 1
                     fnValidCellValue = True 'fnValidHours(sText)
            End Select
    End Select
End Function

Public Function fnInValidData(txtBox As Textbox) As Boolean
    
    Select Case txtBox.TabIndex
         Case txtEmployeeNumber.TabIndex
                fnInValidData = Not fnValidNumber(txtBox)
         Case txtEmployeeName.TabIndex
                fnInValidData = False 'Not fnValidName(txtBox)
         Case txtSSN.TabIndex
                fnInValidData = Not fnValidSSN(txtBox)
    End Select
End Function

Public Function fnExit() As Boolean
    
    If nDataStatus = DATA_CHANGED Then
        If Not tfnCancelExit(t_szEXIT_MESSAGE) Then
            Exit Function
        End If
    End If
    
    fnExit = True
End Function

'=================
'private functions
'=================
Private Function fnValidHours(szText As String) As Boolean
    
    fnValidHours = False
    If Trim(szText) = "" Then
        tgcEditor.ErrorMessage(COL_HOUR) _
            = "You must enter an Hour/Dollar Number"
        Exit Function
    End If
    fnValidHours = True
End Function

Private Sub fnGetTotal()
    Dim Index As Integer
    Dim recordIndex As Integer
    Dim recordArray() As Variant
    Dim toTalRowNumber As Integer
    Dim dValueofDollar As Double
    If Trim(txtTotalDollars.Text) = Trim(txtTotal.Text) Then
        Exit Sub
    End If
    toTalRowNumber = tgcEditorPrft.RowCount
    If toTalRowNumber <= 0 Then
        Exit Sub
    End If
    dValueofDollar = 0
    For Index = 0 To toTalRowNumber - 1 Step 1
        tgcEditorPrft.GetRow recordArray(), recordIndex, Index
        dValueofDollar = dValueofDollar + tfnRound(recordArray(1), 2)
    Next Index
    txtTotal.Text = tfnFormatDecimal(dValueofDollar, 2)
End Sub

Private Sub fnGetTotalDollars()
    Dim Index As Integer
    Dim recordIndex As Integer
    Dim recordArray() As Variant
    Dim toTalRowNumber As Integer
    Dim dValueofDollar As Double
    On Error GoTo errorDollar
        toTalRowNumber = tgcEditor.RowCount
        If toTalRowNumber <= 0 Then
            Exit Sub
        End If
        dValueofDollar = 0
        For Index = 0 To toTalRowNumber - 1 Step 1
            tgcEditor.GetRow recordArray(), recordIndex, Index
            dValueofDollar = dValueofDollar + tfnRound(recordArray(4), 2)
        Next Index
        txtTotalDollars.Text = tfnFormatDecimal(dValueofDollar, 2)
    Exit Sub
errorDollar:
    Resume Next
End Sub

Private Sub subCheckEmpNum()
    
    If cValidate.FirstInvalidInput < 0 Then
        subEnableEmployee False
        'efraInnerFrame.Enabled = True
        tblTimeCard.Enabled = True
        tblProfitCenter.Enabled = True
        If t_nFormMode = EDIT_MODE Then
            fnFillTimeCard
            nDataStatus = DATA_LOADED
        End If
        
    End If
'    tblTimeCard.Enabled = True
    subSetFocus tblTimeCard
End Sub

Private Sub fnFillProfitTable()
    ' this sub is used to caclulate the sum of different name
    ' A(n, 1),..... A(n, 9)----> B(n, 1+....9)
    Dim nCenterName(1, 20) As Variant
    Dim nCenterList() As Variant
    Dim Index As Integer
    Dim inDex2 As Integer
    Dim toTalRowNumber As Integer
    Dim toTal2 As Integer
    Dim recordArray() As Variant
    Dim recordIndex As Integer
    Dim NumberCenter As Integer
    Dim CenterExsit As Boolean
    Dim sTempInteger As Double
    
    On Error GoTo errorFillTable
'    If Trim(txtTotalDollars.Text) = Trim(txtTotal.Text) And t_nFormMode = EDIT_MODE Then
'          Exit Sub
'    End If
    
 
    
    toTalRowNumber = tgcEditor.RowCount
    For Index = 0 To 20 Step 1
          nCenterName(0, Index) = -1
          nCenterName(1, Index) = 0
    Next Index
    NumberCenter = 0
    For Index = 0 To toTalRowNumber - 1 Step 1
    'Get a row then check if The profit center exsist or not
    ' If it exsist, then add amount of hours
    ' if it does not exsist, then add a new profit center, NumberCenter+
        CenterExsit = False
        tgcEditor.GetRow recordArray(), recordIndex, Index
            For inDex2 = 0 To NumberCenter Step 1
                If Trim(nCenterName(0, inDex2)) = Trim(recordArray(1)) Then
                    sTempInteger = tfnRound(nCenterName(1, inDex2), 2) + tfnRound(recordArray(4), 2)
                    nCenterName(1, inDex2) = tfnFormatDecimal(sTempInteger, 2)
                    CenterExsit = True
                    Exit For
                End If
            Next inDex2
        If CenterExsit = False Then
              nCenterName(0, NumberCenter) = Trim(recordArray(1))
              nCenterName(1, NumberCenter) = tfnRound(recordArray(4), 2)
              NumberCenter = NumberCenter + 1
        End If
    Next Index
    
    If NumberCenter <= 0 Then
        Exit Sub
    End If
    ReDim nCenterList(1, NumberCenter - 1)
    For Index = 0 To NumberCenter - 1 Step 1
        nCenterList(0, Index) = nCenterName(0, Index)
        nCenterList(1, Index) = nCenterName(1, Index)
    Next Index
    tgcEditorPrft.ClearData
    tgcEditorPrft.FillWithArray nCenterList
    toTal2 = tgcEditorPrft.RowCount
    tgcEditorPrft.GetRow recordArray(), recordIndex, toTal2 - 1
    If Trim(recordArray(0)) = "" Then
        tgcEditorPrft.DeleteRow tgcEditorPrft.Bookmark(toTal2 - 1)
    End If
    fnGetTotal
    
    Exit Sub
errorFillTable:
     Resume Next
End Sub

Private Function subCheckTotalPrft() As Boolean
    Dim curRow As Integer
    Dim TotalRow As Integer
    Dim prftArray(5) As Variant
    Dim rowIndex As Integer
    
     subCheckTotalPrft = True
    TotalRow = tgcEditor.RowCount - 1
    If TotalRow <= 5 Then
        Exit Function
    End If

    prftArray(0) = tgcEditor.CellValue(1, 0)
    Dim newIndex As Integer
    For newIndex = 1 To 5
        prftArray(newIndex) = -1
    Next newIndex
    
    Dim prftSum As Integer
    Dim prftDuplFlag As Boolean
    
    prftSum = 0
    For rowIndex = 1 To TotalRow
        prftDuplFlag = False
        For newIndex = 0 To prftSum
            
            If tgcEditor.CellValue(1, rowIndex) = prftArray(newIndex) Then
                prftDuplFlag = True
                Exit For
            End If
            
        Next newIndex
        If Not prftDuplFlag Then
            prftSum = prftSum + 1
            prftArray(prftSum) = tgcEditor.CellValue(1, rowIndex)
            If prftSum > 4 Then
                subCheckTotalPrft = False
                
                Exit Function
            End If
        End If
    Next rowIndex
   
End Function

Private Function fnFillTimeCard() As Boolean
    Dim sSql As String
    Dim rsTemp As Recordset
    Dim Index As Long
    Dim n As Integer
    On Error GoTo errorFilltalble
    If b_TableFilled Then
           fnFillTimeCard = False
        Exit Function
    End If
    sSql = " SELECT prh_date,prh_prft_ctr,prh_pay_code,prh_pay_type,prh_hours,prh_source FROM pr_hours WHERE prh_empno=" & txtEmployeeNumber.Text
    sSql = sSql & " ORDER BY prh_date,prh_prft_ctr ,prh_pay_code  "
    
    On Error GoTo errH
    
    If tfnLockRow("prfhours", "pr_hours", sSql) Then
        tgcEditor.FillWithSQL t_dbMainDatabase, sSql
    End If
    fnFillTimeCard = True
    b_TableFilled = True
    fnGetTotalDollars
    fnFillProfitTable
    fnGetTotal
    tblTimeCard.SetFocus
    If tgcEditor.RowCount > 0 Then
        cmdDeleteBtn.Enabled = True
        'subkeepRecord
    End If
    
Exit Function
errorFilltalble:
    Resume Next
    Exit Function
errH:
    tfnErrHandler "fnFill TimeCard", sSql
    
End Function

Private Sub fnEnableControls(beOff As Boolean)
    txtEmployeeNumber.Enabled = beOff
    subEnableSearchbtn cmdEmployeeNumber, beOff
    txtEmployeeName.Enabled = beOff
    subEnableSearchbtn cmdEmployeeName, beOff
    txtSSN.Enabled = beOff
    subEnableSearchbtn cmdSSN, beOff
    tblTimeCard.Enabled = beOff
    tblProfitCenter.Enabled = beOff
End Sub

Private Sub subEnableRefresh(bStatus As Boolean)

    cmdRefreshSelectBtn.Enabled = bStatus
    
    If bStatus Then
        cmdDeleteBtn.Enabled = False
    End If
    
End Sub

Public Sub subEnableSearchbtn(cmdButton As Control, bYesNo As Boolean)

    If bYesNo Then
        cmdButton.Picture = frmContext.LoadPicture(SEARCH_UP)
    Else
        cmdButton.Picture = frmContext.LoadPicture(SEARCH_DOWN)
    End If
        cmdButton.Enabled = bYesNo
    
End Sub

Private Sub subSetupCombos()
    If tgcDropdown Is Nothing Then
        Set tgcDropdown = CreateObject(t_szOLECOMBO)
        Set tgcDropdown.DBEngine = t_engFactor
        With tgcDropdown
            Set .Form = m_Form
            Set .DataBase = t_dbMainDatabase
            Set .DataLink = datComboDropDown
            Set .Table = tblComboDropdown
        End With
    End If
    
    With tgcDropdown
        .AddCombo
        .SQL = "SELECT empno, empname, empssn FROM " & sTempTable
        .AddComboBox txtEmployeeNumber, cmdEmployeeNumber, "empno", tgcDropdown.SQL_LONG_TYPE
        .AddComboBox txtEmployeeName, cmdEmployeeName, "empname", tgcDropdown.SQL_STRING_TYPE(42)
        .AddComboBox txtSSN, cmdSSN, "empssn", tgcDropdown.SQL_STRING_TYPE(11)
    End With
End Sub

Private Sub subSetValidtion()
    Set cValidate = New cValidateInput
    With cValidate
    Set .StatusBar = ffraStatusbar
    Set .Form = m_Form
    .AddEditBox txtEmployeeNumber, "Enter Employee Number."
    .AddEditBox txtEmployeeName, "Enter Employee Name."
    .AddEditBox txtSSN, "Enter Social Security Number."
    .ESCControl = tblComboDropdown
    .MinTabIndex = tbToolbar.TabIndex
    .MaxTabIndex = efraBridge.TabIndex + 1
    Set .ControlForFocus = efraBridge
    Set .LastBox = txtEmployeeNumber
    .SetFirstControls cmdDeleteBtn, cmdUpdateInsertBtn, cmdExitCancelBtn, cmdExitBtn
    End With
End Sub

Private Sub subSetGrid()
    Set tgcEditor = New clsTGSpreadSheet
    With tgcEditor
        Set .Table = tblTimeCard
        Set .Form = m_Form
        Set .StatusBar = ffraStatusbar
        sDecimalString = tfnDecimalPattern(8, 2)
        .AddEditColumn COL_DATE, "Enter Clock-in Date", szDatePattern
        .AddEditColumn COL_PRFT, "Enter Profit Center", szIntegerPattern
        .AddEditColumn COL_CODE, "Enter Pay Code", "^P{1,4}$"
        '.AddEditColumn COL_TYPE, "Enter Pay Type", "^P{0,2}$"   '***
        .AddEditColumn COL_HOUR, "Enter Hours or Dollars", sDecimalString
        COL_SOURCE = tgcEditor.AddHiddenField("prh_source")
        .DisplayFormat(COL_PRFT) = "#####"
        .DisplayFormat(4) = "###,###,##0.00"
        
        .ClearData
    End With
    
    Set tgcEditorPrft = New clsTGSpreadSheet
    Set tgcEditorPrft.Table = tblProfitCenter
    Set tgcEditorPrft.Form = m_Form
    Set tgcEditorPrft.StatusBar = ffraStatusbar
    tgcEditorPrft.AddEditColumn 0, "NOT editable at Profit Center Total table", "^#{1,5}$"
    tgcEditorPrft.AddEditColumn 1, "NOT editable at Profit Center Total table"
    tgcEditorPrft.DisplayFormat(1) = "###,###,##0.00"
    tgcEditorPrft.ClearData
    tgcEditorPrft.RemoveEditColumn 0
    tgcEditorPrft.RemoveEditColumn 1
End Sub

Public Sub Form_Initialize() 'called before Form_Load
    
    t_bStartupFlag = True
    t_bDataChanged = False
    t_bUpdateTable = False
    
    t_nFormMode = IDLE_MODE
    
    ' ** change the help file for the application
    sTempTable = "temp_pr_master" 'tai
    sTempTable1 = "temp_pr_pay" 'tai
    bWarned = False
    bTableExists = False
End Sub

Private Function fnDeleteData() As Boolean
    Dim sSQL2 As String
    Dim sSql As String
    Dim rsTemp As Recordset
    Dim recordIndex As Integer
    Dim CurentRow As Integer
    Dim sStringForT As String * 1
    
    CurentRow = tgcEditor.GetCurrentRowNumber
    
    On Error GoTo ErrorDeleteData
    If tgcEditor.CellValue(COL_SOURCE, CurentRow) = "T" Then
        fnDeleteData = False
        MsgBox "This record can't be deleted!"
        Exit Function
    End If
    sSQL2 = "UPDATE pr_hours SET prh_hours = 0, prh_chk_lnk = -1 "
    sSQL2 = sSQL2 & " WHERE prh_empno= " & txtEmployeeNumber.Text
    sSQL2 = sSQL2 & " AND prh_date= " & tfnDateString(tgcEditor.CellValue(0, CurentRow), True)
    sSQL2 = sSQL2 & " AND prh_prft_ctr = " & tfnRound(tgcEditor.CellValue(1, CurentRow))
    sSQL2 = sSQL2 & " AND prh_pay_code = " & tfnSQLString(tgcEditor.CellValue(2, CurentRow))
    sSQL2 = sSQL2 & " AND prh_pay_type = " & tfnSQLString(tgcEditor.CellValue(3, CurentRow))
  On Error GoTo errHH
   
    t_dbMainDatabase.ExecuteSQL sSQL2
    tgcEditor.CellValue(4, CurentRow) = 0
    fnDeleteData = True
    Exit Function
ErrorDeleteData:
    tfnErrHandler "fnDeleteData", sSql
    fnDeleteData = False
    Exit Function
errHH:
    tfnErrHandler "fnDeleteData", sSQL2
    fnDeleteData = False
End Function

Private Function fnInsertUpdate() As Boolean
    If Not subCheckGridDuplicate Then
        Exit Function
    End If
    If t_nFormMode = ADD_MODE Then
        fnInsertUpdate = fnInsert
    End If
    If t_nFormMode = EDIT_MODE Then
        fnInsertUpdate = fnUpdate
    End If
End Function

Private Function subCheckGridDuplicate() As Boolean
    Dim Index As Integer
    Dim RowCount As Integer
    Dim index1 As Integer
    Dim ColIndex As Integer
    Dim DuplicateFlag As Boolean
    
    RowCount = tgcEditor.RowCount - 1
    For Index = 0 To RowCount
        For index1 = (Index + 1) To RowCount
                DuplicateFlag = True
                For ColIndex = 0 To 3
                    If tgcEditor.CellValue(ColIndex, Index) <> tgcEditor.CellValue(ColIndex, index1) Then
                        DuplicateFlag = False
                        Exit For
                    End If
                Next ColIndex
                If DuplicateFlag Then
                    tfnSetStatusBarError "The row #" & Val(Index + 1) & " and row# " & Val(index1 + 1) & " are identical! "
                    subCheckGridDuplicate = False
                    cmdUpdateInsertBtn.Enabled = False
                    Exit Function
                End If
        Next index1
    Next Index
    subCheckGridDuplicate = True
        
    
End Function

Private Function fnUpdate() As Boolean  'delete first then insert back, since no unique key
                                        'in table pr_hours
                                        'yan
    Dim Index As Long
    Dim sSQL1 As String
    Dim sSql As String
    Dim rsTemp As Recordset
    Dim recordArray() As Variant
    Dim recordIndex As Integer
    Dim CurentRow As Integer
    Dim sStringForT As String * 1
    Dim TotalForLoop As Integer
    Dim toTalRowNumber As Integer
  
    Dim CurRowChanged As Boolean
    Dim ColIndex As Integer
    
    toTalRowNumber = tgcEditor.RowCount
    If toTalRowNumber <= 0 Then
        fnUpdate = False
        Exit Function
    End If
    
On Error GoTo errorUpdate2
    sSql = " delete from pr_hours where prh_empno = " & txtEmployeeNumber.Text
    t_dbMainDatabase.ExecuteSQL sSql
    
    For Index = 0 To toTalRowNumber - 1 Step 1
        CurRowChanged = False
        tgcEditor.GetRow recordArray(), recordIndex, Index
        For ColIndex = 0 To 4
            If recordArray(ColIndex) <> tgcEditor.CellValue(ColIndex, Index) Then
                CurRowChanged = True
                Exit For
            End If
        Next ColIndex
        
        sSql = " INSERT INTO pr_hours ( prh_empno, prh_ssn, prh_date, "
        sSql = sSql & " prh_prft_ctr, prh_shl, prh_pay_code, prh_pay_type, "
        sSql = sSql & " prh_hours, prh_mgr_ovride, prh_host_ovride, "
        sSql = sSql & " prh_host_added, prh_chk_lnk, prh_source ) "
        sSql = sSql & " VALUES ( " & txtEmployeeNumber.Text & " ,"
        sSql = sSql & tfnSQLString(txtSSN) & ","
        sSql = sSql & tfnSQLString(tgcEditor.CellValue(COL_DATE, Index)) & ","
        sSql = sSql & tfnSQLString(tgcEditor.CellValue(COL_PRFT, Index)) & ","
        sSql = sSql & " NULL, "
        sSql = sSql & tfnSQLString(tgcEditor.CellValue(COL_CODE, Index)) & ","
        sSql = sSql & tfnSQLString(tgcEditor.CellValue(COL_TYPE, Index)) & ","
        sSql = sSql & tfnSQLString(tgcEditor.CellValue(COL_HOUR, Index)) & ","
        sSql = sSql & " '0', "
        If (tgcEditor.CellValue(COL_SOURCE, Index) = "T") And CurRowChanged Then
            sSql = sSql & " '1', "
        Else
            sSql = sSql & " '0', "
        End If
        sSql = sSql & " '1', "
        If tgcEditor.CellValue(COL_HOUR, Index) = "" Or tgcEditor.CellValue(COL_HOUR, Index) = 0 Then
            sSql = sSql & " -1, "
        Else
            sSql = sSql & " 0, "
        End If
        sSql = sSql & tfnSQLString(tgcEditor.CellValue(COL_SOURCE, Index)) & " )"
         
         
         
         On Error GoTo errorUpdate2
         t_dbMainDatabase.ExecuteSQL sSql
    Next
    fnUpdate = True
    Exit Function
errorUpdate:
    tfnErrHandler "fnUpdate", sSQL1
    fnUpdate = False
Exit Function
errorUpdate2:
    tfnErrHandler "fnUpdate", sSql
    fnUpdate = False
End Function

Private Function fnInsert() As Boolean
    Dim Index As Integer
    Dim sSql As String
    Dim rsTemp As Recordset
    Dim recordArray() As Variant
    Dim recordIndex As Integer
    Dim CurentRow As Integer
    Dim sStringForT As String * 1
    Dim TotalForLoop As Integer
    Dim toTalRowNumber As Integer
    On Error GoTo errorInsert
    toTalRowNumber = tgcEditor.RowCount
    tgcEditor.GetRow recordArray(), recordIndex, toTalRowNumber - 1
    If Trim(recordArray(0)) = "" Or Trim(recordArray(1)) = "" _
        Or Trim(recordArray(2)) = "" Then 'Or Trim(recordArray(3)) = "" Then 'tai comment out last ""
        TotalForLoop = toTalRowNumber - 2
    Else
        TotalForLoop = toTalRowNumber - 1
    End If
    For Index = 0 To TotalForLoop Step 1
        tgcEditor.GetRow recordArray(), recordIndex, Index
        sSql = " INSERT INTO pr_hours ( prh_empno, prh_ssn, prh_date, "
        sSql = sSql & " prh_prft_ctr, prh_shl, prh_pay_code, prh_pay_type, "
        sSql = sSql & " prh_hours, prh_mgr_ovride, prh_host_ovride, "
        sSql = sSql & " prh_host_added, prh_chk_lnk, prh_source ) "
        sSql = sSql & " VALUES ( " & txtEmployeeNumber.Text & " ,"
        sSql = sSql & tfnSQLString(txtSSN) & ","
        sSql = sSql & tfnDateString(recordArray(0), True) & ","
        sSql = sSql & Trim(recordArray(1)) & ","
        sSql = sSql & " NULL, "
        sSql = sSql & tfnSQLString(Trim(recordArray(2))) & ","
        sSql = sSql & tfnSQLString(Trim(recordArray(3))) & ","
        sSql = sSql & tfnRound(recordArray(4), 2) & ","
        sSql = sSql & " '0', '0', '1', "
        If Trim(recordArray(4)) = "" Or tfnRound(recordArray(4), 2) = 0 Then
            sSql = sSql & " -1, "
        Else
            sSql = sSql & " 0, "
        End If
        sSql = sSql & " 'P' )"
        t_dbMainDatabase.ExecuteSQL sSql
    Next Index
    fnInsert = True
    Exit Function
errorInsert:
    tfnErrHandler "fnInsert", sSql
    fnInsert = False
                 
End Function

Private Function fnRefresh()
    tgcEditor.ClearData
    nDataStatus = DATA_INI
    b_TableFilled = False
    fnFillTimeCard
    b_TableFilled = True
    fnGetTotalDollars
    fnFillProfitTable
    fnGetTotal
    cmdRefreshSelectBtn.Enabled = False
    tblTimeCard.SetFocus
    fnRefresh = True
End Function

Private Sub subAddEmployee()
    
    If tblTimeCard.Enabled Then
        subSetFocus tblTimeCard
    ElseIf txtEmployeeNumber.Enabled Then
        subSetFocus txtEmployeeNumber
    ElseIf cmdAddBtn.Enabled Then
        subSetFocus cmdAddBtn
    Else
        subSetFocus cmdExitCancelBtn
    End If
End Sub

Private Sub subCancel()

    If nDataStatus >= DATA_CHANGED Then 'If t_bDataChanged = True Then

        If tfnCancelExit(t_szCANCEL_MESSAGE) = False Then
            Exit Sub
        End If
    End If
    If t_nFormMode = EDIT_MODE Then
        tfnUnlockRow
    End If
    
    tfnResetScreen 'reset all the buttons
    
    t_nFormMode = IDLE_MODE
    
    t_bDataChanged = False
End Sub

'======================
'form support functions
'======================
Private Sub tfnSetStatusBarMessage(szMessage As String)
    
  '  If t_bStartupFlag Then
  '      Exit Sub
  '  End If
    
    ffraStatusbar.ForeColor = STANDARD_TEXT_COLOR
    ffraStatusbar.Font.Bold = False
    ffraStatusbar.Caption = szMessage
    ffraStatusbar.Refresh

End Sub

Private Sub tfnSetStatusBarError(szErrorMessage As String, Optional vNoBeep As Variant)
    
    ffraStatusbar.ForeColor = ERROR_TEXT_COLOR
    ffraStatusbar.Font.Bold = True
    ffraStatusbar.Caption = szErrorMessage
    If IsMissing(vNoBeep) Then
        Beep
    End If
    ffraStatusbar.Refresh

End Sub

Private Sub tfnSetStatusBarCorrect(szCorrectMessage As String)
    
    ffraStatusbar.ForeColor = CORRECT_TEXT_COLOR
    ffraStatusbar.Font.Bold = True
    ffraStatusbar.Caption = szCorrectMessage
    ffraStatusbar.Refresh

End Sub

Private Sub tfnResetScreen()
    
    On Error Resume Next
    txtEmployeeNumber.Text = ""
    txtEmployeeName.Text = ""
    txtSSN.Text = ""
    txtTotalDollars.Text = ""
    txtTotal.Text = ""
    
    cmdDeleteBtn.Enabled = False
    cmdUpdateInsertBtn.Enabled = False
    cmdRefreshSelectBtn.Enabled = False
    
    cmdUpdateInsertBtn.Caption = t_szCAPTION_UPDATE
    cValidate.ResetFlags
    
    tfnSetStatusBarMessage t_szADD
    
    frmContext.ButtonEnabled(PROFITCENTER_UP) = True
    frmContext.ButtonEnabled(EMP_MST_UP) = True
    
    frmContext.ButtonEnabled(CANCEL_UP) = False
    cmdExitCancelBtn.Enabled = False
    mnuCancel.Enabled = False
    
    tgcEditor.ClearData
    tgcEditorPrft.ClearData
    
    bNeverDelet = True
    b_TableFilled = False
    
    fnEnableControls False
    DoEvents
    tblTimeCard.Enabled = False
    
    tgcEditor.ClearData
    tgcEditorPrft.ClearData
    DoEvents
    
    cmdAddBtn.Enabled = True
    cmdEditBtn.Enabled = True
    subSetFocus cmdAddBtn
End Sub

Private Sub subCreateTempTable()
    
    Dim sSql As String
    Dim sTempTemp As String
    Dim sTempTemp1 As String
    
    #If PROTOTYPE Then
        Exit Sub
    #End If
    
    On Error GoTo err
    
    sTempTemp = "temp_table1"
    sTempTemp1 = "temp_table2"
        
    sSql = "SELECT DISTINCT prm_empno AS empno, trim(prm_last_name) || ', ' || trim(prm_first_name) " _
       & "AS empname, prm_ssn AS empssn FROM pr_master,pr_security " _
       & "WHERE prs_code>=prm_security_code AND prs_name=" & tfnSQLString(tfnGetUserName) _
       & " INTO TEMP " & sTempTable
     
    fnExecuteSQL (sSql)
                        
    sSql = "SELECT prpa_pay_code AS pcode, " _
        & "prpa_desc AS pdesc, prpa_tmi_type AS ptype " _
        & "FROM pr_pay WHERE " _
        & "(prpa_type = 'P' AND prpa_calc_method = 'H' OR (prpa_type = 'N' AND prpa_calc_method = 'D'))" _
        & " Union"
    sSql = sSql & " SELECT prpa_pay_code AS pcode, " _
        & "prpa_desc AS pdesc, prpa_tmi_ottype AS ptype " _
        & "FROM pr_pay WHERE " _
        & "(prpa_type = 'P' AND prpa_calc_method = 'H' " _
        & "OR (prpa_type = 'N' AND prpa_calc_method = 'D')) INTO TEMP " _
        & sTempTemp
    fnExecuteSQL sSql
                    
    sSql = "SELECT prep_empno AS pempno, prep_pay_code, " _
        & "prep_year AS pyear FROM pr_emp_pay " _
        & "WHERE prep_active = 'Y' INTO TEMP " & sTempTemp1
    fnExecuteSQL sSql

    sSql = "SELECT pyear, pcode, ptype, pdesc, pempno " _
        & "FROM " & sTempTemp & ", " & sTempTemp1 _
        & " WHERE pcode = prep_pay_code INTO TEMP " & sTempTable1
    fnExecuteSQL sSql
Exit Sub
err:
    tfnErrHandler "subcreatetemptable", sSql
    
End Sub

Private Function fnValidNumber(txtBox As Textbox)
        
    fnValidNumber = False
    
    If Trim(txtBox.Text) = "" Then
        cValidate.SetErrorMessage txtBox, "You must enter an Employee Number"
        Exit Function
    End If
    
    Dim sSql As String
    
    On Error GoTo err
    
    sSql = "SELECT empno FROM " & sTempTable _
        & " WHERE empno = " & CLng(txtBox.Text)
        
    If GetRecordCount(sSql) <= 0 Then
        Exit Function
    End If
    
    sSql = " select * from pr_master where prm_empno = " & tfnSQLString(txtBox)
    Dim rsTemp As Recordset
    Set rsTemp = t_dbMainDatabase.OpenRecordset(sSql, dbOpenSnapshot, dbSQLPassThrough)
    If rsTemp!prm_date_termed <> "" Then
        cValidate.SetErrorMessage txtBox, "This is a terminated employee! Can't modify the pay hours."
        Exit Function
    End If
    Dim empSecurityId As String
    empSecurityId = rsTemp!prm_security_code
    

    sSql = "select prs_code from pr_security where prs_name = " & tfnSQLString(tfnGetUserName)
    Set rsTemp = t_dbMainDatabase.OpenRecordset(sSql, dbOpenSnapshot, dbSQLPassThrough)
    Dim userSecurityID As String
    If GetRecordCount(sSql) <= 0 Then
        MsgBox "The user doesn't have access to any employee records."
        cValidate.SetErrorMessage txtBox, "The user doesn't have access to any employee records."
        Exit Function
    End If
    userSecurityID = rsTemp!prs_code
    'sam on 05/15/00 per bob on 279344
    If userSecurityID < empSecurityId Then   'sam added
    ' If userSecurityID <= empSecurityId Then  'sam commented out
        cValidate.SetErrorMessage txtBox, "The employee has higher security code, can't be edited."
        Exit Function
    End If
    
    
    If t_nFormMode = EDIT_MODE Then
        sSql = "SELECT prh_empno FROM pr_hours WHERE prh_empno = " & CLng(txtBox.Text)
    Else
        sSql = "SELECT pempno FROM " & sTempTable1 & " WHERE pempno = " & CLng(txtBox.Text)
    End If
                
    If GetRecordCount(sSql) > 0 Then
        If t_nFormMode = EDIT_MODE Then
            Dim sSqlLock As String
            sSqlLock = "SELECT * FROM pr_hours WHERE prh_empno = " _
                 & CLng(txtBox.Text)
            If Not tfnLockRow("prfhours", "pr_hours", sSqlLock) Then
                subSetFocus txtEmployeeNumber
                subTextSelected
                Exit Function
            End If
        End If
        fnValidNumber = True
    Else
        cValidate.SetErrorMessage txtBox, "Entered invalid Employee Number"
    End If
        
    Exit Function

err:
    tfnErrHandler "fnValidNumber", sSql

End Function

Private Function fnValidSSN(txtBox As Textbox)
    
    fnValidSSN = False
    
    If Trim(txtBox.Text) = "" Then
        cValidate.SetErrorMessage txtBox, "You must enter an Employee SSN"
        Exit Function
    End If
    
    Dim sSql As String
    Dim rsTemp As Recordset
    
    On Error GoTo err
    
    sSql = "SELECT empno, empname, empssn FROM " & sTempTable _
        & " WHERE empssn = " & tfnSQLString(txtBox.Text)
        
    If GetRecordSet(rsTemp, sSql) <= 0 Then
        Exit Function
    End If
    
    If t_nFormMode = EDIT_MODE Then
        sSql = "SELECT prh_ssn FROM pr_hours WHERE prh_ssn = " & tfnSQLString(txtBox.Text)
    Else
        sSql = "SELECT pempno FROM " & sTempTable1 & " WHERE pempno = " & rsTemp!empno
    End If
                
    If GetRecordCount(sSql) > 0 Then
        If t_nFormMode = EDIT_MODE Then
            Dim sSqlLock As String
            sSqlLock = "SELECT * FROM pr_hours WHERE prh_ssn = " & tfnSQLString(txtBox.Text)
            If Not tfnLockRow("prfhours", "pr_hours", sSqlLock) Then
                'cmdExitCancelBtn_Click
                subSetFocus txtEmployeeNumber
                subTextSelected
                Exit Function
            End If
        End If
        fnValidSSN = True
    Else
        cValidate.SetErrorMessage txtBox, "Entered invalid Employee SSN"
    End If
    
    Exit Function

err:
    tfnErrHandler "fnValidSSN", sSql
    
End Function

Private Sub subReSetComboSql()
    
    Dim sSql As String
    Dim sSQL1 As String
    
    sSql = "SELECT empno, empname, empssn FROM " & sTempTable
    
    If t_nFormMode = EDIT_MODE Then
        sSql = sSql & " WHERE empno IN (SELECT prh_empno FROM pr_hours)"
    Else
        sSql = sSql & " WHERE empno IN (SELECT pempno FROM " & sTempTable1 & ")"
    End If
    sSql = sSql & " and empno in (select prm_empno from pr_master where prm_date_termed is null) "
    tgcDropdown.ComboSQL(txtEmployeeNumber) = sSql
        
End Sub

Private Sub subEnableEmployee(bFlag As Boolean)
    txtEmployeeNumber.Enabled = bFlag
    txtEmployeeName.Enabled = bFlag
    txtSSN.Enabled = bFlag
    subEnableSearchButton cmdEmployeeNumber, False
    subEnableSearchButton cmdEmployeeName, False
    subEnableSearchButton cmdSSN, False
End Sub

Private Sub subSetFormStatus()
    
    If t_nFormMode = EDIT_MODE _
        And nDataStatus = DATA_CHANGED Then
            'subEnableDelete False
            subEnableRefresh True
    End If
    
    If tgcEditor.ValidData And fnGetField(tgcEditor.CellValue(0, 0)) <> "" And _
        nDataStatus = DATA_CHANGED Then
            cmdUpdateInsertBtn.Enabled = True
    Else
        cmdUpdateInsertBtn.Enabled = False
    End If
    If t_nFormMode = EDIT_MODE _
        And nDataStatus = DATA_INI Then
            cmdDeleteBtn.Enabled = False
            subEnableRefresh False
            cmdUpdateInsertBtn.Enabled = False
    End If
                                
End Sub

Private Function fnValidDate(ByVal szText As String, _
                                  ByVal lRow As Long) As Boolean
            
    fnValidDate = False
    If Trim(szText) = "" Then
        tgcEditor.ErrorMessage(COL_DATE) _
            = "You must enter a Clock-in Date"
        Exit Function
    End If
    
    'see if is a date
    Dim nError As Integer
    If Not IsFactorDate(szText, nError) Then 'the format of sztext has been changed here
        If nError = 1 Then
            tgcEditor.ErrorMessage(COL_DATE) = "Not a date format"
        Else
            tgcEditor.ErrorMessage(COL_DATE) = "Not a valid date"
        End If
        Exit Function
    End If
    
    
    'col_prft is not empty
    
    Dim sSql As String
    Dim rsTemp As Recordset
            
    'Check duplication in table
    On Error GoTo err
    
    'check in table gl_period
    sSql = "SELECT glp_status" _
        & " FROM gl_period WHERE glp_end_dt >= " & tfnSQLString(CDate(szText)) _
        & " AND glp_beg_dt <= " & tfnSQLString(CDate(szText))
    If GetRecordSet(rsTemp, sSql) > 0 Then
        If rsTemp!glp_status = "O" Then
            fnValidDate = True
        ElseIf rsTemp!glp_status = "W" Then
            fnValidDate = True
            If Not bWarned Then
                MsgBox "This G/L Period is about to close!", vbExclamation
                bWarned = True
            Else
                bWarned = False
            End If
        ElseIf rsTemp!glp_status = "C" Then
            tgcEditor.ErrorMessage(COL_DATE) _
                = "This G/L Period is already closed"
        Else
            tgcEditor.ErrorMessage(COL_DATE) _
                = "Invalid G/L Period Status"
        End If
    Else
        tgcEditor.ErrorMessage(COL_DATE) _
            = "There is no valid G/L Period containing this Date"
    End If
    
    Exit Function

err:
    tfnErrHandler "fnvaliddate", sSql
    
End Function

Private Function fnValidProfitCtr(ByVal szText As String, _
                                  ByVal lRow As Long) As Boolean
            
    fnValidProfitCtr = False
    If Trim(szText) = "" Then
        tgcEditor.ErrorMessage(COL_PRFT) _
            = "You must enter a Profit Center Number"
        Exit Function
    End If
            
    If Not IsDate(tgcEditor.CellValue(COL_DATE, lRow)) Then
        tblTimeCard.col = COL_DATE
        Exit Function
    End If
    
    On Error GoTo err
    
    'col_date is not empty
    Dim sDate As Date
    Dim sCode As String
    Dim sType As String
    
    sDate = tfnFormatDate(tgcEditor.CellValue(COL_DATE, lRow))
    sCode = fnGetField(tgcEditor.CellValue(COL_CODE, lRow))
    sType = fnGetField(tgcEditor.CellValue(COL_TYPE, lRow))
    
    Dim sSql As String
    Dim rsTemp As Recordset
        
    'check if in sys_prft_ctr
    sSql = " SELECT prft_ctr, prft_name FROM sys_prft_ctr " _
        & "WHERE prft_ctr = " & CLng(szText)
    If GetRecordCount(sSql) <= 0 Then
        tgcEditor.ErrorMessage(COL_PRFT) _
            = "Invalid Profit Center Number"
        Exit Function
    End If
            
    'Check duplication in table
    sSql = "SELECT prh_empno FROM pr_hours " _
        & " WHERE prh_empno = " & CLng(txtEmployeeNumber.Text) _
        & " AND prh_prft_ctr = " & CInt(szText) _
        & " AND prh_date = " & tfnSQLString(tfnDateString(sDate)) _
        & " AND prh_pay_code = " & tfnSQLString(sCode) _
        & " AND prh_pay_type = " & tfnSQLString(sType) _

     'only for add_mode, since in edit mode everyrecord is on grid, so valify grid is enough
    If GetRecordCount(sSql) > 0 And t_nFormMode = ADD_MODE Then
        tgcEditor.ErrorMessage(COL_PRFT) _
            = "Date/Profit Center/Pay Code already exists for this Employee"
        Exit Function
    End If
    
    If Not subCheckTotalPrft Then
        
        tgcEditor.ErrorMessage(COL_PRFT) = "Can't allow more than 5 different profit centers! "
        fnValidProfitCtr = False
        Exit Function
    End If
    fnValidProfitCtr = True
    
    Exit Function

err:
    tfnErrHandler "fnvalidprofitctr", sSql
    
End Function

Private Function fnValidPayCode(szText As String) As Boolean
    
    fnValidPayCode = False
    If Trim(szText) = "" Then
        tgcEditor.ErrorMessage(COL_CODE) _
            = "You must enter a Pay Code"
        Exit Function
    End If
    
    On Error GoTo err
    
    Dim sSql As String
    Dim lRow As String
    
    lRow = tgcEditor.GetCurrentRowNumber
    
    sSql = "SELECT prpa_pay_code FROM pr_pay " _
        & " WHERE prpa_pay_code = " & tfnSQLString(szText)
    
    If GetRecordCount(sSql) > 0 Then
        'fnValidPayCode = True
    Else
        tgcEditor.ErrorMessage(COL_CODE) _
            = "Pay Code is not exist"
        Exit Function
    End If
    
    'check in table pr_emp_pay
    Dim nYear As Integer
    nYear = Year(CDate(tblTimeCard.Columns(COL_DATE).Value))
    If nYear <= 0 Then
        Exit Function
    End If
    sSql = "SELECT distinct prep_pay_code FROM pr_emp_pay,pr_pay where prep_year = "
    sSql = sSql & tfnSQLString(nYear) & " and (prep_active is NULL or prep_active <> 'N') and prpa_pay_code=prep_pay_code"
    sSql = sSql & " and ((prpa_type='N'  and prpa_calc_method='D') or "
    sSql = sSql & " (prpa_type='P'  and prpa_calc_method='H')) "
    sSql = sSql & " and prep_empno=" & txtEmployeeNumber.Text
    If GetRecordCount(sSql) <= 0 Then
        tgcEditor.ErrorMessage(COL_CODE) _
            = "This Pay Code is invalid "
        Exit Function
    End If
'''
    If Not IsDate(tgcEditor.CellValue(COL_DATE, lRow)) Then
        tblTimeCard.col = COL_DATE
        Exit Function
    End If
    
    'do not check unique in add mode at paycode
    'will check at pay-type
    If t_nFormMode = ADD_MODE Then
        'fnValidPayCode = True  '***
        'Exit Function
    End If
    
    'col_date is not empty
    Dim sDate As Date
    Dim nPrft As Integer
    Dim sType As String
    
    sDate = tfnFormatDate(tgcEditor.CellValue(COL_DATE, lRow))
    nPrft = CInt(tgcEditor.CellValue(COL_PRFT, lRow))
    sType = fnGetField(tgcEditor.CellValue(COL_TYPE, lRow))
    
    Dim rsTemp As Recordset
                    
    'Check duplication in table
    sSql = "SELECT prh_empno FROM pr_hours " _
        & " WHERE prh_empno = " & CLng(txtEmployeeNumber.Text) _
        & " AND prh_prft_ctr = " & nPrft _
        & " AND prh_date = " & tfnSQLString(tfnDateString(sDate)) _
        & " AND prh_pay_code = " & tfnSQLString(szText) _
        & " AND prh_pay_type = " & tfnSQLString(sType) _

    'only for add_mode, since in edit mode everyrecord is on grid, so valify grid is enough
    If GetRecordCount(sSql) > 0 And t_nFormMode = ADD_MODE Then
        tgcEditor.ErrorMessage(COL_CODE) _
            = "Date/Profit Center/Pay Code already exists for this Employee"
        Exit Function
    End If

    ' Check duplication in the grid
    Dim vData() As Variant
    Dim lRowCount As Long
    
    Dim vData1() As Variant
    Dim lRowCount1 As Long
    
    Dim vData2() As Variant
    Dim lRowCount2 As Long
    
    Dim vData3() As Variant
    Dim lRowCount3 As Long
    
    Dim k As Long

    tgcEditor.GetColumn vData, lRowCount, COL_DATE
    tgcEditor.GetColumn vData1, lRowCount1, COL_PRFT
    tgcEditor.GetColumn vData2, lRowCount2, COL_CODE
    tgcEditor.GetColumn vData3, lRowCount3, COL_TYPE

    If lRowCount1 < lRowCount Then
        lRowCount = lRowCount1
    End If
    If lRowCount2 < lRowCount Then
        lRowCount = lRowCount2
    End If
    If lRowCount3 < lRowCount Then
        lRowCount = lRowCount3
    End If

    If lRowCount > 1 Then
        For k = 0 To lRowCount - 1
            If k <> lRow Then
                If fnGetField(vData1(k)) = fnGetField(nPrft) _
                    And tfnFormatDate(vData(k)) = sDate _
                    And fnGetField(vData2(k)) = szText _
                    And fnGetField(vData3(k)) = sType Then
                    tgcEditor.ErrorMessage(COL_CODE) _
                        = "Date/Profit Center/Pay Code already exists for this Employee"
                    Exit Function
                End If
            End If
        Next
    End If
        
'''
    fnValidPayCode = True
    Exit Function

err:
    tfnErrHandler "fnvalidpaycode", sSql
            
End Function

Private Sub subResetFloatingSQL2()

    Dim nYear As Integer
    Dim curDate As Date
    Dim curPrft As Integer
    
    Dim lRow As Long
    
    lRow = tgcEditor.GetCurrentRowNumber
    
    If Not IsNumeric(txtEmployeeNumber.Text) Then
        Exit Sub
    End If
    If Not IsDate(tblTimeCard.Columns(COL_DATE).Value) Then
        tblTimeCard.col = COL_DATE
        Exit Sub
    End If
    
    curDate = CDate(tblTimeCard.Columns(COL_DATE).Value)
    curPrft = tfnRound(tblTimeCard.Columns(COL_PRFT).Value)
    
    On Error GoTo err
    
    nYear = Year(CDate(tblTimeCard.Columns(COL_DATE).Value))
    If nYear <= 0 Then
        Exit Sub
    End If
    
    Dim sSql As String
    sSql = "SELECT distinct prep_pay_code ,prpa_desc  FROM pr_emp_pay,pr_pay where prep_year = "
    sSql = sSql & tfnSQLString(nYear) & " and (prep_active is NULL or prep_active <> 'N') and prpa_pay_code=prep_pay_code"
    sSql = sSql & " and ((prpa_type='N'  and prpa_calc_method='D') or "
    sSql = sSql & " (prpa_type='P'  and prpa_calc_method='H')) "
    sSql = sSql & " and prep_empno=" & txtEmployeeNumber.Text
    
    tgcFEditor.SetSQL COL_CODE, sSql
    
    Exit Sub
    
err:
    tfnErrHandler "subresetfloatingsql2", sSql

End Sub

Private Function fnPartSql() As String
    
    Dim sSql As String
    Dim sTable As String
    
    fnPartSql = ""
    sTable = "grid_record"
    
    On Error Resume Next
    sSql = "DROP TABLE " & sTable
    If bTableExists Then
        fnExecuteSQL sSql
        bTableExists = False
    End If
    
    On Error GoTo err
    sSql = "CREATE TEMP TABLE " & sTable _
        & "(tdate char(10), tprft char(6), tcode char(4), ttype char(2))"
    fnExecuteSQL sSql
    bTableExists = True
    
    If tgcEditor.RowCount < 2 Then
        Exit Function
    End If
    
    Dim k As Long
    Dim lRow As Long
    Dim gDate As String
    Dim gPrft As String
    Dim gCode As String
    Dim gType As String
    Dim curDate As String
    Dim curPrft As String
    
    lRow = tgcEditor.GetCurrentRowNumber
    curDate = fnGetField(tgcEditor.CellValue(COL_DATE, lRow))
    curPrft = fnGetField(tgcEditor.CellValue(COL_PRFT, lRow))
    
    For k = 0 To tgcEditor.RowCount - 1
        If k <> lRow Then
            gDate = fnGetField(tgcEditor.CellValue(COL_DATE, k))
            gPrft = fnGetField(tgcEditor.CellValue(COL_PRFT, k))
            gCode = fnGetField(tgcEditor.CellValue(COL_CODE, k))
            gType = fnGetField(tgcEditor.CellValue(COL_TYPE, k))
            sSql = "INSERT INTO " & sTable _
                & " (tdate, tprft, tcode, ttype) " _
                & "VALUES(" & tfnSQLString(gDate) & ", " & tfnSQLString(gPrft) & " , " _
                & tfnSQLString(gCode) & ", " & tfnSQLString(gType) & ")"
            fnExecuteSQL sSql
        End If
    Next
    fnPartSql = " AND " & tfnSQLString(curDate) & " || '***' || " & tfnSQLString(curPrft) _
        & " || '***' || trim(pcode) || '***' || trim(ptype) " _
        & " NOT IN(SELECT trim(tdate) || '***' || trim(tprft) || '***' " _
        & " || trim(tcode) || '***' || trim(ttype) FROM " & sTable & ")"
    Exit Function
err:
    tfnErrHandler "fnparsql", sSql
End Function

Private Sub subSetFloatingDropDown()
    
    Dim sSQL1 As String
    Dim sSQL2 As String
    
    Set tgcFEditor = New clsFloatingDropDown
    
    With tgcFEditor
         Set .SearchButton = cmdFloatingBtn
         Set .MainTable = tblTimeCard
         Set .DropDownTable = tblFloating
         Set .DataBase = t_dbMainDatabase
         Set .DataLink = datFloating
         Set .Form = m_Form
         Set .EditClass = tgcEditor
         
         .DefaultCursorControl = True
         
         .AddDropDown 1
         sSQL1 = " SELECT prft_ctr, prft_name FROM  sys_prft_ctr "
         .DropDownSQL(1) = sSQL1
         .DropDownType(1) = tgcFEditor.FactorSerachButton
         .AddColumn COL_PRFT
         .ColumnCaption(COL_PRFT) = "Profit Center"
         .ColumnType(COL_PRFT) = tgcFEditor.COLUMN_TYPE_INTEGER
         .ColumnDataField(COL_PRFT) = "prft_ctr"
         .AddExtraColumn "Name", "prft_name", 3150
         
         tgcFEditor.AddDropDown 2
                  
         .AddColumn COL_CODE
         .ColumnCaption(COL_CODE) = "Pay Code"
         .ColumnType(COL_CODE) = .COLUMN_TYPE_STRING
         .ColumnDataField(COL_CODE) = "prep_pay_code" ',prpa_desc "pcode"
         '***
         '.AddColumn COL_TYPE
         '.ColumnCaption(COL_TYPE) = "Pay Type"
         '.ColumnType(COL_TYPE) = .COLUMN_TYPE_STRING
         '.ColumnDataField(COL_TYPE) = "ptype"
         '***
         .AddExtraColumn "Pay Description", "prpa_desc", 2500
    
    End With
End Sub

Private Sub Class_Terminate()
    Set m_Form = Nothing
    Set tbToolbar = Nothing
    Set ffraStatusbar = Nothing
    Set txtEmployeeNumber = Nothing
    Set cmdEmployeeNumber = Nothing
    Set txtEmployeeName = Nothing
    Set cmdEmployeeName = Nothing
    Set txtSSN = Nothing
    Set cmdSSN = Nothing
    Set datComboDropDown = Nothing
    Set datFloating = Nothing
    Set tblComboDropdown = Nothing
    Set tblTimeCard = Nothing
    Set tblProfitCenter = Nothing
    Set tblFloating = Nothing
    Set cmdFloatingBtn = Nothing
    Set txtTotalDollars = Nothing
    Set txtTotal = Nothing
    Set cmdAddBtn = Nothing
    Set cmdEditBtn = Nothing
    Set cmdDeleteBtn = Nothing
    Set cmdUpdateInsertBtn = Nothing
    Set cmdRefreshSelectBtn = Nothing
    Set cmdExitCancelBtn = Nothing
    Set mnuCancel = Nothing
End Sub

Private Sub subTextSelected()

On Error Resume Next
    Dim txtBox As Control
    Set txtBox = m_Form.ActiveControl
    
    If TypeOf txtBox Is Textbox Then
        txtBox.SelStart = 0
        txtBox.SelLength = Len(Trim(txtBox.Text))
    End If
    
End Sub

' A generally used utility function, execute SQL and take care of errors
Private Function fnExecuteSQL(strSQL As String, _
                             Optional vCaller As Variant, _
                             Optional vMsg As Variant, _
                             Optional vDB As Variant) As Integer

    Dim objDB As DataBase
    
    If IsMissing(vDB) Then
        Set objDB = t_dbMainDatabase
    Else
        Set objDB = vDB
    End If
    
    On Error GoTo errExecute
    If objDB Is t_dbMainDatabase Then
        fnExecuteSQL = objDB.ExecuteSQL(strSQL)
    Else
        objDB.Execute strSQL
        fnExecuteSQL = 0
    End If

    On Error GoTo 0
    Exit Function

errExecute:
    Dim bShow As Boolean
    
    #If DEVELOP Then
        subShowODBCError vMsg, strSQL
    #Else
        bShow = Not IsMissing(vMsg)
        If IsMissing(vCaller) Then
            tfnErrHandler "fnExecuteSQL", strSQL, , bShow
        Else
            tfnErrHandler "fnExecuteSQL\vCaller", strSQL, , bShow
        End If
    #End If
    fnExecuteSQL = -1
    
End Function

Private Sub subEnableSearchButton(ByRef ctrlButton As FactorFrame, _
                                 ByVal bStatus As Boolean)

    ctrlButton.Style = 3  'command button
    ctrlButton.ShowFocusRect = True 'show a rectangular if focused
    ctrlButton.Enabled = bStatus
    If bStatus Then
        ctrlButton.Picture = frmContext.LoadPicture(SEARCH_UP)
    Else
        ctrlButton.Picture = frmContext.LoadPicture(SEARCH_DOWN)
    End If
    
End Sub


