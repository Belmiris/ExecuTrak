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
Private nsArray() As Variant
Private nsArrayCenter() As Variant
Private bNeverDelet As Boolean
Private sDecimalString As String

Private sTempPrMaster As String
Private sTempPrPay As String
Private nDataStatus As Integer
Private Const DATA_INI = 0
Private Const DATA_LOADED = 1
Private Const DATA_CHANGED = 2
Private Const COL_DATE = 0
Private Const COL_PRFT = 1
Private Const COL_CODE = 2
'Private Const COL_TYPE = 3
Private Const COL_HOUR = 3
'Private COL_SOURCE As Integer
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
   
    tgcEditor.AddEditColumn COL_DATE, "Enter Clock-in Date", szDatePattern
    tgcEditor.AddEditColumn COL_PRFT, "Enter Profit Center", szIntegerPattern
    tgcEditor.AddEditColumn COL_CODE, "Enter Pay Code", "^P{1,4}$"
    
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
    'tgcEditor.RemoveEditColumn COL_TYPE '0
    tgcEditor.RemoveEditColumn COL_DATE
    tgcEditor.RemoveEditColumn COL_PRFT
    tgcEditor.RemoveEditColumn COL_CODE
    
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
        
        If fnDeleteData() Then
            tgcEditor.DeleteRow
            fnGetTotalDollars
            fnFillProfitTable
            bNeverDelet = False
            tfnSetStatusBarCorrect ("The Record was Deleted Successfully.")
            
            If tgcEditor.RowCount < 1 Then
                tfnResetScreen
            End If
            
        End If
    
    End If
    
End Sub

Public Sub cmdUpdateInsertBtn_Click()
    tfnSetStatusBarMessage t_szEDIT
    
    If cValidate.FirstInvalidInput > 0 Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    t_bUpdateTable = fnInsertUpdate
    Screen.MousePointer = vbDefault
    
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
        Screen.MousePointer = vbHourglass
        bDone = fnFillTimeCard()
        Screen.MousePointer = vbDefault
    Else
        Exit Sub
    End If
    
    If bDone Then
        nDataStatus = DATA_LOADED
        subEnableRefresh False
        cmdUpdateInsertBtn.Enabled = False
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

Public Sub tblFloating_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    tgcFEditor.MouseUp y
End Sub

Public Sub tblFloating_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    tgcFEditor.RowColChange tblFloating
End Sub

Public Sub tblProfitCenter_AfterColEdit(ByVal ColIndex As Integer)
    tgcEditorPrft.AfterColEdit ColIndex
End Sub

Public Sub tblProfitCenter_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
     tgcEditorPrft.BeforeColEdit ColIndex, KeyAscii, Cancel
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

Public Sub tblProfitCenter_SelChange(Cancel As Integer)
    Cancel = True
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
        tgcEditor.CellValue(0, iCurrentRow) = tfnFormatDate(tgcEditor.CellValue(0, iCurrentRow))
    End If
    
    If ColIndex = 2 Then
        On Error Resume Next
        tgcEditor.CellValue(2, iCurrentRow) = UCase(tgcEditor.CellValue(2, iCurrentRow))
    End If
    
    fnGetTotalDollars
    
    If t_nFormMode = EDIT_MODE Then
        
        If iCurrentRow = tgcEditor.RowCount - 1 Then
            fnFillProfitTable
        End If
        
    End If
    
    subSetFormStatus
 
End Sub

Public Sub tblTimeCard_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    tgcEditor.BeforeColEdit ColIndex, KeyAscii, Cancel
    cmdUpdateInsertBtn.Enabled = False
End Sub

Public Sub tblTimeCard_Change()
    tgcEditor.Change
    
    If tblTimeCard.col = COL_DATE Then
        tblTimeCard.Columns(COL_PRFT).Value = ""
        tgcEditor.CellValue(COL_PRFT, tgcEditor.GetCurrentRowNumber) = ""
        tblTimeCard.Columns(COL_CODE).Value = ""
        tgcEditor.CellValue(COL_CODE, tgcEditor.GetCurrentRowNumber) = ""
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
        
        If tgcFEditor.ValidSelection Then
            tblTimeCard_AfterColEdit COL_PRFT
            tblTimeCard_AfterColEdit COL_CODE
        End If
        
    End If
    
'    If t_nFormMode = EDIT_MODE Then
'        tgcFEditor.GotFocus
'
'        If tgcFEditor.ValidSelection Then
'            tblTimeCard_AfterColEdit COL_CODE
'        End If
'
'    End If
'
End Sub

Public Sub tblTimeCard_KeyDown(KeyCode As Integer, Shift As Integer)
    
    tgcFEditor.Keypress tblTimeCard, KeyCode
      
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
    'fnGetTotalDollars
    'subSetFormStatus
End Sub

Public Sub tblTimeCard_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
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
    
    If LastCol = 3 Or LastCol = 1 Then
        fnFillProfitTable
    End If
    
    fnGetTotalDollars
    'subSetFormStatus
End Sub

Public Sub tblTimeCard_Scroll(Cancel As Integer)
    tgcFEditor.Scroll
End Sub

Public Sub tblTimeCard_SelChange(Cancel As Integer)
    Cancel = True
End Sub

Public Sub tblTimeCard_UnboundReadData(ByVal RowBuf As RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
    tgcEditor.ReadData RowBuf, StartLocation, ReadPriorRows
End Sub

Public Sub txtEmployeeName_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    frmContext.MouseDown Button, EMP_MST_UP
End Sub

Public Sub txtEmployeeNumber_Change()
    
    txtEmployeeName.Text = ""
    txtSSN.Text = ""
    tgcDropdown.Change txtEmployeeNumber
    cValidate.Change txtEmployeeNumber
    'nDataStatus = DATA_CHANGED
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

Public Sub txtEmployeeNumber_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    frmContext.MouseDown Button, EMP_MST_UP
End Sub

Public Sub txtSSN_Change()
    
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

Public Sub txtSSN_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
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
                     fnValidCellValue = fnValidHours(sText)
'                     fnValidCellValue = True
'                Case 4
'                    fnValidCellValue = fnValidHours(sText)
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
        dValueofDollar = dValueofDollar + tfnRound(recordArray(COL_HOUR), 2)
    Next Index
        
    txtTotalDollars.Text = tfnFormatDecimal(dValueofDollar, 2)
    Exit Sub
errorDollar:
    Resume Next
End Sub

Private Sub subCheckEmpNum()

    'comment out if statement, we don't need validate here junsong 03/14/01
    'If cValidate.FirstInvalidInput < 0 Then
    subEnableEmployee False
    tblTimeCard.Enabled = True
    tblProfitCenter.Enabled = True
    
    If t_nFormMode = EDIT_MODE Then
        Screen.MousePointer = vbHourglass
        fnFillTimeCard
        nDataStatus = DATA_LOADED
        Screen.MousePointer = vbDefault
    End If
    
    'End If
    
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
                sTempInteger = tfnRound(nCenterName(1, inDex2), 2) + tfnRound(recordArray(COL_HOUR), 2)
                nCenterName(1, inDex2) = tfnFormatDecimal(sTempInteger, 2)
                CenterExsit = True
                Exit For
            End If
        
        Next inDex2
    
        If CenterExsit = False Then
              nCenterName(0, NumberCenter) = Trim(recordArray(1))
              nCenterName(1, NumberCenter) = tfnRound(recordArray(3), 2)
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

Private Function fnFillTimeCard() As Boolean
    Dim sSql As String
    Dim rsTemp As Recordset
    Dim Index As Long
    Dim n As Integer
    On Error GoTo errorFilltalble
    
    sSql = " SELECT bh_date,bh_prft_ctr,bh_pay_code,bh_hours FROM bonus_hold "
    sSql = sSql & " WHERE bh_empno= " & tfnSQLString(txtEmployeeNumber)
    sSql = sSql & " AND bh_chk_link = 0 "
    sSql = sSql & " AND bh_pay_code IN (SELECT DISTINCT pcode FROM " & sTempPrPay
    sSql = sSql & " WHERE pyear = Year(bh_date) "
    sSql = sSql & " AND pempno= " & tfnSQLString(txtEmployeeNumber) & ")"
    sSql = sSql & " ORDER BY bh_date,bh_prft_ctr ,bh_pay_code  "
    
    On Error GoTo errH
    
    If tfnLockRow("bonus_hold", "bh_hours", sSql) Then
        tgcEditor.FillWithSQL t_dbMainDatabase, sSql, , True
    End If
    
    fnFillTimeCard = True
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
        .SQL = "SELECT empno, empname, empssn FROM " & sTempPrMaster
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
        'COL_SOURCE = tgcEditor.AddHiddenField("prh_source")
        .DisplayFormat(COL_PRFT) = "#####"
        .DisplayFormat(COL_HOUR) = "###,###,##0.00"
        
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
    sTempPrMaster = "temp_pr_master" 'tai
    sTempPrPay = "temp_pr_pay" 'tai
    bWarned = False
    bTableExists = False
End Sub

Private Function fnDeleteData() As Boolean
    Dim sSql As String
    Dim rsTemp As Recordset
    Dim CurentRow As Integer
    
    CurentRow = tgcEditor.GetCurrentRowNumber
    
    On Error GoTo ErrorDeleteData
    
    sSql = "DELETE FROM bonus_hold  "
    sSql = sSql & " WHERE bh_empno= " & tfnSQLString(txtEmployeeNumber)
    sSql = sSql & " AND bh_date= " & tfnDateString(tgcEditor.CellValue(0, CurentRow), True)
    sSql = sSql & " AND bh_prft_ctr = " & tfnRound(tgcEditor.CellValue(1, CurentRow))
    sSql = sSql & " AND bh_pay_code = " & tfnSQLString(tgcEditor.CellValue(2, CurentRow))

    On Error GoTo ErrorDeleteData
    
    t_dbMainDatabase.ExecuteSQL sSql
    fnDeleteData = True
    Exit Function
    
ErrorDeleteData:
    tfnErrHandler "fnDeleteData", sSql
    fnDeleteData = False
End Function

Private Function fnInsertUpdate() As Boolean
    
    If Not fnCheckGridDuplicate Then
        Exit Function
    End If
    
    If t_nFormMode = ADD_MODE Then
        fnInsertUpdate = fnInsert
    End If
    
    If t_nFormMode = EDIT_MODE Then
        fnInsertUpdate = fnUpdate
    End If
    
End Function

Private Function fnCheckGridDuplicate() As Boolean
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
                    fnCheckGridDuplicate = False
                    cmdUpdateInsertBtn.Enabled = False
                    Exit Function
                End If
        
        Next index1
    
    Next Index
    
    fnCheckGridDuplicate = True
        
    
End Function

Private Function fnUpdate() As Boolean
    Dim Index As Long
    Dim sSql As String
    Dim rsTemp As Recordset
    Dim toTalRowNumber As Integer
  
    toTalRowNumber = tgcEditor.RowCount
    
    If toTalRowNumber <= 0 Then
        fnUpdate = False
        Exit Function
    End If
    
    For Index = 0 To toTalRowNumber - 1
        
        sSql = "UPDATE bonus_hold SET bh_hours = " & tfnSQLString(tgcEditor.CellValue(COL_HOUR, Index))
        sSql = sSql & " WHERE bh_empno = " & tfnRound(txtEmployeeNumber)
        sSql = sSql & " AND bh_date = " & tfnDateString(tgcEditor.CellValue(COL_DATE, Index), True)
        sSql = sSql & " AND bh_prft_ctr = " & tfnRound(tgcEditor.CellValue(COL_PRFT, Index))
        sSql = sSql & " AND bh_pay_code = " & tfnSQLString(tgcEditor.CellValue(COL_CODE, Index))
        
         On Error GoTo errorUpdate
         t_dbMainDatabase.ExecuteSQL sSql
    Next
    
    fnUpdate = True
    Exit Function
    
errorUpdate:
    tfnErrHandler "fnUpdate", sSql
    fnUpdate = False
End Function

Private Function fnInsert() As Boolean
    Dim Index As Integer
    Dim sSql As String
    Dim rsTemp As Recordset
    Dim recordArray() As Variant
    Dim recordIndex As Integer
    Dim toTalRowNumber As Integer
    
    On Error GoTo errorInsert
    toTalRowNumber = tgcEditor.RowCount - 1
    
    For Index = 0 To toTalRowNumber
        tgcEditor.GetRow recordArray(), recordIndex, Index
        
        sSql = " INSERT INTO bonus_hold(bh_empno, bh_prft_ctr, bh_pay_code, "
        sSql = sSql & "bh_hours, bh_date, bh_check_amount, bh_override, bh_chk_link) VALUES ("
        sSql = sSql & tfnSQLString(txtEmployeeNumber) & "," & tfnRound(recordArray(COL_PRFT)) & ","
        sSql = sSql & tfnSQLString(recordArray(COL_CODE)) & "," & tfnSQLString(recordArray(COL_HOUR)) & ","
        sSql = sSql & tfnDateString(recordArray(COL_DATE), True) & ",0, 'Y', 0)"
        
        t_dbMainDatabase.ExecuteSQL sSql
    Next Index
    
    fnInsert = True
    Exit Function

errorInsert:
    tfnErrHandler "fnInsert", sSql
    fnInsert = False
                 
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
    
    nDataStatus = DATA_INI
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
    
    fnEnableControls False
    DoEvents
    tblTimeCard.Enabled = False
    
    DoEvents
    
    cmdAddBtn.Enabled = True
    cmdEditBtn.Enabled = True
    subSetFocus cmdAddBtn
End Sub

Private Sub subCreateTempTable()
    
    Dim sSql As String
    
    #If PROTOTYPE Then
        Exit Sub
    #End If
    
    On Error GoTo errHanlder
        
    sSql = "SELECT DISTINCT prm_empno AS empno, trim(prm_last_name) || ', ' || trim(prm_first_name) " _
       & "AS empname, prm_ssn AS empssn FROM pr_master,pr_security " _
       & "WHERE prs_code>=prm_security_code AND prs_name=" & tfnSQLString(tfnGetUserName) _
       & " AND prm_date_termed IS NULL " _
       & " INTO TEMP " & sTempPrMaster
     
    fnExecuteSQL sSql
                        
    sSql = "SELECT prpa_pay_code AS pcode, " _
        & "prpa_desc AS pdesc, prep_year As pyear, prep_empno AS pempno " _
        & "FROM pr_pay, pr_emp_pay WHERE  prep_pay_code = prpa_pay_code AND prep_active = 'Y'" _
        & " AND(prpa_type = 'P' AND prpa_calc_method = 'H' OR (prpa_type = 'N' AND prpa_calc_method = 'D'))" _
        & " INTO TEMP " & sTempPrPay
        
    fnExecuteSQL sSql
                        
    Exit Sub

errHanlder:
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
    
    sSql = "SELECT empno FROM " & sTempPrMaster _
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
        sSql = "SELECT bh_empno FROM bonus_hold WHERE bh_chk_link = 0 AND bh_empno = " & CLng(txtBox.Text)
    Else
        sSql = "SELECT pempno FROM " & sTempPrPay & " WHERE pempno = " & CLng(txtBox.Text)
    End If
                
    If GetRecordCount(sSql) > 0 Then
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
    
    sSql = "SELECT empno, empname, empssn FROM " & sTempPrMaster _
        & " WHERE empssn = " & tfnSQLString(txtBox.Text)
        
    If GetRecordSet(rsTemp, sSql) <= 0 Then
        Exit Function
    End If
    
    If t_nFormMode = EDIT_MODE Then
        sSql = "SELECT pempno FROM bonus_hold, " & sTempPrPay & " WHERE pempno = " & rsTemp!empno
        sSql = sSql & " AND pempno = bh_empno AND bh_chk_link = 0"
    Else
        sSql = "SELECT pempno FROM " & sTempPrPay & " WHERE pempno = " & rsTemp!empno
    End If
                
    If GetRecordCount(sSql) > 0 Then
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
    
    If t_nFormMode = EDIT_MODE Then
        sSql = "select distinct empno, empname, empssn FROM " & sTempPrMaster & ",bonus_hold "
        sSql = sSql & " WHERE empno = bh_empno and bh_chk_link = 0 "
        sSql = sSql & " AND bh_pay_code IN (SELECT DISTINCT pcode FROM " & sTempPrPay & ")"
    Else
        sSql = "select distinct empno, empname, empssn FROM " & sTempPrMaster & "," & sTempPrPay
        sSql = sSql & " WHERE empno = pempno"
    End If

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
    
    If t_nFormMode = EDIT_MODE And nDataStatus = DATA_INI Then
        cmdDeleteBtn.Enabled = False
        subEnableRefresh False
        cmdUpdateInsertBtn.Enabled = False
    End If
                                
End Sub

Private Function fnValidDate(ByVal szText As String, _
                                  ByVal lRow As Long) As Boolean
            
    Dim sSql As String
    Dim rsTemp As Recordset
    
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
    
    'david 02/22/2001
    If fnValidDate Then
        
        If fnGetField(tgcEditor.CellValue(COL_PRFT, lRow)) = "" Then
            'get default profit center for the employee
            sSql = "SELECT prm_prft_ctr1"
            sSql = sSql & " FROM pr_master"
            sSql = sSql & " WHERE prm_empno = " & tfnRound(txtEmployeeNumber)
            
            If GetRecordSet(rsTemp, sSql) > 0 Then
                tgcEditor.CellValue(COL_PRFT, lRow) = tfnRound(rsTemp!prm_prft_ctr1)
            End If
            
        End If
        
    End If
    
    Exit Function

err:
    tfnErrHandler "fnValidDate", sSql
    
End Function

Private Function fnValidProfitCtr(ByVal szText As String, _
                                  ByVal lRow As Long) As Boolean
            
    Dim sDate As Date
    Dim sCode As String
    Dim sSql As String
    Dim rsTemp As Recordset
                
    fnValidProfitCtr = False
    
    If Trim(szText) = "" Then
        tgcEditor.ErrorMessage(COL_PRFT) = "You must enter a Profit Center Number"
        Exit Function
    End If
            
    If Not IsDate(tgcEditor.CellValue(COL_DATE, lRow)) Then
        tblTimeCard.col = COL_DATE
        Exit Function
    End If
    
    On Error GoTo errHandler
    
    sSql = " SELECT prft_ctr, prft_name FROM sys_prft_ctr, pr_master "
    sSql = sSql & " WHERE prm_empno = " & tfnRound(txtEmployeeNumber)
    sSql = sSql & " AND prft_ctr = " & tfnRound(szText)
    sSql = sSql & " AND (prft_ctr = prm_prft_ctr1 OR prft_ctr = prm_prft_ctr2 "
    sSql = sSql & " OR prft_ctr = prm_prft_ctr3 OR prft_ctr = prm_prft_ctr4 "
    sSql = sSql & " OR prft_ctr = prm_prft_ctr5)"
        
    If GetRecordCount(sSql) <= 0 Then
        tgcEditor.ErrorMessage(COL_PRFT) = "Invalid Profit Center Number"
        Exit Function
    End If
            
    sDate = tfnFormatDate(tgcEditor.CellValue(COL_DATE, lRow))
    sCode = fnGetField(tgcEditor.CellValue(COL_CODE, lRow))
            
    sSql = "SELECT bh_empno FROM bonus_hold " _
        & " WHERE bh_empno = " & tfnRound(txtEmployeeNumber.Text) _
        & " AND bh_prft_ctr = " & tfnRound(szText) _
        & " AND bh_date = " & tfnSQLString(tfnDateString(sDate)) _
        & " AND bh_pay_code = " & tfnSQLString(sCode)

    If GetRecordCount(sSql) > 0 And t_nFormMode = ADD_MODE Then
        tgcEditor.ErrorMessage(COL_PRFT) = "Date/Profit Center/Pay Code already exists for this Employee"
        Exit Function
    End If
    
    fnValidProfitCtr = True
    
    Exit Function

errHandler:
    tfnErrHandler "fnValidProfitCtr", sSql
    
End Function

Private Function fnValidPayCode(szText As String) As Boolean
    Dim sSql As String
    Dim lRow As Long
    Dim nYear As Integer
    Dim sDate As String
    Dim nPrft As Integer
    Dim rsTemp As Recordset
    
    fnValidPayCode = False
    
    If Trim(szText) = "" Then
        tgcEditor.ErrorMessage(COL_CODE) = "You must enter a Pay Code"
        Exit Function
    End If
    
    On Error GoTo errHandler
    
    lRow = tgcEditor.GetCurrentRowNumber
    nYear = Year(CDate(tblTimeCard.Columns(COL_DATE).Value))
    
    If nYear <= 0 Then
        Exit Function
    End If
    
    sSql = " SELECT pcode FROM " & sTempPrPay
    sSql = sSql & " WHERE pcode = " & tfnSQLString(szText)
    sSql = sSql & " AND pempno = " & tfnRound(txtEmployeeNumber)
    sSql = sSql & " AND pyear = " & tfnSQLString(nYear)
    
    If GetRecordCount(sSql) <= 0 Then
        tgcEditor.ErrorMessage(COL_CODE) = "Pay Code does not exist"
        Exit Function
    End If
    
    sDate = tfnFormatDate(tgcEditor.CellValue(COL_DATE, lRow))
    nPrft = tfnRound(tgcEditor.CellValue(COL_PRFT, lRow))
                    
    sSql = "SELECT bh_empno FROM bonus_hold " _
        & " WHERE bh_empno = " & tfnRound(txtEmployeeNumber.Text) _
        & " AND bh_prft_ctr = " & nPrft _
        & " AND bh_date = " & tfnDateString(sDate, True) _
        & " AND bh_pay_code = " & tfnSQLString(szText)

    If GetRecordCount(sSql) > 0 And t_nFormMode = ADD_MODE Then
        tgcEditor.ErrorMessage(COL_CODE) = "Date/Profit Center/Pay Code already exists for this Employee"
        Exit Function
    End If

    If fnCheckDuplicate(sDate, nPrft, szText, lRow) Then
        tgcEditor.ErrorMessage(COL_CODE) = "Date/Profit Center/Pay Code already entered"
        Exit Function
    End If
    
    fnValidPayCode = True
    Exit Function

errHandler:
    tfnErrHandler "fnValidPayCode", sSql
            
End Function

Private Function fnCheckDuplicate(sDate As String, nPrft As Integer, sPayCode As String, lRow As Long) As Boolean
    ' Check duplication in the grid
    Dim vData() As Variant
    Dim lRowCount As Long
    Dim vData1() As Variant
    Dim lRowCount1 As Long
    Dim vData2() As Variant
    Dim lRowCount2 As Long
    Dim k As Long
    
    tgcEditor.GetColumn vData, lRowCount, COL_DATE
    tgcEditor.GetColumn vData1, lRowCount1, COL_PRFT
    tgcEditor.GetColumn vData2, lRowCount2, COL_CODE

    If lRowCount1 < lRowCount Then
        lRowCount = lRowCount1
    End If
    
    If lRowCount2 < lRowCount Then
        lRowCount = lRowCount2
    End If
    
    If lRowCount > 1 Then
        
        For k = 0 To lRowCount - 1
            
            If k <> lRow Then
                
                If fnGetField(vData1(k)) = fnGetField(nPrft) _
                   And tfnFormatDate(vData(k)) = sDate _
                   And fnGetField(vData2(k)) = sPayCode Then
                   
                     fnCheckDuplicate = True
                End If
                
            End If
            
        Next
        
    End If

End Function

Private Sub subResetFloatingSQL2()

    Dim nYear As Integer
    Dim sSql As String
    Dim lRow As Long
    
    
    If Not IsNumeric(txtEmployeeNumber.Text) Then
        Exit Sub
    End If
    
    If Not IsDate(tblTimeCard.Columns(COL_DATE).Value) Then
        tblTimeCard.col = COL_DATE
        Exit Sub
    End If
    
    On Error GoTo errHandler
    
    'sql from profit center
    sSql = "SELECT prft_ctr, prft_name FROM sys_prft_ctr, pr_master "
    sSql = sSql & " WHERE prm_empno = " & tfnRound(txtEmployeeNumber)
    sSql = sSql & " AND (prft_ctr = prm_prft_ctr1 OR prft_ctr = prm_prft_ctr2 "
    sSql = sSql & " OR prft_ctr = prm_prft_ctr3 OR prft_ctr = prm_prft_ctr4 "
    sSql = sSql & " OR prft_ctr = prm_prft_ctr5)"
    
    tgcFEditor.SetSQL COL_PRFT, sSql
    
    nYear = Year(CDate(tblTimeCard.Columns(COL_DATE).Value))
    
    If nYear <= 0 Then
        Exit Sub
    End If
    
    'sql for pay code

    sSql = "SELECT pcode, pdesc FROM " & sTempPrPay
    sSql = sSql & " WHERE pyear = " & tfnSQLString(nYear)
    sSql = sSql & " AND pempno = " & tfnRound(txtEmployeeNumber)
    tgcFEditor.SetSQL COL_CODE, sSql
    Exit Sub
    
errHandler:
    tfnErrHandler "subresetfloatingsql2", sSql

End Sub

Private Sub subSetFloatingDropDown()
    Dim sSql As String
    
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
         'sSql = " SELECT prft_ctr, prft_name FROM  sys_prft_ctr "
         '.DropDownSQL(1) = sSql
         '.DropDownType(1) = tgcFEditor.FactorSerachButton
         .AddColumn COL_PRFT
         .ColumnCaption(COL_PRFT) = "Profit Center"
         .ColumnType(COL_PRFT) = tgcFEditor.COLUMN_TYPE_INTEGER
         .ColumnDataField(COL_PRFT) = "prft_ctr"
         .AddExtraColumn "Name", "prft_name", 2500
         
         tgcFEditor.AddDropDown 2
                  
         .AddColumn COL_CODE
         .ColumnCaption(COL_CODE) = "Pay Code"
         .ColumnType(COL_CODE) = .COLUMN_TYPE_STRING
         .ColumnDataField(COL_CODE) = "pcode" ',prpa_desc "pcode"
         .AddExtraColumn "Pay Description", "pdesc", 1850
    
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


