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
Private txtEHPrftCtr As Textbox
Private cmdEHPrftCtr As FactorFrame
Private txtEHPrftName As Textbox
Private cmdEHPrftName As FactorFrame
Private txtEHDate As Textbox
Private cmdEHDate As FactorFrame
Private datComboDropDown As Data
Private datFloating As Data
Private tblComboDropdown As TDBGrid
Private tblTimeCard As TDBGrid
Private tblFloating As TDBGrid
Private cmdFloatingBtn As PictureBox
Private cmdAddBtn  As FactorFrame
Private cmdEditBtn  As FactorFrame
Private cmdDeleteBtn  As FactorFrame
Private cmdUpdateInsertBtn As FactorFrame
Private cmdRefreshSelectBtn As FactorFrame
Private cmdExitCancelBtn As FactorFrame
Private cmdExitBtn As FactorFrame
Private efraBridge As FactorFrame
'The control for the pay code selection
Private efraEHPayCode As FactorFrame
Private txtEHPayCode As Textbox
Private txtEHPayCodeDesc As Textbox
Private cmdEHPayCode As FactorFrame
Private cmdEHPayCodeDesc As FactorFrame
''''''''
'Vijaya on 12/12/03 Magic#421627
Private txtEHMgrHours As Textbox
Private bMgrHoursValid As Boolean
''''''''''''''''''''''''''''''''
Private lblEHTotalHours As Label
Private lblEHTotalPayCode1 As Label
Private lblEHTotalPayCode2 As Label
Private lblEHTotalPayCode3 As Label

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
Private tgcFEditor As clsFloatingDropDown
Private bNeverDelet As Boolean
Private sDecimalString As String

Private sTempPrMaster As String
'david 03/11/2003  #403275
'make it public so that can be accessed by clsOTProcessor
Public sTempPrPay As String
''''''''''''''''''''''''''
Private nDataStatus As Integer
Private Const DATA_INI = 0
Private Const DATA_LOADED = 1
Private Const DATA_CHANGED = 2
Private Const COL_EMPNO = 0
Private Const COL_EMPNAME = 1
Private Const COL_HOUR = 2
Private Const COL_PAYCODE1 = 3
Private Const COL_PAYCODE2 = 4
Private Const COL_PAYCODE3 = 5
Private Const MAXCols As Integer = 5
Private Const sDefaultPAYCODE1 As String = "Paycode 1"
Private Const sDefaultPAYCODE2 As String = "Paycode 2"
Private Const sDefaultPAYCODE3 As String = "Paycode 3"
Private Const sHOURLYPAYCODE As String = "7HRP"
Private g_sOrigPayCode1 As String
Private g_sOrigPayCode2 As String
Private g_sOrigPayCode3 As String
Private bWarned As Boolean 'for warning of Date with glp_status ='W'

Private bTableExists As Boolean
Private g_lOriginalRowCount As Long

'david 12/18/2003  #421627
Private lMgrEmpNbr As Long
''''''''''''''''''''''''''
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

Property Set EHPrftCtrTextBox(txtPrftCtr As Textbox)
    Set txtEHPrftCtr = txtPrftCtr
End Property

Property Set EHPrftCtrButton(cmdPrftCtr As FactorFrame)
    Set cmdEHPrftCtr = cmdPrftCtr
End Property

Property Set EHPrftNameTextBox(txtPrftName As Textbox)
    Set txtEHPrftName = txtPrftName
End Property

Property Set EHPayCodeTextBox(txtBox As Textbox)
    Set txtEHPayCode = txtBox
End Property

Property Set EHPayCodeDescTextBox(txtBox As Textbox)
    Set txtEHPayCodeDesc = txtBox
End Property

Property Set EHPayCodeFactorFrame(efraFrame As FactorFrame)
    Set efraEHPayCode = efraFrame
End Property

Property Set EHPayCodeButton(cmdButton As FactorFrame)
    Set cmdEHPayCode = cmdButton
End Property

Property Set EHPayCodeDescButton(cmdButton As FactorFrame)
    Set cmdEHPayCodeDesc = cmdButton
End Property

Property Set EHPrftNameButton(cmdPrftName As FactorFrame)
    Set cmdEHPrftName = cmdPrftName
End Property

Property Set EHDATETextBox(txtDate As Textbox)
    Set txtEHDate = txtDate
End Property

Property Set EHDATEButton(cmdButton As FactorFrame)
    Set cmdEHDate = cmdButton
End Property

'Vijaya on 12/12/03 Magic#421627
Property Set EHMGRHOURSTextBox(txtMgrHours As Textbox)
    Set txtEHMgrHours = txtMgrHours
End Property
''''''''''''''''''''''''''''''''

Property Set EHTotalHoursLabel(lbl As Label)
    Set lblEHTotalHours = lbl
End Property

Property Set EHTotalPayCode1(lbl As Label)
    Set lblEHTotalPayCode1 = lbl
End Property

Property Set EHTotalPayCode2(lbl As Label)
    Set lblEHTotalPayCode2 = lbl
End Property

Property Set EHTotalPayCode3(lbl As Label)
    Set lblEHTotalPayCode3 = lbl
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

Property Set FloatingGrid(tblGrid As TDBGrid)
    Set tblFloating = tblGrid
End Property

Property Set FloatingButton(Pic As PictureBox)
    Set cmdFloatingBtn = Pic
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
    subResetFloatingSQL
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

Public Sub efraBridge_LostFocus()
    cValidate.LostFocus efraBridge
End Sub

Public Sub cmdAddBtn_GotFocus()
    tfnSetStatusBarMessage t_szADD
    On Error Resume Next
    DoEvents
    tgcEditor.ClearData
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
    subSetEditColumns ADD_MODE
    tblTimeCard.Enabled = False
    'Vijaya on 12/12/03 Magic#421627
    If txtEHMgrHours.Visible Then
        txtEHMgrHours.Enabled = False
    End If
    ''''''''''''''''''''''''''''''''
    txtEHPrftCtr.SetFocus
    
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
    subSetEditColumns EDIT_MODE
    tblTimeCard.Enabled = False
    'Vijaya on 12/12/03 Magic#421627
    If txtEHMgrHours.Visible Then
        txtEHMgrHours.Enabled = False
    End If
    ''''''''''''''''''''''''''''''''
    txtEHPrftCtr.SetFocus
End Sub

Public Sub cmdDeleteBtn_Click()
    Dim bResponse As Integer
    Dim sStyle As String, sMsg As String
    Dim lCurRow As Long
    
    sMsg = "Are you sure you want to delete this record?"
    sStyle = vbYesNo + vbQuestion + vbDefaultButton2
    bResponse = MsgBox(sMsg, sStyle)
    
    If bResponse = vbYes Then
        lCurRow = tgcEditor.GetCurrentRowNumber
        
        If fnDeleteData(lCurRow) Then
            
            If lCurRow <= g_lOriginalRowCount Then
                g_lOriginalRowCount = g_lOriginalRowCount - 1
            End If
            
            tgcEditor.DeleteRow
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

'====
'Grid
'====
Public Sub tblTimeCard_AfterColEdit(ByVal ColIndex As Integer)
    tgcEditor.AfterColEdit ColIndex
    nDataStatus = DATA_CHANGED
    subCalculateTotal
    subSetFormStatus
End Sub

Public Sub tblTimeCard_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    tgcEditor.BeforeColEdit ColIndex, KeyAscii, Cancel
    
    If ColIndex <= COL_EMPNAME Then
            
        If tgcEditor.GetCurrentRowNumber < g_lOriginalRowCount Then
            Cancel = True
            Exit Sub
        End If
    
    Else
        
        Select Case ColIndex
            Case COL_PAYCODE1
            
                If tblTimeCard.Columns(ColIndex).Caption = sDefaultPAYCODE1 Then
                    Cancel = True
                End If
                
            Case COL_PAYCODE2
                
                If tblTimeCard.Columns(ColIndex).Caption = sDefaultPAYCODE2 Then
                    Cancel = True
                End If
                
            Case COL_PAYCODE3
                
                If tblTimeCard.Columns(ColIndex).Caption = sDefaultPAYCODE3 Then
                    Cancel = True
                End If
                
        End Select
        
    End If

End Sub

Public Sub tblTimeCard_Change()
    tgcEditor.Change
    cmdUpdateInsertBtn.Enabled = False
    
    If tblTimeCard.Col = COL_EMPNO Then
        tgcEditor.CellValue(COL_EMPNAME, tgcEditor.GetCurrentRowNumber) = ""
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
            nDataStatus = DATA_CHANGED
            subSetFormStatus
        End If
        
    End If
    
End Sub

Public Sub tblTimeCard_KeyDown(KeyCode As Integer, Shift As Integer)
    '7-11 Required the quick key entered, so the following code are used
    'to hanlde this situation, but it may cause the validtion not work
    'correctly.
    
    Dim bKeyReturnEntered As Boolean
    
    If tblTimeCard.Col = COL_HOUR Then

        If KeyCode = vbKeyReturn Then

            If t_nFormMode = ADD_MODE Then

                If tgcEditor.GetCurrentRowNumber < g_lOriginalRowCount Then
                    bKeyReturnEntered = True
                    KeyCode = 0
                End If

            Else

                If tgcEditor.GetCurrentRowNumber < tgcEditor.RowCount - 1 Then
                    bKeyReturnEntered = True
                    KeyCode = 0
                End If

            End If

        End If

    End If

    tgcFEditor.Keypress tblTimeCard, KeyCode
      
    If tgcFEditor.ValidSelection Then
        nDataStatus = DATA_CHANGED
    End If
    
    tgcEditor.KeyDown KeyCode, Shift
    
    If tblTimeCard.Col = COL_HOUR Then

        If bKeyReturnEntered Then
            
            If t_nFormMode = ADD_MODE And tgcEditor.GetCurrentRowNumber = 0 And g_lOriginalRowCount > 0 Then
                'force rowcolchange event when in the first row to reset the last column in class tgceditor
                tgcEditor.RowColChange 0, COL_PAYCODE1
            End If
            
            If t_nFormMode = ADD_MODE Then
                SendKeys "{right}"
                DoEvents
                SendKeys "{left}"
            End If
            
            SendKeys "{down}"
        End If

    End If
    
End Sub

Public Sub tblTimeCard_KeyPress(KeyAscii As Integer)
    
    If Not tgcEditor.Keypress(KeyAscii) Then
        KeyAscii = 0
    End If
        
    tgcFEditor.Keypress tblTimeCard, KeyAscii

End Sub

Public Sub tblTimeCard_LostFocus()
    tgcEditor.LostFocus
    tgcFEditor.LostFocus tblTimeCard
End Sub

Public Sub tblTimeCard_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If tblTimeCard.Col = 1 And Button = RIGHT_BUTTON Then
        frmContext.MouseDown Button, EMP_MST_UP
    End If
    
End Sub

Public Sub tblTimeCard_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim nCurCol As Integer
    Dim lCurRow As Long
    
    If t_bStartupFlag Then
        Exit Sub
    End If
    
    lCurRow = tgcEditor.GetCurrentRowNumber
    
    If tgcFEditor.RowColChange(tblTimeCard) Then
        tgcEditor.RowColChange LastRow, LastCol
    End If
    
    If t_nFormMode = ADD_MODE Then
    
        If tblTimeCard.Col <= COL_EMPNAME Then
                    
            If lCurRow <= g_lOriginalRowCount Then
                tgcEditor.GotoNextColumn COL_EMPNO
            End If
        
        End If
        
        nCurCol = tblTimeCard.Col
        
        If tgcEditor.ValidCell(COL_EMPNO, lCurRow) Then
            
            If nCurCol = COL_EMPNAME Then
                tgcEditor.GotoNextColumn LastCol
            End If
        
        End If
        
    End If
    
    
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

Public Sub txtEHPrftName_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    frmContext.MouseDown Button, PRFTCNTR_UP
End Sub

Public Sub txtEHPrftCtr_Change()
    tgcDropdown.Change txtEHPrftCtr
    cValidate.Change txtEHPrftCtr
End Sub

Public Sub txtEHPrftCtr_Click()
    tgcDropdown.Click txtEHPrftCtr
End Sub

Public Sub txtEHPrftCtr_GotFocus()
    tgcDropdown.GotFocus txtEHPrftCtr
    cValidate.GotFocus txtEHPrftCtr
    Screen.MousePointer = vbDefault
    
    If tgcDropdown.SingleRecordSelected Then
        subSetFocus txtEHDate
    End If
    
End Sub

Public Sub txtEHPrftCtr_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        Screen.MousePointer = vbHourglass
    End If

    If tgcDropdown.Keypress(txtEHPrftCtr, KeyAscii) Then
        cValidate.Keypress txtEHPrftCtr, KeyAscii
    Else
    
        If KeyAscii = vbKeyReturn Then
        
            If tgcDropdown.SingleRecordSelected Then
                subSetFocus txtEHDate
            End If
        
        End If
        
        KeyAscii = 0
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Public Sub txtEHPrftCtr_LostFocus()
   
    If cValidate.LostFocus(txtEHPrftCtr, cmdEHPrftCtr, txtEHPrftName, cmdEHPrftName, tblComboDropdown) Then
        tgcDropdown.LostFocus txtEHPrftCtr
    End If

End Sub

Public Sub cmdEHPrftCtr_Click()
    Screen.MousePointer = vbHourglass
    tgcDropdown.Click cmdEHPrftCtr
    Screen.MousePointer = vbDefault
End Sub

Public Sub txtEHPrftName_Change()
    tgcDropdown.Change txtEHPrftName
    cValidate.Change txtEHPrftName
End Sub

Public Sub txtEHPrftName_Click()
    tgcDropdown.Click txtEHPrftName
End Sub

Public Sub txtEHPrftName_GotFocus()
    tgcDropdown.GotFocus txtEHPrftName
    cValidate.GotFocus txtEHPrftName
    Screen.MousePointer = vbDefault
    
    If tgcDropdown.SingleRecordSelected Then
        subSetFocus txtEHDate
    End If
    
End Sub

Public Sub txtEHPrftName_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        Screen.MousePointer = vbHourglass
    End If

    If tgcDropdown.Keypress(txtEHPrftName, KeyAscii) Then
        cValidate.Keypress txtEHPrftName, KeyAscii
    Else
    
        If KeyAscii = vbKeyReturn Then
            
            If tgcDropdown.SingleRecordSelected Then
                subSetFocus txtEHDate
            End If
        
        End If
        
        KeyAscii = 0
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Public Sub txtEHPrftName_LostFocus()
    
    If cValidate.LostFocus(txtEHPrftCtr, cmdEHPrftCtr, txtEHPrftName, cmdEHPrftName, tblComboDropdown) Then
        tgcDropdown.LostFocus txtEHPrftName
    End If

End Sub

Public Sub cmdEHPrftName_Click()
    Screen.MousePointer = vbHourglass
    tgcDropdown.Click cmdEHPrftName
    Screen.MousePointer = vbDefault
End Sub

Public Sub cmdEHDate_Click()
    Screen.MousePointer = vbHourglass
    subReSetComboSql
    tgcDropdown.Click cmdEHDate
    Screen.MousePointer = vbDefault
End Sub

Public Sub txtEHPrftCtr_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    frmContext.MouseDown Button, EMP_MST_UP
End Sub

Public Sub txtEHDate_Change()
    tgcDropdown.Change txtEHDate
    cValidate.Change txtEHDate
    tfnRegExpControlChange txtEHDate, szDatePattern
End Sub

Public Sub txtEHDate_GotFocus()
    'subTextSelected
    tgcDropdown.GotFocus txtEHDate
    cValidate.GotFocus txtEHDate
    
    If tgcDropdown.SingleRecordSelected Then
  
        If tblTimeCard.Enabled Then
            subSetFocus tblTimeCard
        Else
            subSetFocus efraBridge
        End If
    
    End If
    
End Sub

Public Sub txtEHDate_KeyPress(KeyAscii As Integer)

    tfnRegExpControlKeyPress txtEHDate, KeyAscii, szDatePattern

    If t_nFormMode = EDIT_MODE Then
        
        If KeyAscii = vbKeyReturn Then
            subReSetComboSql
        End If
        
        If tgcDropdown.Keypress(txtEHDate, KeyAscii) Then
            cValidate.Keypress txtEHDate, KeyAscii
        Else
        
            If KeyAscii = vbKeyReturn Then
                
                If tgcDropdown.SingleRecordSelected Then
                      
                    If tblTimeCard.Enabled Then
                        subSetFocus tblTimeCard
                    Else
                        subSetFocus efraBridge
                    End If
                
                End If
            
            End If
            
            KeyAscii = 0
        End If
        
    Else
    
        If KeyAscii = vbKeyReturn Then
            
            If tblTimeCard.Enabled Then
                subSetFocus tblTimeCard
            Else
                subSetFocus efraBridge
            End If
            
            KeyAscii = 0
        Else
            cValidate.Keypress txtEHDate, KeyAscii
        End If
        
    End If
End Sub

Public Sub txtEHDate_LostFocus()

    If cValidate.LostFocus(txtEHDate, cmdEHDate, tblComboDropdown) Then
        tgcDropdown.LostFocus txtEHDate
        
        If cValidate.ValidInput(txtEHDate) Then
            txtEHDate = tfnFormatDate(txtEHDate)
            
            If fnFillTimeCard() Then
                subEnableProfitCtr False
                tblTimeCard.Enabled = True
                'Vijaya on 12/12/03 Magic#421627
                txtEHMgrHours.Enabled = tblTimeCard.Enabled
                ''''''''''''''''''''''''''''''''
                subSetFocus tblTimeCard
            End If
            
        End If
        
    End If
    
End Sub

Public Sub txtEHPayCode_Change()
    tgcDropdown.Change txtEHPayCode
    cValidate.Change txtEHPayCode
End Sub

Public Sub txtEHPayCode_GotFocus()
    
    tgcDropdown.GotFocus txtEHPayCode
    cValidate.GotFocus txtEHPayCode
    Screen.MousePointer = vbDefault
    
    If tgcDropdown.SingleRecordSelected Then
        subSetFocus tblTimeCard
    End If
    
End Sub

Public Sub txtEHPayCode_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        Screen.MousePointer = vbHourglass
    End If

    If tgcDropdown.Keypress(txtEHPayCode, KeyAscii) Then
        cValidate.Keypress txtEHPayCode, KeyAscii
    Else
    
        If KeyAscii = vbKeyReturn Then
            
            If tgcDropdown.SingleRecordSelected Then
                subSetFocus tblTimeCard
            End If
        
        End If
        
        KeyAscii = 0
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Public Sub txtEHPayCode_LostFocus()
    
    If cValidate.LostFocus(txtEHPayCode, cmdEHPayCode, txtEHPayCodeDesc, cmdEHPayCode, tblComboDropdown) Then
        tgcDropdown.LostFocus txtEHPayCode
        
        If cValidate.ValidInput(txtEHPayCode) Then
            subSetPaycodeToHeader
        End If
        
    End If

End Sub

Public Sub cmdEHPayCode_Click()
    Screen.MousePointer = vbHourglass
    tgcDropdown.Click cmdEHPayCode
    Screen.MousePointer = vbDefault
End Sub

Public Sub txtEHPayCodeDesc_Change()
    tgcDropdown.Change txtEHPayCodeDesc
    cValidate.Change txtEHPayCodeDesc
End Sub

Public Sub txtEHPayCodeDesc_GotFocus()
    Set cValidate.LastBox = txtEHPayCodeDesc
    tgcDropdown.GotFocus txtEHPayCodeDesc
    cValidate.GotFocus txtEHPayCodeDesc
    Screen.MousePointer = vbDefault
    
    If tgcDropdown.SingleRecordSelected Then
        subSetFocus tblTimeCard
    End If
    
End Sub

Private Sub subSetPaycodeToHeader()
    Dim nCol As Integer
    
    If m_Form.ActiveControl Is cmdEHPayCode Or m_Form.ActiveControl Is cmdEHPayCodeDesc Then
        Exit Sub
    End If
    
    nCol = tfnRound(txtEHPayCode.Tag)
    
    If nCol > COL_HOUR And Trim(txtEHPayCode) <> "" Then
        tblTimeCard.Columns(nCol).Caption = txtEHPayCode
        txtEHPayCode = ""
    End If
    
    efraEHPayCode.ZOrder 1
    efraEHPayCode.Visible = False
    subSetFocus tblTimeCard
    
    If txtEHDate.Tag <> "" Then
        tblTimeCard.Col = txtEHDate.Tag
    End If
    
End Sub

Public Sub tblTimeCard_HeadClick(ByVal ColIndex As Integer)
    
    If fnCheckHeaderOK(ColIndex) Then
        efraEHPayCode.Visible = True
        efraEHPayCode.ZOrder 0
        txtEHPayCode.Tag = ColIndex
        subSetFocus txtEHPayCode
    End If
    
End Sub

Public Sub txtEHPayCodeDesc_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        Screen.MousePointer = vbHourglass
    End If

    If tgcDropdown.Keypress(txtEHPayCodeDesc, KeyAscii) Then
        cValidate.Keypress txtEHPayCodeDesc, KeyAscii
    Else
    
        If KeyAscii = vbKeyReturn Then
            
            If tgcDropdown.SingleRecordSelected Then
                subSetFocus tblTimeCard
            End If
        
        End If
        
        KeyAscii = 0
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Public Sub txtEHPayCodeDesc_LostFocus()
    
    If cValidate.LostFocus(txtEHPayCode, cmdEHPayCode, txtEHPayCodeDesc, cmdEHPayCode, tblComboDropdown) Then
        tgcDropdown.LostFocus txtEHPayCodeDesc
        
        If cValidate.ValidInput(txtEHPayCode) Then
            subSetPaycodeToHeader
        End If
        
    End If

End Sub

Public Sub cmdEHPayCodeDesc_Click()
    Screen.MousePointer = vbHourglass
    tgcDropdown.Click cmdEHPayCodeDesc
    Screen.MousePointer = vbDefault
End Sub

Public Function fnValidCellValue(tgTable As TDBGrid, ByVal nCol As Integer, ByVal nRow As Long, sText As String) As Boolean
    
    Select Case tgTable.TabIndex
        Case tblTimeCard.TabIndex
            Select Case nCol
                Case COL_EMPNO
                    fnValidCellValue = fnValidEmpNo(sText, nCol, nRow)
                Case COL_EMPNAME
                    'Make sure we have valid Employee Number
                    If tgcEditor.ValidCell(COL_EMPNO, nRow) Then
                        fnValidCellValue = True
                    Else
                        fnValidCellValue = False
                    End If
                Case COL_HOUR, COL_PAYCODE1, COL_PAYCODE2, COL_PAYCODE3
                    fnValidCellValue = fnValidHours(sText, nCol, nRow)
            End Select
    End Select
End Function

Public Function fnInValidData(txtBox As Textbox) As Boolean
    
    Select Case txtBox.TabIndex
        Case txtEHPrftCtr.TabIndex
            fnInValidData = Not fnValidProfitCtr(txtBox)
        Case txtEHPrftName.TabIndex
            fnInValidData = False
        Case txtEHDate.TabIndex
            fnInValidData = Not fnValidDate(txtBox)
        Case txtEHPayCode.TabIndex
            fnInValidData = Not fnValidPayCode(txtBox)
        Case txtEHPayCodeDesc.TabIndex
            fnInValidData = False
        'Vijaya on 12/12/03 Magic#421627
        Case txtEHMgrHours.TabIndex
            fnInValidData = Not fnValidMgrHours(txtBox)
        ''''''''''''''''''''''''''''''''
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
Private Function fnValidHours(szText As String, nCol As Integer, lRow As Long) As Boolean
    Dim strSQL As String
    Dim sPayCode As String
    
    If Trim(szText) = "" Or tfnRound(szText, 2) = 0# Then
        fnValidHours = True
        Exit Function
    End If

    If Not tgcEditor.ValidCell(COL_EMPNO, lRow) Then
        fnValidHours = False
        tgcEditor.ErrorMessage(nCol) = "Select Valid Employee first"
        Exit Function
    End If
    
    Select Case nCol
        Case COL_HOUR
            sPayCode = sHOURLYPAYCODE

        Case COL_PAYCODE1
            
            If tblTimeCard.Columns(nCol).Caption = sDefaultPAYCODE1 Then
                tgcEditor.ErrorMessage(nCol) = "No Pay Code, can't enter data here"
                fnValidHours = False
                Exit Function
            Else
                sPayCode = tblTimeCard.Columns(nCol).Caption
            End If
            
        Case COL_PAYCODE2
            
            If tblTimeCard.Columns(nCol).Caption = sDefaultPAYCODE2 Then
                tgcEditor.ErrorMessage(nCol) = "No Pay Code, can't enter data here"
                fnValidHours = False
                Exit Function
            Else
                sPayCode = tblTimeCard.Columns(nCol).Caption
            End If
            
        Case COL_PAYCODE3
            
            If tblTimeCard.Columns(nCol).Caption = sDefaultPAYCODE3 Then
                tgcEditor.ErrorMessage(nCol) = "No Pay Code, can't enter data here"
                fnValidHours = False
                Exit Function
            Else
                sPayCode = tblTimeCard.Columns(nCol).Caption
            End If
            
        Case Else
            fnValidHours = True
            Exit Function
    End Select
    
    If sPayCode <> "" Then
        strSQL = "SELECT  * FROM " & sTempPrPay
        strSQL = strSQL & " WHERE pcode = " & tfnSQLString(sPayCode)
        strSQL = strSQL & " AND pempno = " & tfnRound(tgcEditor.CellValue(COL_EMPNO, lRow))
        'add by junsong 01/21/2003, call 398327
        strSQL = strSQL & " AND pYear = " & Year(CDate(txtEHDate))
    Else
        tgcEditor.ErrorMessage(nCol) = "No Pay Code, can't enter data here"
        fnValidHours = False
        Exit Function
    End If
    
    If GetRecordCount(strSQL) <= 0 Then
        tgcEditor.ErrorMessage(nCol) = "Pay Code " & sPayCode & " was not set for this employee, can't enter data here"
        fnValidHours = False
    Else
        'check the prft ctr/date/emploee/code
        strSQL = "SELECT * FROM bonus_hold WHERE bh_prft_ctr = " & tfnRound(txtEHPrftCtr)
        strSQL = strSQL & " AND bh_date = " & tfnDateString(txtEHDate, True)
        strSQL = strSQL & " AND bh_pay_code = " & tfnSQLString(sPayCode)
        strSQL = strSQL & " AND bh_empno = " & tfnRound(tgcEditor.CellValue(COL_EMPNO, lRow))
    
        If GetRecordCount(strSQL) > 0 And t_nFormMode = ADD_MODE Then
            tgcEditor.ErrorMessage(nCol) = "Date/Prft Ctr/Pay Code already exists for this employee, can't enter data."
        Else
            fnValidHours = True
        End If
        
    End If

End Function

Private Function fnFillTimeCard() As Boolean
    Dim strSQL As String
    Dim strLock As String
    Dim rsTemp As Recordset
    Dim sLockUser As String
    Dim aryData() As Variant
    Dim lOldEmpNo As Long
    Dim sPayCode1 As String
    Dim sPayCode2 As String
    Dim sPayCode3 As String
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    If m_Form.ActiveControl Is cmdExitBtn Or m_Form.ActiveControl Is cmdExitCancelBtn Then
        Exit Function
    End If
    
    If cValidate.FirstInvalidInput >= 0 Then
        MsgBox "Please select valid profit center and clock-in date", vbExclamation
        Exit Function
    End If
    
    g_lOriginalRowCount = -1
    tfnSetStatusBarCorrect "Loading data, Please wait..."
    m_Form.MousePointer = vbHourglass
    
    If t_nFormMode = ADD_MODE Then
        strSQL = "SELECT empno, empname, jobclass FROM " & sTempPrMaster
        strSQL = strSQL & " WHERE prft_ctr1 = " & tfnRound(txtEHPrftCtr)
        strSQL = strSQL & " AND pay_type = 'H'"
        strSQL = strSQL & " ORDER BY jobclass, empno"
    Else
        strSQL = " SELECT DISTINCT jobclass, empno, empname, bh_pay_code, bh_hours FROM bonus_hold,  " & sTempPrMaster & "," & sTempPrPay
        strSQL = strSQL & " WHERE bh_prft_ctr = " & tfnRound(txtEHPrftCtr)
        'strSQL = strSQL & " AND (bh_prft_Ctr = prft_ctr1 OR bh_prft_ctr = prft_ctr2  OR bh_prft_ctr = prft_ctr3 "
        'strSQL = strSQL & " OR bh_prft_Ctr = prft_ctr4 OR bh_prft_ctr = prft_ctr5)"
        strSQL = strSQL & " AND bh_chk_link = 0 "
        strSQL = strSQL & " AND bh_empno = empno "
        strSQL = strSQL & " AND bh_date = " & tfnDateString(txtEHDate, True)
        strSQL = strSQL & " AND bh_pay_code = pcode "
        strSQL = strSQL & " AND pyear = Year(bh_date) "
        strSQL = strSQL & " AND pempno= empno"
        'strSQL = strSQL & " AND bh_pay_code IN (SELECT DISTINCT pcode FROM " & sTempPrPay
        'strSQL = strSQL & " WHERE pyear = Year(bh_date) "
        'strSQL = strSQL & " AND pempno= empno)"
        strSQL = strSQL & " ORDER BY jobclass, empno"
    End If
        
    strLock = "bh_prft_ctr = " & tfnRound(txtEHPrftCtr) & " AND bh_date = " & tfnDateString(txtEHDate, True)
    
    If Not tfnLockRow("bonus_hold", "bh_hours", strLock, , sLockUser) Then
        MsgBox "The data was locked by user " & sLockUser, vbExclamation
        m_Form.MousePointer = vbDefault
        Exit Function
    End If
        
    If GetRecordSet(rsTemp, strSQL, nDB_REMOTE, "fnFillTimeCard", True) <= 0 Then
        
        If t_nFormMode = ADD_MODE Then
            'STILL DOESN'T WORK
            'This follwoing code work with the code in the tfnresetScreen, see comment there
            tblTimeCard.Col = COL_EMPNO
            GoTo SKIPLoadData
        Else
            MsgBox "NO record selected.", vbInformation
            m_Form.MousePointer = vbDefault
            Exit Function
        End If
        
    End If
        
    ReDim aryData(MAXCols, rsTemp.RecordCount - 1)
    lOldEmpNo = -1
    i = -1
    
    Do Until rsTemp.EOF
        
        If t_nFormMode = ADD_MODE Then
            i = i + 1
            aryData(COL_EMPNO, i) = fnGetField(rsTemp!EmpNO)
            aryData(COL_EMPNAME, i) = fnGetField(rsTemp!empName)
            aryData(COL_HOUR, i) = ""
            aryData(COL_PAYCODE1, i) = ""
            aryData(COL_PAYCODE2, i) = ""
            aryData(COL_PAYCODE3, i) = ""
        Else
            'In edit mode only four possible pay code
            If lOldEmpNo <> tfnRound(rsTemp!EmpNO) Then
                i = i + 1
                aryData(COL_EMPNO, i) = fnGetField(rsTemp!EmpNO)
                aryData(COL_EMPNAME, i) = fnGetField(rsTemp!empName)
                lOldEmpNo = tfnRound(rsTemp!EmpNO)
            End If
            
            If fnGetField(rsTemp!bh_Hours) <> "" And fnGetField(rsTemp!bh_pay_code) <> "" Then
            
                If UCase(fnGetField(rsTemp!bh_pay_code)) = sHOURLYPAYCODE Then
                    aryData(COL_HOUR, i) = tfnRound(rsTemp!bh_Hours, 2)
                Else
                
                    If UCase(fnGetField(rsTemp!bh_pay_code)) = sPayCode1 Then
                        aryData(COL_PAYCODE1, i) = tfnRound(rsTemp!bh_Hours, 2)
                    ElseIf UCase(fnGetField(rsTemp!bh_pay_code)) = sPayCode2 Then
                        aryData(COL_PAYCODE2, i) = tfnRound(rsTemp!bh_Hours, 2)
                    ElseIf UCase(fnGetField(rsTemp!bh_pay_code)) = sPayCode3 Then
                        aryData(COL_PAYCODE3, i) = tfnRound(rsTemp!bh_Hours, 2)
                    Else
                        
                        If sPayCode1 = "" Then
                            aryData(COL_PAYCODE1, i) = tfnRound(rsTemp!bh_Hours, 2)
                            sPayCode1 = UCase(fnGetField(rsTemp!bh_pay_code))
                            tblTimeCard.Columns(COL_PAYCODE1).Caption = fnGetField(rsTemp!bh_pay_code)
                        ElseIf sPayCode2 = "" Then
                            aryData(COL_PAYCODE2, i) = tfnRound(rsTemp!bh_Hours, 2)
                            sPayCode2 = UCase(fnGetField(rsTemp!bh_pay_code))
                            tblTimeCard.Columns(COL_PAYCODE2).Caption = fnGetField(rsTemp!bh_pay_code)
                        ElseIf sPayCode3 = "" Then
                            aryData(COL_PAYCODE3, i) = tfnRound(rsTemp!bh_Hours, 2)
                            sPayCode3 = UCase(fnGetField(rsTemp!bh_pay_code))
                            tblTimeCard.Columns(COL_PAYCODE3).Caption = fnGetField(rsTemp!bh_pay_code)
                        End If
                    
                    End If
                    
                End If
                
            End If
            
        End If
        
        rsTemp.MoveNext
    Loop
    
    If i >= 0 Then
        ReDim Preserve aryData(MAXCols, i)
        tgcEditor.FillWithArray aryData, True, True
        
        If sPayCode1 <> "" Then
            g_sOrigPayCode1 = sPayCode1
        End If
        
        If sPayCode2 <> "" Then
            g_sOrigPayCode2 = sPayCode2
        End If
        
        If sPayCode3 <> "" Then
            g_sOrigPayCode3 = sPayCode3
        End If
        
        'tempory fix the problem that the column was not to set to HOUR when the data was loaded
        If t_nFormMode = ADD_MODE Then
            tgcEditor.GotoNextColumn COL_EMPNO
        End If
        
        'Vijaya on 12/12/03 Magic#421627
        If txtEHMgrHours.Visible Then
            Dim sDate As String
            Dim sMsg As String
            strSQL = "SELECT zzsemh_empno, zzsemh_date, zzsemh_hours FROM zzse_manager_hour"
            strSQL = strSQL & " WHERE zzsemh_prft_ctr =" & tfnRound(txtEHPrftCtr)
            strSQL = strSQL & " AND zzsemh_date =" & tfnDateString(txtEHDate, True)
            If GetRecordSet(rsTemp, strSQL, , "fnFillTimeCard") > 0 Then
                txtEHMgrHours = Format(fnGetField(rsTemp!zzsemh_hours), "0.00")
                sDate = fnGetField(rsTemp!zzsemh_date)
                strSQL = "SELECT prm_date_termed FROM  pr_master"
                strSQL = strSQL & " WHERE prm_empno=" & tfnRound(fnGetField(rsTemp!zzsemh_empno))
                If GetRecordSet(rsTemp, strSQL, , "fnFillTimeCard") > 0 Then
                    If fnGetField(rsTemp!prm_date_termed) <> "" Then
                        If CDate(rsTemp!prm_date_termed) < CDate(sDate) Then
                            DoEvents
                            sMsg = "Existing Manager has been terminated prior to the reporting date."
                            sMsg = sMsg & " Re-enter manager's hour or enter zero to update"
                            MsgBox sMsg, vbExclamation
                        Else
                            bMgrHoursValid = True
                        End If
                    Else
                        bMgrHoursValid = True
                    End If
                End If
            End If
        End If
        ''''''''''''''''''''''''''''''''
        subCalculateTotal
        g_lOriginalRowCount = i
    End If
    
SKIPLoadData:
    
    fnFillTimeCard = True
    'tblTimeCard.SetFocus
    
    If tgcEditor.RowCount > 0 Then
        cmdDeleteBtn.Enabled = True
    End If
    
    m_Form.MousePointer = vbDefault
    Exit Function

ErrorHandler:
    tfnErrHandler "fnFill TimeCard", strSQL
    
End Function

Private Sub fnEnableControls(beOff As Boolean)
    txtEHPrftCtr.Enabled = beOff
    subEnableSearchbtn cmdEHPrftCtr, beOff
    txtEHPrftName.Enabled = beOff
    subEnableSearchbtn cmdEHPrftName, beOff
    txtEHDate.Enabled = beOff
    
    If t_nFormMode = ADD_MODE Then
        subEnableSearchButton cmdEHDate, False
    Else
        subEnableSearchButton cmdEHDate, beOff
    End If
    
    tblTimeCard.Enabled = beOff
    
    'Vijaya on 12/12/03 Magic#421627
    txtEHMgrHours.Enabled = beOff
    ''''''''''''''''''''''''''''''''
    
    subEnableSearchButton cmdEHPayCode, beOff
    subEnableSearchButton cmdEHPayCodeDesc, beOff
    efraEHPayCode.Visible = False
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
        .AddComboBox txtEHPrftCtr, cmdEHPrftCtr, "prft_ctr", tgcDropdown.SQL_INT_TYPE
        .AddComboBox txtEHPrftName, cmdEHPrftName, "prft_name", tgcDropdown.SQL_STRING_TYPE(40)
        
        .AddCombo
        .AddComboBox txtEHDate, cmdEHDate, "bh_date", tgcDropdown.SQL_DATE_TYPE
        
        .AddCombo
        .AddComboBox txtEHPayCode, cmdEHPayCode, "pcode", tgcDropdown.SQL_STRING_TYPE(4)
        .AddComboBox txtEHPayCodeDesc, cmdEHPayCodeDesc, "pdesc", tgcDropdown.SQL_STRING_TYPE(20)
    End With
    
End Sub

Private Sub subSetValidtion()
    Set cValidate = New cValidateInput
    With cValidate
    Set .StatusBar = ffraStatusbar
    Set .Form = m_Form
    .AddEditBox txtEHPrftCtr, "Enter the Payroll Profit Center."
    .AddEditBox txtEHPrftName, "Enter the Payroll Profit Center Name."
    .AddEditBox txtEHDate, "Enter the Clock-In Date."
    .AddEditBox txtEHPayCode, "Enter Pay Code"
    .AddEditBox txtEHPayCodeDesc, "Enter Pay Code Description"
    
    'Vijaya on 12/12/03 Magic#421627
    .AddEditBox txtEHMgrHours, "Enter Manager Hours"
    ''''''''''''''''''''''''''''''''
    
    '.ESCControl = tblComboDropdown
    .ESCControl = cmdUpdateInsertBtn
    .ESCControl = cmdExitCancelBtn
    .MinTabIndex = tbToolbar.TabIndex
    .MaxTabIndex = efraBridge.TabIndex + 1
    Set .ControlForFocus = efraBridge
    Set .LastBox = txtEHMgrHours        'txtEHPayCodeDesc
    .SetFirstControls cmdDeleteBtn, cmdUpdateInsertBtn
    End With
End Sub

Private Sub subSetGrid()
    Set tgcEditor = New clsTGSpreadSheet
    With tgcEditor
        Set .Table = tblTimeCard
        Set .Form = m_Form
        Set .StatusBar = ffraStatusbar
        sDecimalString = tfnDecimalPattern(8, 2)
        .AddEditColumn COL_EMPNO, "Enter Employee Number", szLongPattern
        .AddEditColumn COL_EMPNAME, "Enter Employee Name", "^P{1,60}$"
        .AddESCColumn COL_EMPNAME, COL_EMPNO
        .AddEditColumn COL_HOUR, "Enter Hours", sDecimalString
        .AddEditColumn COL_PAYCODE1, "Enter Hour", sDecimalString
        .AddEditColumn COL_PAYCODE2, "Enter Hour", sDecimalString
        .AddEditColumn COL_PAYCODE3, "Enter Hour", sDecimalString
        .DisplayFormat(COL_HOUR) = "###,###,##0.00"
        .DisplayFormat(COL_PAYCODE1) = "###,###,##0.00"
        .DisplayFormat(COL_PAYCODE2) = "###,###,##0.00"
        .DisplayFormat(COL_PAYCODE3) = "###,###,##0.00"
        
        .AddSortColumn COL_EMPNO, COL_EMPNO, .NUMERIC_TYPE, .ASCENDING, .CASE_INSENSITIVE
        .AddSortColumn COL_EMPNAME, COL_EMPNAME, .STRING_TYPE, .ASCENDING, .CASE_SENSITIVE
        .AllowAddNew = False
        .ClearData
    End With
    
End Sub

Private Sub subSetEditColumns(nFormMode As Integer)
    
    With tgcEditor
        
        Select Case nFormMode
            
            Case ADD_MODE
                .AddEditColumn COL_EMPNO, "Enter Employee Number", szLongPattern
                .AddEditColumn COL_EMPNAME, "Enter Employee Name", "^P{1,60}$"
                .AddEditColumn COL_HOUR, "Enter Hours", sDecimalString
                .AddEditColumn COL_PAYCODE1, "Enter Hour", sDecimalString
                .AddEditColumn COL_PAYCODE2, "Enter Hour", sDecimalString
                .AddEditColumn COL_PAYCODE3, "Enter Hour", sDecimalString
                .SortByColumn = False
                .AllowAddNew = True
            Case EDIT_MODE
                .RemoveEditColumn COL_EMPNO
                .RemoveEditColumn COL_EMPNAME
                .SortByColumn = True
                .AllowAddNew = False
        End Select
    
    End With
End Sub

Public Sub Form_Initialize() 'called before Form_Load
    
    t_bStartupFlag = True
    t_bDataChanged = False
    t_bUpdateTable = False
    
    t_nFormMode = IDLE_MODE
    
    ' ** change the help file for the application
    sTempPrMaster = "temp_pr_master"
    sTempPrPay = "temp_pr_pay"
    bWarned = False
    bTableExists = False
End Sub

Private Function fnDeleteData(lCurRow As Long) As Boolean
    Dim sSql As String
    Dim sPayCodeList As String
        
    sPayCodeList = fnGetPayCodeList(True)
    
    If t_nFormMode = ADD_MODE Then
        fnDeleteData = True
    Else
        sSql = "DELETE FROM bonus_hold  "
        sSql = sSql & " WHERE bh_prft_ctr = " & tfnRound(txtEHPrftCtr)
        sSql = sSql & " AND bh_date= " & tfnDateString(txtEHDate, True)
        sSql = sSql & " AND bh_empno = " & tfnRound(tgcEditor.CellValue(COL_EMPNO, lCurRow))
        sSql = sSql & " AND bh_pay_code IN " & sPayCodeList
    
        fnDeleteData = fnExecuteSQL(sSql, "fnDeleteData")
    End If
    
End Function

Private Function fnInsertUpdate() As Boolean
    
    If t_nFormMode = ADD_MODE Then
        fnInsertUpdate = fnInsert
    End If
    
    If t_nFormMode = EDIT_MODE And tgcEditor.GetChangedRowCount > 0 Then
        fnInsertUpdate = fnUpdate
    Else
        fnInsertUpdate = True
    End If
        
    'Vijaya on 12/12/03 Magic#421627
    If fnInsertUpdate Then
        fnInsertUpdate = fnInsertzzse_manager_hour
    End If
    ''''''''''''''''''''''''''''''''
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
    Dim lIndex As Long
    Dim strSQL As String
    Dim sPayCode As String
    Dim rsTemp As Recordset
    Dim vData() As Variant
    Dim nCount As Integer
    
    'Delete all data first
    For lIndex = 0 To tgcEditor.RowCount - 1
        
        If Not fnDeleteData(lIndex) Then
            Exit Function
        End If
        
    Next
    
    'Reinsert All data
    fnUpdate = fnInsert()
    
End Function

Private Function fnInsert() As Boolean
    Dim lIndex As Long
    Dim strSQL As String
    Dim strSQLHeader As String
    Dim sPayCode As String
    Dim rsTemp As Recordset
    Dim vData() As Variant
    Dim nCount As Integer
    
    For lIndex = 0 To tgcEditor.RowCount - 1
        tgcEditor.GetRow vData(), nCount, lIndex
        
        strSQLHeader = " INSERT INTO bonus_hold(bh_empno, bh_prft_ctr, bh_date, bh_check_amount,  "
        strSQLHeader = strSQLHeader & "bh_override, bh_chk_link, bh_pay_code, bh_hours) VALUES ("
        strSQLHeader = strSQLHeader & tfnRound(vData(COL_EMPNO)) & "," & tfnRound(txtEHPrftCtr) & ","
        strSQLHeader = strSQLHeader & tfnDateString(txtEHDate, True) & ",0, 'Y', 0,"

        If tfnRound(vData(COL_HOUR), 2) <> 0# Then
            strSQL = strSQLHeader & tfnSQLString(sHOURLYPAYCODE) & "," & tfnRound(vData(COL_HOUR), 2) & ")"
            
            If fnExecuteSQL(strSQL, "fnInsert") < 0 Then
                Exit Function
            End If
            
        End If
        
        If tfnRound(vData(COL_PAYCODE1), 2) <> 0# Then
            sPayCode = tblTimeCard.Columns(COL_PAYCODE1).Caption
            strSQL = strSQLHeader & tfnSQLString(sPayCode) & "," & tfnRound(vData(COL_PAYCODE1), 2) & ")"
            
            If fnExecuteSQL(strSQL, "fnInsert") < 0 Then
                Exit Function
            End If
            
        End If
        
        If tfnRound(vData(COL_PAYCODE2), 2) <> 0# Then
            sPayCode = tblTimeCard.Columns(COL_PAYCODE2).Caption
            strSQL = strSQLHeader & tfnSQLString(sPayCode) & "," & tfnRound(vData(COL_PAYCODE2), 2) & ")"
            
            If fnExecuteSQL(strSQL, "fnInsert") < 0 Then
                Exit Function
            End If
            
        End If
        
        If tfnRound(vData(COL_PAYCODE3), 2) <> 0# Then
            sPayCode = tblTimeCard.Columns(COL_PAYCODE3).Caption
            strSQL = strSQLHeader & tfnSQLString(sPayCode) & "," & tfnRound(vData(COL_PAYCODE3), 2) & ")"
            
            If fnExecuteSQL(strSQL, "fnInsert") < 0 Then
                Exit Function
            End If
            
        End If
        
    Next lIndex
    
    fnInsert = True
    Exit Function
                 
End Function

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
    txtEHPrftCtr.Text = ""
    txtEHPrftName.Text = ""
    txtEHDate.Text = ""
    
    'Vijaya on 12/12/03 Magic#421627
    txtEHMgrHours = ""
    bMgrHoursValid = False
    lMgrEmpNbr = -1
    ''''''''''''''''''''''''''''''''
    
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
    
    tblTimeCard.Columns(COL_PAYCODE1).Caption = sDefaultPAYCODE1
    tblTimeCard.Columns(COL_PAYCODE2).Caption = sDefaultPAYCODE2
    tblTimeCard.Columns(COL_PAYCODE3).Caption = sDefaultPAYCODE3
    tgcEditor.ClearData
    subCalculateTotal
    
    g_sOrigPayCode1 = sDefaultPAYCODE1
    g_sOrigPayCode2 = sDefaultPAYCODE2
    g_sOrigPayCode3 = sDefaultPAYCODE3
    bNeverDelet = True
    fnEnableControls False
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
                                
    sSql = "SELECT prm_empno AS empno, trim(prm_last_name) || ', ' || trim(prm_first_name) " _
       & "AS empname, prm_job_class AS jobclass, prm_prft_ctr1 AS prft_ctr1, prm_prft_ctr2 as prft_ctr2, " _
       & " prm_prft_ctr3 as prft_ctr3, prm_prft_ctr4 AS prft_ctr4, prm_prft_ctr5 AS prft_ctr5, prm_pay_type as pay_type " _
       & " FROM pr_master" _
       & " WHERE prm_date_termed IS NULL " _
       & " INTO TEMP " & sTempPrMaster
    
    fnExecuteSQL sSql
                                
    sSql = "SELECT prpa_pay_code AS pcode, prpa_desc AS pdesc, prep_year As pyear, prep_empno AS pempno " _
        & " FROM pr_pay, pr_emp_pay WHERE  prep_pay_code = prpa_pay_code AND prep_active = 'Y'" _
        & " AND(prpa_type = 'P' AND prpa_calc_method = 'H' OR (prpa_type = 'N' AND prpa_calc_method = 'D'))" _
        & " INTO TEMP " & sTempPrPay
        
    fnExecuteSQL sSql
                        
    Exit Sub

errHanlder:
    tfnErrHandler "subcreatetemptable", sSql
    
End Sub

Private Function fnValidEmpNo(sText As String, nCol As Integer, lRow As Long) As Boolean
    Dim strSQL As String
    Dim rsTemp As Recordset
'    Dim sEmpSecurityId As String
'    Dim SUserSecurityID As String
    Dim i As Long
    Dim sEmpName As String
    
    Dim sPrftCtrList As String
    Dim sTemp As String
    Dim nCountPrftCtr As Integer
    
    fnValidEmpNo = False
    
    If Trim(sText) = "" Then
        tgcEditor.ErrorMessage(nCol) = "You must enter an Employee Number"
        Exit Function
    End If
    
    If t_nFormMode = ADD_MODE Then
        strSQL = "SELECT empno, empName FROM " & sTempPrMaster & " WHERE empno = " & tfnRound(sText)
    Else
        strSQL = "SELECT bh_empno, empName FROM bonus_hold, " & sTempPrMaster
        strSQL = strSQL & " WHERE bh_empno = empno AND bh_chk_link = 0 AND bh_empno = " & tfnRound(sText)
    End If
    
    If GetRecordSet(rsTemp, strSQL, nDB_REMOTE, "fnValidEmpNO") <= 0 Then
        tgcEditor.ErrorMessage(nCol) = "Entered invalid Employee Number"
        Exit Function
    Else
        sEmpName = fnGetField(rsTemp!empName)
    End If
    
    strSQL = "SELECT * FROM pr_master WHERE prm_empno = " & tfnSQLString(sText)
    
    If GetRecordSet(rsTemp, strSQL, nDB_REMOTE, "fnValidEmpNO") > 0 Then
    
        If fnGetField(rsTemp!prm_date_termed) <> "" Then
            tgcEditor.ErrorMessage(nCol) = "This is a terminated employee! Can't modify the pay hours."
            Exit Function
        End If
    
'        sEmpSecurityId = fnGetField(rsTemp!prm_security_code)
'
'        strSQL = "select prs_code from pr_security where prs_name = " & tfnSQLString(tfnGetUserName)
'
'        If GetRecordSet(rsTemp, strSQL, nDB_REMOTE, "fnValidEmpNo") <= 0 Then
'            tgcEditor.ErrorMessage(nCol) = "The user doesn't have access to any employee records."
'            Exit Function
'        End If
'
'        SUserSecurityID = fnGetField(rsTemp!prs_code)
'
'        If SUserSecurityID < sEmpSecurityId Then
'            tgcEditor.ErrorMessage(nCol) = "The employee has higher security code, can't be edited."
'            Exit Function
'        End If
        
        'david 12/27/2002  #394605
        'need to also check the profit center in pr_master for the employee
        'the bh_prft_ctr + prm_prft_ctrX +
'        strSQL = "SELECT count(DISTINCT bh_prft_ctr) as prftCount FROM bonus_hold"
'        strSQL = strSQL & " WHERE bh_chk_link = 0 "
'        'don't care date
'        'strSQL = strSQL & " AND bh_date = " & tfnDateString(txtEHDate, True)
'        strSQL = strSQL & " AND bh_prft_ctr <> " & tfnRound(txtEHPrftCtr)
'        strSQL = strSQL & " AND bh_empno = " & tfnRound(sText)
'        strSQL = strSQL & " AND bh_pay_code = " & tfnSQLString(sHOURLYPAYCODE)
'
'        If tfnRound(rsTemp!prftCount) >= 5 Then
'            tgcEditor.ErrorMessage(nCol) = "Employee may only have hours for five separate profit centers in a single pay period."
'            Exit Function
'        End If

        'david 12/27/2002  #394605
        'need to check the profit center in pr_master for the employee
        nCountPrftCtr = 0
        sPrftCtrList = ""
        
        For i = 1 To 5
            sTemp = fnGetField(rsTemp.Fields("prm_prft_ctr" & i))
            
            If sTemp <> "" Then
                If sPrftCtrList <> "" Then
                    sPrftCtrList = sPrftCtrList + ","
                End If
                
                sPrftCtrList = sPrftCtrList + sTemp
                
                If sTemp <> txtEHPrftCtr Then
                    nCountPrftCtr = nCountPrftCtr + 1
                End If
            End If
        Next i

        strSQL = "SELECT DISTINCT bm_eligible_pc FROM bonus_master"
        strSQL = strSQL & " WHERE bm_empno = " & tfnRound(sText)
        strSQL = strSQL & " AND bm_eligible_pc <> " & tfnRound(txtEHPrftCtr)
        strSQL = strSQL & " AND bm_eligible_pc NOT IN (" + sPrftCtrList + ")"
        
        If GetRecordSet(rsTemp, strSQL, nDB_REMOTE, "fnValidEmpNo") < 0 Then
            tgcEditor.ErrorMessage(nCol) = "Database Access Error"
            Exit Function
        End If
        
        'add bm_eligible_pc into the list
        While Not rsTemp.EOF
            sTemp = fnGetField(rsTemp!bm_eligible_pc)
            
            If sTemp <> "" Then
                If sPrftCtrList <> "" Then
                    sPrftCtrList = sPrftCtrList + ","
                End If
                
                sPrftCtrList = sPrftCtrList + sTemp
                nCountPrftCtr = nCountPrftCtr + 1
            End If
            
            rsTemp.MoveNext
        Wend
        
        strSQL = "SELECT count(DISTINCT bh_prft_ctr) as prftCount FROM bonus_hold"
        strSQL = strSQL & " WHERE bh_chk_link = 0 "
        strSQL = strSQL & " AND bh_prft_ctr <> " & tfnRound(txtEHPrftCtr)
        strSQL = strSQL & " AND bh_empno = " & tfnRound(sText)
        strSQL = strSQL & " AND bh_pay_code = " & tfnSQLString(sHOURLYPAYCODE)
        'david 12/31/2002  #393054-2
        strSQL = strSQL & " AND bh_prft_ctr NOT IN (" + sPrftCtrList + ")"
        '''''''''''''''''
        
        If GetRecordSet(rsTemp, strSQL, nDB_REMOTE, "fnValidEmpNo") <= 0 Then
            tgcEditor.ErrorMessage(nCol) = "Database Access Error"
            Exit Function
        End If
        
        nCountPrftCtr = nCountPrftCtr + tfnRound(rsTemp!prftCount)
        ''''''''''''''''''''''''''
            
        If nCountPrftCtr >= 5 Then
            tgcEditor.ErrorMessage(nCol) = "Employee may only have hours for five separate profit centers in a single pay period."
            Exit Function
        End If
    End If

    If t_nFormMode = ADD_MODE Then
        
        For i = 0 To tgcEditor.RowCount - 1
            
            If i <> lRow Then
                
                If tfnRound(tgcEditor.CellValue(COL_EMPNO, i)) = tfnRound(sText) Then
                    tgcEditor.ErrorMessage(nCol) = "Duplicate Employees Entered"
                    fnValidEmpNo = False
                    Exit Function
                End If
                
            End If
        Next
        
    End If
    
    tgcEditor.CellValue(COL_EMPNAME, lRow) = sEmpName
    fnValidEmpNo = True
End Function


Private Sub subReSetComboSql()
    Dim strSQL As String
    Dim sPayCodeList As String

    sPayCodeList = fnGetPayCodeList
    
    If t_nFormMode = EDIT_MODE Then
        strSQL = "SELECT DISTINCT prft_ctr, prft_name FROM sys_prft_ctr, bonus_hold"
        strSQL = strSQL & " WHERE prft_ctr = bh_prft_ctr"
    Else
        strSQL = "SELECT prft_ctr, prft_name FROM sys_prft_ctr"
    End If

    tgcDropdown.ComboSQL(txtEHPrftCtr) = strSQL
    
    
    strSQL = "SELECT DISTINCT pcode, pdesc FROM " & sTempPrPay
    strSQL = strSQL & " WHERE pcode NOT IN " & sPayCodeList
    
    tgcDropdown.ComboSQL(txtEHPayCode) = strSQL
    
    strSQL = "SELECT DISTINCT bh_date FROM bonus_hold WHERE bh_chk_link = 0"
    
    
    If cValidate.ValidInput(txtEHPrftCtr) Then
        strSQL = strSQL & " AND bh_prft_ctr = " & tfnRound(txtEHPrftCtr)
    End If
    
    tgcDropdown.ComboSQL(txtEHDate) = strSQL
End Sub

Private Sub subEnableProfitCtr(bFlag As Boolean)
    txtEHPrftCtr.Enabled = bFlag
    txtEHPrftName.Enabled = bFlag
    txtEHDate.Enabled = bFlag
    subEnableSearchButton cmdEHPrftCtr, False
    subEnableSearchButton cmdEHPrftName, False
    subEnableSearchButton cmdEHDate, False
End Sub

Private Sub subSetFormStatus()
    
    If t_nFormMode = EDIT_MODE And nDataStatus = DATA_CHANGED Then
        subEnableRefresh True
    End If
    
    If cValidate.FirstInvalidInput < 0 And tgcEditor.ValidData And fnGetField(tgcEditor.CellValue(0, 0)) <> "" And _
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

Private Function fnGetPayCodeList(Optional bDeleting As Boolean = False) As String
    Dim sPayCodeList As String
    
    sPayCodeList = "('" & sHOURLYPAYCODE & "'"
    
    If tblTimeCard.Columns(COL_PAYCODE1).Caption <> sDefaultPAYCODE1 Then
        
        If bDeleting Then
            sPayCodeList = sPayCodeList & ",'" & g_sOrigPayCode1 & "'"
        Else
            sPayCodeList = sPayCodeList & ",'" & Trim(tblTimeCard.Columns(COL_PAYCODE1).Caption) & "'"
        End If
        
    End If
    
    If tblTimeCard.Columns(COL_PAYCODE2).Caption <> sDefaultPAYCODE2 Then
    
        If bDeleting Then
            sPayCodeList = sPayCodeList & ",'" & g_sOrigPayCode2 & "'"
        Else
            sPayCodeList = sPayCodeList & ",'" & Trim(tblTimeCard.Columns(COL_PAYCODE2).Caption) & "'"
        End If
        
    End If
    
    If tblTimeCard.Columns(COL_PAYCODE3).Caption <> sDefaultPAYCODE3 Then
        
        If bDeleting Then
            sPayCodeList = sPayCodeList & ",'" & g_sOrigPayCode3 & "'"
        Else
            sPayCodeList = sPayCodeList & ",'" & Trim(tblTimeCard.Columns(COL_PAYCODE3).Caption) & "'"
        End If
        
    End If
    
    sPayCodeList = sPayCodeList & ")"

    fnGetPayCodeList = sPayCodeList
End Function

Private Function fnValidDate(txtBox As Textbox) As Boolean
    Dim sDate As String
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    sDate = Trim(txtBox.Text)
    
    If Not IsFactorDate(sDate) Then 'the format of sztext has been changed here
        cValidate.SetErrorMessage txtEHDate, "Invalid Date entered"
        Exit Function
    End If
    
    'check in table gl_period
    strSQL = "SELECT glp_status FROM gl_period "
    strSQL = strSQL & " WHERE glp_end_dt >= " & tfnDateString(sDate, True)
    strSQL = strSQL & " AND glp_beg_dt <= " & tfnDateString(sDate, True)
    
    If GetRecordSet(rsTemp, strSQL, nDB_REMOTE, "fnvalidDate") > 0 Then
        
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
            cValidate.SetErrorMessage txtEHDate, "This G/L Period is already closed"
        Else
            cValidate.SetErrorMessage txtEHDate, "Invalid G/L Period Status"
        End If
        
    Else
        cValidate.SetErrorMessage txtEHDate, "There is no valid G/L Period containing this Date"
    End If
    
End Function


Private Function fnValidProfitCtr(txtBox As Textbox) As Boolean
    Dim sCode As String
    Dim strSQL As String
    Dim rsTemp As Recordset
                
    fnValidProfitCtr = False
    
    sCode = Trim(txtBox.Text)
    
    If Trim(sCode) = "" Then
        cValidate.SetErrorMessage txtEHPrftCtr, "You must enter a Profit Center Number"
        Exit Function
    End If
                    
    strSQL = "SELECT prft_ctr, prft_name FROM sys_prft_ctr WHERE prft_ctr = " & tfnRound(sCode)
    
    If GetRecordSet(rsTemp, strSQL) <= 0 Then
        cValidate.SetErrorMessage txtEHPrftCtr, "Invalid Profit Center Number"
        Exit Function
    Else
        tgcDropdown.SetText txtEHPrftCtr, fnGetField(rsTemp!prft_name)
    End If
            
'    sDate = tfnFormatDate(tgcEditor.CellValue(COL_EMPNO, lRow))
'    sCode = fnGetField(tgcEditor.CellValue(COL_CODE, lRow))
'
'    sSql = "SELECT bh_empno FROM bonus_hold " _
'        & " WHERE bh_empno = " & tfnRound(txtEHPrftCtr.Text) _
'        & " AND bh_prft_ctr = " & tfnRound(szText) _
'        & " AND bh_date = " & tfnSQLString(tfnDateString(sDate)) _
'        & " AND bh_pay_code = " & tfnSQLString(sCode)
'
'    If GetRecordCount(sSql) > 0 And t_nFormMode = ADD_MODE Then
'        tgcEditor.ErrorMessage(COL_EMPNAME) = "Date/Profit Center/Pay Code already exists for this Employee"
'        Exit Function
'    End If
'
    fnValidProfitCtr = True
End Function

Private Function fnValidPayCode(txtBox As Textbox) As Boolean
    Dim strSQL As String
    Dim lRow As Long
    Dim sDate As String
    Dim nPrft As Integer
    Dim rsTemp As Recordset
    Dim szText As String
    
    szText = Trim(txtBox.Text)
    
    If Trim(szText) = "" Then
        fnValidPayCode = True
        Exit Function
    End If
      
    strSQL = " SELECT pcode FROM " & sTempPrPay
    strSQL = strSQL & " WHERE pcode = " & tfnSQLString(szText)
    
    If GetRecordCount(strSQL) <= 0 Then
        cValidate.SetErrorMessage txtBox, "Pay Code does not exist"
        Exit Function
    End If
    
    If txtEHPayCode.Tag <> "" Then
        
        If szText = sHOURLYPAYCODE Then
            cValidate.SetErrorMessage txtBox, "Can't use Pay Code " & sHOURLYPAYCODE
        Else
            fnValidPayCode = True
            
            Select Case CInt(txtEHPayCode.Tag)
                Case COL_PAYCODE1
                    
                    If szText = tblTimeCard.Columns(COL_PAYCODE2).Caption _
                        Or szText = tblTimeCard.Columns(COL_PAYCODE3).Caption Then
                        cValidate.SetErrorMessage txtBox, "This Pay Code was selected."
                        fnValidPayCode = False
                    End If
                    
                    
                Case COL_PAYCODE2
                    
                    If szText = tblTimeCard.Columns(COL_PAYCODE1).Caption _
                        Or szText = tblTimeCard.Columns(COL_PAYCODE3).Caption Then
                        cValidate.SetErrorMessage txtBox, "This Pay Code was selected."
                        fnValidPayCode = False
                    End If
                Case COL_PAYCODE3
                    
                    If szText = tblTimeCard.Columns(COL_PAYCODE2).Caption _
                        Or szText = tblTimeCard.Columns(COL_PAYCODE1).Caption Then
                        cValidate.SetErrorMessage txtBox, "This Pay Code was selected."
                        fnValidPayCode = False
                    End If
                    
            End Select
            
        End If
        
    End If
    
    Exit Function
End Function

Private Sub subResetFloatingSQL()
    Dim strSQL As String
    
    'all employee, don't care the profit center
    strSQL = "SELECT empno, empname FROM " & sTempPrMaster
    
    tgcFEditor.SetSQL COL_EMPNO, strSQL

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
         
         .AddColumn COL_EMPNO
         .ColumnCaption(COL_EMPNO) = "Employee #"
         .ColumnType(COL_EMPNO) = tgcFEditor.COLUMN_TYPE_INTEGER
         .ColumnDataField(COL_EMPNO) = "empno"
         .SetUnique COL_EMPNO, True
         
         .AddColumn COL_EMPNAME
         .ColumnCaption(COL_EMPNAME) = "Employee Name"
         .ColumnType(COL_EMPNAME) = tgcFEditor.COLUMN_TYPE_STRING
         .ColumnDataField(COL_EMPNAME) = "empname"
         
    End With
    
End Sub

Private Sub Class_Terminate()
    Set m_Form = Nothing
    Set tbToolbar = Nothing
    Set ffraStatusbar = Nothing
    Set txtEHPrftCtr = Nothing
    Set cmdEHPrftCtr = Nothing
    Set txtEHPrftName = Nothing
    Set cmdEHPrftName = Nothing
    Set txtEHDate = Nothing
    Set datComboDropDown = Nothing
    Set tblComboDropdown = Nothing
    Set tblTimeCard = Nothing
    Set tblFloating = Nothing
    Set cmdFloatingBtn = Nothing
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

Private Function fnCheckHeaderOK(nCol As Integer) As Boolean
    Dim i As Long
    
    If nCol <= COL_HOUR Then
        fnCheckHeaderOK = False
    Else
        fnCheckHeaderOK = True
        
        For i = 0 To tgcEditor.RowCount - 1
            
            If tfnRound(tgcEditor.CellValue(nCol, i), 2) > 0 Then
                fnCheckHeaderOK = False
                Exit For
            End If
            
        Next
        
    End If
    
End Function

Private Sub subCalculateTotal()
    Dim i As Long
    Dim dTotalHours As Single
    Dim dTotalPayCode1 As Single
    Dim dTotalPayCode2 As Single
    Dim dTotalPayCode3 As Single
    
    For i = 0 To tgcEditor.RowCount - 1
        dTotalHours = dTotalHours + tfnRound(tgcEditor.CellValue(COL_HOUR, i), 2)
        dTotalPayCode1 = dTotalPayCode1 + tfnRound(tgcEditor.CellValue(COL_PAYCODE1, i), 2)
        dTotalPayCode2 = dTotalPayCode2 + tfnRound(tgcEditor.CellValue(COL_PAYCODE2, i), 2)
        dTotalPayCode3 = dTotalPayCode3 + tfnRound(tgcEditor.CellValue(COL_PAYCODE3, i), 2)
    Next
    
    lblEHTotalHours.Caption = CStr(Format(dTotalHours, "#,##0.00"))
    lblEHTotalPayCode1.Caption = CStr(Format(dTotalPayCode1, "#,##0.00"))
    lblEHTotalPayCode2.Caption = CStr(Format(dTotalPayCode2, "#,##0.00"))
    lblEHTotalPayCode3.Caption = CStr(Format(dTotalPayCode3, "#,##0.00"))
End Sub

'Vijaya on 12/12/03 Magic#421627
Public Sub txtEHMgrHours_Change()
    cValidate.Change txtEHMgrHours
    tfnRegExpControlChange txtEHMgrHours, tfnDecimalPattern(8, 2)
    cmdUpdateInsertBtn.Enabled = False
    nDataStatus = DATA_CHANGED
End Sub

Public Sub txtEHMgrHours_GotFocus()
    Set cValidate.LastBox = txtEHMgrHours
    cValidate.GotFocus txtEHMgrHours
    subSelectIt txtEHMgrHours
End Sub

Public Sub txtEHMgrHours_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        subSetFormStatus
        If cmdUpdateInsertBtn.Enabled = True Then
            subSetFocus cmdUpdateInsertBtn
        Else
            subSetFocus efraBridge
        End If
        KeyAscii = 0
    Else
        tfnRegExpControlKeyPress txtEHMgrHours, KeyAscii, tfnDecimalPattern(8, 2)
        cValidate.Keypress txtEHMgrHours, KeyAscii
    End If
End Sub

Public Sub txtEHMgrHours_LostFocus()
    cValidate.LostFocus txtEHMgrHours
    subSetFormStatus
    If frmZZSEBPRC.ActiveControl Is cmdExitCancelBtn Then
        If cmdUpdateInsertBtn.Enabled Then
            subSetFocus cmdUpdateInsertBtn
        End If
    End If
End Sub

Private Sub subSelectIt(v As Textbox)
    On Error Resume Next
    v.SelStart = 0
    v.SelLength = Len(Trim(v.Text))
End Sub

Private Function fnValidMgrHours(txtBox As Textbox) As Boolean
    
    'This control is Optional, Null or 0 valid
    If Trim(txtBox) = "" Or Trim(txtBox) = 0 Then
        fnValidMgrHours = True
        Exit Function
    End If
    
    If Not IsNumeric(txtBox) Then
        cValidate.SetErrorMessage txtBox, "Manager Hours must be a numeric"
        Exit Function
    End If
    
    If bMgrHoursValid Then
        fnValidMgrHours = True
        Exit Function
    End If
    
    'david 12/18/2003  #421627
    'return error message if any
    'will set the lMgrEmpNbr variable to a valid manager's emp nbr
    Dim sErrMsg As String
    sErrMsg = fnCheckMgrEmpNbr()
    
    If sErrMsg <> "" Then
        cValidate.SetErrorMessage txtBox, sErrMsg
        Exit Function
    End If
    ''''''''''''''''''''''''''
    
    fnValidMgrHours = True

End Function

Public Function fnInsertzzse_manager_hour() As Boolean
    Const FUNC_NAME As String = "fnInsertzzse_manager_hour"
    
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    If Not txtEHMgrHours.Visible Then
        fnInsertzzse_manager_hour = True
        Exit Function
    End If
    
    If Trim(txtEHMgrHours) = "" Or Trim(txtEHMgrHours) = 0 Then
        strSQL = "DELETE FROM zzse_manager_hour"
        strSQL = strSQL & " WHERE zzsemh_prft_ctr =" & tfnRound(txtEHPrftCtr)
        strSQL = strSQL & " AND zzsemh_date =" & tfnDateString(txtEHDate, True)
        If fnExecuteSQL(strSQL, FUNC_NAME) < 0 Then
            Exit Function
        End If
    End If
    'Empty or 0 we don't need to Insert..
    If Trim(txtEHMgrHours) = "" Or Trim(txtEHMgrHours) = 0 Then
        fnInsertzzse_manager_hour = True
        Exit Function
    End If
    
    strSQL = "SELECT zzsemh_empno FROM zzse_manager_hour"
    strSQL = strSQL & " WHERE zzsemh_prft_ctr = " & tfnRound(txtEHPrftCtr)
    strSQL = strSQL & " AND zzsemh_date =" & tfnDateString(txtEHDate, True)
    If GetRecordSet(rsTemp, strSQL, , FUNC_NAME) < 0 Then
        Exit Function
    End If
    
    If rsTemp.RecordCount > 0 Then
        strSQL = "UPDATE zzse_manager_hour SET"
        If Not bMgrHoursValid Then
            strSQL = strSQL & " zzsemh_empno=" & lMgrEmpNbr & ","
        End If
        strSQL = strSQL & " zzsemh_hours=" & Format(txtEHMgrHours, "0.00")
        strSQL = strSQL & " WHERE zzsemh_prft_ctr=" & tfnRound(txtEHPrftCtr)
        strSQL = strSQL & " AND zzsemh_date=" & tfnDateString(txtEHDate, True)
    Else
        strSQL = "INSERT INTO zzse_manager_hour (zzsemh_empno, zzsemh_prft_ctr,"
        strSQL = strSQL & " zzsemh_date,zzsemh_hours) VALUES (" & lMgrEmpNbr
        strSQL = strSQL & " ," & tfnRound(txtEHPrftCtr) & "," & tfnDateString(txtEHDate, True)
        strSQL = strSQL & "," & Format(txtEHMgrHours, "0.00") & ")"
    End If
    
    If fnExecuteSQL(strSQL, FUNC_NAME) < 0 Then
        Exit Function
    End If
    
    fnInsertzzse_manager_hour = True
    
End Function
'End of Vijaya code

'return error message if any
'will set the lMgrEmpNbr variable to a valid manager's emp nbr
Private Function fnCheckMgrEmpNbr() As String
    
    Const SUB_NAME As String = "fnCheckMgrEmpNbr"
    
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim sMsg As String
    Dim sPrftCtrList As String
    
    strSQL = "SELECT prm_empno FROM pr_master WHERE prm_job_class  = 'A'"
    strSQL = strSQL & " AND prm_emp_level  = 3"
    strSQL = strSQL & " AND " & tfnDateString(txtEHDate, True) & " BETWEEN prm_date_hired AND prm_date_termed"
    strSQL = strSQL & " AND prm_prft_ctr1 = " & tfnRound(txtEHPrftCtr)
    
    If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
        fnCheckMgrEmpNbr = "Failed to access database"
        Exit Function
    End If
    
    If rsTemp.RecordCount = 0 Then
        strSQL = "SELECT prm_empno FROM pr_master WHERE prm_job_class  = 'A'"
        strSQL = strSQL & " AND prm_emp_level  = 3"
        strSQL = strSQL & " AND prm_date_termed IS NULL"
        strSQL = strSQL & " AND prm_prft_ctr1 = " & tfnRound(txtEHPrftCtr)
        
        If GetRecordSet(rsTemp, strSQL, , SUB_NAME) < 0 Then
            fnCheckMgrEmpNbr = "Failed to access database"
            Exit Function
        End If
    End If
    
    If rsTemp.RecordCount = 0 Then
        sMsg = "No manager was found for the profit center."
        sMsg = sMsg & " Use the Payroll Employee Master Maintenance to add a manager"
        sMsg = sMsg & " for the profit center. Manager's hours cannot be entered until"
        sMsg = sMsg & " one manager exists."
        DoEvents
        MsgBox sMsg, vbExclamation
        fnCheckMgrEmpNbr = "Manager not found"
        Exit Function
    End If
    
    If rsTemp.RecordCount > 1 Then
        While Not rsTemp.EOF
            If sPrftCtrList = "" Then
                sPrftCtrList = tfnRound(fnGetField(rsTemp!prm_empno))
            Else
                sPrftCtrList = sPrftCtrList & "," & tfnRound(fnGetField(rsTemp!prm_empno))
            End If
            rsTemp.MoveNext
        Wend
        sMsg = "There is more than one manager assigned to this profit center."
        sMsg = sMsg & " Use the Payroll Employee Master Maintenance to add a termination"
        sMsg = sMsg & " date for the invalid manager(s). Manager's hours cannot be inserted"
        sMsg = sMsg & " until only one manager exists."
        DoEvents
        MsgBox sMsg, vbExclamation
        fnCheckMgrEmpNbr = "More than one manager (" & sPrftCtrList & ") found"
        Exit Function
    End If
    
    lMgrEmpNbr = tfnRound(rsTemp!prm_empno)
    
End Function
