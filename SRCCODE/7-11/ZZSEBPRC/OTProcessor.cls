VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOTProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Form As Form
Private tbToolbar As Toolbar
Private ffraStatusbar As FactorFrame
Private pbStatus As ProgressBar
Private txtOTWeek1BeginDate As Textbox
Private txtOTWeek1EndDate As Textbox
Private txtOTWeek2BeginDate As Textbox
Private txtOTWeek2EndDate As Textbox
Private lstOTLog As ListBox
Private cmdOTPrintBtn As FactorFrame
Private cmdOTProcessBtn  As FactorFrame
Private cmdUpdateInsertBtn As FactorFrame
Private cmdExitCancelBtn As FactorFrame
Private cmdExitBtn As FactorFrame
Private efraBridge As FactorFrame
Private mnuCancel As Menu


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

Private cValidate As cValidateInput
Private Const sHOURLYPAYCODE As String = "7HRP"
Private Const sHourPayCodeList As String = "('7HRP', '7HRA', '7HRD', '7HRI', '7HRO')"
Enum MaintenanceType
    TypeHold = 0
    TypeInsert
    TypeUpdate
    TypeDelete
End Enum

Private Type BONUS_DATA
    EmpNO As Long
    PrftCtr As Integer
    payCode As String
    Hours As Single
    bonusDate As Date
    DoType As MaintenanceType
End Type

Private Type EMP_HOURS
    EmpNO As Long
    PrftCtr As Integer
    nWeek As Integer
    OrigHours As Single
    RegHours As Single
    OTHOurs As Single
    Parsed As Boolean
End Type

Private g_ArrayBonus() As BONUS_DATA

Private g_ArrayHours() As EMP_HOURS
Private g_lEmpHoursCount As Long
Private g_bOverTimeRecordFound As Boolean
Private g_bOverPayCodeNotSet As Boolean
Private g_lBonusCount As Long
Private Const INCARRAYZISE As Integer = 10
Private Const MAXWEEKHOURS As Integer = 40
'Private g_sOverTimePayCode As String
'Those variable is used to calculate hours for prft and employee
Private g_dWeek1EmployeeHours As Single
Private g_dWeek2EmployeeHours As Single
Private g_lHoldEmpNO As Long

'david 03/11/2003  #403275
Private bOverTimePayCodeSet As Boolean
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

Property Set myProgressBar(pbar As ProgressBar)
    Set pbStatus = pbar
End Property

Property Set OTWeek1BeginDate(txtBox As Textbox)
    Set txtOTWeek1BeginDate = txtBox
End Property

Property Set OTWeek1EndDate(txtBox As Textbox)
    Set txtOTWeek1EndDate = txtBox
End Property

Property Set OTWeek2BeginDate(txtBox As Textbox)
    Set txtOTWeek2BeginDate = txtBox
End Property

Property Set OTWeek2EndDate(txtBox As Textbox)
    Set txtOTWeek2EndDate = txtBox
End Property

Property Set OTProcessLog(lstBox As ListBox)
    Set lstOTLog = lstBox
    'make the font size 6 other can't view entire row
    'was set at the design time, but size was changed after compiled.
    lstOTLog.Font.Size = 6
End Property

Property Set OTPrintButton(cmdButton As FactorFrame)
    Set cmdOTPrintBtn = cmdButton
End Property

Property Set OTProcessButton(cmdButton As FactorFrame)
    Set cmdOTProcessBtn = cmdButton
End Property

Property Set UpdateInsertButton(btn As FactorFrame)
    Set cmdUpdateInsertBtn = btn
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

Public Sub Form_Initialize() 'called before Form_Load
    
    t_bStartupFlag = True
    t_bDataChanged = False
    t_bUpdateTable = False
    
    t_nFormMode = IDLE_MODE
    
End Sub

Public Sub efraBridge_GotFocus()
    cValidate.GotFocus efraBridge
End Sub

Public Sub efraBridge_LostFocus()
    cValidate.LostFocus efraBridge
End Sub

Public Sub Form_Load()
    subSetValidtion
    tfnResetScreen 'set the default screen
End Sub

Private Sub subSetValidtion()
    Set cValidate = New cValidateInput
    
    With cValidate
        Set .StatusBar = ffraStatusbar
        Set .Form = m_Form
        .AddEditBox txtOTWeek1BeginDate, "Enter Week 1 Begin Date"
        .AddEditBox txtOTWeek1EndDate, "Enter Week 1 End Date."
        .AddEditBox txtOTWeek2BeginDate, "Enter Week 2 Begin Date"
        .AddEditBox txtOTWeek2EndDate, "Enter Week 2 End Date"
        .ESCControl = cmdOTProcessBtn
        .ESCControl = cmdExitCancelBtn
        .ESCControl = cmdExitBtn
        .MinTabIndex = tbToolbar.TabIndex
        .MaxTabIndex = efraBridge.TabIndex + 1
        Set .ControlForFocus = efraBridge
        Set .LastBox = txtOTWeek2EndDate
        .SetFirstControls cmdOTProcessBtn, cmdExitCancelBtn
    End With
    
End Sub

Private Sub tfnResetScreen()
    
    On Error Resume Next
    txtOTWeek1BeginDate.Text = ""
    txtOTWeek1EndDate.Text = ""
    txtOTWeek2BeginDate.Text = ""
    txtOTWeek2EndDate.Text = ""
    lstOTLog.Clear
    
    cmdUpdateInsertBtn.Enabled = False
    cmdUpdateInsertBtn.Caption = t_szCAPTION_INSERT
    cmdOTPrintBtn.Enabled = False
    cmdOTProcessBtn.Enabled = False
    cValidate.ResetFlags
    
    frmContext.ButtonEnabled(CANCEL_UP) = False
    cmdExitCancelBtn.Enabled = False
    mnuCancel.Enabled = False
    SubEnableTextBox True
    subSetFocus txtOTWeek1BeginDate
    t_bDataChanged = False
End Sub

Private Sub SubEnableTextBox(bEnabled As Boolean)
    txtOTWeek1BeginDate.Enabled = bEnabled
    txtOTWeek1EndDate.Enabled = bEnabled
    txtOTWeek2BeginDate.Enabled = bEnabled
    txtOTWeek2EndDate.Enabled = bEnabled
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

Public Sub txtOTWeek1BeginDate_Change()
    cValidate.Change txtOTWeek1BeginDate
    subDataChanged
    t_bDataChanged = True
End Sub

Public Sub txtOTWeek1BeginDate_GotFocus()
    cValidate.GotFocus txtOTWeek1BeginDate
    subTextSelected
End Sub

Public Sub txtOTWeek1BeginDate_KeyPress(KeyAscii As Integer)
    tfnRegExpControlDateKeyPress txtOTWeek1BeginDate, KeyAscii
    cValidate.Keypress txtOTWeek1BeginDate, KeyAscii
    
    If KeyAscii = vbKeyReturn Then
        subSetFocus txtOTWeek1EndDate
        KeyAscii = 0
    End If
    
End Sub

Public Sub txtOTWeek1BeginDate_LostFocus()
    
    If cValidate.LostFocus(txtOTWeek1BeginDate) Then
        subCheckButtonStatus
    End If
    
End Sub

Public Sub txtOTWeek1EndDate_Change()
    cValidate.Change txtOTWeek1EndDate
    subDataChanged
    t_bDataChanged = True
End Sub

Public Sub txtOTWeek1EndDate_GotFocus()
    cValidate.GotFocus txtOTWeek1EndDate
    subTextSelected
End Sub

Public Sub txtOTWeek1EndDate_KeyPress(KeyAscii As Integer)
    tfnRegExpControlDateKeyPress txtOTWeek1EndDate, KeyAscii
    cValidate.Keypress txtOTWeek1EndDate, KeyAscii
    
    If KeyAscii = vbKeyReturn Then
        subSetFocus txtOTWeek2BeginDate
        KeyAscii = 0
    End If
    
End Sub

Public Sub txtOTWeek1EndDate_LostFocus()
    
    If cValidate.LostFocus(txtOTWeek1EndDate) Then
        subCheckButtonStatus
    End If
    
End Sub

Public Sub txtOTWeek2BeginDate_Change()
    cValidate.Change txtOTWeek2BeginDate
    subDataChanged
    t_bDataChanged = True
End Sub

Public Sub txtOTWeek2BeginDate_GotFocus()
    cValidate.GotFocus txtOTWeek2BeginDate
    subTextSelected
End Sub

Public Sub txtOTWeek2BeginDate_KeyPress(KeyAscii As Integer)
    tfnRegExpControlDateKeyPress txtOTWeek2BeginDate, KeyAscii
    cValidate.Keypress txtOTWeek2BeginDate, KeyAscii
    
    If KeyAscii = vbKeyReturn Then
        subSetFocus txtOTWeek2EndDate
        KeyAscii = 0
    End If
    
End Sub

Public Sub txtOTWeek2BeginDate_LostFocus()
    
    If cValidate.LostFocus(txtOTWeek2BeginDate) Then
        subCheckButtonStatus
    End If
    
End Sub

Public Sub txtOTWeek2EndDate_Change()
    cValidate.Change txtOTWeek2EndDate
    subDataChanged
    t_bDataChanged = True
End Sub

Public Sub txtOTWeek2EndDate_GotFocus()
    cValidate.GotFocus txtOTWeek2EndDate
    subTextSelected
End Sub

Public Sub txtOTWeek2EndDate_KeyPress(KeyAscii As Integer)
    
    tfnRegExpControlDateKeyPress txtOTWeek2EndDate, KeyAscii
    cValidate.Keypress txtOTWeek2EndDate, KeyAscii
    
    If KeyAscii = vbKeyReturn Then
        subCheckButtonStatus
        
        If cmdOTProcessBtn.Enabled Then
            subSetFocus cmdOTProcessBtn
        Else
            subSetFocus efraBridge
            'SendKeys "{TAB}"
        End If
        
        KeyAscii = 0
        
    End If
    
End Sub

Public Sub txtOTWeek2EndDate_LostFocus()
        
    If cValidate.LostFocus(txtOTWeek2EndDate) Then
        subCheckButtonStatus
    End If
    
End Sub

Public Sub cmdUpdateInsertBtn_Click()
    Screen.MousePointer = vbHourglass
    t_bUpdateTable = fnInsertUpdate
    Screen.MousePointer = vbDefault
    
    If t_bUpdateTable Then
        cmdUpdateInsertBtn.Enabled = False
        t_bDataChanged = False
        tfnResetScreen
    End If
    
End Sub

Public Sub cmdUpdateInsertBtn_GotFocus()
    tfnSetStatusBarMessage "Insert"
End Sub

Public Sub cmdOTPrint_Click()
    Dim sReportTitle As String
    Dim sArrReport() As String
    Dim nHeader As Integer
    Dim i As Long
    
    tfnSetStatusBarCorrect "Printing Report, Please wait..."
    Screen.MousePointer = vbHourglass
    
    sReportTitle = "Overtime Processing Report"
    nHeader = 3
    
    If Trim(txtOTWeek2BeginDate) <> "" Then
        nHeader = 4
    End If
        
    subSetReportID "ZZSEBPRC"
     
    If lstOTLog.ListCount > 0 Then
        ReDim sArrReport(lstOTLog.ListCount + nHeader - 1)
        
        sArrReport(0) = Space(54) & "Week 1 " & Trim(txtOTWeek1BeginDate) & " - " & Trim(txtOTWeek1EndDate)
        
        If nHeader = 4 Then
            sArrReport(1) = Space(54) & "Week 2 " & Trim(txtOTWeek2BeginDate) & " - " & Trim(txtOTWeek2EndDate)
        End If
        
        sArrReport(nHeader - 2) = ""
        sArrReport(nHeader - 1) = String(132, "=")
        
        For i = 0 To lstOTLog.ListCount - 1
            sArrReport(i + nHeader) = lstOTLog.List(i)
        Next
        
        If Not fnSetupPrinter(vbPRORLandscape) Then
            tfnSetStatusBarError "Failed to print report"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    
        If Not fnSendToPrinter(sArrReport(), sReportTitle) Then
            tfnSetStatusBarError "Failed to print report"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    
    End If
    
    Screen.MousePointer = vbDefault
    tfnSetStatusBarCorrect "Printing Done."
End Sub

Public Sub cmdOTPrint_GotFocus()
    tfnSetStatusBarMessage "Print"
End Sub

Public Sub cmdOtProcess_GotFocus()
    tfnSetStatusBarMessage "Process"
End Sub

Public Sub cmdOTProcess_Click()
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim lEmpNo As Long
    Dim sEmpName As String
    Dim lCount As Long
    Dim i As Long
    
    g_bOverTimeRecordFound = False
    g_bOverPayCodeNotSet = False
    bOverTimePayCodeSet = True
    
    'First Check the date set are correctly or not.
    If cValidate.FirstInvalidInput >= 0 Then
        MsgBox "Please set correct date range!", vbExclamation
        GoTo EXITSUB
    End If
    
    If Trim(txtOTWeek2BeginDate) <> "" And Trim(txtOTWeek2BeginDate) <= Trim(txtOTWeek1EndDate) Then
        MsgBox "The Week 2 begin date can't be earlier than Week 1 end date", vbExclamation
        GoTo EXITSUB
    End If
    
    SubEnableTextBox False
    cmdOTProcessBtn.Enabled = False
    lstOTLog.Clear
    tfnSetStatusBarCorrect "Processing Data, Please wait..."
    Screen.MousePointer = vbHourglass
    
    'GetOverTimeCode
'    g_sOverTimePayCode = fnGetOverTimePayCode()
'    If g_sOverTimePayCode = "" Then
'        subWriteLog "Overtime Pay Code for " & sHOURLYPAYCODE & " Pay Code cannot be determined, Processing is Stopped"
'        GoTo EXITSUB
'    End If
    
    'Build the Employee Array
    
    If Not fnBuildEmployeeHoursArray() Then
        GoTo EXITSUB
    End If
    
    ReDim g_ArrayBonus(UBound(g_ArrayHours)) As BONUS_DATA
    g_lBonusCount = 0
    lEmpNo = -1
    
    subBuildReportHeader
    'ananlyze empno one by one
    For i = 0 To UBound(g_ArrayHours)
        
        If Not g_ArrayHours(i).Parsed Then
            lEmpNo = g_ArrayHours(i).EmpNO
            tfnSetStatusBarCorrect "Parsing hours information for employee " & lEmpNo & "..."
    
            g_lHoldEmpNO = -1
            
            If Not fnParseOneEmployee(lEmpNo) Then
                GoTo EXITSUB
            End If
            
            subSetEmployeeParsed lEmpNo
        End If
        
        subSetProgress (i + 1) / (UBound(g_ArrayHours) + 1) * 100
            
    Next
        
EXITSUB:
    subSetProgress 0
    
    If Not g_bOverTimeRecordFound Then
        If Not bOverTimePayCodeSet Then
            cmdOTPrintBtn.Enabled = True
            subSetFocus cmdOTPrintBtn
            'force statusbar message to show
            DoEvents
            tfnSetStatusBarError "Process failed"
        Else
            subWriteLog "No Record selected"
            tfnSetStatusBarCorrect "Process Done."
        End If
    ElseIf g_bOverPayCodeNotSet Then
        tfnSetStatusBarError "Process failed"
    Else
        cmdUpdateInsertBtn.Enabled = True
        cmdOTPrintBtn.Enabled = True
        subSetFocus cmdOTPrintBtn
        'force statusbar message to show
        DoEvents
        tfnSetStatusBarCorrect "Process Done."
        
        If g_lBonusCount > 0 Then
            ReDim Preserve g_ArrayBonus(g_lBonusCount - 1)
        End If
    
    End If
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub subSetEmployeeParsed(lEmpNo As Long)
    Dim i As Long
    
    For i = 0 To UBound(g_ArrayHours)
        
        If g_ArrayHours(i).EmpNO = lEmpNo Then
            g_ArrayHours(i).Parsed = True
        End If
        
    Next
    
End Sub

Private Function fnParseOneEmployee(lEmpNo As Long) As Boolean
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim nPrftCtr As Integer
    Dim sPrftName As String
    Dim dDate As Date
    Dim dHours As Single
    Dim nPrevPrftCtr As Integer
    Dim dOTHours As Single
    Dim dRegHours As Single
    Dim sMsg As String
    
    On Error GoTo ErrorHandler
    
    'david 03/11/2003  #403275
    'validate ot paycode for the employee
    Dim bError As Boolean
    sMsg = fnValidOtPaycode(lEmpNo, sHourPayCodeList, bError)
    
    fnParseOneEmployee = Not bError
    
    If sMsg <> "" Then
        subWriteLog sMsg
        bOverTimePayCodeSet = False
        Exit Function
    End If
    ''''''''''''''''''''''''''
    
    g_dWeek1EmployeeHours = 0
    g_dWeek2EmployeeHours = 0
    
    strSQL = "SELECT bh_empno, bh_prft_ctr, bh_hours, bh_date, bh_pay_code FROM bonus_hold "
    strSQL = strSQL & " WHERE bh_empno = " & lEmpNo
    strSQL = strSQL & " AND bh_chk_link =  0"
    'strSQL = strSQL & " AND bh_pay_code = " & tfnSQLString(sHOURLYPAYCODE)
    strSQL = strSQL & " AND bh_pay_code IN " & sHourPayCodeList
    
    If Trim(txtOTWeek2BeginDate) <> "" Then
        strSQL = strSQL & " AND ((bh_date BETWEEN " & tfnDateString(txtOTWeek1BeginDate, True)
        strSQL = strSQL & " AND " & tfnSQLString(txtOTWeek1EndDate, True) & ")"
        strSQL = strSQL & " OR (bh_date BETWEEN " & tfnDateString(txtOTWeek2BeginDate, True)
        strSQL = strSQL & " AND " & tfnSQLString(txtOTWeek2EndDate, True) & "))"
    Else
        strSQL = strSQL & " AND bh_date BETWEEN " & tfnDateString(txtOTWeek1BeginDate, True)
        strSQL = strSQL & " AND " & tfnSQLString(txtOTWeek1EndDate, True)
    End If
    
    strSQL = strSQL & " ORDER BY bh_empno,bh_prft_ctr, bh_date"
    
    If GetRecordSet(rsTemp, strSQL, nDB_REMOTE, "fnParseOneEmployee") <= 0 Then
        Exit Function
    End If
    
    nPrevPrftCtr = -1
    
    Do Until rsTemp.EOF
        dDate = rsTemp!bh_date
        dHours = tfnRound(rsTemp!bh_Hours, 2)
        nPrftCtr = tfnRound(rsTemp!bh_prft_ctr)
        
        If nPrevPrftCtr <> nPrftCtr Then
            g_dWeek1EmployeeHours = 0
            g_dWeek2EmployeeHours = 0
            subBuildMessage lEmpNo, nPrftCtr, False
            nPrevPrftCtr = nPrftCtr
        End If
        
        subSearchHoursForEmpInPrft lEmpNo, nPrftCtr, dDate, dRegHours, dOTHours
        
        If Not fnWriteToHoldArray(rsTemp, dRegHours, dOTHours) Then
            Exit Function
        End If
        
        rsTemp.MoveNext
    Loop
    
    subBuildMessage lEmpNo, nPrftCtr, True
    fnParseOneEmployee = True
    Exit Function
ErrorHandler:
    subWriteLog Err.Description
    fnParseOneEmployee = False
End Function

Private Sub subBuildReportHeader()
    Dim sMsg As String
    
    sMsg = Space(65) & "Week 1" & Space(21) & "Week 2"
    subWriteLog sMsg
    
    sMsg = Space(51) & String(6, "-") & "Before" & String(9, "-") & "After" & String(9, "-")
    sMsg = sMsg & "Before" & String(7, "-") & "After" & String(6, "-")
    subWriteLog sMsg
    
    sMsg = "Employee # Employee Name                 Profit Center   Hours   Reg Hours OT Hours   Hours Reg Hours OT Hours"
    subWriteLog sMsg
    subWriteLog String(110, "-")
End Sub

Private Sub subSearchHoursForEmpInPrft(ByVal lEmpNo As Long, ByVal nPrftCtr As Integer, _
            ByVal dDate As Date, ByRef dRegHours As Single, ByRef dOTHours As Single)

    Dim i As Long
    Dim nWeek As Integer
    
    dRegHours = 0
    dOTHours = 0
    
    If DateDiff("d", dDate, CDate(txtOTWeek1EndDate)) >= 0 Then
        nWeek = 1
    Else
        nWeek = 2
    End If
    
    For i = 0 To UBound(g_ArrayHours)
    
        With g_ArrayHours(i)
            
            If .EmpNO = lEmpNo And .PrftCtr = nPrftCtr And .nWeek = nWeek Then
                dRegHours = .RegHours
                dOTHours = .OTHOurs
            End If
            
        End With
        
    Next
    
End Sub

Private Function subBuildMessage(lEmpNo As Long, nPrftCtr As Integer, bEmpTotal As Boolean)
    Dim i As Long
    Dim sLine As String
    Dim sEmpName As String
    Dim dTotalRegHours1 As Single
    Dim dTotalRegHours2 As Single
    Dim dAfterRegHours1 As Single
    Dim dAfterOTHours1 As Single
    Dim dAfterRegHours2 As Single
    Dim dAfterOTHours2 As Single
    Dim dEmployeeTotal1 As Single
    Dim dEmployeeRegHours1 As Single
    Dim dEmployeeOTHours1 As Single
    Dim dEmployeeTotal2 As Single
    Dim dEmployeeRegHours2 As Single
    Dim dEmployeeOTHours2 As Single
    
    For i = 0 To UBound(g_ArrayHours)
        
        With g_ArrayHours(i)
            
            If .EmpNO = lEmpNo Then
            
                If .PrftCtr = nPrftCtr Then
                
                    If .nWeek = 1 Then
                        dTotalRegHours1 = .OrigHours
                        dAfterRegHours1 = .RegHours
                        dAfterOTHours1 = .OTHOurs
                    Else
                        dTotalRegHours2 = .OrigHours
                        dAfterRegHours2 = .RegHours
                        dAfterOTHours2 = .OTHOurs
                    End If
                    
                End If
                        
                If .nWeek = 1 Then
                    dEmployeeTotal1 = dEmployeeTotal1 + .OrigHours
                    dEmployeeRegHours1 = dEmployeeRegHours1 + .RegHours
                    dEmployeeOTHours1 = dEmployeeOTHours1 + .OTHOurs
                Else
                    dEmployeeTotal2 = dEmployeeTotal2 + .OrigHours
                    dEmployeeRegHours2 = dEmployeeRegHours2 + .RegHours
                    dEmployeeOTHours2 = dEmployeeOTHours2 + .OTHOurs
                End If
                
            End If
            
        End With
        
    Next
    
    If Not bEmpTotal Then
        
        If dTotalRegHours1 > 0# Or dTotalRegHours2 > 0# Then
        
            If g_lHoldEmpNO <> lEmpNo Then
                sLine = CStr(lEmpNo) & Space(11 - Len(CStr(lEmpNo)))
                sEmpName = fnGetEmployeeName(lEmpNo)
                sLine = sLine & Trim(Mid(sEmpName, 1, 30)) & Space(30 - Len(Trim(Mid(sEmpName, 1, 30))))
                g_lHoldEmpNO = lEmpNo
            Else
                sLine = Space(41)
            End If
            
            sLine = sLine & CStr(nPrftCtr) & Space(16 - Len(CStr(nPrftCtr)))
            
            If dTotalRegHours1 > 0# Then
                sLine = sLine & CStr(dTotalRegHours1) & Space(8 - Len(CStr(dTotalRegHours1)))
                sLine = sLine & CStr(dAfterRegHours1) & Space(10 - Len(CStr(dAfterRegHours1)))
                sLine = sLine & CStr(dAfterOTHours1) & Space(11 - Len(CStr(dAfterOTHours1)))
            Else
                sLine = sLine & Space(29)
            End If
            
            If dTotalRegHours2 > 0# Then
                sLine = sLine & CStr(dTotalRegHours2) & Space(9 - Len(CStr(dTotalRegHours2)))
                sLine = sLine & CStr(dAfterRegHours2) & Space(10 - Len(CStr(dAfterRegHours2)))
                sLine = sLine & CStr(dAfterOTHours2) & Space(11 - Len(CStr(dAfterOTHours2)))
            End If
            
        End If
        
        If dEmployeeTotal1 > 40# Or dEmployeeTotal2 > 40# Then
            g_bOverTimeRecordFound = True
            subWriteLog sLine
        End If
        
    Else
        If dEmployeeTotal1 > 0 Then
            sLine = Space(57) & String(6, "-") & Space(2) & String(16, "-")
        Else
            sLine = Space(81)
        End If
        
        If dEmployeeTotal2 > 0 Then
            sLine = sLine & Space(5) & String(6, "-") & Space(2) & String(16, "-")
        End If
        
        If dEmployeeTotal1 > 40 Or dEmployeeTotal2 > 40 Then
            subWriteLog sLine
        End If
        
        If dEmployeeTotal1 > 0 Then
            sLine = Space(57) & CStr(dEmployeeTotal1) & Space(8 - Len(CStr(dEmployeeTotal1))) & CStr(dEmployeeRegHours1) & Space(10 - Len(CStr(dEmployeeRegHours1))) & CStr(dEmployeeOTHours1) & Space(10 - Len(CStr(dEmployeeOTHours1)))
        Else
            sLine = Space(85)
        End If
        
        If dEmployeeTotal2 > 0 Then
            sLine = sLine & Space(1) & CStr(dEmployeeTotal2) & Space(9 - Len(CStr(dEmployeeTotal2))) & CStr(dEmployeeRegHours2) & Space(10 - Len(CStr(dEmployeeRegHours2))) & CStr(dEmployeeOTHours2)
        End If
        
        If dEmployeeTotal1 > 40 Or dEmployeeTotal2 > 40 Then
            g_bOverTimeRecordFound = True
            subWriteLog sLine
            subWriteLog ""
        End If
        
    End If
    
End Function

Private Function fnGetEmployeeName(lEmpNo As Long) As String
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    strSQL = "SELECT TRIM(prm_last_name) || ', ' || TRIM(prm_first_name) AS empname "
    strSQL = strSQL & " FROM pr_master  WHERE prm_empno = " & lEmpNo
    
    If GetRecordSet(rsTemp, strSQL, nDB_REMOTE, "fnGetEmployeeName") > 0 Then
        fnGetEmployeeName = fnGetField(rsTemp!empName)
    End If
    
    
End Function

'Get the Total Amount for emploee for the period and for the Profit Center
Private Function fnBuildEmployeeHoursArray() As Boolean
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim dTotalEmpHours As Single
    Dim lCount As Long
    Dim i As Long
    Dim nOldEmpNo As Long
    Dim dEmpRegHours As Double
    Dim dEmpOTHours As Double
    
    'Build array for week 1
    tfnSetStatusBarCorrect "Building the employee hours information for week 1..."
    
    strSQL = "SELECT bh_empno, bh_prft_ctr, SUM(bh_hours) as Hours FROM bonus_hold"
    strSQL = strSQL & " WHERE bh_chk_link =  0"
    'strSQL = strSQL & " AND bh_pay_code = " & tfnSQLString(sHOURLYPAYCODE)
    strSQL = strSQL & " AND bh_pay_code IN " & sHourPayCodeList
    strSQL = strSQL & " AND bh_date BETWEEN " & tfnDateString(txtOTWeek1BeginDate, True)
    strSQL = strSQL & " AND " & tfnSQLString(txtOTWeek1EndDate, True)
    strSQL = strSQL & " AND bh_empno IN(SELECT b2.bh_empno FROM bonus_hold b2 "
    strSQL = strSQL & " WHERE bh_date BETWEEN " & tfnDateString(txtOTWeek1BeginDate, True)
    strSQL = strSQL & " AND " & tfnSQLString(txtOTWeek1EndDate, True)
    strSQL = strSQL & " AND bh_chk_link =  0"
    'strSQL = strSQL & " AND bh_pay_code = " & tfnSQLString(sHOURLYPAYCODE)
    strSQL = strSQL & " AND bh_pay_code IN " & sHourPayCodeList
    strSQL = strSQL & " GROUP BY b2.bh_empno "
    strSQL = strSQL & " HAVING SUM(b2.bh_hours) > 40)"
    strSQL = strSQL & " GROUP BY bh_empno, bh_prft_ctr"
    strSQL = strSQL & " ORDER BY bh_empno, bh_prft_ctr"
    
    If GetRecordSet(rsTemp, strSQL, nDB_REMOTE, "fnBuildEmployeeHoursArray") < 0 Then
        fnBuildEmployeeHoursArray = False
        Exit Function
    ElseIf rsTemp.RecordCount > 0 Then
        
        subSetProgress 0.001
        
        ReDim g_ArrayHours(rsTemp.RecordCount - 1)
        nOldEmpNo = -1
        dEmpOTHours = 0#
        dEmpRegHours = 0#
        
        Do Until rsTemp.EOF
            
            If nOldEmpNo <> tfnRound(rsTemp!bh_empno) Then
                nOldEmpNo = tfnRound(rsTemp!bh_empno)
            End If
                
            With g_ArrayHours(lCount)
                .EmpNO = tfnRound(rsTemp!bh_empno)
                .nWeek = 1
                .PrftCtr = tfnRound(rsTemp!bh_prft_ctr)
                .OrigHours = tfnRound(rsTemp!Hours, 2)
                
                .Parsed = False
                dTotalEmpHours = fnGetTotalHoursForEmployee(.EmpNO, .nWeek)
                
                If dTotalEmpHours = 0# Then
                    .RegHours = 0
                    .OTHOurs = 0
                ElseIf dTotalEmpHours <= MAXWEEKHOURS Then
                    .RegHours = .OrigHours
                    .OTHOurs = 0
                Else
                    'If this is last record for this employee, make the total regular hours equal to the MAxWEEKHOURS
                    'Junsong 12/06/2002 call 391714
                    If fnIsLastRecordForEmployee(rsTemp, nOldEmpNo) Then
                        .RegHours = tfnRound(MAXWEEKHOURS - dEmpRegHours, 2)
                        .OTHOurs = tfnRound((dTotalEmpHours - MAXWEEKHOURS - dEmpOTHours), 2)
                        dEmpOTHours = 0#
                        dEmpRegHours = 0#
                    Else
                        .RegHours = tfnRound(MAXWEEKHOURS * (.OrigHours / dTotalEmpHours), 2)
                        .OTHOurs = tfnRound((dTotalEmpHours - MAXWEEKHOURS) * (.OrigHours / dTotalEmpHours), 2)
                        dEmpOTHours = tfnRound(dEmpOTHours + .OTHOurs, 2)
                        dEmpRegHours = tfnRound(dEmpRegHours + .RegHours, 2)
                    End If
                    
                End If
                
            End With
                
            lCount = lCount + 1
            i = i + 1
            subSetProgress (i / rsTemp.RecordCount) * 100
            rsTemp.MoveNext
        Loop
        
    End If
    
    
    If Trim(txtOTWeek2BeginDate) <> "" Then
        'Build array for week 2
        tfnSetStatusBarCorrect "Building the employee hours information for week 2..."
    
        strSQL = "SELECT bh_empno, bh_prft_ctr, SUM(bh_hours) as Hours FROM bonus_hold"
        strSQL = strSQL & " WHERE bh_chk_link =  0"
'        strSQL = strSQL & " AND bh_pay_code = " & tfnSQLString(sHOURLYPAYCODE)
        strSQL = strSQL & " AND bh_pay_code IN " & sHourPayCodeList
        strSQL = strSQL & " AND bh_date BETWEEN " & tfnDateString(txtOTWeek2BeginDate, True)
        strSQL = strSQL & " AND " & tfnSQLString(txtOTWeek2EndDate, True)
        strSQL = strSQL & " AND bh_empno IN(SELECT b2.bh_empno FROM bonus_hold b2 "
        strSQL = strSQL & " WHERE bh_date BETWEEN " & tfnDateString(txtOTWeek2BeginDate, True)
        strSQL = strSQL & " AND " & tfnSQLString(txtOTWeek2EndDate, True)
        strSQL = strSQL & " AND bh_chk_link =  0"
        'strSQL = strSQL & " AND bh_pay_code = " & tfnSQLString(sHOURLYPAYCODE)
        strSQL = strSQL & " AND bh_pay_code IN " & sHourPayCodeList
        strSQL = strSQL & " GROUP BY b2.bh_empno "
        strSQL = strSQL & " HAVING SUM(b2.bh_hours) > 40)"
        strSQL = strSQL & " GROUP BY bh_empno, bh_prft_ctr"
        strSQL = strSQL & " ORDER BY bh_empno, bh_prft_ctr"
        
        If GetRecordSet(rsTemp, strSQL, nDB_REMOTE, "fnBuildEmployeeHoursArray") > 0 Then
    
            If lCount > 0 Then
                ReDim Preserve g_ArrayHours(UBound(g_ArrayHours) + rsTemp.RecordCount)
            Else
                ReDim g_ArrayHours(rsTemp.RecordCount - 1)
            End If
            
            i = 0
            nOldEmpNo = -1
            dEmpOTHours = 0#
            dEmpRegHours = 0#
        
            Do Until rsTemp.EOF
                    
                If nOldEmpNo <> tfnRound(rsTemp!bh_empno) Then
                    nOldEmpNo = tfnRound(rsTemp!bh_empno)
                End If

                With g_ArrayHours(lCount)
                    .EmpNO = tfnRound(rsTemp!bh_empno)
                    .nWeek = 2
                    .Parsed = False
                    .PrftCtr = tfnRound(rsTemp!bh_prft_ctr)
                    .OrigHours = tfnRound(rsTemp!Hours, 2)
                    dTotalEmpHours = fnGetTotalHoursForEmployee(.EmpNO, .nWeek)
                
                    If dTotalEmpHours = 0# Then
                        .RegHours = 0
                        .OTHOurs = 0
                    ElseIf dTotalEmpHours <= MAXWEEKHOURS Then
                        .RegHours = .OrigHours
                        .OTHOurs = 0
                    Else
                        'If this is last record for this employee, make the total regular hours equal to the MAxWEEKHOURS
                        'Junsong 12/06/2002 call 391714
                        If fnIsLastRecordForEmployee(rsTemp, nOldEmpNo) Then
                            .RegHours = tfnRound(MAXWEEKHOURS - dEmpRegHours, 2)
                            .OTHOurs = tfnRound((dTotalEmpHours - MAXWEEKHOURS - dEmpOTHours), 2)
                            dEmpOTHours = 0#
                            dEmpRegHours = 0#
                        Else
                            .RegHours = tfnRound(MAXWEEKHOURS * (.OrigHours / dTotalEmpHours), 2)
                            .OTHOurs = tfnRound((dTotalEmpHours - MAXWEEKHOURS) * (.OrigHours / dTotalEmpHours), 2)
                            dEmpOTHours = tfnRound(dEmpOTHours + .OTHOurs, 2)
                            dEmpRegHours = tfnRound(dEmpRegHours + .RegHours, 2)
                        End If
                    
                    End If
                    
                End With
            
                lCount = lCount + 1
                i = i + 1
                subSetProgress (i / rsTemp.RecordCount) * 100
                rsTemp.MoveNext
            Loop
        
        End If
    
    End If
        
    fnBuildEmployeeHoursArray = (lCount >= 1)
End Function

Private Function fnGetTotalHoursForEmployee(lEmpNo As Long, nWeek As Integer) As Single
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    strSQL = "SELECT SUM(bh_hours) as Hours FROM bonus_hold"
    strSQL = strSQL & " WHERE bh_empno = " & lEmpNo
    strSQL = strSQL & " AND bh_chk_link =  0"
    'strSQL = strSQL & " AND bh_pay_code = " & tfnSQLString(sHOURLYPAYCODE)
    strSQL = strSQL & " AND bh_pay_code IN " & sHourPayCodeList
    
    If nWeek = 1 Then
        strSQL = strSQL & " AND bh_date BETWEEN " & tfnDateString(txtOTWeek1BeginDate, True)
        strSQL = strSQL & " AND " & tfnSQLString(txtOTWeek1EndDate, True)
    Else
        strSQL = strSQL & " AND bh_date BETWEEN " & tfnDateString(txtOTWeek2BeginDate, True)
        strSQL = strSQL & " AND " & tfnSQLString(txtOTWeek2EndDate, True)
    End If
    
    If GetRecordSet(rsTemp, strSQL, nDB_REMOTE, "fnGetTotalHoursForEmployee") >= 0 Then
        fnGetTotalHoursForEmployee = tfnRound(rsTemp!Hours, 2)
    End If
    
End Function

Private Function fnIsLastRecordForEmployee(rsTemp As Recordset, nOldEmpNo As Long) As Boolean
    'On Error Resume Next
    
    If rsTemp.EOF Then
        fnIsLastRecordForEmployee = True
    Else
        rsTemp.MoveNext
            
        If rsTemp.EOF Then
            fnIsLastRecordForEmployee = True
        Else
        
            If tfnRound(rsTemp!bh_empno) = nOldEmpNo Then
                fnIsLastRecordForEmployee = False
            Else
                fnIsLastRecordForEmployee = True
            End If
            
        End If
        
        rsTemp.MovePrevious
    
    End If
    
End Function

'This funciton calculate the hours for employee and hold the record need to be inserted, updated and deleted later
'The Rule is: search the date the accumurated hours is greater than distributed regular hours, found then
'Update that date and remove all the records occured later than that date.

Private Function fnWriteToHoldArray(ByVal rsTemp As Recordset, dRegHours As Single, dOTHours As Single) As Boolean
    Dim dLeftHours As Single
    Dim dWeekHoursAfterAdd As Single
    Dim dWeekHoursbeforeAdd As Single
    Dim bOverTime As Boolean
    Dim dOverDate As Date
    Dim dDate As Date
    Dim dHours As Single
    Dim lEmpNo As Long
    Dim nPrftCtr As Integer
    Dim rsClone As Recordset
    Dim sPayCode As String
    Dim sOverPayCode As String
    On Error GoTo EXITHERE
    
    dDate = rsTemp!bh_date
    lEmpNo = tfnRound(rsTemp!bh_empno)
    nPrftCtr = tfnRound(rsTemp!bh_prft_ctr)
    dHours = tfnRound(rsTemp!bh_Hours, 2)
    sPayCode = Trim(rsTemp!bh_pay_code & "")
    
    'if both is 0, means no overtime, don't change them
    If dRegHours = 0# And dOTHours = 0# Then
        fnWriteToHoldArray = True
        subWriteToBonusArray dDate, lEmpNo, nPrftCtr, TypeHold, dHours, sPayCode
        Exit Function
    End If
    
    If txtOTWeek2BeginDate <> "" And DateDiff("d", dDate, CDate(txtOTWeek1EndDate)) < 0 Then
        dWeekHoursbeforeAdd = g_dWeek2EmployeeHours
        g_dWeek2EmployeeHours = g_dWeek2EmployeeHours + tfnRound(dHours, 2)
        dWeekHoursAfterAdd = g_dWeek2EmployeeHours
    Else
        dWeekHoursbeforeAdd = g_dWeek1EmployeeHours
        g_dWeek1EmployeeHours = g_dWeek1EmployeeHours + tfnRound(dHours, 2)
        dWeekHoursAfterAdd = g_dWeek1EmployeeHours
    End If
    
    If dWeekHoursbeforeAdd > dRegHours Then
        subWriteToBonusArray dDate, lEmpNo, nPrftCtr, TypeDelete, 0, sPayCode
        
    ElseIf dWeekHoursAfterAdd > dRegHours Then
        dLeftHours = dRegHours - dWeekHoursbeforeAdd
        subWriteToBonusArray dDate, lEmpNo, nPrftCtr, TypeUpdate, dLeftHours, sPayCode
            
        If dOTHours > 0 Then
            bOverTime = True
        End If
        
    ElseIf dWeekHoursAfterAdd = dRegHours Then
        subWriteToBonusArray dDate, lEmpNo, nPrftCtr, TypeHold, dHours, sPayCode
    ElseIf dWeekHoursAfterAdd < dRegHours Then
        
        If Not fnIsLatestRecordForEmpAndPrft(rsTemp) Then
            subWriteToBonusArray dDate, lEmpNo, nPrftCtr, TypeHold, dHours, sPayCode
        Else
            
            If dOTHours > 0# Then
                bOverTime = True
            End If
            
        End If
        
    End If
    
    If bOverTime Then
        sOverPayCode = fnGetOverTimePayCode(sPayCode)
        
        If sOverPayCode = "" Then
            g_bOverPayCodeNotSet = True
            subWriteLog "Over time pay code is not set up for " & sPayCode
            Exit Function
        End If
        
        If fnIsLatestRecordForEmpAndPrft(rsTemp) Then
            subWriteToBonusArray dDate, lEmpNo, nPrftCtr, TypeInsert, dOTHours, sOverPayCode
        Else
            Dim dHoursOnDateOverReg As Single
            Dim nLevel As Integer
            Dim dTempHours As Single
            Dim nHowManyLeft As Integer
            
            nLevel = 0
            dHoursOnDateOverReg = 0
            
            Set rsClone = rsTemp
            
            'from the latest date to earliest date, fill in with the OT hour
            Do Until dHoursOnDateOverReg = -1
                dOverDate = dDate
                subGetLasestData dHoursOnDateOverReg, dOverDate, rsClone, nLevel, nHowManyLeft
                
                If dOTHours <= 0 Then
                    Exit Do
                End If
                
                If dHoursOnDateOverReg = -1 Then
                    Exit Do
                End If
                    
                'If only one day left then, then add OT Hours to this date.
                If nHowManyLeft = 1 Then
                
                    If dLeftHours < dHours Then
                        dOTHours = dOTHours - dHoursOnDateOverReg
                        dTempHours = dHoursOnDateOverReg
                    Else
                        dTempHours = dOTHours
                        dOTHours = 0
                    End If
                    
                ElseIf dHoursOnDateOverReg >= dOTHours Then
                    dTempHours = dOTHours
                    dOTHours = 0
                Else
                    dOTHours = dOTHours - dHoursOnDateOverReg
                    dTempHours = dHoursOnDateOverReg
                End If
                
                If dTempHours > 0# Then
                    subWriteToBonusArray dOverDate, lEmpNo, nPrftCtr, TypeInsert, dTempHours, sOverPayCode
                End If
                
                nLevel = nLevel + 1
            Loop
        
        End If
        
    End If
    
    fnWriteToHoldArray = True
    Exit Function
EXITHERE:
    subWriteLog Err.Description
    fnWriteToHoldArray = False
End Function

Private Sub subWriteToBonusArray(dDate As Date, lEmpNo As Long, nPrftCtr As Integer, nType As MaintenanceType, dHours As Single, sPayCode As String)
    
    If g_lBonusCount >= UBound(g_ArrayBonus) Then
        ReDim Preserve g_ArrayBonus(UBound(g_ArrayBonus) + INCARRAYZISE)
    End If

    With g_ArrayBonus(g_lBonusCount)
         .bonusDate = dDate
         .EmpNO = lEmpNo
         .PrftCtr = nPrftCtr
         .DoType = nType
         .Hours = tfnRound(dHours, 2)
         .payCode = sPayCode
         'Debug.Print CStr(.bonusDate) & vbTab & CStr(.EmpNO) & vbTab & CStr(.PrftCtr) & vbTab & CStr(.payCode) & vbTab & CStr(.DoType) & vbTab & CStr(.Hours)
    End With

    g_lBonusCount = g_lBonusCount + 1
End Sub

Private Function fnIsLatestRecordForEmpAndPrft(ByVal rsTemp As Recordset) As Boolean
    Dim lEmpNo  As Long
    Dim nPrftCtr As Integer
    Dim dDate As Date
    
    On Error GoTo EXITHERE
    dDate = rsTemp!bh_date
    lEmpNo = tfnRound(rsTemp!bh_empno)
    nPrftCtr = tfnRound(rsTemp!bh_prft_ctr)
    rsTemp.MoveNext
    
    If rsTemp.EOF Then
        rsTemp.MovePrevious
        fnIsLatestRecordForEmpAndPrft = True
        Exit Function
    End If
    
    If lEmpNo = tfnRound(rsTemp!bh_empno) And nPrftCtr = tfnRound(rsTemp!bh_prft_ctr) Then
        
        If DateDiff("d", dDate, CDate(txtOTWeek1EndDate)) >= 0 Then
        
            If DateDiff("d", rsTemp!bh_date, CDate(txtOTWeek1EndDate)) >= 0 Then
                fnIsLatestRecordForEmpAndPrft = False
            End If
        
        Else
        
            If DateDiff("d", rsTemp!bh_date, dDate) < 0 Then
                fnIsLatestRecordForEmpAndPrft = False
            End If
            
        End If
            
    Else
        fnIsLatestRecordForEmpAndPrft = True
    End If
    
    rsTemp.MovePrevious
    Exit Function
EXITHERE:
    Err.Clear
    fnIsLatestRecordForEmpAndPrft = True
End Function

Private Sub subGetLasestData(ByRef dHoursOnDateOverReg As Single, ByRef dOTDate As Date, ByVal rsTemp As Recordset, ByRef nLevel As Integer, ByRef nHowManyLeft As Integer)
    Dim lEmpNo  As Long
    Dim nPrftCtr As Integer
    Dim dDate As Date
    Dim bEnd As Boolean
    Dim i As Integer
    Dim lNext As Long
    Dim lPrevious As Long
    Dim vOrigPos As Variant
    
    dHoursOnDateOverReg = -1
    On Error GoTo EXITHERE
    
    vOrigPos = rsTemp.Bookmark
    
    dDate = rsTemp!bh_date
    lEmpNo = tfnRound(rsTemp!bh_empno)
    nPrftCtr = tfnRound(rsTemp!bh_prft_ctr)
    
    Do Until bEnd
        rsTemp.MoveNext
        lNext = lNext + 1
        
        If rsTemp.EOF Then
            Exit Do
        End If
        
        If lEmpNo <> tfnRound(rsTemp!bh_empno) Or nPrftCtr <> tfnRound(rsTemp!bh_prft_ctr) Then
            bEnd = True
            Exit Do
        End If
        
        If DateDiff("d", dDate, CDate(txtOTWeek1EndDate)) >= 0 Then
        
            If DateDiff("d", rsTemp!bh_date, CDate(txtOTWeek1EndDate)) < 0 Then
                Exit Do
            End If
        
        Else
            'get the latest one
        End If
        
    Loop

    For i = 0 To nLevel
        lPrevious = lPrevious + 1
        rsTemp.MovePrevious
    Next
    
    nHowManyLeft = lNext - lPrevious
    
    dHoursOnDateOverReg = tfnRound(rsTemp!bh_Hours, 2)
    dOTDate = CDate(rsTemp!bh_date)
    rsTemp.Bookmark = vOrigPos
    Exit Sub
    
EXITHERE:
    On Error Resume Next
    Err.Clear
    rsTemp.Bookmark = vOrigPos
    dHoursOnDateOverReg = -1
End Sub

Private Function fnGetOverTimePayCode(sPayCode As String) As String
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    strSQL = "SELECT prpa_tmi_ottype FROM pr_pay WHERE prpa_pay_code = " & tfnSQLString(sPayCode)
    
    If GetRecordSet(rsTemp, strSQL, nDB_REMOTE, "fnGetOverTimePayCode") > 0 Then
        strSQL = "SELECT prpa_pay_code FROM pr_pay WHERE prpa_tmi_type = " & tfnSQLString(fnGetField(rsTemp!prpa_tmi_ottype))
        
        If GetRecordSet(rsTemp, strSQL, nDB_REMOTE, "fnGetOverTimePayCode") = 1 Then
            fnGetOverTimePayCode = fnGetField(rsTemp!prpa_pay_code)
        End If
        
    End If
    
End Function
Public Sub mnuCancel_Click()
    subCancel
End Sub

Private Sub subCancel()

    If t_bDataChanged = True Then

        If tfnCancelExit(t_szCANCEL_MESSAGE) = False Then
            Exit Sub
        End If
        
    End If
    
    tfnResetScreen
    t_nFormMode = IDLE_MODE
End Sub

Public Function fnInValidData(txtBox As Textbox) As Boolean
    Dim sDate As String
    
    sDate = Trim(txtBox.Text)
    
    If Not sDate = "" Then
        
        If Not IsFactorDate(sDate) Then
            cValidate.SetErrorMessage txtBox, "Invalid Date entered"
            fnInValidData = True
            Exit Function
        Else
            sDate = tfnFormatDate(sDate)
        End If
    
    End If
    
    Select Case txtBox.TabIndex
        Case txtOTWeek1BeginDate.TabIndex
                    
            If sDate = "" Then
                cValidate.SetErrorMessage txtBox, "The Begin Date of Week 1 is required"
                fnInValidData = True
            Else
                
                If cValidate.ValidInput(txtOTWeek1EndDate) And txtOTWeek1EndDate <> "" Then
                    
                    If DateDiff("d", sDate, txtOTWeek1EndDate) < 0 Then
                        cValidate.SetErrorMessage txtBox, "The Begin Date of week 1 is later than End Date"
                        fnInValidData = True
                    End If
                    
                End If
                
            End If
            
            
            If Not fnInValidData Then
                
                If txtOTWeek1EndDate = "" Then
                    cValidate.SetInputValid txtOTWeek1EndDate, True
                    txtOTWeek1EndDate = CStr(DateAdd("d", 6, CDate(sDate)))
                End If
            
                If txtOTWeek2BeginDate = "" Then
                    cValidate.SetInputValid txtOTWeek2BeginDate, True
                    txtOTWeek2BeginDate = CStr(DateAdd("d", 7, CDate(sDate)))
                End If
                
                If txtOTWeek2EndDate = "" Then
                    cValidate.SetInputValid txtOTWeek2EndDate, True
                    txtOTWeek2EndDate = CStr(DateAdd("d", 13, CDate(sDate)))
                End If
                
            End If
            
        Case txtOTWeek1EndDate.TabIndex
            
            If sDate = "" Then
                cValidate.SetErrorMessage txtBox, "The End Date of Week 1 is required"
                fnInValidData = True
            Else
                
                If cValidate.ValidInput(txtOTWeek1BeginDate) And txtOTWeek1BeginDate <> "" Then
                    
                    If DateDiff("d", sDate, txtOTWeek1BeginDate) > 0 Then
                        cValidate.SetErrorMessage txtBox, "The End Date of week 1 is earlier than Begin Date"
                        fnInValidData = True
                    End If
                    
                End If
                
            End If
            
        Case txtOTWeek2BeginDate.TabIndex
            
            If sDate = "" Then
                
                If Trim(txtOTWeek2EndDate) <> "" And cValidate.ValidInput(txtOTWeek2EndDate) Then
                    cValidate.SetErrorMessage txtBox, "The End date of Week is not empty, Begin Date is required."
                    fnInValidData = True
                ElseIf Trim(txtOTWeek2EndDate) = "" Then
                    cValidate.SetInputValid txtOTWeek2EndDate, True
                End If
                
            Else
            
                If cValidate.ValidInput(txtOTWeek1EndDate) And txtOTWeek1EndDate <> "" Then
                    
                    If DateDiff("d", sDate, txtOTWeek1EndDate) >= 0 Then
                        cValidate.SetErrorMessage txtBox, "The Begin Date of week 2 is earlier than End Date of Week 1"
                        fnInValidData = True
                    End If
                    
                End If
                
                If cValidate.ValidInput(txtOTWeek2EndDate) And txtOTWeek2EndDate <> "" Then
                    
                    If DateDiff("d", sDate, txtOTWeek2EndDate) < 0 Then
                        cValidate.SetErrorMessage txtBox, "The Begin Date of week 2 is later than End Date"
                        fnInValidData = True
                    End If
                    
                End If
                
                If Not fnInValidData Then
                    
                    If Trim(txtOTWeek2EndDate) = "" Then
                        cValidate.SetInputValid txtOTWeek2EndDate, False
                    End If
                    
                End If
                
            End If
    
        Case txtOTWeek2EndDate.TabIndex
        
            If sDate = "" Then
                
                If Trim(txtOTWeek2BeginDate) <> "" And cValidate.ValidInput(txtOTWeek2BeginDate) Then
                    cValidate.SetErrorMessage txtBox, "The Begin date of Week is not empty, End Date is required."
                    fnInValidData = True
                ElseIf Trim(txtOTWeek2BeginDate) = "" Then
                    cValidate.SetInputValid txtOTWeek2BeginDate, True
                End If
                
            Else
                
                If cValidate.ValidInput(txtOTWeek2BeginDate) And txtOTWeek2BeginDate <> "" Then
                    
                    If DateDiff("d", sDate, txtOTWeek2BeginDate) > 0 Then
                        cValidate.SetErrorMessage txtBox, "The End Date of week 2 is earlier than Begin Date"
                        fnInValidData = True
                    End If
                    
                End If
                
                If Not fnInValidData Then
                    
                    If Trim(txtOTWeek2BeginDate) = "" Then
                        cValidate.SetInputValid txtOTWeek2BeginDate, False
                    End If
                    
                End If
                
            End If
            
    End Select
    
    If Not fnInValidData Then
        txtBox.Text = sDate
    End If

End Function

Private Sub subDataChanged()
    cmdOTProcessBtn.Enabled = False
    
    frmContext.ButtonEnabled(CANCEL_UP) = True
    cmdExitCancelBtn.Enabled = True
    mnuCancel.Enabled = True
    
End Sub

Private Sub subCheckButtonStatus()
    
    If cValidate.FirstInvalidInput < 0 Then
        cmdOTProcessBtn.Enabled = True
    Else
        cmdOTProcessBtn.Enabled = False
    End If
    
End Sub

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

Private Sub subWriteLog(sMsg As String)
    lstOTLog.AddItem sMsg
    lstOTLog.Refresh
End Sub

Private Sub subSetProgress(sngPercent As Single)
    
    If sngPercent > 0# Then
        
        If Not pbStatus.Visible Then
            pbStatus.ZOrder 0
            pbStatus.Visible = True
        End If
    
    Else
        pbStatus.Visible = False
    End If
    
    pbStatus.value = sngPercent
End Sub

Private Function fnInsertUpdate() As Boolean
    Dim i As Long
    Dim strSQL As String
    
    
    If g_bOverTimeRecordFound Then
        
        For i = 0 To UBound(g_ArrayBonus)
            
            With g_ArrayBonus(i)
            
                Select Case .DoType
                    
                    Case TypeDelete
                        strSQL = "DELETE FROM bonus_hold"
                        strSQL = strSQL & " WHERE bh_empno = " & tfnRound(.EmpNO)
                        strSQL = strSQL & " AND bh_prft_ctr = " & tfnRound(.PrftCtr)
                        strSQL = strSQL & " AND bh_pay_code = " & tfnSQLString(.payCode)
                        strSQL = strSQL & " AND bh_date = " & tfnDateString(.bonusDate, True)
                    Case TypeHold
                        strSQL = ""
                    Case TypeInsert
                        'Usually this is for overtime paycodem, but the system may already have overtime paycode with hours set.
                        If Not fnExistRecordInSystem(.EmpNO, .PrftCtr, .payCode, .bonusDate) Then
                            strSQL = "INSERT INTO bonus_hold(bh_empno, bh_prft_ctr, bh_date, bh_check_amount,  "
                            strSQL = strSQL & "bh_override, bh_chk_link, bh_pay_code, bh_hours) VALUES ("
                            strSQL = strSQL & tfnRound(.EmpNO) & "," & tfnRound(.PrftCtr) & ","
                            strSQL = strSQL & tfnDateString(.bonusDate, True) & ",0, 'Y', 0,"
                            strSQL = strSQL & tfnSQLString(.payCode) & "," & tfnRound(.Hours, 2) & ")"
                        Else
                            
                            If MsgBox("There is a record with pay code " & .payCode & " for employee " & .EmpNO _
                                    & " in profit Center " & .PrftCtr & " in " & CStr(.bonusDate) _
                                    & ",do you want to override this data?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                                strSQL = "UPDATE bonus_hold SET bh_hours = " & tfnRound(.Hours, 2)
                                strSQL = strSQL & " WHERE bh_empno = " & tfnRound(.EmpNO)
                                strSQL = strSQL & " AND bh_prft_ctr = " & tfnRound(.PrftCtr)
                                strSQL = strSQL & " AND bh_pay_code = " & tfnSQLString(.payCode)
                                strSQL = strSQL & " AND bh_date = " & tfnDateString(.bonusDate, True)
                            Else
                                strSQL = ""
                            End If
                                    
                        End If
                        
                    Case TypeUpdate
                    
                        strSQL = "UPDATE bonus_hold SET bh_hours = " & tfnRound(.Hours, 2)
                        strSQL = strSQL & " WHERE bh_empno = " & tfnRound(.EmpNO)
                        strSQL = strSQL & " AND bh_prft_ctr = " & tfnRound(.PrftCtr)
                        strSQL = strSQL & " AND bh_pay_code = " & tfnSQLString(.payCode)
                        strSQL = strSQL & " AND bh_date = " & tfnDateString(.bonusDate, True)
                    
                End Select
                
                
                If strSQL <> "" Then
                
                    If Not fnExecuteSQL(strSQL) Then
                        fnInsertUpdate = False
                        Exit Function
                    End If
                    
                End If
                
            End With
            
        Next
        
    End If

    fnInsertUpdate = True
End Function

Private Function fnExistRecordInSystem(lEmpNo As Long, nPrftCtr As Integer, sPayCode As String, dDate As Date) As Boolean
    Dim strSQL As String
                      
    strSQL = "SELECT bh_empno FROM bonus_hold "
    strSQL = strSQL & " WHERE bh_empno = " & tfnRound(lEmpNo)
    strSQL = strSQL & " AND bh_prft_ctr = " & tfnRound(nPrftCtr)
    strSQL = strSQL & " AND bh_pay_code = " & tfnSQLString(sPayCode)
    strSQL = strSQL & " AND bh_date = " & tfnDateString(dDate, True)
    
    If GetRecordCount(strSQL) > 0 Then
        fnExistRecordInSystem = True
    Else
        fnExistRecordInSystem = False
    End If

End Function

Private Function fnExecuteSQL(szSQL As String, _
                              Optional nDB As Variant, _
                              Optional szCalledFrom As Variant, _
                              Optional bShowError As Variant) As Boolean
                
      Dim szMsg As String
      
      On Error GoTo SQLError
      
      If IsMissing(nDB) Then
       nDB = nDB_REMOTE
      End If
    
      Select Case nDB
        
        Case nDB_LOCAL
           dbLocal.Execute szSQL
        Case nDB_REMOTE
           t_dbMainDatabase.ExecuteSQL szSQL
      End Select
      
      fnExecuteSQL = True
      Exit Function
      
SQLError:
      fnExecuteSQL = False
      If IsMissing(szCalledFrom) Then
         szCalledFrom = ""
      End If
      If IsMissing(bShowError) Then
         bShowError = True
      End If
      tfnErrHandler "fnExecuteSQL, " & szCalledFrom, szSQL, bShowError
      On Error GoTo 0
End Function

'david 03/11/2003  #403275
'validate ot paycode for the employee
'return error message
Private Function fnValidOtPaycode(lEmpNo As Long, sHourPayCodeList As String, bError As Boolean) As String
    Dim strSQL As String
    Dim aryPayCodes() As String
    Dim i As Integer
    Dim j As Integer
    Dim sOverPayCode As String
    Dim lCount As Long
    Dim sInvalidOtPayCode As String
    Dim nYear(4) As Integer
    
    'get the year
    On Error Resume Next
    nYear(1) = Year(tfnFormatDate(txtOTWeek1BeginDate))
    nYear(2) = Year(tfnFormatDate(txtOTWeek1EndDate))
    
    If nYear(2) = nYear(1) Then
        nYear(2) = 0
    End If
    
    If Trim(txtOTWeek2BeginDate) <> "" Then
        nYear(3) = Year(tfnFormatDate(txtOTWeek2BeginDate))
        nYear(4) = Year(tfnFormatDate(txtOTWeek2EndDate))
    
        If nYear(4) = nYear(3) Then
            nYear(4) = 0
        Else
            If nYear(4) = nYear(1) Or nYear(4) = nYear(2) Then
                nYear(4) = 0
            End If
        End If
        
        If nYear(3) = nYear(1) Or nYear(3) = nYear(2) Then
            nYear(3) = 0
        End If
    End If
    
    On Error GoTo 0
    On Error GoTo errTrap
    
    'constant
    'sHourPayCodeList = "('7HRP', '7HRA', '7HRD', '7HRI', '7HRO')"
    aryPayCodes = Split(Mid(sHourPayCodeList, 2, Len(sHourPayCodeList) - 2), ",")
    bError = False
    
    For i = 0 To UBound(aryPayCodes)
        If aryPayCodes(i) <> "" Then
            sOverPayCode = fnGetOverTimePayCode(Mid(aryPayCodes(i), 2, Len(aryPayCodes(i)) - 2))
            
            If sOverPayCode <> "" Then
                For j = 1 To 4
                    If nYear(j) > 0 Then
                        strSQL = "SELECT  * FROM " & frmZZSEBPRC.objHours.sTempPrPay
                        strSQL = strSQL & " WHERE pcode = " & tfnSQLString(sOverPayCode)
                        strSQL = strSQL & " AND pempno = " & lEmpNo
                        strSQL = strSQL & " AND pYear = " & nYear(j)
                        
                        lCount = GetRecordCount(strSQL)
                        
                        If lCount < 0 Then
                            fnValidOtPaycode = "Failed to access database when checking overtime paycode."
                            bError = True
                            Exit Function
                        ElseIf lCount = 0 Then
                            If sInvalidOtPayCode <> "" Then
                                sInvalidOtPayCode = sInvalidOtPayCode + ","
                            End If
                            
                            sInvalidOtPayCode = sInvalidOtPayCode + sOverPayCode
                        End If
                    End If
                Next j
            End If
        End If
    Next i

    If sInvalidOtPayCode <> "" Then
        fnValidOtPaycode = "Pay Code (" & sInvalidOtPayCode & ") not set for employee#" & lEmpNo
    End If
    
    Exit Function
    
errTrap:
    fnValidOtPaycode = Err.Number & " - " & Err.Description
    tfnErrHandler "", False
    bError = True
End Function
