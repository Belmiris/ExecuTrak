VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cValidateInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***********************************************************'
'
' Copyright (c) 1996 FACTOR, A Division of W.R.Hess Company
'
' Module name   : CVALID.CLS
' Date          : Feb. 23, 1996
' Programmer(s) : Qinggang Ma
'
' This module defines a class used for input validation.

Option Explicit

    Private Const TYPE_BACK_ONCE = 1
    Private Const TYPE_BACK_ALWAYS = 2
    Private Const TYPE_NO_ESCAPE = 3

    Private Const STATUS_ERROR_FREE = 0
    Private Const STATUS_ERRORS = 1
    Private Const STATUS_ERROR_WARNING = 2

    Private Const FOCUS_GOT = 1
    Private Const FOCUS_LOST = 2
    
    Private Const ESC_LEVEL_0 = -1
    Private Const ESC_LEVEL_1 = 0
    Private Const ESC_LEVEL_2 = 1
    
    Private Const FLAG_TEXT_NOCHANGE = 0
    Private Const FLAG_TEXT_CHANGED = 1
    
    Private Const STANDARD_TEXT_COLOR = &H0&
    Private Const ERROR_TEXT_COLOR As Long = &HFF&

    Private Type tpFocusControl
        m_txtEdit       As Control
        m_nType         As Integer
        m_sPrompt       As String
        m_sErrorMsgs    As String
        m_nStatusFlag   As Integer
        m_nInputFlag    As Integer
        m_nTextChanged  As Integer
        m_nChangeFlag   As Integer
        m_bValidMessage As Boolean
    End Type
    Private Type tpESCapable
        m_objControl As Object
        m_nLevel As Integer
    End Type

    Private frmParent As Form
    Private lblStatus As Control
    
    'Variables for focus control
    Private txtLastAsscBox As Textbox
    Private txtLastBox As Textbox
    Private ctrlForFocus As Object
    Private ctrlFirstControls() As Object
    Private nFirstControls As Integer
    Private bDefaultValidMessage As Boolean
    Private tpControls() As tpFocusControl
    Private bLocked As Boolean
    Private nCount As Integer
    
    Private arryESCControls() As tpESCapable

    Private nESCCount As Integer
    Private bNeedBack As Boolean
    Private nPreviousControl As Integer
    Private nLastFocus As Integer
    Private bDisplayMessage As Boolean
    Private nESCLevelDefault As Integer
    Private nMinTabIndex As Integer
    Private nMaxTabIndex As Integer
    Private bResetFlag As Boolean
    Private bShowBusyMessage As Boolean

    Private nFlagFocus As Integer

    Private nValidating As Integer
Public Function BoxInList(txtBox As Object) As Boolean
    Dim nNum As Integer
    nNum = fnBoxNumber(txtBox.TabIndex)
    If nNum >= 0 Then
        BoxInList = True
    Else
        BoxInList = False
    End If
End Function


Public Function CheckFocus() As Integer
    'Return focus status
    CheckFocus = nLastFocus
    nLastFocus = -1

End Function


Property Set ControlForFocus(ctrl As Control)
    Set ctrlForFocus = ctrl
    If nMaxTabIndex < ctrl.TabIndex Then
        nMaxTabIndex = ctrl.TabIndex
    End If
    If nMinTabIndex > ctrl.TabIndex - 1 Then
        nMinTabIndex = ctrl.TabIndex - 1
    End If
End Property

Property Get ControlForFocus() As Control
    Set ControlForFocus = ctrlForFocus
End Property

Public Property Get DataChanged(txtBox As Object) As Boolean
    Dim nNum As Integer
    nNum = fnBoxNumber(txtBox.TabIndex)
    If nNum >= 0 Then
        If tpControls(nNum).m_nChangeFlag = FLAG_TEXT_CHANGED Then
            DataChanged = True
        Else
            DataChanged = False
        End If
    End If
End Property

Public Property Let DataChanged(txtBox As Object, bTemp As Boolean)
    Dim nNum As Integer
    nNum = fnBoxNumber(txtBox.TabIndex)
    If nNum >= 0 Then
        If bTemp Then
            tpControls(nNum).m_nChangeFlag = FLAG_TEXT_CHANGED
        Else
            tpControls(nNum).m_nChangeFlag = FLAG_TEXT_NOCHANGE
        End If
    End If
End Property

Public Function IsESCControl(Optional ctrl As Variant) As Boolean
    IsESCControl = False
    If IsMissing(ctrl) Then
        If frmParent.ActiveControl.TabIndex <= nMinTabIndex Then
            IsESCControl = True
        End If
    Else
        If ctrl.TabIndex <= nMinTabIndex Then
            IsESCControl = True
        End If
    End If
End Function

Property Get ESCLevel1() As Integer
    ESCLevel1 = ESC_LEVEL_1
End Property
Property Get ESCLevel2() As Integer
    ESCLevel2 = ESC_LEVEL_2
End Property

Property Let GreenMessageDefault(bTemp As Boolean)
    bDefaultValidMessage = bTemp
End Property

Property Let DisplayMessage(bFlag As Boolean)
    bDisplayMessage = bFlag
End Property

Property Get DisplayMessage() As Boolean
    DisplayMessage = bDisplayMessage
End Property

Property Let ESCLevel(nTemp As Integer)
    nESCLevelDefault = nTemp
End Property

Private Function fnGetFirstControl() As Object

    Dim i As Integer
    
    Set fnGetFirstControl = Nothing
    For i = 0 To nFirstControls
        With ctrlFirstControls(i)
            If .Enabled And .Visible Then
                Set fnGetFirstControl = ctrlFirstControls(i)
                Exit For
            End If
        End With
    Next

End Function

Property Let GreenMessage(txtBox As Textbox, bTemp As Boolean)
    Dim nIndex As Integer
    nIndex = fnBoxNumber(txtBox.TabIndex)
    tpControls(nIndex).m_bValidMessage = bTemp
End Property

Property Get LastBox() As Textbox
    If nPreviousControl >= 0 And nPreviousControl < nCount Then
        Set LastBox = tpControls(nPreviousControl).m_txtEdit
    Else
        Set LastBox = txtLastBox
    End If
End Property

Property Let Locked(bStatus As Boolean)
    bLocked = bStatus
End Property

Property Get Locked() As Boolean
    Locked = bLocked
End Property


'---------------------------------------------------------------------------------------
' Procedure : LostFocus
' DateTime  :
' Author    : DenBorg
' Modified  : 11/21/2005 16:14
'
' Purpose   : This function's logic has been moved to LostFocus2(), and any required
'             modifications needs to be done in LostFocus2(). It still does exactly
'             the same thing, but uses a Collection object instead of a ParamArray.
'
'             This existing function was left in to allow existing code to work as
'             designed. It simply builds a collection from the information in the
'             ParamArray, and then calls LostFocus2().
'
'             This was done to make writing Factor Combo Box classes possible
'             (a piece of code cannot dynamically change the number of parameters
'              passes to a function call)
'---------------------------------------------------------------------------------------
'
Public Function LostFocus(txtBox As Object, ParamArray arryControls()) As Boolean
    Dim colControls As Collection
    Dim Item        As Long
    
    Set colControls = New Collection
    For Item = 0 To UBound(arryControls)
        colControls.Add arryControls(Item)
    Next 'Item
    
    LostFocus = LostFocus2(txtBox, colControls)
End Function
Property Let MinTabIndex(nTemp As Integer)
    nMinTabIndex = nTemp
End Property

Property Let MaxTabIndex(nTemp As Integer)
    nMaxTabIndex = nTemp
End Property

Public Sub SetChangeFlags(Optional vFlag As Variant)
    Dim nNum As Integer
    Dim bFlag As Boolean

    If IsMissing(vFlag) Then
        bFlag = True
    Else
        bFlag = vFlag
    End If
    For nNum = 0 To nCount - 1
        If bFlag Then
            tpControls(nNum).m_nChangeFlag = FLAG_TEXT_CHANGED
        Else
            tpControls(nNum).m_nChangeFlag = FLAG_TEXT_NOCHANGE
        End If
    Next
End Sub

Public Sub SetESCLevel(oCtrl As Object, nLevel As Integer)
    Dim i As Integer
    For i = 0 To nESCCount - 1
        If frmParent.ActiveControl Is arryESCControls(i).m_objControl Then
            arryESCControls(i).m_nLevel = nLevel
            Exit For
        End If
    Next i
End Sub

Public Sub SetFirstControls(ParamArray arryTemp() As Variant)
    'The order in which the first control gets focus depends on the last variable.
    ' If the last variable is a boolean indicating whether the input order is the
    ' order of getting focus. A value of true means it is.
    'The default value is false.
    
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim bFlag As Boolean

    nFirstControls = UBound(arryTemp)
    If IsObject(arryTemp(nFirstControls)) Then
        bFlag = True
    Else
        bFlag = arryTemp(nFirstControls)
        nFirstControls = nFirstControls - 1
    End If
    ReDim ctrlFirstControls(nFirstControls)
    If bFlag Then
        For i = 0 To nFirstControls
            For j = 0 To i - 1
                If ctrlFirstControls(j).TabIndex > arryTemp(i).TabIndex Then
                    Exit For
                End If
            Next j
            For k = i - 1 To j Step -1
                Set ctrlFirstControls(k + 1) = ctrlFirstControls(k)
            Next k
            Set ctrlFirstControls(j) = arryTemp(i)
        Next i
    Else
        If nFirstControls >= 0 Then
            ReDim ctrlFirstControls(nFirstControls)
            For i = 0 To nFirstControls
                Set ctrlFirstControls(i) = arryTemp(i)
            Next i
        End If
    End If
End Sub

Private Function fnValidate(ByVal nNum As Integer) As Boolean
    Dim bRet As Boolean
    Dim nCount As Integer
    Dim sMsg As String
    Dim lColor As Long
    Dim bBold As Boolean
    
#If FACTOR_MENU = 1 Then
    fnValidate = True
    Exit Function
#End If
    
    'david 12/20/99
    'comment out DoEvents
    'fixed the validation issue in IQFFCPMT
    'Ma cannot tell weather this will affect the validation of other program or not.
    
    nCount = 0
    While nValidating >= 0 And nCount < 10
        nCount = nCount + 1
        'DoEvents
    Wend
    
    If nValidating >= 0 Then
        'It is validating
        If nValidating = nNum Then
            'Assume wrong, because we are validating the currrent record
            tpControls(nNum).m_nStatusFlag = STATUS_ERRORS
        Else
            'Return the flag, because we are validating
            fnValidate = tpControls(nNum).m_nStatusFlag
        End If
        Exit Function
    End If
    If Not bResetFlag Then
        If bShowBusyMessage Then
            sMsg = lblStatus.Caption
            lColor = lblStatus.ForeColor
            bBold = lblStatus.Font.Bold
            subShowMessage "Validating entry. Please wait . . ."
        End If
    End If
    nValidating = nNum
    With tpControls(nNum)
        If .m_nTextChanged = FLAG_TEXT_CHANGED Then
            On Error GoTo errTrap
            bRet = frmParent.fnInvalidData(.m_txtEdit)
            If bRet Then
                .m_nStatusFlag = STATUS_ERROR_WARNING
            Else
                .m_nStatusFlag = STATUS_ERROR_FREE
            End If
            .m_nTextChanged = FLAG_TEXT_NOCHANGE
        End If
    End With
    If Not bResetFlag Then
        If bShowBusyMessage Then
            lblStatus.Caption = sMsg
            lblStatus.ForeColor = lColor
            lblStatus.Font.Bold = bBold
            lblStatus.Refresh
        End If
    End If
    nValidating = -1
    fnValidate = bRet
    
    Exit Function
    
errTrap:
    MsgBox "An Error has occurred in Validation Class." + vbCrLf + vbCrLf + "Error Code: " & Err.Number _
        & vbCrLf & "Error Desc: " & Err.Description + vbCrLf, vbExclamation
    nValidating = -1
    fnValidate = bRet
End Function

Property Let ShowBusyMessage(bFlag As Boolean)
    bShowBusyMessage = bFlag
End Property

Private Sub subShowMessage(sMsg As String)
'    frmContext.ShowSBMessage sMsg
'Exit Sub
    lblStatus.ForeColor = STANDARD_TEXT_COLOR
    lblStatus.Font.Bold = False
    lblStatus.Caption = sMsg
    lblStatus.Refresh
End Sub

Public Function FocusIsLost(txtBox As Object, _
                            ParamArray arryControls()) As Boolean
    Dim i As Integer

    If frmParent.ActiveControl Is Nothing Then
        Exit Function
    End If
    If frmParent.ActiveControl.TabIndex = txtBox.TabIndex Then
        FocusIsLost = False
        Exit Function
    End If
    For i = 0 To UBound(arryControls)
        If frmParent.ActiveControl.TabIndex = arryControls(i).TabIndex Then
            FocusIsLost = False
            Exit Function
        End If
    Next
    FocusIsLost = True
End Function

Property Set LastBox(txtBox As Textbox)
    Set txtLastAsscBox = txtBox
    Set txtLastBox = txtBox
End Property
Property Get NotEscapable() As Integer
    NotEscapable = TYPE_NO_ESCAPE
End Property

Public Sub SetInputValid(txtBox As Object, bStatus As Boolean)
    Dim nBox As Integer
    nBox = fnBoxNumber(txtBox.TabIndex)
    If nBox >= 0 Then
        With tpControls(nBox)
            .m_nTextChanged = FLAG_TEXT_NOCHANGE
            .m_nInputFlag = 1
            If bStatus Then
                .m_nStatusFlag = STATUS_ERROR_FREE
            Else
                .m_nStatusFlag = STATUS_ERROR_WARNING
            End If
        End With
    End If
End Sub

Public Sub ShowErrorMessage(txtBox As Object)
    Dim nNum As Integer
    Dim nFlg As Integer
    Dim nErr As Integer
    
    nNum = fnBoxNumber(txtBox.TabIndex)
    If nNum >= 0 Then
        With tpControls(nNum)
            nFlg = .m_nInputFlag
            .m_nInputFlag = 1
            nErr = .m_nStatusFlag
            If .m_nStatusFlag <> STATUS_ERROR_FREE Then
                .m_nStatusFlag = STATUS_ERROR_WARNING
                Beep
            End If
            fnShowStatus nNum
            .m_nInputFlag = nFlg
            .m_nStatusFlag = nErr
        End With
    End If
End Sub

Property Get Textbox(nIndex As Integer) As Control
    If nIndex >= 0 And nIndex < nCount Then
        Set Textbox = tpControls(nIndex).m_txtEdit
    Else
        Set Textbox = Nothing
    End If
End Property

Public Sub Change(txtBox As Object)
    Dim nNum As Integer
    If Not bLocked Then
        nNum = fnBoxNumber(txtBox.TabIndex)
        If nNum >= 0 Then
            tpControls(nNum).m_nTextChanged = FLAG_TEXT_CHANGED
            tpControls(nNum).m_nChangeFlag = FLAG_TEXT_CHANGED
            Exit Sub
        End If
    End If
End Sub

Property Let ESCControl(oCtrl As Object)
    ReDim Preserve arryESCControls(nESCCount)
    Set arryESCControls(nESCCount).m_objControl = oCtrl
    nESCCount = nESCCount + 1
End Property

Public Function FirstChanged() As Integer
    Dim i As Integer
    FirstChanged = -1
    
    For i = 0 To nCount - 1
        If tpControls(i).m_nInputFlag > 0 Then
            FirstChanged = i
            Exit For
        End If
    Next i
End Function

Private Function fnEscapable() As Integer
    Dim i As Integer
    On Error Resume Next
    fnEscapable = ESC_LEVEL_0
    For i = 0 To nESCCount - 1
        If frmParent.ActiveControl.TabIndex = arryESCControls(i).m_objControl.TabIndex Then
            fnEscapable = arryESCControls(i).m_nLevel
            Exit For
        End If
    Next i
    On Error GoTo 0
End Function

Public Function FirstInvalidInput() As Integer
    Dim i As Integer
    FirstInvalidInput = -1
    
    For i = 0 To nCount - 1
        fnValidate i
        If tpControls(i).m_nStatusFlag > STATUS_ERROR_FREE Then
            FirstInvalidInput = i
            Exit For
        End If
    Next i
End Function

Public Sub ResetFlags(Optional txtBox, Optional vInputFlag)
    Dim i As Integer
    If IsMissing(txtBox) Then
        If frmParent.ActiveControl Is Nothing Then
            Exit Sub
        End If
        If frmParent.ActiveControl.Enabled Then
            On Error Resume Next
            frmParent.ActiveControl.SetFocus
            DoEvents
            On Error GoTo 0
        End If
        bResetFlag = True
        For i = 0 To nCount - 1
            With tpControls(i)
                .m_nChangeFlag = FLAG_TEXT_NOCHANGE
                .m_nTextChanged = FLAG_TEXT_CHANGED
                .m_sErrorMsgs = ""
                fnValidate i
                .m_nInputFlag = 0
            End With
        Next i
        bResetFlag = False
        nPreviousControl = 0
        If txtLastBox Is Nothing Then
            nFlagFocus = FOCUS_LOST
        Else
            If frmParent.ActiveControl.TabIndex <> txtLastBox.TabIndex Then
                nFlagFocus = FOCUS_LOST
            End If
        End If
        bNeedBack = False
    Else
        i = fnBoxNumber(txtBox.TabIndex)
        If i >= 0 Then
            With tpControls(i)
                .m_nTextChanged = FLAG_TEXT_CHANGED
                fnValidate i
                If IsMissing(vInputFlag) Then
                    .m_nInputFlag = 0
                Else
                    If vInputFlag Then
                        .m_nInputFlag = 0
                    End If
                End If
                .m_sErrorMsgs = ""
                .m_nChangeFlag = FLAG_TEXT_NOCHANGE
            End With
        End If
        nValidating = -1
    End If
    
    nLastFocus = -1
End Sub

Public Sub AddEditBox(Box As Object, _
                      sProm As String, _
                      Optional nVType As Variant)
                      
    ReDim Preserve tpControls(nCount)
    
    With tpControls(nCount)
        Set .m_txtEdit = Box
        .m_sPrompt = sProm
        .m_sErrorMsgs = ""
        .m_bValidMessage = bDefaultValidMessage
        If IsMissing(nVType) Then
            .m_nType = TYPE_BACK_ONCE
        Else
            .m_nType = nVType
        End If
        .m_nTextChanged = FLAG_TEXT_CHANGED
        If nMinTabIndex > Box.TabIndex - 1 Then
            nMinTabIndex = Box.TabIndex - 1
        End If
        If nMaxTabIndex < Box.TabIndex + 1 Then
            nMaxTabIndex = Box.TabIndex + 1
        End If
    End With
    nCount = nCount + 1
End Sub

Private Function fnBoxNumber(nIndex As Integer) As Integer
    Dim i
    fnBoxNumber = -1
    For i = 0 To nCount - 1
        If tpControls(i).m_txtEdit.TabIndex = nIndex Then
            fnBoxNumber = i
            Exit For
        End If
    Next i
End Function

Private Sub fnShowStatus(nControlNumber As Integer)
    Dim sAddMessage As String
    With tpControls(nControlNumber)
        If .m_nStatusFlag > STATUS_ERRORS And .m_nInputFlag > 0 Then
            sAddMessage = "Invalid input; " & .m_sPrompt
            lblStatus.ForeColor = ERROR_TEXT_COLOR
            lblStatus.Font.Bold = True
            If .m_sErrorMsgs <> "" Then
                sAddMessage = .m_sErrorMsgs
            End If
'            frmContext.ShowSBError sAddMessage
        Else
            If .m_nInputFlag > 0 And .m_nStatusFlag = STATUS_ERROR_FREE And .m_bValidMessage Then
                lblStatus.Font.Bold = False
                lblStatus.ForeColor = &H8000&  'Green text
                sAddMessage = "Data is valid"
'                frmContext.ShowSBRight sAddMessage
            Else
                sAddMessage = .m_sPrompt
                lblStatus.ForeColor = STANDARD_TEXT_COLOR
                lblStatus.Font.Bold = False
'                frmContext.ShowSBMessage sAddMessage
            End If
        End If
    End With
    lblStatus.Caption = sAddMessage
    lblStatus.Refresh
End Sub

Property Set Form(fForm As Object)
    Set frmParent = fForm
End Property

Public Sub GotFocus(txtBox As Object)
    Dim Index As Integer
    
    If bLocked Then
        If txtBox.TabIndex = ctrlForFocus.TabIndex Then
            SendKeys "{TAB}"
        Else
            'david 05/16/2002
            'need to show some message on the status bar
            lblStatus.ForeColor = STANDARD_TEXT_COLOR
            lblStatus.Font.Bold = False
            lblStatus.Caption = "This field cannot be changed"
            lblStatus.Refresh
        End If
        
        Exit Sub
    End If
    
    If Not ctrlForFocus Is Nothing Then
        If txtBox.TabIndex = ctrlForFocus.TabIndex Then
            Dim ctrlFirst As Object
            Dim txtLast As Textbox
            
            If txtLastAsscBox Is Nothing Then
                Set txtLast = tpControls(nCount - 1).m_txtEdit
            Else
                Set txtLast = txtLastAsscBox
            End If
            
            Set ctrlFirst = fnGetFirstControl
            
            If ctrlFirst Is Nothing Then
                If nESCCount > 0 Then
                    'DenBorg - 12/06/2005
                    'Modified it to search for the first ESC Control
                    'that is both Visible and Enabled, instead of just
                    'simply setting it to the first ESC Control in the
                    'array arryESCControls
                    For Index = 0 To UBound(arryESCControls)
                        With arryESCControls(0).m_objControl
                            If .Visible And .Enabled Then
                                Set ctrlFirst = arryESCControls(Index).m_objControl
                                Exit For
                            End If
                        End With
                    Next 'Index
                End If
            End If
            
            If nFlagFocus = FOCUS_GOT Then
                If LostFocus(txtLastBox) Then
                    If ctrlFirst Is Nothing Then
                        SendKeys "{TAB}"
                    Else
                        If ctrlFirst.Enabled And ctrlFirst.Visible Then
                            On Error Resume Next
                            ctrlFirst.SetFocus
                            nFlagFocus = FOCUS_LOST
                        End If
                    End If
                End If
            Else
                If Not txtLast Is Nothing Then
                    SendKeys "{TAB}", True
                End If
            End If
            
            Exit Sub
        End If
    End If
    
    If Not txtLastAsscBox Is Nothing Then
        If txtBox.TabIndex = txtLastAsscBox.TabIndex Then
            nFlagFocus = FOCUS_GOT
        End If
    End If
    
    If Not txtLastBox Is Nothing Then
        If txtBox.TabIndex = txtLastBox.TabIndex Then
            nFlagFocus = FOCUS_GOT
        End If
    End If
    
    Dim nNum As Integer
    nNum = fnBoxNumber(txtBox.TabIndex)
    
    If nNum >= 0 Then
        fnValidate nNum
        
        If bDisplayMessage Then
            fnShowStatus nNum
        Else
            bDisplayMessage = True
        End If
        
        nFlagFocus = FOCUS_GOT
    End If
End Sub

Public Sub Keypress(txtBox As Object, nKeyAscii As Integer, Optional bMsg As Variant)
    Dim nNum As Integer
    If Not bLocked Then
        nNum = fnBoxNumber(txtBox.TabIndex)
        If nNum >= 0 Then
            If nKeyAscii = vbKeyReturn Then
                fnValidate nNum
    '            fnShowStatus nNum
            Else
                tpControls(nNum).m_nInputFlag = 1
            End If
        End If
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : LostFocus2
' DateTime  : 11/21/2005 16:26
' Author    : DenBorg
'
' Purpose   : This function is the LostFocus() function, except that it uses a
'             Collection instead of a ParamArray. It still does exactly the same thing,
'             but uses a Collection object instead of a ParamArray.
'
'             This existing LostFocus() function was left in to allow existing code to
'             work as designed. It simply builds a collection from the information in
'             the ParamArray, and then calls LostFocus2().
'
'             This was done to make writing Factor Combo Box classes possible
'             (a piece of code cannot dynamically change the number of parameters
'              passes to a function call)
'---------------------------------------------------------------------------------------
'
Public Function LostFocus2(txtBox As Object, Optional colControls As Collection) As Boolean
    Dim nNum As Integer
    Dim bBack As Boolean
    Dim i As Integer
    Dim nEscapable As Integer
    
    If colControls Is Nothing Then
        Set colControls = New Collection
    End If
    
    LostFocus2 = True
    If txtBox Is Nothing Or bLocked Then
        Exit Function
    End If

    DoEvents
    
    If frmParent.ActiveControl Is Nothing Then
        Exit Function
    End If
    
    If TypeOf frmParent.ActiveControl Is Textbox Then
        If frmParent.ActiveControl.TabIndex < txtBox.TabIndex Then
            nFlagFocus = FOCUS_LOST
            Exit Function
        End If
    End If

    If frmParent.ActiveControl.TabIndex = txtBox.TabIndex Then
        LostFocus2 = False
        Exit Function
    End If
    
'    If Not txtLastBox Is txtBox Then
'        If txtLastAsscBox Is Nothing Then
'            Set txtLastBox = tpControls(0).m_txtEdit
'        Else
'            If txtLastAsscBox.TabIndex = txtBox.TabIndex Then
'                Set txtLastBox = txtBox
'            Else
'                For nNum = 0 To UBound(arryControls)
'                    If txtLastAsscBox.TabIndex = arryControls(nNum).TabIndex Then
'                        Set txtLastBox = txtBox
'                    End If
'                Next nNum
'            End If
'        End If
'    End If
    For i = 1 To colControls.Count  'UBound(arryControls)
        If frmParent.ActiveControl.TabIndex = colControls(i).TabIndex Then
            LostFocus2 = False
            nFlagFocus = FOCUS_GOT
            Exit Function
        End If
    Next
            
    If frmParent.ActiveControl.TabIndex <= nMinTabIndex Or frmParent.ActiveControl.TabIndex > nMaxTabIndex Then
'        If tpControls(nNum).m_nType <> TYPE_NO_ESCAPE Then
            LostFocus2 = True
            nFlagFocus = FOCUS_LOST
            Exit Function
'        End If
    End If
    
    nNum = fnBoxNumber(txtBox.TabIndex)
    If nNum >= 0 Then
        If bNeedBack Then
            bNeedBack = False
            If nPreviousControl = nNum Then
                'bNeedBack = False
            Else
                nFlagFocus = FOCUS_GOT
                Exit Function
            End If
        End If
        If Not ctrlForFocus Is Nothing Then
            If Not TypeOf frmParent.ActiveControl Is Textbox Then
                If frmParent.ActiveControl Is ctrlForFocus Then
                    nFlagFocus = FOCUS_GOT
                Else
                    nFlagFocus = FOCUS_LOST
                End If
            End If
        End If
        nEscapable = fnEscapable
        If nEscapable = ESC_LEVEL_2 Then
            If tpControls(nNum).m_nType <> TYPE_NO_ESCAPE Then
                Exit Function
            End If
        End If

        fnValidate nNum
        With tpControls(nNum)
            If .m_nStatusFlag > STATUS_ERROR_FREE Then
                .m_nStatusFlag = STATUS_ERROR_WARNING
            End If
            If .m_nStatusFlag > STATUS_ERROR_FREE Then
                If .m_nInputFlag <= 1 Or .m_nType = TYPE_BACK_ALWAYS Then
                    .m_nInputFlag = 2
                    bNeedBack = True
                    If nEscapable <> ESC_LEVEL_0 Then
                        Exit Function
                    End If
                    If .m_txtEdit.Enabled Then
                        Beep
                        On Error Resume Next
                        .m_txtEdit.SetFocus
                        .m_txtEdit.SelStart = 0
                        .m_txtEdit.SelLength = Len(.m_txtEdit.Text)
                        nLastFocus = nNum
                    End If
                    LostFocus2 = False
                Else
                    bNeedBack = False
                    nLastFocus = -1
                    Exit Function
                End If
            End If
        End With
        nPreviousControl = nNum
    End If
    Exit Function
    
rsmNext:
    Resume Next
End Function
Property Get BackAlways() As Integer
    BackAlways = TYPE_BACK_ALWAYS
End Property

Property Get BackOnce() As Integer
    BackOnce = TYPE_BACK_ONCE
End Property
Public Sub SetErrorMessage(oBox As Object, Msg As String)
    Dim nBox As Integer
    nBox = fnBoxNumber(oBox.TabIndex)
    If nBox >= 0 Then
        tpControls(nBox).m_sErrorMsgs = Msg
    End If
End Sub

Public Sub SetPrompt(oBox As Object, Msg As String)
    Dim nBox As Integer
    nBox = fnBoxNumber(oBox.TabIndex)
    If nBox >= 0 Then
        tpControls(nBox).m_sPrompt = Msg
    End If
End Sub

Public Sub SetFocus(nCtrl As Integer)
    Const nTrialNumber As Integer = 1
    Dim nTrail As Integer

    nTrail = 0
    
    If nCtrl >= 0 And nCtrl <= nCount Then
        On Error GoTo errSetFocus
        tpControls(nCtrl).m_txtEdit.SetFocus
        On Error GoTo 0
    End If
    Exit Sub
    
errSetFocus:
    If nTrail < nTrialNumber Then
        nTrail = nTrail + 1
        DoEvents
        Resume
    End If
    
End Sub

Property Set StatusBar(lblBar As Object)
    Set lblStatus = lblBar
End Property

Public Function Validating() As Boolean
    If nValidating >= 0 Then
        Validating = True
    Else
        Validating = False
    End If
End Function

Property Get ValidInput(txtBox As Object) As Boolean
    Dim nNum As Integer
    nNum = fnBoxNumber(txtBox.TabIndex)
    If nNum >= 0 Then
        fnValidate nNum
        If tpControls(nNum).m_nStatusFlag > STATUS_ERROR_FREE Then
            ValidInput = False
        Else
            ValidInput = True
        End If
    Else
        ValidInput = True
    End If

End Property

Property Let ValidInput(txtBox As Object, bStatus As Boolean)
    Dim nNum As Integer
    nNum = fnBoxNumber(txtBox.TabIndex)

    If nNum >= 0 Then
        With tpControls(nNum)
            If bStatus Then
                .m_nStatusFlag = STATUS_ERROR_FREE
            Else
                .m_nStatusFlag = STATUS_ERRORS
            End If
            .m_nTextChanged = FLAG_TEXT_NOCHANGE
        End With
    End If
End Property

Private Sub Class_Initialize()
    nCount = 0
    nESCCount = 0
    nESCLevelDefault = ESC_LEVEL_1
    bNeedBack = False
    nPreviousControl = 1000
    bDisplayMessage = True
    bDefaultValidMessage = True
    nMinTabIndex = 1000
    nMaxTabIndex = 0
    nFirstControls = -1
    nValidating = -1
    bShowBusyMessage = False
    Set txtLastBox = Nothing
End Sub

Private Sub Class_Terminate()
    Dim i As Integer
    Set frmParent = Nothing
    Set lblStatus = Nothing
    For i = 0 To nCount - 1
        Set tpControls(i).m_txtEdit = Nothing
    Next i
    For i = 0 To nESCCount - 1
        Set arryESCControls(i).m_objControl = Nothing
    Next i
    For i = 0 To nFirstControls - 1
        Set ctrlFirstControls(i) = Nothing
    Next i
    Set ctrlForFocus = Nothing
    Set txtLastBox = Nothing
    Set txtLastAsscBox = Nothing
End Sub

Public Function ErrorMessage(txtBox As Object) As String
    Dim nNum As Integer
    Dim nFlg As Integer
    Dim nErr As Integer
    
    nNum = fnBoxNumber(txtBox.TabIndex)
    If nNum >= 0 Then
        ErrorMessage = fnGetErrorMessage(nNum, True)
    End If
End Function

Private Function fnGetErrorMessage(nControlNumber As Integer, Optional bJustGiveMeTheErrMsg As Boolean = False) As String
    Dim sAddMessage As String
    
    With tpControls(nControlNumber)
        If .m_nStatusFlag > STATUS_ERRORS And (.m_nInputFlag > 0 Or bJustGiveMeTheErrMsg) Then
            sAddMessage = "Data is not valid"
            If .m_sErrorMsgs <> "" Then
                sAddMessage = .m_sErrorMsgs
            End If
        Else
            sAddMessage = ""
        End If
    End With
    
    fnGetErrorMessage = sAddMessage
End Function

