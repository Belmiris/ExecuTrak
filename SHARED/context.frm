VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmContext 
   Caption         =   "Toolbar Kit"
   ClientHeight    =   1488
   ClientLeft      =   696
   ClientTop       =   5388
   ClientWidth     =   2160
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1488
   ScaleWidth      =   2160
   Visible         =   0   'False
   Begin VB.PictureBox pctStatusbar 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   288
      ScaleHeight     =   240
      ScaleWidth      =   1596
      TabIndex        =   0
      Top             =   912
      Width           =   1596
      Begin VB.Label lblStatusbar 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   156
         Left            =   48
         TabIndex        =   1
         Top             =   36
         Width           =   1164
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   324
      Top             =   132
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   16777215
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Menu mnuDropDown 
      Caption         =   "Dropnown"
      Begin VB.Menu mnuContextItems 
         Caption         =   "Any"
         Index           =   0
      End
      Begin VB.Menu mnuContextItems 
         Caption         =   "Any"
         Index           =   1
      End
      Begin VB.Menu mnuContextSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "&Cancel"
      End
   End
End
Attribute VB_Name = "frmContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private Const t_szOLETBKit As String = "TBKIT.clsToolbar"
    Private Const TBKIT_DLL_PATH = "OLE\TBKIT.DLL"
    
    Private Const CONTROL_TB = 0
    Private Const CONTROL_TB_FRAME = 1
    Private Const CONTROL_TB_RIGHT = 2
    Private Const CONTROL_TB_PANEL = 3
    Private Const DELAY_FOR_START = 10
    
    Private objToolbar As Object
    Private frmMainForm As Form
    Private sLastMessage As String
    Private m_nMenuItems As Integer

    'Api definitions
    Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
    End Type
    Private Type POINTAPI
        x As Long
        y As Long
    End Type

    Private Declare Function GetMenu Lib "user32" ( _
        ByVal hwnd As Long) As Long
    
    Private Declare Function GetCursorPos Lib "user32" ( _
        lpPoint As POINTAPI) As Long

    Private Declare Function GetSubMenu Lib "user32" ( _
        ByVal hMenu As Long, _
        ByVal nPos As Long) As Long

    Private Declare Function TrackPopupMenu Lib "user32" ( _
        ByVal hMenu As Long, _
        ByVal wFlags As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal nReserved As Long, _
        ByVal hwnd As Long, _
        lprc As RECT) As Long

    Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long


    'david 10/30/00
    Private sngMainFormHeight As Single
    Private sngMainFormWidth As Single

'
'Function : ShowPopup - shows a context menu
'Variables: pointer to the form, submenu index
'Return   : none
'
Private Sub fnShowPopup()

    Const nSubMenu = 0
    
    ' This is where to put menu
    Dim nWhereX, nWhereY As Single
    
    'Top Level Menu
    Dim hMenu As Integer
    
    'Sub Menu to popup
    Dim hSubMenu As Integer
    
    'Window rectangle to show popup in
    Dim rctMainWindow As RECT
    
    'Pos of Current Mouse Pointer
    Dim pntPosition As POINTAPI
    
    'A TEmp variable to hold the return value from TrackPopup
    Dim nTemp As Integer
    
    'Get the Mouse pointer position
    Call GetCursorPos(pntPosition)
    
    'Set up the variables to hold the screen size
    'Screen .ScaleMode = TWIPS
    rctMainWindow.Top = 0
    rctMainWindow.Left = 0
    rctMainWindow.Right = Screen.Width \ Screen.TwipsPerPixelX
    rctMainWindow.Bottom = Screen.Height \ Screen.TwipsPerPixelY
    
    ' Don't Put the menu right under the mousepointer
    nWhereX = pntPosition.x + 3
    nWhereY = pntPosition.y + 5
    
    'Get the top level menu
    hMenu = GetMenu(frmContext.hwnd)
    
    'Get the submenu that they want
    hSubMenu = GetSubMenu(hMenu, nSubMenu)
    
    'Popup the menu
    nTemp = TrackPopupMenu(hSubMenu, 2, nWhereX, nWhereY, 0, frmContext.hwnd, rctMainWindow)

End Sub


Public Function AddButton(sCaption As String, _
                          ByVal nUpKey As Integer, _
                          Optional vAddToMenu As Variant, _
                          Optional vCallBack As Variant) As Boolean
    If Not objToolbar Is Nothing Then
        AddButton = objToolbar.AddButton(sCaption, nUpKey, vAddToMenu, vCallBack)
        subCheckError
    End If
End Function

Public Sub AddSeparator(Optional vAddToTB As Variant, _
                        Optional vAddToMenu As Variant)
    If Not objToolbar Is Nothing Then
        objToolbar.AddSeparator
    End If
End Sub

Public Sub BeginSetupTBMainMenu(frmTemp As Object, _
                                ParamArray objControls())
    Dim I As Integer
    Dim bInitSet As Boolean
'    Set objToolbar = New clsToolbar
    Set frmMainForm = frmTemp
    bInitSet = True
    On Error GoTo errSetup
    Set objToolbar = CreateObject(t_szOLETBKit)
    With objToolbar
        Set .ContextForm = Me
        Set .FactorOLE = t_oleObject
        Set .TBImageList = ImageList1
'        .SetMenu mnuContextItems
        #If FACTOR_MENU < 0 Then
            .FactorMenu = True
        #Else
            .FactorMenu = False
        #End If
        For I = 0 To UBound(objControls)
            If Not IsMissing(objControls(I)) Then
                Select Case I
                    Case CONTROL_TB
                        Set objToolbar.tbToolbar = objControls(I)
                    Case CONTROL_TB_FRAME
                        Set objToolbar.fraToolbar = objControls(I)
                    Case CONTROL_TB_RIGHT
                        Set objToolbar.fraToolbarRight = objControls(I)
                    Case CONTROL_TB_PANEL
                        Set objToolbar.fraPanel = objControls(I)
                End Select
            End If
        Next I
        If objToolbar.tbToolbar Is Nothing Then
            Set objToolbar.tbToolbar = frmTemp.tbToolbar
        End If
        If objToolbar.fraToolbar Is Nothing Then
            Set objToolbar.fraToolbar = frmTemp.efraToolBar
        End If
        .BeginSetupTBMainMenu frmTemp
    End With
    subCheckError
    
    Exit Sub
errSetup:
    If Err.Number Then
        If bInitSet Then
            If frmMainForm.RegisterDll(TbkitDllPath, False) Then
                bInitSet = False
                Resume
            End If
        End If
    End If
End Sub

Public Sub BeginSetupToolbar(frmTemp As Form, _
                             ParamArray objControls())
    Dim I As Integer

'    Set objToolbar = New clsToolbar
    Set frmMainForm = frmTemp
    
    sngMainFormHeight = frmMainForm.Height
    sngMainFormWidth = frmMainForm.Width
    
    On Error GoTo errSetup1
    Set objToolbar = CreateObject(t_szOLETBKit)
    With objToolbar
        Set .ContextForm = Me
        Set .FactorOLE = t_oleObject
        Set .TBImageList = ImageList1
        #If FACTOR_MENU < 0 Then
            .FactorMenu = True
        #Else
            .FactorMenu = False
        #End If
        For I = 0 To UBound(objControls)
            If Not IsMissing(objControls(I)) Then
                Select Case I
                    Case CONTROL_TB
                        Set objToolbar.tbToolbar = objControls(I)
                    Case CONTROL_TB_FRAME
                        Set objToolbar.fraToolbar = objControls(I)
                    Case CONTROL_TB_RIGHT
                        Set objToolbar.fraToolbarRight = objControls(I)
                    Case CONTROL_TB_PANEL
                        Set objToolbar.fraPanel = objControls(I)
                End Select
            End If
        Next I
        If objToolbar.fraToolbar Is Nothing Then
            Set objToolbar.fraToolbar = frmTemp.efraToolBar
        End If
        If objToolbar.tbToolbar Is Nothing Then
            Set objToolbar.tbToolbar = frmTemp.tbToolbar
        End If
        On Error Resume Next
        If objToolbar.fraToolbarRight Is Nothing Then
            Set objToolbar.fraToolbarRight = frmTemp.efraToolBarRight
        End If
        If objToolbar.fraPanel Is Nothing Then
            Set objToolbar.fraPanel = frmTemp.efraToolBar.FMName
        End If
        .BeginSetupToolbar frmTemp
    End With
    subCheckError
    Exit Sub
errSetup1:
    If Err.Number = 429 Then
        MsgBox "Toolbar OLE is not registered properly. Please contact Factor."
    End If
End Sub


Public Sub ButtonClick(Button As Button)
    If Not objToolbar Is Nothing Then
        subShowBusyState True, Button.Key
        objToolbar.ButtonClick Button.Key
        tfnWaitSeconds DELAY_FOR_START
        subShowBusyState False, Button.Key
        subCheckError
    End If
End Sub

Property Get ButtonEnabled(ByVal nID As Integer) As Boolean
    If Not objToolbar Is Nothing Then
        On Error Resume Next
        ButtonEnabled = objToolbar.ButtonEnabled(nID)
    End If
End Property

Property Get ButtonVisible(ByVal nID As Integer) As Boolean
    If Not objToolbar Is Nothing Then
        On Error Resume Next
        ButtonVisible = objToolbar.ButtonVisible(nID)
    End If
End Property


Property Let ButtonEnabled(ByVal nID As Integer, _
                           ByVal bFlag As Boolean)
    If Not objToolbar Is Nothing Then
        On Error Resume Next
        objToolbar.ButtonEnabled(nID) = bFlag
    End If
End Property

Property Let ButtonVisible(ByVal nID As Integer, _
                           ByVal bFlag As Boolean)
    If Not objToolbar Is Nothing Then
        On Error Resume Next
        objToolbar.ButtonVisible(nID) = bFlag
    End If
End Property

Public Sub EndSetupToolbar()
    Dim nTHeight As Integer
    
    If Not objToolbar Is Nothing Then
        On Error Resume Next
'        Set objToolbar.StatusBar = frmMainForm.sbStatusBar
        objToolbar.EndSetupToolbar
'        SetParent pctStatusbar.hWnd, frmMainForm.sbStatusBar.hWnd
'        pctStatusbar.Move 0, 0, frmMainForm.sbStatusBar.Panels(1).Width, frmMainForm.sbStatusBar.Height
'        nTHeight = pctStatusbar.TextHeight("A")
'        ffraStatusbar.Move Screen.TwipsPerPixelX * 2, (pctStatusbar.Height - nTHeight) / 2, pctStatusbar.Width, nTHeight
        subCheckError
    End If
    m_nMenuItems = 1
End Sub


Property Let HelpFile(sFile As String)
    If Not objToolbar Is Nothing Then
        objToolbar.HelpFile = sFile
    End If
End Property


Public Sub ShowSBMessage(sMsg As String)
    On Error Resume Next
    frmMainForm.ffraStatusbar.ForeColor = STANDARD_TEXT_COLOR
    frmMainForm.ffraStatusbar.Font.Bold = False
    frmMainForm.ffraStatusbar.Caption = sMsg
    frmMainForm.ffraStatusbar.Refresh
End Sub

Public Sub ShowSBError(sMsg As String)
    On Error Resume Next
    frmMainForm.ffraStatusbar.ForeColor = ERROR_TEXT_COLOR
    frmMainForm.ffraStatusbar.Font.Bold = True
    frmMainForm.ffraStatusbar.Caption = sMsg
    frmMainForm.ffraStatusbar.Refresh
End Sub

Public Sub ShowSBRight(sMsg As String)
    On Error Resume Next
    frmMainForm.ffraStatusbar.Font.Bold = False
    frmMainForm.ffraStatusbar.ForeColor = &H8000&  'Green text
    frmMainForm.ffraStatusbar.Caption = sMsg
    frmMainForm.ffraStatusbar.Refresh
End Sub

Public Sub FormResize()
    On Error Resume Next
    
    frmMainForm.Height = sngMainFormHeight
    frmMainForm.Width = sngMainFormWidth
    
    If Not objToolbar Is Nothing Then
        objToolbar.Resize
    End If
End Sub

Public Function LoadPicture(ByVal nID As Integer) As Object
    If Not objToolbar Is Nothing Then
        Set LoadPicture = objToolbar.LoadPicture(nID)
    End If
End Function

Public Sub MenuClick(ByVal nIdx As Integer)
    If Not objToolbar Is Nothing Then
        subShowBusyState True, nIdx
        objToolbar.MenuClick nIdx
        tfnWaitSeconds DELAY_FOR_START
        subShowBusyState False, nIdx
        subCheckError
    End If
End Sub


Public Sub MouseDown(ByVal Button As Integer, _
                     ByVal nKey As Integer, _
                     ParamArray vExKeys() As Variant)
    If objToolbar Is Nothing Then
        Exit Sub
'        objToolbar.MouseDown Button, nKey, vExKey
    End If
    If Button <> vbRightButton Then
        Exit Sub
    End If
    Const DEFAULT_MENU = 0
    Const ADDITIONAL_MENU = 1
    
    Dim sCap As String
    Dim sTag As String
    Dim bEnabled As Boolean
    Dim I As Integer
    
    objToolbar.GetMenuInfo sCap, sTag, bEnabled, nKey
    mnuContextItems(DEFAULT_MENU).Caption = sCap
    mnuContextItems(DEFAULT_MENU).Tag = sTag
    mnuContextItems(DEFAULT_MENU).Enabled = bEnabled
    mnuContextItems(ADDITIONAL_MENU).Visible = False
    If Not IsMissing(vExKeys) Then
        For I = 0 To UBound(vExKeys)
        If IsNumeric(vExKeys(I)) Then
            If I + 1 > m_nMenuItems Then
                Load mnuContextItems(I + 1)
                m_nMenuItems = m_nMenuItems + 1
            End If
            objToolbar.GetMenuInfo sCap, sTag, bEnabled, Val(vExKeys(I))
            mnuContextItems(I + 1).Caption = sCap
            mnuContextItems(I + 1).Visible = True
            mnuContextItems(I + 1).Tag = sTag
            mnuContextItems(I + 1).Enabled = bEnabled
        End If
        Next I
    End If

    fnShowPopup
    
End Sub


Public Function RunItem(ByVal nID As Integer) As Integer
    If Not objToolbar Is Nothing Then
        subShowBusyState True, nID
        objToolbar.RunItem nID
        subShowBusyState False, nID
        subCheckError
        If objToolbar.ErrorCode = 0 Then
            RunItem = True
        Else
            RunItem = False
        End If
    End If
End Function

Public Function RunProgram(sProgram As String) As Boolean
    #If FACTOR_MENU < 0 Then
        If Not t_oleObject Is Nothing Then
            subShowBusyState True, VENDOR_UP
            RunProgram = t_oleObject.RunExe(sProgram)
            tfnWaitSeconds DELAY_FOR_START
            subShowBusyState False, VENDOR_UP
            subCheckError
        End If
    #Else
        MsgBox "This program must run from Factor Maim Menu", vbOKOnly + vbCritical, App.Title
    #End If
End Function

Private Sub subCheckError()
    Dim nCode As Integer
    
    If Not objToolbar Is Nothing Then
        nCode = objToolbar.ErrorCode
        If nCode <> 0 Then
            #If DEVELOP Then
                MsgBox objToolbar.ErrorMessage(nCode), vbExclamation + vbOK
            #Else
                #If FACTOR_MENU < 0 Then
                    MsgBox objToolbar.ErrorMessage(nCode), vbExclamation + vbOK
                #Else
                    'Must run from factor menu
                    MsgBox objToolbar.ErrorMessage(nCode), vbExclamation + vbOK
                #End If
            #End If
        End If
    End If
End Sub

Private Sub subShowBusyState(bFlag As Boolean, _
                             Vkey As Variant)
    
    On Error Resume Next
    If bFlag Then
        Screen.MousePointer = vbHourglass
        sLastMessage = frmMainForm.ffraStatusbar.Caption
        If objToolbar.IsModule(Vkey) Then
            frmMainForm.tfnSetStatusBarMessage "Launching program. Please wait . . ."
        End If
    Else
        If objToolbar.IsModule(Vkey) Then
            frmMainForm.tfnSetStatusBarMessage sLastMessage
        End If
        Screen.MousePointer = vbDefault
    End If
    
End Sub

Public Function TbkitDllPath() As String
    If Val(App.Minor) < 20 Then
        TbkitDllPath = TBKIT_DLL_PATH
    Else
        TbkitDllPath = fnGetFactorPath + "\" + TBKIT_DLL_PATH
    End If
End Function

Public Sub TBMouseMove()
    If Not objToolbar Is Nothing Then
        objToolbar.TBMouseMove
    End If
End Sub

Private Sub Form_Unload(CANCEL As Integer)
    Set objToolbar = Nothing
    Set frmMainForm = Nothing
End Sub


Private Sub mnuContextItems_Click(Index As Integer)
    If Not objToolbar Is Nothing Then
        subShowBusyState True, Index
'        objToolbar.ContextMenuClick Index
        objToolbar.CMenuClick mnuContextItems(Index).Tag
        subShowBusyState False, Index
        subCheckError
    End If
End Sub

'david 10/12/00
Private Function fnGetFactorPath() As String
    Dim sTemp As String
    Dim nPosi As Integer
    
    sTemp = UCase(App.Path)
    
    nPosi = InStrRev(sTemp, "\")
    
    If nPosi > 0 Then
        fnGetFactorPath = Left(sTemp, nPosi - 1)
    Else
        fnGetFactorPath = ""
    End If
End Function

