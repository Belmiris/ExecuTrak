VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmContext 
   Caption         =   "Toolbar Kit"
   ClientHeight    =   1485
   ClientLeft      =   690
   ClientTop       =   5385
   ClientWidth     =   2160
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1485
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
      ScaleWidth      =   1590
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
      _ExtentX        =   794
      _ExtentY        =   794
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
'david 08/07/2003
'delay too much
'Private Const DELAY_FOR_START = 10
Private Const DELAY_FOR_START = 3

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
    
Private Type VS_FIXEDFILEINFO
  dwSignature As Long
  dwStrucVersionl As Integer     '  e.g. = &h0000 = 0
  dwStrucVersionh As Integer     '  e.g. = &h0042 = .42
  dwFileVersionMSl As Integer    '  e.g. = &h0003 = 3
  dwFileVersionMSh As Integer    '  e.g. = &h0075 = .75
  dwFileVersionLSl As Integer    '  e.g. = &h0000 = 0
  dwFileVersionLSh As Integer    '  e.g. = &h0031 = .31
  dwProductVersionMSl As Integer '  e.g. = &h0003 = 3
  dwProductVersionMSh As Integer '  e.g. = &h0010 = .1
  dwProductVersionLSl As Integer '  e.g. = &h0000 = 0
  dwProductVersionLSh As Integer '  e.g. = &h0031 = .31
  dwFileFlagsMask As Long        '  = &h3F for version "0.42"
  dwFileFlags As Long            '  e.g. VFF_DEBUG Or VFF_PRERELEASE
  dwFileOS As Long               '  e.g. VOS_DOS_WINDOWS16
  dwFileType As Long             '  e.g. VFT_DRIVER
  dwFileSubtype As Long          '  e.g. VFT2_DRV_KEYBOARD
  dwFileDateMS As Long           '  e.g. 0
  dwFileDateLS As Long           '  e.g. 0
End Type

Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal length As Long)
'

'
'Function : ShowPopup - shows a context menu
'Variables: pointer to the form, submenu index
'Return   : none
'
Private Sub fnShowPopup()

    Const nSubMenu = 0
    
    ' This is where to put menu
    Dim nWhereX, nWhereY As Single
    
    'changed by junsong 02/03/2003 integer to long
    'Get run time error 6 in win 2000
    
    'Top Level Menu
    'Dim hMenu As Integer
    Dim hMenu As Long
    
    'Sub Menu to popup
    'Dim hSubMenu As Integer
    Dim hSubMenu As Long
    
    'Window rectangle to show popup in
    Dim rctMainWindow As RECT
    
    'Pos of Current Mouse Pointer
    Dim pntPosition As POINTAPI
    
    'A TEmp variable to hold the return value from TrackPopup
    'Dim nTemp As Integer
    Dim nTemp As Long
    
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
    Dim i As Integer
    Dim bInitSet As Boolean

    sngMainFormHeight = frmTemp.Height
    sngMainFormWidth = frmTemp.Width

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
        For i = 0 To UBound(objControls)
            If Not IsMissing(objControls(i)) Then
                Select Case i
                    Case CONTROL_TB
                        Set objToolbar.tbToolbar = objControls(i)
                    Case CONTROL_TB_FRAME
                        Set objToolbar.fraToolbar = objControls(i)
                    Case CONTROL_TB_RIGHT
                        Set objToolbar.fraToolbarRight = objControls(i)
                    Case CONTROL_TB_PANEL
                        Set objToolbar.fraPanel = objControls(i)
                End Select
            End If
        Next i
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
    If Err.number Then
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
    Dim i As Integer

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
        For i = 0 To UBound(objControls)
            If Not IsMissing(objControls(i)) Then
                Select Case i
                    Case CONTROL_TB
                        Set objToolbar.tbToolbar = objControls(i)
                    Case CONTROL_TB_FRAME
                        Set objToolbar.fraToolbar = objControls(i)
                    Case CONTROL_TB_RIGHT
                        Set objToolbar.fraToolbarRight = objControls(i)
                    Case CONTROL_TB_PANEL
                        Set objToolbar.fraPanel = objControls(i)
                End Select
            End If
        Next i
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
    If Err.number = 429 Then
        MsgBox "Toolbar OLE is not registered properly. Please contact Factor."
    End If
End Sub


Public Sub ButtonClick(Button As MSComctlLib.Button)
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

Public Sub FormResize(Optional bResize As Boolean = True)
    On Error Resume Next
    
    If bResize Then
        frmMainForm.Height = sngMainFormHeight
        frmMainForm.Width = sngMainFormWidth
    End If
    
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
    Dim i As Integer
    
    objToolbar.GetMenuInfo sCap, sTag, bEnabled, nKey
    mnuContextItems(DEFAULT_MENU).Caption = sCap
    mnuContextItems(DEFAULT_MENU).Tag = sTag
    mnuContextItems(DEFAULT_MENU).Enabled = bEnabled
    mnuContextItems(ADDITIONAL_MENU).Visible = False
    If Not IsMissing(vExKeys) Then
        For i = 0 To UBound(vExKeys)
        If IsNumeric(vExKeys(i)) Then
            If i + 1 > m_nMenuItems Then
                Load mnuContextItems(i + 1)
                m_nMenuItems = m_nMenuItems + 1
            End If
            objToolbar.GetMenuInfo sCap, sTag, bEnabled, Val(vExKeys(i))
            mnuContextItems(i + 1).Caption = sCap
            mnuContextItems(i + 1).Visible = True
            mnuContextItems(i + 1).Tag = sTag
            mnuContextItems(i + 1).Enabled = bEnabled
        End If
        Next i
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

Public Function RunProgram(sProgram As String, Optional sModuleID As String = "") As Boolean
    #If FACTOR_MENU < 0 Then
        If Not t_oleObject Is Nothing Then
            'david 08/05/2003  #412827-9
            If sModuleID = "" Then
                On Error Resume Next
                sModuleID = Trim(frmMainForm!efraToolBar.FMName)
            End If
            ''''''''''''''''''''''''''''
            
            subShowBusyState True, VENDOR_UP
            
            Dim bNewFactmenu As Boolean
            bNewFactmenu = GetFactmenuVersion() > "3.27.0001"
            
            If bNewFactmenu Then
                RunProgram = t_oleObject.RunExe(sProgram, , sModuleID)
            Else
                RunProgram = t_oleObject.RunExe(sProgram)
            End If
            
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

Private Sub Form_Unload(Cancel As Integer)
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
    
    sTemp = UCase(App.path)
    
    nPosi = InStrRev(sTemp, "\")
    
    If nPosi > 0 Then
        fnGetFactorPath = Left(sTemp, nPosi - 1)
    Else
        fnGetFactorPath = ""
    End If
End Function

Public Function GetFactmenuVersion() As String

    On Error GoTo GetFileVersionData_Error
    
    Dim sInfo As String
    Dim lResult As Long
    Dim iDelim As Integer
    Dim lHandle As Long
    Dim lSizeof As Long
    Dim sFile As String

    sFile = fnGetFactorPath + "\factmenu.exe"
    lHandle = 0
    
    'how big is the Version Info block?
    lSizeof = GetFileVersionInfoSize(sFile, lHandle)
    
    If lSizeof > 0 Then
        sInfo = String$(lSizeof, 0)
        
        lResult = GetFileVersionInfo(ByVal sFile, 0&, ByVal lSizeof, ByVal sInfo)
        
        If lResult Then
            sInfo = Replace(sInfo, Chr(0), "")
    
            iDelim = InStr(sInfo, "ProductVersion")
            If iDelim > 0 Then
                iDelim = iDelim + Len("ProductVersion")
                GetFactmenuVersion = Mid$(sInfo, iDelim, 9)
                Exit Function
            End If
        Else
            GoTo invalid_file_info_error
        End If
    Else
        GoTo invalid_file_info_error
    End If

GetFileVersionData_Exit:

Exit Function

GetFileVersionData_Error:
    Resume GetFileVersionData_Exit

invalid_file_info_error:
    GoTo GetFileVersionData_Exit
End Function

Public Function GetProgramVersion(sPathFile As String) As String

    On Error GoTo GetFileVersionData_Error
    
    Dim sInfo As String
    Dim lResult As Long
    Dim iDelim As Integer
    Dim lHandle As Long
    Dim lSizeof As Long

    lHandle = 0
    
    'how big is the Version Info block?
    lSizeof = GetFileVersionInfoSize(sPathFile, lHandle)
    
    If lSizeof > 0 Then
        sInfo = String$(lSizeof, 0)
        
        lResult = GetFileVersionInfo(ByVal sPathFile, 0&, ByVal lSizeof, ByVal sInfo)
        
        If lResult Then
            sInfo = Replace(sInfo, Chr(0), "")
            
            'NOTE: THE ProductVersion may be appended one extra digit
            'e.g. 3.27.0002 will be shown as 3.27.00024
            'it should be ok if we just want to check the version for earlier or latest.
            'e.g. If ProductVersion > "3.27.0001" and ProductVersion <= "3.27.0004" Then ...
            'if it needs to be display on the screen, then make sure you fix it before showing it.
            GetProgramVersion = fnExtractInfo(sInfo, "ProductVersion")
        Else
            GoTo invalid_file_info_error
        End If
    Else
        GoTo invalid_file_info_error
    End If

GetFileVersionData_Exit:

Exit Function

GetFileVersionData_Error:
    Resume GetFileVersionData_Exit

invalid_file_info_error:
    GoTo GetFileVersionData_Exit
End Function

Private Function fnExtractInfo(sInfo As String, sSearchFor As String) As String
    Dim iDelim As Integer
    Dim iEnd As Integer
    Dim sTemp As String
    Dim i As Integer
    Dim n As Integer
    
    iDelim = InStr(sInfo, sSearchFor)
    
    If iDelim > 0 Then
        iDelim = iDelim + Len(sSearchFor)
        sTemp = Mid$(sInfo, iDelim)
        iEnd = -1
        
        For i = 1 To Len(sTemp)
            n = Asc(Mid(sTemp, i, 1))
            If (n < 32 Or n > 126) And n <> 169 And n <> 174 Then
                iEnd = i - 1
                Exit For
            End If
        Next i
        
        If iEnd = -1 Then
            iEnd = Len(sTemp)
        End If
        
        fnExtractInfo = Left(sTemp, iEnd)
    End If
End Function

'this function will return the File Version or Product Version which is same
'as the file properties version from the windows explorer.
Public Function GetProductVersion(ByVal FullFileName As String, Optional bShowMsg As Boolean = False) As String

 Dim rc As Long
 Dim lDummy As Long
 Dim sBuffer() As Byte
 Dim lBufferLen As Long
 Dim lVerPointer As Long
 Dim udtVerBuffer As VS_FIXEDFILEINFO
 Dim lVerbufferLen As Long


 '*** Get size ****
 lBufferLen = GetFileVersionInfoSize(FullFileName, lDummy)
 
 If lBufferLen < 1 Then
    If bShowMsg Then
        MsgBox "No Version Info available for File: " + FullFileName + "."
    End If
    Exit Function
 End If

    '**** Store info to udtVerBuffer struct ****
    ReDim sBuffer(lBufferLen)
    rc = GetFileVersionInfo(FullFileName, 0&, lBufferLen, sBuffer(0))
    rc = VerQueryValue(sBuffer(0), "\", lVerPointer, lVerbufferLen)
    MoveMemory udtVerBuffer, lVerPointer, Len(udtVerBuffer)

 
    '**** Determine File Version number ****
    GetProductVersion = Format$(udtVerBuffer.dwFileVersionMSh) & "." & Format$(udtVerBuffer.dwFileVersionMSl) & "." & Format$(udtVerBuffer.dwFileVersionLSh) & "." & Format$(udtVerBuffer.dwFileVersionLSl)
End Function
