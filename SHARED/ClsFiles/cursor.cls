VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCursor"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit
'***********************************************************'
'
' Copyright (c) 1996 FACTOR, A Division of W.R.Hess Company
'
' Module name   : CURSOR.CLS
' Date          : March 20, 1997
' Programmer(s) : Qinggang Ma
'
' This module defines a class for mouse pointer control

'   Public Functions and properties:

'   For set up:
'   1. MainForm (property set): Set the main form
'   2. AddControl: Add a control in whose mouse pointer will not be changed
'   3. AddScreen: Add a screen for mouse cursor control
'   4. AddArea: Add an area in which mouse pointer is controlled
'   5. SetLeftTop: Set the left and top of area area
'   6. SetRightBottom: Set the right and bottom of an area
'   7. SetAll: Set the left-top and right-bottom of an area

'   For changing cursor
'   8. CurrentScreen: Choose the current screen(for multiple screens)
'   9. SetCursorState: Set the mousepointer in a given area
'   10. SetToNoDrop: Set the mousepointer to no drop for the given area
'   11. SetToDefault: Set the mousepointer to default for the given area

'   Event in timer:
'   12. SetMousePointer

Private Const MF_BYPOSITION = &H400
Private Const STATE_VISIBLE = &H80
Private Const DEFAULT_STATE = vbDefault
Private Const MOUSE_POINTER_NOT_SET = -1

#If Win32 Then
    Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
    Private Declare Function GetMenuState Lib "user32" (ByVal hMenu As Long, ByVal wID As Long, ByVal wFlags As Long) As Long
    Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
    Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
    
    Private Type POINTAPI
        x As Long
        y As Long
    End Type
    
    Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
    End Type
#Else
    Private Declare Sub GetCursorPos Lib "USER" (lpPoint As POINTAPI)
    Private Declare Function GetMenu Lib "USER" (ByVal hwnd As Integer) As Integer
    Private Declare Function GetMenuItemCount Lib "USER" (ByVal hMenu As Integer) As Integer
    Private Declare Function GetMenuState Lib "USER" (ByVal hMenu As Integer, ByVal wID As Integer, ByVal wFlags As Integer) As Integer
    Private Declare Function GetParent Lib "USER" (ByVal hwnd As Integer) As Integer
    Private Declare Sub GetWindowRect Lib "USER" (ByVal hwnd As Integer, lpRect As RECT)
    Private Declare Function ScreenToClient Lib "USER" (ByVal hwnd As Integer, lpPoint As POINTAPI) As Boolean
    
    Private Type POINTAPI 'Point structure
        x As Integer
        y As Integer
    End Type
    
    Private Type RECT 'used for tooltip and context menus
        Left As Integer
        Top As Integer
        Right As Integer
        Bottom As Integer
    End Type
#End If

Private Type tpRegions
    m_nLeft As Integer
    m_nTop As Integer
    m_nRight As Integer
    m_nBottom As Integer
    m_nCursorState As Integer
    m_nAreaID As Integer
    m_ctrlFlag As Control
End Type

Private Type tpScreen
    m_rectRegions() As tpRegions
    m_nRegionCount As Integer
    m_nScreenID As Integer
End Type

Private frmParent As Form
Private arryScreens() As tpScreen
Private m_nScreenCount As Integer
Private m_nCurrentScreen As Integer
Private m_nCurrentArea As Integer
Private bEnabled As Boolean

Private m_arryControls() As Object   'Control
Dim m_nControlCount As Integer


Public Function AddArea(Optional vArea As Variant, _
                        Optional nMouse As Variant, _
                        Optional bSetArea As Variant, _
                        Optional ByVal vScreenID As Variant) As Integer
    Dim nID As Integer
    Dim nIndex As Integer
    Dim nArea As Integer
    Dim ctrlFlag As Control
    
    If IsMissing(vScreenID) Then
        nIndex = m_nCurrentScreen
    Else
        nID = val(vScreenID)
        nIndex = fnIndexScreen(nID)
    End If
    
    AddArea = -1
    If nIndex >= 0 Then
        With arryScreens(nIndex)
            m_nCurrentArea = .m_nRegionCount
            If .m_nRegionCount = 0 Then
                ReDim .m_rectRegions(.m_nRegionCount)
            Else
                ReDim Preserve .m_rectRegions(.m_nRegionCount)
            End If
            .m_nRegionCount = .m_nRegionCount + 1
            With .m_rectRegions(m_nCurrentArea)
                If IsMissing(nMouse) Then
                    .m_nCursorState = MOUSE_POINTER_NOT_SET
                Else
                    .m_nCursorState = val(nMouse)
                End If
                If IsMissing(vArea) Then
                    .m_nAreaID = m_nCurrentArea
                    AddArea = .m_nAreaID
                Else
                    If IsObject(vArea) Then
                        Set .m_ctrlFlag = vArea
                        AddArea = .m_ctrlFlag.TabIndex
                        SetAll .m_ctrlFlag
                    Else
                        .m_nAreaID = val(vArea)
                        AddArea = .m_nAreaID
                    End If
                End If
            End With
        End With
    End If
    
End Function

Public Sub AddControl(objCtrl As Object)
    Dim i As Integer
    
    For i = 0 To m_nControlCount - 1
        If objCtrl Is m_arryControls(i) Then
            Exit Sub
        End If
    Next
    If m_nControlCount = 0 Then
        ReDim m_arryControls(m_nControlCount)
    Else
        ReDim Preserve m_arryControls(m_nControlCount)
    End If
    
    Set m_arryControls(m_nControlCount) = objCtrl
    m_nControlCount = m_nControlCount + 1
End Sub

Public Sub AddControls(ParamArray objCtrls() As Variant)
    Dim i As Integer
    m_nControlCount = UBound(objCtrls) + 1
    
    ReDim m_arryControls(m_nControlCount)
    For i = 0 To m_nControlCount - 1
        If IsObject(objCtrls(i)) Then
            Set m_arryControls(i) = objCtrls(i)
        End If
    Next
End Sub

Public Function AddScreen(Optional vID As Variant) As Integer
    Dim nID As Integer
    If IsMissing(vID) Then
        nID = m_nScreenCount
    Else
        nID = val(vID)
    End If
    If m_nScreenCount = 0 Then
        ReDim arryScreens(m_nScreenCount)
    Else
        ReDim Preserve arryScreens(m_nScreenCount)
    End If
    arryScreens(m_nScreenCount).m_nScreenID = nID

    m_nCurrentScreen = m_nScreenCount
    m_nScreenCount = m_nScreenCount + 1

    AddScreen = nID
    
End Function
Property Let Enabled(bFlag As Boolean)
    bEnabled = bFlag
End Property

Property Get Enabled() As Boolean
    Enabled = bEnabled
End Property

Private Function fnIndexArea(nArea As Integer, nScreen As Integer) As Integer

    Dim i As Integer
    
    fnIndexArea = -1
    If nScreen >= 0 And nScreen < m_nScreenCount Then
        With arryScreens(nScreen)
            For i = 0 To .m_nRegionCount - 1
                With .m_rectRegions(i)
                    If .m_ctrlFlag Is Nothing Then
                        If nArea = .m_nAreaID Then
                            fnIndexArea = i
                            Exit For
                        End If
                    Else
                        If nArea = .m_ctrlFlag.TabIndex Then
                            fnIndexArea = i
                            Exit For
                        End If
                    End If
                End With
            Next
        End With
    End If
End Function

Private Function fnIndexScreen(ByVal nScreenID As Integer) As Integer
    Dim i As Integer
    
    If nScreenID >= 0 And nScreenID < m_nScreenCount Then
        If arryScreens(nScreenID).m_nScreenID = nScreenID Then
            fnIndexScreen = nScreenID
            Exit Function
        End If
    End If
    
    fnIndexScreen = -1
    For i = 0 To m_nScreenCount - 1
        If arryScreens(i).m_nScreenID = nScreenID Then
            fnIndexScreen = i
            Exit For
        End If
    Next
    
End Function

Property Set MainForm(objForm As Form)
    Set frmParent = objForm
    bEnabled = True
End Property

Public Sub SetAll(ctrl As Control, _
                  Optional vArea As Variant, _
                  Optional vScreen As Variant)
    SetLeftTop ctrl, , , vArea, vScreen
    SetRightBottom ctrl, , , vArea, vScreen
End Sub

Public Sub SetLeftTop(Optional vCtrl As Variant, _
                      Optional vLeft As Variant, _
                      Optional vTop As Variant, _
                      Optional vArea As Variant, _
                      Optional vScreen As Variant)
    Dim nArea As Integer
    Dim nScreen As Integer
    Dim nLeft As Integer
    Dim nTop As Integer
    Dim nLeft0 As Integer
    Dim nTop0 As Integer
    Dim nIndex As Integer
    Dim objCtrl As Object
    
    If IsMissing(vScreen) Then
        nIndex = m_nCurrentScreen
    Else
        nScreen = val(vScreen)
        nIndex = fnIndexScreen(nScreen)
    End If
    If IsMissing(vArea) Then
        nArea = m_nCurrentArea
    Else
        If IsObject(vArea) Then
            nArea = fnIndexArea(vArea.TabIndex, nIndex)
        Else
            nArea = fnIndexArea(val(vArea), nIndex)
        End If
    End If
    
    If nIndex >= 0 Then
        If IsMissing(vLeft) Then
            nLeft0 = 0
        Else
            nLeft0 = val(vLeft)
        End If
        If IsMissing(vTop) Then
            nTop0 = 0
        Else
            nTop0 = val(vTop)
        End If
        If IsMissing(vCtrl) Then
            nLeft = nLeft0
            nTop = nTop0
        Else
            If IsObject(vCtrl) Then
                Set objCtrl = vCtrl
                subGetOffsets nLeft, nTop, objCtrl
                nLeft = nLeft - nLeft0
                nTop = nTop - nTop0
            End If
        End If
        With arryScreens(nIndex)
            If nArea >= 0 And nArea < .m_nRegionCount Then
                .m_rectRegions(nArea).m_nLeft = nLeft
                .m_rectRegions(nArea).m_nTop = nTop
            End If
        End With
    End If

End Sub

Public Sub SetCursorState(nMP As Integer, _
                          Optional vAreaID As Variant, _
                          Optional vScreen As Variant)

    Dim nID As Integer
    Dim nIndex As Integer
    Dim nArea As Integer
    Dim nAreaID As Integer
    
    If IsMissing(vScreen) Then
        nIndex = m_nCurrentScreen
    Else
        nID = val(vScreen)
        nIndex = fnIndexScreen(nID)
    End If
    
    If nIndex >= 0 Then
        If IsMissing(vAreaID) Then
            nArea = m_nCurrentArea
        Else
            If IsObject(vAreaID) Then
                nAreaID = vAreaID.TabIndex
            Else
                nAreaID = val(vAreaID)
            End If
            nArea = fnIndexArea(nAreaID, nIndex)
        End If
        With arryScreens(nIndex)
            If nArea >= 0 And nArea < .m_nRegionCount Then
                .m_rectRegions(nArea).m_nCursorState = nMP
            End If
        End With
    End If

End Sub

Public Sub SetToNoDrop(objArea As Variant, Optional vScreen As Variant)
    Dim nAreaID As Integer
    If IsObject(objArea) Then
        nAreaID = objArea.TabIndex
    Else
        nAreaID = val(objArea)
    End If
    SetCursorState vbNoDrop, nAreaID, vScreen
    
End Sub

Public Sub SetToDefault(objArea As Variant, Optional vScreen As Variant)
    Dim nAreaID As Integer
    If IsObject(objArea) Then
        nAreaID = objArea.TabIndex
    Else
        nAreaID = val(objArea)
    End If
    SetCursorState vbDefault, nAreaID, vScreen
    
End Sub

Public Sub SetRightBottom(Optional vCtrl As Variant, _
                          Optional vWidth As Variant, _
                          Optional vHeight As Variant, _
                          Optional vArea As Variant, _
                          Optional vScreen As Variant)
    Dim nArea As Integer
    Dim nScreen As Integer
    Dim nLeft As Integer
    Dim nTop As Integer
    Dim nWidth As Integer
    Dim nHeight As Integer
    Dim nIndex As Integer
    Dim objCtrl As Object
    
    If IsMissing(vScreen) Then
        nIndex = m_nCurrentScreen
    Else
        nScreen = val(vScreen)
        nIndex = fnIndexScreen(nScreen)
    End If
    If IsMissing(vArea) Then
        nArea = m_nCurrentArea
    Else
        If IsObject(vArea) Then
            nArea = fnIndexArea(vArea.TabIndex, nIndex)
        Else
            nArea = fnIndexArea(val(vArea), nIndex)
        End If
    End If
    
    If nIndex >= 0 Then
        If IsMissing(vWidth) Then
            nWidth = 0
        Else
            nWidth = val(vWidth)
        End If
        If IsMissing(vHeight) Then
            nHeight = 0
        Else
            nHeight = val(vHeight)
        End If
        If IsMissing(vCtrl) Then
            nLeft = 0
            nTop = 0
        Else
            If IsObject(vCtrl) Then
                Set objCtrl = vCtrl
                subGetOffsets nLeft, nTop, objCtrl
                nLeft = nLeft + objCtrl.Width
                nTop = nTop + objCtrl.Height
            End If
        End If
        nLeft = nLeft + nWidth
        nTop = nTop + nHeight
        With arryScreens(nIndex)
            If nArea >= 0 And nArea < .m_nRegionCount Then
                .m_rectRegions(nArea).m_nRight = nLeft
                .m_rectRegions(nArea).m_nBottom = nTop
            End If
        End With
    End If

End Sub

Property Let CurrentScreen(ByVal nID As Integer)
    Dim nIndex As Integer
    
    nIndex = fnIndexScreen(nID)
    m_nCurrentScreen = nIndex
End Property

Property Get CurrentScreen() As Integer
    
    If m_nCurrentScreen < 0 Then
        CurrentScreen = m_nCurrentScreen
    Else
        CurrentScreen = arryScreens(m_nCurrentScreen).m_nScreenID
    End If

End Property

#If Win32 Then
Private Sub subGetLeftTop(nLeft As Single, nTop As Single, ByVal nParenthWnd As Long, ByVal n_hWnd As Long)
#Else
Private Sub subGetLeftTop(nLeft As Single, nTop As Single, ByVal nParenthWnd As Integer, ByVal n_hWnd As Integer)
#End If
    'Returns the Left and Top of a control
    Dim myRect As RECT
    Dim myPoint As POINTAPI
    Dim nLimit As Integer
    
    nLimit = frmParent.Width

    GetWindowRect n_hWnd, myRect

    myPoint.x = myRect.Left
    myPoint.y = myRect.Top
    
    If ScreenToClient(nParenthWnd, myPoint) Then
        nLeft = myPoint.x * Screen.TwipsPerPixelX     'For textbox
        If nLeft < -nLimit Or nLeft > nLimit Then
            nLeft = Screen.TwipsPerPixelX
        End If
        nTop = myPoint.y * Screen.TwipsPerPixelY       'For textbox
        If nTop < 0 Or nTop > frmParent.Height Then
            nTop = Screen.TwipsPerPixelX
        End If
    End If

End Sub

Private Sub subGetOffsets(nLeft As Integer, _
                          nTop As Integer, _
                          txtBox As Object)

    'Find out the offsets of the table for a combo
    #If Win32 Then
        Dim nTbl_hWnd As Long
        Dim n_hWnd As Long
        Dim n_ParenthWnd As Long
    #Else
        Dim nTbl_hWnd As Integer
        Dim n_hWnd As Integer
        Dim n_ParenthWnd As Integer
    #End If
    Dim nClientLeft As Single
    Dim nClientTop As Single
    
    nTbl_hWnd = frmParent.hwnd
    n_hWnd = txtBox.hwnd
    n_ParenthWnd = GetParent(n_hWnd)
    nLeft = 0
    nTop = 0
    While n_hWnd <> nTbl_hWnd And n_ParenthWnd > 0
        subGetLeftTop nClientLeft, nClientTop, n_ParenthWnd, n_hWnd
        nLeft = nLeft + nClientLeft
        nTop = nTop + nClientTop
        n_hWnd = n_ParenthWnd
        n_ParenthWnd = GetParent(n_hWnd)
    Wend
End Sub

Public Sub SetMousePointer(Optional vScreen As Variant)
    If frmParent.WindowState = vbMinimized Or Not bEnabled Then
        Exit Sub
    End If
    Dim myPoint As POINTAPI
    Dim nCursorX As Single
    Dim nCursorY As Single
    Dim nMouse As Integer
    Dim nScreen As Integer
    Dim i As Integer
    
    If IsMissing(vScreen) Then
        nScreen = m_nCurrentScreen
    Else
        nScreen = vScreen
    End If
    
    GetCursorPos myPoint
    
    ScreenToClient frmParent.hwnd, myPoint
    nCursorX = myPoint.x * Screen.TwipsPerPixelX
    nCursorY = myPoint.y * Screen.TwipsPerPixelY
    
    nMouse = vbDefault

    If nScreen >= 0 And nScreen < m_nScreenCount Then
        With arryScreens(nScreen)
            For i = 0 To .m_nRegionCount - 1
                With .m_rectRegions(i)
                    If nCursorX > .m_nLeft Then
                        If nCursorX < .m_nRight Then
                            If nCursorY > .m_nTop Then
                                If nCursorY < .m_nBottom Then
                                    If .m_nCursorState = MOUSE_POINTER_NOT_SET Then
                                        If .m_ctrlFlag Is Nothing Then
                                            nMouse = DEFAULT_STATE
                                        Else
                                            If .m_ctrlFlag.Enabled Then
                                                nMouse = DEFAULT_STATE
                                            Else
                                                nMouse = vbNoDrop
                                            End If
                                        End If
                                    Else
                                        nMouse = .m_nCursorState
                                    End If
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                End With
            Next
        End With
    End If
    For i = 0 To m_nControlCount - 1
        With m_arryControls(i)
            If .Visible Then
                If TypeOf m_arryControls(i) Is Form Then
                    nCursorX = nCursorX + frmParent.Left + (frmParent.Width - frmParent.ScaleWidth) \ 2
                    nCursorY = nCursorY + frmParent.Top + frmParent.Height - frmParent.ScaleHeight
                End If
                If nCursorX > .Left Then
                    If nCursorX < .Left + .Width Then
                        If nCursorY > .Top Then
                            If nCursorY < .Top + .Height Then
                                nMouse = DEFAULT_STATE
                                'nMouse = Screen.MousePointer
                            End If
                        End If
                    End If
                End If
            End If
        End With
    Next
    If fnMenuPopUp Then
        nMouse = DEFAULT_STATE
    End If
    If nMouse <> Screen.MousePointer Then
        Screen.MousePointer = nMouse
    End If
End Sub

Private Function fnMenuPopUp() As Boolean
    #If Win32 Then
        Dim n_hMenu As Long
        Dim nState As Long
        Dim nCount As Long
        Dim i As Long
    #Else
        Dim n_hMenu As Integer
        Dim nState As Integer
        Dim nCount As Integer
        Dim i As Integer
    #End If
    Dim bFlag As Boolean
    
    n_hMenu = GetMenu(frmParent.hwnd)
    nCount = GetMenuItemCount(n_hMenu)
    
    If nCount = 0 Then
        fnMenuPopUp = True
    Else
        For i = 0 To nCount - 1
            nState = GetMenuState(n_hMenu, i, MF_BYPOSITION)
            bFlag = nState And STATE_VISIBLE
            If bFlag Then
                Exit For
            End If
        Next
        fnMenuPopUp = bFlag
    End If

End Function



Private Sub Class_Terminate()
    Dim i As Integer
    Dim j As Integer
    
    For i = 0 To m_nScreenCount - 1
        For j = 0 To arryScreens(i).m_nRegionCount - 1
            Set arryScreens(i).m_rectRegions(j).m_ctrlFlag = Nothing
        Next j
    Next i
    For i = 0 To m_nControlCount - 1
        Set m_arryControls(i) = Nothing
    Next i
    Set frmParent = Nothing
    
End Sub


