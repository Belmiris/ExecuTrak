VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTextTip"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'***********************************************************'
'
' Copyright (c) 1998 FACTOR, A Division of W.R.Hess Company
'
'Programmer: Weigong Jiang
'Date:10/06/98
'Objective:
'(1)Showing all the information if Control is not big enough ---STYLE_EXT_TEXT
'(2)Showing extra info such as description
'
'****Note*****
'(1)Some of the functions are copied from tgcombo class
'(2)it is only works for text box now
'*************************************
' Usage:
'(1)Declare varible:
'  private CTextTip as clsTextTip
'  Put a picture Box in the form, say picTextTip
'  You may need to set the Appearance to 0 at design time to
'  have a better looking picture
'(2)Set UP
' Example:
'   Private Sub subSetTextTip()
'     Set CTextTip = New clsTextTip
'     Set CTextTip.Form = Me
'     Set CTextTip.PictureBox = picTextTip
    
'     CTextTip.AddControl txtLastName, CTextTip.GetStyleExt, CTextTip.PositionBottom, 60
'     CTextTip.AddControl txtTerms
'     CTextTip.AddControl txtVendorClass
'     CTextTip.AddControl txt1099FormType, , CTextTip.PositionLeft
'   End Sub
'(3) Call CTextTip.ShowTips in timer event
'(4) Set TAG Property when you want to show extra info
'********************************************************************
Option Explicit

    Private Const POS_TOP As Integer = 1
    Private Const POS_BOTTOM As Integer = 2
    Private Const POS_LEFT As Integer = 3
    Private Const POS_RIGHT As Integer = 4

    Private Const SHOW_DELAY = 1 'delay 1 second
    Private lLastTimeShow As Long
    Private Const STYLE_EXT_TEXT As Integer = 1   'Show complete text
    Private Const STYLE_GET_DESC As Integer = 2   'Show description
    
    Private Const MODE_TIMER = 1
    Private Const MODE_MOUSE = 2

    Private Type tpControles
        m_Control As Control
        m_nFlag As Integer 'Show style
        m_nPosition As Integer
        m_nMaxCharEachLine As Integer 'Set this guy if you want max width
    End Type
    
    Private frmParent As Form
    Private picTextFrame As PictureBox
    
    Private arryControls() As tpControles
    Private nControlCount As Integer

    Private m_nMode As Integer

    #If Win32 Then
        Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
        Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
        Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
        Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
    #Else
        Private Declare Function SetParent Lib "USER" (ByVal hWndChild As Integer, ByVal hWndNewParent As Integer) As Integer
        Private Declare Function GetParent Lib "USER" (ByVal hWnd As Integer) As Integer
        Private Declare Sub GetWindowRect Lib "USER" (ByVal hWnd As Integer, lpRect As RECT)
        Private Declare Function ScreenToClient Lib "USER" (ByVal hWnd As Integer, lpPoint As POINTAPI) As Boolean
    #End If

Private Sub subGetOffsets(nLeft As Integer, _
                         nTop As Integer, _
                         cntrl As Control)

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
    
    nTbl_hWnd = frmParent.hWnd
    n_hWnd = cntrl.hWnd
    n_ParenthWnd = GetParent(n_hWnd)
    nLeft = 0
    nTop = cntrl.Height
    While n_hWnd <> nTbl_hWnd And n_ParenthWnd > 0
        subGetLeftTop nClientLeft, nClientTop, n_ParenthWnd, n_hWnd
        nLeft = nLeft + nClientLeft
        nTop = nTop + nClientTop
        n_hWnd = n_ParenthWnd
        n_ParenthWnd = GetParent(n_hWnd)
    Wend
End Sub

Public Sub AddControl(ByVal txtBox As Control, Optional vExtOrDESC, _
        Optional vPosition, Optional vDisplayCharCount)

    On Error Resume Next
    ReDim Preserve arryControls(nControlCount) As tpControles
    
    With arryControls(nControlCount)
        Set .m_Control = txtBox
        If IsMissing(vExtOrDESC) Then
            .m_nFlag = STYLE_GET_DESC
        Else
            .m_nFlag = vExtOrDESC
        End If
        If IsMissing(vPosition) Then
            .m_nPosition = POS_TOP
        Else
            .m_nPosition = vPosition
        End If
        If IsMissing(vDisplayCharCount) Then
            .m_nMaxCharEachLine = -1
        Else
            .m_nMaxCharEachLine = vDisplayCharCount
        End If
    End With
    nControlCount = nControlCount + 1

End Sub
Public Sub SetMouseMode()
    m_nMode = MODE_MOUSE
End Sub
Public Sub SetTimerMode()
    m_nMode = MODE_TIMER
End Sub
Private Function IsCursorOverControl(ByVal ctrl As Control) As Boolean

    Dim lpPoint As POINTAPI 'holds the current cursor screen coordinates
    Dim lpRect As RECT      'holds the screen coordinates for the efraToolBar frame

    Call GetCursorPos(lpPoint)                   'get the current cursor position in screen coordinates
    Call GetWindowRect(ctrl.hWnd, lpRect) 'get the screen coordinates for the toolbar

    'if the cursor is in the toolbar rectangle then

    IsCursorOverControl = (lpPoint.x >= lpRect.Left And lpPoint.x <= lpRect.Right And lpPoint.y >= lpRect.Top And lpPoint.y <= lpRect.Bottom)
    
End Function
Public Sub ShowTips()
    'This function should be call by a timer event.

    Dim txtBox As Control
    Dim nIndex As Integer
    Dim bFound As Boolean
    Static nLastIndexDisplayed As Integer 'for fast moving mouse only
                                          'We want to show the next tip but old one is still showing
    If m_nMode = MODE_MOUSE Then
        Exit Sub
    End If
    
    m_nMode = MODE_MOUSE
    bFound = False
    For nIndex = 0 To nControlCount - 1
        If IsCursorOverControl(arryControls(nIndex).m_Control) Then
            bFound = True
            Exit For
        End If
    Next
    
    If Not bFound Then
        picTextFrame.Visible = False
        picTextFrame.ZOrder 1
        m_nMode = MODE_TIMER
        Exit Sub
    Else
        If (Not picTextFrame.Visible) Or (nLastIndexDisplayed <> nIndex) Then
            subGetAndPrintTip nIndex
            nLastIndexDisplayed = nIndex
        End If
    End If
    m_nMode = MODE_TIMER
End Sub
Property Get GetStyleDesc() As Integer
    GetStyleDesc = STYLE_GET_DESC
End Property

Property Get GetStyleExt() As Integer
    GetStyleExt = STYLE_EXT_TEXT
End Property
Private Sub subGetLeftTop(nLeft As Single, nTop As Single, ByVal nParenthWnd As Integer, ByVal n_hWnd As Integer)
    
    Dim myRect As RECT
    Dim myPoint As POINTAPI
    Dim nLimit As Integer
    
    nLimit = frmParent.Width

    GetWindowRect n_hWnd, myRect

    myPoint.x = myRect.Left
    myPoint.y = myRect.Top
    
    If ScreenToClient(nParenthWnd, myPoint) Then
        nLeft = myPoint.x * Screen.TwipsPerPixelX     'For Control
        If nLeft < -nLimit Or nLeft > nLimit Then
            nLeft = Screen.TwipsPerPixelX
        End If
        nTop = myPoint.y * Screen.TwipsPerPixelY       'For Control
        If nTop < 0 Or nTop > frmParent.Height Then
            nTop = Screen.TwipsPerPixelX
        End If
    End If

End Sub

Private Sub subWrapText(s As String, _
    noChar As Integer, dWidth As Single, dHeight As Single)
    Dim i As Integer
    Dim nPos As Integer
    Dim sTemp As String
    Dim sOut As String
    
    On Error GoTo NoWrap
    'Kill the return key if alredy there
    Do While i < 10
       nPos = InStr(s, vbCrLf)
   
       If nPos > 0 Then
          s = Left(s, nPos - 1) & Right(s, Len(s) - nPos + 1)
       Else
          Exit Do
       End If
       i = i + 1
    Loop
    
    Do While s <> ""
        i = noChar
        Do While Mid(s, i + 1, 1) <> Space(1)
            
            If i >= Len(s) Then
                Exit Do
            Else
                i = i + 1
            End If
        Loop
        sTemp = Left(s, i)
        s = Mid(s, i + 1, Len(s))
        If dWidth < picTextFrame.TextWidth(sTemp) + picTextFrame.TextWidth(Space(4)) Then
            dWidth = picTextFrame.TextWidth(sTemp) + picTextFrame.TextWidth(Space(4))
        End If
        dHeight = dHeight + picTextFrame.TextHeight(sTemp)
        sOut = sOut & Space(1) & LTrim(sTemp) & vbCrLf
    Loop
    
    s = sOut
    'If dHeight < 300 Then
     '   dHeight = dHeight * 3 / 2
    If dHeight < 1000 Then
        dHeight = dHeight * 5 / 4
    Else
        dHeight = dHeight * 10 / 9
    End If
    Exit Sub
NoWrap:
    dHeight = 300
    dWidth = 2000
    On Error GoTo 0
End Sub

Private Sub subGetAndPrintTip(nIndex As Integer)
     
    Dim sMsg As String
    Dim nLeft As Integer
    Dim nWidth As Integer
    Dim nTop As Integer
    Dim nHeight As Integer
    Dim vMsg As Variant
    
    Dim dWidth As Single
    Dim dHeight As Single
  
    On Error Resume Next
   ' If Timer - lLastTimeShow < SHOW_DELAY Then
   '     Exit Sub
   ' End If
    With arryControls(nIndex)
        Select Case .m_nFlag
            Case STYLE_EXT_TEXT
                sMsg = .m_Control
                If picTextFrame.TextWidth(sMsg & "    ") <= .m_Control.Width Then
                    sMsg = ""
                End If
            Case STYLE_GET_DESC
                sMsg = .m_Control.Tag
        End Select
        
        If sMsg = "" Then
            picTextFrame.Visible = False
            picTextFrame.ZOrder 1
            Exit Sub
        End If

        If .m_nMaxCharEachLine <= 0 Then
            picTextFrame.Width = picTextFrame.TextWidth(sMsg) + picTextFrame.TextWidth(szSPACE) * 4
            picTextFrame.Height = picTextFrame.TextHeight(sMsg) * 5 / 4
            sMsg = Space(1) & Trim(sMsg)
        Else
            'subInsertCRLF sMsg, .m_nMaxCharEachLine, dWidth, dHeight
            subWrapText sMsg, .m_nMaxCharEachLine, dWidth, dHeight
            picTextFrame.Width = dWidth
            picTextFrame.Height = dHeight
        End If
        
        subGetOffsets nLeft, nTop, .m_Control
        
        #If DEVELOP Then
    '    Dim nTextBox_left As Integer
    '    Dim nTextBox_Top As Integer
    '    nTextBox_left = nLeft
    '    nTextBox_Top = nTop - arryControls(nIndex).m_Control.Height
        #End If
        Select Case arryControls(nIndex).m_nPosition
              Case POS_TOP
                  nTop = nTop - .m_Control.Height - picTextFrame.Height
              Case POS_BOTTOM
                  'do nothing
              Case POS_LEFT
                  nTop = nTop - .m_Control.Height
                  nLeft = nLeft - picTextFrame.Width
              Case POS_RIGHT
                  nTop = nTop - .m_Control.Height
                  nLeft = nLeft + picTextFrame.Width
        End Select
        
    End With
    #If DEVELOP Then
      
     '   subMoingEffect nTextBox_left, nTextBox_Top, arryControls(nIndex).m_Control.Height, arryControls(nIndex).m_Control.Width, _
                       nLeft, nTop, dHeight, dWidth
    #End If
    picTextFrame.Top = nTop           'set the top corner coordinate
    picTextFrame.Left = nLeft       'set the left corner coordinate
       
    picTextFrame.Cls
    picTextFrame.Visible = True
    picTextFrame.ZOrder 0
    picTextFrame.Print sMsg
    lLastTimeShow = Timer
End Sub
#If DEVELOP Then
    Private Sub subMoingEffect(a1, b1, c1, d1, a2, b2, c2, d2)
        Dim n As Integer
        Const nStep = 1
        picTextFrame.Visible = True
        picTextFrame.ZOrder 0
        For n = 0 To nStep
           picTextFrame.Left = (nStep - n) / nStep * a1 + n / nStep * a2
           picTextFrame.Top = (nStep - n) / nStep * b1 + n / nStep * b2
           picTextFrame.Height = (nStep - n) / nStep * c1 + n / nStep * c2
           picTextFrame.Width = (nStep - n) / nStep * d1 + n / nStep * d2
           DoEvents
        Next
    End Sub
#End If
Private Sub subInsertCRLF(szText As String, _
             lenOfEachLine As Integer, dWidth As Single, dHeight As Single)
       Dim szTemp As String
       Dim intLines As Integer
       Dim k As Integer
       Dim intLenOfText As Integer
       Dim strOut As String
       
       strOut = ""
       intLenOfText = Len(szText)
       intLines = intLenOfText \ lenOfEachLine
      
       For k = 0 To intLines
          szTemp = Mid(szText, k * lenOfEachLine + 1, lenOfEachLine)
          If dWidth < picTextFrame.TextWidth(szTemp) Then
                dWidth = picTextFrame.TextWidth(szTemp)
          End If
          dHeight = dHeight + picTextFrame.TextHeight(szTemp)
          
          If strOut = "" Then
               strOut = szTemp
          Else
               strOut = strOut & vbCrLf & szTemp
          End If
       Next k
       szText = strOut
       dWidth = dWidth + picTextFrame.TextWidth(szSPACE) + 4
       dHeight = dHeight * 5 / 4
End Sub

Property Set Form(oForm As Form)
    Set frmParent = oForm
End Property

Property Get PositionLeft() As Integer
    PositionLeft = POS_LEFT
End Property
Property Get PositionRight() As Integer
    PositionRight = POS_RIGHT
End Property
Property Get PositionTop() As Integer
    PositionTop = POS_TOP
End Property

Property Get PositionBottom() As Integer
    PositionBottom = POS_BOTTOM
End Property

Property Set PictureBox(oPic As Picture)
    Set picTextFrame = oPic
    picTextFrame.BorderStyle = 1
    picTextFrame.Visible = False
    picTextFrame.AutoRedraw = True
    picTextFrame.BackColor = &HC0FFFF
    '&HC0FFFF    'Light Yellow
                             '&H80000005  ' White
                             
End Property

Private Sub Class_Initialize()
    nControlCount = 0
    m_nMode = MODE_TIMER
End Sub



Private Sub Class_Terminate()
    m_nMode = MODE_MOUSE
End Sub
