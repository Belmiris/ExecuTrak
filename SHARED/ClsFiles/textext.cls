VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsColumnExtension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***********************************************************'
'
' Copyright (c) 1996 FACTOR, A Division of W.R.Hess Company
'
' Module name   : TEXTEXT.CLS
' Date          : OCTOBER 1, 1996
' Programmer(s) : Qinggang Ma
'
' This module implementing a class to show extra information in a grid column.
'Usage:

'Controls Required
'   A True Grid
'   A Picturebox (set the Appearance property to 0-flat)
'Call back function
'   A public function fnGetText(tgTable as TDBGrid, ByVal nCol as integer, ByVal nRow as Integer) As String
'       must be added in the form calling this class
'Event Calls (in the true grid)
'   1. MouseDown Button, X, Y
'   2. MouseMove X, Y
'   3. MouseUp
'Event Calls (In Timer Mode)
'   In Timer event call ShowColumnExt
'Set up
'   1.  Properties need to set:
'           a. Form
'           b. PictureBox
'           c. StatusBar
'           d. Table
'   2.  Methods set column:
'           a. AddColumn nCol, vPrompt, vPos, nFlag, vMsg
'               nCol: The column index
'               vPrompt: The prompt shown in the statusbar
'               vMsg: Message to show if any
'               vPos: Position where the box to appear (See Property position)
'               nFlag: What to show
'   3. Auxilary properties and methods
'           a. GetText, ShowCellText, ShowMessage
'               Return constants indicating what text to show
'           b. Message: set the message to show
'           c. Position: Set the position of the box
'           d. PositionDefault: Set the default position
'           e. PositionBottom, PositionLeft, PositionRight, PositionTop
'               Return constants for position to show the box
'           f. Prompt Set prompt
'           g. Style: Set the style indicating what text to show
'           h. StyleDefault: Set the default style
'   David Chai: 09/14/2004  #434264-2
'       Add Multi-line tooltip feature
'           AddColumn nCol, [vPrompt], [vPos], [nFlag], [vMsg], [bMultiLineTooltip=False], _
                [sStringForCrLf=vbCrLf], [nMaxLinesInTooltip=15]
'
'   David 09/14/2004  #434264-2
'   Fixed to allow the tooltip picture box control to be put
'   into a container that is not same as the true grid control
'   NOTE: However, the container control of tooltip picture box must be
'   a parent container control of the true grid if they are not same.


Option Explicit
    
Private Const POS_TOP As Integer = 1
Private Const POS_BOTTOM As Integer = 2
Private Const POS_LEFT As Integer = 3
Private Const POS_RIGHT As Integer = 4

Private Const EXTN_RIGHT As Integer = 0
Private Const EXTN_LEFT As Integer = 1

Private Const SHOW_DELAY = 0.5
Private Const SHOW_DELAY_INITIAL = 1

Private Const STYLE_CELLTEXT As Integer = 1  'Show cell text
Private Const STYLE_MESSAGE As Integer = 2   'Show message
Private Const STYLE_GETTEXT As Integer = 3   'Show text from fnGetText

Private Const MODE_TIMER = 1
Private Const MODE_MOUSE = 2

Private Type tpColumns
    m_nColumn As Integer
    m_sMessage As String
    m_sPrompt As String
    m_nPosition As Integer
    m_nExtension As Integer
    m_nFlag As Integer
    m_bMultiLine As Boolean
    m_sCrLf As String
    m_nMaxLines As Integer
End Type

Private frmParent As Form
Private lblStatusbar As Control
Private tgTable As Object
Private picTextFrame As PictureBox

Private arryColumns() As tpColumns
Private nColumnCount As Integer

Private nLastRow As Integer
Private nLastColumn As Integer
Private nLastMessageCol As Integer
Private nLastMessageRow As Integer
Private nPositionDefault As Integer
Private nStyleDefault As Integer
Private m_lOriginalRow As Integer
Private m_lOriginalCol As Integer
Private m_sSavePropmt As String
Private m_bExternalPrompt As Boolean
Private m_bMouseDown As Boolean
Private m_nMode As Integer
Private m_fLastTimeShow As Single
Private m_bUseEditorRow As Boolean
Private m_fShowDelay As Single
Private m_fEditDelay As Single
Private m_bClicked As Boolean
Private m_lLastFirstRow As Long

Private m_sngToolTipTopOffset As Single
Private m_sngToolTipLeftOffset As Single

#If Win32 Then
    Private Declare Function ScreenToClient Lib "user32" _
            (ByVal hwnd As Long, _
             lpPoint As POINTAPI) As Long
#Else
    Private Declare Function ScreenToClient Lib "USER" _
            (ByVal hwnd As Integer, _
             lpPoint As POINTAPI) As Boolean
#End If

Public Sub AddColumn(ByVal nCol As Integer, _
                     Optional vPrompt As Variant, _
                     Optional ByVal vPos As Variant, _
                     Optional ByVal nFlag As Variant, _
                     Optional vMsg As Variant, _
                     Optional bMultiLineTooltip As Boolean = False, _
                     Optional sStringForCrLf As String = vbCrLf, _
                     Optional nMaxLinesInTooltip As Integer = 15)

    'add change property on the fly feature
    Dim nIndex As Integer
    Dim bColFound As Boolean
    
    nIndex = fnColumn(nCol)
    
    If nIndex >= 0 Then
        bColFound = True
    Else
        nIndex = nColumnCount
        ReDim Preserve arryColumns(nIndex)
    End If
    
    With arryColumns(nIndex)
        .m_nColumn = nCol
        If IsMissing(vPos) Then
            .m_nPosition = nPositionDefault
        Else
            .m_nPosition = vPos
        End If
        If IsMissing(vMsg) Then
            .m_sMessage = ""
        Else
            .m_sMessage = vMsg
        End If
        If IsMissing(vPrompt) Then
            .m_sPrompt = ""
        Else
            .m_sPrompt = vPrompt
        End If
        If IsMissing(nFlag) Then
            .m_nFlag = nStyleDefault
        Else
            .m_nFlag = nFlag
        End If
        
        .m_bMultiLine = bMultiLineTooltip
        .m_sCrLf = sStringForCrLf
        .m_nMaxLines = nMaxLinesInTooltip
    End With
    
    If Not bColFound Then
        nColumnCount = nColumnCount + 1
    End If

End Sub

Public Sub Click(objCtrl As Object)
    Dim Y As Integer
    Dim nRow As Integer
    
    If objCtrl Is picTextFrame Then
        Y = objCtrl.Top + objCtrl.Height / 2
        nRow = fnCurrentRow(Y)
'        m_fLastTimeShow = Timer + m_fEditDelay
        m_bClicked = True
        'On Error Resume Next
        tgTable.SetFocus
        If nRow >= 0 Then
            tgTable.Row = nRow
        End If
        Y = objCtrl.Left + Screen.TwipsPerPixelX * 2
        nRow = fnCurrentColumn(Y)
        If nRow >= 0 Then
            tgTable.Col = nRow
        End If
        picTextFrame.Visible = False
    End If
End Sub

Property Let EditDelay(ByVal fDelay As Single)
    m_fEditDelay = fDelay
End Property

Public Property Let MultiLine(nCol As Integer, bSet As Boolean)
    Dim nIndex As Integer
    
    nIndex = fnColumn(nCol)
    If nIndex >= 0 Then
        arryColumns(nIndex).m_bMultiLine = bSet
    End If
End Property

Public Property Let NewLineCharacter(nCol As Integer, sStringForCrLf As String)
    Dim nIndex As Integer
    
    nIndex = fnColumn(nCol)
    If nIndex >= 0 Then
        arryColumns(nIndex).m_sCrLf = sStringForCrLf
    End If
End Property

Private Function fnColWidth(nCol As Integer) As Integer
    Dim i As Integer
    Dim nLastCol As Integer
    Dim nWidth As Integer

    If tgTable.Splits.Count = 0 Or (tgTable.Splits.Count > 0 And tgTable.Split = tgTable.Splits.Count - 1) Then
        With tgTable
            For i = 0 To .Columns.Count - 1
                If .Columns(i).Visible Then
                    nLastCol = i
                End If
            Next
            nWidth = .Columns(nCol).Width
            If nCol = nLastCol Then
                nWidth = .Width - .Columns(nCol).Left
            End If
        End With
    Else
        nWidth = tgTable.Columns(nCol).Width
    End If
    fnColWidth = nWidth
End Function

Private Function fnLastCol() As Integer
    Dim i As Integer
    Dim nWidth As Integer
    
    With tgTable
        fnLastCol = .Columns.Count - 1
        nWidth = .Columns(tgTable.LeftCol).Left
        For i = tgTable.LeftCol To .Columns.Count - 1
            If .Columns(i).Visible Then
                nWidth = nWidth + .Columns(i).Width
                If nWidth > .Width Then
                    fnLastCol = i
                    Exit For
                End If
            End If
        Next
    End With
End Function


Property Let PositionDefault(nDef As Integer)
    nPositionDefault = nDef
End Property
Public Function RowColChange() As Boolean
    RowColChange = Not m_bMouseDown
    If m_bMouseDown Then
        tgTable.Row = m_lOriginalRow
        tgTable.Col = m_lOriginalCol
        m_bMouseDown = False
    End If
End Function

Public Sub SetExtendLeft(ByVal nCol As Integer)
    
    Dim nIndex As Integer
    
    nIndex = fnColumn(nCol)
    If nIndex >= 0 Then
        arryColumns(nIndex).m_nExtension = EXTN_LEFT
    End If
    
End Sub

Public Sub SetExtendRight(ByVal nCol As Integer)
    
    Dim nIndex As Integer
    
    nIndex = fnColumn(nCol)
    If nIndex >= 0 Then
        arryColumns(nIndex).m_nExtension = EXTN_RIGHT
    End If

End Sub

Public Sub SetMouseMode()
    m_nMode = MODE_MOUSE
End Sub
Public Sub SetTimerMode()
    m_nMode = MODE_TIMER
End Sub

Public Sub ShowColumnExt()
    'This function should be call by a timer event.

    If tgTable Is Nothing Or m_nMode = MODE_MOUSE Then
        Exit Sub
    End If
    If Not (tgTable.Visible And tgTable.Enabled) Then
        Exit Sub
    End If
    
    Dim lpPoint As POINTAPI
    Dim X As Single
    Dim Y As Single
    Dim nRow As Integer
    Dim lFirstRow As Long
    
    GetCursorPos lpPoint
    ScreenToClient tgTable.hwnd, lpPoint
    X = lpPoint.X * Screen.TwipsPerPixelX
    Y = lpPoint.Y * Screen.TwipsPerPixelY
    With tgTable
        If Y < 0 Or Y > .Height Then
            m_fLastTimeShow = Timer + SHOW_DELAY_INITIAL
            picTextFrame.Visible = False
            Exit Sub
        End If
        If X < 0 Or X > .Width Then
            m_fLastTimeShow = Timer + SHOW_DELAY_INITIAL
            picTextFrame.Visible = False
            Exit Sub
        End If
    End With
    If Not IsNull(tgTable.FirstRow) Then
        lFirstRow = Val(tgTable.FirstRow)
    End If
    If lFirstRow <> m_lLastFirstRow Then
        m_lLastFirstRow = lFirstRow
        picTextFrame.Visible = False
        Exit Sub
    End If
    nRow = fnCurrentRow(Y)
    If nRow < 0 Then
        picTextFrame.Visible = False
        Exit Sub
    End If
    If m_bClicked Then
        If nRow = nLastRow Then
            Exit Sub
        Else
            m_bClicked = False
        End If
    End If
    
    If Timer - m_fLastTimeShow < m_fShowDelay Then
        If nRow <> nLastRow Then
            picTextFrame.Visible = False
        End If
        Exit Sub
    End If

    Dim nCol As Integer
'    Dim bFlag As Boolean
    nCol = fnCurrentColumn(X)
    If nCol >= tgTable.LeftCol Then
        nCol = fnColumn(nCol)
    Else
        nCol = -1
    End If
    
    If nCol >= 0 Then
        fnShowText nCol, nRow
        m_lOriginalRow = Y
        m_lOriginalCol = X
    Else
        picTextFrame.Visible = False
    End If
    
End Sub


Property Let ShowDelay(ByVal fDelay As Single)
    m_fShowDelay = fDelay
End Property

Property Let StyleDefault(nDef As Integer)
    nStyleDefault = nDef
End Property


Property Let Style(ByVal nCol As Integer, ByVal nFlag As Integer)
    '3 styles:
        ' 1. Show the text in the cell (ShowCellText)
        ' 2. Show a message given in the setup (ShowMessage)
        ' 3. Show text supplied by a call back function

    Dim nIndex As Integer
    
    nIndex = fnColumn(nCol)
    If nIndex >= 0 Then
        arryColumns(nIndex).m_nFlag = nFlag
    End If
End Property

Property Get ShowCellText() As Integer
    ShowCellText = STYLE_CELLTEXT
End Property

Property Get GetText() As Integer
    GetText = STYLE_GETTEXT
End Property

Property Get ShowMessage() As Integer
    ShowMessage = STYLE_MESSAGE
End Property

'
Private Function fnColumn(ByVal nCol As Integer) As Integer

    fnColumn = -1
    Dim i As Integer
    For i = 0 To nColumnCount - 1
        If nCol = arryColumns(i).m_nColumn Then
            fnColumn = i
            Exit For
        End If
    Next
End Function

Private Function fnCurrentRow(ByVal Y As Single) As Integer

    fnCurrentRow = tgTable.RowContaining(Y)
Exit Function

    Dim i As Integer
    
    fnCurrentRow = -1
    For i = 0 To tgTable.VisibleRows - 1
        If Y > tgTable.RowTop(i) Then
            If Y < tgTable.RowTop(i) + tgTable.RowHeight Then
                fnCurrentRow = i
                Exit For
            End If
        End If
    Next

End Function

Private Sub fnShowText(ByVal nCol As Integer, ByVal nRow As Integer)
    
    If IsNull(tgTable.Row) Then
        Exit Sub
    End If
'    If tgTable.Row < 0 Then
'        Exit Sub
'    End If

    Dim sMsg As String
    Dim nLeft As Integer
    Dim nWidth As Integer
    Dim nTop As Integer
    Dim nLines As Integer
    Dim nColWidth As Integer
    Dim nHeight As Integer
    
    Dim nColT As Integer
    Dim vMsg As Variant
    Dim lRow As Long
    Dim sNewMsg As String
    
    If IsNull(tgTable.RowBookmark(nRow)) Then
        picTextFrame.Visible = False
        Exit Sub
    End If
    If Trim(tgTable.RowBookmark(nRow)) = "" Then
        picTextFrame.Visible = False
        Exit Sub
    End If
    
    If nRow <> nLastRow Then
        picTextFrame.Visible = False
    Else
        If nCol = nLastColumn Then
            If picTextFrame.Visible Then
                Exit Sub
            End If
        End If
    End If
    
    nColT = arryColumns(nCol).m_nColumn
    Select Case arryColumns(nCol).m_nFlag
        Case STYLE_CELLTEXT
            vMsg = tgTable.Columns(arryColumns(nCol).m_nColumn).CellValue(tgTable.RowBookmark(nRow))
            If IsNull(vMsg) Then
                sMsg = ""
            Else
                sMsg = vMsg
            End If
        Case STYLE_MESSAGE
            sMsg = arryColumns(nCol).m_sMessage
        Case STYLE_GETTEXT
            If m_bUseEditorRow Then
                If IsNull(tgTable.FirstRow) Then
                    lRow = nRow
                Else
                    lRow = Val(tgTable.FirstRow) + nRow
                End If
                sMsg = frmParent.fnGetText(tgTable, arryColumns(nCol).m_nColumn, lRow)
            Else
                sMsg = frmParent.fnGetText(tgTable, arryColumns(nCol).m_nColumn, nRow)
            End If
    End Select
    If sMsg = "" Then
        Exit Sub
    End If
    
    If sMsg <> "" Then
        nWidth = picTextFrame.TextWidth(sMsg) + 72
        nColWidth = fnColWidth(nColT)
        
        If arryColumns(nCol).m_nFlag <> STYLE_CELLTEXT Or nWidth > nColWidth Or _
           arryColumns(nCol).m_bMultiLine Then
            If arryColumns(nCol).m_bMultiLine Then
                'get the multi-line message info, reset the tooltiop's width and height
                subGetMultiLineInfo sMsg, arryColumns(nCol).m_sCrLf, nWidth, _
                    nLines, sNewMsg, arryColumns(nCol).m_nMaxLines
                
                If sNewMsg <> "" Then
                    sMsg = sNewMsg
                End If
                
                nHeight = (tgTable.RowHeight - Screen.TwipsPerPixelY) * nLines '+ Screen.TwipsPerPixelY
            Else
                nHeight = tgTable.RowHeight + Screen.TwipsPerPixelY
            End If
            
            If tgTable.Container Is picTextFrame.Container Then
                nTop = tgTable.Top
                nLeft = tgTable.Left
            Else
                nTop = tgTable.Top + m_sngToolTipTopOffset
                nLeft = tgTable.Left + m_sngToolTipLeftOffset
            End If
            
            Select Case arryColumns(nCol).m_nPosition
                Case POS_TOP
                    nTop = nTop + tgTable.RowTop(nRow) - nHeight
                    nLeft = nLeft + tgTable.Columns(nColT).Left
                Case POS_BOTTOM
                    nTop = nTop + tgTable.RowTop(nRow) + tgTable.RowHeight + Screen.TwipsPerPixelY
                    nLeft = nLeft + tgTable.Columns(nColT).Left
                Case POS_LEFT
                    nTop = nTop + tgTable.RowTop(nRow)
                    nLeft = nLeft + tgTable.Columns(nColT).Left
                Case POS_RIGHT
                    nTop = nTop + tgTable.RowTop(nRow)
                    nLeft = nLeft + tgTable.Columns(nColT).Left + nColWidth   'tgTable.Columns(nColT).Width
            End Select
            
            If arryColumns(nCol).m_nExtension = EXTN_LEFT Then
                If arryColumns(nCol).m_nPosition = POS_RIGHT Then
                    nLeft = nLeft - nWidth
                Else
                    nLeft = nLeft - (nWidth - nColWidth)
                End If
            End If
            If nLeft + nWidth > tgTable.Left + tgTable.Width Then
                nLeft = nLeft - ((nLeft + nWidth) - (tgTable.Left + tgTable.Width)) 'tgTable.Width - nWidth
                If nLeft < 0 Then
                    nLeft = 0
                End If
            End If
            If nColT = fnLastCol Then
                nLeft = nLeft - 180
            End If
            picTextFrame.Move nLeft, nTop, nWidth, nHeight
            picTextFrame.Cls
            picTextFrame.Refresh
            picTextFrame.Visible = True
            picTextFrame.ZOrder 0
            DoEvents
            picTextFrame.Print sMsg
        End If
    End If
    nLastRow = nRow
    nLastColumn = nCol

    m_fLastTimeShow = Timer
End Sub

Property Set Form(oForm As Object)
    Set frmParent = oForm
End Property

Property Let message(ByVal nCol As Integer, sMsg As String)
    Dim nIndex As Integer
    
    nIndex = fnColumn(nCol)
    If nIndex >= 0 Then
        arryColumns(nIndex).m_sMessage = sMsg
    End If

End Property

Public Sub MouseMove(ByVal X As Single, ByVal Y As Single)
    Dim nCol As Integer
    Dim bFlag As Boolean
    
    If m_nMode = MODE_TIMER Then
        If Abs(m_lOriginalRow - Y) >= 60 Then
            picTextFrame.Visible = False
            Exit Sub
        End If
        If Abs(m_lOriginalCol - X) >= 60 Then
            picTextFrame.Visible = False
        End If
    ElseIf m_nMode = MODE_MOUSE Then
    
        If nLastColumn < 0 Then
            Exit Sub
        End If
        
        With tgTable
            If X > .Columns(nLastColumn).Left Then
                If X < .Columns(nLastColumn).Left + .Columns(nLastColumn).Width Then
                    If nLastRow >= 0 Then
                        If Y > .RowTop(nLastRow) Then
                            If Y < .RowTop(nLastRow) + .RowHeight Then
                                bFlag = False
                            End If
                        End If
                    End If
                End If
            End If
        End With
        If bFlag Then
            picTextFrame.Visible = False
        End If
        fnShowPrompt fnCurrentColumn(X), fnCurrentRow(Y)
    End If
End Sub

Private Sub fnShowPrompt(ByVal nCol1 As Integer, ByVal lRow As Integer)
    If lblStatusbar Is Nothing Then
        Exit Sub
    End If
    
    Dim nCol As Integer
    nCol = fnColumn(nCol1)
    
    If nCol < 0 Or lRow < 0 Then
        lblStatusbar.Caption = m_sSavePropmt
        m_bExternalPrompt = True
        lblStatusbar.ForeColor = &H80000012
        lblStatusbar.Refresh
        nLastMessageCol = nCol1
        nLastMessageRow = lRow
    Else
        If arryColumns(nCol).m_sPrompt <> "" And (nLastMessageCol <> nCol1 Or nLastMessageRow * lRow < 0) Then
            If m_bExternalPrompt Then
                m_sSavePropmt = lblStatusbar.Caption
            End If
            m_bExternalPrompt = False
            lblStatusbar.Caption = "Hold left mouse button down to " & arryColumns(nCol).m_sPrompt
            lblStatusbar.ForeColor = &H80000012
            lblStatusbar.Refresh
            nLastMessageCol = nCol1
            nLastMessageRow = lRow
        End If
    End If
End Sub

Public Sub MouseUp()
    If m_nMode <> MODE_MOUSE Then
        Exit Sub
    End If
    picTextFrame.Visible = False
'    tgTable.col = m_lOriginalCol
End Sub


Public Sub MouseDown(ByVal Button As Integer, ByVal X As Single, ByVal Y As Single)
    If m_nMode <> MODE_MOUSE Then
        Exit Sub
    End If

    If tgTable Is Nothing Then
        Exit Sub
    End If
    If Button <> vbLeftButton Then
        Exit Sub
    End If
    
    Dim nCol As Integer
    Dim bFlag As Boolean

    nLastColumn = fnCurrentColumn(X)
    
    nCol = fnColumn(nLastColumn)
    bFlag = True
    If nCol >= 0 Then
        With tgTable.Columns(nLastColumn)
            If X >= .Left Then
                If X <= .Left + .Width Then
                    m_lOriginalRow = tgTable.Row
                    m_lOriginalCol = tgTable.Col
                    fnShowText nCol, Y
                    m_bMouseDown = True
                    bFlag = False
                End If
            End If
        End With
    End If
    If bFlag Then
        picTextFrame.Visible = False
    End If
End Sub

Property Let Position(ByVal nCol As Integer, ByVal nPos As Integer)
    Dim nIndex As Integer
    
    nIndex = fnColumn(nCol)
    If nIndex >= 0 Then
        arryColumns(nIndex).m_nPosition = nPos
    End If
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

Property Set PictureBox(oPic As Object)
    Dim ctrl As Control
    Dim bContainerFound As Boolean
    
    Set picTextFrame = oPic
    picTextFrame.Visible = False
    picTextFrame.BackColor = &HC0FFFF    'Light Yellow
                             '&H80000005  ' White
    If Not tgTable Is Nothing Then
        Set picTextFrame.Font = tgTable.Font
    
        Set ctrl = tgTable.Container
        bContainerFound = False
        
        Do While Not ctrl Is Nothing
            If ctrl Is picTextFrame.Container Then
                bContainerFound = True
                Exit Do
            End If
            
            m_sngToolTipTopOffset = m_sngToolTipTopOffset + ctrl.Top
            m_sngToolTipLeftOffset = m_sngToolTipLeftOffset + ctrl.Left
            
            Set ctrl = ctrl.Container
        Loop
        
        If Not bContainerFound Then
            MsgBox "Tooltip is not set on proper container.", vbCritical
            m_sngToolTipTopOffset = 0
            m_sngToolTipLeftOffset = 0
        End If
    End If
End Property

Private Function fnCurrentColumn(ByVal X As Single) As Integer
    
    fnCurrentColumn = tgTable.ColContaining(X)
'Exit Function
    Dim i As Integer
    
    fnCurrentColumn = -1
    For i = 0 To tgTable.Columns.Count - 1
        If X > tgTable.Columns(i).Left Then
            If X < tgTable.Columns(i).Left + tgTable.Columns(i).Width Then
                fnCurrentColumn = i
                Exit For
            End If
        End If
    Next
End Function

Property Let Prompt(ByVal nCol As Integer, sPrompt As String)
    Dim nIndex As Integer
    
    nIndex = fnColumn(nCol)
    If nIndex >= 0 Then
        arryColumns(nIndex).m_sPrompt = sPrompt
    End If

End Property

Property Set StatusBar(oBar As Object)
    Set lblStatusbar = oBar
End Property

Property Set Table(tgTemp As Object)
    Dim ctrl As Control
    Dim bContainerFound As Boolean
    
    Set tgTable = tgTemp
    If Not picTextFrame Is Nothing Then
        Set picTextFrame.Font = tgTable.Font
        
        Set ctrl = tgTable.Container
        bContainerFound = False
        
        Do While Not ctrl Is Nothing
            If ctrl Is picTextFrame.Container Then
                bContainerFound = True
                Exit Do
            End If
            
            m_sngToolTipTopOffset = m_sngToolTipTopOffset + ctrl.Top
            m_sngToolTipLeftOffset = m_sngToolTipLeftOffset + ctrl.Left
            
            Set ctrl = ctrl.Container
        Loop
        
        If Not bContainerFound Then
            MsgBox "Tooltip is not set on proper container.", vbCritical
            m_sngToolTipTopOffset = 0
            m_sngToolTipLeftOffset = 0
        End If
    End If
End Property

Property Let UseEditorRow(bTemp As Boolean)
    m_bUseEditorRow = bTemp
End Property

Private Sub Class_Initialize()
    nColumnCount = 0
    nLastMessageCol = -1
    nLastMessageRow = -1
    nPositionDefault = POS_LEFT
    nStyleDefault = STYLE_CELLTEXT
    m_fShowDelay = SHOW_DELAY
    m_fEditDelay = SHOW_DELAY_INITIAL
    m_bClicked = False
    m_lLastFirstRow = 0
    m_bExternalPrompt = True
    m_bUseEditorRow = False
    m_nMode = MODE_TIMER

    m_sngToolTipTopOffset = 0
    m_sngToolTipLeftOffset = 0
End Sub

Private Sub Class_Terminate()
    Set frmParent = Nothing
    Set lblStatusbar = Nothing
    Set tgTable = Nothing
    Set picTextFrame = Nothing
End Sub

Private Sub subGetMultiLineInfo(ByVal sMsg As String, ByVal sCrLf As String, ByRef nMaxTextWidth As Integer, _
                                ByRef nLines As Integer, Optional ByRef sNewMsg As String = "", _
                                Optional ByVal nMaxLines As Integer = 5)
    
    Const sDotDotDot As String = " ....."
    
    Dim i As Integer
    Dim aryString() As String
    Dim nWidth As Integer
    Dim bCrLfRemoved As Boolean
    Dim bLinesExceeded As Boolean
    
    nLines = 1
    nMaxTextWidth = 1000
    sNewMsg = ""
    
    On Error GoTo ErrTrap
    aryString = Split(sMsg, sCrLf)
    
    nLines = UBound(aryString) + 1
    
    If nLines = 1 Then
        nMaxTextWidth = picTextFrame.TextWidth(sMsg) + 72
        Exit Sub
    End If
    
    'remove the crlf from the bottom of the text
    i = nLines - 1
    For i = i To 0 Step -1
        If Trim(aryString(i)) <> "" Then
            Exit For
        End If

        nLines = nLines - 1
        bCrLfRemoved = True
    Next i
    
    If nLines > nMaxLines Then
        bLinesExceeded = True
        nLines = nMaxLines
        sNewMsg = ""
    End If
    
    nMaxTextWidth = -1
    
    For i = 0 To nLines - 1
        If Trim(aryString(i)) <> "" Then
            nWidth = picTextFrame.TextWidth(aryString(i)) + 72
            
            If nWidth > nMaxTextWidth Then
                nMaxTextWidth = nWidth
            End If
        End If
    Next i
    
    If bLinesExceeded Or bCrLfRemoved Then
        sNewMsg = Trim(aryString(0))
        
        For i = 1 To nLines - 1
            sNewMsg = sNewMsg + vbCrLf + Trim(aryString(i))
        Next i
        
        If bLinesExceeded Then
            sNewMsg = sNewMsg + sDotDotDot
            nMaxTextWidth = nMaxTextWidth + picTextFrame.TextWidth(sDotDotDot) + 50
        End If
    End If
    
    Exit Sub
ErrTrap:
    'do nothing
End Sub
