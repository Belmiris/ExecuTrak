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
                     Optional vMsg As Variant)

    ReDim Preserve arryColumns(nColumnCount)
    
    With arryColumns(nColumnCount)
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
    End With
    nColumnCount = nColumnCount + 1

End Sub


Public Sub Click(objCtrl As Object)
    Dim y As Integer
    Dim nRow As Integer
    
    If objCtrl Is picTextFrame Then
        y = objCtrl.Top + objCtrl.Height / 2
        nRow = fnCurrentRow(y)
'        m_fLastTimeShow = Timer + m_fEditDelay
        m_bClicked = True
        'On Error Resume Next
        tgTable.SetFocus
        If nRow >= 0 Then
            tgTable.Row = nRow
        End If
        y = objCtrl.Left + Screen.TwipsPerPixelX * 2
        nRow = fnCurrentColumn(y)
        If nRow >= 0 Then
            tgTable.Col = nRow
        End If
        picTextFrame.Visible = False
    End If
End Sub


Property Let EditDelay(ByVal fDelay As Single)
    m_fEditDelay = fDelay
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
    Dim x As Single
    Dim y As Single
    Dim nRow As Integer
    Dim lFirstRow As Long
    
    GetCursorPos lpPoint
    ScreenToClient tgTable.hwnd, lpPoint
    x = lpPoint.x * Screen.TwipsPerPixelX
    y = lpPoint.y * Screen.TwipsPerPixelY
    With tgTable
        If y < 0 Or y > .Height Then
            m_fLastTimeShow = Timer + SHOW_DELAY_INITIAL
            picTextFrame.Visible = False
            Exit Sub
        End If
        If x < 0 Or x > .Width Then
            m_fLastTimeShow = Timer + SHOW_DELAY_INITIAL
            picTextFrame.Visible = False
            Exit Sub
        End If
    End With
    If Not IsNull(tgTable.FirstRow) Then
        lFirstRow = val(tgTable.FirstRow)
    End If
    If lFirstRow <> m_lLastFirstRow Then
        m_lLastFirstRow = lFirstRow
        picTextFrame.Visible = False
        Exit Sub
    End If
    nRow = fnCurrentRow(y)
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
    nCol = fnCurrentColumn(x)
    If nCol >= tgTable.LeftCol Then
        nCol = fnColumn(nCol)
    Else
        nCol = -1
    End If
    
    If nCol >= 0 Then
        fnShowText nCol, nRow
        m_lOriginalRow = y
        m_lOriginalCol = x
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

Private Function fnCurrentRow(ByVal y As Single) As Integer

    fnCurrentRow = tgTable.RowContaining(y)
Exit Function

    Dim i As Integer
    
    fnCurrentRow = -1
    For i = 0 To tgTable.VisibleRows - 1
        If y > tgTable.RowTop(i) Then
            If y < tgTable.RowTop(i) + tgTable.RowHeight Then
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
    Dim nColWidth As Integer
    Dim nHeight As Integer
    
    Dim nColT As Integer
    Dim vMsg As Variant
    Dim lRow As Long

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
                    lRow = val(tgTable.FirstRow) + nRow
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
        If arryColumns(nCol).m_nFlag <> STYLE_CELLTEXT Or nWidth > nColWidth Then
            nHeight = tgTable.RowHeight + Screen.TwipsPerPixelY
            nTop = tgTable.Top
            nLeft = tgTable.Left
            Select Case arryColumns(nCol).m_nPosition
                Case POS_TOP
                    nTop = nTop + tgTable.RowTop(nRow) - picTextFrame.Height
                    nLeft = nLeft + tgTable.Columns(nColT).Left
                Case POS_BOTTOM
                    nTop = nTop + tgTable.RowTop(nRow) + picTextFrame.Height
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

Public Sub MouseMove(ByVal x As Single, ByVal y As Single)
    Dim nCol As Integer
    Dim bFlag As Boolean
    
    If m_nMode = MODE_TIMER Then
        If Abs(m_lOriginalRow - y) >= 60 Then
            picTextFrame.Visible = False
            Exit Sub
        End If
        If Abs(m_lOriginalCol - x) >= 60 Then
            picTextFrame.Visible = False
        End If
    ElseIf m_nMode = MODE_MOUSE Then
    
        If nLastColumn < 0 Then
            Exit Sub
        End If
        
        With tgTable
            If x > .Columns(nLastColumn).Left Then
                If x < .Columns(nLastColumn).Left + .Columns(nLastColumn).Width Then
                    If nLastRow >= 0 Then
                        If y > .RowTop(nLastRow) Then
                            If y < .RowTop(nLastRow) + .RowHeight Then
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
        fnShowPrompt fnCurrentColumn(x), fnCurrentRow(y)
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


Public Sub MouseDown(ByVal Button As Integer, ByVal x As Single, ByVal y As Single)
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

    nLastColumn = fnCurrentColumn(x)
    
    nCol = fnColumn(nLastColumn)
    bFlag = True
    If nCol >= 0 Then
        With tgTable.Columns(nLastColumn)
            If x >= .Left Then
                If x <= .Left + .Width Then
                    m_lOriginalRow = tgTable.Row
                    m_lOriginalCol = tgTable.Col
                    fnShowText nCol, y
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
    Set picTextFrame = oPic
    picTextFrame.Visible = False
    picTextFrame.BackColor = &HC0FFFF    'Light Yellow
                             '&H80000005  ' White
    If Not tgTable Is Nothing Then
        Set picTextFrame.Font = tgTable.Font
    End If
End Property

Private Function fnCurrentColumn(ByVal x As Single) As Integer
    
    fnCurrentColumn = tgTable.ColContaining(x)
'Exit Function
    Dim i As Integer
    
    fnCurrentColumn = -1
    For i = 0 To tgTable.Columns.Count - 1
        If x > tgTable.Columns(i).Left Then
            If x < tgTable.Columns(i).Left + tgTable.Columns(i).Width Then
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
    Set tgTable = tgTemp
    If Not picTextFrame Is Nothing Then
        Set picTextFrame.Font = tgTable.Font
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
End Sub

Private Sub Class_Terminate()
    Set frmParent = Nothing
    Set lblStatusbar = Nothing
    Set tgTable = Nothing
    Set picTextFrame = Nothing
End Sub


