VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SizingKit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ARY_INC = 10

Private Type tpControlSize
    m_objControl As Object
    m_nLeft As Integer
    m_nTop As Integer
    m_nWidth As Integer
    m_nHeight As Integer
    m_nSizeWhat As Integer
End Type

Private aryControlSize() As tpControlSize
Private m_nControlCount As Integer
Private csFormSize As tpControlSize

Private m_nMinimumWidth As Integer
Private m_nMinimumHeight As Integer
Private m_nLastWindowState As Integer
Private m_nMaxNormalWindowState As Integer
'

Public Sub AddControl(objControl As Control, _
                      ByVal nSize As Integer)
    
    If m_nControlCount = 0 Then
        ReDim aryControlSize(ARY_INC)
    Else
        If UBound(aryControlSize) < m_nControlCount Then
            ReDim Preserve aryControlSize(ARY_INC + m_nControlCount)
        End If
    End If
    With aryControlSize(m_nControlCount)
        Set .m_objControl = objControl
        .m_nLeft = objControl.Left
        .m_nTop = objControl.Top
        .m_nWidth = objControl.Width
        .m_nHeight = objControl.Height
        .m_nSizeWhat = nSize
    End With
    m_nControlCount = m_nControlCount + 1

End Sub

Private Function fnTestFlag(ByVal nFlag As Integer, _
                            ByVal nTester As Integer) As Boolean
                            
    fnTestFlag = (nFlag And nTester) > 0
    
End Function


Property Set MainForm(objForm As Form)
    With csFormSize
        Set .m_objControl = objForm
        .m_nLeft = objForm.Left
        .m_nTop = objForm.Top
        .m_nWidth = objForm.Width
        .m_nHeight = objForm.Height
    End With
End Property

Property Let MinimumWidth(ByVal nTemp As Integer)
    m_nMinimumWidth = nTemp
End Property

Property Let MinimumHeight(ByVal nTemp As Integer)
    m_nMinimumHeight = nTemp
End Property

Public Sub Resize()

    Dim i As Integer
    Dim nLeft As Integer
    Dim nWidth As Integer
    Dim nTop As Integer
    Dim nHeight As Integer
    Dim bSize As Boolean
    Dim fFactorWidth As Single
    Dim fFactorHeight As Single
    
    If csFormSize.m_objControl.WindowState = vbMinimized Then
        m_nLastWindowState = csFormSize.m_objControl.WindowState
        Exit Sub
    End If
    
    If m_nLastWindowState = vbMinimized And csFormSize.m_objControl.WindowState = m_nMaxNormalWindowState Then
        m_nLastWindowState = csFormSize.m_objControl.WindowState
        Exit Sub
    End If
    
    'david 10/05/2001
    If csFormSize.m_objControl.WindowState = m_nMaxNormalWindowState Then
        m_nLastWindowState = csFormSize.m_objControl.WindowState
        Exit Sub
    End If
    
    If csFormSize.m_objControl.Width < m_nMinimumWidth Or csFormSize.m_objControl.Height < m_nMinimumHeight Then
        If csFormSize.m_objControl.Width < m_nMinimumWidth Then
            nWidth = m_nMinimumWidth
        Else
            nWidth = csFormSize.m_objControl.Width
        End If
        If csFormSize.m_objControl.Height < m_nMinimumHeight Then
            nHeight = m_nMinimumHeight
        Else
            nHeight = csFormSize.m_objControl.Height
        End If
        With csFormSize.m_objControl
            .Move .Left, .Top, nWidth, nHeight
        End With
        m_nLastWindowState = csFormSize.m_objControl.WindowState
        Exit Sub
    End If
    
    fFactorWidth = csFormSize.m_objControl.Width / csFormSize.m_nWidth
    fFactorHeight = csFormSize.m_objControl.Height / csFormSize.m_nHeight
    
    For i = 0 To m_nControlCount - 1
        With aryControlSize(i)
            If fnTestFlag(.m_nSizeWhat, SCT_KEEP_WIDTH) Then
                nWidth = .m_nWidth
            Else
                nWidth = .m_nWidth * fFactorWidth
            End If
            
            If fnTestFlag(.m_nSizeWhat, SCT_KEEP_HEIGHT) Then
                nHeight = .m_nHeight
            Else
                nHeight = .m_nHeight * fFactorHeight
            End If
            
            If fnTestFlag(.m_nSizeWhat, SCT_KEEP_LEFT) And fnTestFlag(.m_nSizeWhat, SCT_KEEP_RIGHT) Then
                nLeft = .m_nLeft
                nWidth = csFormSize.m_objControl.Width - nLeft - (csFormSize.m_nWidth - (.m_nLeft + .m_nWidth))
            Else
                If fnTestFlag(.m_nSizeWhat, SCT_KEEP_LEFT) Then
                    nLeft = .m_nLeft
                Else
                    nLeft = .m_nLeft * fFactorWidth
                End If
                
                If fnTestFlag(.m_nSizeWhat, SCT_KEEP_RIGHT) Then
                    nLeft = csFormSize.m_objControl.Width - nWidth - (csFormSize.m_nWidth - (.m_nLeft + .m_nWidth))
                End If
            End If
            
            If fnTestFlag(.m_nSizeWhat, SCT_KEEP_TOP) And fnTestFlag(.m_nSizeWhat, SCT_KEEP_BOTTOM) Then
                nTop = .m_nTop
                nHeight = csFormSize.m_objControl.Height - nTop - (csFormSize.m_nHeight - (.m_nTop + .m_nHeight))
            Else
                If fnTestFlag(.m_nSizeWhat, SCT_KEEP_TOP) Then
                    nTop = .m_nTop
                Else
                    nTop = .m_nTop * fFactorHeight
                End If
                
                If fnTestFlag(.m_nSizeWhat, SCT_KEEP_BOTTOM) Then
                    nTop = csFormSize.m_objControl.Height - nHeight - (csFormSize.m_nHeight - (.m_nTop + .m_nHeight))
                End If
            End If
            
            .m_objControl.Move nLeft, nTop, nWidth, nHeight
        End With
    Next
    
    m_nLastWindowState = csFormSize.m_objControl.WindowState
    m_nMaxNormalWindowState = csFormSize.m_objControl.WindowState
End Sub


Private Sub Class_Initialize()
    m_nControlCount = 0
    m_nLastWindowState = -1
    m_nMaxNormalWindowState = -1
End Sub


