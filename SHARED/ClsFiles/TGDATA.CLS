VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTrueGridData"
Attribute VB_Creatable = True
Attribute VB_Exposed = False
' clsTrueGridData
'
' This class is used strictly as a container.  Another module will contain
' a collection of this class.
'
' This class is used by modTrueGrid and should not be explicitly used by
' anything else.
'
' Michael A. Donahue

Option Explicit

' each registered Trud DBGrid gets one of the following class instantiations
' created for it

Public m_nOriginalRowHeight As Integer          ' original row height
Public m_nOriginalRowFontSize As Integer        ' original row font
Public m_nOriginalHeadFontSize As Integer       ' original header font
Public m_colOriginalColumnWidths As Collection  ' original column widths
Public m_colPatterns As Collection              ' column regular expressions
Public m_colErrorMsgs As Collection             ' column error messages
Public m_colStatusMsgs As Collection            ' column status bar messages
Public m_colColumnsRequired As Collection       ' column required flags
Public m_tblGrid As Object                      ' pointer to the True DBGrid
Public m_lblStatusBar As Object                 ' pointer to the status bar
Public m_frmParent As Object                    ' form that the control is on

Private Sub Class_Initialize()

    fnInit
    
End Sub

Private Sub Class_Terminate()
    
    fnInit

End Sub

Private Sub fnInit()
    
    m_nOriginalRowHeight = 0
    m_nOriginalRowFontSize = 0
    m_nOriginalHeadFontSize = 0
    Set m_colOriginalColumnWidths = Nothing
    Set m_colPatterns = Nothing
    Set m_colErrorMsgs = Nothing
    Set m_colStatusMsgs = Nothing
    Set m_colColumnsRequired = Nothing
    Set m_tblGrid = Nothing
    Set m_lblStatusBar = Nothing
    Set m_frmParent = Nothing
    
End Sub

