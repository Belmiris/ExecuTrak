VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsColorCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This Class is required when USE_ARRAY is set to 0
'When USE_ARRAY is set to -1, this Class is not required.

Option Explicit

Private m_nCol As Integer
Private m_nCondCount As Integer
Private m_sCondition() As String
Private m_lForeColor() As Long
Private m_lBackColor() As Long
Private m_sCellStyle() As String
'

Public Sub AddCondition(nCol As Integer, ParamArray aryColorCond() As Variant)
    Dim i As Integer
    Dim j As Integer
    
    i = -1
    
    On Error Resume Next
    i = UBound(aryColorCond(0))
    On Error GoTo 0
    
    If i < 0 Then
        Exit Sub
    End If
    
    'need to have 3 parm in a group
    'namely Condition, ForeColor, and BackColor
    'or
    'need to have 2 parm in a group
    'namely Condition and CellStyle constant (e.g. COLHEADING, ALIGNLEFT, FONTBOLD)
    
    'syntax for condition check: XopY,
    'X means old cell value
    'X can also be [col,row], means a cell value col,row
    'op can be (==, <>, >>, <<, GT, LT),
    'Y can be ~ means new cell value,
    'Y can be a constant number,
    'Y can be [col,row] means a cell value in col,row
    'CellStyle:
    'COLHEADING means the cell value is a caption for column heading
    'ALIGNLEFT means the left justified the cell value
    'FONTBOLD means show the fonts in bold typeface
    
    If (i + 1) Mod 3 <> 0 And (i + 1) Mod 2 <> 0 Then
        Exit Sub
    End If
    
    m_nCol = nCol
    
    If IsNumeric(aryColorCond(0)(1)) Then
        m_nCondCount = (i + 1) / 3 - 1
    Else
        m_nCondCount = (i + 1) / 2 - 1
    End If
    
    ReDim m_sCondition(ConditionCount)
    ReDim m_lForeColor(ConditionCount)
    ReDim m_lBackColor(ConditionCount)
    ReDim m_sCellStyle(ConditionCount)
    
    j = -1
    
    If IsNumeric(aryColorCond(0)(1)) Then
        For i = 0 To ConditionCount
            j = j + 1
            m_sCondition(i) = aryColorCond(0)(j)
            j = j + 1
            m_lForeColor(i) = aryColorCond(0)(j)
            j = j + 1
            m_lBackColor(i) = aryColorCond(0)(j)
        Next
    Else
        For i = 0 To ConditionCount
            j = j + 1
            m_sCondition(i) = aryColorCond(0)(j)
            j = j + 1
            m_sCellStyle(i) = aryColorCond(0)(j)
        Next
    End If
End Sub

Public Sub AppendCondition(objNewColorCondition As clsColorCondition)
    Dim i As Integer
    Dim j As Integer
    Dim sCondition As String
    Dim lForeColor As Long
    Dim lBackColor As Long
    Dim sCellStyle As String
    
    If objNewColorCondition.ConditionCount + 1 <= 0 Then
        Exit Sub
    End If
    
    i = -1
    
    On Error Resume Next
    i = UBound(m_sCondition)
    On Error GoTo 0
    
    j = i + objNewColorCondition.ConditionCount + 1
    
    ReDim Preserve m_sCondition(j)
    ReDim Preserve m_lForeColor(j)
    ReDim Preserve m_lBackColor(j)
    ReDim Preserve m_sCellStyle(j)
    
    j = -1
    
    For i = i + 1 To i + objNewColorCondition.ConditionCount + 1
        j = j + 1
        If objNewColorCondition.GetCondition(j, sCondition, lForeColor, lBackColor, sCellStyle) Then
            m_sCondition(i) = sCondition
            m_lForeColor(i) = lForeColor
            m_lBackColor(i) = lBackColor
            m_sCellStyle(i) = sCellStyle
        End If
    Next
End Sub

Property Get ConditionCount() As Integer
    ConditionCount = m_nCondCount
End Property

Public Function GetCondition(ndx As Integer, sCondition As String, lForeColor As Long, lBackColor As Long, sCellStyle As String) As Boolean
    On Error GoTo errTrap
    
    sCondition = m_sCondition(ndx)
    lForeColor = m_lForeColor(ndx)
    lBackColor = m_lBackColor(ndx)
    sCellStyle = m_sCellStyle(ndx)
    
    GetCondition = True
    Exit Function
    
errTrap:
    'no nothing
    MsgBox "Error occurred in GetCondition()"
End Function

Private Sub Class_Terminate()
    ReDim m_sCondition(0)
    ReDim m_lForeColor(0)
    ReDim m_lBackColor(0)
    ReDim m_sCellStyle(0)
End Sub
