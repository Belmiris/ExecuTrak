VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsColorCondiition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This Class is required when USE_ARRAY is set to 0
'When USE_ARRAY is set to -1, this Class is not required.

Option Explicit

Public m_nCol As Integer
Public m_nCondCount As Integer
Private m_sCondition() As String
Private m_lForeColor() As Long
Private m_lBackColor() As Long
'

Public Sub AddCondition(nCol As Integer, ParamArray aryColorCond() As Variant)
    Dim i As Integer
    
    i = -1
    
    On Error Resume Next
    i = UBound(aryColorCond(0))
    On Error GoTo 0
    
    If i < 0 Then
        Exit Sub
    End If
    
    'need to have 3 parm in a group
    'namely Condition, ForeColor, and BackColor
    
    'syntax for condition check: XopY,
    'X means old cell value, op can be (==, <>, >>, <<, GT, LT),
    'Y can be ~ means new cell value,
    'Y can be a constant number,
    'Y can be [Z] means a cell value in ColZ
    
    If (i + 1) Mod 3 <> 0 Then
        Exit Sub
    End If
    
    m_nCol = nCol
    m_nCondCount = (i + 1) / 3 - 1
    
    ReDim m_sCondition(ConditionCount)
    ReDim m_lForeColor(ConditionCount)
    ReDim m_lBackColor(ConditionCount)
    
    For i = 0 To ConditionCount
        m_sCondition(i) = aryColorCond(0)(i)
        m_lForeColor(i) = aryColorCond(0)(i + 1)
        m_lBackColor(i) = aryColorCond(0)(i + 2)
    Next
End Sub

Property Get ConditionCount() As Integer
    ConditionCount = m_nCondCount
End Property

Public Function GetCondition(ndx As Integer, sCondition As String, lForeColor As Long, lBackColor As Long) As Boolean
    On Error GoTo errTrap
    
    sCondition = m_sCondition(ndx)
    lForeColor = m_lForeColor(ndx)
    lBackColor = m_lBackColor(ndx)
    
    GetCondition = True
    Exit Function
    
errTrap:
    'no nothing
    MsgBox "Error occurred in GetCondition()"
End Function
