VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DetachedRecordsetItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Name As String
Private m_Value As Variant
Private m_Populated As Boolean

Private Sub Class_Initialize()
    m_Populated = False
End Sub

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
    Value = m_Value
End Property

Public Sub Populate(rsTemp As Recordset, fieldIndex As Long)
    
    If m_Populated Then _
        Err.Raise -1000, "DetachedRecordsetItem.Populate", "DetachedRecordsetItem already populated"
    
    If rsTemp Is Nothing Then _
        Err.Raise -1001, "DetachedRecordsetItem.Populate", "Empty rsTemp passed to DetachedRecordsetItem.Populate"
    
    If rsTemp.EOF Then _
        Err.Raise -1002, "DetachedRecordsetItem.Populate", "rsTemp passed to DetachedRecordsetItem.Populate is EOF"
    
    If fieldIndex < 0 Then _
        Err.Raise -1003, "DetachedRecordsetItem.Populate", "Field index passed to DetachedRecordsetItem.Populate is invalid (" & fieldIndex & ")"
    
    If fieldIndex >= rsTemp.Fields.Count Then _
        Err.Raise -1004, "DetachedRecordsetItem.Populate", "Field index passed to DetachedRecordsetItem.Populate is too large (" & fieldIndex & ")"
    
    m_Name = rsTemp.Fields(fieldIndex).Name
    m_Value = rsTemp.Fields(fieldIndex).Value
    
End Sub

Public Sub SetProperties(fieldName As String, fieldValue As Variant)

    If m_Populated Then _
        Err.Raise -1010, "DetachedRecordsetItem.Populate", "DetachedRecordsetItem already populated"

    m_Name = fieldName
    m_Value = fieldValue

End Sub
