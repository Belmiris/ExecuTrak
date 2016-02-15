VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DetachedRecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_records As Collection
Private m_index As Long
'

Public Property Get BOF() As Boolean
    
    BOF = Me.EOF
    
End Property

Public Property Get EOF() As Boolean
    
    EOF = True
    If Not m_records Is Nothing Then
        If m_records.Count > 0 Then
            If m_index <= m_records.Count Then
                EOF = False
            End If
        End If
    End If
    
End Property

Public Property Get Fields() As Collection
Attribute Fields.VB_UserMemId = 0
    
    If m_records Is Nothing Then
        Err.Raise -5352, "DetachedRecordset.Fields", "No records"
        Exit Property
    End If
    
    If m_index < 0 Then
        Err.Raise -5353, "DetachedRecordset.Fields", "No current record"
        Exit Property
    End If
    
    If m_index > m_records.Count Then
        Err.Raise -5354, "DetachedRecordset.Fields", "Record index past end of collection"
        Exit Property
    End If
    
    Set Fields = m_records(m_index)

End Property

Public Property Get RecordCount() As Long
    
    If Not m_records Is Nothing Then
        RecordCount = m_records.Count
    End If

End Property

Private Sub Class_Initialize()
'    Dim colFields As Collection
        
    ' Keep
    Set m_records = New Collection
    m_index = 1
    
'    ' Testing - OLD
'    Set colFields = New Collection
'    colFields.Add "Value 1.1", "Field1"
'    colFields.Add "Value 1.2", "Field2"
'    colFields.Add "Value 1.3", "Field3"
'    m_records.Add colFields
'
'    Set colFields = New Collection
'    colFields.Add "Value 2.1", "Field1"
'    colFields.Add "Value 2.2", "Field2"
'    colFields.Add "Value 2.3", "Field3"
'    m_records.Add colFields
    
'    m_index = 1
    
'    ' Testing - NEW
'    Dim drsItem As DetachedRecordsetItem
'
'    ' Record 1
'    Set colFields = New Collection
'
'    Set drsItem = New DetachedRecordsetItem
'    drsItem.SetProperties "Field1", "Value 1.1"
'    colFields.Add drsItem, "Field1"
'
'    Set drsItem = New DetachedRecordsetItem
'    drsItem.SetProperties "Field2", "Value 1.2"
'    colFields.Add drsItem, "Field2"
'
'    Set drsItem = New DetachedRecordsetItem
'    drsItem.SetProperties "Field3", "Value 1.3"
'    colFields.Add drsItem, "Field3"
'
'    m_records.Add colFields
'
'    ' Record 2
'    Set colFields = New Collection
'
'    Set drsItem = New DetachedRecordsetItem
'    drsItem.SetProperties "Field1", "Value 2.1"
'    colFields.Add drsItem, "Field1"
'
'    Set drsItem = New DetachedRecordsetItem
'    drsItem.SetProperties "Field2", "Value 2.2"
'    colFields.Add drsItem, "Field2"
'
'    Set drsItem = New DetachedRecordsetItem
'    drsItem.SetProperties "Field3", "Value 2.3"
'    colFields.Add drsItem, "Field3"
'
'    m_records.Add colFields
'
'    m_index = 1
    
End Sub

Public Sub CopyCurrentRecord(rsTemp As Recordset)
    Dim colFields As Collection
    Dim drsItem As DetachedRecordsetItem
    Dim fieldName As String
    Dim idxField As Long
    
    If m_records Is Nothing Then
        Set m_records = New Collection
    End If
    
    If rsTemp Is Nothing Then _
        Err.Raise -1000, "DetachedRecordset.CopyRecordset", "Empty rsTemp passed"
        
    Set colFields = New Collection
    For idxField = 0 To (rsTemp.Fields.Count - 1)
        Set drsItem = New DetachedRecordsetItem
        drsItem.Populate rsTemp, idxField
        colFields.Add drsItem, drsItem.Name
    Next
    
    m_records.Add colFields
    
End Sub

Public Sub MoveFirst()

    m_index = 1

End Sub

Public Sub MoveLast()

    If Not m_records Is Nothing Then
        If m_records.Count > 0 Then
            m_index = m_records.Count
        Else
            m_index = 1
        End If
    End If

End Sub

Public Sub MoveNext()
    Dim idxNew As Long
    
    idxNew = m_index + 1
    If Not m_records Is Nothing Then
        If idxNew <= (m_records.Count + 1) Then
            m_index = idxNew
        End If
    End If
    
End Sub

