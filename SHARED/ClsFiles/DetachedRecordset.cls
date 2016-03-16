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

' When using the DetachedRecordset class to replace DAO Recordset objects
' in a program, make sure the the program does not use any unsupported Recordset
' Properties or functions.
' THE COMPILER WILL NOT CATCH IT IF AN UNSUPPORTED FUNCTION / PROPERTY IS CALLED!
' THERE WILL BE AN ERROR AT RUNTIME - I.E. WHEN A CUSTOMER MAY BE USING IT.

'    rsTemp.AbsolutePosition            * Not supported
'    rsTemp.AddNew                      * Not supported
'    rsTemp.BatchCollisionCount         * Not supported
'    rsTemp.BatchCollisions             * Not supported
'    rsTemp.BatchSize                   * Not supported
'    'rsTemp.BOF
'    rsTemp.Bookmark                    * Not supported
'    rsTemp.Bookmarkable                * Not supported
'    rsTemp.CacheSize                   * Not supported
'    rsTemp.CacheStart                  * Not supported
'    rsTemp.Cancel                      * Not supported
'    rsTemp.CancelUpdate                * Not supported
'    rsTemp.Clone                       * Not supported
'    rsTemp.Close                       * Not supported
'    rsTemp.CopyQueryDef                * Not supported
'    rsTemp.DateCreated                 * Not supported
'    rsTemp.Delete                      * Not supported
'    rsTemp.Edit                        * Not supported
'    rsTemp.EditMode                    * Not supported
'    'rsTemp.EOF
'    'rsTemp.Fields
'    rsTemp.FillCache                   * Not supported
'    rsTemp.Filter                      * Not supported
'    rsTemp.FindFirst                   * Not supported
'    rsTemp.FindLast                    * Not supported
'    rsTemp.FindNext                    * Not supported
'    rsTemp.FindPrevious                * Not supported
'    rsTemp.GetRows                     * Not supported
'    rsTemp.Index                       * Not supported
'    rsTemp.LastModified                * Not supported
'    rsTemp.LastUpdated                 * Not supported
'    rsTemp.LockEdits                   * Not supported
'    rsTemp.Move
'    'rsTemp.MoveFirst
'    'rsTemp.MoveLast
'    'rsTemp.MoveNext
'    rsTemp.MovePrevious                * Not supported
'    rsTemp.Name                        * Not supported
'    rsTemp.NextRecordset               * Not supported
'    rsTemp.NoMatch                     * Not supported
'    rsTemp.OpenRecordset               * Not supported
'    rsTemp.PercentPosition             * Not supported
'    rsTemp.Properties                  * Not supported
'    'rsTemp.RecordCount
'    rsTemp.RecordStatus                * Not supported
'    rsTemp.Requery                     * Not supported
'    rsTemp.Restartable                 * Not supported
'    rsTemp.Seek                        * Not supported
'    rsTemp.Sort                        * Not supported
'    rsTemp.StillExecuting              * Not supported
'    rsTemp.Transactions                * Not supported
'    rsTemp.Type                        * Not supported
'    rsTemp.Updatable                   * Not supported
'    rsTemp.Update                      * Not supported
'    rsTemp.UpdateOptions               * Not supported
'    rsTemp.ValidationRule              * Not supported
'    rsTemp.ValidationText              * Not supported

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

Public Sub CloseRecordset()
    Class_Initialize
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
        Err.Raise -1000, "DetachedRecordset.CopyCurrentRecord", "DetachedRecordset.CopyCurrentRecord Empty rsTemp passed"
        
    Set colFields = New Collection
    For idxField = 0 To (rsTemp.Fields.Count - 1)
        Set drsItem = New DetachedRecordsetItem
        drsItem.Populate rsTemp, idxField
        colFields.Add drsItem, drsItem.Name
    Next
    
    m_records.Add colFields
    
End Sub

Public Sub CopyRecordset(rsTemp As Recordset, Optional bMoveFirst As Boolean = True)

    If rsTemp Is Nothing Then _
        Err.Raise -2000, "DetachedRecordset.CopyRecordset", "DetachedRecordset.CopyRecordset Empty rsTemp passed"

    Class_Initialize
    
    If rsTemp.RecordCount > 0 Then
        If bMoveFirst Then rsTemp.MoveFirst
    
        While Not rsTemp.EOF
            CopyCurrentRecord rsTemp
            rsTemp.MoveNext
        Wend
    End If
    
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

' *******************************************************************
' Unsuppored methods / properties
' Some are not legal names (VB6 sucks)
' *******************************************************************

'Public Function AbsolutePosition(Optional ParamArray xarrayx() As Variant) As Variant
'    Err.Raise -1313, "DetachedRecordset", "The method 'AbsolutePosition' is not supported by DetachedRecordset"
'End Function
 
'Public Function AddNew(ParamArray xarrayx() Qs Variant) As Variant
'    Err.Raise -1313, "DetachedRecordset", "The method 'AddNew' is not supported by DetachedRecordset"
'End Function
 
Public Function BatchCollisionCount(ParamArray xarrayx() As Variant) As Variant
    Err.Raise -1313, "DetachedRecordset", "The method 'BatchCollisionCount' is not supported by DetachedRecordset"
End Function
 
Public Function BatchCollisions(ParamArray xarrayx() As Variant) As Variant
    Err.Raise -1313, "DetachedRecordset", "The method 'BatchCollisions' is not supported by DetachedRecordset"
End Function
 
Public Function BatchSize(ParamArray xarrayx() As Variant) As Variant
    Err.Raise -1313, "DetachedRecordset", "The method 'BatchSize' is not supported by DetachedRecordset"
End Function
 
Public Function Bookmark(ParamArray xarrayx() As Variant) As Variant
    Err.Raise -1313, "DetachedRecordset", "The method 'Bookmark' is not supported by DetachedRecordset"
End Function
 
Public Function Bookmarkable(ParamArray xarrayx() As Variant) As Variant
    Err.Raise -1313, "DetachedRecordset", "The method 'Bookmarkable' is not supported by DetachedRecordset"
End Function
 
Public Function CacheSize(ParamArray xarrayx() As Variant) As Variant
    Err.Raise -1313, "DetachedRecordset", "The method 'CacheSize' is not supported by DetachedRecordset"
End Function
 
Public Function CacheStart(ParamArray xarrayx() As Variant) As Variant
    Err.Raise -1313, "DetachedRecordset", "The method 'CacheStart' is not supported by DetachedRecordset"
End Function
 
Public Function Cancel(ParamArray xarrayx() As Variant) As Variant
    Err.Raise -1313, "DetachedRecordset", "The method 'Cancel' is not supported by DetachedRecordset"
End Function
 
Public Function CancelUpdate(ParamArray xarrayx() As Variant) As Variant
    Err.Raise -1313, "DetachedRecordset", "The method 'CancelUpdate' is not supported by DetachedRecordset"
End Function
 
Public Function Clone(ParamArray xarrayx() As Variant) As Variant
    Err.Raise -1313, "DetachedRecordset", "The method 'Clone' is not supported by DetachedRecordset"
End Function
 
'Public Function Close(ParamArray xarrayx() As Variant) As Variant
'    Err.Raise -1313, "DetachedRecordset", "The method 'Close' is not supported by DetachedRecordset"
'End Function
 
Public Function CopyQueryDef(ParamArray xarrayx() As Variant) As Variant
    Err.Raise -1313, "DetachedRecordset", "The method 'CopyQueryDef' is not supported by DetachedRecordset"
End Function
 
Public Function DateCreated(ParamArray xarrayx() As Variant) As Variant
    Err.Raise -1313, "DetachedRecordset", "The method 'DateCreated' is not supported by DetachedRecordset"
End Function
 
Public Function Delete(ParamArray xarrayx() As Variant) As Variant
    Err.Raise -1313, "DetachedRecordset", "The method 'Delete' is not supported by DetachedRecordset"
End Function
 
Public Function Edit(ParamArray xarrayx() As Variant) As Variant
    Err.Raise -1313, "DetachedRecordset", "The method 'Edit' is not supported by DetachedRecordset"
End Function
 
Public Function EditMode(ParamArray xarrayx() As Variant) As Variant
    Err.Raise -1313, "DetachedRecordset", "The method 'EditMode' is not supported by DetachedRecordset"
End Function
 
Public Function FillCache(ParamArray xarrayx() As Variant) As Variant
    Err.Raise -1313, "DetachedRecordset", "The method 'FillCache' is not supported by DetachedRecordset"
End Function
 
Public Function Filter(ParamArray xarrayx() As Variant) As Variant
    Err.Raise -1313, "DetachedRecordset", "The method 'Filter' is not supported by DetachedRecordset"
End Function
 
Public Function FindFirst(ParamArray xarrayx() As Variant) As Variant
    Err.Raise -1313, "DetachedRecordset", "The method 'FindFirst' is not supported by DetachedRecordset"
End Function
 
Public Function FindLast(ParamArray xarrayx() As Variant) As Variant
    Err.Raise -1313, "DetachedRecordset", "The method 'FindLast' is not supported by DetachedRecordset"
End Function
 
Public Function FindNext(ParamArray xarrayx() As Variant) As Variant
    Err.Raise -1313, "DetachedRecordset", "The method 'FindNext' is not supported by DetachedRecordset"
End Function
 
Public Function FindPrevious(ParamArray xarrayx() As Variant) As Variant
    Err.Raise -1313, "DetachedRecordset", "The method 'FindPrevious' is not supported by DetachedRecordset"
End Function
 
Public Function GetRows(ParamArray xarrayx() As Variant) As Variant
    Err.Raise -1313, "DetachedRecordset", "The method 'GetRows' is not supported by DetachedRecordset"
End Function
 
Public Function Index(ParamArray xarrayx() As Variant) As Variant
    Err.Raise -1313, "DetachedRecordset", "The method 'Index' is not supported by DetachedRecordset"
End Function
 
Public Function LastModified(ParamArray xarrayx() As Variant) As Variant
    Err.Raise -1313, "DetachedRecordset", "The method 'LastModified' is not supported by DetachedRecordset"
End Function
 
Public Function LastUpdated(ParamArray xarrayx() As Variant) As Variant
    Err.Raise -1313, "DetachedRecordset", "The method 'LastUpdated' is not supported by DetachedRecordset"
End Function
 
Public Function LockEdits(ParamArray xarrayx() As Variant) As Variant
    Err.Raise -1313, "DetachedRecordset", "The method 'LockEdits' is not supported by DetachedRecordset"
End Function
 
Public Function Move(ParamArray xarrayx() As Variant) As Variant
    Err.Raise -1313, "DetachedRecordset", "The method 'Move' is not supported by DetachedRecordset"
End Function
 
Public Function MovePrevious(ParamArray xarrayx() As Variant) As Variant
    Err.Raise -1313, "DetachedRecordset", "The method 'MovePrevious' is not supported by DetachedRecordset"
End Function
 
Public Function Name(ParamArray xarrayx() As Variant) As Variant
    Err.Raise -1313, "DetachedRecordset", "The method 'Name' is not supported by DetachedRecordset"
End Function
 
Public Function NextRecordset(ParamArray xarrayx() As Variant) As Variant
    Err.Raise -1313, "DetachedRecordset", "The method 'NextRecordset' is not supported by DetachedRecordset"
End Function
 
Public Function NoMatch(ParamArray xarrayx() As Variant) As Variant
    Err.Raise -1313, "DetachedRecordset", "The method 'NoMatch' is not supported by DetachedRecordset"
End Function
 
Public Function OpenRecordset(ParamArray xarrayx() As Variant) As Variant
    Err.Raise -1313, "DetachedRecordset", "The method 'OpenRecordset' is not supported by DetachedRecordset"
End Function
 
Public Function PercentPosition(ParamArray xarrayx() As Variant) As Variant
    Err.Raise -1313, "DetachedRecordset", "The method 'PercentPosition' is not supported by DetachedRecordset"
End Function
 
Public Function Properties(ParamArray xarrayx() As Variant) As Variant
    Err.Raise -1313, "DetachedRecordset", "The method 'Properties' is not supported by DetachedRecordset"
End Function
 
Public Function RecordStatus(ParamArray xarrayx() As Variant) As Variant
    Err.Raise -1313, "DetachedRecordset", "The method 'RecordStatus' is not supported by DetachedRecordset"
End Function
 
Public Function Requery(ParamArray xarrayx() As Variant) As Variant
    Err.Raise -1313, "DetachedRecordset", "The method 'Requery' is not supported by DetachedRecordset"
End Function
 
Public Function Restartable(ParamArray xarrayx() As Variant) As Variant
    Err.Raise -1313, "DetachedRecordset", "The method 'Restartable' is not supported by DetachedRecordset"
End Function
 
'Public Function Seek(ParamArray xarrayx() As Variant) As Variant
'    Err.Raise -1313, "DetachedRecordset", "The method 'Seek' is not supported by DetachedRecordset"
'End Function
 
Public Function Sort(ParamArray xarrayx() As Variant) As Variant
    Err.Raise -1313, "DetachedRecordset", "The method 'Sort' is not supported by DetachedRecordset"
End Function
 
Public Function StillExecuting(ParamArray xarrayx() As Variant) As Variant
    Err.Raise -1313, "DetachedRecordset", "The method 'StillExecuting' is not supported by DetachedRecordset"
End Function
 
Public Function Transactions(ParamArray xarrayx() As Variant) As Variant
    Err.Raise -1313, "DetachedRecordset", "The method 'Transactions' is not supported by DetachedRecordset"
End Function
 
'Public Function Type(ParamArray xarrayx() As Variant) As Variant
'    Err.Raise -1313, "DetachedRecordset", "The method 'Type' is not supported by DetachedRecordset"
'End Function
 
Public Function Updatable(ParamArray xarrayx() As Variant) As Variant
    Err.Raise -1313, "DetachedRecordset", "The method 'Updatable' is not supported by DetachedRecordset"
End Function
 
Public Function Update(ParamArray xarrayx() As Variant) As Variant
    Err.Raise -1313, "DetachedRecordset", "The method 'Update' is not supported by DetachedRecordset"
End Function
 
Public Function UpdateOptions(ParamArray xarrayx() As Variant) As Variant
    Err.Raise -1313, "DetachedRecordset", "The method 'UpdateOptions' is not supported by DetachedRecordset"
End Function
 
Public Function ValidationRule(ParamArray xarrayx() As Variant) As Variant
    Err.Raise -1313, "DetachedRecordset", "The method 'ValidationRule' is not supported by DetachedRecordset"
End Function
 
Public Function ValidationText(ParamArray xarrayx() As Variant) As Variant
    Err.Raise -1313, "DetachedRecordset", "The method 'ValidationText' is not supported by DetachedRecordset"
End Function
 

