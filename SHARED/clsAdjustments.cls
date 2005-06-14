VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAdjustments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private colAdjustments As New Collection

Public Function Add(oAdjustment As clsAdjustment)
Attribute Add.VB_UserMemId = 0
    colAdjustments.Add oAdjustment
End Function

Public Function Count() As Long
   Count = colAdjustments.Count
End Function

Public Sub Delete(ByVal Index As Variant)
   colAdjustments.Remove Index
End Sub

Public Function Item(ByVal Index As Variant) As clsAdjustment
   Set Item = colAdjustments.Item(Index)
End Function

' NewEnum must return the IUnknown interface of a
' collection's enumerator.
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
   Set NewEnum = colAdjustments.[_NewEnum]
End Function

