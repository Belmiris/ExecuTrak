VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBBLineItem"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit
    
    Private m_colFormula As Collection
    Private m_sProductCode As String
    Private m_lProductLink As Long
    Private m_dQuantity As Double

    Private m_nFormulaCount As Integer
Public Sub AddBlendItem(sCode As String, lLink As Long, dQuantity As Double)
    Dim clsTemp As New clsBBFormula
    If m_colFormula Is Nothing Then
        Set m_colFormula = New Collection
        m_nFormulaCount = 1
    End If
    clsTemp.Code = sCode
    clsTemp.Link = lLink
    clsTemp.Quantity = dQuantity
    clsTemp.FormulaNumber = m_nFormulaCount
    m_colFormula.AddItem clsTemp
    
End Sub

Property Get Code() As String
    Code = m_sProductCode
End Property
Property Let Code(sCode As String)
    m_sProductCode = sCode
End Property

Public Function ConvertToFraction() As Boolean
    If m_colFormula Is Nothing Then
        Exit Function
    End If
    
'    If m_colFormula.count = 1 Then
        'If the count is 1, the product is not blended
'        sProductLink = m_colFormula.Link
'        sProductCode = m_colFormula.Code
'        Set m_colFormula = Nothing
'        Exit Function
'    End If

    Dim i As Integer
    Dim dSum As Double
    
    dSum = 0
    With m_colFormula
        For k = 1 To m_nFormulaCount
            For i = 1 To .count
                If .FormulaNumber = k Then
                    dSum = dSum + .Quantity
                End If
            Next
            If dSum <> 0 Then
                For i = 1 To .count
                    If .FormulaNumber = k Then
                        .Quantity = .Quantity / dSum
                    End If
                Next
                ConvertToFraction = True
            Else
                ConvertToFraction = False
            End If
        Next k
    End With
End Function

Public Sub CreateFormula()
    If m_colFormula Is Nothing Then
        Set m_colFormula = New Collection
    End If
    m_nFormulaCount = m_nFormulaCount + 1
End Sub

Property Set Formula(colTemp As Collection)
    Set m_colFormula = colTemp
End Property


Property Get Formula() As Collection
    Set Formula = m_colFormula
End Property

Property Let Link(lLink As Long)
    m_lProductLink = lLink
End Property

Property Get Link() As Long
    Link = m_lProductLink
End Property

Property Let Quantity(dQty As Double)
    m_dQuantity = dQty
End Property

Property Get Quantity() As Double
    Quantity = m_dQuantity
End Property

Private Sub Class_Initialize()
    Set m_colFormula = Nothing
End Sub


