VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDDSingle"
Attribute VB_Creatable = True
Attribute VB_Exposed = False
Option Explicit
'New
    Private txtBox As Object
    Private cmdButton As Object
    
    Private sColCaption As String
    Private sRegPattern As String
    Private nFieldType As Integer
    Private nColumnWidth() As Integer
    Private nColumnCount As Integer
'Old
    Private GridArray() As Variant
    Private nDimension As Integer

    Private nTableLeft As Integer
    Private nTableTop As Integer
    Private nTableWidth As Integer
    Private nVisibleRows As Integer
    Private nPosition As Integer
    Private nHeadLines As Integer
    Private sTextSource As String
Public Sub AddColumn(nW As Integer)
    If nColumnCount = 0 Then
        ReDim nColumnWidth(nColumnCount)
    Else
        ReDim Preserve nColumnWidth(nColumnCount)
    End If
    nColumnWidth(nColumnCount) = nW
    nColumnCount = nColumnCount + 1
End Sub


Property Get ColumnCount() As Integer
    ColumnCount = nColumnCount
End Property

Property Let ColumnCount(nCol As Integer)
    nColumnCount = nCol
End Property

Public Function Dimension() As Integer
    Dimension = nDimension
End Function

Public Function RowCount() As Integer
    Select Case nDimension
        Case 1
            RowCount = UBound(GridArray) + 1
        Case 2
            RowCount = UBound(GridArray, 2) + 1
    End Select
End Function

Property Let Position(nPos As Integer)
    nPosition = nPos
End Property
Property Get Position() As Integer
    Position = nPosition
End Property


Public Sub GetArray(A() As Variant)
    Dim i As Integer
    Select Case nDimension
        Case 1
            ReDim A(0, UBound(GridArray))
            For i = 0 To UBound(GridArray)
                A(0, i) = GridArray(i)
            Next
        Case 2
            Dim j As Integer
            ReDim A(UBound(GridArray), UBound(GridArray, 2))
            For i = 0 To UBound(GridArray, 2)
                For j = 0 To UBound(GridArray, 1)
                    A(j, i) = GridArray(j, i)
                Next j
            Next i
    End Select
End Sub

Public Sub SetArray(A() As Variant)
    Dim i As Integer

    If nColumnCount <= 1 Then
        nDimension = 1
        ReDim GridArray(UBound(A))
        For i = 0 To UBound(A)
            GridArray(i) = A(i)
        Next
    Else
        nDimension = 2
        Dim j As Integer
        ReDim GridArray(UBound(A), UBound(A, 2))
        For i = 0 To UBound(A, 2)
            For j = 0 To UBound(A, 1)
                GridArray(j, i) = A(j, i)
            Next j
        Next i
    End If

End Sub

Property Let TableLeft(L As Integer)
    nTableLeft = L
End Property
Property Get TableLeft() As Integer
    TableLeft = nTableLeft
End Property



Property Let TableTop(T As Integer)
    nTableTop = T
End Property
Property Get TableTop() As Integer
    TableTop = nTableTop
End Property

Property Let TableWidth(w As Integer)
    nTableWidth = w
End Property
Property Get TableWidth() As Integer
    TableWidth = nTableWidth
End Property
Property Get HeadLines() As Integer
    HeadLines = nHeadLines
End Property

Property Let HeadLines(nLines As Integer)
    nHeadLines = nLines
End Property

Property Let TextSource(sSrc As String)
    sTextSource = sSrc
End Property
Property Get TextSource() As String
    TextSource = sTextSource
End Property

Property Let VisibleRows(nRows As Integer)
    nVisibleRows = nRows
End Property
Property Get VisibleRows() As Integer
    VisibleRows = nVisibleRows
End Property
Property Let Caption(sCap As String)
    sColCaption = sCap
End Property
Property Get Caption() As String
    Caption = sColCaption
End Property



Property Set Button(b As Object)
    Set cmdButton = b
End Property
Property Get Button() As Object
    Set Button = cmdButton
End Property

Property Get ComboType() As Integer
    ComboType = nFieldType
End Property

Property Let ComboType(nType As Integer)
    
    nFieldType = nType

End Property
Property Get ColWidth(nIndex As Integer) As Integer
    If nIndex >= 0 And nIndex < nColumnCount Then
        ColWidth = nColumnWidth(nIndex)
    Else
        ColWidth = -1
    End If
End Property

Property Let ColWidth(nIndex As Integer, nWidth As Integer)
    If nIndex >= 0 And nIndex < nColumnCount Then
        nColumnWidth(nIndex) = nWidth
    End If
End Property


Property Get RegExpPattern() As String
    RegExpPattern = sRegPattern
End Property
Property Let RegExpPattern(sPattern As String)
    sRegPattern = sPattern
End Property


Property Set TextBox(Box As Object)
    Set txtBox = Box
End Property

Property Get TextBox() As Object
    Set TextBox = txtBox
End Property

Private Sub Class_Initialize()
    nColumnCount = 0
End Sub


