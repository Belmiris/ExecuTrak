VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTGCursorControl"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit
    
    Private Const STATUS_INIT = 0
    Private Const STATUS_FILLED = 1

    Private Type tpColumns
        m_nStatus As Integer
        m_nNextCols() As Integer
        m_nNextCount As Integer
    End Type
    
    Private m_clsEdit As Object
    Private m_nLastCol As Integer
    Private m_nLastRow As Integer
    Private m_nSeries As Integer
    Private m_bRowColChanging As Boolean
    
    Private m_arryControls() As tpColumns
    Private m_nCount As Integer
    
Public Sub AddItem(nCol As Integer, _
                   ParamArray vAssc() As Variant)
    Dim i As Integer
    
    If m_nCount = 0 Then
        ReDim m_arryControls(0)
    Else
        ReDim Preserve m_arryControls(m_nCount)
    End If
    
    With m_arryControls(m_nCount)
        .m_nStatus = STATUS_INIT
        .m_nNextCount = UBound(vAssc) + 1
        If .m_nNextCount >= 0 Then
            ReDim .m_nNextCols(.m_nNextCount)
            .m_nNextCols(0) = nCol
            For i = 1 To .m_nNextCount
                .m_nNextCols(i) = vAssc(i - 1)
            Next
        End If
    End With
    m_nCount = m_nCount + 1
End Sub

Property Set EditClass(objTemp As Object)
    Set m_clsEdit = objTemp
    m_bRowColChanging = False
End Property


Private Function fnColumnSet() As Integer
    Dim i As Integer
    
    fnColumnSet = -1
    With m_clsEdit.Table
        For i = 0 To m_nCount - 1
            If fnListIndex(i, .col) >= 0 Then
'            If .col = m_arryControls(i).m_nCurrent _
'            Or .col = m_arryControls(i).m_nNextOne _
'            Or .col = m_arryControls(i).m_nNextTwo Then
                fnColumnSet = i
                Exit For
            End If
        Next i
    End With

End Function

Private Function fnListIndex(nIndex As Integer, ByVal nCol As Integer) As Integer
    
    Dim i As Integer
    
    fnListIndex = -1
    With m_arryControls(nIndex)
        For i = 0 To .m_nNextCount
            If nCol = .m_nNextCols(i) Then
                fnListIndex = i
                Exit For
            End If
        Next
    End With
End Function

Public Sub InputSelected()
Exit Sub
    Dim nRet As Integer
    nRet = fnColumnSet
    If nRet >= 0 Then
        m_nSeries = nRet
        m_arryControls(nRet).m_nStatus = STATUS_FILLED
        m_clsEdit.GotoNextColumn m_nLastCol
    End If
End Sub

Public Function RowColChange() As Boolean
Exit Function
    Dim nRet As Integer
    
    If m_bRowColChanging Then
'        Exit Function
    End If
    m_bRowColChanging = True

    nRet = fnColumnSet
    
    RowColChange = False
    If nRet >= 0 Then
        RowColChange = True
        If m_arryControls(nRet).m_nStatus = STATUS_FILLED Then
            If m_nSeries = nRet Then
                If fnListIndex(nRet, m_clsEdit.Table.col) >= 0 Then
                    m_clsEdit.GotoNextColumn m_nLastCol
                End If
            End If
            m_arryControls(nRet).m_nStatus = STATUS_INIT
        End If
    End If
    m_nSeries = nRet
    m_nLastCol = m_clsEdit.Table.col
    m_nLastRow = m_clsEdit.Table.Row
    m_bRowColChanging = False
End Function

