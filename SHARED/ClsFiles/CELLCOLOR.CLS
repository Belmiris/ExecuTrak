VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCellColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_tgTable As TDBGrid
Private m_tgcEditor As clsTGSpreadSheet

#If USE_ARRAY Then
    Private Type CELL_COLOR
        nCol As Integer
        'syntax for condition check: XopY,
        'X means old cell value, op can be (==, <>, >>, <<, GT, LT),
        'Y can be ~ means new cell value,
        'Y can be a constant number,
        'Y can be [Z] means a cell value in ColZ
        sCondition As String
        lForeColor As Long
        lBackColor As Long
    End Type

    Private m_nCellColorCount As Integer
    Private m_aryCellColor() As CELL_COLOR
#Else
    Private colColorCondition As New Collection
#End If
'

Property Set Table(tbl As TDBGrid)
    Set m_tgTable = tbl
End Property

Property Get Table() As TDBGrid
    Set Table = m_tgTable
End Property

Property Set EditorClass(tgm As clsTGSpreadSheet)
    Set m_tgcEditor = tgm
    Set Table = m_tgcEditor.Table
End Property

Public Sub AddColumn4CellColor(nCol As Integer, ParamArray aryColorCond() As Variant)
    Dim i As Integer
    Dim j As Integer
    Dim n As Integer
        
    On Error GoTo errTrap
    
    i = -1
    
    On Error Resume Next
    i = UBound(aryColorCond)
    On Error GoTo 0
    
    If i < 0 Then
        Exit Sub
    End If
    
    'need to have 3 parm in a group
    'namely Condition, ForeColor, and BackColor
    
    'syntax for condition check: XopY,
    'X means old cell value
    'X can also be [col,row], means a cell value col,row
    'op can be (==, <>, >>, <<, GT, LT),
    'Y can be ~ means new cell value,
    'Y can be a constant number,
    'Y can be [col,row] means a cell value in col,row
    
    If (i + 1) Mod 3 <> 0 Then
        Exit Sub
    End If
    
    i = (i + 1) / 3
    
    #If Not USE_ARRAY Then
        'david 02/07/2004
        'TESTING PHASE: USE_COLLECTION
        'use collection to speed up the search
        'performance should be better than USE_ARRAY
        Dim objColorCondition As New clsColorCondiition
        objColorCondition.AddCondition nCol, aryColorCond
        
        colColorCondition.Add objColorCondition, "C" & nCol
    #Else
        n = m_nCellColorCount + i
        ReDim Preserve m_aryCellColor(n)
        
        j = -1
        For i = m_nCellColorCount + 1 To n
            With m_aryCellColor(i)
                .nCol = nCol
                j = j + 1
                .sCondition = aryColorCond(j)
                j = j + 1
                .lForeColor = aryColorCond(j)
                j = j + 1
                .lBackColor = aryColorCond(j)
            End With
        Next
        
        m_nCellColorCount = n
    #End If
    
    Exit Sub

errTrap:
    'do nothing
End Sub

Public Sub RemoveColumn4CellColor(nCol As Integer)
    On Error Resume Next
    colColorCondition.Remove "C" & nCol
End Sub

Property Get IsColumnForCellColor(nCol As Integer) As Boolean
    #If Not USE_ARRAY Then
        Dim objColorCondition As clsColorCondiition
        
        On Error Resume Next
        Set objColorCondition = colColorCondition.Item("C" & nCol)
        
        If Not objColorCondition Is Nothing Then
            IsColumnForCellColor = True
            Exit Property
        End If
    #Else
        Dim i As Integer
        
        For i = 0 To m_nCellColorCount
            If m_aryCellColor(i).nCol = nCol Then
                IsColumnForCellColor = True
                Exit Property
            End If
        Next

    #End If
End Property

Public Sub FetchCellStyle(ByVal Condition As Integer, _
                          ByVal Split As Integer, _
                          Bookmark As Variant, _
                          ByVal Col As Integer, _
                          ByVal CellStyle As Object)
    
    Dim lRow As Long
    Dim nPosi As Integer
    Dim sLeftSide As String
    Dim sOP As String
    Dim sRightSide As String
    Dim sLeftValue As String
    Dim sRightValue As String
    
    Dim nCol2 As Integer
    Dim lRow2 As Long
    
    Dim i As Integer
    Dim sCondition As String
    Dim lForeColor As Long
    Dim lBackColor As Long
    Dim bSetCellColor As Boolean
    
    On Error GoTo errTrap
    
    lRow = val(Bookmark)
    
    If lRow < 0 Then
        Exit Sub
    End If
    
    #If Not USE_ARRAY Then
        'david 02/07/2004
        'TESTING PHASE: USE_COLLECTION
        'use collection to speed up the search
        'performance should be better than USE_ARRAY
        
        Dim objColorCondition As clsColorCondiition
        'get the color condition
        On Error Resume Next
        Set objColorCondition = colColorCondition.Item("C" & Col)
        On Error GoTo 0
        
        If Not objColorCondition Is Nothing Then
            For i = 0 To objColorCondition.ConditionCount
                If Not objColorCondition.GetCondition(i, sCondition, lForeColor, lBackColor) Then
                    Exit Sub
                End If
                
                'syntax for condition check: XopY,
                'X means old cell value
                'X can also be [col,row], means a cell value col,row
                'op can be (==, <>, >>, <<, GT, LT),
                'Y can be ~ means new cell value,
                'Y can be a constant number,
                'Y can be [col,row] means a cell value in col,row
                
                'get the op
                sOP = "=="
                nPosi = InStr(sCondition, sOP)
                If nPosi <= 0 Then
                    sOP = "<>"
                    nPosi = InStr(sCondition, sOP)
                    If nPosi <= 0 Then
                        sOP = ">>"
                        nPosi = InStr(sCondition, sOP)
                        If nPosi <= 0 Then
                            sOP = "<<"
                            nPosi = InStr(sCondition, sOP)
                            If nPosi <= 0 Then
                                sOP = "GT"
                                nPosi = InStr(sCondition, sOP)
                                If nPosi <= 0 Then
                                    sOP = "LT"
                                    nPosi = InStr(sCondition, sOP)
                                End If
                            End If
                        End If
                    End If
                End If
                
                If nPosi <= 0 Then
                    'op is not valid
                    Exit Sub
                End If
                
                'get left hand side and right hand side
                sLeftSide = Left(sCondition, nPosi - 1)
                sRightSide = Mid(sCondition, nPosi + 2)
                
                If UCase(Left(sRightSide, 1)) = "~" Then   'compare old value to a new value in the same cell
                    sLeftValue = m_tgcEditor.OldValue & ""
                    sRightValue = m_tgcEditor.CellValue(Col, lRow) & ""
                Else  'compare to a cellvalue in another column [] or compare to a constant
                    If UCase(sLeftSide) = "X" Then
                        sLeftValue = m_tgcEditor.CellValue(Col, lRow) & ""
                    Else
                        sLeftValue = sLeftSide
                        
                        If Left(sLeftValue, 1) = "[" Then
                            'get another col, row
                            'remove ]
                            sLeftValue = Mid(sLeftValue, 2, Len(sLeftValue) - 2)
                            
                            nCol2 = InStr(sLeftValue, ",")
                            
                            If nCol2 > 0 Then
                                If Mid(sLeftValue, nCol2 + 1) = "" Then
                                    lRow2 = lRow
                                Else
                                    lRow2 = Mid(sLeftValue, nCol2 + 1)
                                End If
                                nCol2 = Left(sLeftValue, nCol2 - 1)
                            Else
                                lRow2 = lRow
                                nCol2 = sLeftValue
                            End If
                        
                            sLeftValue = m_tgcEditor.CellValue(nCol2, lRow2) & ""
                        Else  'is a constant
                            'okay
                        End If
                    End If
                    
                    sRightValue = sRightSide
                    
                    If UCase(Mid(sCondition, 4, 1)) = "[" Then
                        'get another col, row
                        'remove ]
                        sRightValue = Left(sRightValue, Len(sRightValue) - 1)
                        
                        nCol2 = InStr(sRightValue, ",")
                        
                        If nCol2 > 0 Then
                            If Mid(sRightValue, nCol2 + 1) = "" Then
                                lRow2 = lRow
                            Else
                                lRow2 = Mid(sRightValue, nCol2 + 1)
                            End If
                            nCol2 = Left(sRightValue, nCol2 - 1)
                        Else
                            lRow2 = lRow
                            nCol2 = sRightValue
                        End If
                    
                        sRightValue = m_tgcEditor.CellValue(nCol2, lRow2) & ""
                    Else  'is a constant
                        'okay
                    End If
                    
                    Select Case UCase(sOP)
                    Case "<>"  'string/numeric not equal
                        'If UCase(Mid(sCondition, 4, 1)) = "~" Then  'compare old value to a new value in the same cell
                        '    If m_tgcEditor.CellValueChanged(Col, lRow) Then
                        '        Exit For
                        '    End If
                        'Else
                            If Trim(sLeftValue) <> Trim(sRightValue) Then
                                Exit For
                            End If
                        'End If
                    Case "=="  'string/numeric equal
                        'If UCase(Mid(sCondition, 4, 1)) = "~" Then  'compare old value to a new value in the same cell
                        '    If Not m_tgcEditor.CellValueChanged(Col, lRow) Then
                        '        Exit For
                        '    End If
                        'Else
                            If Trim(sLeftValue) = Trim(sRightValue) Then
                                Exit For
                            End If
                        'End If
                    Case "<<"  'numeric less than
                        If val(sLeftValue) < val(sRightValue) Then
                            Exit For
                        End If
                    Case ">>"  'numeric greater than
                        If val(sLeftValue) > val(sRightValue) Then
                            Exit For
                        End If
                    Case "LT"  'string/date less than
                        If Trim(sLeftValue) < Trim(sRightValue) Then
                            Exit For
                        End If
                    Case "GT"  'string/date greater than
                        If Trim(sLeftValue) > Trim(sRightValue) Then
                            Exit For
                        End If
                    End Select
                End If
            Next
        
            'if condition is true, set the color
            If i <= objColorCondition.ConditionCount Then
                bSetCellColor = True
            End If
        End If
    #Else
        For i = 0 To m_nCellColorCount
            If Col = m_aryCellColor(i).nCol Then
                sCondition = m_aryCellColor(i).sCondition
                
                If UCase(Mid(sCondition, 4, 1)) = "~" Then  'compare old value to a new value in the same cell
                    sLeftValue = m_tgcEditor.OldValue & ""
                    sRightValue = m_tgcEditor.CellValue(Col, lRow) & ""
                Else  'compare to a cellvalue in another column [] or compare to a constant
                    sLeftValue = m_tgcEditor.CellValue(Col, lRow) & ""
                    sRightValue = Mid(sCondition, 5)
                    
                    If UCase(Mid(sCondition, 4, 1)) = "[" Then
                        'get another col, row
                        Dim nCol2 As Integer
                        Dim lRow2 As Long
                        
                        'remove ]
                        sRightValue = Left(sRightValue, Len(sRightValue) - 1)
                        
                        nCol2 = InStr(sRightValue, ",")
                        
                        If nCol2 > 0 Then
                            lRow2 = Mid(sRightValue, nCol2 + 1)
                            nCol2 = Left(sRightValue, nCol2 - 1)
                        Else
                            lRow2 = lRow
                            nCol2 = sRightValue
                        End If
                    
                        sRightValue = m_tgcEditor.CellValue(nCol2, lRow2) & ""
                    Else  'is a constant
                        'okay
                    End If
                End If
                
                Select Case UCase(Mid(sCondition, 2, 2))
                Case "<>"  'string/numeric not equal
                    If UCase(Mid(sCondition, 4, 1)) = "~" Then  'compare old value to a new value in the same cell
                        If m_tgcEditor.CellValueChanged(Col, lRow) Then
                            Exit For
                        End If
                    Else
                        If Trim(sLeftValue) <> Trim(sRightValue) Then
                            Exit For
                        End If
                    End If
                Case "=="  'string/numeric equal
                    If UCase(Mid(sCondition, 4, 1)) = "~" Then  'compare old value to a new value in the same cell
                        If Not m_tgcEditor.CellValueChanged(Col, lRow) Then
                            Exit For
                        End If
                    Else
                        If Trim(sLeftValue) = Trim(sRightValue) Then
                            Exit For
                        End If
                    End If
                Case "<<"  'numeric less than
                    If val(sLeftValue) < val(sRightValue) Then
                        Exit For
                    End If
                Case ">>"  'numeric greater than
                    If val(sLeftValue) > val(sRightValue) Then
                        Exit For
                    End If
                Case "LT"  'string/date less than
                    If Trim(sLeftValue) < Trim(sRightValue) Then
                        Exit For
                    End If
                Case "GT"  'string/date greater than
                    If Trim(sLeftValue) > Trim(sRightValue) Then
                        Exit For
                    End If
                End Select
            End If
        Next

        'if condition is true, set the color
        If i <= m_nCellColorCount Then
            lForeColor = m_aryCellColor(i).lForeColor
            lBackColor = m_aryCellColor(i).lBackColor
            bSetCellColor = True
        End If
    #End If
    
    'if condition is true, set the color
    If bSetCellColor Then
        CellStyle.ForeColor = lForeColor
        CellStyle.BackColor = lBackColor
    ElseIf Condition = dbgSelectedRow + dbgMarqueeRow Or Condition = dbgSelectedRow + dbgMarqueeRow + dbgCurrentCell Or _
       Condition = dbgSelectedRow Then    'selected row(s)
        CellStyle.ForeColor = &H8000000E
        CellStyle.BackColor = &H8000000D
    End If
    
    Exit Sub
    
errTrap:
    'do nothing
End Sub

Private Sub Class_Initialize()
    #If USE_ARRAY Then
        m_nCellColorCount = -1
    #End If
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    
    #If Not USE_ARRAY Then
        While colColorCondition.Count > 0
            colColorCondition.Remove 1
        Wend
    #End If
    
    Set m_tgTable = Nothing
    Set m_tgcEditor = Nothing
End Sub

