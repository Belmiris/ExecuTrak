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

Private Const m_COLHEADING As String = "C-H"
Private Const m_ALIGNLEFT As String = "A-L"
Private Const m_FONTBOLD As String = "F-B"

Private m_tgTable As TDBGrid
Private m_tgcEditor As clsTGSpreadSheet
Private colColorCondition As New Collection
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
    Static bAlreadyShowedMsg As Boolean
    
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
    'op can be (==, <>, <=, >=, =, <, >, NE, EQ, LE, GE, LT, GT),
    'Y can be ~ means new cell value,
    'Y can be a constant number,
    'Y can be [col,row] means a cell value in col,row
    
    Dim sMsg As String
    sMsg = "Invalid number of arguments. S/B: " + vbCrLf _
        + "Col, condition1, forecolor1, backcolor1, condition2, , forecolor2, backcolor2, ..." + vbCrLf + vbCrLf _
        + "Need to have 3 parm in a group, namely Condition, ForeColor, and BackColor" + vbCrLf _
        + "Syntax for condition check: XopY" + vbCrLf _
        + "X means old cell value" + vbCrLf _
        + "X can also be [col,row], means a cell value col,row" + vbCrLf _
        + "op can be (==, <>, <=, >=, =, <, >, NE, EQ, LE, GE, LT, GT)," + vbCrLf _
        + "Y can be ~ means new cell value" + vbCrLf _
        + "Y can be a constant number" + vbCrLf _
        + "Y can be [col,row] means a cell value in col,row"
    
    If (i + 1) Mod 3 <> 0 Then
        If Not bAlreadyShowedMsg Then
            MsgBox sMsg
            bAlreadyShowedMsg = True
        End If
        
        Exit Sub
    Else
        If Not IsNumeric(aryColorCond(1)) Then
            If Not bAlreadyShowedMsg Then
                MsgBox sMsg
                bAlreadyShowedMsg = True
            End If
            
            Exit Sub
        End If
    End If
    
    Dim objColorCondition As New clsColorCondition
    objColorCondition.AddCondition nCol, aryColorCond
    
    Dim objOldColorCondition As clsColorCondition
    
    If IsColConditionExists(nCol, objOldColorCondition) Then
        'append the condition into the existing array
        objOldColorCondition.AppendCondition objColorCondition
    Else
        colColorCondition.Add objColorCondition, "C" & nCol
    End If
    
    m_tgTable.Columns(nCol).FetchStyle = True
    
    Exit Sub

errTrap:
    'do nothing
End Sub

Public Sub AddColumn4CellStyle(nCol As Integer, ParamArray aryStyleCond() As Variant)
    Static bAlreadyShowedMsg As Boolean
    
    Dim i As Integer
    Dim j As Integer
    Dim n As Integer
        
    On Error GoTo errTrap
    
    i = -1
    
    On Error Resume Next
    i = UBound(aryStyleCond)
    On Error GoTo 0
    
    If i < 0 Then
        Exit Sub
    End If
    
    'need to have 2 parm in a group
    'namely Condition and CellStyle constant (e.g. COLHEADING, ALIGNLEFT, FONTBOLD)
    
    'syntax for condition check: XopY,
    'X means old cell value
    'X can also be [col,row], means a cell value col,row
    'op can be (==, <>, <=, >=, =, <, >, NE, EQ, LE, GE, LT, GT),
    'Y can be ~ means new cell value,
    'Y can be a constant number,
    'Y can be [col,row] means a cell value in col,row
    'CellStyle:
    'COLHEADING means the cell value is a caption for column heading
    'ALIGNLEFT means the left justified the cell value
    'FONTBOLD means show the fonts in bold typeface
    
    Dim sMsg As String
    sMsg = "Invalid number of arguments. S/B: " + vbCrLf _
        + "Col, condition1, cellstyle1, condition2, cellstyle2, ..." + vbCrLf + vbCrLf _
        + "Need to have 2 parm in a group, namely Condition and CellStyle constant (e.g. COLHEADING, ALIGNLEFT, FONTBOLD)" + vbCrLf _
        + "Syntax for condition check: XopY" + vbCrLf _
        + "X means old cell value" + vbCrLf _
        + "X can also be [col,row], means a cell value col,row" + vbCrLf _
        + "op can be (==, <>, <=, >=, =, <, >, NE, EQ, LE, GE, LT, GT)," + vbCrLf _
        + "Y can be ~ means new cell value" + vbCrLf _
        + "Y can be a constant number" + vbCrLf _
        + "Y can be [col,row] means a cell value in col,row" + vbCrLf _
        + "CellStyle:" + vbCrLf _
        + "COLHEADING means the cell value is a caption for column heading" + vbCrLf _
        + "ALIGNLEFT means the left justified the cell value" + vbCrLf _
        + "FONTBOLD means show the fonts in bold typeface"
    
    If (i + 1) Mod 2 <> 0 Then
        If Not bAlreadyShowedMsg Then
            MsgBox sMsg
            bAlreadyShowedMsg = True
        End If
        
        Exit Sub
    Else
        If IsNumeric(aryStyleCond(1)) Then
            If Not bAlreadyShowedMsg Then
                MsgBox sMsg
                bAlreadyShowedMsg = True
            End If
            
            Exit Sub
        End If
    End If
    
    Dim objColorCondition As New clsColorCondition
    objColorCondition.AddCondition nCol, aryStyleCond
    
    Dim objOldColorCondition As clsColorCondition
    
    If IsColConditionExists(nCol, objOldColorCondition) Then
        'append the condition into the existing array
        objOldColorCondition.AppendCondition objColorCondition
    Else
        colColorCondition.Add objColorCondition, "C" & nCol
    End If
    
    m_tgTable.Columns(nCol).FetchStyle = True
    
    Exit Sub

errTrap:
    'do nothing
End Sub

Public Function IsColConditionExists(nCol As Integer, Optional objColorCondition As clsColorCondition = Nothing) As Boolean
    On Error GoTo errNoExists
    
    If Not colColorCondition Is Nothing Then
        Set objColorCondition = colColorCondition.Item("C" & nCol)
        IsColConditionExists = True
    End If
    
    Exit Function
    
errNoExists:
    'do nothing, return false
End Function

Public Sub RemoveColumn4CellColor(nCol As Integer)
    On Error Resume Next
    Dim objOldColorCondition As clsColorCondition
    
    If IsColConditionExists(nCol, objOldColorCondition) Then
        Set objOldColorCondition = Nothing
        colColorCondition.Remove "C" & nCol
    End If
    
    m_tgTable.Columns(nCol).FetchStyle = False
End Sub

Property Get COLHEADING() As String
    COLHEADING = m_COLHEADING
End Property

Property Get ALIGNLEFT() As String
    ALIGNLEFT = m_ALIGNLEFT
End Property

Property Get FONTBOLD() As String
    FONTBOLD = m_FONTBOLD
End Property

Property Let FetchStyle(nCol As Integer, bOnOff As Boolean)
    m_tgTable.Columns(nCol).FetchStyle = bOnOff
End Property

Property Get IsColumnForCellColor(nCol As Integer) As Boolean
    Dim objColorCondition As clsColorCondition
    
    On Error Resume Next
    Set objColorCondition = colColorCondition.Item("C" & nCol)
    
    If Not objColorCondition Is Nothing Then
        IsColumnForCellColor = True
        Exit Property
    End If
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
    Dim sCellStyle As String
    Dim bSetCellColor As Boolean
    Dim bColorIsSet As Boolean
    
    On Error GoTo errTrap
    
    lRow = Val(Bookmark)
    
    If lRow < 0 Then
        Exit Sub
    End If
    
    Dim objColorCondition As clsColorCondition
    'get the color condition
    On Error Resume Next
    Set objColorCondition = colColorCondition.Item("C" & Col)
    On Error GoTo 0
    
    If Not objColorCondition Is Nothing Then
        bColorIsSet = False
        
        For i = 0 To objColorCondition.ConditionCount
            If Not objColorCondition.GetCondition(i, sCondition, lForeColor, lBackColor, sCellStyle) Then
                Exit Sub
            End If
            
            'syntax for condition check: XopY,
            'X means old cell value
            'X can also be [col,row], means a cell value col,row
            'op can be (==, <>, <=, >=, =, <, >, NE, EQ, LE, GE, LT, GT),
            'Y can be ~ means new cell value,
            'Y can be a constant number,
            'Y can be [col,row] means a cell value in col,row
            'CellStyle:
            'COLHEADING means the cell value is a caption for column heading
            'ALIGNLEFT means the left justified the cell value
            'FONTBOLD means show the fonts in bold typeface
            
            'get the op

'            'same logic as the if blocks below
'            'the "if block" logic may have better performance
'            Dim aryOP()
'            aryOP = Array("<>", "==", "<=", ">=", "=", "<", ">", "NE", "EQ", "LE", "GE", "LT", "GT")
'
'            sOP = ""
'
'            For nCol2 = 0 To UBound(aryOP)
'                nPosi = InStr(sCondition, aryOP(nCol2))
'                If nPosi > 0 Then
'                    sOP = aryOP(nCol2)
'                    Exit For
'                End If
'            Next nCol2
'
'            If sOP = "" Then
'                'op is not valid
'                Exit Sub
'            End If

            'same logic as the for loop above
            'but the "if block" logic may have better performance
            sOP = "<>"
            nPosi = InStr(sCondition, sOP)
            If nPosi <= 0 Then
                sOP = "=="
                nPosi = InStr(sCondition, sOP)
                If nPosi <= 0 Then
                    sOP = "<="
                    nPosi = InStr(sCondition, sOP)
                    If nPosi <= 0 Then
                        sOP = ">="
                        nPosi = InStr(sCondition, sOP)
                        If nPosi <= 0 Then
                            sOP = "="
                            nPosi = InStr(sCondition, sOP)
                            If nPosi <= 0 Then
                                sOP = "<"
                                nPosi = InStr(sCondition, sOP)
                                If nPosi <= 0 Then
                                    sOP = ">"
                                    nPosi = InStr(sCondition, sOP)
                                    If nPosi <= 0 Then
                                        sOP = "NE"
                                        nPosi = InStr(sCondition, sOP)
                                        If nPosi <= 0 Then
                                            sOP = "EQ"
                                            nPosi = InStr(sCondition, sOP)
                                            If nPosi <= 0 Then
                                                sOP = "LE"
                                                nPosi = InStr(sCondition, sOP)
                                                If nPosi <= 0 Then
                                                    sOP = "GE"
                                                    nPosi = InStr(sCondition, sOP)
                                                    If nPosi <= 0 Then
                                                        sOP = "LT"
                                                        nPosi = InStr(sCondition, sOP)
                                                        If nPosi <= 0 Then
                                                            sOP = "GT"
                                                            nPosi = InStr(sCondition, sOP)
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
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
            sRightSide = Mid(sCondition, nPosi + Len(sOP))
            
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
                
                If Left(sRightValue, 1) = "[" Then
                    'get another col, row
                    'remove ]
                    sRightValue = Mid(sRightValue, 2, Len(sRightValue) - 2)
                    
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
            End If
            
            bSetCellColor = False
            
            Select Case UCase(sOP)
            Case "<>"  'numeric/date/string not equal to
                If IsNumeric(sLeftValue) And IsNumeric(sRightValue) Then
                    If Val(sLeftValue) <> Val(sRightValue) Then
                        bSetCellColor = True
                    End If
                ElseIf IsDate(sLeftValue) And IsDate(sRightValue) Then
                    If CDate(sLeftValue) <> CDate(sRightValue) Then
                        bSetCellColor = True
                    End If
                Else  'string compare
                    If Trim(sLeftValue) <> Trim(sRightValue) Then
                        bSetCellColor = True
                    End If
                End If
            Case "==", "=" 'numeric/date/string equal to
                If IsNumeric(sLeftValue) And IsNumeric(sRightValue) Then
                    If Val(sLeftValue) = Val(sRightValue) Then
                        bSetCellColor = True
                    End If
                ElseIf IsDate(sLeftValue) And IsDate(sRightValue) Then
                    If CDate(sLeftValue) = CDate(sRightValue) Then
                        bSetCellColor = True
                    End If
                Else  'string compare
                    If Trim(sLeftValue) = Trim(sRightValue) Then
                        bSetCellColor = True
                    End If
                End If
            Case "<="  'numeric/date/string less than and equal to
                If IsNumeric(sLeftValue) And IsNumeric(sRightValue) Then
                    If Val(sLeftValue) <= Val(sRightValue) Then
                        bSetCellColor = True
                    End If
                ElseIf IsDate(sLeftValue) And IsDate(sRightValue) Then
                    If CDate(sLeftValue) <= CDate(sRightValue) Then
                        bSetCellColor = True
                    End If
                Else  'string compare
                    If Trim(sLeftValue) <= Trim(sRightValue) Then
                        bSetCellColor = True
                    End If
                End If
            Case ">="  'numeric/date/string greater than and equal to
                If IsNumeric(sLeftValue) And IsNumeric(sRightValue) Then
                    If Val(sLeftValue) >= Val(sRightValue) Then
                        bSetCellColor = True
                    End If
                ElseIf IsDate(sLeftValue) And IsDate(sRightValue) Then
                    If CDate(sLeftValue) >= CDate(sRightValue) Then
                        bSetCellColor = True
                    End If
                Else  'string compare
                    If Trim(sLeftValue) >= Trim(sRightValue) Then
                        bSetCellColor = True
                    End If
                End If
            Case "<"  'numeric/date/string less than
                If IsNumeric(sLeftValue) And IsNumeric(sRightValue) Then
                    If Val(sLeftValue) < Val(sRightValue) Then
                        bSetCellColor = True
                    End If
                ElseIf IsDate(sLeftValue) And IsDate(sRightValue) Then
                    If CDate(sLeftValue) < CDate(sRightValue) Then
                        bSetCellColor = True
                    End If
                Else  'string compare
                    If Trim(sLeftValue) < Trim(sRightValue) Then
                        bSetCellColor = True
                    End If
                End If
            Case ">"  'numeric/date/string greater than
                If IsNumeric(sLeftValue) And IsNumeric(sRightValue) Then
                    If Val(sLeftValue) > Val(sRightValue) Then
                        bSetCellColor = True
                    End If
                ElseIf IsDate(sLeftValue) And IsDate(sRightValue) Then
                    If CDate(sLeftValue) > CDate(sRightValue) Then
                        bSetCellColor = True
                    End If
                Else  'string compare
                    If Trim(sLeftValue) > Trim(sRightValue) Then
                        bSetCellColor = True
                    End If
                End If
            Case "NE"  'string not equal to
                If Trim(sLeftValue) <> Trim(sRightValue) Then
                    bSetCellColor = True
                End If
            Case "EQ"  'string equal to
                If Trim(sLeftValue) = Trim(sRightValue) Then
                    bSetCellColor = True
                End If
            Case "LE"  'string less than and equal to
                If Trim(sLeftValue) <= Trim(sRightValue) Then
                    bSetCellColor = True
                End If
            Case "GE"  'string greater than and equal to
                If Trim(sLeftValue) >= Trim(sRightValue) Then
                    bSetCellColor = True
                End If
            Case "LT"  'string less than
                If Trim(sLeftValue) < Trim(sRightValue) Then
                    bSetCellColor = True
                End If
            Case "GT"  'string greater than
                If Trim(sLeftValue) > Trim(sRightValue) Then
                    bSetCellColor = True
                End If
            End Select
        
            'if condition is true, set the color or cellstyle
            If bSetCellColor Then
                If sCellStyle <> "" Then  'is a caption for column heading
                    If sCellStyle = COLHEADING Then   'is a caption for column heading
                        CellStyle = m_tgTable.Styles("Heading")
                        CellStyle.Alignment = dbgLeft
                        CellStyle.Locked = True
                    ElseIf sCellStyle = ALIGNLEFT Then    'is a caption for column heading
                        CellStyle.Alignment = dbgLeft
                    ElseIf sCellStyle = FONTBOLD Then     'is a caption for column heading
                        CellStyle.Font.Bold = True
                    End If
                Else
                    If UCase(sRightSide) = COLHEADING Then   'is a caption for column heading
                        CellStyle = m_tgTable.Styles("Heading")
                        CellStyle.Alignment = dbgLeft
                        CellStyle.Locked = True
                    ElseIf UCase(sRightSide) = ALIGNLEFT Then    'is a caption for column heading
                        CellStyle.Alignment = dbgLeft
                    ElseIf UCase(sRightSide) = FONTBOLD Then     'is a caption for column heading
                        CellStyle.Font.Bold = True
                    Else
                        If lForeColor >= 0 Then
                            CellStyle.ForeColor = lForeColor
                        End If
                        If lBackColor >= 0 Then
                            CellStyle.BackColor = lBackColor
                        End If
                    End If
                End If
            End If
        Next
    End If
    
    Exit Sub

errTrap:
    'do nothing
End Sub

Private Sub Class_Initialize()
    'nothing
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    
    While colColorCondition.Count > 0
        colColorCondition.Remove 1
    Wend
    
    Set m_tgTable = Nothing
    Set m_tgcEditor = Nothing
End Sub

