Attribute VB_Name = "modComboSubs"
'   Public functions:
'       1. Sub fnComboAddItem(txtPrim As TextBox, ParamArray txtScnd() As Variant)
'       2. Sub fnComboGetText(txtCode As TextBox)
'       3. Sub fnComboResetFlags()
'       4. Sub fnComboSetText(txtCode As TextBox, Optional vText As Variant)
'       5. Function fnGetSecondaryBox(txtCode As TextBox, Optional vPos As Variant) As TextBox
'   Event calls:
'       1. fnComboChange(txtCodeP As TextBox)
'       2. Sub fnComboLostFocus(txtCode As TextBox, ParamArray arryControls())
        
Option Explicit

    Private Type tpCombos
        txtPrimary As textbox
        txtSecondary As Control
        bTextFilled As Boolean
        sText As String
    End Type
    Private arryCombo() As tpCombos
    Private nComboCount As Integer


Public Sub fnComboAddItem(txtPrim As Control, ParamArray txtScnd() As Variant)
    Dim nCount As Integer
    Dim i As Integer
    
    nCount = UBound(txtScnd) + 1
    ReDim Preserve arryCombo(nComboCount + nCount)

    For i = 0 To nCount - 1
        With arryCombo(nComboCount + i)
            Set .txtPrimary = txtPrim
            Set .txtSecondary = txtScnd(i)
        End With
    Next
    nComboCount = nComboCount + nCount
End Sub

Public Sub fnComboSetText(txtCode As Control, Optional vText As Variant)
    Dim nIdx As Integer

    If IsMissing(vText) Then
        nIdx = -1
        Do
            nIdx = fnComboGetIndexP(txtCode, nIdx + 1)
            If nIdx >= 0 Then
                With arryCombo(nIdx)
                    If TypeOf .txtSecondary Is textbox Then
                        .sText = .txtSecondary.Text
                    Else
                        .sText = .txtSecondary.Caption
                    End If
                    .bTextFilled = True
                End With
            End If
        Loop Until nIdx < 0
    Else
        nIdx = fnComboGetIndexS(txtCode)
        If nIdx >= 0 Then
            With arryCombo(nIdx)
                If IsNull(vText) Then
                    .sText = ""
                Else
                    .sText = Trim(vText)
                End If
                .bTextFilled = True
            End With
        End If
    End If
End Sub

Public Sub fnComboChange(txtCode As textbox)
    Dim nIdx As Integer

    nIdx = -1
    Do
        nIdx = fnComboGetIndexP(txtCode, nIdx + 1)
        If nIdx >= 0 Then
            arryCombo(nIdx).bTextFilled = False
        End If
    Loop Until nIdx < 0

End Sub

Public Sub fnComboGetText(txtCode As textbox)
    
    Dim nIdx As Integer

    nIdx = -1
    Do
        nIdx = fnComboGetIndexP(txtCode, nIdx + 1)
        If nIdx >= 0 Then
            With arryCombo(nIdx)
                If .bTextFilled Then
                    If TypeOf .txtSecondary Is textbox Then
                        .txtSecondary.Text = .sText
                    Else
                        .txtSecondary.Caption = .sText
                    End If
                End If
            End With
        End If
    Loop Until nIdx < 0

End Sub

Private Function fnComboGetIndexP(txtCode As textbox, nStart As Integer)
    Dim i As Integer
    
    fnComboGetIndexP = -2
    If nStart < 0 Then
        Exit Function
    End If

    For i = nStart To nComboCount - 1
        If txtCode.TabIndex = arryCombo(i).txtPrimary.TabIndex Then
            fnComboGetIndexP = i
            Exit Function
        End If
    Next i
End Function

Private Function fnComboGetIndexS(txtCode As Control)
    Dim i As Integer
    
    fnComboGetIndexS = -1
    If txtCode Is Nothing Then
        Exit Function
    End If
    For i = 0 To nComboCount - 1
        If txtCode.TabIndex = arryCombo(i).txtSecondary.TabIndex Then
            fnComboGetIndexS = i
            Exit Function
        End If
    Next i
End Function

Public Sub fnComboKeyPress(txtCode As textbox)

End Sub

Public Sub fnComboLostFocus(txtCode As textbox, ParamArray arryControls())
    Dim nIdx As Integer

    For nIdx = 0 To UBound(arryControls)
        If frmMain.ActiveControl.TabIndex = arryControls(nIdx).TabIndex Then
            Exit Sub
        End If
    Next nIdx
    
    nIdx = -1
    Do
        nIdx = fnComboGetIndexP(txtCode, nIdx + 1)
        If nIdx >= 0 Then
        With arryCombo(nIdx)
            If .bTextFilled Then
                If TypeOf .txtSecondary Is textbox Then
                    .txtSecondary.Text = .sText
                Else
                    .txtSecondary.Caption = .sText
                End If
            End If
        End With
        End If
    Loop Until nIdx < 0

End Sub

Public Sub fnComboResetFlags()
    Dim i As Integer
    For i = 0 To nComboCount - 1
        arryCombo(i).bTextFilled = False
    Next i
End Sub

Public Function fnGetSecondaryBox(txtCode As textbox, Optional vPos As Variant) As Control
    Dim nIdx As Integer
    Dim j As Integer
    Dim nPos As Integer
    
    If IsMissing(vPos) Then
        nPos = 1
    Else
        nPos = vPos
    End If

    nIdx = fnComboGetIndexP(txtCode, 0)
    If nIdx >= 0 Then
        For j = 1 To nPos - 1
            nIdx = nIdx + 1
        Next
        If nIdx < nComboCount Then
            Set fnGetSecondaryBox = arryCombo(nIdx).txtSecondary
        Else
            Set fnGetSecondaryBox = Nothing
        End If
    End If

End Function


