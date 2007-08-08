Attribute VB_Name = "modUPCValidateConvert"
Option Explicit

Public Function fnExpandUPC(sUPC As String, sErrMsg, Optional sUPCType As String, Optional bExpanded As Boolean = False) As String
    Dim sUPCCode As String
    
    sUPC = Trim(sUPC)
    sErrMsg = ""
    fnExpandUPC = ""
    
    If sUPC = "" Then
        Exit Function
    End If
    
    fnExpandUPC = sUPC
    If IsMissing(sUPCType) Then
        sUPCType = ""
    End If
    
    sErrMsg = ""
        
    If Len(sUPC) > 0 And Not IsNumeric(sUPC) Then
        sErrMsg = sUPCType & "UPC come with invalid characters in it."
        Exit Function
    End If
    
    If Len(sUPC) = 15 Then
        sUPC = fnRmLZeros(sUPC)
        If Len(sUPC) < 6 Or Len(sUPC) > 14 Or Len(sUPC) = 9 Then
            sErrMsg = sUPCType & " Invalid UPC code."
            Exit Function
        End If
    End If
    
    If Len(sUPC) < 6 Or Len(sUPC) > 14 Or Len(sUPC) = 9 Then
        fnExpandUPC = sUPC
        sErrMsg = sUPCType & " Invalid UPC length."
        Exit Function
    End If
        
    Select Case Len(sUPC)
        Case 6
            sUPCCode = "0" & fnDecompress6To10(sUPC)
            sUPCCode = sUPCCode & fnUPCLastDigit(sUPCCode)
        Case 7
            sUPCCode = "0" & fnDecompress6To10(Mid(sUPC, 1, 6)) & Right(sUPC, 1)
            If Right(sUPC, 1) <> fnUPCLastDigit(Left(sUPCCode, 11)) Then
                sUPCCode = sUPC
                sErrMsg = sUPCType & " UPC/EAN can not be determined."
            End If
        Case 8
            If bExpanded Then
                sUPCCode = sUPC
            Else
                sUPCCode = Left(sUPC, 7) & fnUPCLastDigit(Left(sUPC, 7))
                If sUPCCode <> sUPC Then
                    'sUPCCode = sUPC
                    sErrMsg = sUPCType & " Check digit is invalid for the UPC/EAN code."
                End If
            End If
        Case 10
            sUPCCode = "0" & sUPC & fnUPCLastDigit("0" & sUPC)
        Case 11
            If Left(sUPC, 1) = "0" Then
                sUPCCode = sUPC & fnUPCLastDigit(sUPC)
            Else
                'It might have two situations: the first digit is missing or the last one is missing.
                'According to Bob, we Assume the first one is missing first, then we calculate the checksum
                'If it's equal, it's OK, Otherwise, we assume the last one is missing. then calculate the checksum
                'and put it at the back
                
                sUPCCode = "0" & sUPC
                
                If Right(sUPC, 1) <> fnUPCLastDigit(Left(sUPCCode, 11)) Then
                    sUPCCode = sUPC & fnUPCLastDigit(sUPC)
                End If
                
            End If
        Case 12 'Validate the UPC
            sUPCCode = Left(sUPC, 11) & fnUPCLastDigit(Left(sUPC, 11))
            If sUPCCode <> sUPC Then
                sErrMsg = sUPCType & " Invalid check digit for UPC-A code."
            End If
        Case 13
            If bExpanded Then
                sUPCCode = sUPC
            Else
                sUPCCode = Left(sUPC, 12) & fnUPCLastDigit(Left(sUPC, 12))
                If sUPCCode <> sUPC Then
                    'sUPCCode = sUPC
                    sErrMsg = sUPCType & " Invalid check digit for the EAN-13 code."
                End If
            End If
        Case 14
            sUPCCode = Left(sUPC, 13) & fnUPCLastDigit(Left(sUPC, 13))
            
            If sUPCCode <> sUPC Then
                sErrMsg = sUPCType & " Invalid check digit for UPC-A extended code."
            End If
            
    End Select
    
    fnExpandUPC = sUPCCode
    
End Function

Public Function fnUPCLastDigit(ByVal sUPCCode As String) As String
    Dim nSumEven As Integer
    Dim nSumOdd As Integer
    Dim sUPCLastDigit As String
    Dim nMaxLen As Integer
    Dim i As Integer
    Dim sUPCRightMost As String
    
On Error GoTo errHandler
    
    ' Actually the lenght all UPC code here must be 7, 11, 12 or 13.
    ' Algorithm: Assign odd/even to each character moving from right to left
    
    nMaxLen = Len(sUPCCode)
    
    sUPCRightMost = ""
    For i = nMaxLen To 1 Step -1
        sUPCRightMost = sUPCRightMost & Mid(sUPCCode, i, 1)
    Next i
    
    nSumEven = 0
    nSumOdd = 0
    For i = 1 To nMaxLen
        If i Mod 2 = 0 Then
            nSumEven = nSumEven + Mid(sUPCRightMost, i, 1)
        Else
            nSumOdd = nSumOdd + Mid(sUPCRightMost, i, 1)
        End If
    Next i
    
    sUPCLastDigit = 10 - Right(CStr((nSumOdd * 3) + nSumEven), 1)
    If sUPCLastDigit = "10" Then sUPCLastDigit = "0"
    
    fnUPCLastDigit = sUPCLastDigit
    Exit Function
    
errHandler:
    MsgBox " modUPCValidateConvert.fnUPCLastDigit: " & vbCrLf _
         & " This UPC code " & sUPCCode & " not correct." & vbCrLf _
         & " Please contact FACTOR."
    fnUPCLastDigit = "0"
    
End Function

Public Function fnDecompress6To10(sUPC As String) As String
    Dim sUPCCode As String
    
    If Len(sUPC) <> 6 Then
        MsgBox "ERROR!"
        Exit Function
    End If
    
    Select Case Right(sUPC, 1)
        Case 0, 1, 2
            sUPCCode = Left(sUPC, 2) & Mid(sUPC, 6, 1) & "0000" & Mid(sUPC, 3, 3)
        Case 3
            sUPCCode = Left(sUPC, 3) & "00000" & Mid(sUPC, 4, 2)
        Case 4
            sUPCCode = Left(sUPC, 4) & "00000" & Mid(sUPC, 5, 1)
        Case Else
            sUPCCode = Left(sUPC, 5) & "0000" & Right(sUPC, 1)
    End Select
    
    fnDecompress6To10 = sUPCCode

End Function
Public Function fnRmLZeros(sUPC As String) As String
    Dim i As Integer
    Dim retStr As String
    
    ' Whenever the length of sUPC or retStr is less than 6
    ' then the sUPC is not valid UPC, it will be taken care
    ' by the caller function
    
    retStr = sUPC
    If Len(retStr) < 6 Then
        fnRmLZeros = retStr
        Exit Function
    End If
    
    For i = 1 To Len(sUPC)
        If Mid(sUPC, i, 1) = "0" Then
            retStr = Right(retStr, Len(retStr) - 1)
        Else
            Exit For
        End If
        If Len(retStr) < 6 Then
            Exit For
        End If
    Next
    
    fnRmLZeros = retStr
    
End Function

