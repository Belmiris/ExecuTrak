Attribute VB_Name = "ModMyText"
' this module deal with text

Option Explicit

'next function will link the text of the first field of
' the recordset.
' the difference with fngettext is : No trim here
Public Function fnGetTextFromRecordset(ByVal rsTemp As Recordset, Optional bTrim) As String
    
    Dim i As Integer
    Dim k As Integer
    Dim sText As String
    Dim sTemp As String
    Dim nPos As Integer
    
    On Error Resume Next
    sText = ""
    fnGetTextFromRecordset = ""
    If rsTemp Is Nothing Then
       Exit Function
    End If
    If rsTemp.RecordCount < 1 Then
       Exit Function
    End If
    rsTemp.MoveLast
    rsTemp.MoveFirst
    k = rsTemp.RecordCount
    For i = 1 To k
        If Not IsNull(rsTemp.Fields(0)) Then
           sTemp = rsTemp.Fields(0)
           ' This is for 'HP' type display only 1/27/98
           ' (1) Replace chr(13) vb vbcrlf
           ' (2) Right trim extra space
           If Not IsMissing(bTrim) Then ' only in this case ,do something
               If CBool(bTrim) Then
                    If Right(sTemp, 1) = szSPACE Then
                        sTemp = RTrim(rsTemp.Fields(0)) & szSPACE
                    End If
                    nPos = InStr(sTemp, Chr(13))
                    If nPos <> 0 And InStr(sTemp, vbCrLf) = 0 Then
                        sTemp = Left(sTemp, nPos - 1) & vbCrLf & LTrim(Right(sTemp, Len(sTemp) - nPos))
                    End If
               End If
           End If
           sText = sText & sTemp
        End If
        rsTemp.MoveNext
    Next i
    fnGetTextFromRecordset = sText
  
End Function

'(1) This function will split sztext in to blocks
'(2) Each Block has a max of lenOfEachLine char's
'(3) It will return the number of blocks and the split blocks
'    will be stored in the myArray()
'When we update: if fnSplitText returns n
'Call    for i=0 to n
'        next
Public Function fnSplitText(szText As String, myArray() As String, _
             lenOfEachLine As Integer) As Integer
       Dim szTemp As String
       Dim intLines As Integer
       Dim k As Integer
       Dim intLenOfText As Integer
       
       intLenOfText = Len(szText)
       intLines = intLenOfText \ lenOfEachLine
      
       For k = 0 To intLines
          ReDim Preserve myArray(k)
          myArray(k) = Mid(szText, k * lenOfEachLine + 1, lenOfEachLine)
        
       Next k
       fnSplitText = intLines
            
End Function


Public Function myReplace(sText As String, ByVal sKey As String, _
                                     ByVal sNew As String) As String
   
    Dim nIdx As Integer
    Dim nPos As Integer
    
    If Trim(sText) = "" Then
        myReplace = ""
        Exit Function
    End If
    
    nIdx = 1
    nPos = InStr(nIdx, sText, sKey) ' find the start position of the key
    
    While nPos <> 0
        sText = Left(sText, nPos - 1) & sNew & Right(sText, Len(sText) - nPos - Len(sKey) + 1)
        nIdx = nPos + Len(sKey)
        nPos = InStr(nIdx, sText, sKey)
    Wend
    myReplace = sText
End Function

