Attribute VB_Name = "modProcessText"
Option Explicit

Private Declare Function SendMessageByVal Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Const EM_GETLINECOUNT = &HBA
Const EM_GETFIRSTVISIBLELINE = &HCE
Const EM_LINEFROMCHAR = &HC9
Const EM_LINEINDEX = &HBB
Const EM_LINELENGTH = &HC1
Const EM_FMTLINES = &HC8
Const EM_GETLINE = &HC4


'************************************************************************************
'   Functions to implement the multi-line text box
'************************************************************************************

' Return an array with all the lines in the multiline textbox.
' If the second optional argument is True, the trailing CR-LF is preserved.

Public Function fnGetAllLines(tb As Textbox, Optional KeepHardLineBreaks As Boolean) As String()
    Dim result() As String, I As Long
    
    SendMessageByVal tb.hwnd, EM_FMTLINES, True, 0
    result() = Split(tb.Text, vbCrLf)
    
    For I = 0 To UBound(result)
        
        If Right$(result(I), 1) = vbCr Then
            result(I) = Left$(result(I), Len(result(I)) - 1)
        ElseIf KeepHardLineBreaks Then
            result(I) = result(I) & vbCrLf
        End If
        
    Next
    
    SendMessageByVal tb.hwnd, EM_FMTLINES, False, 0
    fnGetAllLines = result()
End Function

'This Function will process the text in the text control, and split the text, put the result into
'table, Return the total number of lines and put all lines into array sParam()
Public Function fnBuildMultiLines(sParam() As String, _
                           ByVal stext As String, _
                           sDelim As String, _
                           Optional nMaxLen As Variant = 70, _
                           Optional vStart As Variant, _
                           Optional vEnd As Variant) As Integer
    
    Const nArrayInc As Integer = 5
    Dim nLineStart As Integer
    Dim nLineEnd As Integer
    Dim nEndOfText As Integer
    Dim sTemp As String
    Dim I, j, nLines As Integer
    
    'On Error Resume Next
    
    If Trim(stext) = "" Or sDelim = " " Then
        Exit Function
    End If
    
    If IsMissing(vStart) Then
        nLineStart = 1
    Else
        nLineStart = vStart
    End If
    
    
    If IsMissing(vEnd) Then
        nEndOfText = Len(stext)
    Else
        nEndOfText = vEnd
    End If
    
    If nLineStart < 1 Then nLineStart = 1
    
    nLineEnd = 1
    
    j = 0
    ReDim sParam(nArrayInc)
    
    While nLineStart <= nEndOfText
        nLineEnd = InStr(nLineStart, stext, sDelim)
        
        If nLineEnd = 0 Or nLineEnd > nEndOfText Then
            nLineEnd = nEndOfText + 1
        End If
        
        sTemp = Mid$(stext, nLineStart, nLineEnd - nLineStart)
        
        If RTrim(sTemp) <> "" Then sTemp = RTrim(sTemp)
        
        If sTemp <> "" Or nLineEnd = nLineStart Then
            
            If Len(sTemp) Mod nMaxLen = 0 And Len(sTemp) > 0 Then
                nLines = Len(sTemp) \ nMaxLen - 1
            Else
                nLines = Len(sTemp) \ nMaxLen
            End If
            
            If j + nLines + 1 > UBound(sParam) Then
                ReDim Preserve sParam(j + nLines + nArrayInc + 1)
            End If
            
            For I = 0 To nLines
                sParam(j) = Mid(sTemp, I * nMaxLen + 1, nMaxLen)
                
                If Len(sParam(j)) < nMaxLen Then
                    sParam(j) = sParam(j) + Space(nMaxLen - Len(sParam(j)))
                End If
                Debug.Print sParam(j)
                j = j + 1
            Next I
            
        End If
        
        nLineStart = nLineEnd + Len(sDelim)
        
    Wend
    
    j = j - 1
    
    Do While j >= 0

        If Trim(sParam(j)) <> "" Then
            Exit Do
        End If

        j = j - 1
    Loop

    If j < 0 Then j = 0
    
    ReDim Preserve sParam(j)
    
    fnBuildMultiLines = j
    
End Function

'   The following function accepts a recordset, and returns a string which can
'   be put into the text box.
Public Function fnGetMultiLines(rsTemp As Recordset, Optional nMaxLen As Variant = 70, Optional fieldNum As Variant) As String
    Dim sTemp As String
    Dim stext As String
    Dim sLastText As String

    If rsTemp.RecordCount > 0 Then
        
        If IsMissing(fieldNum) Then
            fieldNum = 0
        End If
        
        sTemp = ""
        sLastText = ""
        
        While Not rsTemp.EOF
            
            If Not IsNull(rsTemp.Fields(fieldNum)) Then
                stext = rsTemp.Fields(fieldNum)
                
                If Right(sLastText, 1) <> " " Then
                    
                    If Len(sLastText) + fnGetSpaceforText(stext) > nMaxLen Then
                        
                        If Trim(stext) = "" Then
                            sTemp = sTemp + vbCrLf
                        Else
                            sTemp = sTemp + RTrim(stext)
                        End If
                        
                    Else
                        sTemp = sTemp + stext
                    End If
                
                Else
                    sTemp = RTrim(sTemp) + vbCrLf + RTrim(stext)
                End If
                
            Else
                sTemp = sTemp + Space(nMaxLen * 2) + vbCrLf + ""
                stext = ""
            End If
            
            sLastText = stext
            rsTemp.MoveNext
        Wend
        
    End If
    
    fnGetMultiLines = RTrim(sTemp)
End Function

Private Function fnGetSpaceforText(stext As String) As Integer
    Dim I As Integer
    Dim nSpaceNum As Integer
    
    If Len(stext) <= 0 Then
        Exit Function
    End If
    
    For I = Len(stext) To 1 Step -1
        
        If Asc(Mid(stext, I, 1)) <> 32 Then
            Exit For
        End If
        
        nSpaceNum = nSpaceNum + 1
    Next
    
    fnGetSpaceforText = nSpaceNum
End Function
