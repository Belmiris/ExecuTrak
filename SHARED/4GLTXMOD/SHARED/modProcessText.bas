Attribute VB_Name = "modProcessText"
'***********************************************
'This Module is used to process mutiline textbox
'***********************************************

Option Explicit

Private Declare Function SendMessageByVal Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Const EM_GETLINECOUNT = &HBA
Const EM_GETFIRSTVISIBLELINE = &HCE
Const EM_LINEFROMCHAR = &HC9
Const EM_LINEINDEX = &HBB
Const EM_LINELENGTH = &HC1
Const EM_FMTLINES = &HC8
Const EM_GETLINE = &HC4

' Return the number of lines in the control.
Public Function GetLineCount(tb As Textbox) As Long
    GetLineCount = SendMessageByVal(tb.hwnd, EM_GETLINECOUNT, 0, 0)
End Function

' Return the index of the first visible line
' (0 for the first text line in the control).
' When applied to a single-line control, return the
' index of the first visible character
' (0 for the first character in the control).

Public Function GetFirstVisibleLine(tb As Textbox) As Long
    GetFirstVisibleLine = SendMessageByVal(tb.hwnd, EM_GETFIRSTVISIBLELINE, 0, 0)
End Function

' Return the number of the line that contains the specified character.
' Both line and character numbers are zero-based.

Public Function LineFromChar(tb As Textbox, ByVal charIndex As Long) As Long
    LineFromChar = SendMessageByVal(tb.hwnd, EM_LINEFROMCHAR, charIndex, 0)
End Function

' Return the character offset of the first character of a line.

Public Function LineIndex(tb As Textbox, ByVal lineNum As Long) As Long
    LineIndex = SendMessageByVal(tb.hwnd, EM_LINEINDEX, lineNum, 0)
End Function

' Return the length of the specified line.

Public Function LineLength(tb As Textbox, ByVal lineNum As Long) As Long
    Dim charOffset As Long
    ' Retrieve the character offset of the first character of the line.
    charOffset = SendMessageByVal(tb.hwnd, EM_LINEINDEX, lineNum, 0)
    ' Now it is possible to retrieve the length of the line.
    LineLength = SendMessageByVal(tb.hwnd, EM_LINELENGTH, charOffset, 0)
End Function

' Return the specified line.

Public Function GetLine(tb As Textbox, ByVal lineNum As Long) As String
    Dim charOffset As Long, lineLen As Long
    ' Retrieve the character offset of the first character of the line.
    charOffset = SendMessageByVal(tb.hwnd, EM_LINEINDEX, lineNum, 0)
    ' Now it is possible to retrieve the length of the line.
    lineLen = SendMessageByVal(tb.hwnd, EM_LINELENGTH, charOffset, 0)
    ' Extract the line text.
    GetLine = Mid$(tb.Text, charOffset + 1, lineLen)
End Function

' Get the line/column coordinates of a given character (both are zero-based).
' If charIndex is negative, it returns the coordinates of the caret

Public Sub GetLineColumn(tb As Textbox, ByVal charIndex As Long, line As Long, column As Long)
    If charIndex < 0 Then charIndex = tb.SelStart
    ' Get the line number.
    line = SendMessageByVal(tb.hwnd, EM_LINEFROMCHAR, charIndex, 0)
    ' Get the column number by subtracting the line's start
    ' index from the caret position
    column = tb.SelStart - SendMessageByVal(tb.hwnd, EM_LINEINDEX, line, 0)
End Sub

'************************************************************************************
'   Functions to implement the multi-line text box
'   This modules may cause problem when the space exists at the position of nMaxlen -1 in the text box, so set the size of textbox
'   to hold maxLen characters, and use courrier new font name and 11 font size may avoid this problem.
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

Public Function fnDisplayLines(rsTemp As Recordset, Optional nMaxLen As Variant = 70, Optional fieldNum As Variant) As String
    Dim sTemp As String
    
    If rsTemp Is Nothing Then
        Exit Function
    End If
    
    If rsTemp.RecordCount = 0 Then
        Exit Function
    End If
    
    If IsMissing(fieldNum) Then
        fieldNum = 0
    End If
    
    Do Until rsTemp.EOF
        sTemp = sTemp & rsTemp.Fields(fieldNum) & vbCrLf
        rsTemp.MoveNext
    Loop
    
    fnDisplayLines = Right(sTemp, Len(sTemp) - 2)
    
End Function


'This Function will process the text in the text control, and split the text, put the result into
'table, Return the total number of lines and put all lines into array sParam()
Public Function fnBuildMultiLines(sParam() As String, _
                           ByVal sText As String, _
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
    
    If Trim(sText) = "" Or sDelim = " " Then
        fnBuildMultiLines = -1
        Exit Function
    End If
    
    If IsMissing(vStart) Then
        nLineStart = 1
    Else
        nLineStart = vStart
    End If
    
    
    If IsMissing(vEnd) Then
        nEndOfText = Len(sText)
    Else
        nEndOfText = vEnd
    End If
    
    If nLineStart < 1 Then nLineStart = 1
    
    nLineEnd = 1
    
    j = 0
    ReDim sParam(nArrayInc)
    
    While nLineStart <= nEndOfText
        nLineEnd = InStr(nLineStart, sText, sDelim)
        
        If nLineEnd = 0 Or nLineEnd > nEndOfText Then
            nLineEnd = nEndOfText + 1
        End If
        
        sTemp = Mid$(sText, nLineStart, nLineEnd - nLineStart)
        
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
    Dim sText As String
    Dim sLastText As String

    If rsTemp.RecordCount > 0 Then
        
        If IsMissing(fieldNum) Then
            fieldNum = 0
        End If
        
        sTemp = ""
        sLastText = ""
        
        While Not rsTemp.EOF
            
            If Not IsNull(rsTemp.Fields(fieldNum)) Then
                sText = rsTemp.Fields(fieldNum)
                
                If Right(sLastText, 1) <> " " Then
                    
                    If Len(sLastText) + fnGetSpaceforText(sText) > nMaxLen Then
                        
                        If Trim(sText) = "" Then
                            sTemp = sTemp + vbCrLf
                        Else
                            sTemp = sTemp + RTrim(sText)
                        End If
                        
                    Else
                        sTemp = sTemp + sText
                    End If
                
                Else
                    sTemp = RTrim(sTemp) + vbCrLf + RTrim(sText)
                End If
                
            Else
                sTemp = sTemp + Space(nMaxLen * 2) + vbCrLf + ""
                sText = ""
            End If
            
            sLastText = sText
            rsTemp.MoveNext
        Wend
        
    End If
    
    fnGetMultiLines = RTrim(sTemp)
End Function

Private Function fnGetSpaceforText(sText As String) As Integer
    Dim I As Integer
    Dim nSpaceNum As Integer
    
    If Len(sText) <= 0 Then
        Exit Function
    End If
    
    For I = Len(sText) To 1 Step -1
        
        If Asc(Mid(sText, I, 1)) <> 32 Then
            Exit For
        End If
        
        nSpaceNum = nSpaceNum + 1
    Next
    
    fnGetSpaceforText = nSpaceNum
End Function
