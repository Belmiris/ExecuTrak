Attribute VB_Name = "modJQIU"
Option Explicit

Private Declare Function SendMessageByVal Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Const EM_GETLINECOUNT = &HBA
Const EM_GETFIRSTVISIBLELINE = &HCE
Const EM_LINEFROMCHAR = &HC9
Const EM_LINEINDEX = &HBB
Const EM_LINELENGTH = &HC1
Const EM_FMTLINES = &HC8
Const EM_GETLINE = &HC4

' Return an array with all the lines in the multiline textbox.
' If the second optional argument is True, the trailing CR-LF is preserved.

Public Function fnGetAllLines(tb As Textbox, Optional KeepHardLineBreaks As Boolean) As String()
    Dim result() As String, i As Long
    
    SendMessageByVal tb.hwnd, EM_FMTLINES, True, 0
    result() = Split(tb.Text, vbCrLf)
    
    For i = 0 To UBound(result)
        
        If Right$(result(i), 1) = vbCr Then
            result(i) = Left$(result(i), Len(result(i)) - 1)
        ElseIf KeepHardLineBreaks Then
            result(i) = result(i) & vbCrLf
        End If
        
    Next
    
    SendMessageByVal tb.hwnd, EM_FMTLINES, False, 0
    fnGetAllLines = result()
End Function


