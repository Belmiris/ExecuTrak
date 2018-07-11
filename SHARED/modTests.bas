Attribute VB_Name = "modTests"
' This module was created to identify scenarios which need testing

'Print G9126Test1(1.0825)
' 1.08249998092651
'
'Print G9126Test1(1.0825, 1)
' 1.0825
'
'Print G9126Test1(1.0825, 2)
' 1.0825
'
Public Function G9126Test1(sng As Single, Optional Scenario As Integer = 0) As Double
    
    Select Case Scenario
        Case 1 ' Cstr()
            G9126Test1 = CDbl(CStr(sng))
        Case 2 ' tfnRound()
            G9126Test1 = CDbl(tfnRound(sng, 8))
        Case Else ' raw
            G9126Test1 = CDbl(sng)
    End Select
End Function

'Print G8967Test1(1.0825)
' 1.08249998092651
'
'Print G8967Test1(1.0825, 1)
' 1.0825
'
'Print G8967Test1(1.0825, 2)
' 1.0825
'
Public Function G8967Test1(sng As Single, Optional Scenario As Integer = 0) As Double
    
    Select Case Scenario
        Case 1 ' Cstr()
            G8967Test1 = CDbl(CStr(sng)) + 1
        Case 2 ' tfnRound()
            G8967Test1 = CDbl(tfnRound(sng, 6)) + 1
        Case Else ' raw +1
            G8967Test1 = CDbl(sng + 1)
    End Select
End Function

